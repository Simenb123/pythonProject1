from __future__ import annotations
from typing import Dict, Optional, List, Tuple, Set
import os
import duckdb

from . import settings as S
from .detect import detect_csv_options, read_headers, duck_quote, sql_lit

# Ekstra tekstfelt vi forsøker å hente (hvis de finnes i CSV)
EXTRA_TEXT_COLUMNS: Dict[str, str] = {
    "share_class":     "Aksjeklasse",
    "owner_country":   "Landkode",
    "owner_zip_place": "Postnr/sted",
}

# Forventet skjema – opprettes eksplisitt (GUI forventer disse 10)
SCHEMA_COLS: List[Tuple[str, str]] = [
    ("company_orgnr",      "VARCHAR"),
    ("company_name",       "VARCHAR"),
    ("owner_orgnr",        "VARCHAR"),
    ("owner_name",         "VARCHAR"),
    ("share_class",        "VARCHAR"),
    ("owner_country",      "VARCHAR"),
    ("owner_zip_place",    "VARCHAR"),
    ("shares_owner_num",   "DOUBLE"),
    ("shares_company_num", "DOUBLE"),
    ("ownership_pct",      "DOUBLE"),
]

def open_conn() -> duckdb.DuckDBPyConnection:
    return duckdb.connect(S.DB_PATH, read_only=False)

# ---------- header-normalisering ----------
def _norm_header(h: str) -> str:
    return h.strip().strip('"').strip("'").strip()

def _resolve_headers(headers: List[str], wanted: Dict[str, str]) -> Dict[str, str]:
    """Match ønskede headere mot faktiske (robust mot ekstra \" i start/slutt)."""
    norm_map = {_norm_header(h): h for h in headers}
    out: Dict[str, str] = {}
    for k, v in wanted.items():
        out[k] = norm_map.get(_norm_header(v), v)
    return out

# ---------- SQL-hjelpere ----------
def _clean_number_expr(col_sql_quoted: str) -> str:
    """Tall: strip anførselstegn, fjern space/NBSP/smal NBSP/punktum/komma → TRY_CAST DOUBLE."""
    as_str = f"TRIM(BOTH '\"' FROM CAST({col_sql_quoted} AS VARCHAR))"
    expr = (
        f"REPLACE(REPLACE(REPLACE(REPLACE(REPLACE("
        f"{as_str}, ' ', ''), '\u00A0', ''), '\u202F', ''), '.', ''), ',', '')"
    )
    return f"TRY_CAST({expr} AS DOUBLE)"

def _build_read_csv(opts: Dict[str, Optional[str]]) -> Tuple[str, List[object]]:
    parts = ["?"]; params: List[object] = [None]  # csv_path settes senere
    if opts.get("delim"): parts.append(f"delim='{sql_lit(opts['delim'])}'")
    parts += ["header=true", "union_by_name=true", "sample_size=-1", "max_line_size=10000000",
              f"strict_mode={'true' if opts.get('strict', True) else 'false'}"]
    if not opts.get("strict", True): parts.append("ignore_errors=true")
    if opts.get("encoding"): parts.append(f"encoding='{sql_lit(opts['encoding'])}'")
    if opts.get("quote") is not None:  parts.append(f"quote='{sql_lit(opts['quote'])}'")
    if opts.get("escape") is not None: parts.append(f"escape='{sql_lit(opts['escape'])}'")
    return "read_csv_auto(" + ", ".join(parts) + ")", params

def _expected_columns() -> Set[str]:
    return {name for name, _ in SCHEMA_COLS}

def _existing_columns(conn: duckdb.DuckDBPyConnection) -> Set[str]:
    rows = conn.execute(
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema=current_schema() AND table_name='shareholders' "
        "ORDER BY ordinal_position"
    ).fetchall()
    return {r[0] for r in rows}

# ---------- bygg / ensure ----------
def ensure_db(csv_path: str,
              db_path: str,
              delimiter: str = S.DELIMITER,
              column_map: Optional[Dict[str, str]] = None,
              force: bool = False) -> None:
    """
    Bygg/oppdater DB. Lager skjema eksplisitt (10 kolonner) og fyller det via INSERT SELECT.
    Backfiller 'Antall aksjer selskap' per company_orgnr (vindu) og beregner ownership_pct ved behov.
    """
    if column_map is None:
        column_map = S.COLUMN_MAP
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Fant ikke CSV: {csv_path}")

    # 1) Detekter dialekt + header
    opts = detect_csv_options(csv_path, delimiter)
    headers = read_headers(csv_path, opts)

    # 2) Map ønskede headere til faktiske (robust mot "…")
    colmap = _resolve_headers(headers, column_map)
    num_src = _resolve_headers(headers, {
        "shares_owner_num": S.COUNT_COLUMNS["shares_owner"],
        "shares_company_num": S.COUNT_COLUMNS["shares_company"],
    })

    needs_compute = colmap.get("ownership_pct") == "__COMPUTE_FROM_COUNTS__"
    required = [colmap["company_orgnr"], colmap["company_name"], colmap["owner_orgnr"], colmap["owner_name"]]
    if needs_compute:
        required += [num_src["shares_owner_num"], num_src["shares_company_num"]]
    else:
        required.append(colmap["ownership_pct"])

    missing = [h for h in required if _norm_header(h) not in {_norm_header(x) for x in headers}]
    if missing:
        raise ValueError(
            "Kolonnemapping matcher ikke CSV‑headerne.\n\n"
            + "Mangler: " + ", ".join(missing) + "\n\n"
            + "Tilgjengelige headere: " + ", ".join(headers[:40]) + (" …" if len(headers) > 40 else "")
        )

    # 3) Trenger vi rebuild?
    meta = S.load_meta()
    csv_mtime = os.path.getmtime(csv_path)
    changed = meta.get("csv_path") != csv_path or meta.get("csv_mtime") != csv_mtime

    con = duckdb.connect(db_path)
    try:
        have_cols = _existing_columns(con)
        needs_schema_upgrade = not _expected_columns().issubset(have_cols)
        if force or changed or not have_cols or needs_schema_upgrade:
            # 3a) Drop & lag tomt skjema eksplisitt
            con.execute("DROP TABLE IF EXISTS shareholders")
            cols_sql = ", ".join(f"{n} {t}" for n, t in SCHEMA_COLS)
            con.execute(f"CREATE TABLE shareholders ({cols_sql})")

            # 3b) Bygg SELECT vi setter inn
            read_csv_sql, params = _build_read_csv(opts); params[0] = csv_path

            comp_org = f"TRIM(BOTH '\"' FROM TRIM(CAST({duck_quote(colmap['company_orgnr'])} AS VARCHAR)))"
            comp_nam = f"TRIM(BOTH '\"' FROM TRIM(CAST({duck_quote(colmap['company_name'])}  AS VARCHAR)))"
            ownr_org = f"NULLIF(TRIM(BOTH '\"' FROM TRIM(CAST({duck_quote(colmap['owner_orgnr'])}  AS VARCHAR))), '')"
            ownr_nam = f"NULLIF(TRIM(BOTH '\"' FROM TRIM(CAST({duck_quote(colmap['owner_name'])}  AS VARCHAR))), '')"

            # Tekstkolonner
            txt_parts: List[str] = []
            for out_name, csv_header in EXTRA_TEXT_COLUMNS.items():
                if _norm_header(csv_header) in {_norm_header(h) for h in headers}:
                    src = duck_quote(_resolve_headers(headers, {out_name: csv_header})[out_name])
                    txt_parts.append(f" NULLIF(TRIM(BOTH '\"' FROM TRIM(CAST({src} AS VARCHAR)))), '') AS {out_name},")
                else:
                    txt_parts.append(f" CAST(NULL AS VARCHAR) AS {out_name},")

            # Tallfelter + backfill totalsum per selskap
            own_num = _clean_number_expr(duck_quote(num_src["shares_owner_num"]))
            tot_raw = _clean_number_expr(duck_quote(num_src["shares_company_num"]))
            tot_num = f"COALESCE({tot_raw}, MAX({tot_raw}) OVER (PARTITION BY {comp_org}))"

            # Prosent
            if needs_compute:
                pct_expr = f"CASE WHEN {tot_num} IS NULL OR {tot_num}=0 THEN NULL ELSE ({own_num}/{tot_num})*100 END"
            else:
                pct_col = duck_quote(colmap["ownership_pct"])
                pct_str = f"TRIM(BOTH '\"' FROM CAST({pct_col} AS VARCHAR))"
                pct_expr = f"TRY_CAST(REPLACE({pct_str}, ',', '.') AS DOUBLE)"

            insert_sql = (
                "INSERT INTO shareholders "
                "SELECT "
                f" {comp_org} AS company_orgnr,"
                f" {comp_nam} AS company_name,"
                f" {ownr_org} AS owner_orgnr,"
                f" {ownr_nam} AS owner_name,"
                + "".join(txt_parts) +
                f" {own_num} AS shares_owner_num,"
                f" {tot_num} AS shares_company_num,"
                f" {pct_expr} AS ownership_pct "
                f"FROM {read_csv_sql}"
            )
            con.execute(insert_sql, params)

            # 3c) Indekser & metadata
            con.execute("CREATE INDEX IF NOT EXISTS idx_sh_company ON shareholders(company_orgnr)")
            con.execute("CREATE INDEX IF NOT EXISTS idx_sh_owner   ON shareholders(owner_orgnr)")
            con.execute("VACUUM")
            meta.update({
                "csv_path": csv_path,
                "csv_mtime": csv_mtime,
                "column_map": colmap,
                "delimiter": opts.get("delim", delimiter),
                "encoding": opts.get("encoding"),
                "quote": opts.get("quote"),
                "escape": opts.get("escape"),
                "strict": opts.get("strict", True),
            })
            S.save_meta(meta)
    finally:
        con.close()

# ---------- SELECT‑hjelpere ----------
def _col_or_null(have: Set[str], name: str, sqltype: str) -> str:
    return name if name in have else f"CAST(NULL AS {sqltype}) AS {name}"

def _expr_or_null(have: Set[str], name: str, sqltype: str) -> str:
    return name if name in have else f"CAST(NULL AS {sqltype})"

# ---------- Queries brukt av GUI/graf ----------
def list_columns(conn: duckdb.DuckDBPyConnection) -> List[str]:
    rows = conn.execute(
        "SELECT column_name FROM information_schema.columns "
        "WHERE table_schema=current_schema() AND table_name='shareholders' "
        "ORDER BY ordinal_position"
    ).fetchall()
    return [r[0] for r in rows]

def search_companies(conn: duckdb.DuckDBPyConnection, term: str, by: str, limit: int = 200):
    if by == "orgnr":
        sql = ("SELECT DISTINCT company_orgnr, company_name "
               "FROM shareholders WHERE company_orgnr LIKE ? ORDER BY company_orgnr LIMIT ?")
        params = [f"%{term}%", limit]
    else:
        sql = ("SELECT DISTINCT company_orgnr, company_name "
               "FROM shareholders WHERE company_name ILIKE ? ORDER BY company_name LIMIT ?")
        params = [f"%{term}%", limit]
    return conn.execute(sql, params).fetchall()

def get_owners_full(conn: duckdb.DuckDBPyConnection, company_orgnr: str):
    """Radene GUI viser – robust mot manglende kolonner."""
    have = _existing_columns(conn)
    select_parts = [
        _col_or_null(have, "owner_orgnr","VARCHAR"),
        _col_or_null(have, "owner_name","VARCHAR"),
        _col_or_null(have, "share_class","VARCHAR"),
        _col_or_null(have, "owner_country","VARCHAR"),
        _col_or_null(have, "owner_zip_place","VARCHAR"),
        _col_or_null(have, "shares_owner_num","DOUBLE"),
        _col_or_null(have, "shares_company_num","DOUBLE"),
        _col_or_null(have, "ownership_pct","DOUBLE"),
    ]
    sql = ("SELECT " + ", ".join(select_parts) +
           " FROM shareholders WHERE company_orgnr = ? "
           "ORDER BY (ownership_pct IS NULL), ownership_pct DESC, owner_name")
    return conn.execute(sql, [company_orgnr]).fetchall()

def get_owners_agg_owner(conn: duckdb.DuckDBPyConnection, company_orgnr: str):
    """SUM per eier (uavhengig av aksjeklasse) – for graf oppstrøms."""
    have = _existing_columns(conn)
    own = _expr_or_null(have, "shares_owner_num", "DOUBLE")
    tot = _expr_or_null(have, "shares_company_num", "DOUBLE")
    sql = (
        "SELECT owner_orgnr, owner_name, "
        "       CAST(NULL AS VARCHAR) AS share_class, "
        "       CAST(NULL AS VARCHAR) AS owner_country, "
        "       CAST(NULL AS VARCHAR) AS owner_zip_place, "
        f"       SUM({own}) AS shares_owner_num, "
        f"       MAX({tot}) AS shares_company_num, "
        f"       CASE WHEN MAX({tot}) IS NULL OR MAX({tot})=0 THEN NULL "
        f"            ELSE SUM({own})/MAX({tot})*100 END AS ownership_pct "
        "FROM shareholders WHERE company_orgnr=? "
        "GROUP BY owner_orgnr, owner_name "
        "ORDER BY (ownership_pct IS NULL), ownership_pct DESC, owner_name"
    )
    return conn.execute(sql, [company_orgnr]).fetchall()

def get_children_agg_company(conn: duckdb.DuckDBPyConnection, owner_orgnr: str):
    """SUM per datterselskap (nedstrøms)."""
    have = _existing_columns(conn)
    own = _expr_or_null(have, "shares_owner_num", "DOUBLE")
    tot = _expr_or_null(have, "shares_company_num", "DOUBLE")
    sql = (
        "SELECT company_orgnr, company_name, shares_owner_num, shares_company_num, ownership_pct "
        "FROM ( "
        "  SELECT company_orgnr, MAX(company_name) AS company_name, "
        f"         SUM({own}) AS shares_owner_num, "
        f"         MAX({tot}) AS shares_company_num, "
        f"         CASE WHEN MAX({tot}) IS NULL OR MAX({tot})=0 THEN NULL "
        f"              ELSE SUM({own})/MAX({tot})*100 END AS ownership_pct "
        "  FROM shareholders WHERE owner_orgnr=? "
        "  GROUP BY company_orgnr "
        ") t "
        "ORDER BY (ownership_pct IS NULL), ownership_pct DESC, company_name"
    )
    return conn.execute(sql, [owner_orgnr]).fetchall()
