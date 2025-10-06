"""
Aksjonærregister – Tkinter søk + orgkart (DuckDB)

Dette er en robust, kjørbar startapp for å søke i en stor CSV (3–4M rader) og
visualisere eierskap oppstrøms/nedstrøms. Den håndterer:
- CSV-autodeteksjon (delimiter/encoding/quote/escape) med tolerant fallback
- Kolonnemapping til **din header** (Orgnr; Selskap; Navn aksjonær; Fødselsår/orgnr; …)
- Beregner eierandel (%) når kolonnen mangler: Antall aksjer / Antall aksjer selskap
- GUI (Tkinter) når tilgjengelig, ellers CLI-fallback
- Enkle selvtester (python app.py --test)
"""
from __future__ import annotations

import os
import sys
import json
import glob
import tempfile
from dataclasses import dataclass
from typing import List, Optional, Dict, Tuple

import duckdb

# ---------------------
# Betinget GUI/visnings-avhengigheter
# ---------------------
try:
    import tkinter as tk  # type: ignore
    from tkinter import ttk, messagebox, filedialog  # type: ignore
    TK_AVAILABLE = True
except Exception:
    tk = None  # type: ignore
    ttk = messagebox = filedialog = None  # type: ignore
    TK_AVAILABLE = False

try:
    from graphviz import Digraph
except Exception:
    Digraph = None  # type: ignore

try:
    from PIL import Image, ImageTk  # type: ignore
    PIL_AVAILABLE = True
except Exception:
    Image = ImageTk = None  # type: ignore
    PIL_AVAILABLE = False

import webbrowser

# ========================
# KONFIGURASJON
# ========================
SERVER_CSV_DIR = r""  # f.eks. r"\\\\server\\share\\aksjonaerregister"
CSV_PATTERN = "*.csv"

CSV_PATH = "./aksjonaerregister.csv"
DB_PATH = "./aksjonaerregister.duckdb"
DELIMITER = ","  # endre til ";" ved semikolon
META_PATH = "./build_meta.json"

# Kolonnemapping til OPPRINNELIG HEADER du viste
# Orgnr; Selskap; Aksjeklasse; Navn aksjonær; Fødselsår/orgnr; Postnr/sted; Landkode; Antall aksjer; Antall aksjer selskap
COLUMN_MAP: Dict[str, str] = {
    "company_orgnr": "Orgnr",
    "company_name": "Selskap",
    "owner_orgnr": "Fødselsår/orgnr",
    "owner_name": "Navn aksjonær",
    # Mangler eksplisitt % i fila → vi beregner fra counts
    "ownership_pct": "__COMPUTE_FROM_COUNTS__",
}
COUNT_COLUMNS: Dict[str, str] = {
    "shares_owner": "Antall aksjer",
    "shares_company": "Antall aksjer selskap",
}

MAX_DEPTH_UP = 3
MAX_DEPTH_DOWN = 2

# ========================
# DATABASE
# ========================
SCHEMA_SQL = (
    "CREATE TABLE IF NOT EXISTS shareholders AS "
    "SELECT CAST(NULL AS VARCHAR) AS company_orgnr, "
    "CAST(NULL AS VARCHAR) AS company_name, "
    "CAST(NULL AS VARCHAR) AS owner_orgnr, "
    "CAST(NULL AS VARCHAR) AS owner_name, "
    "CAST(NULL AS DOUBLE) AS ownership_pct WHERE FALSE;"
)

@dataclass
class Owner:
    orgnr: Optional[str]
    name: str
    pct: Optional[float]

# ========================
# HJELPERE
# ========================

def _sql_lit(s: str) -> str:
    return s.replace("'", "''")


def duck_quote(col: str) -> str:
    return '"' + col.replace('"', '""') + '"'


def save_meta(meta: dict) -> None:
    try:
        with open(META_PATH, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def load_meta() -> dict:
    try:
        with open(META_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def latest_csv_in_dir(folder: str, pattern: str = CSV_PATTERN) -> Optional[str]:
    try:
        files = sorted(glob.glob(os.path.join(folder, pattern)), key=os.path.getmtime, reverse=True)
        return files[0] if files else None
    except Exception:
        return None


# ========================
# CSV-DETeksjon (robust)
# ========================

def detect_csv_options(csv_path: str, default_delim: str = DELIMITER) -> Dict[str, Optional[str]]:
    attempts: List[Dict[str, Optional[str]]] = []
    delims = [default_delim, ";", ",", "\t", "|"]
    encs = [None, "utf8", "latin1", "iso-8859-1", "windows-1252"]
    quotes = [None, '"', "'"]
    escapes = [None, '"', "'"]
    # Start med auto
    attempts.append({"delim": None, "encoding": None, "quote": None, "escape": None, "strict": True})
    for d in delims:
        for e in encs:
            for q in quotes:
                for esc in escapes:
                    attempts.append({"delim": d, "encoding": e, "quote": q, "escape": esc, "strict": True})
    # Tolerant
    for d in delims:
        for e in encs:
            attempts.append({"delim": d, "encoding": e, "quote": None, "escape": None, "strict": False})

    con = duckdb.connect()
    try:
        for op in attempts:
            try:
                parts = ["?"]
                params = [csv_path]
                if op.get("delim") is not None:
                    parts.append(f"delim='{_sql_lit(op['delim'])}'")
                parts += [
                    "header=true",
                    "union_by_name=true",
                    "sample_size=-1",
                    "max_line_size=10000000",
                    f"strict_mode={'true' if op.get('strict', True) else 'false'}",
                ]
                if not op.get("strict", True):
                    parts.append("ignore_errors=true")
                if op.get("encoding"):
                    parts.append(f"encoding='{_sql_lit(op['encoding'])}'")
                if op.get("quote"):
                    parts.append(f"quote='{_sql_lit(op['quote'])}'")
                if op.get("escape"):
                    parts.append(f"escape='{_sql_lit(op['escape'])}'")
                sql = "SELECT * FROM read_csv_auto(" + ", ".join(parts) + ") LIMIT 1"
                con.execute(sql, params).fetchall()
                return op
            except Exception:
                continue
        return {"delim": None, "encoding": None, "quote": None, "escape": None, "strict": False}
    finally:
        con.close()


def _read_headers(csv_path: str, opts: Dict[str, Optional[str]]) -> List[str]:
    con = duckdb.connect()
    try:
        parts = ["?"]
        params = [csv_path]
        if opts.get("delim"):
            parts.append(f"delim='{_sql_lit(opts['delim'])}'")
        parts += ["header=true", "union_by_name=true", "sample_size=2000"]
        if opts.get("encoding"):
            parts.append(f"encoding='{_sql_lit(opts['encoding'])}'")
        if opts.get("quote"):
            parts.append(f"quote='{_sql_lit(opts['quote'])}'")
        if opts.get("escape"):
            parts.append(f"escape='{_sql_lit(opts['escape'])}'")
        sql = "SELECT * FROM read_csv_auto(" + ", ".join(parts) + ") LIMIT 0"
        df = con.execute(sql, params).fetchdf()
        return list(df.columns)
    finally:
        con.close()


# ========================
# BUILD / ENSURE DB
# ========================

def ensure_db(csv_path: str, db_path: str, delimiter: str = DELIMITER, column_map: Optional[Dict[str, str]] = None) -> None:
    if column_map is None:
        column_map = COLUMN_MAP
    if not os.path.exists(csv_path):
        raise FileNotFoundError(f"Fant ikke CSV: {csv_path}")

    opts = detect_csv_options(csv_path, delimiter)
    headers = _read_headers(csv_path, opts)

    needs_compute = column_map.get("ownership_pct") == "__COMPUTE_FROM_COUNTS__"
    required = [column_map["company_orgnr"], column_map["company_name"], column_map["owner_orgnr"], column_map["owner_name"]]
    if needs_compute:
        required += [COUNT_COLUMNS["shares_owner"], COUNT_COLUMNS["shares_company"]]
    else:
        required.append(column_map["ownership_pct"])

    missing = [h for h in required if h not in headers]
    if missing:
        msg = (
            "Kolonnemapping matcher ikke CSV-headerne.\n\n"
            + "Mangler (forventet→finnes ikke): " + ", ".join(missing) + "\n\n"
            + "Tilgjengelige headere i CSV: " + ", ".join(headers[:40]) + (" …" if len(headers) > 40 else "") + "\n\n"
            + "Løsning: Juster COLUMN_MAP eller bruk veiviser."
        )
        raise ValueError(msg)

    meta = load_meta()
    csv_mtime = os.path.getmtime(csv_path)
    must_build = (not os.path.exists(db_path)) or meta.get("csv_path") != csv_path or meta.get("csv_mtime") != csv_mtime

    con = duckdb.connect(db_path)
    try:
        con.execute(SCHEMA_SQL)
        res = con.execute("SELECT COUNT(*) FROM shareholders").fetchone()[0]
        if must_build or res == 0:
            def build_src(o: Dict[str, Optional[str]]) -> Tuple[str, List[object]]:
                parts = ["?"]
                params: List[object] = [csv_path]
                if o.get("delim"):
                    parts.append(f"delim='{_sql_lit(o['delim'])}'")
                parts += [
                    "header=true",
                    "union_by_name=true",
                    "sample_size=-1",
                    "max_line_size=10000000",
                    f"strict_mode={'true' if o.get('strict', True) else 'false'}",
                ]
                if not o.get("strict", True):
                    parts.append("ignore_errors=true")
                if o.get("encoding"):
                    parts.append(f"encoding='{_sql_lit(o['encoding'])}'")
                if o.get("quote"):
                    parts.append(f"quote='{_sql_lit(o['quote'])}'")
                if o.get("escape"):
                    parts.append(f"escape='{_sql_lit(o['escape'])}'")

                # Pct-uttrykk
                if needs_compute:
                    own = duck_quote(COUNT_COLUMNS["shares_owner"])  # teller
                    tot = duck_quote(COUNT_COLUMNS["shares_company"])  # nevner
                    clean_own = f"TRY_CAST(REPLACE(REPLACE(REPLACE({own}, ' ', ''), '.', ''), ',', '') AS DOUBLE)"
                    clean_tot = f"TRY_CAST(REPLACE(REPLACE(REPLACE({tot}, ' ', ''), '.', ''), ',', '') AS DOUBLE)"
                    pct_expr = f"CASE WHEN {clean_tot} IS NULL OR {clean_tot} = 0 THEN NULL ELSE ({clean_own} / {clean_tot}) * 100 END"
                else:
                    pct_expr = f"TRY_CAST(REPLACE({duck_quote(column_map['ownership_pct'])}, ',', '.') AS DOUBLE)"

                sql = (
                    "SELECT "
                    f" TRIM({duck_quote(column_map['company_orgnr'])}) AS company_orgnr,"
                    f" TRIM({duck_quote(column_map['company_name'])}) AS company_name,"
                    f" NULLIF(TRIM({duck_quote(column_map['owner_orgnr'])}), '') AS owner_orgnr,"
                    f" TRIM({duck_quote(column_map['owner_name'])}) AS owner_name,"
                    f" {pct_expr} AS ownership_pct"
                    " FROM read_csv_auto(" + ", ".join(parts) + ")"
                )
                return sql, params

            try:
                src, p = build_src(opts)
                con.execute("DELETE FROM shareholders")
                con.execute("INSERT INTO shareholders " + src, p)
            except Exception:
                fallback = dict(opts)
                fallback["strict"] = False
                src, p = build_src(fallback)
                con.execute("DELETE FROM shareholders")
                con.execute("INSERT INTO shareholders " + src, p)
            con.execute("VACUUM")
            meta.update({
                "csv_path": csv_path,
                "csv_mtime": csv_mtime,
                "column_map": column_map,
                "delimiter": opts.get("delim", delimiter),
                "encoding": opts.get("encoding"),
                "quote": opts.get("quote"),
                "escape": opts.get("escape"),
                "strict": opts.get("strict", True),
                "count_columns": COUNT_COLUMNS if needs_compute else None,
            })
            save_meta(meta)
    finally:
        con.close()


def open_conn() -> duckdb.DuckDBPyConnection:
    return duckdb.connect(DB_PATH, read_only=False)


# ========================
# QUERY-FUNKSJONER
# ========================

def search_companies(conn: duckdb.DuckDBPyConnection, term: str, by: str, limit: int = 200):
    if by == "orgnr":
        sql = (
            "SELECT DISTINCT company_orgnr, company_name "
            "FROM shareholders WHERE company_orgnr LIKE ? ORDER BY company_orgnr LIMIT ?"
        )
        params = [f"%{term}%", limit]
    else:
        sql = (
            "SELECT DISTINCT company_orgnr, company_name "
            "FROM shareholders WHERE company_name ILIKE ? ORDER BY company_name LIMIT ?"
        )
        params = [f"%{term}%", limit]
    return duckdb.execute(sql, params).fetchall()


def get_owners(conn: duckdb.DuckDBPyConnection, company_orgnr: str) -> List[Owner]:
    sql = (
        "SELECT owner_orgnr, owner_name, ownership_pct "
        "FROM shareholders WHERE company_orgnr = ? "
        "ORDER BY (ownership_pct IS NULL), ownership_pct DESC"
    )
    rows = conn.execute(sql, [company_orgnr]).fetchall()
    return [Owner(orgnr=r[0], name=r[1], pct=r[2]) for r in rows]


def get_subs(conn: duckdb.DuckDBPyConnection, owner_orgnr: str) -> List[Owner]:
    sql = (
        "SELECT company_orgnr, company_name, ownership_pct "
        "FROM shareholders WHERE owner_orgnr = ? "
        "ORDER BY (ownership_pct IS NULL), ownership_pct DESC"
    )
    rows = conn.execute(sql, [owner_orgnr]).fetchall()
    return [Owner(orgnr=r[0], name=r[1], pct=r[2]) for r in rows]


# ========================
# GRAPHVIZ
# ========================

def render_graph(conn: duckdb.DuckDBPyConnection, company_orgnr: str, company_name: str, mode: str = "both",
                 max_up: int = MAX_DEPTH_UP, max_down: int = MAX_DEPTH_DOWN) -> Optional[str]:
    if Digraph is None:
        return None
    g = Digraph("eierskap", format="png")
    g.attr(rankdir="TB", fontsize="10", labelloc="t", label=f"Eierskapstre for\n{company_name} ({company_orgnr})")

    seen = set()

    def node_label(name: str, orgnr: Optional[str]) -> str:
        base = name or "(Ukjent navn)"
        if orgnr:
            base += f"\n{orgnr}"
        return base

    root = f"C:{company_orgnr}"
    g.node(root, node_label(company_name, company_orgnr), shape="box", style="rounded,filled", fillcolor="lightgrey")
    seen.add(root)

    def walk_up(target: str, d: int):
        if d >= max_up:
            return
        for ow in get_owners(conn, target):
            nid = f"U:{ow.orgnr or ow.name}"
            if nid not in seen:
                g.node(nid, node_label(ow.name, ow.orgnr), shape="ellipse")
                seen.add(nid)
            g.edge(nid, f"C:{target}", label=(f"{ow.pct:.2f}%" if isinstance(ow.pct, float) else ""))
            if ow.orgnr:
                walk_up(ow.orgnr, d + 1)

    def walk_down(owner: str, d: int):
        if d >= max_down:
            return
        for sb in get_subs(conn, owner):
            nid = f"D:{sb.orgnr or sb.name}"
            if nid not in seen:
                g.node(nid, node_label(sb.name, sb.orgnr), shape="box", style="rounded")
                seen.add(nid)
            g.edge(f"C:{owner}", nid, label=(f"{sb.pct:.2f}%" if isinstance(sb.pct, float) else ""))
            if sb.orgnr:
                walk_down(sb.orgnr, d + 1)

    if mode in ("up", "both"):
        walk_up(company_orgnr, 0)
    if mode in ("down", "both"):
        walk_down(company_orgnr, 0)

    tmp = tempfile.gettempdir()
    out = os.path.join(tmp, f"org_{company_orgnr}_{mode}.png")
    try:
        g.render(filename=os.path.splitext(out)[0], cleanup=True)
    except Exception:
        return None
    return out


# ========================
# TKINTER-UI (valgfritt)
# ========================
if TK_AVAILABLE:
    class App(tk.Tk):
        def __init__(self) -> None:
            super().__init__()
            self.title("Aksjonærregister – søk og visualisering (DuckDB)")
            self.geometry("1200x780")
            self.conn: Optional[duckdb.DuckDBPyConnection] = None
            self.graph_img = None
            self._build_menu()
            self._build_ui()
            self._open_db()

        def _build_menu(self) -> None:
            m = tk.Menu(self)
            fm = tk.Menu(m, tearoff=0)
            fm.add_command(label="Velg CSV og bygg DB…", command=self._import_csv)
            fm.add_separator()
            fm.add_command(label="Avslutt", command=self.destroy)
            m.add_cascade(label="Fil", menu=fm)
            hm = tk.Menu(m, tearoff=0)
            hm.add_command(label="Graphviz nedlasting", command=lambda: webbrowser.open("https://graphviz.org/download/"))
            m.add_cascade(label="Hjelp", menu=hm)
            self.config(menu=m)

        def _build_ui(self) -> None:
            row = ttk.Frame(self, padding=8); row.pack(fill=tk.X)
            ttk.Label(row, text="Søk:").pack(side=tk.LEFT)
            self.entry = ttk.Entry(row, width=40); self.entry.pack(side=tk.LEFT, padx=6)
            self.entry.bind("<Return>", lambda e: self._do_search())
            self.by = tk.StringVar(value="navn")
            ttk.Radiobutton(row, text="Navn", variable=self.by, value="navn").pack(side=tk.LEFT)
            ttk.Radiobutton(row, text="Orgnr", variable=self.by, value="orgnr").pack(side=tk.LEFT)
            ttk.Button(row, text="Søk", command=self._do_search).pack(side=tk.LEFT, padx=6)

            cols = ("company_orgnr", "company_name")
            self.tree = ttk.Treeview(self, columns=cols, show="headings", height=14)
            self.tree.heading("company_orgnr", text="Orgnr")
            self.tree.heading("company_name", text="Selskap")
            self.tree.column("company_orgnr", width=160, anchor=tk.W)
            self.tree.column("company_name", width=560, anchor=tk.W)
            self.tree.pack(fill=tk.BOTH, expand=False, padx=8, pady=(6, 8))
            self.tree.bind("<<TreeviewSelect>>", self._on_select)

            ctrl = ttk.Frame(self, padding=(8, 0)); ctrl.pack(fill=tk.X)
            self.sel_label = ttk.Label(ctrl, text="Velg et selskap ovenfor…"); self.sel_label.pack(side=tk.LEFT)
            self.mode = tk.StringVar(value="both")
            ttk.Radiobutton(ctrl, text="Oppstrøms", variable=self.mode, value="up").pack(side=tk.RIGHT)
            ttk.Radiobutton(ctrl, text="Nedstrøms", variable=self.mode, value="down").pack(side=tk.RIGHT)
            ttk.Radiobutton(ctrl, text="Begge", variable=self.mode, value="both").pack(side=tk.RIGHT)

            depth = ttk.Frame(self, padding=(8, 0)); depth.pack(fill=tk.X)
            ttk.Label(depth, text="Dybde opp:").pack(side=tk.LEFT)
            self.spin_up = tk.Spinbox(depth, from_=0, to=6, width=3); self.spin_up.delete(0, tk.END); self.spin_up.insert(0, str(MAX_DEPTH_UP)); self.spin_up.pack(side=tk.LEFT, padx=(4,12))
            ttk.Label(depth, text="Dybde ned:").pack(side=tk.LEFT)
            self.spin_down = tk.Spinbox(depth, from_=0, to=6, width=3); self.spin_down.delete(0, tk.END); self.spin_down.insert(0, str(MAX_DEPTH_DOWN)); self.spin_down.pack(side=tk.LEFT, padx=(4,12))
            ttk.Button(depth, text="Vis orgkart", command=self._show_graph).pack(side=tk.RIGHT)

            self.canvas = ttk.Label(self); self.canvas.pack(fill=tk.BOTH, expand=True, padx=8, pady=8)

        def _open_db(self) -> None:
            if not os.path.exists(DB_PATH) and os.path.exists(CSV_PATH):
                try:
                    ensure_db(CSV_PATH, DB_PATH)
                except Exception as e:
                    messagebox.showerror("Importfeil", str(e))
                    return
            try:
                self.conn = open_conn()
            except Exception as e:
                messagebox.showerror("Feil ved åpning", str(e))

        def _import_csv(self) -> None:
            path = filedialog.askopenfilename(title="Velg CSV", filetypes=[("CSV", "*.csv"), ("Alle filer", "*.*")])
            if not path:
                return
            try:
                ensure_db(path, DB_PATH)
                if self.conn:
                    self.conn.close()
                self.conn = open_conn()
                messagebox.showinfo("Ferdig", "Database er bygget og klar!")
            except Exception as e:
                messagebox.showerror("Importfeil", str(e))

        def _do_search(self) -> None:
            if not self.conn:
                messagebox.showwarning("Ingen DB", "Åpne/bygg databasen først.")
                return
            term = self.entry.get().strip()
            rows = search_companies(self.conn, term, self.by.get())
            for i in self.tree.get_children():
                self.tree.delete(i)
            for orgnr, name in rows:
                self.tree.insert("", tk.END, values=(orgnr, name))

        def _on_select(self, _evt=None) -> None:
            sel = self.tree.selection()
            if not sel:
                return
            vals = self.tree.item(sel[0], "values")
            self.sel_label.config(text=f"Valgt: {vals[1]} ({vals[0]})")

        def _show_graph(self) -> None:
            if not self.conn:
                return
            sel = self.tree.selection()
            if not sel:
                messagebox.showinfo("Ingen valgt", "Velg et selskap fra listen først.")
                return
            orgnr, name = self.tree.item(sel[0], "values")
            try:
                up = int(self.spin_up.get()); down = int(self.spin_down.get())
            except Exception:
                up, down = MAX_DEPTH_UP, MAX_DEPTH_DOWN
            path = render_graph(self.conn, orgnr, name, mode=self.mode.get(), max_up=up, max_down=down)
            if not path:
                messagebox.showwarning("Ingen graf", "Kunne ikke generere graf (mangler Graphviz eller feil ved rendering).")
                return
            if not PIL_AVAILABLE:
                messagebox.showinfo("Lagret", f"Graf generert: {path}\n(Pillow ikke installert – viser ikke bilde i appen)")
                return
            try:
                img = Image.open(path)
                max_w = self.winfo_width() - 40; max_h = self.winfo_height() - 300
                w, h = img.size; scale = min(max_w / w, max_h / h, 1.0)
                if scale < 1.0:
                    img = img.resize((int(w * scale), int(h * scale)), Image.LANCZOS)
                self.graph_img = ImageTk.PhotoImage(img)
                self.canvas.configure(image=self.graph_img)
            except Exception as e:
                messagebox.showerror("Visningsfeil", str(e))

else:
    App = None  # type: ignore

# ========================
# CLI-FALLBACK
# ========================

def _print_cli_help() -> None:
    help_txt = (
        "CLI-bruk (uten GUI):\n"
        "  --cli-help\n"
        "  --cli-build --csv <sti> [--delimiter \";\"|\",\"]\n"
        "  --cli-search <term> --by navn|orgnr [--limit N] [--db <sti>]\n"
        "  --cli-graph --orgnr <orgnr> --name <navn> [--mode both|up|down] [--max-up N] [--max-down N] [--db <sti>]\n"
    )
    print(help_txt)


def _argval(args: List[str], key: str, default: Optional[str] = None) -> Optional[str]:
    if key in args:
        i = args.index(key)
        if i + 1 < len(args):
            return args[i + 1]
    return default


def _has(args: List[str], key: str) -> bool:
    return key in args


def _cli_main(argv: List[str]) -> None:
    if _has(argv, "--cli-help") or not any(a.startswith("--cli-") for a in argv):
        _print_cli_help(); return

    if _has(argv, "--cli-build"):
        csv = _argval(argv, "--csv", CSV_PATH)
        delim = _argval(argv, "--delimiter", DELIMITER)
        if not csv:
            print("Mangler --csv <sti>"); return
        ensure_db(csv, DB_PATH, delimiter=delim, column_map=COLUMN_MAP)
        print(f"Bygd DB: {DB_PATH} fra {csv}"); return

    if _has(argv, "--cli-search"):
        term = _argval(argv, "--cli-search", ""); by = _argval(argv, "--by", "navn"); lim = int(_argval(argv, "--limit", "20") or 20)
        db = _argval(argv, "--db", DB_PATH); con = duckdb.connect(db)
        try:
            rows = search_companies(con, term, by, lim)
            for r in rows: print(f"{r[0]}\t{r[1]}")
        finally:
            con.close(); return

    if _has(argv, "--cli-graph"):
        orgnr = _argval(argv, "--orgnr"); name = _argval(argv, "--name", orgnr or ""); mode = _argval(argv, "--mode", "both")
        up = int(_argval(argv, "--max-up", str(MAX_DEPTH_UP)) or MAX_DEPTH_UP); dn = int(_argval(argv, "--max-down", str(MAX_DEPTH_DOWN)) or MAX_DEPTH_DOWN)
        db = _argval(argv, "--db", DB_PATH)
        if not orgnr: print("Mangler --orgnr <orgnr>"); return
        con = duckdb.connect(db)
        try:
            out = render_graph(con, orgnr, name, mode=mode, max_up=up, max_down=dn)
            print(out if out else "Kunne ikke generere graf (Graphviz ikke tilgjengelig?)")
        finally:
            con.close(); return

    _print_cli_help()

# ========================
# SELVTESTER (kalles med --test)
# ========================

def _write_sample_csv(path: str, delim: str = ",") -> None:
    txt = (
        "selskap_orgnr{d}selskap_navn{d}eier_orgnr{d}eier_navn{d}eierandel_prosent\n"
        "910000001{d}Alpha AS{d}910000010{d}Holding ASA{d}60\n"
        "910000001{d}Alpha AS{d}{d}Ola Nordmann{d}40\n"
        "910000010{d}Holding ASA{d}910000020{d}TopHold AS{d}100\n"
        "910000002{d}Beta AS{d}910000010{d}Holding ASA{d}35\n"
    ).format(d=delim)
    with open(path, "w", encoding="utf-8") as f:
        f.write(txt)


def _selftest() -> None:
    print("Kjører selvtester…")
    tmp = tempfile.gettempdir()
    test_csv = os.path.join(tmp, "_test_aksjonarregister.csv")
    test_db = os.path.join(tmp, "_test_aksjonarregister.duckdb")
    if os.path.exists(test_db): os.remove(test_db)
    _write_sample_csv(test_csv, ",")
    ensure_db(test_csv, test_db, delimiter=",", column_map=COLUMN_MAP)
    con = duckdb.connect(test_db)
    try:
        rows = search_companies(con, "Alpha", by="navn", limit=10)
        assert any(r[1] == "Alpha AS" for r in rows)
        rows2 = search_companies(con, "alpha", by="navn", limit=10)
        assert any(r[1] == "Alpha AS" for r in rows2)
        rows = search_companies(con, "910000001", by="orgnr", limit=10)
        assert any(r[0] == "910000001" for r in rows)
        # Graphviz test (hopp hvis ikke installert)
        if Digraph is not None:
            out = render_graph(con, "910000001", "Alpha AS", mode="both", max_up=2, max_down=2)
            assert out and os.path.exists(out)
        print("Alle tester OK!")
    finally:
        con.close()


if __name__ == "__main__":
    args = sys.argv[1:]
    if "--test" in args:
        _selftest()
    else:
        if TK_AVAILABLE:
            app = App(); app.mainloop()
        else:
            print(
                "Tkinter er ikke installert i dette miljøet.\n"
                "- Kjør `python app.py --cli-help` for CLI-bruk uten GUI, eller\n"
                "- Installer Tkinter og kjør `python app.py` for GUI."
            )
            if any(a.startswith("--cli-") for a in args):
                _cli_main(args)
            # ellers avslutter vi stille
