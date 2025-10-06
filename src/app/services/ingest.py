# -*- coding: utf-8 -*-
from __future__ import annotations
import os, sys, json, re, hashlib, datetime as dt
from pathlib import Path
from typing import Tuple, Dict, Any, Optional

import pandas as pd

# Stioppsett: støtt både "app.services.*" og "services.*"
SRC = Path(__file__).resolve().parents[2]
if str(SRC) not in sys.path:
    sys.path.insert(0, str(SRC))

def _imports():
    try:
        from app.services.clients import load_meta
        from app.services.versioning import resolve_active_raw_file
        from app.services.io import read_raw
        from app.services.mapping import ensure_mapping_interactive, standardize_with_mapping
        from app.services.regnskapslinjer import try_map_saldobalanse_to_regnskapslinjer
    except Exception:
        from services.clients import load_meta                          # type: ignore
        from services.versioning import resolve_active_raw_file          # type: ignore
        from services.io import read_raw                                 # type: ignore
        from services.mapping import ensure_mapping_interactive, standardize_with_mapping  # type: ignore
        from services.regnskapslinjer import try_map_saldobalanse_to_regnskapslinjer      # type: ignore
    return (load_meta, resolve_active_raw_file, read_raw,
            ensure_mapping_interactive, standardize_with_mapping,
            try_map_saldobalanse_to_regnskapslinjer)

(load_meta, resolve_active_raw_file, read_raw,
 ensure_mapping_interactive, standardize_with_mapping,
 try_map_saldobalanse_to_regnskapslinjer) = _imports()

# ------------------------- hjelpere -------------------------

_CANON_PRIORITY = {
    "konto":     ["kontonr","kontonummer","konto nr","konto","accountno","account no","account"],
    "kontonavn": ["kontonavn","kontonamn","kontotekst","account name","accountname"],
    "dato":      ["dato","bokføringsdato","post date","postdate","transdate","date"],
    "bilagsnr":  ["bilagsnr","bilagsnummer","bilag","voucher","voucher nr","voucherno","doknr","document no"],
    "tekst":     ["tekst","beskrivelse","description","post text","mottaker","faktura","narrative"],
}

def _find_col_priority(df: pd.DataFrame, key: str) -> Optional[str]:
    for name in _CANON_PRIORITY.get(key, []):
        for c in df.columns:
            if c.casefold() == name.casefold():
                return c
    return None

def _bnr_key(x) -> Optional[str]:
    if x is None: return None
    return re.sub(r"[^0-9a-z]+", "", str(x).lower()) or None

def _konto_key_series(s: pd.Series) -> pd.Series:
    if pd.api.types.is_numeric_dtype(s):
        return pd.to_numeric(s, errors="coerce").round(0).astype("Int64").astype(str)
    return s.astype(str).str.replace(r"\D", "", regex=True)

def _choose_konto_kontonavn(df: pd.DataFrame) -> pd.DataFrame:
    """
    Sikrer at 'konto' er kontonummer (sifre) og 'kontonavn' er tekst.
    Håndterer tilfeller der det finnes både 'kontonr' og en tekstlig 'konto'.
    """
    k_num = _find_col_priority(df, "konto")
    k_txt = _find_col_priority(df, "kontonavn")

    existing_konto = None
    for c in df.columns:
        if c.casefold() == "konto":
            existing_konto = c; break

    if k_num and existing_konto and k_num != existing_konto:
        s = df[existing_konto].astype(str)
        if s.str.fullmatch(r"\d+").mean() < 0.5:
            if "kontonavn" not in df.columns:
                df = df.rename(columns={existing_konto: "kontonavn"})
            else:
                df = df.rename(columns={existing_konto: "konto_tekst"})
            df = df.rename(columns={k_num: "konto"})
            return df

    if k_num and "konto" not in df.columns:
        df = df.rename(columns={k_num: "konto"})

    if "konto" in df.columns and "kontonavn" not in df.columns and k_txt and k_txt != "konto":
        df = df.rename(columns={k_txt: "kontonavn"})

    return df

def _remove_summary_rows(df: pd.DataFrame) -> pd.DataFrame:
    """Fjern rader som ser ut som summer/«Totalt beløp»."""
    txt_cols = [c for c in ("konto","kontonavn","tekst") if c in df.columns]
    if not txt_cols: return df
    patt = re.compile(r"(totalt\s*bel|^sum( |$))", re.IGNORECASE)
    mask = pd.Series(True, index=df.index)
    for c in txt_cols:
        s = df[c].astype(str)
        bad = s.str.contains(patt, na=False, regex=True)
        if "konto" in df.columns:
            kontostr = df["konto"].astype(str)
            bad = bad & ~kontostr.str.fullmatch(r"\d+", na=False)
        mask &= ~bad
    return df[mask].reset_index(drop=True)

def _sha256_file(p: Path, block: int = 1 << 20) -> str:
    h = hashlib.sha256()
    with p.open("rb") as f:
        while True:
            b = f.read(block)
            if not b: break
            h.update(b)
    return h.hexdigest()

def _sha256_json(o: Any) -> str:
    s = json.dumps(o, sort_keys=True, default=str, ensure_ascii=False).encode("utf-8")
    return hashlib.sha256(s).hexdigest()

def _dataset_paths(version_dir: Path, source: str) -> Tuple[Path, Path]:
    processed = version_dir / "processed"
    processed.mkdir(parents=True, exist_ok=True)
    return processed / f"{source}.parquet", processed / f"{source}.manifest.json"

def _write_dataset(df: pd.DataFrame, out: Path) -> Tuple[Path, str]:
    """Forsøk Parquet, fallback til pickle hvis pyarrow mangler."""
    try:
        df.to_parquet(out, index=False)
        return out, "parquet"
    except Exception:
        out_pkl = out.with_suffix(".pkl")
        df.to_pickle(out_pkl)
        return out_pkl, "pickle"

def _read_dataset(p: Path) -> pd.DataFrame:
    if p.suffix == ".parquet":
        return pd.read_parquet(p)
    if p.suffix == ".pkl":
        return pd.read_pickle(p)
    raise ValueError(f"Ukjent datasettformat: {p}")

def _canonicalize(df: pd.DataFrame, source: str) -> pd.DataFrame:
    df = _choose_konto_kontonavn(df)

    if source == "saldobalanse":
        if "endring" not in df.columns and {"inngående balanse","utgående balanse"} <= set(df.columns):
            ib = pd.to_numeric(df["inngående balanse"], errors="coerce")
            ub = pd.to_numeric(df["utgående balanse"], errors="coerce")
            df["endring"] = ub - ib
        first = [c for c in ["konto","kontonavn","inngående balanse","endring","utgående balanse"] if c in df.columns]
        df = df[first + [c for c in df.columns if c not in first]]
    else:  # hovedbok
        if "bilagsnr" in df.columns:
            df["__bnr_key__"] = df["bilagsnr"].map(_bnr_key)
        if "dato" in df.columns:
            try: df = df.sort_values("dato").reset_index(drop=True)
            except Exception: pass

    return df

# ------------------------- Hoved-API -------------------------

def ensure_parquet_fresh(parent,
                         root_dir: Path,
                         client: str,
                         year: int,
                         source: str,  # "hovedbok" | "saldobalanse"
                         vtype: str    # "ao" | "interim" | "versjon"
                         ) -> Tuple[Path, Dict[str, Any]]:
    """
    Returnerer (dataset_path, manifest). Regenererer datasett når råfil eller mapping er endret.
    """
    meta = load_meta(root_dir, client)
    raw_path = resolve_active_raw_file(root_dir, client, year, source, vtype, meta)
    if not raw_path:
        raise FileNotFoundError(f"Ingen aktiv versjon for {source}/{vtype} {year}.")

    version_dir = Path(raw_path).parents[1]  # …/vYYYY…/
    dataset_path, manifest_path = _dataset_paths(version_dir, source)

    raw_sha = _sha256_file(Path(raw_path))

    # Hent mapping (viser dialog første gang) – bruk preview for mapping
    df_prev, _ = read_raw(Path(raw_path))
    mapping = ensure_mapping_interactive(parent, root_dir, client, year, source, df_prev.head(1000))
    mapping_sha = _sha256_json(mapping)

    # Hvis manifest finnes og matcher, bruk det
    if manifest_path.exists():
        try:
            old = json.loads(manifest_path.read_text(encoding="utf-8"))
            if old.get("raw_sha256") == raw_sha and old.get("mapping_sha256") == mapping_sha:
                data_path = Path(old.get("dataset_path", str(dataset_path)))
                if data_path.exists():
                    return data_path, old
        except Exception:
            pass  # fall through

    # Regenerer datasett
    df_raw, _ = read_raw(Path(raw_path))
    df_std = standardize_with_mapping(df_raw, mapping=mapping,
                                      parse_dates=True,
                                      numeric_fields=("beløp","mvabeløp","inngående balanse","utgående balanse"))
    df_std = _canonicalize(_remove_summary_rows(df_std), source)

    out_path, fmt = _write_dataset(df_std, dataset_path)

    # ---------------- REGNSKAPSLINJER (kun for saldobalanse) ----------------
    regn_meta: Dict[str, Any] = {}
    if source == "saldobalanse":
        rows_df, agg_df, regn_meta = try_map_saldobalanse_to_regnskapslinjer(df_std)
        if rows_df is not None and agg_df is not None:
            # skriv side-datasett
            side_dir = dataset_path.parent
            rows_out, rows_fmt = _write_dataset(rows_df, side_dir / "saldobalanse__regnskapslinjer_rows.parquet")
            agg_out,  agg_fmt  = _write_dataset(agg_df,  side_dir / "saldobalanse__regnskapslinjer_agg.parquet")
            regn_meta.update({
                "rows_path": str(rows_out),
                "rows_format": rows_fmt,
                "agg_path": str(agg_out),
                "agg_format": agg_fmt,
            })

    # Statistikk/manifest
    nrows, ncols = int(df_std.shape[0]), int(df_std.shape[1])
    uniq_konto = None
    if "konto" in df_std.columns:
        uniq_konto = int(_konto_key_series(df_std["konto"]).nunique(dropna=True))

    first_date = last_date = None
    if "dato" in df_std.columns:
        try:
            s = pd.to_datetime(df_std["dato"], errors="coerce")
            first_date = pd.to_datetime(s.min()).date().isoformat()
            last_date  = pd.to_datetime(s.max()).date().isoformat()
        except Exception:
            pass

    mani = {
        "schema_version": 1,
        "dataset": source,
        "created_at": dt.datetime.now().isoformat(timespec="seconds"),
        "dataset_path": str(out_path),
        "format": fmt,
        "raw_file": str(raw_path),
        "raw_sha256": raw_sha,
        "mapping_sha256": mapping_sha,
        "row_count": nrows,
        "col_count": ncols,
        "unique_accounts": uniq_konto,
        "first_date": first_date,
        "last_date": last_date,
        "columns": list(df_std.columns),
        "dtypes": {c: str(df_std[c].dtype) for c in df_std.columns},
    }
    if regn_meta:
        mani["regnskapslinjer"] = regn_meta

    manifest_path.write_text(json.dumps(mani, ensure_ascii=False, indent=2), encoding="utf-8")
    return out_path, mani

def load_canonical_dataset(dataset_path: Path) -> pd.DataFrame:
    return _read_dataset(Path(dataset_path))


# Valgfri CLI for backfill
if __name__ == "__main__":
    import argparse
    from app.services.clients import get_clients_root  # type: ignore
    p = argparse.ArgumentParser()
    p.add_argument("--client", required=True)
    p.add_argument("--year", type=int, required=True)
    p.add_argument("--source", choices=["hovedbok","saldobalanse"], required=True)
    p.add_argument("--type", dest="vtype", choices=["ao","interim","versjon"], required=True)
    a = p.parse_args()
    root = get_clients_root()
    ds_path, mani = ensure_parquet_fresh(None, root, a.client, a.year, a.source, a.vtype)
    print("OK", ds_path, "rows:", mani.get("row_count"))
