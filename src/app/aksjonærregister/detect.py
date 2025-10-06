from __future__ import annotations
"""
Robust CSV-detektering og headerlesing.

Endringer:
- Prøver også quote='' (deaktivert sitering) for filer der hele linjer er sitert.
- Sender quote/escape til DuckDB også når verdien er tom streng.
"""
from typing import Dict, List, Optional
import glob
import os
import duckdb

from . import settings as S

# --- Små hjelpere ---
def duck_quote(col: str) -> str:
    """Sikker quoting av kolonnenavn for SQL-strenger."""
    return '"' + col.replace('"', '""') + '"'

def sql_lit(s: str) -> str:
    """Minimal escaping for opsjonsverdier i SQL (enkelsitert)."""
    return s.replace("'", "''")

def latest_csv_in_dir(folder: str, pattern: str = S.CSV_PATTERN) -> Optional[str]:
    try:
        files = sorted(glob.glob(os.path.join(folder, pattern)), key=os.path.getmtime, reverse=True)
        return files[0] if files else None
    except Exception:
        return None

# --- Deteksjon ---
def detect_csv_options(csv_path: str, default_delim: str = S.DELIMITER) -> Dict[str, Optional[str]]:
    """Prøver matrise av dialekter, returnerer vellykket kombinasjon."""
    attempts: List[Dict[str, Optional[str]]] = []
    delims  = [default_delim, ";", ",", "\t", "|"]
    encs    = [None, "utf8", "latin1", "iso-8859-1", "windows-1252"]
    # NB: tom streng "" => deaktiver sitering helt
    quotes  = [None, '"', "'", ""]
    escapes = [None, '"', "'"]

    # Auto først
    attempts.append({"delim": None, "encoding": None, "quote": None, "escape": None, "strict": True})
    # Kombinasjoner (strict)
    for d in delims:
        for e in encs:
            for q in quotes:
                for esc in escapes:
                    attempts.append({"delim": d, "encoding": e, "quote": q, "escape": esc, "strict": True})
    # Tolerant
    for d in delims:
        for e in encs:
            for q in quotes:
                attempts.append({"delim": d, "encoding": e, "quote": q, "escape": None, "strict": False})

    con = duckdb.connect()
    try:
        for op in attempts:
            try:
                parts = ["?"]
                params = [csv_path]
                if op.get("delim") is not None:
                    parts.append(f"delim='{sql_lit(op['delim'])}'")
                parts += [
                    "header=true", "union_by_name=true", "sample_size=-1", "max_line_size=10000000",
                    f"strict_mode={'true' if op.get('strict', True) else 'false'}",
                ]
                if not op.get("strict", True): parts.append("ignore_errors=true")
                if op.get("encoding"): parts.append(f"encoding='{sql_lit(op['encoding'])}'")
                if op.get("quote") is not None:  parts.append(f"quote='{sql_lit(op['quote'])}'")
                if op.get("escape") is not None: parts.append(f"escape='{sql_lit(op['escape'])}'")
                sql = "SELECT * FROM read_csv_auto(" + ", ".join(parts) + ") LIMIT 1"
                con.execute(sql, params).fetchall()
                return op
            except Exception:
                continue
        # Siste utvei
        return {"delim": default_delim, "encoding": "utf8", "quote": "", "escape": None, "strict": False}
    finally:
        con.close()

def read_headers(csv_path: str, opts: Dict[str, Optional[str]]) -> List[str]:
    """Les bare headerne (LIMIT 0) med oppgitte opsjoner."""
    con = duckdb.connect()
    try:
        parts = ["?"]
        params = [csv_path]
        if opts.get("delim"):
            parts.append(f"delim='{sql_lit(opts['delim'])}'")
        parts += ["header=true", "union_by_name=true", "sample_size=2000"]
        if opts.get("encoding"):
            parts.append(f"encoding='{sql_lit(opts['encoding'])}'")
        if opts.get("quote") is not None:
            parts.append(f"quote='{sql_lit(opts['quote'])}'")
        if opts.get("escape") is not None:
            parts.append(f"escape='{sql_lit(opts['escape'])}'")
        sql = "SELECT * FROM read_csv_auto(" + ", ".join(parts) + ") LIMIT 0"
        df = con.execute(sql, params).fetchdf()
        return list(df.columns)
    finally:
        con.close()
