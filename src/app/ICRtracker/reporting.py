# -*- coding: utf-8 -*-
"""
reporting.py
Enkle rapport- og loggefunksjoner (CSV + valgfri SQLite-historikk).
"""
from __future__ import annotations

import csv
import sqlite3
from pathlib import Path
from typing import Iterable, Mapping

def write_csv(path: Path, rows: Iterable[Mapping[str,object]], fieldnames: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        w.writeheader()
        for r in rows:
            w.writerow({k: r.get(k, "") for k in fieldnames})

# -- SQLite auditlog (valgfri)

def open_audit(db_path: Path) -> sqlite3.Connection:
    db_path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.executescript("""
    CREATE TABLE IF NOT EXISTS findings (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        ts DATETIME DEFAULT CURRENT_TIMESTAMP,
        source TEXT,
        client_orgnr TEXT,
        client_name TEXT,
        direction TEXT,
        related_orgnr TEXT,
        related_name TEXT,
        related_type TEXT,
        stake_percent REAL,
        shares INTEGER,
        company_orgnr TEXT,
        company_name TEXT,
        fuzzy_score INTEGER,
        flag_client_crosshit INTEGER
    );
    """)
    return conn

def log_findings(conn: sqlite3.Connection, rows: Iterable[Mapping[str,object]], source: str) -> None:
    sql = """INSERT INTO findings
        (source, client_orgnr, client_name, direction, related_orgnr, related_name, related_type,
         stake_percent, shares, company_orgnr, company_name, fuzzy_score, flag_client_crosshit)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)"""
    data = []
    for r in rows:
        data.append((
            source,
            r.get("client_orgnr",""), r.get("client_name",""), r.get("direction",""),
            r.get("related_orgnr",""), r.get("related_name",""), r.get("related_type",""),
            r.get("stake_percent"), r.get("shares"),
            r.get("company_orgnr",""), r.get("company_name",""),
            r.get("fuzzy_score"), 1 if r.get("flag_client_crosshit") else 0
        ))
    if data:
        conn.executemany(sql, data)
        conn.commit()
