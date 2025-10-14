# -*- coding: utf-8 -*-
"""
db_schema_tools.py
Kjør denne (Run) for å liste tabeller/kolonner i en SQLite-DB,
slik at du enkelt kan fylle inn riktige feltnavn i db_compat_adapter.py.
"""
from __future__ import annotations
import sqlite3
from pathlib import Path

DB_PATH = Path(r"F:\Dokument\Kildefiler\aksjonarregister.db")  # ← sett riktig

def main():
    conn = sqlite3.connect(str(DB_PATH))
    conn.row_factory = sqlite3.Row
    cur = conn.execute("SELECT name, type FROM sqlite_master WHERE type IN ('table','view') ORDER BY name;")
    rows = cur.fetchall()
    print(f"Tables/views in {DB_PATH}:")
    for r in rows:
        name = r["name"]
        print(f" - {r['type']:5}  {name}")
        try:
            cols = conn.execute(f"PRAGMA table_info('{name}');").fetchall()
            if cols:
                print("     columns:", ", ".join(c[1] for c in cols))
        except Exception:
            pass

if __name__ == "__main__":
    main()
