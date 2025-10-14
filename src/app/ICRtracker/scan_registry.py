# -*- coding: utf-8 -*-
"""
scan_registry.py
Kommandolinjeverktøy for å:
  - importere aksjonærregister (CSV) til SQLite
  - skanne klientliste mot registeret
  - skrive CSV-rapport + (valgfri) SQLite auditlog

Kjørbar direkte (Run i IDE) ELLER via -m.
"""

# ====== BOOTSTRAP så fila kan kjøres direkte ======
if __name__ == "__main__" and __package__ is None:
    import sys, pathlib
    here = pathlib.Path(__file__).resolve()
    # prøv å finne <prosjektrot>/src automatisk
    for up in range(2, 6):  # parents[2] .. parents[5]
        cand = here.parents[up] / "src"
        if cand.exists():
            sys.path.insert(0, str(cand))
            __package__ = "app.ICRtracker"
            break
# ==================================================

from __future__ import annotations
import argparse
from pathlib import Path

from .registry_db import import_csv_to_db, open_db
from .matcher import load_clients, scan_all_clients
from .reporting import write_csv, open_audit, log_findings

FIELDS = [
    "client_orgnr","client_name","direction","related_orgnr","related_name","related_type",
    "stake_percent","shares","company_orgnr","company_name","fuzzy_score","flag_client_crosshit"
]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--import-csv", help="Rå CSV fra aksjonærregisteret (stor)")
    ap.add_argument("--db", required=True, help="SQLite databasefil for aksjonærregisteret")
    ap.add_argument("--clients", help="Klientliste (xlsx/csv)")
    ap.add_argument("--out", help="CSV-rapport ut")
    ap.add_argument("--min-name-score", type=int, default=90)
    ap.add_argument("--audit-db", help="SQLite for historikk/rapportlogg")
    args = ap.parse_args()

    db_path = Path(args.db)

    # 1) Importer CSV -> DB (valgfritt)
    if args.import_csv:
        import_csv_to_db(Path(args.import_csv), db_path)
        print(f"Import ferdig -> {db_path}")

    # 2) Skann klienter (valgfritt)
    if args.clients and args.out:
        clients = load_clients(Path(args.clients))
        print(f"Klienter lastet: {len(clients)}")
        conn = open_db(db_path)
        rows = scan_all_clients(conn, clients, min_name_score=args.min_name_score)
        print(f"Funn: {len(rows)}")
        write_csv(Path(args.out), rows, FIELDS)
        print(f"Rapport skrevet: {args.out}")

        if args.audit_db:
            aconn = open_audit(Path(args.audit_db))
            log_findings(aconn, rows, source="scan_registry")
            print(f"Auditlog ført til: {args.audit_db}")

if __name__ == "__main__":
    main()
