# -*- coding: utf-8 -*-
"""
scan_registry.py
Kommandolinjeverktøy for å:
  - importere aksjonærregister (CSV) til SQLite
  - skanne klientliste mot registeret
  - skrive CSV-rapport + (valgfri) SQLite auditlog

Eksempler:
  python -m ICRtracker.scan_registry --import-csv "D:\\data\\aksjonarregister2024.csv" --db "D:\\data\\ar.db"
  python -m ICRtracker.scan_registry --db "D:\\data\\ar.db" --clients "F:\\Dokument\\Kildefiler\\BHL AS klientliste - kopi.xlsx" --out "F:\\Dokument\\Kildefiler\\irc\\rapporter\\ar_match_report.csv"
"""
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
