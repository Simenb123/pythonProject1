# -*- coding: utf-8 -*-
"""
scan_registry.py
Kommandolinjeverktøy for å:
  - (valgfritt) importere AR CSV -> DB via registry_db
  - skanne klientliste mot registeret (via matcher + ar_bridge)
  - skrive CSV-rapport + (valgfri) SQLite auditlog

Kjørbar direkte (Run i IDE) ELLER via -m.
"""
from __future__ import annotations

# ====== BOOTSTRAP så fila kan kjøres direkte ======
if __name__ == "__main__" and __package__ is None:
    import sys, pathlib
    here = pathlib.Path(__file__).resolve()
    for up in range(2, 6):
        cand = here.parents[up] / "src"
        if cand.exists():
            sys.path.insert(0, str(cand))
            __package__ = "app.ICRtracker"
            break
# ==================================================

import argparse
from pathlib import Path

# --- AR-kobling (bro først, fallback til registry_db) ---
try:
    from .ar_bridge import open_db  # bruker matcher for resten
    print("scan_registry: ar_bridge aktiv (aksjonaerregister).")
except Exception:
    from .registry_db import open_db  # type: ignore
    from .registry_db import import_csv_to_db  # type: ignore
    print("scan_registry: registry_db (standard).")

from .matcher import load_clients, scan_all_clients
from .reporting import write_csv, open_audit, log_findings

# ---------- Standardverdier (så du kan trykke Run uten arguments) ----------
DEFAULT_DB       = r"F:\Dokument\Kildefiler\aksjonarregister.db"  # For registry_db-varianten (ikke brukt av ar_bridge)
DEFAULT_CLIENTS  = r"F:\Dokument\Kildefiler\BHL AS klientliste - kopi.xlsx"
DEFAULT_OUT      = r"F:\Dokument\Kildefiler\irc\rapporter\ar_match_report.csv"
DEFAULT_AUDIT_DB = ""  # f.eks. r"F:\Dokument\Kildefiler\irc\rapporter\audit_findings.db"

FIELDS = [
    "client_orgnr","client_name","direction","related_orgnr","related_name","related_type",
    "stake_percent","shares","company_orgnr","company_name","fuzzy_score","flag_client_crosshit"
]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--import-csv", help="Rå CSV for registry_db (ikke brukt av ar_bridge)")
    ap.add_argument("--db", default=DEFAULT_DB, help="SQLite databasefil (registry_db). Ikke nødvendig med ar_bridge.")
    ap.add_argument("--clients", default=DEFAULT_CLIENTS, help="Klientliste (xlsx/csv)")
    ap.add_argument("--out", default=DEFAULT_OUT, help="CSV-rapport ut")
    ap.add_argument("--min-name-score", type=int, default=90)
    ap.add_argument("--audit-db", default=DEFAULT_AUDIT_DB, help="SQLite for historikk/rapportlogg")
    args = ap.parse_args()

    db_path = Path(args.db)

    # 1) (Valgfritt) Importer CSV -> DB (kun hvis du bruker registry_db-varianten)
    if args.import_csv:
        try:
            import_csv_to_db  # type: ignore[attr-defined]
        except Exception:
            print("Import ikke tilgjengelig via ar_bridge – hoppet over.")
        else:
            import_csv_to_db(Path(args.import_csv), db_path)  # type: ignore[attr-defined]
            print(f"Import ferdig -> {db_path}")

    # 2) Skann klienter
    if args.clients and args.out:
        clients_path = Path(args.clients)
        out_path     = Path(args.out)

        if not clients_path.exists():
            print(f"ADVARSEL: Fant ikke klientliste: {clients_path}")
            return

        clients = load_clients(clients_path)
        print(f"Klienter lastet: {len(clients)}")

        conn = open_db(db_path)  # ar_bridge ignorerer stien og bruker egen (via settings)
        rows = scan_all_clients(conn, clients, min_name_score=args.min_name_score)
        print(f"Funn: {len(rows)}")

        out_path.parent.mkdir(parents=True, exist_ok=True)
        write_csv(out_path, rows, FIELDS)
        print(f"Rapport skrevet: {out_path}")

        if args.audit_db:
            aconn = open_audit(Path(args.audit_db))
            log_findings(aconn, rows, source="scan_registry")
            print(f"Auditlog ført til: {args.audit_db}")
    else:
        # Ingen clients/out: helsesjekk på DB
        try:
            _ = open_db(db_path)
            print("DB/AR-tilkobling OK.")
            print("Tips: legg inn --clients og --out for å generere rapport.")
        except Exception as e:
            print(f"Kunne ikke åpne DB/AR: {e}")

if __name__ == "__main__":
    main()
