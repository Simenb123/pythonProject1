# -*- coding: utf-8 -*-
"""
scan_registry.py
Kommandolinjeverktøy for å:
  - importere aksjonærregister (CSV) til SQLite (valgfritt)
  - skanne klientliste mot registeret
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

# --- AR-kobling (adapter først, fallback til standard) ---
try:
    from .db_compat_adapter import open_db, get_owners, companies_owned_by, normalize_orgnr  # noqa: F401
    print("Bruker db_compat_adapter (eksisterende AR-DB).")
except Exception:
    from .registry_db import import_csv_to_db, open_db  # type: ignore
    print("Bruker registry_db (standard holdings/companies).")

from .matcher import load_clients, scan_all_clients
from .reporting import write_csv, open_audit, log_findings

# ---------- Standardverdier (så du kan trykke Run uten arguments) ----------
DEFAULT_DB       = r"F:\Dokument\Kildefiler\aksjonarregister.db"
DEFAULT_CLIENTS  = r"F:\Dokument\Kildefiler\BHL AS klientliste - kopi.xlsx"
DEFAULT_OUT      = r"F:\Dokument\Kildefiler\irc\rapporter\ar_match_report.csv"
DEFAULT_AUDIT_DB = ""  # f.eks. r"F:\Dokument\Kildefiler\irc\rapporter\audit_findings.db"

FIELDS = [
    "client_orgnr","client_name","direction","related_orgnr","related_name","related_type",
    "stake_percent","shares","company_orgnr","company_name","fuzzy_score","flag_client_crosshit"
]

def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--import-csv", help="Rå CSV fra aksjonærregisteret (stor)")
    ap.add_argument("--db", default=DEFAULT_DB, help="SQLite databasefil for aksjonærregisteret")
    ap.add_argument("--clients", default=DEFAULT_CLIENTS, help="Klientliste (xlsx/csv)")
    ap.add_argument("--out", default=DEFAULT_OUT, help="CSV-rapport ut")
    ap.add_argument("--min-name-score", type=int, default=90)
    ap.add_argument("--audit-db", default=DEFAULT_AUDIT_DB, help="SQLite for historikk/rapportlogg")
    args = ap.parse_args()

    db_path = Path(args.db)

    # 1) Importer CSV -> DB (valgfritt – kun støttet av registry_db-varianten)
    if args.import_csv:
        try:
            import_csv_to_db  # type: ignore[attr-defined]
        except Exception:
            print("Import ikke tilgjengelig via adapter – hoppet over.")
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

        conn = open_db(db_path)
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
            print(f"DB OK: {db_path}")
            print("Tips: legg inn --clients og --out for å generere rapport.")
        except Exception as e:
            print(f"Kunne ikke åpne DB: {e}")

if __name__ == "__main__":
    main()
