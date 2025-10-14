# -*- coding: utf-8 -*-
"""
sanity_import_ar.py
Sjekker om 'app.aksjonaerregister' kan importeres, og gjør en enkel ping.
Kjør denne med Run i PyCharm.
"""
from __future__ import annotations

# --- bootstrap: legg til <prosjektrot>/src på sys.path når du kjører direkte ---
if __name__ == "__main__" and __package__ is None:
    import sys, pathlib
    here = pathlib.Path(__file__).resolve()
    for up in range(2, 7):
        cand = here.parents[up] / "src"
        if cand.exists():
            sys.path.insert(0, str(cand))
            __package__ = "app.ICRtracker"
            break

import importlib.util
import traceback

def main():
    # 1) Finn spesifikasjonen (forteller om modulen kan finnes)
    spec = importlib.util.find_spec("app.aksjonaerregister.db")
    print("find_spec('app.aksjonaerregister.db') ->", spec)

    if spec is None:
        print("\n❌ Finner ikke 'app.aksjonaerregister.db'.")
        print("   Sjekk at:")
        print("   - 'src' er markert som Sources Root i PyCharm")
        print("   - det finnes __init__.py i 'src/app' og 'src/app/aksjonaerregister'")
        print("   - mappen heter nøyaktig 'aksjonaerregister' og filen 'db.py' finnes der")
        return

    # 2) Prøv faktisk import + ping
    try:
        from app.aksjonaerregister import db as ar_db
        print("✅ Import OK:", ar_db.__file__)
        # noen installasjoner trenger duckdb – hvis det mangler, får du ModuleNotFoundError her
    except Exception:
        print("❌ Import feilet:")
        traceback.print_exc()
        return

    # 3) Prøv å åpne en tilkobling (hva open_conn heter i din modul)
    try:
        if hasattr(ar_db, "open_conn"):
            conn = ar_db.open_conn()
            print("✅ open_conn() OK:", type(conn))
            # prøv en helt enkel spørring om mulig (tåler at det ikke er duckdb)
            try:
                res = conn.execute("SELECT 1").fetchone()
                print("Ping DB:", res)
            except Exception as e:
                print("Ping DB: (hopper over) ->", e)
        else:
            print("ℹ️  Fant ikke 'open_conn' i ar_db – sjekk API i din db.py")
    except Exception:
        print("❌ Åpning av tilkobling feilet:")
        traceback.print_exc()

if __name__ == "__main__":
    main()
