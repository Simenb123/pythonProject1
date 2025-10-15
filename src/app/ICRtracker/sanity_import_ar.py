# -*- coding: utf-8 -*-
"""
sanity_import_ar.py
Kjør denne med Run. Den bekrefter at 'app.aksjonærregister' kan importeres,
og at vi faktisk får en tilkobling fra AR-modulen (DuckDB).
"""
from __future__ import annotations

# --- Robust bootstrap: legg mappen som inneholder "app" på sys.path ---
if __name__ == "__main__" and __package__ is None:
    import sys, pathlib
    here = pathlib.Path(__file__).resolve()
    # Gå oppover og finn en mappe som har 'app/__init__.py'
    for parent in (here.parent, *here.parents):
        if (parent / "app" / "__init__.py").exists():
            sys.path.insert(0, str(parent))
            break

import importlib.util
import traceback
import sys

def main():
    print("sys.path[0] =", sys.path[0])

    # Sjekk både 'aksjonærregister' og ascii-fallback 'aksjonaerregister'
    for mod in ("app.aksjonærregister.db", "app.aksjonaerregister.db"):
        try:
            spec = importlib.util.find_spec(mod)
        except Exception as e:
            print(f"find_spec({mod!r}) ga feil:", repr(e))
            spec = None
        print(f"find_spec({mod!r}) ->", spec)

    # Prøv faktisk import med korrekt navn
    try:
        from app.aksjonærregister import db as ar_db  # type: ignore
        print("✅ Import OK:", ar_db.__file__)
    except Exception:
        print("❌ Import med 'app.aksjonærregister' feilet – forsøker ascii-fallback…")
        try:
            from app.aksjonaerregister import db as ar_db  # type: ignore
            print("✅ Import OK (ascii-fallback):", ar_db.__file__)
        except Exception:
            print("❌ Import feilet:")
            traceback.print_exc()
            return

    # 3) Åpne tilkobling (open_conn) og ping DB
    try:
        if hasattr(ar_db, "open_conn"):
            conn = ar_db.open_conn()
            print("✅ open_conn() OK:", type(conn))
            try:
                res = conn.execute("SELECT 1").fetchone()
                print("Ping DB:", res)
            except Exception as e:
                print("Ping DB (hopper over)->", e)
        else:
            print("ℹ️  Fant ikke 'open_conn' i ar_db – sjekk API i din db.py")
    except Exception:
        print("❌ Åpning av tilkobling feilet:")
        traceback.print_exc()

if __name__ == "__main__":
    main()
