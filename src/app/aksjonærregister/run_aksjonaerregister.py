# -*- coding: utf-8 -*-
"""
Enkel oppstartfil – kjør denne i PyCharm med Run/Debug.
Tilpasset mappestrukturen:  src/app/aksjonaerregister

- Hvis Tkinter finnes: starter GUI (app.aksjonaerregister.ui_tk.App)
- Ellers: CLI-hjelp / CLI-kommandoer
"""
from __future__ import annotations
import os, sys

def _ensure_on_path():
    # Legg til prosjektrot og src i sys.path (så 'app' er importbar)
    here = os.path.abspath(os.getcwd())
    src = os.path.join(here, "src")
    roots = [here, src]
    for p in roots:
        if os.path.isdir(p) and p not in sys.path:
            sys.path.insert(0, p)

def main():
    _ensure_on_path()

    # Prøv først pakke under app/, deretter toppnivå
    App = None
    cli_main = None
    try:
        from app.aksjonaerregister.ui_tk import App as _App  # type: ignore
        from app.aksjonaerregister.cli import main as _cli_main  # type: ignore
        App, cli_main = _App, _cli_main
    except Exception:
        from aksjonaerregister.ui_tk import App as _App  # type: ignore
        from aksjonaerregister.cli import main as _cli_main  # type: ignore
        App, cli_main = _App, _cli_main

    # GUI hvis Tkinter er tilgjengelig og ingen CLI-argumenter er gitt
    try:
        import tkinter  # noqa: F401
        TK = True
    except Exception:
        TK = False

    argv = sys.argv[1:]
    if TK and not argv:
        app = App(); app.mainloop()
    else:
        cli_main(argv if argv else ["-h"])

if __name__ == "__main__":
    main()
