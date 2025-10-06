# -*- coding: utf-8 -*-
from __future__ import annotations
"""
Kjør denne fila direkte selv om den ligger i:  src/app/aksjonaerregister/
Robust oppstarter som setter opp pakkekontekst og laster ui_tk/cli
med fullt kvalifiserte navn slik at relative imports fungerer.

Bruk:
  Run uten argumenter  -> GUI (hvis Tkinter finnes)
  Run med argumenter   -> CLI, f.eks.:
      build --csv "C:\\sti\\fil.csv" --delimiter ";"
      search "alpha" --by navn --limit 20
      graph --orgnr 910000001 --name "Alpha AS" --mode both
"""
import os, sys, importlib.util, types

# ---- finn stiene vi trenger ----
THIS    = os.path.abspath(__file__)                 # .../src/app/aksjonaerregister/run_*.py
PKG_DIR = os.path.dirname(THIS)                     # .../src/app/aksjonaerregister
APP_DIR = os.path.dirname(PKG_DIR)                  # .../src/app
SRC_DIR = os.path.dirname(APP_DIR)                  # .../src
PROJ_DIR= os.path.dirname(SRC_DIR)                  # prosjektrot

for p in (PROJ_DIR, SRC_DIR, APP_DIR, PKG_DIR):
    if p and os.path.isdir(p) and p not in sys.path:
        sys.path.insert(0, p)

def _ensure_pkg_namespace(base: str) -> None:
    """Lag nødvendige pakke-noder i sys.modules med __path__ slik at relative imports virker."""
    if base == "app.aksjonaerregister":
        if "app" not in sys.modules:
            app_pkg = types.ModuleType("app")
            app_pkg.__path__ = [APP_DIR]
            sys.modules["app"] = app_pkg
        if base not in sys.modules:
            aksj_pkg = types.ModuleType(base)
            aksj_pkg.__path__ = [PKG_DIR]
            sys.modules[base] = aksj_pkg
    elif base == "aksjonaerregister":
        if base not in sys.modules:
            pkg = types.ModuleType(base)
            pkg.__path__ = [PKG_DIR]
            sys.modules[base] = pkg

def _load_pkg_module(base: str, modname: str, path: str):
    """Last modul fra filsti, men under gitt pakkebase (for relative imports)."""
    _ensure_pkg_namespace(base)
    fullname = f"{base}.{modname}"
    spec = importlib.util.spec_from_file_location(fullname, path)
    if spec is None or spec.loader is None:
        raise ImportError(f"Cannot create spec for {fullname} from {path}")
    module = importlib.util.module_from_spec(spec)
    # Sørg for at relative imports (from .foo import ...) vet hvor 'parent' er:
    module.__package__ = base
    sys.modules[fullname] = module
    spec.loader.exec_module(module)  # type: ignore[attr-defined]
    return module

def _try_load():
    ui_path  = os.path.join(PKG_DIR, "ui_tk.py")
    cli_path = os.path.join(PKG_DIR, "cli.py")
    # 1) forsøk med app.aksjonaerregister (din layout)
    try:
        ui  = _load_pkg_module("app.aksjonaerregister", "ui_tk", ui_path)
        cli = _load_pkg_module("app.aksjonaerregister", "cli",   cli_path)
        return ui, cli
    except Exception:
        # 2) fallback uten "app."
        ui  = _load_pkg_module("aksjonaerregister", "ui_tk", ui_path)
        cli = _load_pkg_module("aksjonaerregister", "cli",   cli_path)
        return ui, cli

def _tk_available():
    try:
        import tkinter  # noqa
        return True
    except Exception:
        return False

def main(argv=None):
    ui, cli = _try_load()
    args = list(sys.argv[1:] if argv is None else argv)
    if _tk_available() and not args:
        app = ui.App()
        app.mainloop()
    else:
        cli.main(args if args else ["-h"])

if __name__ == "__main__":
    main()
