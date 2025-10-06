from __future__ import annotations
"""
Entry-point for `python -m aksjonaerregister`.

- Hvis Tkinter er tilgjengelig: start GUI (ui_tk.App)
- Hvis ikke: vis CLI-hjelp (samme som `python -m aksjonaerregister.cli -h`)
"""
import sys

# Sjekk Tkinter
try:
    import tkinter  # noqa: F401
    TK_AVAILABLE = True
except Exception:
    TK_AVAILABLE = False

from . import settings as S
from .ui_tk import App  # type: ignore
from .cli import main as cli_main

def main(argv: list[str] | None = None) -> None:
    argv = list(sys.argv[1:] if argv is None else argv)
    if TK_AVAILABLE:
        app = App()
        app.mainloop()
    else:
        # Ingen Tkinter → gå til CLI
        if not argv or argv == ["-h"] or argv == ["--help"]:
            cli_main(["-h"])
        else:
            cli_main(argv)

if __name__ == "__main__":  # pragma: no cover
    main()
