# ui_theme.py – Sprint 5 (tema-bytte) – 2025-06-02
# ---------------------------------------------------------------------------
from __future__ import annotations

import json
import logging
import tkinter as tk
from pathlib import Path
from typing import Iterable

try:
    import ttkbootstrap as ttkb

    _HAS_BOOTSTRAP = True
except ImportError:
    import tkinter.ttk as ttkb  # type: ignore

    _HAS_BOOTSTRAP = False

logger = logging.getLogger(__name__)

CONFIG = Path.home() / ".bilagsuttrekk.config"
DEFAULT_THEME = "flatly" if _HAS_BOOTSTRAP else "default"


def load_theme() -> str:
    if CONFIG.exists():
        try:
            return json.loads(CONFIG.read_text()).get("theme", DEFAULT_THEME)
        except Exception:
            logger.warning("Korrupt konfig – bruker default theme.")
    return DEFAULT_THEME


def save_theme(name: str):
    try:
        CONFIG.write_text(json.dumps({"theme": name}))
    except Exception:
        logger.warning("Kunne ikke lagre tema-valg.")


def init_style() -> "ttkb.Style":
    """Må kalles *før* første Tk-vindu lages."""
    theme = load_theme()
    if _HAS_BOOTSTRAP:
        style = ttkb.Style(theme=theme)
    else:
        style = ttkb.Style()
        style.theme_use(theme if theme in style.theme_names() else "default")
    return style


def available_themes() -> Iterable[str]:
    if _HAS_BOOTSTRAP:
        return ttkb.Style().theme_names()
    return ttkb.Style().theme_names()


def set_theme(theme_name: str):
    """Endre tema på alle levende Tk-vinduer."""
    save_theme(theme_name)
    for w in map(tk._get_default_root, tk._default_root.children.values() if tk._default_root else []):  # type: ignore
        try:
            w.style.theme_use(theme_name)  # ttkbootstrap
        except Exception:
            try:
                w.tk.call("ttk::style", "theme", "use", theme_name)  # standard ttk
            except Exception:
                pass
