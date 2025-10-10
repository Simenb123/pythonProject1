# app_logging.py – 2025-06-03
# ---------------------------------------------------------------------------
"""Felles logging-oppsett – kalles én gang tidlig i programmet."""
from __future__ import annotations

import logging
from typing import Any

_FMT = "%(asctime)s [%(levelname)s] %(name)s: %(message)s"


def configure(level: int = logging.INFO, *, fmt: str = _FMT, **kwargs: Any) -> None:
    logging.basicConfig(level=level, format=fmt, **kwargs)
