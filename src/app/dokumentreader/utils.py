from __future__ import annotations

import re
from datetime import date
from decimal import Decimal, InvalidOperation
from typing import Optional

from dateutil import parser as dateparser
from unidecode import unidecode

# ——— Locale/tegnsett ———
NORWEGIAN_MONTHS = {
    "januar": 1, "februar": 2, "mars": 3, "april": 4, "mai": 5, "juni": 6,
    "juli": 7, "august": 8, "september": 9, "oktober": 10, "november": 11, "desember": 12
}
NOR_MONTHS_ABBR = {
    "jan": 1, "feb": 2, "mar": 3, "apr": 4, "mai": 5, "jun": 6,
    "jul": 7, "aug": 8, "sep": 9, "okt": 10, "nov": 11, "des": 12
}

# NBSP-varianter som ofte er tusenskilletegn i PDF-er
NBSP_CHARS = "\u00A0\u202F\u2009"  # NBSP, Narrow NBSP, Thin Space

# Tallsøk som godtar NBSP-varianter
NUMBER_RE = rf"[0-9][0-9 {NBSP_CHARS}.,]*[0-9]|[0-9]"


def normspace(s: str) -> str:
    return re.sub(r"\s+", " ", s or "").strip()


def clean_text_for_search(s: str) -> str:
    return normspace(unidecode(s or ""))


def parse_decimal(s: str) -> Optional[Decimal]:
    """Robust tallparser for norske/engelske skilletegn og NBSP-varianter."""
    if s is None:
        return None
    s = s.strip()
    for ch in NBSP_CHARS:
        s = s.replace(ch, " ")
    s = re.sub(r"[^\d,.\-]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return Decimal(s)
    except InvalidOperation:
        return None


def parse_date_any(s: str) -> Optional[date]:
    """dd.mm.yyyy, yyyy-mm-dd og norske månedsnavn/forkortelser (jan./des.)."""
    if not s:
        return None
    t = s.strip()
    for name, num in {**NORWEGIAN_MONTHS, **NOR_MONTHS_ABBR}.items():
        t = re.sub(rf"\b{name}\.?\b", str(num), t, flags=re.I)
    try:
        return dateparser.parse(t, dayfirst=True, fuzzy=True).date()
    except Exception:
        return None


def detect_currency(text: str) -> str:
    if re.search(r"\bNOK\b", text, re.I): return "NOK"
    if re.search(r"\bSEK\b", text, re.I): return "SEK"
    if re.search(r"\bDKK\b", text, re.I): return "DKK"
    if "€" in text or re.search(r"\bEUR\b", text, re.I): return "EUR"
    if "$" in text or re.search(r"\bUSD\b", text, re.I): return "USD"
    if "£" in text or re.search(r"\bGBP\b", text, re.I): return "GBP"
    if re.search(r"\bkr\b", text, re.I): return "NOK"
    return "NOK"


def try_int(s: str) -> Optional[int]:
    try:
        return int(s)
    except Exception:
        return None
