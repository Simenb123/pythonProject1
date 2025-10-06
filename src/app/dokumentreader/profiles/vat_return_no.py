from __future__ import annotations
from typing import Any, Dict, List, Tuple
import re
import os

from .base import DocProfile, as_result

# NB: Ulike systemer skriver skattemelding MVA forskjellig.
# Heuristikk: se etter feltnavn/linjer, ‘termin’, ‘sum utgående’, ‘sum inngående’, ‘fradrag’, ‘til gode’, ‘å betale’.
RE_TERM  = re.compile(r"\b(termin|periode)\b.*?(\d{1,2})\s*[-/]\s*(\d{4})", re.I)
RE_OUT   = re.compile(r"(sum\s+utg(?:ående)?\s*mva|utgående mva).*?([\d .,\u00A0\u202F]+)", re.I)
RE_IN    = re.compile(r"(sum\s+inng(?:ående)?\s*mva|inngående mva).*?([\d .,\u00A0\u202F]+)", re.I)
RE_DUE   = re.compile(r"(å\s*betale|betal(?:es)?).*?([\d .,\u00A0\u202F]+)", re.I)
RE_TILG  = re.compile(r"(til\s*gode).*?([\d .,\u00A0\u202F]+)", re.I)

def _to_decimal(s: str) -> float | None:
    s = s.replace("\u00A0", " ").replace("\u202F", " ")
    s = re.sub(r"[^\d,.\- ]", "", s)
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return None

class VatReturnNoProfile(DocProfile):
    name = "vat_return_no"
    description = "Skattemelding MVA (summer per termin)"

    def detect(self, page1_text: str, full_text: str) -> float:
        t = full_text.lower()
        score = 0
        for w in ["skattemelding", "merverdiavgift", "mva-melding", "mva-kode", "mva-koder", "termin"]:
            if w in t: score += 0.2
        if re.search(RE_OUT, full_text): score += 0.4
        if re.search(RE_IN,  full_text): score += 0.3
        return min(1.0, score)

    def parse(self, path: str, page1_text: str, full_text: str) -> Dict[str, Any]:
        # grov oppsummering
        out = {
            "file_name": os.path.basename(path),
            "term": None,
            "year": None,
            "amounts": {
                "outgoing_vat": None,
                "incoming_vat": None,
                "net_due": None,
                "net_refund": None
            },
            "note": None
        }
        m = RE_TERM.search(full_text)
        if m:
            out["term"] = int(m.group(2))
            out["year"] = int(m.group(3))
        m = RE_OUT.search(full_text)
        if m: out["amounts"]["outgoing_vat"] = _to_decimal(m.group(2))
        m = RE_IN.search(full_text)
        if m: out["amounts"]["incoming_vat"] = _to_decimal(m.group(2))
        m = RE_DUE.search(full_text)
        if m: out["amounts"]["net_due"] = _to_decimal(m.group(2))
        m = RE_TILG.search(full_text)
        if m: out["amounts"]["net_refund"] = _to_decimal(m.group(2))

        return as_result("vat_return_no", out)
