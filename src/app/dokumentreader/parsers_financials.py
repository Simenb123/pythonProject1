from __future__ import annotations
# --- bootstrap ---
import os, sys, re
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# ------------------
from typing import Dict
from decimal import Decimal
from app.dokumentreader.utils import normspace, parse_decimal, detect_currency

# Forenklet skjemaparser for "Resultatregnskap" og "Balanse".
# Vi ser etter seksjonsoverskrifter, og linjer på formen:
#   "Driftsinntekter ................... 1 234 567,00"
#   "Sum eiendeler ..................... 9 876,00"
# NB: Dette er en robust baseline – du kan senere lage leverandørprofiler (YAML).

AMOUNT = r"\d{1,3}(?:[ .\u00A0\u202F]\d{3})*(?:,\d{2})?"
LINE_RE = re.compile(r"^(.{3,80}?)[ .\u00A0\u202F]*(" + AMOUNT + r")\s*$")

HEAD_RS = re.compile(r"\bResultat[- ]?regnskap\b", re.I)
HEAD_BAL = re.compile(r"\bBalanse\b", re.I)

def _collect_section(lines: list[str], start_idx: int) -> list[str]:
    """Samle linjer frem til neste store overskrift eller blank blokk."""
    buf = []
    for ln in lines[start_idx+1:]:
        s = normspace(ln)
        if not s:
            if buf and not any(buf[-1]):  # to blanks på rad
                break
        # stopp om ny seksjon
        if HEAD_RS.search(s) or HEAD_BAL.search(s):
            break
        buf.append(s)
    return buf

def parse_financial_statement(text: str) -> Dict[str, Dict[str, Decimal] | str]:
    lines = [l.rstrip() for l in text.splitlines()]
    currency = detect_currency(text)
    res_income: Dict[str, Decimal] = {}
    res_balance: Dict[str, Decimal] = {}

    for i, l in enumerate(lines):
        s = normspace(l)
        if HEAD_RS.search(s):
            body = _collect_section(lines, i)
            for b in body:
                m = LINE_RE.match(b)
                if not m:
                    continue
                key = normspace(m.group(1).strip(" .·:-"))
                amt = parse_decimal(m.group(2))
                if amt is not None:
                    res_income[key] = amt
        if HEAD_BAL.search(s):
            body = _collect_section(lines, i)
            for b in body:
                m = LINE_RE.match(b)
                if not m:
                    continue
                key = normspace(m.group(1).strip(" .·:-"))
                amt = parse_decimal(m.group(2))
                if amt is not None:
                    res_balance[key] = amt

    return {
        "currency": currency,
        "income_statement": res_income,
        "balance_sheet": res_balance,
    }
