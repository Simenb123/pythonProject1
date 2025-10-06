from __future__ import annotations
import os, sys, re
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)

from app.dokumentreader.utils import NBSP_CHARS, NUMBER_RE

MONEY = NUMBER_RE

# Krev 'no|number' etter 'invoice' for å unngå at "Invoice Date" matches som fakturanr
RE_INVOICE_NO = re.compile(
    r"(?:faktura\s*(?:nr|nummer)|invoice\s*(?:no|number))\s*[:#]?\s*([A-Z0-9][A-Z0-9_/\-\. ]{1,})",
    re.I
)

RE_ORDER_NO = re.compile(r"bestillings?nr[:#]?\s*([A-Z0-9\-\/]+)", re.I)

# KID – tåler NBSP/mellomrom
RE_KID = re.compile(r"\bKID(?:\s*[-:]?\s*nr\.?)?\s*[:\-]?\s*([0-9" + NBSP_CHARS + r"\s]{6,30})\b", re.I)

# Orgnr / VAT / Company No (Norge + generiske engelske varianter)
RE_ORG_ANY = re.compile(
    r"\b(?:Org\.?\s*Nr|Company\s*(?:No|Number)|Registration\s*(?:No|Number))[: ]*\s*((?:NO\s*)?\d{9}(?:\s*MVA)?)",
    re.I
)

# Dato-etiketter
RE_DATE_LABELED = re.compile(r"(fakturadato|invoice\s*date|issue\s*date|date)\s*[:\-]?\s*([^\n]+)", re.I)
RE_DUE_LABELED  = re.compile(r"(forfallsdato|forfaller|due\s*date|payment\s*due)\s*[:\-]?\s*([^\n]+)", re.I)

# Summer
RE_TOTAL_INCL = re.compile(
    r"(?:total(?:sum)?|beløp\s*å\s*betale|amount\s*due|grand\s*total|total\s*payable|total\s*amount)\s*[:\-]?\s*(" + MONEY + r")\b",
    re.I
)
RE_TOTAL_EXCL = re.compile(
    r"(?:sum\s*å\s*betale\s*eks(?:l|kl)\.?\s*mva\.?|sum\s*eks(?:\s*mva)?|subtotal|net\s*total|total\s*excl\.?)\s*[:\-]?\s*(" + MONEY + r")\b",
    re.I
)

# MVA-beløp – ikke fang sats-prosent (… 25 %) som "beløp"
RE_VAT_AMOUNT = re.compile(
    r"(?:mva(?:-?beløp)?|merverdiavgift|vat(?:\s*(?:amount|total))?)"
    r"(?:\s*\d{1,2}[.,]?\d{0,2}\s*%)?"
    r"\s*(?:nok|kr|eur|usd|gbp)?\s*[:\-]?\s*(" + MONEY + r")(?!\s*%)",
    re.I
)

# MVA-sats
RE_VAT_RATE = re.compile(
    r"(?:\b(?:mva|vat)\b[^%\n]{0,15}(\d{1,2}[.,]?\d{0,2})\s*%|(\d{1,2}[.,]?\d{0,2})\s*%\s*(?:mva|vat)\b)",
    re.I
)

# --- Bakoverkompat: gammelt navn fra tidligere kode ---
RE_VATNO_ANY = RE_ORG_ANY
