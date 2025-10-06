from __future__ import annotations
# --- bootstrap ---
import os, sys, re
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# ------------------
from app.dokumentreader.doc_types import DocumentType

# Utvidet keyword-basert klassifisering.
KEYS = {
    DocumentType.INVOICE: [
        r"\bFaktura\b", r"\bFaktura\s*nr\b", r"\bKID\b",
        r"\bBeløp å betale\b", r"\bMVA\b",
        r"\bInvoice\b", r"\bInvoice\s*(?:no|number)\b", r"\bAmount\s*due\b", r"\bVAT\b",
        r"\bInvoice\s*date\b", r"\bDue\s*date\b",
    ],
    DocumentType.FINANCIAL_STATEMENT: [
        r"\bResultatregnskap\b", r"\bBalanse\b", r"\bÅrsregnskap\b", r"\bNoter?\b",
        r"\bDriftsinntekter\b", r"\bSum\s+eiendeler\b",
    ],
    DocumentType.TAX_RETURN: [
        r"\bSkattemelding\b", r"\bSelvangivelse\b", r"\bSkattepliktig\b",
        r"\bMidlertidige forskjeller\b", r"\bPermanente forskjeller\b",
        r"\bEgenkapitalavstemming\b", r"\bBegrensning av rentefradrag\b",
        r"\bSpesifikasjonsutskrift\b",  # Maestro utskrift
        r"\bSkattepliktig inntekt\b",
        r"\bPost\s*\d{2,4}\b",
    ],
}

def classify_text(text: str) -> tuple[DocumentType, float]:
    t = text or ""
    scores = {dt: 0 for dt in KEYS}
    for dt, pats in KEYS.items():
        for p in pats:
            if re.search(p, t, flags=re.I):
                scores[dt] += 1

    # Sterke signaler for skattemelding
    if re.search(r"\bSkattemelding\b", t, re.I):
        scores[DocumentType.TAX_RETURN] += 3
    if re.search(r"\bSpesifikasjonsutskrift\b", t, re.I) and re.search(r"(Skattepliktig|Midlertidige|Permanente|rentefradrag|Egenkapitalavstemming)", t, re.I):
        scores[DocumentType.TAX_RETURN] += 2

    best = max(scores.items(), key=lambda kv: kv[1])
    if best[1] == 0:
        return DocumentType.UNKNOWN, 0.0
    total = sum(scores.values()) or 1
    conf = best[1] / total
    return best[0], conf
