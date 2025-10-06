from __future__ import annotations
# --- bootstrap ---
import os, sys, re
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# ------------------
from typing import Dict, List
from app.dokumentreader.utils import normspace, parse_decimal, detect_currency
from app.dokumentreader.doc_types import KVPair

# Tall (robust – tåler NBSP-varianter)
AMOUNT = r"\d{1,3}(?:[ .\u00A0\u202F]\d{3})*(?:,\d{2})?"

# Linjetyper
POST_LINE  = re.compile(r"^(Post\s*\d{2,4}.*?)\s+(" + AMOUNT + r")$", re.I)
# Krev BOKSTAV i nøkkelen (hindrer at '938 998 582' blir feiltolket som '938: 998 582')
KV_COLON   = re.compile(r"^(?=.*[A-Za-zÆØÅæøå])(.{2,150}?)[ :]\s*(" + AMOUNT + r")$")
# Maestro: 'Label  I år [I fjor]' -> ta første beløp (I år)
LINE_YEAR  = re.compile(r"^(.{3,120}?)\s+(" + AMOUNT + r")(?:\s+" + AMOUNT + r")?\s*$")
# Ja/Nei-spørsmål
QA_RE      = re.compile(r"^(.+?\?)\s*(Ja|Nei)\s*$", re.I)

# Meta
ORG_RE  = re.compile(r"\b(\d{3}\s?\d{3}\s?\d{3})\b")
YEAR_RE = re.compile(r"\b(20\d{2}|19\d{2})\b")

# Støy (linjer vi ikke vil putte i fields)
NOISE = re.compile(r"(?:^| )(Nei|Ja|Side\s*\d+|Spesifikasjonsutskrift|Maestro\s+Årsoppgjør)\b", re.I)

# Hvilke linjer vi vil løfte til 'posts' (nøkkeltall)
INTERESTING_KEYS = (
    "årsresultat", "driftsresultat", "resultat før skattekostnad",
    "næringsinntekt", "netto inntekt",
    "sum salgsinntekt", "sum driftsinntekt", "sum driftskostnad",
    "sum eiendeler", "egenkapital og gjeld",
    "inngående egenkapital", "utgående egenkapital",
    "sum underskudd til fremføring",
)

def _clean_amount(s: str) -> float | None:
    d = parse_decimal(s)
    return float(d) if d is not None else None

def parse_tax_return(text: str) -> Dict[str, object]:
    """
    Skattemelding/Spesifikasjonsutskrift (Maestro m.fl.)
    Returnerer:
      - taxpayer_name, orgnr, income_year
      - fields: fri nøkkel/verdi (inkl. Ja/Nei)
      - posts: utvalgte hovedtall ('I år')
    """
    currency = detect_currency(text)
    fields: List[KVPair] = []
    posts: Dict[str, float] = {}

    lines = [normspace(l) for l in (text or "").splitlines()]

    # --- meta (navn/orgnr/år) fra topp
    taxpayer_name = None
    orgnr = None
    income_year = None
    for i, ln in enumerate(lines[:60]):
        if not orgnr:
            m = ORG_RE.search(ln)
            if m:
                orgnr = m.group(1).replace(" ", "")
                # bruk linjen over som navn hvis meningsfull
                for j in range(i-1, max(-1, i-6), -1):
                    s = lines[j].strip()
                    if s and not ORG_RE.search(s) and not YEAR_RE.search(s):
                        taxpayer_name = s
                        break
        if not income_year:
            cand = [y for y in YEAR_RE.findall(ln) if y.startswith("20")]
            if cand:
                income_year = cand[0]
        if taxpayer_name and orgnr and income_year:
            break

    # --- hovedsløyfe
    for raw in lines:
        ln = normspace(raw)
        if not ln or NOISE.search(ln):
            continue

        # Post xxxx ...
        m = POST_LINE.match(ln)
        if m:
            posts[m.group(1)] = _clean_amount(m.group(2)) or 0.0
            continue

        # Maestro-rad: Label  I år [I fjor]
        m3 = LINE_YEAR.match(ln)
        if m3:
            label = m3.group(1).strip(" .·:-")
            first = m3.group(2)
            # hopp over tabell-overskrifter
            if re.match(r"^(Nr|Betegnelse|I år|I fjor)$", label, re.I):
                continue
            val = _clean_amount(first)
            if val is not None:
                fields.append(KVPair(key=label, value=m3.group(2)))
                l = label.lower()
                if any(k in l for k in INTERESTING_KEYS):
                    posts[label] = val
            continue

        # Label: beløp (må inneholde bokstaver)
        m2 = KV_COLON.match(ln)
        if m2:
            key = m2.group(1).strip()
            val_str = m2.group(2)
            fields.append(KVPair(key=key, value=val_str))
            l = key.lower()
            val = _clean_amount(val_str)
            if val is not None and any(k in l for k in INTERESTING_KEYS):
                posts[key] = val
            continue

        # Ja/Nei-spørsmål
        mq = QA_RE.match(ln)
        if mq:
            fields.append(KVPair(key=mq.group(1).strip(), value=mq.group(2).title()))
            continue

    return {
        "currency": currency,
        "taxpayer_name": taxpayer_name,
        "orgnr": orgnr,
        "income_year": income_year,
        "fields": [f.dict() for f in fields],
        "posts": posts
    }
