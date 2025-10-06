from __future__ import annotations

# --- BOOTSTRAP så scriptet kan kjøres direkte med Run ---
import os, sys
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# --------------------------------------------------------

import re
from decimal import Decimal, ROUND_HALF_UP
from typing import Any, Dict, List

from app.dokumentreader.utils import (
    normspace, clean_text_for_search, parse_decimal, parse_date_any,
    detect_currency, NUMBER_RE
)
from app.dokumentreader.regexes import (
    RE_INVOICE_NO, RE_ORDER_NO, RE_KID, RE_ORG_ANY, RE_VATNO_ANY,
    RE_DATE_LABELED, RE_DUE_LABELED,
    RE_TOTAL_INCL, RE_TOTAL_EXCL, RE_VAT_AMOUNT, RE_VAT_RATE
)
from app.dokumentreader.models import LineItemModel, PartyModel

HEADER_ALIASES = {
    "description": ["beskrivelse", "vare", "produkt", "tjeneste", "artikkel", "description", "item"],
    "qty": ["antall", "mengde", "qty", "quantity", "stk"],
    "unit": ["enhet", "unit", "stk", "mnd", "hour", "time", "tim"],
    "unit_price": ["pris", "enhetspris", "á", "a pris", "unit price", "stk pris"],
    "vat_rate": ["mva", "mva %", "vat", "vat %"],
    "line_total": ["sum", "beløp", "amount", "total", "linjesum"]
}


def extract_key_fields(text: str) -> Dict[str, Any]:
    res: Dict[str, Any] = {}
    clean = text or ""

    # Fakturanr / bestillingsnr / KID
    m = RE_INVOICE_NO.search(clean)
    if m:
        res["invoice_number"] = normspace(m.group(1))
    m = RE_ORDER_NO.search(clean)
    if m:
        res["order_reference"] = m.group(1)
    m = RE_KID.search(clean)
    if m:
        res["kid_number"] = re.sub(r"\s+", "", m.group(1))

    # Orgnr/VAT (selger)
    m = RE_ORG_ANY.search(clean)
    if m:
        val = normspace(m.group(1))
        digits_m = re.search(r"(\d{9})", val)
        if digits_m:
            res["seller_orgnr"] = digits_m.group(1)
        if "MVA" in val.upper() or "NO" in val.upper():
            res["seller_vatno"] = val

    if not res.get("seller_vatno"):
        m = RE_VATNO_ANY.search(clean)
        if m:
            res["seller_vatno"] = normspace(m.group(1))
            digits_m = re.search(r"(\d{9})", res["seller_vatno"])
            if digits_m and not res.get("seller_orgnr"):
                res["seller_orgnr"] = digits_m.group(1)

    # Datoer
    m = RE_DATE_LABELED.search(clean)
    if m:
        dt = parse_date_any(m.group(2))
        res["invoice_date"] = dt.isoformat() if dt else None
    m = RE_DUE_LABELED.search(clean)
    if m:
        dd = parse_date_any(m.group(2))
        res["due_date"] = dd.isoformat() if dd else None

    # Valuta
    res["currency"] = detect_currency(clean)

    # Summer
    m = RE_TOTAL_EXCL.search(clean)
    if m:
        res["subtotal_excl_vat"] = parse_decimal(m.group(1))
    m = RE_VAT_AMOUNT.search(clean)
    if m:
        res["vat_amount"] = parse_decimal(m.group(1))
    m = RE_TOTAL_INCL.search(clean)
    if m:
        res["total_incl_vat"] = parse_decimal(m.group(1))

    # MVA-satser
    rates = set()
    for match in RE_VAT_RATE.finditer(clean):
        cand = match.group(1) or match.group(2)
        val = parse_decimal(cand) if cand else None
        if val is not None:
            rates.add(val)
    if rates:
        res["vat_rates"] = sorted(rates)

    return res


def _score_company_line(s: str) -> int:
    score = 0
    # Preferer selskapsnavn (AS/ASA/A/S)
    if re.search(r"\b(A/S|AS|ASA)\b", s): score += 3
    # Ingen tall -> ofte navn
    if not re.search(r"\d", s): score += 2
    # Typiske adresseord/land – trekk ned
    if re.search(r"\b(postboks|pb|gate|gaten|vei|veien|road|street|all[eé]|oslo|bergen|trondheim|stavanger|norge|norway)\b", s, re.I):
        score -= 3
    # Postnummer
    if re.search(r"\b\d{4}\b", s): score -= 2
    if len(s) > 50: score -= 1
    return score


def pick_parties(text: str, max_scan_lines: int = 120):
    """
    Heuristikk for selger/kjøper.
    Forbedringer:
      - selger: støtter 'Org. Nr' *og* 'Company/VAT/Registration No'
      - kjøper: bruker 'To: …' hvis finnes; hopper også over 'Invoice' i topplinjer
    """
    lines = [l for l in (text or "").splitlines() if l.strip()]
    seller = PartyModel()
    buyer = PartyModel()

    # SELGER: finn org-/company-/vat-anker, se noen linjer over og velg best scorede kandidat
    anchor_re = re.compile(r"(Org\.?\s*Nr|VAT\s*No|Company\s*(?:No|Number)|Registration\s*(?:No|Number))", re.I)
    best = ("", -999)
    for i, l in enumerate(lines[:max_scan_lines]):
        if anchor_re.search(l):
            for j in range(i-1, max(-1, i-10), -1):
                cand = normspace(lines[j])
                if re.search(r"\b(Norge|Norway)\b", cand, re.I):
                    continue
                sc = _score_company_line(cand)
                if sc > best[1]:
                    best = (cand, sc)
            break
    seller.name = best[0] or None

    # KJØPER: prøv "To:" først
    for l in lines[:max_scan_lines]:
        m = re.match(r"^\s*To\s*:\s*(.+)$", l, re.I)
        if m:
            buyer.name = normspace(m.group(1))
            break

    # Fallback – første meningsfulle linje i starten (unngå metadata/etiketter)
    if not buyer.name:
        skip = re.compile(r"(Org\.?\s*Nr|Faktura|Invoice|Kunde\s*nr|SWIFT|IBAN|Valuta|Kid\s*nr|Fakturadato|Forfallsdato)", re.I)
        for l in lines[:10]:
            if skip.search(l) or re.search(r"\d{5,}", l):
                continue
            buyer.name = normspace(l)
            break

    return seller, buyer


def map_columns(header_row: List[str]):
    def normalize(h: str) -> str:
        h = clean_text_for_search(h).lower()
        return re.sub(r"[^a-z0-9% ]", "", h)
    mapping = {}
    norm_headers = [normalize(h) for h in header_row]
    for idx, nh in enumerate(norm_headers):
        for target, aliases in HEADER_ALIASES.items():
            if any(a in nh for a in aliases):
                mapping[idx] = target
    return mapping


def parse_line_items_from_tables(tables) -> List[LineItemModel]:
    items: List[LineItemModel] = []
    for df in tables:
        df = df.copy()
        df.columns = [str(c) for c in df.columns]
        header = [normspace(str(x)) for x in df.iloc[0].tolist()]
        colmap = map_columns(header)
        body = df.iloc[1:].reset_index(drop=True)

        if not colmap:
            if 0 not in colmap:
                colmap[0] = "description"
            for idx in range(len(header)-1, -1, -1):
                col_vals = " ".join(body.iloc[:, idx].astype(str).tolist())[:200]
                if re.search(NUMBER_RE, col_vals):
                    colmap[idx] = "line_total"
                    break

        for _, row in body.iterrows():
            raw = {i: str(row.iloc[i]) for i in range(len(row))}
            li = LineItemModel(description="")
            for i, target in colmap.items():
                val = normspace(str(raw.get(i, "")))
                if target == "description":
                    li.description = val
                elif target == "qty":
                    li.quantity = parse_decimal(val)
                elif target == "unit":
                    li.unit = val
                elif target == "unit_price":
                    li.unit_price = parse_decimal(val)
                elif target == "vat_rate":
                    li.vat_rate = parse_decimal(val)
                elif target == "line_total":
                    li.line_total = parse_decimal(val)
            if li.description:
                if li.line_total is None and li.unit_price is not None and li.quantity is not None:
                    try:
                        li.line_total = (li.unit_price * li.quantity).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
                    except Exception:
                        pass
                items.append(li)
    return items


def parse_line_items_from_text(text: str) -> List[LineItemModel]:
    """
    Streng fallback: krev desimaler + 'signalord' (NOK/kr/Total/MVA/Sum/Beløp).
    Dette unngår at tid (0,25) og andre tall blir tolket som linjesummer.
    """
    items: List[LineItemModel] = []
    MONEY_DEC = r"\d{1,3}(?:[ .\u00A0\u202F]\d{3})*,\d{2}"
    SIGNAL = r"(?:\bNOK\b|\bkr\b|mva|vat|total|sum|beløp|amount)"
    lines = [normspace(l) for l in (text or "").splitlines()]
    for ln in lines:
        m = re.search(rf"({MONEY_DEC})\s*$", ln, re.I)
        if not m:
            continue
        if not re.search(SIGNAL, ln, re.I):
            continue
        if re.search(r"(Org\.?\s*Nr|KID|IBAN|SWIFT|Postboks|Bankkonto|Telefon|\+47)", ln, re.I):
            continue
        amount = parse_decimal(m.group(1))
        desc = re.split(r"\s{2,}", ln)[0]
        if desc:
            items.append(LineItemModel(description=desc, line_total=amount))
    return items
