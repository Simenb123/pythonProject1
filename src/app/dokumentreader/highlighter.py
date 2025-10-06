# -*- coding: utf-8 -*-
"""
Bygger highlight-data for fakturaer.
Utvidet til å markere:
- Verdier: invoice_number, kid_number, invoice_date, due_date, subtotal, vat_amount, total, seller/buyer/currency
- Etiketter: norsk + engelsk varianter
Returnerer:
  page_map: {page_index: [(Rect, label, color, kind('value'|'label'), key_or_None), ...]}
  key_hits: {key: [(page_index, Rect), ...]}
"""

from __future__ import annotations
import os, sys, re
from typing import Dict, List, Tuple, Optional
import fitz  # PyMuPDF

# --- bootstrap ---
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# ----------------------------------------------------

from app.dokumentreader.models import InvoiceModel

Rect = fitz.Rect
PageMapEntry = Tuple[Rect, str, str, str, Optional[str]]  # rect, label, color, kind, key
PageMap = Dict[int, List[PageMapEntry]]
KeyHits = Dict[str, List[Tuple[int, Rect]]]

LABELS = [
    # Nummer/KID
    "Faktura nr", "Fakturanr", "Invoice no", "Invoice number",
    "KID", "Kid nr", "KID nr", "Kidnummer",
    # Datoer
    "Fakturadato", "Invoice date", "Issue date", "Dato",
    "Forfallsdato", "Forfaller", "Due date", "Payment due",
    # Summer
    "Beløp å betale", "Amount due", "Totalsum", "Total", "Total payable", "Grand total",
    "Subtotal", "Net total", "MVA", "VAT",
    # Andre
    "Kunde", "Customer", "Selger", "Supplier", "Buyer", "Purchaser", "Valuta", "Currency"
]

DISPLAY = {
    "invoice_number": "Fakturanr",
    "kid_number":     "KID",
    "invoice_date":   "Fakturadato",
    "due_date":       "Forfallsdato",
    "subtotal":       "Eks. mva",
    "vat_amount":     "MVA",
    "total":          "Total",
    "seller_name":    "Selger",
    "seller_org":     "Selger orgnr",
    "seller_vat":     "Selger MVA",
    "buyer_name":     "Kjøper",
    "currency":       "Valuta",
}

COLORS = {
    "label":          "#9e9e9e",
    "invoice_number": "#1976d2",
    "kid_number":     "#8e24aa",
    "invoice_date":   "#2e7d32",
    "due_date":       "#00897b",
    "subtotal":       "#5d4037",
    "vat_amount":     "#6a1b9a",
    "total":          "#d32f2f",
    "seller_name":    "#455a64",
    "seller_org":     "#283593",
    "seller_vat":     "#1e88e5",
    "buyer_name":     "#546e7a",
    "currency":       "#37474f",
}

NBSP = "\u00A0"
NNBSP = "\u202F"


def _search_exact(page: fitz.Page, text: str) -> List[Rect]:
    out: List[Rect] = []
    if not text:
        return out
    try:
        for r in page.search_for(text):
            out.append(r)
    except Exception:
        pass
    return out


def _find_digits_sequence(page: fitz.Page, digits_only: str) -> List[Rect]:
    """Match tallsekvenser (KID/orgnr) selv om PDF har mellomrom/nbsp mellom segmenter."""
    d = re.sub(r"\D", "", digits_only or "")
    if not d:
        return []
    words = page.get_text("words")
    toks = [(w[0], w[1], w[2], w[3], w[4]) for w in words if re.search(r"\d", w[4])]
    toks.sort(key=lambda t: (round(t[1] / 5), t[0]))

    n = len(toks)
    for i in range(n):
        s = ""
        x0, y0, x1, y1 = toks[i][0], toks[i][1], toks[i][2], toks[i][3]
        for j in range(i, min(i + 16, n)):
            s += re.sub(r"\D", "", toks[j][4])
            x0, y0 = min(x0, toks[j][0]), min(y0, toks[j][1])
            x1, y1 = max(x1, toks[j][2]), max(y1, toks[j][3])
            if s == d:
                return [Rect(x0, y0, x1, y1)]
            if len(s) > len(d):
                break
    return []


def _amount_variants(x: Optional[str]) -> List[str]:
    if not x:
        return []
    s = str(x)
    base = s.replace(".", ",")
    out = [base]
    try:
        from decimal import Decimal
        d = Decimal(s)
        grouped = f"{d:,.2f}"                    # 167,890.00
        grouped_no = grouped.replace(",", " ").replace(".", ",") # 167 890,00
        out += [grouped_no, grouped_no.replace(" ", NBSP), grouped_no.replace(" ", NNBSP)]
    except Exception:
        pass
    # unike i rekkefølge
    return list(dict.fromkeys(out))


def _date_variants(iso_str: Optional[str]) -> List[str]:
    """Inkluder engelske månedsnavn-varianter som '2024, February 29'."""
    if not iso_str:
        return []
    try:
        from datetime import date
        y, m, d = [int(p) for p in iso_str.split("-")]
        dt = date(y, m, d)
        return [
            dt.strftime("%Y-%m-%d"),
            dt.strftime("%d.%m.%Y"),
            dt.strftime("%d.%m.%y"),
            dt.strftime("%d. %m. %Y"),
            dt.strftime("%d %b %Y"),    # 19 Feb 2024
            dt.strftime("%d %B %Y"),    # 19 February 2024
            dt.strftime("%b %d, %Y"),   # Feb 19, 2024
            dt.strftime("%B %d, %Y"),   # February 19, 2024
            dt.strftime("%Y, %B %d"),   # 2024, February 19
        ]
    except Exception:
        return [iso_str]


def build_invoice_highlights(pdf_path: str, inv: InvoiceModel) -> tuple[PageMap, KeyHits]:
    page_map: PageMap = {}
    key_hits: KeyHits = {}

    def add(page_no: int, rect: Rect, label: str, color: str, kind: str, key: Optional[str] = None):
        page_map.setdefault(page_no, []).append((rect, label, color, kind, key))
        if key:
            key_hits.setdefault(key, []).append((page_no, rect))

    with fitz.open(pdf_path) as doc:
        # 1) Etiketter – alle sider
        for pno in range(len(doc)):
            page = doc[pno]
            for lab in LABELS:
                for r in _search_exact(page, lab):
                    add(pno, r, lab, COLORS["label"], "label", None)

        # 2) Verdier – søk hele dokumentet
        targets = []
        if inv.invoice_number:
            targets.append(("invoice_number", [inv.invoice_number]))
        if inv.kid_number:
            targets.append(("kid_number", [inv.kid_number]))
        targets.append(("invoice_date", _date_variants(inv.invoice_date)))
        due = inv.payment_terms.due_date if inv.payment_terms else None
        targets.append(("due_date", _date_variants(due)))
        if inv.amounts and inv.amounts.subtotal_excl_vat is not None:
            targets.append(("subtotal", _amount_variants(str(inv.amounts.subtotal_excl_vat))))
        if inv.amounts and inv.amounts.vat_amount is not None:
            targets.append(("vat_amount", _amount_variants(str(inv.amounts.vat_amount))))
        if inv.amounts and inv.amounts.total_incl_vat is not None:
            targets.append(("total", _amount_variants(str(inv.amounts.total_incl_vat))))
        if inv.seller and inv.seller.name:
            targets.append(("seller_name", [inv.seller.name]))
        if inv.seller and inv.seller.org_number:
            targets.append(("seller_org", [inv.seller.org_number]))
        if inv.seller and inv.seller.vat_number:
            targets.append(("seller_vat", [inv.seller.vat_number]))
        if inv.buyer and inv.buyer.name:
            targets.append(("buyer_name", [inv.buyer.name]))
        if inv.amounts and inv.amounts.currency:
            targets.append(("currency", [inv.amounts.currency]))

        for key, variants in targets:
            if not variants:
                continue
            color = COLORS.get(key, "#ff6f00")
            disp = DISPLAY.get(key, key)
            for pno in range(len(doc)):
                page = doc[pno]
                for val in variants:
                    rects = _search_exact(page, val)
                    if not rects and key in {"kid_number", "seller_org"}:
                        rects = _find_digits_sequence(page, val)
                    for r in rects:
                        add(pno, r, disp, color, "value", key)

    return page_map, key_hits
