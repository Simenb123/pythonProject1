from __future__ import annotations

# --- BOOTSTRAP så scriptet kan kjøres direkte med Run ---
import os, sys
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# --------------------------------------------------------

import argparse
import logging
import os
from decimal import Decimal
from typing import List

from app.dokumentreader.extractors import (
    extract_text_blocks_from_pdf, extract_text_from_image, extract_tables, ExtractedText
)
from app.dokumentreader.parsers import (
    extract_key_fields, parse_line_items_from_tables, parse_line_items_from_text, pick_parties
)
from app.dokumentreader.models import (
    InvoiceModel, AmountsModel, PaymentTermsModel, PartyModel, TaxBreakdownModel, model_to_json_text
)

def read_any(path: str) -> ExtractedText:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        return extract_text_blocks_from_pdf(path)
    elif ext in (".png", ".jpg", ".jpeg", ".tif", ".tiff", ".bmp"):
        return extract_text_from_image(path)
    else:
        raise ValueError(f"Ukjent/ikke-støttet filtype: {ext}")

def _page1_text_if_pdf(path: str) -> str:
    if not path.lower().endswith(".pdf"):
        return ""
    try:
        import fitz
        with fitz.open(path) as doc:
            if len(doc) > 0:
                return doc[0].get_text() or ""
    except Exception:
        pass
    return ""

def build_invoice_model(path: str) -> InvoiceModel:
    et = read_any(path)
    text = et.text or ""

    # Nøkkelfelt: prioriter side 1 hvis PDF
    text_page1 = _page1_text_if_pdf(path)
    kf = extract_key_fields(text_page1 or text)
    currency = kf.get("currency", "NOK")

    # Tabeller -> linjeposter (eller streng tekst-fallback)
    tables = []
    if path.lower().endswith(".pdf"):
        try:
            tables = extract_tables(path)
        except Exception:
            tables = []
    line_items = parse_line_items_from_tables(tables) if tables else parse_line_items_from_text(text_page1 or text)

    # Selger/kjøper – bruk også tekst fra side 1
    seller, buyer = pick_parties(text_page1 or text)

    # Summer
    amounts = AmountsModel(
        currency=currency,
        subtotal_excl_vat=kf.get("subtotal_excl_vat"),
        vat_amount=kf.get("vat_amount"),
        total_incl_vat=kf.get("total_incl_vat"),
    )
    if amounts.total_incl_vat is None and amounts.subtotal_excl_vat is not None and amounts.vat_amount is not None:
        amounts.total_incl_vat = (amounts.subtotal_excl_vat + amounts.vat_amount).quantize(Decimal("0.01"))
    if amounts.subtotal_excl_vat is None and amounts.total_incl_vat is not None and amounts.vat_amount is not None:
        amounts.subtotal_excl_vat = (amounts.total_incl_vat - amounts.vat_amount).quantize(Decimal("0.01"))

    # MVA-oppsummering (én totalpost – kan utvides per sats)
    taxes: List[TaxBreakdownModel] = []
    if amounts.vat_amount is not None:
        main_rate = None
        for li in line_items:
            if li.vat_rate is not None:
                main_rate = li.vat_rate
                break
        if main_rate is None:
            main_rate = Decimal("25")
        taxes.append(TaxBreakdownModel(vat_rate=main_rate, vat_amount=amounts.vat_amount))

    # Betalingsvilkår
    payment = PaymentTermsModel(
        due_date=kf.get("due_date"),
        terms_text=kf.get("terms_text"),
    )

    inv = InvoiceModel(
        file_name=os.path.basename(path),
        ocr_used=et.ocr_used,
        invoice_number=kf.get("invoice_number"),
        order_reference=kf.get("order_reference"),
        kid_number=kf.get("kid_number"),
        invoice_date=kf.get("invoice_date"),
        payment_terms=payment,
        seller=PartyModel(
            name=seller.name,
            org_number=kf.get("seller_orgnr"),
            vat_number=kf.get("seller_vatno"),
            address=seller.address
        ),
        buyer=buyer,
        amounts=amounts,
        taxes=taxes,
        line_items=line_items,
        notes=None
    )

    # Konsistenskontroll: linjesum ≈ total
    try:
        if inv.amounts.total_incl_vat and inv.line_items:
            calc_sum = Decimal("0.00")
            for li in inv.line_items:
                if li.line_total is not None:
                    calc_sum += li.line_total
                elif li.unit_price is not None and li.quantity is not None:
                    calc_sum += (li.unit_price * li.quantity)
            if calc_sum and abs(calc_sum - inv.amounts.total_incl_vat) > Decimal("5.00"):
                inv.notes = (inv.notes or "") + f" Varsel: Linjesum ({calc_sum}) avviker fra total ({inv.amounts.total_incl_vat})."
    except Exception:
        pass

    return inv

def main():
    ap = argparse.ArgumentParser(description="Les faktura og skriv strukturert JSON.")
    ap.add_argument("infile", nargs="?", help="PDF/JPG/PNG av faktura (valgfri – filvelger åpnes hvis utelatt)")
    ap.add_argument("--out", help="Skriv JSON til fil")
    ap.add_argument("--pretty", action="store_true", help="Pen JSON (indentert)")
    ap.add_argument("--loglevel", default="WARNING", help="Loggnivå: DEBUG/INFO/WARNING/ERROR")
    args = ap.parse_args()

    logging.basicConfig(level=getattr(logging, args.loglevel.upper(), logging.WARNING))

    # Hvis ikke oppgitt: filvelger (GUI)
    if not args.infile:
        try:
            from tkinter import Tk, filedialog
            root = Tk(); root.withdraw()
            sel = filedialog.askopenfilename(
                title="Velg faktura",
                filetypes=[("Dokumenter", "*.pdf *.png *.jpg *.jpeg *.tif *.tiff")]
            )
            root.update(); root.destroy()
            if not sel:
                ap.error("infile er påkrevd (angi en fil eller velg i dialogen).")
            args.infile = sel
        except Exception:
            ap.error("infile er påkrevd. Eksempel: bare trykk Run og velg en fil, eller angi C:\\path\\faktura.pdf")

    inv = build_invoice_model(args.infile)
    txt = model_to_json_text(inv, pretty=args.pretty)

    if args.out:
        os.makedirs(os.path.dirname(os.path.abspath(args.out)), exist_ok=True)
        with open(args.out, "w", encoding="utf-8") as f:
            f.write(txt)
        print(f"Skrev {args.out}")
    else:
        print(txt)

if __name__ == "__main__":
    main()
