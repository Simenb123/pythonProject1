# -*- coding: utf-8 -*-
from __future__ import annotations
import os, sys

# --- bootstrap ---
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# ------------------

from typing import Tuple, Optional
from app.dokumentreader.extractors import extract_text_blocks_from_pdf, extract_text_from_image
from app.dokumentreader.classifier import classify_text
from app.dokumentreader.doc_types import DocumentType, DocumentEnvelope, FinancialStatementModel, TaxReturnModel, KVPair
from app.dokumentreader.invoice_reader import build_invoice_model
from app.dokumentreader.parsers_financials import parse_financial_statement
from app.dokumentreader.parsers_tax import parse_tax_return
from app.dokumentreader.template_engine import apply_templates


def read_text_and_ocr(path: str) -> Tuple[str, bool]:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pdf":
        et = extract_text_blocks_from_pdf(path)
    else:
        et = extract_text_from_image(path)
    return et.text or "", et.ocr_used


def _parse_invoice(path: str, ocr: bool) -> DocumentEnvelope:
    inv = build_invoice_model(path)
    return DocumentEnvelope(
        file_name=os.path.basename(path),
        ocr_used=ocr or inv.ocr_used,
        doc_type=DocumentType.INVOICE,
        invoice=inv.dict() if hasattr(inv, "dict") else inv.model_dump(),
        raw_text_excerpt=None
    )


def _parse_financials(path: str, txt: str, ocr: bool) -> DocumentEnvelope:
    d = parse_financial_statement(txt)
    fin = FinancialStatementModel(
        currency=d.get("currency"),
        income_statement=d.get("income_statement") or {},
        balance_sheet=d.get("balance_sheet") or {},
    )
    return DocumentEnvelope(
        file_name=os.path.basename(path),
        ocr_used=ocr,
        doc_type=DocumentType.FINANCIAL_STATEMENT,
        financials=fin,
        raw_text_excerpt=(txt[:2000] if txt else None)
    )


def _parse_tax(path: str, txt: str, ocr: bool) -> DocumentEnvelope:
    d = parse_tax_return(txt)
    tr = TaxReturnModel(
        currency=d.get("currency"),
        taxpayer_name=d.get("taxpayer_name"),
        orgnr=d.get("orgnr"),
        income_year=d.get("income_year"),
        fields=[*d.get("fields", [])],
        posts=d.get("posts") or {},
    )
    # --- MALER: bruk lærte ankre som ekstra felter ---
    res = apply_templates(path, DocumentType.TAX_RETURN.value, full_text=txt)
    if res.get("fields"):
        for kv in res["fields"]:
            try:
                tr.fields.append(KVPair(key=kv.get("key"), value=str(kv.get("value"))).dict())  # pydantic v1/v2-kompat
            except Exception:
                tr.fields.append({"key": kv.get("key"), "value": str(kv.get("value"))})

    return DocumentEnvelope(
        file_name=os.path.basename(path),
        ocr_used=ocr,
        doc_type=DocumentType.TAX_RETURN,
        tax_return=tr,
        raw_text_excerpt=(txt[:2000] if txt else None)
    )


def parse_document(path: str, force_profile: Optional[str] = None) -> DocumentEnvelope:
    """
    Orchestrator:
      - force_profile in {"invoice","financials_no","vat_return_no"} tvinger parser
      - ellers klassifiseres dokumentet automatisk
      - etter ordinær parsing, prøver vi maler (Admin/Lær) for å hente ekstra felt
    """
    txt, ocr = read_text_and_ocr(path)

    if force_profile == "invoice":
        return _parse_invoice(path, ocr)
    elif force_profile == "financials_no":
        return _parse_financials(path, txt, ocr)
    elif force_profile == "vat_return_no":
        return _parse_tax(path, txt, ocr)

    doc_type, _conf = classify_text(txt)
    if doc_type == DocumentType.INVOICE:
        env = _parse_invoice(path, ocr)
    elif doc_type == DocumentType.FINANCIAL_STATEMENT:
        env = _parse_financials(path, txt, ocr)
    elif doc_type == DocumentType.TAX_RETURN:
        env = _parse_tax(path, txt, ocr)
    else:
        env = DocumentEnvelope(
            file_name=os.path.basename(path),
            ocr_used=ocr,
            doc_type=DocumentType.UNKNOWN,
            raw_text_excerpt=(txt[:2000] if txt else None)
        )

    # MALER for faktura/regnskap (valgfritt): kan aktiveres på samme måte:
    # res = apply_templates(path, env.doc_type.value, full_text=txt)

    return env
