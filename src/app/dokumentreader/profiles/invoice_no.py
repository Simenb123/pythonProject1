from __future__ import annotations
from typing import Any, Dict
import os

from app.dokumentreader.invoice_reader import build_invoice_model
from .base import DocProfile, as_result

class InvoiceProfile(DocProfile):
    name = "invoice"
    description = "Norsk/engelsk faktura"

    def detect(self, page1_text: str, full_text: str) -> float:
        hits = 0
        for kw in ["Faktura", "Invoice", "Faktura nr", "Invoice number", "Beløp å betale", "Amount due"]:
            if kw.lower() in page1_text.lower():
                hits += 1
        return min(1.0, hits / 3.0)

    def parse(self, path: str, page1_text: str, full_text: str) -> Dict[str, Any]:
        inv = build_invoice_model(path)
        return as_result("invoice", inv.dict() if hasattr(inv, "dict") else inv.model_dump())
