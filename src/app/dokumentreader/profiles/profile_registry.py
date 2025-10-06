from __future__ import annotations
from typing import List, Tuple, Dict
from app.dokumentreader.profiles.base import DocProfile
from app.dokumentreader.profiles.invoice_no import InvoiceProfile
from app.dokumentreader.profiles.financials_no import FinancialsNoProfile
from app.dokumentreader.profiles.vat_return_no import VatReturnNoProfile

ALL_PROFILES: List[DocProfile] = [
    InvoiceProfile(),        # faktura
    FinancialsNoProfile(),   # resultat/balanse
    VatReturnNoProfile(),    # skattemelding MVA
]

def pick_best_profile(page1_text: str, full_text: str) -> Tuple[DocProfile, float]:
    scored = [(p, p.detect(page1_text, full_text)) for p in ALL_PROFILES]
    scored.sort(key=lambda x: x[1], reverse=True)
    return scored[0]
