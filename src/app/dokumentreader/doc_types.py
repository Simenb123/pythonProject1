from __future__ import annotations

# --- bootstrap ---
import os, sys
if __package__ is None or not __package__.startswith("app.dokumentreader"):
    BASE = os.path.abspath(os.path.join(os.path.dirname(__file__), "../../"))
    if BASE not in sys.path:
        sys.path.insert(0, BASE)
# ------------------

from enum import Enum
from typing import Optional, Dict, List
from decimal import Decimal
import json

# Pydantic v1/v2-kompat
try:
    from pydantic import BaseModel, Field
    _PYDANTIC_V2 = True
except Exception:
    from pydantic import BaseModel  # type: ignore
    _PYDANTIC_V2 = False
    def Field(default=None, **kwargs):  # type: ignore
        return default


class DocumentType(str, Enum):
    INVOICE = "invoice"
    FINANCIAL_STATEMENT = "financial_statement"
    TAX_RETURN = "tax_return"
    UNKNOWN = "unknown"


class KVPair(BaseModel):
    key: str
    value: Optional[str] = None


class FinancialStatementModel(BaseModel):
    """Forenklet årsregnskap; vi deler i Resultat & Balanse som key->amount."""
    company: Optional[str] = None
    orgnr: Optional[str] = None
    period: Optional[str] = None
    currency: Optional[str] = "NOK"

    income_statement: Dict[str, Decimal] = Field(default_factory=dict)  # f.eks. {"Driftsinntekter": 12345.00}
    balance_sheet: Dict[str, Decimal] = Field(default_factory=dict)     # f.eks. {"Sum eiendeler": 9999.00}


class TaxReturnModel(BaseModel):
    """Generisk skattemelding (nøkkel/verdi + eventuelle 'poster')."""
    taxpayer_name: Optional[str] = None
    orgnr: Optional[str] = None
    income_year: Optional[str] = None
    currency: Optional[str] = "NOK"

    fields: List[KVPair] = Field(default_factory=list)  # fri nøkkel/verdi
    posts: Dict[str, Decimal] = Field(default_factory=dict)  # f.eks. {"Post 200 Skattepliktig inntekt": 1000.00}


class DocumentEnvelope(BaseModel):
    """Konvolutt som sier hva slags dokument og legger ved relevant modell."""
    file_name: str
    ocr_used: bool = False
    doc_type: DocumentType

    # One-of (kun ett av disse blir fylt ut)
    invoice: Optional[dict] = None            # reuse: models.InvoiceModel.model_dump()
    financials: Optional[FinancialStatementModel] = None
    tax_return: Optional[TaxReturnModel] = None

    raw_text_excerpt: Optional[str] = None  # for feilsøking i GUI

def model_to_json_text(model: BaseModel, pretty: bool = False) -> str:
    indent = 2 if pretty else None
    if _PYDANTIC_V2:
        if hasattr(model, "model_dump_json"):
            return model.model_dump_json(indent=indent)  # type: ignore
        return json.dumps(model.model_dump(mode="json"), ensure_ascii=False, indent=indent)  # type: ignore
    return model.json(ensure_ascii=False, indent=indent)  # type: ignore
