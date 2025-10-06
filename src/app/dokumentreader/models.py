from __future__ import annotations

from typing import Optional, List
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

class PartyModel(BaseModel):
    name: Optional[str] = None
    org_number: Optional[str] = None
    vat_number: Optional[str] = None
    address: Optional[str] = None

class LineItemModel(BaseModel):
    description: str
    quantity: Optional[Decimal] = None
    unit: Optional[str] = None
    unit_price: Optional[Decimal] = None
    vat_rate: Optional[Decimal] = None  # prosent
    line_total: Optional[Decimal] = None

class TaxBreakdownModel(BaseModel):
    vat_rate: Decimal
    vat_amount: Decimal
    taxable_amount: Optional[Decimal] = None

class PaymentTermsModel(BaseModel):
    due_date: Optional[str] = None  # ISO-dato
    terms_text: Optional[str] = None

class AmountsModel(BaseModel):
    currency: str = Field(default="NOK")
    subtotal_excl_vat: Optional[Decimal] = None
    vat_amount: Optional[Decimal] = None
    total_incl_vat: Optional[Decimal] = None

class InvoiceModel(BaseModel):
    file_name: str
    ocr_used: bool = False
    invoice_number: Optional[str] = None
    order_reference: Optional[str] = None
    kid_number: Optional[str] = None

    invoice_date: Optional[str] = None
    payment_terms: PaymentTermsModel = Field(default_factory=PaymentTermsModel)

    seller: PartyModel = Field(default_factory=PartyModel)
    buyer: PartyModel = Field(default_factory=PartyModel)

    amounts: AmountsModel = Field(default_factory=AmountsModel)
    taxes: List[TaxBreakdownModel] = Field(default_factory=list)
    line_items: List[LineItemModel] = Field(default_factory=list)

    notes: Optional[str] = None

def model_to_json_text(model: BaseModel, pretty: bool = False) -> str:
    indent = 2 if pretty else None
    if _PYDANTIC_V2:
        if hasattr(model, "model_dump_json"):
            return model.model_dump_json(indent=indent)  # type: ignore
        return json.dumps(model.model_dump(mode="json"), ensure_ascii=False, indent=indent)  # type: ignore
    # v1
    return model.json(ensure_ascii=False, indent=indent)  # type: ignore
