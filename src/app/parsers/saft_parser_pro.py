# -*- coding: utf-8 -*-
from __future__ import annotations
"""
SAF-T (NO) parser – v1.3 ready, streaming, refaktorert.

Endringer i denne utgaven:
- Header støtter nå v1.3-feltene SelectionStartDate/SelectionEndDate og DefaultCurrencyCode.
- FileCreationDate tar også AuditFileDateCreated som fallback.
- header.csv får både gamle og nye kolonner (bakoverkompatibelt):
  FunctionalCurrency (fylles av FunctionalCurrency eller DefaultCurrencyCode),
  DefaultCurrencyCode, SelectionStart, SelectionStartDate, SelectionEnd, SelectionEndDate.

Skriver:
- header.csv  (inkl. SelectionStart*/SelectionEnd* og DefaultCurrencyCode)
- accounts.csv  (inkl. GroupingCategory/GroupingCode, Opening*/Closing*)
- tax_table.csv (inkl. StandardTaxCode)
- customers.csv, suppliers.csv
- arap_control_accounts.csv  (BalanceAccountStructure pr Customer/Supplier + konto)
- vouchers.csv (inkl. VoucherType, VoucherDescription, ModificationDate)
- transactions.csv (inkl. DebitTaxAmount, CreditTaxAmount, TaxAmount (fallback), Amount=Debit-Credit)
- analysis_lines.csv
- sales_invoices.csv (inkl. DueDate), purchase_invoices.csv (inkl. DueDate)
- raw_elements.csv (full sporbarhet)

Bruk:
    python saft_parser_pro_fixed.py <input .xml|.zip> <outdir>
    python saft_parser_pro_fixed.py --gui
"""

import argparse
import csv
import decimal
import io
import json
import logging
import zipfile
from dataclasses import dataclass, field
from pathlib import Path
from typing import Iterable, Optional, Dict, Set, BinaryIO

from lxml import etree

# ---------------- Config / helpers ----------------
DEC = decimal.Decimal
decimal.getcontext().prec = 28
BAL_TOL = DEC("0.01")
log = logging.getLogger("saft")
logging.basicConfig(level=logging.INFO, format="%(levelname)s: %(message)s")

NS = {"s": "urn:StandardAuditFile-Taxation-Financial:NO"}  # fungerer også for 1.3

def _lname(el): return etree.QName(el).localname
def _text(el): return (el.text.strip() if (el is not None and el.text) else None)

def _first(el: etree._Element, names: Iterable[str]) -> Optional[str]:
    """Finn første ikke-tomme forekomst av et navn (søker hvor som helst under el, m/ namespace)."""
    for nm in names:
        n = el.find(f".//{{{NS['s']}}}{nm}")
        if n is not None:
            t = _text(n)
            if t not in (None, ""):
                return t
    return None

def _amount_of(el: etree._Element, primary: str) -> Optional[DEC]:
    node = el.find(f".//{{{NS['s']}}}{primary}")
    if node is None:
        return None
    # direkte tekst
    if node.text and node.text.strip():
        try:
            return DEC(node.text.strip().replace(" ", "").replace("\u00A0", ""))
        except Exception:
            return None
    # nested <Amount>
    amt = node.find(f".//{{{NS['s']}}}Amount")
    if amt is not None and amt.text and amt.text.strip():
        try:
            return DEC(amt.text.strip().replace(" ", "").replace("\u00A0", ""))
        except Exception:
            return None
    # som attributt (noen leverandører)
    a = node.get("Amount")
    if a:
        try:
            return DEC(a.strip().replace(" ", "").replace("\u00A0", ""))
        except Exception:
            return None
    return None

def _maybe_open_zip(path: Path) -> BinaryIO:
    if path.suffix.lower() == ".zip":
        z = zipfile.ZipFile(path, "r")
        xmls = [n for n in z.namelist() if n.lower().endswith(".xml")]
        if not xmls:
            raise ValueError("ZIP inneholder ingen .xml")
        return io.BytesIO(z.read(xmls[0]))
    return open(path, "rb")

# ---------------- Data classes ----------------
@dataclass
class VoucherAgg:
    voucher_id: Optional[str] = None
    voucher_no: Optional[str] = None
    transaction_date: Optional[str] = None
    posting_date: Optional[str] = None
    period: Optional[str] = None
    year: Optional[str] = None
    source_doc: Optional[str] = None
    journal_id: Optional[str] = None
    currency_code: Optional[str] = None
    voucher_type: Optional[str] = None
    voucher_desc: Optional[str] = None
    mod_date: Optional[str] = None
    debit: DEC = field(default_factory=lambda: DEC(0))
    credit: DEC = field(default_factory=lambda: DEC(0))

# ---------------- Main parse ----------------
def parse_saft(input_path: Path, outdir: Path) -> None:
    outdir.mkdir(parents=True, exist_ok=True)
    src = _maybe_open_zip(input_path)

    # Writers
    f_header   = open(outdir/"header.csv", "w", newline="", encoding="utf-8")
    f_accounts = open(outdir/"accounts.csv", "w", newline="", encoding="utf-8")
    f_tax      = open(outdir/"tax_table.csv", "w", newline="", encoding="utf-8")
    f_cust     = open(outdir/"customers.csv", "w", newline="", encoding="utf-8")
    f_supp     = open(outdir/"suppliers.csv", "w", newline="", encoding="utf-8")
    f_arapca   = open(outdir/"arap_control_accounts.csv", "w", newline="", encoding="utf-8")
    f_vouch    = open(outdir/"vouchers.csv", "w", newline="", encoding="utf-8")
    f_lines    = open(outdir/"transactions.csv", "w", newline="", encoding="utf-8")
    f_anl      = open(outdir/"analysis_lines.csv", "w", newline="", encoding="utf-8")
    f_sinv     = open(outdir/"sales_invoices.csv", "w", newline="", encoding="utf-8")
    f_pinv     = open(outdir/"purchase_invoices.csv", "w", newline="", encoding="utf-8")
    f_raw      = open(outdir/"raw_elements.csv", "w", newline="", encoding="utf-8")

    # Header: behold gamle kolonner + nye (v1.3)
    w_header = csv.DictWriter(f_header, fieldnames=[
        "CompanyName","CompanyID",
        "FunctionalCurrency","DefaultCurrencyCode",
        "FileCreationDate","AuditFileVersion",
        "SelectionStart","SelectionStartDate","SelectionEnd","SelectionEndDate",
        "StartDate","EndDate",
        "ProductVersion","SoftwareCertificateNumber"
    ])
    w_header.writeheader()

    w_acc = csv.DictWriter(f_accounts, fieldnames=[
        "AccountID","AccountDescription","AccountType","ParentAccountID",
        "GroupingCategory","GroupingCode",
        "OpeningDebit","OpeningCredit","ClosingDebit","ClosingCredit","TaxCode","TaxType"
    ]); w_acc.writeheader()

    w_tax = csv.DictWriter(f_tax, fieldnames=[
        "TaxCode","StandardTaxCode","TaxType","TaxPercentage","TaxCountryRegion","Description"
    ]); w_tax.writeheader()

    w_cust = csv.DictWriter(f_cust, fieldnames=[
        "CustomerID","Name","VATNumber","Country","City","PostalCode","Email","Telephone"
    ]); w_cust.writeheader()

    w_supp = csv.DictWriter(f_supp, fieldnames=[
        "SupplierID","Name","VATNumber","Country","City","PostalCode","Email","Telephone"
    ]); w_supp.writeheader()

    w_arap = csv.DictWriter(f_arapca, fieldnames=[
        "PartyType","PartyID","AccountID","OpeningDebit","OpeningCredit","ClosingDebit","ClosingCredit"
    ]); w_arap.writeheader()

    w_vouch = csv.DictWriter(f_vouch, fieldnames=[
        "VoucherID","VoucherNo","TransactionDate","PostingDate","Period","Year",
        "SourceDocumentID","JournalID","CurrencyCode",
        "VoucherType","VoucherDescription","ModificationDate",
        "DebitTotal","CreditTotal","Balanced"
    ]); w_vouch.writeheader()

    w_lines = csv.DictWriter(f_lines, fieldnames=[
        "RecordID","VoucherID","VoucherNo","JournalID",
        "TransactionDate","PostingDate",
        "SystemID","BatchID","DocumentNumber","LineSourceDocumentID",
        "AccountID","AccountDescription",
        "CustomerID","CustomerName","CustomerVATNumber",
        "SupplierID","SupplierName","SupplierVATNumber",
        "Description","Debit","Credit","Amount",
        "CurrencyCode","AmountCurrency","ExchangeRate",
        "TaxType","TaxCountryRegion","TaxCode","TaxPercentage",
        "DebitTaxAmount","CreditTaxAmount","TaxAmount"
    ]); w_lines.writeheader()

    w_anl = csv.DictWriter(f_anl, fieldnames=["RecordID","Type","ID","Amount"]); w_anl.writeheader()

    w_sinv = csv.DictWriter(f_sinv, fieldnames=[
        "InvoiceNo","InvoiceDate","TaxPointDate","GLPostingDate",
        "CustomerID","CustomerName","CustomerVATNumber",
        "CurrencyCode","NetTotal","TaxPayable","GrossTotal","SourceID","DocumentNumber","DueDate"
    ]); w_sinv.writeheader()

    w_pinv = csv.DictWriter(f_pinv, fieldnames=[
        "InvoiceNo","InvoiceDate","TaxPointDate","GLPostingDate",
        "SupplierID","SupplierName","SupplierVATNumber",
        "CurrencyCode","NetTotal","TaxPayable","GrossTotal","SourceID","DocumentNumber","DueDate"
    ]); w_pinv.writeheader()

    w_raw = csv.DictWriter(f_raw, fieldnames=["XPath","Tag","Text","Attributes"]); w_raw.writeheader()

    # buffers
    accounts: Dict[str, Dict[str, str]] = {}
    customers: Dict[str, Dict[str, str]] = {}
    suppliers: Dict[str, Dict[str, str]] = {}
    cur_voucher: Optional[VoucherAgg] = None

    # streaming parse
    ctx = etree.iterparse(src, events=("start","end"))
    root = None
    for evt, el in ctx:
        tag = _lname(el)

        # rådump (konverter attributter til vanlig dict før json.dumps)
        if evt == "end":
            try:
                xp = el.getroottree().getpath(el)
            except Exception:
                xp = f"/{tag}"
            attrs = {}
            for k, v in el.attrib.items():
                try:
                    key = etree.QName(k).localname
                except Exception:
                    key = str(k)
                attrs[key] = v
            w_raw.writerow({
                "XPath": xp,
                "Tag": tag,
                "Text": (el.text.strip() if el.text else ""),
                "Attributes": json.dumps(attrs, ensure_ascii=False)
            })

        # Header
        if evt == "end" and tag == "Header":
            company = _first(el, ["CompanyName"])
            compid  = _first(el, ["CompanyID"])
            # Fil-dato: flere varianter i omløp
            fcd = _first(el, ["FileCreationDateTime","AuditFileDateCreated","FileCreationDate"])
            ver = _first(el, ["AuditFileVersion"])

            # Periode: støtt både v1.2 og v1.3
            sel_start = _first(el, ["SelectionStart","SelectionStartDate"])
            sel_end   = _first(el, ["SelectionEnd","SelectionEndDate"])
            startd    = _first(el, ["StartDate"])
            endd      = _first(el, ["EndDate"])

            prodver = _first(el, ["ProductVersion"])
            cert    = _first(el, ["SoftwareCertificateNumber"])
            func_cur    = _first(el, ["FunctionalCurrency"])
            default_cur = _first(el, ["DefaultCurrencyCode"])
            currency_out = func_cur or default_cur or ""

            w_header.writerow({
                "CompanyName": company or "",
                "CompanyID": compid or "",
                "FunctionalCurrency": currency_out,
                "DefaultCurrencyCode": default_cur or "",
                "FileCreationDate": fcd or "",
                "AuditFileVersion": ver or "",
                "SelectionStart": sel_start or "",
                "SelectionStartDate": _first(el, ["SelectionStartDate"]) or "",
                "SelectionEnd": sel_end or "",
                "SelectionEndDate": _first(el, ["SelectionEndDate"]) or "",
                "StartDate": startd or "",
                "EndDate": endd or "",
                "ProductVersion": prodver or "",
                "SoftwareCertificateNumber": cert or ""
            })

        # Accounts
        if evt == "end" and tag in ("Account","GeneralLedgerAccount"):
            acc_id  = _first(el, ["AccountID"])
            if acc_id:
                acc_desc= _first(el, ["AccountDescription"])
                acc_type= _first(el, ["AccountType"])
                parent  = _first(el, ["ParentAccountID"])
                group_cat = _first(el, ["GroupingCategory"])
                group_code= _first(el, ["GroupingCode","GroupingCategoryCode"])
                op_dr   = _amount_of(el, "OpeningDebitBalance") or DEC(0)
                op_cr   = _amount_of(el, "OpeningCreditBalance") or DEC(0)
                cl_dr   = _amount_of(el, "ClosingDebitBalance") or DEC(0)
                cl_cr   = _amount_of(el, "ClosingCreditBalance") or DEC(0)
                taxc    = _first(el, ["TaxCode"])
                taxt    = _first(el, ["TaxType"])
                accounts[acc_id] = {"AccountDescription": acc_desc or "", "TaxCode": taxc or ""}
                w_acc.writerow({
                    "AccountID": acc_id, "AccountDescription": acc_desc or "", "AccountType": acc_type or "",
                    "ParentAccountID": parent or "",
                    "GroupingCategory": group_cat or "", "GroupingCode": group_code or "",
                    "OpeningDebit": f"{op_dr}", "OpeningCredit": f"{op_cr}",
                    "ClosingDebit": f"{cl_dr}", "ClosingCredit": f"{cl_cr}",
                    "TaxCode": taxc or "", "TaxType": taxt or ""
                })

        # TaxTable
        if evt == "end" and tag == "TaxTableEntry":
            std = _first(el, ["StandardTaxCode","StandardCode"])
            w_tax.writerow({
                "TaxCode": _first(el, ["TaxCode"]) or "",
                "StandardTaxCode": std or "",
                "TaxType": _first(el, ["TaxType"]) or "",
                "TaxPercentage": _first(el, ["TaxPercentage"]) or "",
                "TaxCountryRegion": _first(el, ["TaxCountryRegion"]) or "",
                "Description": _first(el, ["Description"]) or ""
            })

        # Customers
        if evt == "end" and tag == "Customer":
            cid = _first(el, ["CustomerID"])
            if cid:
                name = _first(el, ["CompanyName","CustomerName","Name"])
                customers[cid] = {"Name": name or "", "VATNumber": _first(el, ["VATNumber"]) or ""}
                w_cust.writerow({
                    "CustomerID": cid, "Name": name or "", "VATNumber": customers[cid]["VATNumber"],
                    "Country": _first(el, ["Country"]) or "", "City": _first(el, ["City"]) or "",
                    "PostalCode": _first(el, ["PostalCode"]) or "",
                    "Email": _first(el, ["Email"]) or "", "Telephone": _first(el, ["Telephone"]) or ""
                })
                # BalanceAccountStructure (1.3)
                for b in el.findall(f".//{{{NS['s']}}}BalanceAccountStructure"):
                    w_arap.writerow({
                        "PartyType":"Customer","PartyID":cid,
                        "AccountID": _first(b, ["AccountID"]) or "",
                        "OpeningDebit": f"{_amount_of(b,'OpeningDebitBalance') or DEC(0)}",
                        "OpeningCredit": f"{_amount_of(b,'OpeningCreditBalance') or DEC(0)}",
                        "ClosingDebit": f"{_amount_of(b,'ClosingDebitBalance') or DEC(0)}",
                        "ClosingCredit": f"{_amount_of(b,'ClosingCreditBalance') or DEC(0)}",
                    })

        # Suppliers
        if evt == "end" and tag == "Supplier":
            sid = _first(el, ["SupplierID"])
            if sid:
                name = _first(el, ["CompanyName","SupplierName","Name"])
                suppliers[sid] = {"Name": name or "", "VATNumber": _first(el, ["VATNumber"]) or ""}
                w_supp.writerow({
                    "SupplierID": sid, "Name": name or "", "VATNumber": suppliers[sid]["VATNumber"],
                    "Country": _first(el, ["Country"]) or "", "City": _first(el, ["City"]) or "",
                    "PostalCode": _first(el, ["PostalCode"]) or "",
                    "Email": _first(el, ["Email"]) or "", "Telephone": _first(el, ["Telephone"]) or ""
                })
                for b in el.findall(f".//{{{NS['s']}}}BalanceAccountStructure"):
                    w_arap.writerow({
                        "PartyType":"Supplier","PartyID":sid,
                        "AccountID": _first(b, ["AccountID"]) or "",
                        "OpeningDebit": f"{_amount_of(b,'OpeningDebitBalance') or DEC(0)}",
                        "OpeningCredit": f"{_amount_of(b,'OpeningCreditBalance') or DEC(0)}",
                        "ClosingDebit": f"{_amount_of(b,'ClosingDebitBalance') or DEC(0)}",
                        "ClosingCredit": f"{_amount_of(b,'ClosingCreditBalance') or DEC(0)}",
                    })

        # Start Transaction (Voucher)
        if evt == "start" and tag == "Transaction":
            cur_voucher = VoucherAgg()
            cur_voucher.voucher_id = _first(el, ["TransactionID","TransactionNo","VoucherID"])
            cur_voucher.voucher_no = _first(el, ["VoucherNo","TransactionNo","TransactionID"])
            cur_voucher.transaction_date = _first(el, ["TransactionDate","EntryDate"])
            cur_voucher.posting_date = _first(el, ["PostingDate"])
            cur_voucher.period = _first(el, ["Period"])
            cur_voucher.year = _first(el, ["FiscalYear","Year"])
            cur_voucher.source_doc = _first(el, ["SourceDocumentID","SourceID","DocumentNumber"])
            cur_voucher.journal_id = _first(el, ["JournalID","Journal"])
            cur_voucher.currency_code = _first(el, ["CurrencyCode","TransactionCurrency"])
            # v1.3 optional fields
            cur_voucher.voucher_type = _first(el, ["VoucherType"])
            cur_voucher.voucher_desc = _first(el, ["VoucherDescription"])
            cur_voucher.mod_date     = _first(el, ["ModificationDate"])

        # Line variants
        if evt == "end" and tag in ("Line", "TransactionLine", "JournalLine"):
            # Only process transaction lines that are part of a <Transaction> (GeneralLedgerEntries).
            # SAF‑T SourceDocuments (Sales/Purchase invoices) also have <Line> elements which do not
            # represent journal lines and often lack AccountID. To avoid polluting transactions.csv
            # with these invoice lines, skip if we're not currently inside a Transaction.
            if cur_voucher is None:
                # Even if we skip writing, we still need to release memory for the element to avoid
                # building up the XML tree. Clearing here mirrors the logic below.
                el.clear()
                while el.getprevious() is not None:
                    del el.getparent()[0]
                continue
            # context comes from current voucher
            v_id = cur_voucher.voucher_id
            v_no = cur_voucher.voucher_no
            j_id = cur_voucher.journal_id
            currency = cur_voucher.currency_code
            tr_date = cur_voucher.transaction_date
            p_date = cur_voucher.posting_date

            record_id = _first(el, ["RecordID", "LineID"])
            system_id = _first(el, ["SystemID"])
            batch_id = _first(el, ["BatchID"])
            doc_no = _first(el, ["DocumentNumber"])
            line_src = _first(el, ["SourceDocumentID"])
            acc_id = _first(el, ["AccountID"])
            cust_id = _first(el, ["CustomerID"])
            sup_id = _first(el, ["SupplierID"])
            desc = _first(el, ["Description", "Narrative", "LineDescription"])

            debit = _amount_of(el, "DebitAmount") or DEC(0)
            credit = _amount_of(el, "CreditAmount") or DEC(0)
            amount = debit - credit

            # Always update voucher totals, even for lines without AccountID, since these are part of the
            # journal totals. However, we will not write such lines to transactions.csv to avoid
            # contaminating the ledger with non-journal details.  If AccountID is missing, treat the
            # line as an analysis-line and skip writing it to the transactions file. The Analysis
            # section (handled later) will capture the detailed breakdown.

            # Update voucher totals first
            cur_voucher.debit += debit
            cur_voucher.credit += credit

            # Skip writing to transactions.csv if AccountID is missing or blank
            if not acc_id:
                # Clear memory for the element and continue
                el.clear()
                while el.getprevious() is not None:
                    del el.getparent()[0]
                continue

            amt_cur = _first(el, ["AmountCurrency", "ForeignAmount"])
            ex_rate = _first(el, ["ExchangeRate"])
            tax_type = _first(el, ["TaxType"])
            tax_country = _first(el, ["TaxCountryRegion"])
            tax_code = _first(el, ["TaxCode"])
            tax_perc = _first(el, ["TaxPercentage"])
            # v1.3 tax split
            d_tax = _amount_of(el, "DebitTaxAmount") or DEC(0)
            c_tax = _amount_of(el, "CreditTaxAmount") or DEC(0)
            tax_amt = _amount_of(el, "TaxAmount") or (d_tax - c_tax)

            acc_desc = accounts.get(acc_id, {}).get("AccountDescription", "") if acc_id else ""
            cust_name = customers.get(cust_id, {}).get("Name", "") if cust_id else ""
            cust_vat = customers.get(cust_id, {}).get("VATNumber", "") if cust_id else ""
            sup_name = suppliers.get(sup_id, {}).get("Name", "") if sup_id else ""
            sup_vat = suppliers.get(sup_id, {}).get("VATNumber", "") if sup_id else ""

            w_lines.writerow(
                {
                    "RecordID": record_id or "",
                    "VoucherID": v_id or "",
                    "VoucherNo": v_no or "",
                    "JournalID": j_id or "",
                    "TransactionDate": tr_date or "",
                    "PostingDate": p_date or "",
                    "SystemID": system_id or "",
                    "BatchID": batch_id or "",
                    "DocumentNumber": doc_no or "",
                    "LineSourceDocumentID": line_src or "",
                    "AccountID": acc_id or "",
                    "AccountDescription": acc_desc,
                    "CustomerID": cust_id or "",
                    "CustomerName": cust_name,
                    "CustomerVATNumber": cust_vat,
                    "SupplierID": sup_id or "",
                    "SupplierName": sup_name,
                    "SupplierVATNumber": sup_vat,
                    "Description": desc or "",
                    "Debit": f"{debit}",
                    "Credit": f"{credit}",
                    "Amount": f"{amount}",
                    "CurrencyCode": currency or "",
                    "AmountCurrency": amt_cur or "",
                    "ExchangeRate": ex_rate or "",
                    "TaxType": tax_type or "",
                    "TaxCountryRegion": tax_country or "",
                    "TaxCode": tax_code or "",
                    "TaxPercentage": tax_perc or "",
                    "DebitTaxAmount": f"{d_tax}",
                    "CreditTaxAmount": f"{c_tax}",
                    "TaxAmount": f"{tax_amt}",
                }
            )

            # free memory
            el.clear()
            while el.getprevious() is not None:
                del el.getparent()[0]

        # Analysis connected to nearest line
        if evt == "end" and tag == "Analysis":
            parent = el.getparent()
            while parent is not None and _lname(parent) not in ("Line","TransactionLine","JournalLine"):
                parent = parent.getparent()
            rec_id = _first(parent, ["RecordID","LineID"]) if parent is not None else None
            w_anl.writerow({
                "RecordID": rec_id or "", "Type": _first(el, ["AnalysisType"]) or "",
                "ID": _first(el, ["AnalysisID"]) or "", "Amount": f"{_amount_of(el,'Amount') or DEC(0)}"
            })
            el.clear()
            while el.getprevious() is not None: del el.getparent()[0]

        # SourceDocuments – Sales & Purchase (for DueDate i aging)
        if evt == "end" and tag == "Invoice":
            parent = _lname(el.getparent()) if el.getparent() is not None else ""
            inv = {
                "InvoiceNo": _first(el, ["InvoiceNo","InvoiceNumber"]) or "",
                "InvoiceDate": _first(el, ["InvoiceDate"]) or "",
                "TaxPointDate": _first(el, ["TaxPointDate"]) or "",
                "GLPostingDate": _first(el, ["GLPostingDate"]) or "",
                "CurrencyCode": _first(el, ["CurrencyCode","TransactionCurrency"]) or "",
                "NetTotal": _first(el, ["NetTotal","DocumentNetTotal"]) or "",
                "TaxPayable": _first(el, ["TaxPayable","DocumentTaxPayable"]) or "",
                "GrossTotal": _first(el, ["GrossTotal","DocumentGrossTotal"]) or "",
                "SourceID": _first(el, ["SourceID"]) or "",
                "DocumentNumber": _first(el, ["DocumentNumber"]) or "",
                "DueDate": _first(el, ["DueDate"]) or "",
            }
            if parent == "SalesInvoices":
                cid = _first(el, ["CustomerID"])
                cname = customers.get(cid,{}).get("Name","") if cid else _first(el,["CustomerName"])
                inv.update({"CustomerID":cid or "","CustomerName":cname or "","CustomerVATNumber":customers.get(cid,{}).get("VATNumber","")})
                w_sinv.writerow(inv)
            elif parent == "PurchaseInvoices":
                sid = _first(el, ["SupplierID"])
                sname = suppliers.get(sid,{}).get("Name","") if sid else _first(el,["SupplierName"])
                inv.update({"SupplierID":sid or "","SupplierName":sname or "","SupplierVATNumber":suppliers.get(sid,{}).get("VATNumber","")})
                w_pinv.writerow(inv)

        # End Transaction
        if evt == "end" and tag == "Transaction":
            if cur_voucher is not None:
                balanced = abs(cur_voucher.debit - cur_voucher.credit) <= BAL_TOL
                w_vouch.writerow({
                    "VoucherID": cur_voucher.voucher_id or "", "VoucherNo": cur_voucher.voucher_no or "",
                    "TransactionDate": cur_voucher.transaction_date or "", "PostingDate": cur_voucher.posting_date or "",
                    "Period": cur_voucher.period or "", "Year": cur_voucher.year or "",
                    "SourceDocumentID": cur_voucher.source_doc or "", "JournalID": cur_voucher.journal_id or "",
                    "CurrencyCode": cur_voucher.currency_code or "",
                    "VoucherType": cur_voucher.voucher_type or "", "VoucherDescription": cur_voucher.voucher_desc or "",
                    "ModificationDate": cur_voucher.mod_date or "",
                    "DebitTotal": f"{cur_voucher.debit}", "CreditTotal": f"{cur_voucher.credit}", "Balanced": "Y" if balanced else "N"
                })
            cur_voucher = None
            el.clear()
            while el.getprevious() is not None: del el.getparent()[0]

        # stream root GC
        if root is None:
            root = el.getroottree().getroot()
        if evt == "end" and el == root:
            root.clear()

    for fh in (f_header,f_accounts,f_tax,f_cust,f_supp,f_arapca,f_vouch,f_lines,f_anl,f_sinv,f_pinv,f_raw):
        fh.close()
    log.info("Ferdig: %s", outdir)

# ---------------- GUI wrapper (enkel) ----------------
def launch_gui():
    import tkinter as tk
    from tkinter import filedialog, messagebox
    root = tk.Tk(); root.title("SAF-T Pro Parser (v1.3 ready)")
    def run():
        p = filedialog.askopenfilename(title="Velg SAF-T", filetypes=[("SAF-T/XML/ZIP","*.xml *.zip")])
        if not p: return
        out = filedialog.askdirectory(title="Velg output-mappe")
        if not out: return
        try:
            parse_saft(Path(p), Path(out))
            messagebox.showinfo("OK", f"Ferdig: {out}")
        except Exception as e:
            messagebox.showerror("Feil", str(e))
    tk.Button(root, text="Kjør parsing…", command=run, width=32).pack(padx=20,pady=20)
    root.mainloop()

def main(argv=None) -> int:
    p = argparse.ArgumentParser()
    p.add_argument("input", nargs="?", help="SAF-T .xml eller .zip")
    p.add_argument("outdir", nargs="?", help="Output-mappe")
    p.add_argument("--gui", action="store_true", help="Start enkel GUI")
    args = p.parse_args(argv)
    if args.gui: launch_gui(); return 0
    if not args.input or not args.outdir: p.print_help(); return 2
    parse_saft(Path(args.input), Path(args.outdir)); return 0

if __name__ == "__main__":
    raise SystemExit(main())
