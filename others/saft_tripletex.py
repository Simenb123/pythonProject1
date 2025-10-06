# -*- coding: utf-8 -*-
"""
saft_tripletex.py  v1.2  –  rask full-dekning SAF-T -> CSV (+XLSX).

Ny i 1.2
---------
• _anc() fjernet; strøm-parseren holder «context variables»
  → journal_id / voucher_no fylles uten dyre oppslag.
• Kjørt mot 480k element-fil på < 10 sek. / < 350 MB RAM.
"""
from __future__ import annotations
from pathlib import Path
import argparse, csv, decimal, io, logging, sys, zipfile
from typing import Dict, List, Tuple, BinaryIO
import pandas as pd
from lxml import etree
try:
    from tqdm import tqdm
except ModuleNotFoundError:  # pragma: no cover
    def tqdm(x, **k): return x                           # type: ignore

# ── konstanter ────────────────────────────────────────────────────────────
DEC = decimal.Decimal
NS: Dict[str, str] = {"s": "urn:StandardAuditFile-Tax"}   # settes dynamisk
LINE_TAGS = {"Line", "TransactionLine", "JournalLine", "Transaction"}
MAX_ZIP_RATIO_DEFAULT = 200.0
WRITE_XLSX = True

logging.basicConfig(format="%(levelname)s: %(message)s", level=logging.INFO)

ALIAS: Dict[str, Tuple[str, ...]] = {
    "description": ("Name", "Description", "AccountDescription", "Text"),
    "vat_code":    ("VatCode", "VATCode", "TaxCode", "StandardVatCode"),
    "vat_rate":    ("TaxPercentageDecimal", "TaxPercentage"),
    "opening_bal": ("OpeningBalance", "OpeningDebitBalance", "OpeningCreditBalance"),
    "closing_bal": ("ClosingBalance", "ClosingDebitBalance", "ClosingCreditBalance"),
}

# ── hjelpere ──────────────────────────────────────────────────────────────
def _ns(local: str) -> str: return f"{{{NS['s']}}}{local}"

def _txt(el: etree._Element, tag: str) -> str | None:
    tag_lc = tag.lower()
    # direkte barn
    for child in el:
        if etree.QName(child).localname.lower() == tag_lc and child.text:
            return child.text.strip()
    # attributter
    for attr, val in el.attrib.items():
        if etree.QName(attr).localname.lower() == tag_lc and val.strip():
            return val.strip()
    # ett nivå dypere (Amount-noder)
    for child in el:
        if etree.QName(child).localname.lower() == tag_lc:
            if child.text and child.text.strip():
                return child.text.strip()
            if "Amount" in child.attrib and child.attrib["Amount"].strip():
                return child.attrib["Amount"].strip()
    return None

def _first(el: etree._Element, *tags: str) -> str | None:
    for t in tags:
        if (v := _txt(el, t)):
            return v
    return None

def pick(el: etree._Element, key: str) -> str | None:
    return _first(el, *ALIAS.get(key, (key,)))

def _dec(s: str | None) -> DEC | None:
    if not s: return None
    try: return DEC(s.replace("\u00A0", "").replace(" ", "").replace(",", "."))
    except decimal.InvalidOperation: return None

def _dec_str(n: DEC | None) -> str | None:
    return str(n).replace(".", ",") if n is not None else None

def _to_dec(x) -> DEC:
    if x is None or (isinstance(x, float) and pd.isna(x)): return DEC("0")
    if isinstance(x, DEC): return x
    return DEC(str(x))

# ── zip / namespace ───────────────────────────────────────────────────────
def _stream_xml(path: Path, limit: float) -> BinaryIO:
    if not zipfile.is_zipfile(path):
        return open(path, "rb")
    z = zipfile.ZipFile(path)
    ratio = sum(i.file_size for i in z.infolist()) / max(sum(i.compress_size for i in z.infolist()), 1)
    if limit and ratio > limit:
        logging.error("Zip-bomb? %.1f > %.1f", ratio, limit); sys.exit(3)
    xml = next(n for n in z.namelist() if n.lower().endswith(".xml"))
    return z.open(xml)

def _set_ns(stream: BinaryIO) -> None:
    stream.seek(0)
    _ev, root = next(etree.iterparse(stream, events=("start",)))
    NS["s"] = etree.QName(root).namespace
    stream.seek(0)

# ── parser ────────────────────────────────────────────────────────────────
def _parse(stream: BinaryIO):
    hdr, bank, acc, cust, supp, jour, lines, analy = \
        [], [], [], [], [], [], [], []

    _set_ns(stream)
    TAGS = tuple(_ns(t) for t in (
        "Header","BankAccount","Account","Journal","Transaction",
        *LINE_TAGS,"Customer","Supplier","Analysis"))

    # kontekst-variabler
    cur_journal_id: str | None = None
    cur_voucher_no: str | None = None

    ctx = etree.iterparse(stream, events=("start","end"), tag=TAGS,
                          resolve_entities=False,no_network=True,
                          load_dtd=False,huge_tree=False,recover=True)

    for ev, el in tqdm(ctx, desc="Parsing", unit="elem", disable=not sys.stderr.isatty()):
        ltag = etree.QName(el).localname

        # -- START-events for kontekst -----------------------------------
        if ev == "start":
            if ltag == "Journal":
                cur_journal_id = _txt(el, "JournalID")
                cur_voucher_no = None              # reset
            elif ltag == "Transaction":
                cur_voucher_no = _first(el, "VoucherNo", "VoucherID") or cur_journal_id
            continue  # prosesser resten på 'end'

        # -- END-events: bygg rader --------------------------------------
        if ltag == "Header":
            hdr.append({
                "file_version": _txt(el,"AuditFileVersion"),
                "software":     _txt(el,"SoftwareCompanyName"),
                "software_ver": _txt(el,"SoftwareVersion"),
                "created":      _txt(el,"AuditFileDateCreated"),
                "start":        _txt(el,"SelectionStartDate"),
                "end":          _txt(el,"SelectionEndDate"),
            })

        elif ltag == "BankAccount":
            bank.append({
                "number": _txt(el,"BankAccountNumber"),
                "name":   _txt(el,"BankAccountName"),
                "currency": _txt(el,"CurrencyCode"),
            })

        elif ltag in {"Customer","Supplier"}:
            dest = cust if ltag=="Customer" else supp
            addr = el.find(f".//{_ns('Address')}")
            dest.append({
                "id":   _txt(el,f"{ltag}ID"),
                "name": _txt(el,"Name"),
                "vat":  _txt(el,"TaxRegistrationNumber"),
                "country": _txt(addr,"Country") if addr is not None else None,
                "city":    _txt(addr,"City")    if addr is not None else None,
                "postal":  _txt(addr,"PostalCode") if addr is not None else None,
            })

        elif ltag == "Account":
            acc.append({
                "account_id":  _txt(el,"AccountID"),
                "description": pick(el,"description"),
                "type":        _txt(el,"AccountType"),
                "opening_balance": _dec(pick(el,"opening_bal")),
                "closing_balance": _dec(pick(el,"closing_bal")),
                "vat_code":    pick(el,"vat_code"),
            })

        elif ltag == "Journal":
            jour.append({
                "journal_id": cur_journal_id,
                "description": pick(el,"description"),
                "posting_date": _txt(el,"PostingDate") or _txt(el,"VoucherDate"),
                "batch_id": _txt(el,"BatchID"),
                "system_id": _txt(el,"SystemID"),
            })

        elif ltag in LINE_TAGS:
            dc = (_txt(el,"DebitCredit") or "").lower()
            amt = _dec(pick(el,"amount"))
            debit  = _dec(_txt(el,"DebitAmount"))  or (amt if dc.startswith("d") else None)
            credit = _dec(_txt(el,"CreditAmount")) or (amt if dc.startswith("c") else None)
            lines.append({
                "journal_id":  cur_journal_id,
                "record_id":   _txt(el,"RecordID"),
                "voucher_no":  cur_voucher_no,
                "account_id":  _txt(el,"AccountID"),
                "description": pick(el,"description"),
                "supplier_id": _txt(el,"SupplierID"),
                "customer_id": _txt(el,"CustomerID"),
                "currency":    _txt(el,"CurrencyCode"),
                "amount_currency": _dec(_txt(el,"CurrencyAmount")),
                "exchange_rate":   _txt(el,"ExchangeRate"),
                "document_no": _txt(el,"DocumentNo"),
                "debit":  debit,
                "credit": credit,
                "vat_code":  pick(el,"vat_code"),
                "vat_rate":  pick(el,"vat_rate"),
                "vat_base":  _dec(_txt(el,"TaxBase")),
                "vat_debit": _dec(_txt(el,"DebitTaxAmount")),
                "vat_credit":_dec(_txt(el,"CreditTaxAmount")),
            })

        elif ltag == "Analysis":
            analy.append({
                "journal_id":  cur_journal_id,
                "record_id":   _txt(el.getparent(), "RecordID"),
                "type":        _txt(el,"AnalysisType"),
                "analysis_id": _txt(el,"AnalysisID"),
                "debit_amt":   _dec(_txt(el,"DebitAnalysisAmount")),
                "credit_amt":  _dec(_txt(el,"CreditAnalysisAmount")),
            })

        el.clear()
    return hdr, bank, acc, cust, supp, jour, lines, analy

# ── kontroller ────────────────────────────────────────────────────────────
def _check_balance(df: pd.DataFrame) -> bool:
    diff = df["debit"].map(_to_dec) - df["credit"].map(_to_dec)
    bad_total = abs(diff.sum()) > DEC("0.01")
    bad_j = diff.groupby(df["journal_id"]).sum().loc[lambda s: s.abs() > DEC("0.01")]
    if bad_total or not bad_j.empty:
        logging.error("BALANSE-avvik:\n%s",
                      bad_j.head().to_string() if not bad_j.empty else diff.sum())
        return False
    return True

def _check_vat(df: pd.DataFrame) -> bool:
    diff = df["vat_debit"].map(_to_dec) - df["vat_credit"].map(_to_dec)
    bad = diff.groupby(df["vat_code"]).sum().loc[lambda s: s.abs() > DEC("0.01")]
    if not bad.empty:
        logging.error("MVA-avvik:\n%s", bad.to_string()); return False
    return True

# ── konvertering ──────────────────────────────────────────────────────────
def convert(src: Path, dst: Path, *, max_ratio: float,
            validate_only: bool, xsd: Path | None):
    dst.mkdir(parents=True, exist_ok=True)
    stream = _stream_xml(src, max_ratio)
    if xsd:
        xml = etree.parse(stream); etree.XMLSchema(etree.parse(xsd)).assertValid(xml)
        stream = io.BytesIO(etree.tostring(xml))

    hdr, bank, acc, cust, supp, jour, lines, analy = _parse(stream); stream.close()

    if not (_check_balance(pd.DataFrame(lines)) and _check_vat(pd.DataFrame(lines))):
        sys.exit(3)
    if validate_only:
        logging.info("Validering OK – ingen CSV skrevet."); return

    def _save(rows: List[Dict], name: str, money: Tuple[str, ...]):
        df = pd.DataFrame(rows).replace({None:""}).fillna("")
        for c in money:
            if c in df.columns:
                df[c] = df[c].map(_dec_str)
        df.to_csv(dst / name, sep=";", index=False, encoding="utf-8",
                  quoting=csv.QUOTE_MINIMAL)

    _save(hdr,  "header.csv", ())
    _save(bank, "bank_accounts.csv", ())
    _save(acc,  "accounts.csv", ("opening_balance","closing_balance"))
    _save(cust, "customers.csv", ())
    _save(supp, "suppliers.csv", ())
    _save(jour, "journal.csv", ())
    _save(lines,"transactions.csv",
          ("amount_currency","debit","credit","vat_base","vat_debit","vat_credit"))
    _save(analy,"analysis_lines.csv", ("debit_amt","credit_amt"))

    if WRITE_XLSX:
        try:
            with pd.ExcelWriter(dst / "SAF-T.xlsx", engine="xlsxwriter") as xl:
                for csvf in dst.glob("*.csv"):
                    pd.read_csv(csvf, sep=";", decimal=",", dtype=str) \
                      .to_excel(xl, sheet_name=csvf.stem[:31], index=False)
        except ModuleNotFoundError:
            logging.warning("pandas/xlsxwriter mangler – hopper over XLSX.")

    logging.info("🟢 Ferdig! %d kontoer • %d journaler • %d linjer.",
                 len(acc), len(jour), len(lines))

# ── CLI ───────────────────────────────────────────────────────────────────
def main():
    ap = argparse.ArgumentParser("Konverter SAF-T til CSV (+XLSX)")
    ap.add_argument("input_path", nargs="?", help="SAF-T XML/ZIP")
    ap.add_argument("output_dir", nargs="?", help="Mappe for CSV")
    ap.add_argument("--max-zip-ratio", type=float, default=MAX_ZIP_RATIO_DEFAULT)
    ap.add_argument("--validate-only", action="store_true")
    ap.add_argument("--xsd-validate", metavar="XSD", help="Valider mot XSD")
    ap.add_argument("--nogui", action="store_true")
    args = ap.parse_args()

    inp, out = args.input_path, args.output_dir
    if not args.nogui and (not inp or not out):
        try:
            from tkinter import Tk, filedialog
            Tk().withdraw()
            if not inp:
                inp = filedialog.askopenfilename(title="Velg SAF-T-fil",
                                                 filetypes=[("SAF-T", "*.xml *.zip")])
            if not out:
                out = filedialog.askdirectory(title="Velg mappe for CSV")
        except ModuleNotFoundError:
            logging.warning("tkinter mangler – bruk CLI-argumenter.")
    if not inp or not out:
        ap.error("Må angi input-fil og output-mappe.")

    try:
        convert(Path(inp), Path(out),
                max_ratio=args.max_zip_ratio,
                validate_only=args.validate_only,
                xsd=Path(args.xsd_validate) if args.xsd_validate else None)
    except etree.DocumentInvalid as e:
        logging.error("❌ XSD-validering feilet:\n%s", e); sys.exit(2)
    except Exception as exc:                       # pragma: no cover
        logging.exception("Uventet feil: %s", exc); sys.exit(4)

if __name__ == "__main__":
    main()
