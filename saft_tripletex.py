"""Convert Tripletex SAF-T exports to CSV files."""

from __future__ import annotations

from pathlib import Path
import argparse
import zipfile
from tkinter import Tk, filedialog

import pandas as pd

try:
    from lxml import etree
except ModuleNotFoundError as exc:  # pragma: no cover - helpful message
    raise ModuleNotFoundError(
        "lxml is required to parse SAF-T files. Install it with 'pip install lxml'."
    ) from exc

try:  # pragma: no cover - optional dependency
    from tqdm import tqdm
except ModuleNotFoundError:  # pragma: no cover - fallback
    def tqdm(iterable, **kwargs):
        for item in iterable:
            yield item


NS = {"s": "urn:StandardAuditFile-Tax"}

VAT_CODE_MAP: dict[int, str] = {
    0: "Ingen mva",
    1: "Høy sats",
    11: "Kjøp høy sats",
    12: "Kjøp middels sats",
    14: "Fradragsberettiget utenfor mva-området",
    15: "Fradragsført innland",
    21: "Middels sats",
    31: "Null sats",
}


def _strip_ns(tag: str) -> str:
    if "}" in tag:
        return tag.split("}", 1)[1]
    return tag


def _normalize_date(value: str | None) -> str | None:
    if not value:
        return None
    try:
        return pd.to_datetime(value).strftime("%Y-%m-%d")
    except Exception:
        return value


def _parse_accounts(elem: etree._Element) -> dict:
    return {
        "account_id": elem.findtext("s:AccountID", namespaces=NS),
        "description": elem.findtext("s:AccountDescription", namespaces=NS),
        "type": elem.findtext("s:AccountType", namespaces=NS),
        "vat_code": elem.findtext("s:VatCode", namespaces=NS),
    }


def _parse_journal(elem: etree._Element) -> tuple[dict, list[dict]]:
    j = {
        "journal_id": elem.findtext("s:JournalID", namespaces=NS),
        "description": elem.findtext("s:Description", namespaces=NS),
        "voucher_date": _normalize_date(elem.findtext("s:VoucherDate", namespaces=NS)),
        "voucher_type": elem.findtext("s:VoucherType", namespaces=NS),
    }

    transactions: list[dict] = []
    for trans in elem.findall("s:Transaction", namespaces=NS):
        t = {
            "journal_id": j["journal_id"],
            "voucher_no": trans.findtext("s:VoucherNo", namespaces=NS)
            or trans.findtext("s:VoucherID", namespaces=NS)
            or j["journal_id"],
            "account_id": trans.findtext("s:AccountID", namespaces=NS),
            "debit": trans.findtext("s:DebitAmount", namespaces=NS),
            "credit": trans.findtext("s:CreditAmount", namespaces=NS),
            "currency": trans.findtext("s:Currency/s:Code", namespaces=NS),
            "document_no": trans.findtext("s:DocumentNumber", namespaces=NS),
            "vat_code": trans.findtext("s:TaxInformation/s:VATCode", namespaces=NS),
        }
        transactions.append(t)
    return j, transactions


def _read_saft_xml(file_obj) -> tuple[list[dict], list[dict], list[dict]]:
    accounts: list[dict] = []
    journals: list[dict] = []
    transactions: list[dict] = []

    ns_uri = None
    version = None

    context = etree.iterparse(file_obj, events=("end",), recover=True)
    for event, elem in tqdm(context, desc="Parsing", unit="elem"):
        tag = _strip_ns(elem.tag)
        if tag == "Account":
            accounts.append(_parse_accounts(elem))
        elif tag == "Journal":
            journal, trans = _parse_journal(elem)
            journals.append(journal)
            transactions.extend(trans)
        elif tag == "AuditFile":
            if "}" in elem.tag:
                ns_uri = elem.tag.split("}", 1)[0][1:]
            version = elem.get("version")
        elem.clear()
        while elem.getprevious() is not None:
            del elem.getparent()[0]
    del context

    if ns_uri != NS["s"] or version not in {"1.2", "1.3"}:
        raise ValueError(f"Unsupported SAF-T version: {version}")

    return accounts, journals, transactions


def konverter_saft_tripletex(
    input_path: str | Path,
    output_dir: str | Path,
) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Convert a Tripletex SAF-T file to three CSV files."""

    in_path = Path(input_path)
    out_dir = Path(output_dir)
    if not in_path.exists():
        raise FileNotFoundError(f"Input file not found: {in_path}")
    out_dir.mkdir(parents=True, exist_ok=True)

    if zipfile.is_zipfile(in_path):
        with zipfile.ZipFile(in_path) as zf:
            xml_names = [n for n in zf.namelist() if n.lower().endswith(".xml")]
            if not xml_names:
                raise ValueError("Zip file contains no XML file")
            with zf.open(xml_names[0]) as f:
                accounts, journals, transactions = _read_saft_xml(f)
    else:
        with open(in_path, "rb") as f:
            accounts, journals, transactions = _read_saft_xml(f)

    accounts_df = pd.DataFrame(accounts)
    journals_df = pd.DataFrame(journals)
    transactions_df = pd.DataFrame(transactions)

    for df in [accounts_df, journals_df, transactions_df]:
        if "vat_code" in df.columns:
            df["vat_code"] = pd.to_numeric(df["vat_code"], errors="coerce").astype("Int64")
            df["vat_description"] = df["vat_code"].map(VAT_CODE_MAP)

    if "voucher_date" in journals_df.columns:
        journals_df["voucher_date"] = journals_df["voucher_date"].apply(_normalize_date)

    accounts_df.to_csv(out_dir / "accounts.csv", sep=";", index=False, encoding="utf-8")
    journals_df.to_csv(out_dir / "journal.csv", sep=";", index=False, encoding="utf-8")
    transactions_df.to_csv(out_dir / "transactions.csv", sep=";", index=False, encoding="utf-8")

    return accounts_df, journals_df, transactions_df


def _cli() -> None:
    parser = argparse.ArgumentParser(
        description=(
            "Convert Tripletex SAF-T file (XML or ZIP) to CSV files. "
            "If arguments are omitted, dialog boxes will open."
        )
    )
    parser.add_argument("input_path", nargs="?", help="Path to SAF-T XML or ZIP")
    parser.add_argument("output_dir", nargs="?", help="Directory for CSV output")
    args = parser.parse_args()

    input_path = args.input_path
    output_dir = args.output_dir
    if not input_path or not output_dir:
        Tk().withdraw()
        if not input_path:
            input_path = filedialog.askopenfilename(
                title="Velg SAF-T-fil (XML eller ZIP)",
                filetypes=[("SAF-T filer", "*.xml *.zip"), ("Alle filer", "*.*")],
            )
        if not output_dir:
            output_dir = filedialog.askdirectory(title="Velg mappe for CSV-utdata")

    if not input_path or not output_dir:
        parser.error("input_path and output_dir are required")

    print(f"Input: {input_path}  Output: {output_dir}")
    konverter_saft_tripletex(input_path, output_dir)


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    _cli()

