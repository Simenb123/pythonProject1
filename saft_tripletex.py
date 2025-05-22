from __future__ import annotations
import xml.etree.ElementTree as ET
from pathlib import Path

# Namespace used by Tripletex SAF-T exports
NS = {"s": "urn:StandardAuditFile-Tax"}


def parse(path: str | Path) -> ET.Element:
    """Parse SAF-T XML file and return root element."""
    return ET.parse(path).getroot()


def get_accounts(root: ET.Element) -> list[dict[str, str]]:
    """Return a list of accounts found in the file."""
    accounts = []
    for acct in root.findall(".//s:GeneralLedgerAccounts/s:Account", namespaces=NS):
        accounts.append({
            "id": acct.findtext("s:AccountID", default="", namespaces=NS),
            "description": acct.findtext("s:AccountDescription", default="", namespaces=NS),
            "type": acct.findtext("s:AccountType", default="", namespaces=NS),
        })
    return accounts


def get_journals(root: ET.Element) -> list[dict[str, object]]:
    """Return a list of journals and their transactions."""
    journals = []
    for journal in root.findall(".//s:GeneralLedgerEntries/s:Journal", namespaces=NS):
        transactions = []
        for tx in journal.findall("s:Transaction", namespaces=NS):
            transactions.append({
                "id": tx.findtext("s:TransactionID", default="", namespaces=NS),
                "description": tx.findtext("s:Description", default="", namespaces=NS),
                "period": tx.findtext("s:Period", default="", namespaces=NS),
            })
        journals.append({
            "id": journal.findtext("s:JournalID", default="", namespaces=NS),
            "description": journal.findtext("s:Description", default="", namespaces=NS),
            "transactions": transactions,
        })
    return journals


def get_transactions(root: ET.Element) -> list[dict[str, str]]:
    """Return all transactions from all journals."""
    txs = []
    for tx in root.findall(".//s:GeneralLedgerEntries/s:Journal/s:Transaction", namespaces=NS):
        txs.append({
            "id": tx.findtext("s:TransactionID", default="", namespaces=NS),
            "description": tx.findtext("s:Description", default="", namespaces=NS),
            "period": tx.findtext("s:Period", default="", namespaces=NS),
        })
    return txs


if __name__ == "__main__":
    import argparse
    import pprint

    parser = argparse.ArgumentParser(description="Parse Tripletex SAF-T file")
    parser.add_argument("xml_file", type=Path)
    args = parser.parse_args()

    root = parse(args.xml_file)
    accounts = get_accounts(root)
    journals = get_journals(root)
    transactions = get_transactions(root)

    print(f"Accounts: {len(accounts)}")
    print(f"Journals: {len(journals)}")
    print(f"Transactions: {len(transactions)}")
    pprint.pprint(accounts[:3])
    if journals:
        pprint.pprint(journals[0])
