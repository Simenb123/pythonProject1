from __future__ import annotations
from pathlib import Path
from typing import List, Dict, Any
from xml.etree import ElementTree as ET

NS = {"s": "urn:StandardAuditFile-Tax"}

def parse_saf_t(xml_source: Path | str) -> Dict[str, List[Dict[str, Any]]]:
    """Parse a Tripletex SAF-T XML file and return accounts, journals and transactions"""
    if isinstance(xml_source, (str, Path)):
        tree = ET.parse(str(xml_source))
    else:
        tree = ET.parse(xml_source)
    root = tree.getroot()

    accounts: List[Dict[str, Any]] = []
    for acc in root.findall(".//s:MasterFiles/s:GeneralLedgerAccounts/s:Account", namespaces=NS):
        accounts.append({
            "id": acc.findtext("s:AccountID", namespaces=NS),
            "description": acc.findtext("s:Description", namespaces=NS),
        })

    journals: List[Dict[str, Any]] = []
    transactions: List[Dict[str, Any]] = []

    for journal in root.findall(".//s:GeneralLedgerEntries/s:Journal", namespaces=NS):
        journal_id = journal.findtext("s:JournalID", namespaces=NS)
        journals.append({
            "id": journal_id,
            "description": journal.findtext("s:Description", namespaces=NS),
        })
        for trx in journal.findall("s:Transaction", namespaces=NS):
            transactions.append({
                "journal_id": journal_id,
                "id": trx.findtext("s:TransactionID", namespaces=NS),
                "description": trx.findtext("s:Description", namespaces=NS),
            })

    return {"accounts": accounts, "journals": journals, "transactions": transactions}

if __name__ == "__main__":
    import io
    sample_xml = """
    <AuditFile xmlns=\"urn:StandardAuditFile-Tax\">
      <MasterFiles>
        <GeneralLedgerAccounts>
          <Account>
            <AccountID>1000</AccountID>
            <Description>Cash</Description>
          </Account>
        </GeneralLedgerAccounts>
      </MasterFiles>
      <GeneralLedgerEntries>
        <Journal>
          <JournalID>1</JournalID>
          <Description>Sales</Description>
          <Transaction>
            <TransactionID>t1</TransactionID>
            <Description>sale 1</Description>
          </Transaction>
          <Transaction>
            <TransactionID>t2</TransactionID>
            <Description>sale 2</Description>
          </Transaction>
        </Journal>
      </GeneralLedgerEntries>
    </AuditFile>
    """
    with open("sample_saft.xml", "w") as f:
        f.write(sample_xml)
    data = parse_saf_t("sample_saft.xml")
    print(data)
