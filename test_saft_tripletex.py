import unittest
from saft_tripletex import get_accounts, get_journals, get_transactions, NS
import xml.etree.ElementTree as ET

SAMPLE_XML = f"""
<AuditFile xmlns='{NS['s']}'>
    <GeneralLedgerAccounts>
        <Account>
            <AccountID>1000</AccountID>
            <AccountDescription>Bank</AccountDescription>
            <AccountType>BS</AccountType>
        </Account>
    </GeneralLedgerAccounts>
    <GeneralLedgerEntries>
        <Journal>
            <JournalID>1</JournalID>
            <Description>Journal</Description>
            <Transaction>
                <TransactionID>10</TransactionID>
                <Description>Desc</Description>
                <Period>2024-01</Period>
            </Transaction>
        </Journal>
    </GeneralLedgerEntries>
</AuditFile>
"""

class TestSaftTripletex(unittest.TestCase):
    def setUp(self):
        self.root = ET.fromstring(SAMPLE_XML)

    def test_accounts(self):
        accounts = get_accounts(self.root)
        self.assertEqual(accounts[0]["id"], "1000")
        self.assertEqual(accounts[0]["description"], "Bank")

    def test_journals_and_transactions(self):
        journals = get_journals(self.root)
        self.assertEqual(journals[0]["id"], "1")
        self.assertEqual(journals[0]["transactions"][0]["id"], "10")
        txs = get_transactions(self.root)
        self.assertEqual(txs[0]["description"], "Desc")

if __name__ == '__main__':
    unittest.main()
