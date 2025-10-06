
import unittest, tempfile, csv
from pathlib import Path
from decimal import Decimal as DEC
from saft_parser_pro import parse_and_write, MAX_ZIP_RATIO_DEFAULT

NESTED_XML = """<?xml version="1.0" encoding="UTF-8"?>
<AuditFile xmlns="urn:StandardAuditFile-Tax">
  <Header>
    <AuditFileVersion>1.20</AuditFileVersion>
    <SoftwareCompanyName>Test</SoftwareCompanyName>
    <SoftwareVersion>1.0</SoftwareVersion>
    <AuditFileDateCreated>2025-01-01</AuditFileDateCreated>
    <SelectionStartDate>2025-01-01</SelectionStartDate>
    <SelectionEndDate>2025-01-31</SelectionEndDate>
    <Company>
      <CompanyName>ACME AS</CompanyName>
      <CompanyID>999999999</CompanyID>
    </Company>
  </Header>
  <MasterFiles>
    <GeneralLedgerAccounts>
      <Account>
        <AccountID>3000</AccountID>
        <AccountDescription>Salg</AccountDescription>
        <AccountType>GL</AccountType>
      </Account>
      <Account>
        <AccountID>4000</AccountID>
        <AccountDescription>Kjøp</AccountDescription>
        <AccountType>GL</AccountType>
      </Account>
    </GeneralLedgerAccounts>
    <Customer>
      <CustomerID>C1</CustomerID>
      <Name>Kunde</Name>
      <PaymentTermsDays>14</PaymentTermsDays>
    </Customer>
    <Supplier>
      <SupplierID>S1</SupplierID>
      <Name>Lev</Name>
      <PaymentTermsDays>30</PaymentTermsDays>
    </Supplier>
    <TaxTable>
      <TaxTableEntry>
        <TaxCode>HIGH</TaxCode>
        <TaxPercentage>25</TaxPercentage>
        <Description>MVA 25%</Description>
      </TaxTableEntry>
      <TaxTableEntry>
        <TaxCode>HIGH</TaxCode>
        <TaxPercentage>25</TaxPercentage>
      </TaxTableEntry>
    </TaxTable>
  </MasterFiles>
  <GeneralLedgerEntries>
    <Journal>
      <JournalID>J1</JournalID>
      <Description>Hovedbok</Description>
      <Transaction>
        <TransactionDate>2025-01-10</TransactionDate>
        <VoucherNo>V1</VoucherNo>
        <Line>
          <RecordID>1</RecordID>
          <AccountID>3000</AccountID>
          <CustomerID>C1</CustomerID>
          <PostingDate>2025-01-10</PostingDate>
          <DebitAmount><Amount>100.00</Amount></DebitAmount>
          <TaxAmount><Amount>25.00</Amount></TaxAmount>
        </Line>
        <Line>
          <RecordID>2</RecordID>
          <AccountID>4000</AccountID>
          <SupplierID>S1</SupplierID>
          <PostingDate>2025-01-10</PostingDate>
          <CreditAmount><Amount>100.00</Amount></CreditAmount>
        </Line>
      </Transaction>
    </Journal>
    <Journal>
      <JournalID>J2</JournalID>
      <Transaction>
        <VoucherNo>V2</VoucherNo>
        <Line>
          <RecordID>1</RecordID>
          <AccountID>9999</AccountID>
          <PostingDate>2025-01-15</PostingDate>
          <DebitAmount><Amount>10.00</Amount></DebitAmount>
        </Line>
      </Transaction>
    </Journal>
  </GeneralLedgerEntries>
</AuditFile>
"""

class TestProParser(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.TemporaryDirectory()
        self.tmpdir = Path(self.tmp.name)
        self.xml = self.tmpdir / "nested.xml"
        self.xml.write_text(NESTED_XML, encoding="utf-8")
        self.out = self.tmpdir / "out"

    def tearDown(self):
        self.tmp.cleanup()

    def test_nested_amounts_and_outputs(self):
        parse_and_write(self.xml, self.out, max_ratio=MAX_ZIP_RATIO_DEFAULT, stream=True, write_xlsx=False, strict=False, xsd=None)
        tx = list(csv.DictReader(open(self.out/'transactions.csv', encoding='utf-8'), delimiter=';'))
        # Tre linjer totalt
        self.assertEqual(len(tx), 3)  # 2 i V1 + 1 i V2
        # Første linje: debit 100
        self.assertEqual(tx[0]['debit'], '100.00')
        # Andre linje: credit 100
        self.assertEqual(tx[1]['credit'], '100.00')
        # TaxAmount fanget
        self.assertEqual(tx[0]['vat_amount'], '25.00')

        # Vouchers csv
        vouchers = list(csv.DictReader(open(self.out/'vouchers.csv', encoding='utf-8'), delimiter=';'))
        v1 = [v for v in vouchers if v['voucher_no'] == 'V1'][0]
        v2 = [v for v in vouchers if v['voucher_no'] == 'V2'][0]
        self.assertEqual(DEC(v1['imbalance']), DEC('0'))
        self.assertNotEqual(DEC(v2['imbalance']), DEC('0'))

        # Referanse-integritet: 9999 mangler i accounts
        miss = list(csv.DictReader(open(self.out/'missing_accounts.csv', encoding='utf-8'), delimiter=';'))
        self.assertIn('9999', [m['account_id'] for m in miss])

        # AR/AP aggregater finnes
        self.assertTrue((self.out/'ar_aggregates.csv').exists())
        self.assertTrue((self.out/'ap_aggregates.csv').exists())

if __name__ == '__main__':
    unittest.main(verbosity=2)
