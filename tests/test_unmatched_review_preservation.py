from datetime import date
from decimal import Decimal
import hashlib
from personal_budget_plaid_excel.domain.source_mode import ReconciliationItem, ReconciliationReport, ReconciliationClassification
from pathlib import Path
import tempfile, unittest
from openpyxl import load_workbook
from openpyxl.utils.cell import range_boundaries
from personal_budget_plaid_excel.domain.models import CategoryBudget, LedgerTransaction
from personal_budget_plaid_excel.domain.review import build_review_transactions
from personal_budget_plaid_excel.infrastructure.review_workbook import OpenpyxlReviewWorkbookWriter, PRESERVED_REVIEW_TABLE

class UnmatchedReviewPreservationTests(unittest.TestCase):
    def tx(self, merchant, amount='20.00'):
        return LedgerTransaction(transaction_date=date(2026,8,5),month=date(2026,8,1),merchant=merchant,account='Synthetic Account',amount=Decimal(amount),personal_budget=True,source='synthetic')
    def write(self, path, txs, reconciliation=None):
        return OpenpyxlReviewWorkbookWriter().write(path,date(2026,8,1),[CategoryBudget('Flexible','Food',Decimal('100')),CategoryBudget('Flexible','Other',Decimal('50'))],build_review_transactions(txs,[]), reconciliation=reconciliation)
    def test_unmatched_choice_evidence_and_disposition_survive_reruns_without_active_duplication(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'review.xlsx';old=self.tx('Original Merchant');current=self.tx('Unrelated New Merchant','30.00')
            self.write(p,[old]);w=load_workbook(p);w['Transactions']['E4']='Food';w['Transactions']['F4']='=HYPERLINK("https://mail.google.com/mail/u/0/#all/synthetic-thread","Verified item")';original_key=w['Transactions']['I4'].value;w.save(p)
            self.write(p,[current]);w=load_workbook(p);q=w['Reconciliation'];first=range_boundaries(q.tables[PRESERVED_REVIEW_TABLE].ref)[1];row=first+1
            self.assertEqual(w['Transactions'].max_row,4);self.assertIsNone(w['Transactions']['E4'].value);self.assertEqual(q.cell(row,5).value,'Food');self.assertEqual(q.cell(row,6).value,'=HYPERLINK("https://mail.google.com/mail/u/0/#all/synthetic-thread","Verified item")')
            self.assertEqual(w['_ReviewState']['D2'].value,original_key);self.assertIn('Transactions',w['Budget Summary']['C4'].value);self.assertNotIn('Reconciliation',w['Budget Summary']['C4'].value)
            q.cell(row,9,'Hold — investigate identity');q.cell(row,10,'Human note retained');w['Transactions']['E4']='Other';w.save(p)
            self.write(p,[current]);w=load_workbook(p);q=w['Reconciliation'];self.assertEqual(q.cell(row,5).value,'Food');self.assertEqual(q.cell(row,9).value,'Hold — investigate identity');self.assertEqual(q.cell(row,10).value,'Human note retained');self.assertEqual(w['Transactions']['E4'].value,'Other');self.assertEqual(range_boundaries(q.tables[PRESERVED_REVIEW_TABLE].ref)[3],row)
            self.assertEqual(len(w['Transactions'].data_validations.dataValidation),1);self.assertIn('E4',w['Transactions'].data_validations.dataValidation[0].sqref)
    def test_corrupt_preserved_identity_fails_before_replacing_output(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'review.xlsx';self.write(p,[self.tx('Old')]);w=load_workbook(p);w['Transactions']['E4']='Food';w.save(p);self.write(p,[self.tx('Different','40')]);w=load_workbook(p);w['_ReviewState']['D2']='bad-key';w.save(p);before=p.read_bytes()
            with self.assertRaisesRegex(ValueError,'preserved review identity'):self.write(p,[self.tx('Different','40')])
            self.assertEqual(p.read_bytes(),before)
    def test_confirmed_exact_identity_can_reuse_preserved_choice_without_duplicate_archive(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'review.xlsx';old=self.tx('Old');new=self.tx('New','30');self.write(p,[old]);w=load_workbook(p);w['Transactions']['E4']='Food';w.save(p);self.write(p,[new]);self.write(p,[new,old]);w=load_workbook(p)
            self.assertEqual(w['Transactions'].max_row,5);self.assertEqual(w['Transactions']['E5'].value,'Food');self.assertNotIn(PRESERVED_REVIEW_TABLE,w['Reconciliation'].tables)
    def test_growing_source_queue_does_not_overwrite_preserved_classifications(self):
        with tempfile.TemporaryDirectory() as tmp:
            p=Path(tmp)/'review.xlsx';old=self.tx('Old');new=self.tx('New','30');self.write(p,[old]);w=load_workbook(p);w['Transactions']['E4']='Food';w.save(p);self.write(p,[new])
            items=tuple(ReconciliationItem(hashlib.sha256(str(i).encode()).hexdigest(),'Chase',date(2026,8,i+1),Decimal(i+1),ReconciliationClassification.AMBIGUOUS,'Review needed','No automatic budget change','Synthetic queue note') for i in range(7))
            self.write(p,[new],ReconciliationReport(coverage=(),items=items));w=load_workbook(p);q=w['Reconciliation'];first=range_boundaries(q.tables[PRESERVED_REVIEW_TABLE].ref)[1]
            self.assertGreater(first,22);self.assertEqual(q.cell(first+1,5).value,'Food');self.assertEqual(q['A22'].value,'Chase');self.write(p,[new]);w=load_workbook(p);self.assertEqual(w['Reconciliation'].cell(first+1,5).value,'Food')
if __name__=='__main__':unittest.main()
