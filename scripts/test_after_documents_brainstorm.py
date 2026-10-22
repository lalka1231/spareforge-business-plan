"""Content regression checks; missing revised artifact uses baseline to expose defects."""
import re
import unittest
from pathlib import Path
ROOT=Path(__file__).resolve().parents[1]

class DocumentRevisionTests(unittest.TestCase):
    def test_no_unexpanded_financial_template(self):
        target=ROOT/'docs/FINAL_BUSINESS_PLAN_brainstorm.md'
        text=(target if target.exists() else ROOT/'docs/FINAL_BUSINESS_PLAN.md').read_text()
        self.assertNotIn('{fmt(',text)

    def test_method_monthly_fixed_cost(self):
        target=ROOT/'docs/FINANCE_METHOD_brainstorm.md'
        text=(target if target.exists() else ROOT/'docs/FINANCE_METHOD.md').read_text()
        section=text.split('## Месячная сезонность и cash runway')[1].split('## ')[0]
        self.assertNotIn('350 000',section)
        self.assertIn('490 600',section)

    def test_full_revised_sections(self):
        target=ROOT/'docs/FINAL_BUSINESS_PLAN_brainstorm.md'
        self.assertTrue(target.exists(),'Revised full report not built')
        text=target.read_text()
        self.assertEqual([int(i) for i in re.findall(r'^## (\d+)\.',text,re.M)],list(range(1,23)))
        self.assertIn('AgroVector_After_brainstorm',text)
        self.assertIn('безотзывную закупку',text)
        self.assertIn('21 760',text)
        self.assertNotIn('Не достигнута года',text)

if __name__=='__main__':unittest.main()
