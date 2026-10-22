#!/usr/bin/env python3
"""Independent invariants for the parallel corrected finance export."""
import json
import math
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / 'scripts'))
def corrected_data():
    import build_corrected_finance_brainstorm as corrected
    return corrected.build(write=False)


class CorrectedFinanceTests(unittest.TestCase):
    def test_runway_corrects_legacy_sixth_month_to_fifth(self):
        source = json.loads((ROOT / 'data/finance_summary.json').read_text(encoding='utf-8'))
        reserve = source['cash_runway']['initial_operating_cash_rub']
        fixed = source['yearly_scenarios']['base'][1]['fixed_opex_rub']
        burn = fixed / 12
        expected = math.floor(reserve / burn) + 1
        self.assertEqual(expected, 5)
        corrected = corrected_data()
        self.assertEqual(corrected['corrections']['legacy_no_revenue_first_negative_month'],
                         source['cash_runway']['no_revenue_first_negative_month'])
        self.assertEqual(corrected['cash_runway']['no_revenue_first_negative_month'], expected)

    def setUp(self):
        if self._testMethodName == 'test_runway_corrects_legacy_sixth_month_to_fifth':
            return
        self.data = corrected_data()
        self.rows = self.data['yearly_scenarios']['base']
        self.monthly = self.data['first_year_monthly']

    def test_fcf_identity_all_years_including_t0(self):
        for row in self.rows:
            expected = (row['ebitda_rub'] - row['usn_tax_rub']
                        - row['additional_capex_rub'] + row['reserve_return_rub'])
            if row['period'] == 0:
                expected = -(row['initial_capex_rub'] + row['launch_cost_rub'] + row['reserve_funding_rub'])
            self.assertAlmostEqual(row['free_cash_flow_rub'], expected, places=6)

    def test_monthly_rollforward_and_matches_existing_seasonality(self):
        for i, row in enumerate(self.monthly):
            self.assertAlmostEqual(row['closing_cash_rub'], row['opening_cash_rub'] + row['net_operating_cf_rub'], places=6)
            if i:
                self.assertAlmostEqual(row['opening_cash_rub'], self.monthly[i-1]['closing_cash_rub'], places=6)
        self.assertAlmostEqual(sum(r['net_operating_cf_rub'] for r in self.monthly),
                               self.rows[1]['ebitda_rub'] - self.rows[1]['usn_tax_rub'], places=6)

    def test_no_revenue_monthly_runway_identity(self):
        reserve = self.data['cash_runway']['initial_operating_cash_rub']
        fixed_monthly = self.rows[1]['fixed_opex_rub'] / 12
        for i, row in enumerate(self.data['no_revenue_cash_monthly'], 1):
            self.assertAlmostEqual(row['closing_cash_rub'], reserve - i*fixed_monthly, places=6)
        first = self.data['cash_runway']['no_revenue_first_negative_month']
        self.assertLess(self.data['no_revenue_cash_monthly'][first-1]['closing_cash_rub'], 0)
        self.assertGreaterEqual(self.data['no_revenue_cash_monthly'][first-2]['closing_cash_rub'], 0)

    def test_base_economics_unchanged_from_canonical_source(self):
        with (ROOT / 'data/finance_summary.json').open(encoding='utf-8') as stream:
            canonical = json.load(stream)
        self.assertEqual(self.data['base'], canonical['base'])
        self.assertEqual(self.rows, canonical['yearly_scenarios']['base'])


if __name__ == '__main__':
    unittest.main(verbosity=2)
