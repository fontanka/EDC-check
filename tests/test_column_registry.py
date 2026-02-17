"""
Unit tests for column_registry.py
Tests visit constants, column constants, get_col helper, and validate_columns.
"""

import unittest
from column_registry import VISITS, COL, get_col, validate_columns, CRITICAL_COLUMNS


class TestVISITS(unittest.TestCase):
    """Test visit prefix constants."""

    def test_onsite_visits(self):
        self.assertIn('SBV', VISITS.ONSITE)
        self.assertIn('TV', VISITS.ONSITE)
        self.assertIn('FU1M', VISITS.ONSITE)
        self.assertIn('FU6M', VISITS.ONSITE)

    def test_remote_visits(self):
        self.assertEqual(VISITS.REMOTE, ['FU3M', 'FU3Y', 'FU5Y'])

    def test_no_overlap(self):
        """Onsite and remote should not overlap."""
        overlap = set(VISITS.ONSITE) & set(VISITS.REMOTE)
        self.assertEqual(overlap, set())

    def test_all_visits_comprehensive(self):
        """ALL should include both onsite and remote (except UV and LOGS)."""
        for v in VISITS.ONSITE:
            if v not in ('UV',):
                self.assertIn(v, VISITS.ALL, f"{v} should be in ALL")
        for v in VISITS.REMOTE:
            self.assertIn(v, VISITS.ALL, f"{v} should be in ALL")

    def test_all_visits_ordered(self):
        """ALL visits should be in chronological order."""
        expected_order = ['SBV', 'TV', 'DV', 'FU1M', 'FU3M', 'FU6M',
                         'FU1Y', 'FU2Y', 'FU3Y', 'FU4Y', 'FU5Y']
        self.assertEqual(VISITS.ALL, expected_order)


class TestCOL(unittest.TestCase):
    """Test column name constants."""

    def test_identifiers_no_prefix(self):
        self.assertEqual(COL.SCREENING_NUM, 'Screening #')
        self.assertEqual(COL.SITE_NUM, 'Site #')
        self.assertEqual(COL.STATUS, 'Status')

    def test_procedure_date(self):
        self.assertEqual(COL.PROCEDURE_DATE, 'TV_PR_PRSTDTC')

    def test_eligibility(self):
        self.assertEqual(COL.ELIGIBILITY_DECISION, 'SBV_ELIG_IEORRES_CONF5')

    def test_ae_fields(self):
        self.assertEqual(COL.AE_TERM, 'LOGS_AE_AETERM')
        self.assertEqual(COL.AE_START_DATE, 'LOGS_AE_AESTDTC')
        self.assertEqual(COL.AE_SERIOUS, 'LOGS_AE_AESER')

    def test_death_fields(self):
        self.assertEqual(COL.DEATH_DATE, 'LOGS_DTH_DDDTC')
        self.assertEqual(COL.DEATH_CATEGORY, 'LOGS_DTH_DDRESCAT')
        self.assertEqual(COL.DEATH_REASON, 'LOGS_DTH_DDORRES')

    def test_kccq_fields(self):
        self.assertEqual(COL.KCCQ_OVERALL, 'KCCQ_QSORRES_KCCQ_OVERALL')
        self.assertEqual(COL.KCCQ_CLINICAL, 'KCCQ_QSORRES_KCCQ_CLINICAL')

    def test_vs_ascites_typo_preserved(self):
        """CRF has 'ASCITIS' typo — must be preserved in registry."""
        self.assertEqual(COL.VS_ASCITES, 'VS_CVORRES_ASCITIS')


class TestGetCol(unittest.TestCase):
    """Test visit-prefixed column name builder."""

    def test_basic(self):
        self.assertEqual(get_col('SBV', 'LB_CBC_LBORRES_HGB'), 'SBV_LB_CBC_LBORRES_HGB')

    def test_treatment_visit(self):
        self.assertEqual(get_col('TV', 'PR_PRSTDTC'), 'TV_PR_PRSTDTC')

    def test_follow_up(self):
        self.assertEqual(get_col('FU1M', 'VS_VSORRES_HR'), 'FU1M_VS_VSORRES_HR')

    def test_with_visit_constant(self):
        self.assertEqual(get_col(VISITS.SCREENING, COL.KCCQ_OVERALL), 'SBV_KCCQ_QSORRES_KCCQ_OVERALL')


class TestValidateColumns(unittest.TestCase):
    """Test column validation function."""

    def test_all_present(self):
        cols = ['A', 'B', 'C']
        found, missing = validate_columns(cols, ['A', 'B'])
        self.assertEqual(found, ['A', 'B'])
        self.assertEqual(missing, [])

    def test_some_missing(self):
        cols = ['A', 'B']
        found, missing = validate_columns(cols, ['A', 'C', 'D'])
        self.assertEqual(found, ['A'])
        self.assertEqual(set(missing), {'C', 'D'})

    def test_empty_required(self):
        found, missing = validate_columns(['A', 'B'], [])
        self.assertEqual(found, [])
        self.assertEqual(missing, [])

    def test_empty_columns(self):
        found, missing = validate_columns([], ['A', 'B'])
        self.assertEqual(found, [])
        self.assertEqual(set(missing), {'A', 'B'})


class TestCriticalColumns(unittest.TestCase):
    """Test that CRITICAL_COLUMNS list is well-formed."""

    def test_not_empty(self):
        self.assertGreater(len(CRITICAL_COLUMNS), 0)

    def test_contains_screening_num(self):
        self.assertIn('Screening #', CRITICAL_COLUMNS)

    def test_contains_procedure_date(self):
        self.assertIn('TV_PR_PRSTDTC', CRITICAL_COLUMNS)

    def test_all_strings(self):
        for col in CRITICAL_COLUMNS:
            self.assertIsInstance(col, str)


if __name__ == '__main__':
    unittest.main()
