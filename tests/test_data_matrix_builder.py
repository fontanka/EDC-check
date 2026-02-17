"""Tests for data_matrix_builder — column classification, helpers."""
import sys
import unittest
from datetime import datetime
from unittest.mock import MagicMock

# Mock tkinter before importing the module (not available in headless CI)
sys.modules.setdefault('tkinter', MagicMock())
sys.modules.setdefault('tkinter.ttk', MagicMock())
sys.modules.setdefault('tkinter.messagebox', MagicMock())
sys.modules.setdefault('tkinter.filedialog', MagicMock())

from data_matrix_builder import (
    classify_column, parse_time_minutes, try_parse_date,
)


class TestClassifyColumn(unittest.TestCase):
    """Column type classification."""

    def test_ae_term(self):
        self.assertEqual(classify_column("LOGS_AE_AETERM"), 'ae')

    def test_ae_severity(self):
        self.assertEqual(classify_column("LOGS_AE_AESEV"), 'ae')

    def test_cm_treatment(self):
        self.assertEqual(classify_column("LOGS_CM_CMTRT"), 'cm')

    def test_cm_dose(self):
        self.assertEqual(classify_column("LOGS_CM_CMDOSE"), 'cm')

    def test_mh_term(self):
        self.assertEqual(classify_column("SBV_MH_MHTERM"), 'mh')

    def test_mh_non_mh_column(self):
        """MHTERM without _MH_ prefix should not match."""
        self.assertIsNone(classify_column("MHTERM_ONLY"))

    def test_hfh_date(self):
        self.assertEqual(classify_column("SBV_HFH_HOSTDTC"), 'hfh')

    def test_hmeh_column(self):
        self.assertEqual(classify_column("SBV_HMEH_HOSTDTC"), 'hmeh')

    def test_hmeh_keyword(self):
        self.assertEqual(classify_column("HMEH_OTHER_FIELD"), 'hmeh')

    def test_cvc_column(self):
        self.assertEqual(classify_column("SBV_CVC_CVORRES_CVPM"), 'cvc')

    def test_cvc_without_cvc_prefix(self):
        """CVC fields without _CVC_ prefix should not match."""
        self.assertIsNone(classify_column("CVORRES_CVPM"))

    def test_cvh_column(self):
        self.assertEqual(classify_column("SBV_CVH_PRSTDTC"), 'cvh')

    def test_act_column(self):
        self.assertEqual(classify_column("TV_LB_ACT_LBORRES_ACT"), 'act')

    def test_ae_ref_lbref(self):
        self.assertEqual(classify_column("LOGS_AE_LBREF"), 'ae_ref')

    def test_ae_ref_prref(self):
        self.assertEqual(classify_column("LOGS_AE_PRREF"), 'ae_ref')

    def test_generic_lab(self):
        """Generic lab column should return None (handled by matrix support check)."""
        self.assertIsNone(classify_column("SBV_LB_CBC_LBORRES_HGB"))

    def test_visit_date(self):
        """Visit date column should return None."""
        self.assertIsNone(classify_column("SBV_SV_SVSTDTC"))

    def test_empty(self):
        self.assertIsNone(classify_column(""))


class TestParseTimeMinutes(unittest.TestCase):
    def test_valid_time(self):
        self.assertEqual(parse_time_minutes("14:30"), 14 * 60 + 30)

    def test_midnight(self):
        self.assertEqual(parse_time_minutes("00:00"), 0)

    def test_invalid(self):
        self.assertEqual(parse_time_minutes("not_a_time"), 9999)

    def test_empty(self):
        self.assertEqual(parse_time_minutes(""), 9999)

    def test_single_digit(self):
        self.assertEqual(parse_time_minutes("9:05"), 9 * 60 + 5)


class TestTryParseDate(unittest.TestCase):
    def test_valid_iso(self):
        result = try_parse_date("2025-03-15 10:00:00")
        self.assertIsInstance(result, datetime)

    def test_short_date(self):
        """Short dates (<=10 chars) should return datetime.max."""
        result = try_parse_date("2025-03-15")
        self.assertEqual(result, datetime.max)

    def test_invalid(self):
        result = try_parse_date("not_a_date")
        self.assertEqual(result, datetime.max)

    def test_with_suffix(self):
        """Numeric suffix like ' (2)' should be stripped."""
        result = try_parse_date("2025-03-15 10:00:00 (2)")
        self.assertIsInstance(result, datetime)

    def test_visit_label(self):
        """Visit labels like 'Screening (2025-01-01)' return datetime.max."""
        result = try_parse_date("Screening (2025-01-01)")
        self.assertEqual(result, datetime.max)


if __name__ == '__main__':
    unittest.main()
