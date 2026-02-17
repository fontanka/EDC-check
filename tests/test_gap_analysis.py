"""
Unit tests for gap_analysis.py and ViewBuilder gap-related logic.
Tests gap detection, column-to-visit mapping, and gap count indexing.
"""

import sys
import types
import unittest

# Mock tkinter before importing view_builder (tkinter not available in CI)
_tk_mock = types.ModuleType("tkinter")
_tk_mock.Frame = type("Frame", (), {})
_tk_mock.Label = type("Label", (), {})
_tk_mock.Canvas = type("Canvas", (), {})
_tk_mock.Toplevel = type("Toplevel", (), {})
_tk_mock.Tk = type("Tk", (), {})
_tk_mock.BooleanVar = type("BooleanVar", (), {"__init__": lambda self, **kw: None, "get": lambda self: False})
_tk_mock.StringVar = type("StringVar", (), {"__init__": lambda self, **kw: None, "get": lambda self: ""})
_tk_mock.END = "end"
_tk_mock.X = "x"
_tk_mock.Y = "y"
_tk_mock.BOTH = "both"
_tk_mock.LEFT = "left"
_tk_mock.RIGHT = "right"
_tk_mock.TOP = "top"
_tk_mock.BOTTOM = "bottom"

_ttk_mock = types.ModuleType("tkinter.ttk")
_ttk_mock.Treeview = type("Treeview", (), {})
_ttk_mock.Scrollbar = type("Scrollbar", (), {})
_ttk_mock.Notebook = type("Notebook", (), {})
_ttk_mock.Combobox = type("Combobox", (), {})
_ttk_mock.Progressbar = type("Progressbar", (), {})
_ttk_mock.Style = type("Style", (), {"__init__": lambda self, *a, **kw: None, "configure": lambda self, *a, **kw: None})

_msg_mock = types.ModuleType("tkinter.messagebox")
_msg_mock.showwarning = lambda *a, **kw: None
_msg_mock.showinfo = lambda *a, **kw: None
_msg_mock.showerror = lambda *a, **kw: None

_fd_mock = types.ModuleType("tkinter.filedialog")
_fd_mock.asksaveasfilename = lambda **kw: None

sys.modules.setdefault("tkinter", _tk_mock)
sys.modules.setdefault("tkinter.ttk", _ttk_mock)
sys.modules.setdefault("tkinter.messagebox", _msg_mock)
sys.modules.setdefault("tkinter.filedialog", _fd_mock)

from view_builder import ViewBuilder
from config import VISIT_MAP


class MockApp:
    """Minimal mock for ClinicalDataMasterV30."""
    def __init__(self):
        self.df_main = None
        self.df_ae = None
        self.labels = {}
        self.ae_lookup = {}
        self.current_patient_gaps = []
        self.current_tree_data = {}
        self.sdv_manager = None
        self.sdv_verified_fields = set()


class TestIdentifyColumn(unittest.TestCase):
    """Test _identify_column visit/form detection."""

    def setUp(self):
        self.vb = ViewBuilder(MockApp())

    def test_screening_visit(self):
        visit, form, cat = self.vb._identify_column("SBV_LB_CBC_LBORRES_HGB")
        self.assertEqual(visit, VISIT_MAP["SBV"])

    def test_treatment_visit(self):
        visit, form, cat = self.vb._identify_column("TV_PR_PRSTDTC")
        self.assertEqual(visit, VISIT_MAP["TV"])

    def test_followup_visit(self):
        visit, form, cat = self.vb._identify_column("FU1M_SV_SVSTDTC")
        self.assertEqual(visit, VISIT_MAP["FU1M"])

    def test_logs_prefix(self):
        visit, form, cat = self.vb._identify_column("LOGS_AE_AETERM")
        self.assertEqual(visit, VISIT_MAP["LOGS"])

    def test_unscheduled_fallback(self):
        visit, form, cat = self.vb._identify_column("UNKNOWN_FIELD_XYZ")
        self.assertEqual(visit, "Unscheduled")

    def test_returns_tuple(self):
        result = self.vb._identify_column("SBV_VS_CVORRES_EDEMA")
        self.assertIsInstance(result, tuple)
        self.assertEqual(len(result), 3)

    def test_discharge_visit(self):
        visit, form, cat = self.vb._identify_column("DV_SV_SVSTDTC")
        self.assertEqual(visit, VISIT_MAP["DV"])

    def test_year_followup(self):
        visit, form, cat = self.vb._identify_column("FU1Y_ECHO_LVORRES")
        self.assertEqual(visit, VISIT_MAP["FU1Y"])


class TestAddGap(unittest.TestCase):
    """Test ViewBuilder.add_gap() recording."""

    def setUp(self):
        self.vb = ViewBuilder(MockApp())

    def test_gap_recorded(self):
        gaps = []
        self.vb.add_gap("Screening/Baseline", "Lab", "Hemoglobin", "SBV_LB_CBC_LBORRES_HGB", gaps)
        self.assertEqual(len(gaps), 1)
        self.assertEqual(gaps[0]['visit'], "Screening/Baseline")
        self.assertEqual(gaps[0]['form'], "Lab")
        self.assertEqual(gaps[0]['field'], "Hemoglobin")
        self.assertEqual(gaps[0]['variable'], "SBV_LB_CBC_LBORRES_HGB")
        self.assertIn('timestamp', gaps[0])

    def test_multiple_gaps(self):
        gaps = []
        self.vb.add_gap("Visit1", "Form1", "Field1", "COL1", gaps)
        self.vb.add_gap("Visit2", "Form2", "Field2", "COL2", gaps)
        self.assertEqual(len(gaps), 2)

    def test_gap_has_timestamp(self):
        gaps = []
        self.vb.add_gap("V", "F", "L", "C", gaps)
        self.assertIsNotNone(gaps[0]['timestamp'])


class TestGapCountIndex(unittest.TestCase):
    """Test gap count index building used in _render_tree."""

    def _build_index(self, gaps):
        gap_counts = {}
        for gap in gaps:
            key = (gap.get('visit', ''), gap.get('form', ''))
            gap_counts[key] = gap_counts.get(key, 0) + 1
        return gap_counts

    def test_empty_gaps(self):
        self.assertEqual(self._build_index([]), {})

    def test_single_gap(self):
        gaps = [{'visit': 'Screening/Baseline', 'form': 'Lab', 'field': 'HGB', 'variable': 'X'}]
        idx = self._build_index(gaps)
        self.assertEqual(idx[('Screening/Baseline', 'Lab')], 1)

    def test_multiple_same_form(self):
        gaps = [
            {'visit': 'Screening/Baseline', 'form': 'Lab', 'field': 'HGB', 'variable': 'X'},
            {'visit': 'Screening/Baseline', 'form': 'Lab', 'field': 'WBC', 'variable': 'Y'},
            {'visit': 'Screening/Baseline', 'form': 'Lab', 'field': 'PLT', 'variable': 'Z'},
        ]
        idx = self._build_index(gaps)
        self.assertEqual(idx[('Screening/Baseline', 'Lab')], 3)

    def test_different_visits(self):
        gaps = [
            {'visit': 'Screening/Baseline', 'form': 'Lab', 'field': 'HGB', 'variable': 'X'},
            {'visit': '30-Day Follow Up', 'form': 'Lab', 'field': 'HGB', 'variable': 'Y'},
        ]
        idx = self._build_index(gaps)
        self.assertEqual(idx[('Screening/Baseline', 'Lab')], 1)
        self.assertEqual(idx[('30-Day Follow Up', 'Lab')], 1)

    def test_visit_total(self):
        """Count gaps for a visit across all forms."""
        gaps = [
            {'visit': 'Screening/Baseline', 'form': 'Lab', 'field': 'HGB', 'variable': 'X'},
            {'visit': 'Screening/Baseline', 'form': 'Vitals', 'field': 'BP', 'variable': 'Y'},
            {'visit': '30-Day Follow Up', 'form': 'Lab', 'field': 'HGB', 'variable': 'Z'},
        ]
        idx = self._build_index(gaps)
        visit_total = sum(cnt for (v, f), cnt in idx.items() if v == 'Screening/Baseline')
        self.assertEqual(visit_total, 2)

    def test_different_forms_same_visit(self):
        gaps = [
            {'visit': 'V1', 'form': 'A', 'field': 'F1', 'variable': 'C1'},
            {'visit': 'V1', 'form': 'B', 'field': 'F2', 'variable': 'C2'},
        ]
        idx = self._build_index(gaps)
        self.assertEqual(idx[('V1', 'A')], 1)
        self.assertEqual(idx[('V1', 'B')], 1)


class TestCleanLabel(unittest.TestCase):
    """Test ViewBuilder._clean_label()."""

    def setUp(self):
        self.vb = ViewBuilder(MockApp())

    def test_removes_brackets(self):
        self.assertEqual(self.vb._clean_label("Hemoglobin [HGB]"), "Hemoglobin")

    def test_no_brackets(self):
        self.assertEqual(self.vb._clean_label("Hemoglobin"), "Hemoglobin")

    def test_multiple_brackets(self):
        self.assertEqual(self.vb._clean_label("A [B] C [D]"), "A  C")

    def test_empty(self):
        self.assertEqual(self.vb._clean_label(""), "")

    def test_numeric(self):
        self.assertEqual(self.vb._clean_label(123), "123")


class TestIsMatrixSupportedCol(unittest.TestCase):
    """Test ViewBuilder._is_matrix_supported_col()."""

    def setUp(self):
        self.vb = ViewBuilder(MockApp())

    def test_adverse_event(self):
        self.assertTrue(self.vb._is_matrix_supported_col("Adverse Event"))

    def test_laboratory(self):
        self.assertTrue(self.vb._is_matrix_supported_col("Laboratory"))

    def test_medical_history(self):
        self.assertTrue(self.vb._is_matrix_supported_col("Medical History"))

    def test_concomitant_meds(self):
        self.assertTrue(self.vb._is_matrix_supported_col("Concomitant Medications"))

    def test_not_supported(self):
        self.assertFalse(self.vb._is_matrix_supported_col("Demographics"))

    def test_partial_match(self):
        self.assertTrue(self.vb._is_matrix_supported_col("Laboratory Results"))


class TestViewBuilderCache(unittest.TestCase):
    """Test ViewBuilder cache operations."""

    def setUp(self):
        self.vb = ViewBuilder(MockApp())

    def test_initial_cache_empty(self):
        self.assertEqual(len(self.vb._view_cache), 0)

    def test_invalidate_cache(self):
        self.vb._view_cache["key"] = "value"
        self.vb.invalidate_cache()
        self.assertEqual(len(self.vb._view_cache), 0)

    def test_clear_cache(self):
        self.vb._view_cache["key1"] = "value1"
        self.vb._view_cache["key2"] = "value2"
        self.vb.clear_cache()
        self.assertEqual(len(self.vb._view_cache), 0)


if __name__ == "__main__":
    unittest.main()
