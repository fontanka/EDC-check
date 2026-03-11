"""Tests for dashboard_manager — _preprocess_data and _map_labels_and_aggregate."""
import unittest
import pandas as pd
from unittest.mock import MagicMock
from collections import defaultdict

from dashboard_manager import DashboardManager


def _make_sdv_mgr(modular_data):
    """Create a mock SDVManager with modular_data."""
    mgr = MagicMock()
    mgr.is_loaded.return_value = True
    mgr.modular_data = modular_data
    mgr.form_entry_status = {}
    return mgr


class TestPreprocessData(unittest.TestCase):
    def setUp(self):
        self.df = pd.DataFrame({
            'Subject Screening #': ['101-01', '102-02', '103-03'],
            'Variable Value': ['Hello', '', 'nan'],
            'Hidden': [0, 0, 1],
            'CRA_CONTROL_STATUS': [2, 0, 1],
        })
        self.mgr = _make_sdv_mgr(self.df)
        self.dm = DashboardManager(self.mgr)

    def test_basic_columns_created(self):
        result = self.dm._preprocess_data()
        self.assertIn('Patient', result.columns)
        self.assertIn('HasValue', result.columns)
        self.assertIn('Site', result.columns)

    def test_empty_values_flagged(self):
        result = self.dm._preprocess_data()
        # 'Hello' has value, '' and 'nan' do not
        has_values = result['HasValue'].tolist()
        self.assertEqual(has_values, [True, False, False])

    def test_site_extracted(self):
        result = self.dm._preprocess_data()
        sites = result['Site'].tolist()
        self.assertEqual(sites, ['101', '102', '103'])

    def test_exclusion_filter(self):
        result = self.dm._preprocess_data(excluded_patients=['102-02'])
        self.assertEqual(len(result), 2)
        self.assertNotIn('102-02', result['Patient'].values)

    def test_none_data(self):
        self.mgr.modular_data = None
        result = self.dm._preprocess_data()
        self.assertIsNone(result)


class TestMapLabelsAndAggregate(unittest.TestCase):
    def setUp(self):
        self.df_metric = pd.DataFrame({
            'Patient': ['101-01', '101-01', '102-02'],
            'Site': ['101', '101', '102'],
            'VisitName': ['Screening', 'Screening', 'Treatment'],
            'FormName': ['Demographics', 'Demographics', 'Procedure'],
            'Variable name': ['SBV_DM_BRTHDAT', 'SBV_DM_SEX', 'TV_PR_PRSTDTC'],
            'Value': ['1990-01-01', 'Female', '2025-03-15'],
            'Metric': ['V', '!', 'GAP'],
        })
        self.mgr = _make_sdv_mgr(pd.DataFrame())
        self.dm = DashboardManager(self.mgr)
        self.dm.labels = {
            'SBV_DM_BRTHDAT': 'Date of Birth',
            'SBV_DM_SEX': 'Sex',
        }
        self.dm.suffix_labels = {}
        self.dm.stats = {
            'study': defaultdict(int),
            'site': defaultdict(lambda: defaultdict(int)),
            'patient': defaultdict(lambda: defaultdict(int)),
            'form': defaultdict(lambda: defaultdict(int)),
        }

    def test_labels_mapped(self):
        self.dm._map_labels_and_aggregate(self.df_metric)
        fields = self.dm.details_df['Field'].tolist()
        self.assertIn('Date of Birth', fields)
        self.assertIn('Sex', fields)

    def test_study_stats(self):
        self.dm._map_labels_and_aggregate(self.df_metric)
        self.assertEqual(self.dm.stats['study']['V'], 1)
        self.assertEqual(self.dm.stats['study']['!'], 1)
        self.assertEqual(self.dm.stats['study']['GAP'], 1)

    def test_site_stats(self):
        self.dm._map_labels_and_aggregate(self.df_metric)
        self.assertEqual(self.dm.stats['site']['101']['V'], 1)
        self.assertEqual(self.dm.stats['site']['102']['GAP'], 1)

    def test_is_calculated(self):
        self.dm._map_labels_and_aggregate(self.df_metric)
        self.assertTrue(self.dm.is_calculated)

    def test_fallback_to_variable_name(self):
        """Variable without label should use variable name."""
        self.dm._map_labels_and_aggregate(self.df_metric)
        # TV_PR_PRSTDTC has no label mapping
        fields = self.dm.details_df['Field'].tolist()
        self.assertIn('TV_PR_PRSTDTC', fields)


if __name__ == '__main__':
    unittest.main()
