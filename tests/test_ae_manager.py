"""
Unit tests for ae_manager.py
Tests column mapping, date parsing, filtering, statistics, and screen failure detection.
"""

import unittest
import pandas as pd
import numpy as np
from ae_manager import AEManager


def _make_df_main(rows=None):
    """Create a minimal df_main for testing."""
    if rows is None:
        rows = [
            {'Screening #': '101-01', 'Status': 'Enrolled', 'Site #': '101',
             'TV_PR_PRSTDTC': '2025-03-15'},
            {'Screening #': '101-02', 'Status': 'Screen Failure', 'Site #': '101',
             'TV_PR_PRSTDTC': ''},
            {'Screening #': '102-01', 'Status': 'Enrolled', 'Site #': '102',
             'TV_PR_PRSTDTC': '2025-04-01'},
        ]
    return pd.DataFrame(rows).astype(str)


def _make_df_ae(rows=None):
    """Create a minimal df_ae for testing."""
    if rows is None:
        rows = [
            {'Screening #': '101-01', 'Template number': '1',
             'LOGS_AE_AESER': 'Yes', 'LOGS_AE_AETERM': 'Headache',
             'LOGS_AE_AESEV': 'Mild', 'LOGS_AE_AESTDTC': '2025-03-20',
             'LOGS_AE_AEENDTC': '2025-03-25', 'LOGS_AE_AEONGO': 'No',
             'LOGS_AE_AEOUT': 'Recovered', 'LOGS_AE_AEREL1': 'Not Related',
             'LOGS_AE_AEREL2': 'Not Related', 'LOGS_AE_AEREL3': 'Not Related',
             'LOGS_AE_AEREL4': 'Not Related', 'LOGS_AE_AEREPDAT': '2025-03-20',
             'LOGS_AE_AESHOSP': 'No', 'LOGS_AE_AESLIFE': 'No',
             'LOGS_AE_AESDTH': 'No', 'LOGS_AE_AESDISAB': 'No',
             'LOGS_AE_AESMIE': 'No'},
            {'Screening #': '101-01', 'Template number': '2',
             'LOGS_AE_AESER': 'No', 'LOGS_AE_AETERM': 'Nausea',
             'LOGS_AE_AESEV': 'Moderate', 'LOGS_AE_AESTDTC': '2025-03-10',
             'LOGS_AE_AEENDTC': '', 'LOGS_AE_AEONGO': 'Yes',
             'LOGS_AE_AEOUT': '', 'LOGS_AE_AEREL1': 'Related',
             'LOGS_AE_AEREL2': 'Not Related', 'LOGS_AE_AEREL3': 'Not Related',
             'LOGS_AE_AEREL4': 'Possibly Related', 'LOGS_AE_AEREPDAT': '2025-03-22',
             'LOGS_AE_AESHOSP': 'No', 'LOGS_AE_AESLIFE': 'No',
             'LOGS_AE_AESDTH': 'No', 'LOGS_AE_AESDISAB': 'No',
             'LOGS_AE_AESMIE': 'No'},
            {'Screening #': '101-02', 'Template number': '1',
             'LOGS_AE_AESER': 'No', 'LOGS_AE_AETERM': 'Dizziness',
             'LOGS_AE_AESEV': 'Mild', 'LOGS_AE_AESTDTC': '2025-02-01',
             'LOGS_AE_AEENDTC': '2025-02-05', 'LOGS_AE_AEONGO': 'No',
             'LOGS_AE_AEOUT': 'Recovered', 'LOGS_AE_AEREL1': 'Not Related',
             'LOGS_AE_AEREL2': 'Not Related', 'LOGS_AE_AEREL3': 'Not Related',
             'LOGS_AE_AEREL4': 'Not Related', 'LOGS_AE_AEREPDAT': '2025-02-01',
             'LOGS_AE_AESHOSP': 'No', 'LOGS_AE_AESLIFE': 'No',
             'LOGS_AE_AESDTH': 'No', 'LOGS_AE_AESDISAB': 'No',
             'LOGS_AE_AESMIE': 'No'},
            {'Screening #': '102-01', 'Template number': '1',
             'LOGS_AE_AESER': 'Yes', 'LOGS_AE_AETERM': 'Cardiac arrest',
             'LOGS_AE_AESEV': 'Severe', 'LOGS_AE_AESTDTC': '2025-04-10',
             'LOGS_AE_AEENDTC': '2025-04-10', 'LOGS_AE_AEONGO': 'No',
             'LOGS_AE_AEOUT': 'Fatal', 'LOGS_AE_AEREL1': 'Related',
             'LOGS_AE_AEREL2': 'Related', 'LOGS_AE_AEREL3': 'Not Related',
             'LOGS_AE_AEREL4': 'Related', 'LOGS_AE_AEREPDAT': '2025-04-10',
             'LOGS_AE_AESHOSP': 'Yes', 'LOGS_AE_AESLIFE': 'Yes',
             'LOGS_AE_AESDTH': 'Yes', 'LOGS_AE_AESDISAB': 'No',
             'LOGS_AE_AESMIE': 'No'},
        ]
    return pd.DataFrame(rows).astype(str)


class TestAEManagerColumnMapping(unittest.TestCase):
    """Test that column mapping resolves correctly."""

    def setUp(self):
        self.mgr = AEManager(_make_df_main(), _make_df_ae())

    def test_find_col_primary(self):
        """Primary LOGS_AE_* columns should resolve."""
        self.assertEqual(self.mgr._find_col('AE Term'), 'LOGS_AE_AETERM')
        self.assertEqual(self.mgr._find_col('SAE?'), 'LOGS_AE_AESER')
        self.assertEqual(self.mgr._find_col('Onset Date'), 'LOGS_AE_AESTDTC')
        self.assertEqual(self.mgr._find_col('Severity'), 'LOGS_AE_AESEV')

    def test_find_col_ae_number(self):
        """AE # should resolve to 'Template number'."""
        self.assertEqual(self.mgr._find_col('AE #'), 'Template number')

    def test_find_col_missing(self):
        """Non-existent display name should return None."""
        self.assertIsNone(self.mgr._find_col('NonExistentField'))

    def test_find_col_fallback(self):
        """When primary column is missing, should fall back to alternatives."""
        # Create AE df with only fallback column names
        df_ae = pd.DataFrame([
            {'Screening #': '101-01', 'AE #': '1', 'AETERM': 'Test'}
        ]).astype(str)
        mgr = AEManager(_make_df_main(), df_ae)
        self.assertEqual(mgr._find_col('AE Term'), 'AETERM')


class TestAEManagerDateParsing(unittest.TestCase):
    """Test date cleaning and parsing."""

    def setUp(self):
        self.mgr = AEManager(_make_df_main(), _make_df_ae())

    def test_clean_date_iso(self):
        self.assertEqual(self.mgr._clean_date('2025-03-15T10:30:00'), '2025-03-15')

    def test_clean_date_time_unknown(self):
        self.assertEqual(self.mgr._clean_date('2025-03-15, time unknown'), '2025-03-15')

    def test_clean_date_with_time_space(self):
        self.assertEqual(self.mgr._clean_date('2025-03-15 14:30'), '2025-03-15')

    def test_clean_date_plain(self):
        self.assertEqual(self.mgr._clean_date('2025-03-15'), '2025-03-15')

    def test_parse_date_obj_valid(self):
        from datetime import date
        result = self.mgr._parse_date_obj('2025-03-15')
        self.assertEqual(result, date(2025, 3, 15))

    def test_parse_date_obj_empty(self):
        self.assertIsNone(self.mgr._parse_date_obj(''))
        self.assertIsNone(self.mgr._parse_date_obj(None))

    def test_parse_date_obj_invalid(self):
        self.assertIsNone(self.mgr._parse_date_obj('not-a-date'))

    def test_normalize_boolean(self):
        self.assertEqual(self.mgr._normalize_boolean('Yes'), 'Yes')
        self.assertEqual(self.mgr._normalize_boolean('yes'), 'Yes')
        self.assertEqual(self.mgr._normalize_boolean('y'), 'Yes')
        self.assertEqual(self.mgr._normalize_boolean('1'), 'Yes')
        self.assertEqual(self.mgr._normalize_boolean('No'), 'No')
        self.assertEqual(self.mgr._normalize_boolean('no'), 'No')
        self.assertEqual(self.mgr._normalize_boolean('n'), 'No')
        self.assertEqual(self.mgr._normalize_boolean(''), '')

    def test_is_checked(self):
        self.assertTrue(self.mgr._is_checked('Yes'))
        self.assertTrue(self.mgr._is_checked('checked'))
        self.assertFalse(self.mgr._is_checked('No'))
        self.assertFalse(self.mgr._is_checked(''))
        self.assertFalse(self.mgr._is_checked('nan'))


class TestAEManagerScreenFailures(unittest.TestCase):
    """Test screen failure detection."""

    def test_screen_failures_detected(self):
        mgr = AEManager(_make_df_main(), _make_df_ae())
        sf = mgr.get_screen_failures()
        self.assertIn('101-02', sf)
        self.assertNotIn('101-01', sf)
        self.assertNotIn('102-01', sf)

    def test_screen_failures_empty(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01', 'Status': 'Enrolled'}
        ]).astype(str)
        mgr = AEManager(df_main, _make_df_ae())
        self.assertEqual(mgr.get_screen_failures(), [])

    def test_screen_failures_no_status_col(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01'}
        ]).astype(str)
        mgr = AEManager(df_main, _make_df_ae())
        self.assertEqual(mgr.get_screen_failures(), [])

    def test_screen_failures_cached(self):
        mgr = AEManager(_make_df_main(), _make_df_ae())
        sf1 = mgr.get_screen_failures()
        sf2 = mgr.get_screen_failures()
        self.assertIs(sf1, sf2)  # Same object (cached)


class TestAEManagerPatientData(unittest.TestCase):
    """Test patient AE data retrieval and filtering."""

    def setUp(self):
        self.mgr = AEManager(_make_df_main(), _make_df_ae())

    def test_get_patient_ae_data_basic(self):
        data = self.mgr.get_patient_ae_data('101-01')
        self.assertEqual(len(data), 2)
        terms = {row['AE Term'] for row in data}
        self.assertIn('Headache', terms)
        self.assertIn('Nausea', terms)

    def test_get_patient_ae_data_nonexistent(self):
        data = self.mgr.get_patient_ae_data('999-99')
        self.assertEqual(len(data), 0)

    def test_filter_sae_only(self):
        data = self.mgr.get_patient_ae_data('101-01', {'sae_only': True})
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]['AE Term'], 'Headache')
        self.assertEqual(data[0]['SAE?'], 'Yes')

    def test_filter_exclude_pre_proc(self):
        """Nausea onset 2025-03-10 is before procedure 2025-03-15, should be excluded."""
        data = self.mgr.get_patient_ae_data('101-01', {'exclude_pre_proc': True})
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]['AE Term'], 'Headache')

    def test_filter_device_related(self):
        """Only AEs with non-'Not Related' relationship should pass."""
        data = self.mgr.get_patient_ae_data('101-01', {'device_related_only': True})
        # Nausea has AEREL1='Related', so it passes
        # Headache has all 'Not Related', so it's excluded
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]['AE Term'], 'Nausea')

    def test_filter_onset_cutoff(self):
        data = self.mgr.get_patient_ae_data('101-01', {'onset_cutoff': '2025-03-15'})
        # Only Headache (onset 2025-03-20) is after cutoff
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]['AE Term'], 'Headache')

    def test_filter_report_cutoff(self):
        data = self.mgr.get_patient_ae_data('101-01', {'report_cutoff': '2025-03-21'})
        # Only Nausea (report 2025-03-22) is after cutoff
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]['AE Term'], 'Nausea')

    def test_ongoing_ae_resolution_date(self):
        data = self.mgr.get_patient_ae_data('101-01')
        nausea = [r for r in data if r['AE Term'] == 'Nausea'][0]
        self.assertEqual(nausea['Resolution Date'], 'Ongoing')

    def test_empty_ae_dataframe(self):
        mgr = AEManager(_make_df_main(), pd.DataFrame())
        data = mgr.get_patient_ae_data('101-01')
        self.assertEqual(len(data), 0)


class TestAEManagerSummaryStats(unittest.TestCase):
    """Test summary statistics calculation."""

    def setUp(self):
        self.mgr = AEManager(_make_df_main(), _make_df_ae())

    def test_total_counts(self):
        stats = self.mgr.get_summary_stats()
        self.assertEqual(stats['total_aes'], 4)
        self.assertEqual(stats['total_saes'], 2)
        self.assertEqual(stats['patients_with_aes'], 3)

    def test_fatal_count(self):
        stats = self.mgr.get_summary_stats()
        self.assertEqual(stats['fatal_cases'], 1)

    def test_exclude_patients(self):
        stats = self.mgr.get_summary_stats(excluded_patients=['101-02'])
        self.assertEqual(stats['total_aes'], 3)
        self.assertEqual(stats['patients_with_aes'], 2)

    def test_exclude_screen_failures(self):
        stats = self.mgr.get_summary_stats(exclude_screen_failures=True)
        # 101-02 is a screen failure, should be excluded
        self.assertEqual(stats['total_aes'], 3)
        self.assertEqual(stats['patients_with_aes'], 2)

    def test_outcome_distribution(self):
        stats = self.mgr.get_summary_stats()
        outcomes = stats['outcome_dist']
        self.assertIn('Recovered', outcomes)
        self.assertIn('Fatal', outcomes)

    def test_sae_criteria(self):
        stats = self.mgr.get_summary_stats()
        criteria = stats['sae_criteria']
        self.assertEqual(criteria['Hospitalization'], 1)
        self.assertEqual(criteria['Life-threatening'], 1)
        self.assertEqual(criteria['Death'], 1)

    def test_by_site(self):
        stats = self.mgr.get_summary_stats()
        sites = stats['by_site']
        self.assertIn('101', sites)
        self.assertIn('102', sites)

    def test_top_terms(self):
        stats = self.mgr.get_summary_stats()
        terms = stats['top_terms']
        self.assertIn('Headache', terms)

    def test_relatedness_table(self):
        stats = self.mgr.get_summary_stats()
        rel = stats['relatedness_table']
        self.assertIn('Device', rel)
        self.assertIn('Procedure', rel)
        # 101-01 AE#2 has AEREL1=Related, 102-01 AE#1 has AEREL1=Related
        self.assertEqual(rel['Device']['Related'], 2)

    def test_per_patient_details(self):
        stats = self.mgr.get_summary_stats()
        details = stats['per_patient_details']
        self.assertTrue(len(details) > 0)
        self.assertTrue(any('101-01' in d for d in details))

    def test_empty_ae_df(self):
        mgr = AEManager(_make_df_main(), pd.DataFrame())
        stats = mgr.get_summary_stats()
        self.assertEqual(stats['total_aes'], 0)


class TestAEManagerProcedureDate(unittest.TestCase):
    """Test procedure date lookup."""

    def setUp(self):
        self.mgr = AEManager(_make_df_main(), _make_df_ae())

    def test_procedure_date_found(self):
        from datetime import date
        d = self.mgr._get_procedure_date('101-01')
        self.assertEqual(d, date(2025, 3, 15))

    def test_procedure_date_missing(self):
        d = self.mgr._get_procedure_date('101-02')
        self.assertIsNone(d)

    def test_procedure_date_cached(self):
        self.mgr._get_procedure_date('101-01')
        self.assertIn('101-01', self.mgr._procedure_dates)

    def test_procedure_date_nonexistent_patient(self):
        d = self.mgr._get_procedure_date('999-99')
        self.assertIsNone(d)


class TestAEManagerDatasetExport(unittest.TestCase):
    """Test bulk dataset export."""

    def setUp(self):
        self.mgr = AEManager(_make_df_main(), _make_df_ae())

    def test_get_dataset_all(self):
        data = self.mgr.get_dataset_ae_data()
        self.assertEqual(len(data), 4)
        # Each row should have Patient ID
        for row in data:
            self.assertIn('Patient ID', row)

    def test_get_dataset_filtered(self):
        data = self.mgr.get_dataset_ae_data({'sae_only': True})
        self.assertEqual(len(data), 2)


class TestAEManagerDeathDetails(unittest.TestCase):
    """Test death/mortality data in summary stats."""

    def test_death_details_present(self):
        df_main = pd.DataFrame([
            {'Screening #': '102-01', 'Status': 'Enrolled', 'TV_PR_PRSTDTC': '2025-04-01',
             'LOGS_DTH_DDDTC': '2025-04-10', 'LOGS_DTH_DDRESCAT': 'Cardiovascular',
             'LOGS_DTH_DDORRES': 'Cardiac arrest'},
            {'Screening #': '101-01', 'Status': 'Enrolled', 'TV_PR_PRSTDTC': '2025-03-15',
             'LOGS_DTH_DDDTC': '', 'LOGS_DTH_DDRESCAT': '', 'LOGS_DTH_DDORRES': ''},
        ]).astype(str)
        mgr = AEManager(df_main, _make_df_ae())
        stats = mgr.get_summary_stats()
        death_details = stats['death_details']
        self.assertEqual(len(death_details), 1)
        self.assertEqual(death_details[0]['patient_id'], '102-01')
        self.assertEqual(death_details[0]['mortality_classification'], 'Cardiovascular')
        self.assertEqual(death_details[0]['cause_of_death'], 'Cardiac arrest')

    def test_death_details_empty(self):
        mgr = AEManager(_make_df_main(), _make_df_ae())
        stats = mgr.get_summary_stats()
        # No death columns in default df_main, so no death details
        self.assertEqual(stats['death_details'], [])


if __name__ == '__main__':
    unittest.main()
