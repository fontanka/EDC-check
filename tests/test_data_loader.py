"""
Unit tests for data_loader.py
Tests file detection, sheet parsing, schema validation, and cross-form checks.
"""

import unittest
import os
import tempfile
import pandas as pd
import numpy as np
from datetime import datetime
from data_loader import (
    detect_latest_project_file,
    parse_cutoff_from_filename,
    load_project_file,
    validate_schema,
    validate_cross_form,
    LoadResult,
    _load_repeating_sheet,
    _safe_date,
    _check_fatal_ae_death_consistency,
    _check_procedure_before_followups,
    _check_ae_onset_after_procedure,
)


class TestDetectLatestProjectFile(unittest.TestCase):
    """Test auto-detection of ProjectToOneFile exports."""

    def test_no_files(self):
        with tempfile.TemporaryDirectory() as d:
            result = detect_latest_project_file(d)
            self.assertIsNone(result)

    def test_single_file(self):
        with tempfile.TemporaryDirectory() as d:
            name = "Innoventric_CLD-048_DM_ProjectToOneFile_15-01-2026_10-30_05_(UTC).xlsx"
            # Create an empty file
            open(os.path.join(d, name), 'w').close()
            result = detect_latest_project_file(d)
            self.assertIsNotNone(result)
            path, dt = result
            self.assertTrue(path.endswith(name))
            self.assertEqual(dt, datetime(2026, 1, 15, 10, 30, 5))

    def test_picks_latest(self):
        with tempfile.TemporaryDirectory() as d:
            old = "Innoventric_CLD-048_DM_ProjectToOneFile_01-01-2025_08-00_00_(UTC).xlsx"
            new = "Innoventric_CLD-048_DM_ProjectToOneFile_15-06-2025_14-30_00_(UTC).xlsx"
            for n in (old, new):
                open(os.path.join(d, n), 'w').close()
            result = detect_latest_project_file(d)
            self.assertIsNotNone(result)
            path, dt = result
            self.assertIn("15-06-2025", path)
            self.assertEqual(dt.year, 2025)
            self.assertEqual(dt.month, 6)

    def test_nonexistent_directory(self):
        result = detect_latest_project_file("/nonexistent/path/xyz")
        self.assertIsNone(result)

    def test_underscore_seconds_format(self):
        with tempfile.TemporaryDirectory() as d:
            name = "Innoventric_CLD-048_DM_ProjectToOneFile_04-01-2026_09-45_03_(UTC).xlsx"
            open(os.path.join(d, name), 'w').close()
            result = detect_latest_project_file(d)
            self.assertIsNotNone(result)
            _, dt = result
            self.assertEqual(dt, datetime(2026, 1, 4, 9, 45, 3))


class TestParseCutoffFromFilename(unittest.TestCase):
    """Test cutoff timestamp extraction from filename."""

    def test_standard_format(self):
        dt = parse_cutoff_from_filename(
            "Innoventric_CLD-048_DM_ProjectToOneFile_15-01-2026_10-30_05_(UTC).xlsx"
        )
        self.assertEqual(dt, datetime(2026, 1, 15, 10, 30, 5))

    def test_no_match(self):
        self.assertIsNone(parse_cutoff_from_filename("random_file.xlsx"))

    def test_empty_string(self):
        self.assertIsNone(parse_cutoff_from_filename(""))


class TestLoadRepeatingSheet(unittest.TestCase):
    """Test parsing of repeating-form sheets."""

    def test_basic(self):
        raw = pd.DataFrame([
            ['Col\xa0A', 'Col B', 'Screening #'],
            ['val1', 'val2', '101-01'],
            ['val3', 'val4', '101-02'],
        ])
        df = _load_repeating_sheet(raw)
        self.assertIsNotNone(df)
        self.assertEqual(list(df.columns), ['Col A', 'Col B', 'Screening #'])
        self.assertEqual(len(df), 2)

    def test_empty_sheet(self):
        result = _load_repeating_sheet(pd.DataFrame())
        self.assertIsNone(result)


class TestValidateSchema(unittest.TestCase):
    """Test schema validation."""

    def test_all_present(self):
        cols = ['Screening #', 'Site #', 'Status', 'TV_PR_PRSTDTC']
        warnings = validate_schema(cols, required_cols=cols)
        self.assertEqual(warnings, [])

    def test_missing_columns(self):
        cols = ['Screening #', 'Site #']
        warnings = validate_schema(cols, required_cols=['Screening #', 'Site #', 'MISSING_COL'])
        self.assertEqual(len(warnings), 1)
        self.assertIn('MISSING_COL', warnings[0])


class TestSafeDate(unittest.TestCase):
    """Test safe date parsing utility."""

    def test_valid_iso(self):
        dt = _safe_date('2025-03-15')
        self.assertEqual(dt.year, 2025)
        self.assertEqual(dt.month, 3)

    def test_empty(self):
        self.assertIsNone(_safe_date(''))
        self.assertIsNone(_safe_date(None))
        self.assertIsNone(_safe_date('nan'))
        self.assertIsNone(_safe_date('none'))

    def test_nan(self):
        self.assertIsNone(_safe_date(np.nan))

    def test_invalid(self):
        self.assertIsNone(_safe_date('not-a-date'))


class TestCrossFormFatalAE(unittest.TestCase):
    """Test fatal AE ↔ death form consistency."""

    def test_fatal_ae_no_death_date(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01', 'LOGS_DTH_DDDTC': ''},
        ]).astype(str)
        df_ae = pd.DataFrame([
            {'Screening #': '101-01', 'LOGS_AE_AEOUT': 'Fatal'},
        ]).astype(str)
        issues = []
        _check_fatal_ae_death_consistency(df_main, df_ae, issues)
        self.assertEqual(len(issues), 1)
        self.assertIn('101-01', issues[0])
        self.assertIn('empty', issues[0])

    def test_fatal_ae_with_death_date(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01', 'LOGS_DTH_DDDTC': '2025-04-10'},
        ]).astype(str)
        df_ae = pd.DataFrame([
            {'Screening #': '101-01', 'LOGS_AE_AEOUT': 'Fatal'},
        ]).astype(str)
        issues = []
        _check_fatal_ae_death_consistency(df_main, df_ae, issues)
        self.assertEqual(len(issues), 0)

    def test_no_fatal_aes(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01', 'LOGS_DTH_DDDTC': ''},
        ]).astype(str)
        df_ae = pd.DataFrame([
            {'Screening #': '101-01', 'LOGS_AE_AEOUT': 'Recovered'},
        ]).astype(str)
        issues = []
        _check_fatal_ae_death_consistency(df_main, df_ae, issues)
        self.assertEqual(len(issues), 0)

    def test_no_ae_data(self):
        df_main = pd.DataFrame([{'Screening #': '101-01'}]).astype(str)
        issues = []
        _check_fatal_ae_death_consistency(df_main, None, issues)
        self.assertEqual(len(issues), 0)


class TestCrossFormProcedureBeforeFollowups(unittest.TestCase):
    """Test procedure date precedes follow-up dates."""

    def test_fu_before_procedure(self):
        df_main = pd.DataFrame([{
            'Screening #': '101-01',
            'TV_PR_PRSTDTC': '2025-06-01',
            'FU1M_SV_SVSTDTC': '2025-05-01',  # Before procedure!
        }]).astype(str)
        issues = []
        _check_procedure_before_followups(df_main, issues)
        self.assertEqual(len(issues), 1)
        self.assertIn('FU1M', issues[0])
        self.assertIn('precedes procedure', issues[0])

    def test_fu_after_procedure(self):
        df_main = pd.DataFrame([{
            'Screening #': '101-01',
            'TV_PR_PRSTDTC': '2025-06-01',
            'FU1M_SV_SVSTDTC': '2025-07-01',
        }]).astype(str)
        issues = []
        _check_procedure_before_followups(df_main, issues)
        self.assertEqual(len(issues), 0)

    def test_no_procedure_date(self):
        df_main = pd.DataFrame([{
            'Screening #': '101-01',
            'TV_PR_PRSTDTC': '',
            'FU1M_SV_SVSTDTC': '2025-07-01',
        }]).astype(str)
        issues = []
        _check_procedure_before_followups(df_main, issues)
        self.assertEqual(len(issues), 0)


class TestCrossFormAEOnsetAfterProcedure(unittest.TestCase):
    """Test AE onset date vs procedure date consistency."""

    def test_post_ae_before_procedure(self):
        df_main = pd.DataFrame([{
            'Screening #': '101-01',
            'TV_PR_PRSTDTC': '2025-06-01',
        }]).astype(str)
        df_ae = pd.DataFrame([{
            'Screening #': '101-01',
            'LOGS_AE_AESTDTC': '2025-05-15',
            'LOGS_AE_AEINT': '',  # not marked as pre-procedure
            'LOGS_AE_AETERM': 'Headache',
        }]).astype(str)
        issues = []
        _check_ae_onset_after_procedure(df_main, df_ae, issues)
        self.assertEqual(len(issues), 1)
        self.assertIn('not marked pre-procedure', issues[0])

    def test_pre_procedure_ae_ok(self):
        df_main = pd.DataFrame([{
            'Screening #': '101-01',
            'TV_PR_PRSTDTC': '2025-06-01',
        }]).astype(str)
        df_ae = pd.DataFrame([{
            'Screening #': '101-01',
            'LOGS_AE_AESTDTC': '2025-05-15',
            'LOGS_AE_AEINT': 'Pre-treatment',
            'LOGS_AE_AETERM': 'Headache',
        }]).astype(str)
        issues = []
        _check_ae_onset_after_procedure(df_main, df_ae, issues)
        self.assertEqual(len(issues), 0)

    def test_post_ae_after_procedure_ok(self):
        df_main = pd.DataFrame([{
            'Screening #': '101-01',
            'TV_PR_PRSTDTC': '2025-06-01',
        }]).astype(str)
        df_ae = pd.DataFrame([{
            'Screening #': '101-01',
            'LOGS_AE_AESTDTC': '2025-06-15',
            'LOGS_AE_AEINT': '',
            'LOGS_AE_AETERM': 'Headache',
        }]).astype(str)
        issues = []
        _check_ae_onset_after_procedure(df_main, df_ae, issues)
        self.assertEqual(len(issues), 0)


class TestValidateCrossForm(unittest.TestCase):
    """Test the top-level cross-form validation function."""

    def test_empty_data(self):
        result = LoadResult(df_main=pd.DataFrame())
        issues = validate_cross_form(result)
        self.assertEqual(issues, [])

    def test_combined_issues(self):
        df_main = pd.DataFrame([{
            'Screening #': '101-01',
            'TV_PR_PRSTDTC': '2025-06-01',
            'LOGS_DTH_DDDTC': '',
            'FU1M_SV_SVSTDTC': '2025-05-01',
        }]).astype(str)
        df_ae = pd.DataFrame([
            {'Screening #': '101-01', 'LOGS_AE_AEOUT': 'Fatal',
             'LOGS_AE_AESTDTC': '2025-07-01', 'LOGS_AE_AEINT': '',
             'LOGS_AE_AETERM': 'Cardiac arrest'},
        ]).astype(str)
        result = LoadResult(df_main=df_main, df_ae=df_ae)
        issues = validate_cross_form(result)
        # Should have at least: fatal AE without death date + FU before procedure
        self.assertGreaterEqual(len(issues), 2)


class TestLoadResultDataclass(unittest.TestCase):
    """Test LoadResult container."""

    def test_defaults(self):
        result = LoadResult(df_main=pd.DataFrame())
        self.assertIsNone(result.df_ae)
        self.assertIsNone(result.df_cm)
        self.assertEqual(result.labels, {})
        self.assertEqual(result.warnings, [])


if __name__ == '__main__':
    unittest.main()
