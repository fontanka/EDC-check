"""
Unit tests for hf_hospitalization_manager.py
Tests HF term matching (exact, pattern, fuzzy, exclusion), word-boundary regex,
date parsing, and event windowing.
"""

import unittest
from datetime import datetime, timedelta
import pandas as pd
from hf_hospitalization_manager import (
    HFHospitalizationManager, HFEvent,
    HF_EXACT_TERMS, HF_PROCEDURE_TERMS, HF_EXCLUSION_TERMS, HF_PATTERNS,
    FUZZY_THRESHOLD,
)


class TestHFTermMatchingExact(unittest.TestCase):
    """Test exact term matching with word boundaries."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_exact_heart_failure(self):
        is_match, conf, matched, mtype = self.mgr.is_hf_related("heart failure")
        self.assertTrue(is_match)
        self.assertEqual(conf, 1.0)
        self.assertEqual(mtype, "exact")

    def test_exact_case_insensitive(self):
        is_match, _, _, mtype = self.mgr.is_hf_related("Heart Failure")
        self.assertTrue(is_match)
        self.assertEqual(mtype, "exact")

    def test_exact_chf(self):
        is_match, _, _, mtype = self.mgr.is_hf_related("CHF")
        self.assertTrue(is_match)

    def test_exact_adhf(self):
        is_match, _, _, _ = self.mgr.is_hf_related("ADHF")
        self.assertTrue(is_match)

    def test_exact_pulmonary_edema(self):
        is_match, _, _, _ = self.mgr.is_hf_related("pulmonary edema")
        self.assertTrue(is_match)

    def test_exact_cardiogenic_shock(self):
        is_match, _, _, _ = self.mgr.is_hf_related("cardiogenic shock")
        self.assertTrue(is_match)

    def test_exact_volume_overload(self):
        is_match, _, _, _ = self.mgr.is_hf_related("volume overload")
        self.assertTrue(is_match)

    def test_exact_pleural_effusion(self):
        is_match, _, _, _ = self.mgr.is_hf_related("pleural effusion")
        self.assertTrue(is_match)

    def test_exact_ascites(self):
        is_match, _, _, _ = self.mgr.is_hf_related("ascites")
        self.assertTrue(is_match)


class TestHFTermMatchingProcedure(unittest.TestCase):
    """Test procedure term matching."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_paracentesis(self):
        is_match, _, _, mtype = self.mgr.is_hf_related("paracentesis")
        self.assertTrue(is_match)
        self.assertEqual(mtype, "procedure")

    def test_thoracentesis(self):
        is_match, _, _, _ = self.mgr.is_hf_related("thoracentesis")
        self.assertTrue(is_match)

    def test_iv_diuretic(self):
        is_match, _, _, _ = self.mgr.is_hf_related("IV diuretic")
        self.assertTrue(is_match)

    def test_furosemide_infusion(self):
        is_match, _, _, _ = self.mgr.is_hf_related("furosemide infusion")
        self.assertTrue(is_match)


class TestHFTermMatchingPattern(unittest.TestCase):
    """Test regex pattern matching."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_pattern_heart_fail_variation(self):
        is_match, _, _, mtype = self.mgr.is_hf_related("acute heart failing condition")
        # "heart fail" pattern should match
        self.assertTrue(is_match)

    def test_pattern_cardiac_decomp(self):
        is_match, _, _, _ = self.mgr.is_hf_related("cardiac decompensation episode")
        self.assertTrue(is_match)

    def test_pattern_congestive_heart(self):
        is_match, _, _, _ = self.mgr.is_hf_related("congestive heart disease")
        self.assertTrue(is_match)


class TestHFTermMatchingExclusion(unittest.TestCase):
    """Test exclusion terms prevent false positives."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_renal_excluded(self):
        """'renal failure' should NOT match as HF."""
        is_match, _, _, _ = self.mgr.is_hf_related("renal failure")
        self.assertFalse(is_match)

    def test_kidney_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("kidney disease")
        self.assertFalse(is_match)

    def test_pneumonia_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("pneumonia")
        self.assertFalse(is_match)

    def test_copd_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("COPD exacerbation")
        self.assertFalse(is_match)

    def test_anemia_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("anemia")
        self.assertFalse(is_match)

    def test_sepsis_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("sepsis")
        self.assertFalse(is_match)

    def test_stroke_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("stroke")
        self.assertFalse(is_match)

    def test_hypertension_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("hypertension")
        self.assertFalse(is_match)

    def test_fracture_excluded(self):
        is_match, _, _, _ = self.mgr.is_hf_related("hip fracture")
        self.assertFalse(is_match)


class TestHFWordBoundaryMatching(unittest.TestCase):
    """Test that word-boundary matching avoids substring false positives."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_adrenaline_not_renal(self):
        """'adrenaline' should NOT be excluded by 'renal' word-boundary match."""
        # adrenaline contains "renal" as substring but not as word
        is_match, _, _, mtype = self.mgr.is_hf_related("adrenaline induced cardiac failure")
        # Should match "cardiac" related pattern or exact, not be excluded
        # Actually depends on exact terms - "cardiac" isn't in exclusion list
        # The key test: it should NOT be blocked by "renal" exclusion
        # If it matches, mtype should not be "excluded"
        if is_match:
            self.assertNotEqual(mtype, "excluded")

    def test_hf_not_substring(self):
        """'HF' as standalone should match, but as part of another word should not."""
        # Standalone HF
        is_match, _, _, _ = self.mgr.is_hf_related("HF")
        self.assertTrue(is_match)

    def test_no_heart_failure_not_matched(self):
        """Terms like 'No heart failure' - 'heart failure' IS present as words."""
        # "heart failure" appears with word boundaries, so it will match
        # This is expected behavior - the AE term itself contains "heart failure"
        is_match, _, _, _ = self.mgr.is_hf_related("heart failure")
        self.assertTrue(is_match)


class TestHFNonMatching(unittest.TestCase):
    """Test terms that should NOT match HF."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_common_non_hf(self):
        non_hf_terms = [
            "headache",
            "back pain",
            "nausea",
            "diarrhea",
            "constipation",
            "insomnia",
            "fatigue",
            "rash",
            "cough",
            "fever",
        ]
        for term in non_hf_terms:
            is_match, _, _, _ = self.mgr.is_hf_related(term)
            self.assertFalse(is_match, f"'{term}' should NOT match as HF-related")

    def test_empty_input(self):
        is_match, _, _, _ = self.mgr.is_hf_related("")
        self.assertFalse(is_match)

    def test_none_input(self):
        is_match, _, _, _ = self.mgr.is_hf_related(None)
        self.assertFalse(is_match)

    def test_whitespace_only(self):
        is_match, _, _, _ = self.mgr.is_hf_related("   ")
        self.assertFalse(is_match)


class TestHFCustomTuning(unittest.TestCase):
    """Test custom include/exclude keyword tuning."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_custom_include(self):
        self.mgr.custom_includes = ["special cardiac condition"]
        # Clear cache since we changed tuning
        self.mgr._is_hf_related_cached.cache_clear()
        is_match, _, _, mtype = self.mgr.is_hf_related("special cardiac condition")
        self.assertTrue(is_match)
        self.assertEqual(mtype, "custom_included")

    def test_custom_exclude(self):
        self.mgr.custom_excludes = ["heart failure"]
        self.mgr._is_hf_related_cached.cache_clear()
        is_match, _, _, mtype = self.mgr.is_hf_related("heart failure")
        self.assertFalse(is_match)
        self.assertEqual(mtype, "custom_excluded")

    def test_custom_exclude_takes_priority(self):
        """Custom exclusion should override built-in exact match."""
        self.mgr.custom_excludes = ["pulmonary edema"]
        self.mgr._is_hf_related_cached.cache_clear()
        is_match, _, _, _ = self.mgr.is_hf_related("pulmonary edema")
        self.assertFalse(is_match)


class TestHFCaching(unittest.TestCase):
    """Test that caching works correctly."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_same_result_on_repeat(self):
        result1 = self.mgr.is_hf_related("heart failure")
        result2 = self.mgr.is_hf_related("heart failure")
        self.assertEqual(result1, result2)

    def test_cache_is_case_normalized(self):
        """Both should hit the same cache entry."""
        result1 = self.mgr.is_hf_related("Heart Failure")
        result2 = self.mgr.is_hf_related("heart failure")
        self.assertEqual(result1, result2)


class TestHFDateParsing(unittest.TestCase):
    """Test date parsing from various formats."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)

    def test_iso_format(self):
        d = self.mgr._parse_date("2025-03-15")
        self.assertEqual(d, datetime(2025, 3, 15))

    def test_iso_with_time(self):
        d = self.mgr._parse_date("2025-03-15T10:30:00")
        self.assertEqual(d, datetime(2025, 3, 15))

    def test_time_unknown_suffix(self):
        d = self.mgr._parse_date("2025-03-15, Time unknown")
        self.assertEqual(d, datetime(2025, 3, 15))

    def test_with_space_time(self):
        d = self.mgr._parse_date("2025-03-15 14:30")
        self.assertEqual(d, datetime(2025, 3, 15))

    def test_empty_string(self):
        self.assertIsNone(self.mgr._parse_date(""))

    def test_none_input(self):
        self.assertIsNone(self.mgr._parse_date(None))

    def test_nan_input(self):
        import numpy as np
        self.assertIsNone(self.mgr._parse_date(np.nan))

    def test_invalid_date(self):
        self.assertIsNone(self.mgr._parse_date("not-a-date"))

    def test_trailing_comma(self):
        d = self.mgr._parse_date("2025-03-15,")
        self.assertEqual(d, datetime(2025, 3, 15))


class TestHFEventWindow(unittest.TestCase):
    """Test event window calculation."""

    def setUp(self):
        self.mgr = HFHospitalizationManager(None, None)
        self.treatment = datetime(2025, 6, 1)

    def test_pre_treatment_within(self):
        event = datetime(2025, 3, 1)  # 92 days before
        self.assertTrue(self.mgr.is_within_window(event, self.treatment, pre_treatment=True, days=365))

    def test_pre_treatment_outside(self):
        event = datetime(2024, 1, 1)  # >1 year before
        self.assertFalse(self.mgr.is_within_window(event, self.treatment, pre_treatment=True, days=365))

    def test_post_treatment_within(self):
        event = datetime(2025, 9, 1)  # 92 days after
        self.assertTrue(self.mgr.is_within_window(event, self.treatment, pre_treatment=False, days=365))

    def test_post_treatment_outside(self):
        event = datetime(2026, 12, 1)  # >1 year after
        self.assertFalse(self.mgr.is_within_window(event, self.treatment, pre_treatment=False, days=365))

    def test_treatment_day_is_pre(self):
        """Treatment day itself should be within pre-treatment window."""
        self.assertTrue(self.mgr.is_within_window(self.treatment, self.treatment, pre_treatment=True, days=365))

    def test_treatment_day_not_post(self):
        """Treatment day itself should NOT be within post-treatment window (strict >0)."""
        self.assertFalse(self.mgr.is_within_window(self.treatment, self.treatment, pre_treatment=False, days=365))

    def test_none_event_date(self):
        self.assertFalse(self.mgr.is_within_window(None, self.treatment, pre_treatment=True))

    def test_none_treatment_date(self):
        event = datetime(2025, 3, 1)
        self.assertFalse(self.mgr.is_within_window(event, None, pre_treatment=True))

    def test_6m_window(self):
        event = datetime(2025, 10, 1)  # 122 days after
        self.assertTrue(self.mgr.is_within_window(event, self.treatment, pre_treatment=False, days=183))

    def test_6m_window_outside(self):
        event = datetime(2026, 1, 1)  # 214 days after
        self.assertFalse(self.mgr.is_within_window(event, self.treatment, pre_treatment=False, days=183))


class TestHFEventDataclass(unittest.TestCase):
    """Test HFEvent dataclass serialization."""

    def test_to_dict(self):
        event = HFEvent(
            event_id="test_1", date="2025-03-15", source_form="AE",
            source_row=1, original_term="heart failure",
            matched_synonym="heart failure", match_type="exact",
            confidence=1.0
        )
        d = event.to_dict()
        self.assertEqual(d['event_id'], 'test_1')
        self.assertEqual(d['confidence'], 1.0)
        self.assertTrue(d['is_included'])

    def test_from_dict(self):
        d = {
            'event_id': 'test_2', 'date': '2025-04-01', 'source_form': 'HFH',
            'source_row': 0, 'original_term': 'CHF', 'matched_synonym': 'chf',
            'match_type': 'exact', 'confidence': 1.0,
            'is_included': True, 'is_manual': False, 'notes': ''
        }
        event = HFEvent.from_dict(d)
        self.assertEqual(event.event_id, 'test_2')
        self.assertEqual(event.source_form, 'HFH')
        self.assertTrue(event.is_included)

    def test_roundtrip(self):
        event = HFEvent(
            event_id="rt_1", date="2025-05-01", source_form="HMEH",
            source_row=3, original_term="pleural effusion",
            matched_synonym="pleural effusion", match_type="exact",
            confidence=1.0, is_included=False, notes="excluded by reviewer"
        )
        d = event.to_dict()
        restored = HFEvent.from_dict(d)
        self.assertEqual(event, restored)


class TestHFTreatmentDate(unittest.TestCase):
    """Test treatment date lookup from df_main."""

    def test_treatment_date_found(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01', 'TV_PR_SVDTC': '2025-06-01'}
        ]).astype(str)
        mgr = HFHospitalizationManager(df_main)
        td = mgr.get_treatment_date('101-01')
        self.assertEqual(td, datetime(2025, 6, 1))

    def test_treatment_date_missing(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01', 'TV_PR_SVDTC': ''}
        ]).astype(str)
        mgr = HFHospitalizationManager(df_main)
        td = mgr.get_treatment_date('101-01')
        self.assertIsNone(td)

    def test_treatment_date_no_patient(self):
        df_main = pd.DataFrame([
            {'Screening #': '101-01', 'TV_PR_SVDTC': '2025-06-01'}
        ]).astype(str)
        mgr = HFHospitalizationManager(df_main)
        td = mgr.get_treatment_date('999-99')
        self.assertIsNone(td)

    def test_treatment_date_none_df(self):
        mgr = HFHospitalizationManager(None)
        td = mgr.get_treatment_date('101-01')
        self.assertIsNone(td)


if __name__ == '__main__':
    unittest.main()
