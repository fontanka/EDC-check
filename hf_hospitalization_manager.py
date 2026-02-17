"""
Heart Failure Hospitalization Manager

Parses and tracks HF-related hospitalizations from clinical data sources:
- Pre-treatment: HFH, HMEH, CVH forms (HFH is primary, others deduplicated)
- Post-treatment: AE sheet (using symptom onset date - AESTDTC)

Supports fuzzy matching for HF synonyms and manual event editing.
"""

import pandas as pd
import numpy as np
from datetime import datetime, timedelta
from dataclasses import dataclass, asdict
from difflib import get_close_matches
from functools import lru_cache
from typing import List, Tuple, Dict, Optional
import json
import logging
import os
import re
from difflib import SequenceMatcher

logger = logging.getLogger("ClinicalViewer.HFManager")


# ============================================================================
# HF SYNONYM DICTIONARY
# ============================================================================

# Exact match terms (case-insensitive, checked first)
HF_EXACT_TERMS = [
    # Core HF terms
    "heart failure",
    "hf",
    "chf",
    "congestive heart failure",
    "acute heart failure",
    "acute on chronic heart failure",
    "acute-on-chronic heart failure",
    "chronic heart failure",
    "heart failure exacerbation",
    "hf exacerbation",
    "chf exacerbation",
    "decompensated heart failure",
    "acute decompensated heart failure",
    "adhf",
    "cardiac decompensation",
    "left heart failure",
    "right heart failure",
    "left ventricular failure",
    "right ventricular failure",
    "biventricular failure",
    "cardiogenic shock",
    "cardiogenic pulmonary edema",
    "pulmonary edema",
    "cardiac pulmonary edema",
    "flash pulmonary edema",
    "volume overload",
    "fluid overload",
    "cardiac fluid overload",
    # Fluid accumulation symptoms
    "ascites",
    "cardiac ascites",
    "pericardial effusion",
    "pleural effusion",
    "peripheral edema",
    "lower extremity edema",
    "leg edema",
    "anasarca",
]

# Procedure terms that indicate HF hospitalization (100% related to HF)
HF_PROCEDURE_TERMS = [
    "paracentesis",
    "abdominal paracentesis",
    "therapeutic paracentesis",
    "thoracentesis",
    "pleural drainage",
    "pleural tap",
    "ultrafiltration",
    "aquapheresis",
    "diuretic infusion",
    "iv diuretic",
    "intravenous diuretic",
    "furosemide infusion",
    "lasix infusion",
    "bumetanide infusion",
]

# Pattern-based matching (regex patterns for fuzzy matching)
HF_PATTERNS = [
    r"heart\s*fail",
    r"hf\s+exac",
    r"chf\s+exac",
    r"cardiac\s+decomp",
    r"decomp.*heart",
    r"congest.*heart",
    r"pulmon.*edema",
    r"fluid\s+overload",
    r"volume\s+overload",
]

# EXCLUSION TERMS - if term contains any of these, it's NOT HF-related
# This prevents false positives from fuzzy matching
HF_EXCLUSION_TERMS = [
    # Kidney-related
    "kidney",
    "renal",
    "aki",
    "ckd",
    "nephro",
    "dialysis",
    "creatinine",
    "uremia",
    # Liver-related
    "liver",
    "hepatic",
    "cirrhosis",
    "hepato",
    # Respiratory (not cardiac)
    "copd",
    "asthma",
    "pneumonia",
    "bronchitis",
    "respiratory failure",
    # Blood/Hematological
    "anemia",
    "anaemia",
    # Electrolyte/metabolic disturbances
    "hypokalemia",
    "hyponatremia",
    "hyperkalemia",
    # Blood pressure (not HF)
    "hypotension",
    "hypertension",
    # Other non-HF
    "sepsis",
    "cancer",
    "tumor",
    "fracture",
    "stroke",
    "cva",
    "fall",
    "cellulitis",
    "wound",
    "ulcer",
    "shoulder",
    "premature ventricular contractions",
    "pvc",
]

# Fuzzy matching threshold (0-1, higher = stricter)
# Increased from 0.75 to 0.85 to reduce false positives
FUZZY_THRESHOLD = 0.85


# ============================================================================
# DATA CLASSES
# ============================================================================

@dataclass
class HFEvent:
    """Represents a single HF hospitalization event."""
    event_id: str           # Unique identifier
    date: str               # Event date (ISO format)
    source_form: str        # HFH, HMEH, CVH, or AE
    source_row: int         # Row number in source form
    original_term: str      # Original text from source
    matched_synonym: str    # Which synonym matched
    match_type: str         # 'exact', 'pattern', 'fuzzy', 'procedure', 'manual'
    confidence: float       # Match confidence (0-1)
    is_included: bool = True    # Whether to count this event
    is_manual: bool = False     # Whether manually added/edited
    notes: str = ""         # User notes
    
    def to_dict(self):
        return asdict(self)
    
    @staticmethod
    def from_dict(d):
        return HFEvent(**d)


# ============================================================================
# HF HOSPITALIZATION MANAGER
# ============================================================================

class HFHospitalizationManager:
    """
    Manages HF hospitalization data parsing, matching, and persistence.
    """
    
    def __init__(self, df_main: pd.DataFrame, df_ae: pd.DataFrame = None):
        """
        Initialize with clinical data.
        
        Args:
            df_main: Main sheet DataFrame with patient data
            df_ae: AE sheet DataFrame (optional, for post-treatment events)
        """
        self.df_main = df_main
        self.df_ae = df_ae
        self.manual_edits = {}  # patient_id -> List[HFEvent]
        self._load_manual_edits()
        
        # Custom keyword tuning
        self.custom_includes = []
        self.custom_excludes = []
        self._load_tuning_config()
        
        # Build column mappings for each form
        self._build_column_mappings()
    
    def _build_column_mappings(self):
        """Identify relevant columns for each source form."""
        self.hfh_cols = {}  # Columns containing HFH data
        self.hmeh_cols = {}
        self.cvh_cols = {}
        
        if self.df_main is None:
            return
            
        for col in self.df_main.columns:
            col_str = str(col)
            # HFH columns (Heart Failure History)
            if "_HFH_" in col_str:
                suffix = col_str.split("_HFH_")[-1] if "_HFH_" in col_str else ""
                if suffix not in self.hfh_cols:
                    self.hfh_cols[suffix] = []
                self.hfh_cols[suffix].append(col_str)
            
            # HMEH columns (Hospitalization and Medical Events History)
            if "HMEH_" in col_str or "_HMEH_" in col_str:
                suffix = col_str.split("HMEH_")[-1] if "HMEH_" in col_str else ""
                if suffix not in self.hmeh_cols:
                    self.hmeh_cols[suffix] = []
                self.hmeh_cols[suffix].append(col_str)
            
            # CVH columns (Cardiovascular History)
            if "_CVH_" in col_str:
                suffix = col_str.split("_CVH_")[-1] if "_CVH_" in col_str else ""
                if suffix not in self.cvh_cols:
                    self.cvh_cols[suffix] = []
                self.cvh_cols[suffix].append(col_str)
    
    def _load_manual_edits(self):
        """Load manually edited events from JSON file."""
        try:
            edits_path = os.path.join(os.path.dirname(__file__), "hf_manual_edits.json")
            if os.path.exists(edits_path):
                with open(edits_path, 'r') as f:
                    data = json.load(f)
                    for patient_id, events in data.items():
                        self.manual_edits[patient_id] = [HFEvent.from_dict(e) for e in events]
        except Exception as e:
            logger.warning("Could not load HF manual edits: %s", e)
    
    def _load_tuning_config(self):
        """Load custom tuning keywords from JSON file."""
        try:
            tuning_path = os.path.join(os.path.dirname(__file__), "hf_tuning.json")
            if os.path.exists(tuning_path):
                with open(tuning_path, 'r') as f:
                    data = json.load(f)
                    self.custom_includes = [str(s).lower().strip() for s in data.get('includes', [])]
                    self.custom_excludes = [str(s).lower().strip() for s in data.get('excludes', [])]
        except Exception as e:
            logger.warning("Could not load HF tuning config: %s", e)

    def save_tuning_config(self):
        """Save custom tuning keywords to JSON file."""
        try:
            tuning_path = os.path.join(os.path.dirname(__file__), "hf_tuning.json")
            data = {
                'includes': self.custom_includes,
                'excludes': self.custom_excludes
            }
            with open(tuning_path, 'w') as f:
                json.dump(data, f, indent=2)
        except Exception as e:
            logger.warning("Could not save HF tuning config: %s", e)

    def save_manual_edits(self):
        """Save manually edited events to JSON file."""
        try:
            edits_path = os.path.join(os.path.dirname(__file__), "hf_manual_edits.json")
            data = {}
            for patient_id, events in self.manual_edits.items():
                data[patient_id] = [e.to_dict() for e in events]
            with open(edits_path, 'w') as f:
                json.dump(data, f, indent=2, default=str)
        except Exception as e:
            logger.warning("Could not save HF manual edits: %s", e)
    
    # -------------------------------------------------------------------------
    # HF Term Matching
    # -------------------------------------------------------------------------
    
    @lru_cache(maxsize=1000)
    def _is_hf_related_cached(self, term_lower: str) -> Tuple[bool, float, str, str]:
        """Cached inner implementation of is_hf_related (keyed by normalized term)."""
        return self._is_hf_related_impl(term_lower)

    def is_hf_related(self, term: str) -> Tuple[bool, float, str, str]:
        """
        Check if a term is HF-related.

        Returns:
            Tuple of (is_match, confidence, matched_term, match_type)
        """
        if not term or not isinstance(term, str):
            return False, 0.0, "", ""

        term_lower = term.lower().strip()
        if not term_lower:
            return False, 0.0, "", ""
        return self._is_hf_related_cached(term_lower)

    def _is_hf_related_impl(self, term_lower: str) -> Tuple[bool, float, str, str]:
        """Core HF classification logic (called via cache)."""

        def _wb_match(needle, haystack):
            """Word-boundary match to avoid substring false positives."""
            return bool(re.search(r'\b' + re.escape(needle) + r'\b', haystack))

        # 0. Check custom exclusions FIRST
        for excl in self.custom_excludes:
            if _wb_match(excl, term_lower):
                return False, 0.0, excl, "custom_excluded"

        # 0.1 Check hardcoded exclusions
        for excl in HF_EXCLUSION_TERMS:
            if _wb_match(excl, term_lower):
                return False, 0.0, "", "excluded"

        # 0.2 Check custom inclusions
        for incl in self.custom_includes:
            if _wb_match(incl, term_lower):
                return True, 1.0, incl, "custom_included"

        # 1. Exact match (highest priority)
        for hf_term in HF_EXACT_TERMS:
            if _wb_match(hf_term, term_lower):
                return True, 1.0, hf_term, "exact"

        # 2. Procedure terms
        for proc_term in HF_PROCEDURE_TERMS:
            if _wb_match(proc_term, term_lower):
                return True, 1.0, proc_term, "procedure"
        
        # 3. Pattern matching
        for pattern in HF_PATTERNS:
            if re.search(pattern, term_lower, re.IGNORECASE):
                return True, 0.95, pattern, "pattern"
        
        # 4. Fuzzy matching (for typos, variations)
        # Use get_close_matches for efficient fuzzy search instead of triple-nested loop
        all_hf_terms = HF_EXACT_TERMS + HF_PROCEDURE_TERMS
        matches = get_close_matches(term_lower, all_hf_terms, n=1, cutoff=FUZZY_THRESHOLD)

        if matches:
            best_match = matches[0]
            best_score = SequenceMatcher(None, term_lower, best_match).ratio()
            return True, best_score, best_match, "fuzzy"

        return False, 0.0, "", ""
    
    # -------------------------------------------------------------------------
    # Date Handling
    # -------------------------------------------------------------------------
    
    def get_treatment_date(self, patient_id: str) -> Optional[datetime]:
        """Get the treatment date (TV_PR_SVDTC) for a patient."""
        if self.df_main is None:
            return None
        
        # Find patient row
        mask = self.df_main['Screening #'].astype(str).str.replace('.0', '') == str(patient_id).replace('.0', '')
        patient_rows = self.df_main[mask]
        
        if patient_rows.empty:
            return None
        
        row = patient_rows.iloc[0]
        
        # Look for treatment date column
        treatment_col = None
        for col in self.df_main.columns:
            if 'TV_PR_SVDTC' in str(col):
                treatment_col = col
                break
        
        if treatment_col is None:
            return None
        
        date_val = row.get(treatment_col)
        return self._parse_date(date_val)
    
    def _parse_date(self, date_val) -> Optional[datetime]:
        """Parse a date value from various formats."""
        if pd.isna(date_val) or not date_val:
            return None
        
        date_str = str(date_val).strip()
        
        # Remove ", Time unknown" suffix first (before space splitting)
        if ', Time unknown' in date_str:
            date_str = date_str.replace(', Time unknown', '')
        
        # Remove time portion from ISO format
        if 'T' in date_str:
            date_str = date_str.split('T')[0]
        
        # Remove any remaining time portion after space
        if ' ' in date_str:
            date_str = date_str.split(' ')[0]
        
        # Remove any trailing punctuation
        date_str = date_str.rstrip(',;')
        
        # Try common formats
        for fmt in ['%Y-%m-%d', '%d-%m-%Y', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        
        return None
    
    def is_within_window(self, event_date: datetime, treatment_date: datetime, 
                       pre_treatment: bool, days: int = 365) -> bool:
        """Check if event is within generic window before/after treatment."""
        if event_date is None or treatment_date is None:
            return False
        
        delta = event_date - treatment_date
        
        if pre_treatment:
            # days before treatment (including treatment day as pre)
            return timedelta(days=-days) <= delta <= timedelta(days=0)
        else:
            # days after treatment (treatment day = post)
            return timedelta(days=0) < delta <= timedelta(days=days)
    
    # -------------------------------------------------------------------------
    # Event Parsing
    # -------------------------------------------------------------------------
    
    def _get_patient_row(self, patient_id: str) -> Optional[pd.Series]:
        """Get the patient row from main dataframe."""
        if self.df_main is None:
            return None
        
        mask = self.df_main['Screening #'].astype(str).str.replace('.0', '') == str(patient_id).replace('.0', '')
        patient_rows = self.df_main[mask]
        
        if patient_rows.empty:
            return None
        
        return patient_rows.iloc[0]
    
    def parse_hfh_events(self, patient_id: str) -> List[HFEvent]:
        """
        Parse Heart Failure History events (PRIMARY source).
        
        HFH columns typically contain:
        - HOSTDTC: Hospitalization date
        - HOTERM: Description/reason
        - HONUM: Number of hospitalizations (count field)
        """
        events = []
        row = self._get_patient_row(patient_id)
        
        if row is None:
            return events
        
        # Find HFH columns and parse events
        hfh_data = {}  # Group columns by entry number
        
        for col in self.df_main.columns:
            col_str = str(col)
            
            if "_HFH_" not in col_str and "HFH_" not in col_str:
                continue
            
            val = row.get(col)
            if pd.isna(val) or str(val).strip() in ['', 'nan']:
                continue
            
            # Parse piped data format: #1 / 2024-01-15 / Heart failure admission
            val_str = str(val).strip()
            
            if val_str.startswith('#') and '/' in val_str:
                # Piped format: #entry_num / date / details
                entries = val_str.split('|') if '|' in val_str else [val_str]
                for entry in entries:
                    parts = [p.strip() for p in entry.split('/')]
                    if len(parts) >= 2:
                        entry_num = parts[0].replace('#', '').strip()
                        date_str = parts[1].strip() if len(parts) > 1 else ""
                        term = parts[2].strip() if len(parts) > 2 else ""
                        
                        if entry_num not in hfh_data:
                            hfh_data[entry_num] = {'date': '', 'term': '', 'count': ''}
                        
                        if 'HOSTDTC' in col_str or 'DTC' in col_str:
                            hfh_data[entry_num]['date'] = date_str
                        if 'HOTERM' in col_str or 'TERM' in col_str or 'DESC' in col_str:
                            hfh_data[entry_num]['term'] = term if term else val_str
                        if 'HONUM' in col_str or 'NUM' in col_str:
                            hfh_data[entry_num]['count'] = val_str
            else:
                # Simple value format - check if it contains HF terms
                is_hf, conf, matched, match_type = self.is_hf_related(val_str)
                if is_hf:
                    # Extract date from related column if possible
                    date_col = col_str.replace('HOTERM', 'HOSTDTC').replace('DESC', 'DTC')
                    date_val = row.get(date_col) if date_col in self.df_main.columns else None
                    
                    event = HFEvent(
                        event_id=f"HFH_{patient_id}_{len(events)}",
                        date=str(self._parse_date(date_val) or ""),
                        source_form="HFH",
                        source_row=0,
                        original_term=val_str,
                        matched_synonym=matched,
                        match_type=match_type,
                        confidence=conf
                    )
                    events.append(event)
        
        # Process grouped HFH data
        for entry_num, data in hfh_data.items():
            term = data.get('term', '')
            if not term:
                continue
            
            is_hf, conf, matched, match_type = self.is_hf_related(term)
            if is_hf:
                event = HFEvent(
                    event_id=f"HFH_{patient_id}_{entry_num}",
                    date=data.get('date', ''),
                    source_form="HFH",
                    source_row=int(entry_num) if entry_num.isdigit() else 0,
                    original_term=term,
                    matched_synonym=matched,
                    match_type=match_type,
                    confidence=conf
                )
                events.append(event)
        
        return events
    
    def parse_hmeh_events(self, patient_id: str, existing_dates: set) -> List[HFEvent]:
        """
        Parse Hospitalization and Medical Events History (secondary source).
        Deduplicates against existing dates.
        """
        events = []
        row = self._get_patient_row(patient_id)
        
        if row is None:
            return events
        
        for col in self.df_main.columns:
            col_str = str(col)
                
            if "HMEH_" not in col_str and "_HMEH_" not in col_str:
                continue
            
            val = row.get(col)
            if pd.isna(val) or str(val).strip() in ['', 'nan']:
                continue
            
            val_str = str(val).strip()
            is_hf, conf, matched, match_type = self.is_hf_related(val_str)
            
            if is_hf:
                # Try to extract date from related column
                date_col = col_str.replace('HOTERM', 'HOSTDTC').replace('TERM', 'DTC').replace('DESC', 'DTC')
                date_val = row.get(date_col) if date_col in self.df_main.columns else None
                event_date = self._parse_date(date_val)
                date_str = str(event_date.date()) if event_date else ""
                
                # Deduplicate by date
                if date_str and date_str in existing_dates:
                    continue
                
                event = HFEvent(
                    event_id=f"HMEH_{patient_id}_{len(events)}",
                    date=date_str,
                    source_form="HMEH",
                    source_row=0,
                    original_term=val_str,
                    matched_synonym=matched,
                    match_type=match_type,
                    confidence=conf
                )
                events.append(event)
                if date_str:
                    existing_dates.add(date_str)
        
        return events
    
    def parse_cvh_events(self, patient_id: str, existing_dates: set) -> List[HFEvent]:
        """
        Parse Cardiovascular History (for HF-related procedures like paracentesis).
        Deduplicates against existing dates.
        """
        events = []
        row = self._get_patient_row(patient_id)
        
        if row is None:
            return events
        
        for col in self.df_main.columns:
            col_str = str(col)
            if "_CVH_" not in col_str:
                continue
            
            val = row.get(col)
            if pd.isna(val) or str(val).strip() in ['', 'nan']:
                continue
            
            val_str = str(val).strip()
            is_hf, conf, matched, match_type = self.is_hf_related(val_str)
            
            if is_hf:
                # Try to extract date
                date_col = col_str.replace('PRTRT', 'PRSTDTC').replace('TERM', 'DTC')
                date_val = row.get(date_col) if date_col in self.df_main.columns else None
                event_date = self._parse_date(date_val)
                date_str = str(event_date.date()) if event_date else ""
                
                # Deduplicate
                if date_str and date_str in existing_dates:
                    continue
                
                event = HFEvent(
                    event_id=f"CVH_{patient_id}_{len(events)}",
                    date=date_str,
                    source_form="CVH",
                    source_row=0,
                    original_term=val_str,
                    matched_synonym=matched,
                    match_type=match_type,
                    confidence=conf
                )
                events.append(event)
                if date_str:
                    existing_dates.add(date_str)
        
        return events
    
    def parse_mh_events(self, patient_id: str, existing_dates: set) -> List[HFEvent]:
        """
        Parse Medical History (MH) form.
        Looks for columns with 'MH' but not 'HMEH' (to avoid duplicates).
        """
        events = []
        row = self._get_patient_row(patient_id)
        if row is None:
            return events
            
        for col in self.df_main.columns:
            col_str = str(col)
            # Must contain MH but not HMEH (subset)
            if "MH" not in col_str:
                continue
            if "HMEH" in col_str or "HFH" in col_str:
                continue
                
            val = row.get(col)
            if pd.isna(val) or str(val).strip() in ['', 'nan']:
                continue
            
            val_str = str(val).strip()
            is_hf, conf, matched, match_type = self.is_hf_related(val_str)
            
            if is_hf:
                # Try to extract date
                # Heuristic: Find a corresponding date column. MH usually has MHSTDTC or similar.
                # Try replacing TERM/DESC with STDTC/DTC
                date_col = None
                if 'TERM' in col_str:
                    date_col = col_str.replace('TERM', 'STDTC').replace('TERM', 'DTC')
                elif 'DESC' in col_str:
                    date_col = col_str.replace('DESC', 'STDTC').replace('DESC', 'DTC')
                
                # If we have a generic value column, look for the date column in the same group?
                # Assuming simple suffix replacement for now.
                
                date_val = None
                if date_col and date_col in self.df_main.columns:
                    date_val = row.get(date_col)
                
                # Dedupe
                event_date = self._parse_date(date_val)
                date_str = str(event_date.date()) if event_date else ""
                
                if date_str and date_str in existing_dates:
                    continue
                
                event = HFEvent(
                    event_id=f"MH_{patient_id}_{len(events)}",
                    date=date_str,
                    source_form="MH",
                    source_row=0,
                    original_term=val_str,
                    matched_synonym=matched,
                    match_type=match_type,
                    confidence=conf
                )
                events.append(event)
                if date_str:
                    existing_dates.add(date_str)
                    
        return events
    
    def parse_ae_events(self, patient_id: str) -> List[HFEvent]:
        """
        Parse Adverse Events for post-treatment HF hospitalizations.
        Uses AESTDTC (symptom onset date) for date filtering.
        """
        events = []
        
        if self.df_ae is None or self.df_ae.empty:
            logger.debug("parse_ae_events: No AE data available")
            return events

        # Filter to patient
        mask = self.df_ae['Screening #'].astype(str).str.contains(
            patient_id.replace('-', '-'), na=False
        )
        patient_aes = self.df_ae[mask]
        logger.debug("parse_ae_events(%s): Found %d AE rows", patient_id, len(patient_aes))
        
        for idx, ae_row in patient_aes.iterrows():
            # Get AE term
            ae_term = ae_row.get('LOGS_AE_AETERM', '')
            if pd.isna(ae_term) or not str(ae_term).strip():
                continue
            
            
            ae_term = str(ae_term).strip()
            is_hf, conf, matched, match_type = self.is_hf_related(ae_term)
            
            if is_hf:
                # Get onset date (AESTDTC)
                onset_date = ae_row.get('LOGS_AE_AESTDTC', '')
                event_date = self._parse_date(onset_date)
                date_str = str(event_date.date()) if event_date else ""
                
                # Get AE number for display
                ae_num = ae_row.get('Template number', '')
                
                # Use DataFrame index for unique event_id (ae_num can be duplicated)
                event = HFEvent(
                    event_id=f"AE_{patient_id}_{idx}",
                    date=date_str,
                    source_form="AE",
                    source_row=int(ae_num) if str(ae_num).isdigit() else idx,
                    original_term=ae_term,
                    matched_synonym=matched,
                    match_type=match_type,
                    confidence=conf
                )
                events.append(event)
        
        return events
    
    # -------------------------------------------------------------------------
    # Main API Methods
    # -------------------------------------------------------------------------
    
    def get_pre_treatment_events(self, patient_id: str) -> List[HFEvent]:
        """
        Get all HF events within 1 year BEFORE treatment.
        Priority: HFH (primary) -> HMEH (dedupe) -> CVH (dedupe)
        """
        treatment_date = self.get_treatment_date(patient_id)
        
        # Collect events from each source with deduplication
        existing_dates = set()
        
        # 1. HFH - Primary source
        hfh_events = self.parse_hfh_events(patient_id)
        for e in hfh_events:
            if e.date:
                existing_dates.add(e.date)
        
        # 2. HMEH - Deduplicate
        hmeh_events = self.parse_hmeh_events(patient_id, existing_dates)
        
        # 3. CVH - Deduplicate
        cvh_events = self.parse_cvh_events(patient_id, existing_dates)

        # 4. MH - Medical History (New)
        mh_events = self.parse_mh_events(patient_id, existing_dates)
        
        all_events = hfh_events + hmeh_events + cvh_events + mh_events
        
        # Filter to 1 year before treatment (include events with no date)
        if treatment_date:
            logger.debug(
                "get_pre_treatment(%s): treatment_date=%s, parsed %d events from HFH/HMEH/CVH/MH",
                patient_id, treatment_date.date(), len(all_events),
            )
            filtered = []
            for event in all_events:
                event_date = self._parse_date(event.date)
                if event_date is None:
                    # Include events with missing dates (conservative)
                    logger.debug("  Pre-Event '%s' date=None -> INCLUDED (no date)", event.original_term)
                    filtered.append(event)
                else:
                    is_within = self.is_within_window(event_date, treatment_date, pre_treatment=True, days=365)
                    logger.debug(
                        "  Pre-Event '%s' date=%s -> within_1y=%s",
                        event.original_term, event.date, is_within,
                    )
                    if is_within:
                        filtered.append(event)
            all_events = filtered
            logger.debug("  After filtration: %d events", len(all_events))
        
        # Apply manual edits
        all_events = self._apply_manual_edits(patient_id, all_events, pre_treatment=True)
        
        return all_events
    
    def get_post_treatment_events(self, patient_id: str) -> List[HFEvent]:
        """
        Get all HF events AFTER treatment (from AE sheet).
        Uses AESTDTC (symptom onset date) for filtering.
        Events with no parseable date are included (conservative approach).
        """
        treatment_date = self.get_treatment_date(patient_id)

        # Parse AE events
        ae_events = self.parse_ae_events(patient_id)
        logger.debug("get_post_treatment(%s): %d raw AE events, treatment_date=%s",
                      patient_id, len(ae_events),
                      treatment_date.date() if treatment_date else "None")

        # Filter to post-treatment window (include events with no date)
        if treatment_date:
            filtered = []
            for event in ae_events:
                event_date = self._parse_date(event.date)
                if event_date is None:
                    # Include events with missing dates (conservative)
                    logger.debug("  Post-Event '%s' date=None -> INCLUDED (no date)",
                                 event.original_term)
                    filtered.append(event)
                elif self.is_within_window(event_date, treatment_date, pre_treatment=False, days=365*5):
                    filtered.append(event)
                else:
                    logger.debug("  Post-Event '%s' date=%s -> EXCLUDED (outside window)",
                                 event.original_term, event.date)
            ae_events = filtered
        
        # Apply manual edits
        ae_events = self._apply_manual_edits(patient_id, ae_events, pre_treatment=False)
        
        return ae_events
    
    def _apply_manual_edits(self, patient_id: str, events: List[HFEvent], 
                           pre_treatment: bool) -> List[HFEvent]:
        """Apply manual edits (exclusions, additions) to event list."""
        if patient_id not in self.manual_edits:
            return events
        
        manual = self.manual_edits[patient_id]
        
        # Build lookup of manual events by ID
        manual_by_id = {e.event_id: e for e in manual}
        
        # Update events with manual overrides
        result = []
        for event in events:
            if event.event_id in manual_by_id:
                # Use manual version
                manual_event = manual_by_id[event.event_id]
                result.append(manual_event)
                del manual_by_id[event.event_id]
            else:
                result.append(event)
        
        # Add any manually added events
        for event_id, event in manual_by_id.items():
            if event.is_manual:
                # Check if this event is for the right period
                if pre_treatment and event.source_form in ['HFH', 'HMEH', 'CVH', 'MANUAL_PRE']:
                    result.append(event)
                elif not pre_treatment and event.source_form in ['AE', 'MANUAL_POST']:
                    result.append(event)
        
        return result
    
    def get_patient_summary(self, patient_id: str) -> dict:
        """
        Get summary statistics for a patient.
        
        Returns:
            dict with keys: pre_count, post_count, pre_events, post_events, treatment_date
        """
        treatment_date = self.get_treatment_date(patient_id)
        pre_events = self.get_pre_treatment_events(patient_id)
        post_events = self.get_post_treatment_events(patient_id)
        
        # Calculate 1Y counts (unique dates)
        pre_dates_1y = {e.date for e in pre_events if e.is_included and e.date}
        post_dates_1y = {e.date for e in post_events if e.is_included and e.date}
        pre_count_1y = len(pre_dates_1y)
        post_count_1y = len(post_dates_1y)
        
        # Calculate 6m subsets (unique dates)
        pre_dates_6m = set()
        if treatment_date:
            for e in pre_events:
                if e.is_included and e.date:
                    edate = self._parse_date(e.date)
                    if self.is_within_window(edate, treatment_date, pre_treatment=True, days=183):
                        pre_dates_6m.add(e.date)
        pre_count_6m = len(pre_dates_6m)
        
        post_dates_6m = set()
        if treatment_date:
            for e in post_events:
                if e.is_included and e.date:
                    edate = self._parse_date(e.date)
                    if self.is_within_window(edate, treatment_date, pre_treatment=False, days=183):
                        post_dates_6m.add(e.date)
        post_count_6m = len(post_dates_6m)
        
        return {
            'patient_id': patient_id,
            'treatment_date': str(treatment_date.date()) if treatment_date else "",
            'pre_count': pre_count_6m,  # Keep as 6M for backward compat or clarifying
            'post_count': post_count_6m, # Keep as 6M
            'pre_count_6m': pre_count_6m,
            'post_count_6m': post_count_6m,
            'pre_count_1y': pre_count_1y,
            'post_count_1y': post_count_1y,
            'pre_events': pre_events,
            'post_events': post_events
        }
    
    def get_all_patients_summary(self) -> List[dict]:
        """Get summary for all patients in the dataset."""
        if self.df_main is None:
            return []
        
        summaries = []
        patients = self.df_main['Screening #'].dropna().unique()
        
        for patient_id in patients:
            patient_str = str(patient_id).replace('.0', '')
            summary = self.get_patient_summary(patient_str)
            # Only include patients with treatment date or events
            if summary['treatment_date'] or summary['pre_count'] > 0 or summary['post_count'] > 0:
                summaries.append(summary)
        
        return summaries
    
    def update_event(self, patient_id: str, event: HFEvent):
        """Update or add an event in manual edits."""
        if patient_id not in self.manual_edits:
            self.manual_edits[patient_id] = []
        
        # Find and update existing, or add new
        found = False
        for i, e in enumerate(self.manual_edits[patient_id]):
            if e.event_id == event.event_id:
                self.manual_edits[patient_id][i] = event
                found = True
                break
        
        if not found:
            self.manual_edits[patient_id].append(event)
        
        self.save_manual_edits()
    
    def delete_event(self, patient_id: str, event_id: str):
        """Delete an event from manual edits."""
        if patient_id not in self.manual_edits:
            return
        
        self.manual_edits[patient_id] = [
            e for e in self.manual_edits[patient_id] if e.event_id != event_id
        ]
        self.save_manual_edits()


# ============================================================================
# TESTING / DEMO
# ============================================================================

if __name__ == "__main__":
    # Quick test of HF term matching
    manager = HFHospitalizationManager(None, None)
    
    test_terms = [
        "Heart failure exacerbation",
        "CHF decompensation",
        "Admission for paracentesis",
        "Acute on chronic heart failure",
        "Pneumonia",
        "Cardiac arrhythmia",
        "Fluid overload requiring IV diuretics",
        "hf excerbation",  # typo
        "congestive hearfailure",  # typo
    ]
    
    logging.basicConfig(level=logging.DEBUG)
    logger.info("HF Term Matching Test:")
    logger.info("-" * 60)
    for term in test_terms:
        is_hf, conf, matched, match_type = manager.is_hf_related(term)
        status = "HF" if is_hf else "Not HF"
        logger.info("%s (%.2f) [%-8s] %s", status, conf, match_type, term)
        if is_hf:
            logger.debug("    -> Matched: %s", matched)
    logger.info("-" * 60)
