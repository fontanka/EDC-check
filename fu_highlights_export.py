# FU Highlights Export Module
# Extracts Follow-up Highlights data for copy-paste into presentations

import base64
import datetime
from datetime import timedelta
import difflib
import io
import logging
import re
import matplotlib.pyplot as plt

logger = logging.getLogger("ClinicalViewer.FUHighlights")
import matplotlib.dates as mdates
import numpy as np
import pandas as pd
from matplotlib.patches import Patch

# Visit configuration for FU Highlights
# Note: Echo data uses different prefixes than other assessments
FU_VISITS = {
    "Screening": {"prefix": "SBV_", "echo_prefix": "SBV_ECHO_SPONSOR_", "label": "Baseline", "date_col": "SBV_SV_SVSTDTC"},
    "Discharge": {"prefix": "DV_", "echo_prefix": "DV_ECHO_1D_SPONSOR_", "label": "Discharge", "date_col": "DV_SV_SVSTDTC"},
    "30D": {"prefix": "FU1M_", "echo_prefix": "FU1M_ECHO_1D_SPONSOR_", "label": "30-Day", "date_col": "FU1M_SV_SVSTDTC"},
    "3M": {"prefix": "FU3M_", "echo_prefix": "FU3M_ECHO_1D_SPONSOR_", "label": "3-Month", "date_col": "FU3M_SV_SVSTDTC"},
    "6M": {"prefix": "FU6M_", "echo_prefix": "FU6M_ECHO_1D_SPONSOR_", "label": "6-Month", "date_col": "FU6M_SV_SVSTDTC"},
    "1Y": {"prefix": "FU1Y_", "echo_prefix": "FU1Y_ECHO_1D_SPONSOR_", "label": "1-Year", "date_col": "FU1Y_SV_SVSTDTC"},
    "2Y": {"prefix": "FU2Y_", "echo_prefix": "FU2Y_ECHO_1D_SPONSOR_", "label": "2-Year", "date_col": "FU2Y_SV_SVSTDTC"},
    "4Y": {"prefix": "FU4Y_", "echo_prefix": "FU4Y_ECHO_1D_SPONSOR_", "label": "4-Year", "date_col": "FU4Y_SV_SVSTDTC"},
}

# Visit order for sequential display (Screening through all FUs)
VISIT_ORDER = ["Screening", "Discharge", "30D", "3M", "6M", "1Y", "2Y", "4Y"]

# Parameter mapping: (display_name, column_pattern, value_type)
# {PREFIX} = standard prefix (SBV_, FU1M_, etc.)
# {ECHO_PREFIX} = echo prefix (SBV_ECHO_SPONSOR_, FU1M_ECHO_1D_SPONSOR_, etc.)
# Based on actual DB Variable names from the app tree view and echo_export.py

FU_HIGHLIGHT_PARAMS = [
    # Clinical assessments
    ("NYHA", "{PREFIX}FS_RSORRES_FSNYHA", "value"),
    ("CFS", "{PREFIX}CFSS_RSORRES_CFSS", "value"),
    
    # Physical examination (Edema and Ascites from VS section)
    ("Edema Level (0-4)", "{PREFIX}VS_CVORRES_EDEMA", "value"),
    ("Ascites (0-3)", "{PREFIX}VS_CVORRES_ASCITIS", "value"),
    
    # Echo TR severity - direct column patterns
    ("TR severity (Hepatic backflow)", "{ECHO_PREFIX}FAORRES_ECHO1_SP", "value"),
    ("TR severity (Color Doppler/VCW)", "{ECHO_PREFIX}FAORRES_ECHO2_SP", "value"),
    
    # Cardiac Output from Echo Sponsor - ECHO34 column (was mistakenly ECHO31 which is RVSP)
    ("CO (based on Echo) [L/min]", "SPECIAL_CO_LOGIC", "value"),
    
    # 6MWT - distance walked
    ("6MWT result [m]", "{PREFIX}6MWT_FTORRES_DIS", "value"),  # Note: FU visits use 6MWT_FU_
    
    # KCCQ scores
    ("KCCQ Overall Summary Score", "{PREFIX}KCCQ_QSORRES_KCCQ_OVERALL", "value"),
    ("KCCQ Clinical Summary Score", "{PREFIX}KCCQ_QSORRES_KCCQ_CLINICAL", "value"),
]

VITAL_SIGNS_PARAMS = [
    ("Date of Visit", "{PREFIX}SV_SVSTDTC", "value"),
    ("Systolic Blood Pressure [mmHg]", "{PREFIX}VS_VSORRES_SYSBP", "value"),
    ("Diastolic Blood Pressure [mmHg]", "{PREFIX}VS_VSORRES_DIABP", "value"),
    ("Heart Rate [bpm]", "{PREFIX}VS_VSORRES_HR", "value"),
    ("Respiratory Rate [breaths/min]", "{PREFIX}VS_VSORRES_RESP", "value"),
    ("Weight [kg]", "{PREFIX}VS_VSORRES_WEIGHT", "value"),
]


class FUHighlightsExporter:
    def __init__(self, df_main, fuzzy_confirm_callback=None):
        self.df_main = df_main
        self.fuzzy_confirm_callback = fuzzy_confirm_callback
    
    def simplify_nyha(self, val):
        """Extract just the Roman numeral or number from NYHA string."""
        if not val or val == "NA":
            return val
        # Look for I, II, III, IV or numbers 1-4
        match = re.search(r'\b(I{1,3}V?|IV|[1-4])\b', str(val), re.IGNORECASE)
        if match:
            return match.group(1).upper()
        return val


    def simplify_cfs(self, val):
        # Look for leading number at start of string
        match = re.match(r'^(\d+)', str(val).strip())
        if match:
            return match.group(1)
        return val
    
    def simplify_grade(self, val):
        """Extract grade number (0, 1, 2, 3, 4) from Edema/Ascites text."""
        if not val or val == "NA":
            return "NA"
        import re
        # Look for leading number or "0 -", "1 -", etc.
        match = re.match(r'^(\d+)', str(val).strip())
        if match:
            return match.group(1)
        return val
    
    def get_value(self, row, column_pattern, prefix, echo_prefix=None):
        """Get value from a column using the pattern and prefix."""
        # Special logic for Cardiac Output which changes column ID between visits
        if column_pattern == "SPECIAL_CO_LOGIC":
            if prefix.startswith("SBV"):
                # Screening/Baseline uses ECHO34
                col = f"{echo_prefix}FAORRES_ECHO34_SP"
            else:
                # Follow-ups use ECHO31
                col = f"{echo_prefix}FAORRES_ECHO31_SP"
        else:
            col = column_pattern.replace("{PREFIX}", prefix)
            if echo_prefix:
                col = col.replace("{ECHO_PREFIX}", echo_prefix)
        
        # Try exact match first
        if col in row.index:
            val = row[col]
            if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                return str(val)
        
        # Try partial match (column might have extra suffix)
        for col_name in row.index:
            if str(col_name).startswith(col):
                val = row[col_name]
                if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                    return str(val)
        
        return "NA"
    
    def get_echo_value(self, row, echo_prefix, semantic_key):
        """Get Echo value by searching for semantic pattern in column labels."""
        # Semantic patterns from echo_export.py
        patterns = {
            "echo_tr_hepatic": ["systolic hepatic", "hepatic backflow", "TRSEV_HVB"],
            "echo_tr_core_lab": ["core lab", "device's valve", "TRSEV_CL", "TRSEV1"],
            "echo_co": ["cardiac output", "_CO"],
        }
        
        search_terms = patterns.get(semantic_key, [])
        
        # Search for columns with this prefix that match the semantic pattern
        for col in row.index:
            col_str = str(col)
            if not col_str.startswith(echo_prefix):
                continue
            
            # Check label if available
            label = str(self.labels_map.get(col, col)).lower()
            col_lower = col_str.lower()
            
            for term in search_terms:
                if term.lower() in label or term.lower() in col_lower:
                    val = row[col]
                    if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                        return str(val)
        
        return "NA"

    def parse_date(self, date_str):
        """Parse date string to datetime object."""
        if pd.isna(date_str) or str(date_str).strip() in ['', 'nan', 'NaN']:
            return None
        try:
            return pd.to_datetime(date_str)
        except Exception:
            return None

    def parse_frequency_multiplier(self, freq_str, freq_other_str=""):
        """Parse frequency string and return (multiplier, display_note, override_daily_dose).
        
        Returns tuple: (multiplier, note, override_daily_dose)
        - multiplier: number of doses per day (None if can't calculate)
        - note: display note (e.g., "PRN", "q6h→4x", "continuous")
        - override_daily_dose: exact daily dose if parsed from text (e.g. "6mg am, 4mg pm" -> 10)
        """
        if not freq_str or freq_str.lower() in ['nan', 'none', '']:
            return 1, "", None
        
        freq = freq_str.strip().lower()
        
        # Standard frequencies
        if freq == "once a day" or freq == "qd" or freq == "od":
            return 1, "", None
        elif freq == "twice a day" or freq == "bid":
            return 2, "", None
        elif freq == "3 times a day" or freq == "tid":
            return 3, "", None
        elif freq == "4 times a day" or freq == "qid":
            return 4, "", None
        elif freq == "every other day" or freq == "qod":
            return 0.5, "(every 48h)", None
        elif freq == "as needed":
            return None, "PRN", None
        elif freq == "once":
            return 1, "(single dose)", None
        elif freq == "other":
            # Parse the "other" field
            if freq_other_str:
                return self._parse_other_frequency(freq_other_str)
            return 1, "", None
        
        return 1, "", None
    
    def _parse_other_frequency(self, other_str):
        """Parse 'Other' frequency values like q6h, q8h, continuous, or multi-dose text."""
        other = other_str.strip().lower()
        
        # Check for explicit multiple "mg" dosages (e.g. "6 mg morning, 4 mg evening")
        # Pattern: look for numbers followed by "mg"
        mg_matches = re.findall(r'(\d+(?:\.\d+)?)\s*mg', other)
        if len(mg_matches) > 1:
            total_dose = sum(float(m) for m in mg_matches)
            return None, f"({other_str})", total_dose
            
        # Every other day check (in case it's typed in Other)
        if "every other day" in other or "qod" in other:
             return 0.5, "(every 48h)", None

        # Check for q{N}h pattern (every N hours)
        match = re.match(r'q\s*(\d+)\s*h', other)
        if match:
            interval_hours = int(match.group(1))
            if interval_hours > 0:
                doses_per_day = 24 // interval_hours
                return doses_per_day, f"(q{interval_hours}h→{doses_per_day}x/d)", None
        
        # Continuous infusion
        if "continuous" in other:
            return None, "(continuous)", None
        
        # Unknown - return 1x and show the raw value
        return 1, f"({other_str.strip()})", None
    
    def get_diuretic_rows(self, row, visits_config):
        """Get simplified Daily Loop Diuretic rows split by drug name.
        
        Returns: [ ["Diuretic Treatment (Drug)", val1, val2, ...], ... ]
        Shows FULL DAILY DOSAGE (single dose × frequency).
        """
        # Dictionary to store meds: { "Drugs": { "VisitName": "DailyDose" } }
        meds_found = {} # {"Torsemide": {"Screening": "40", ...}}
        
        # Get raw CM data
        cm_trt = str(row.get("LOGS_CM_CMTRT", ""))
        cm_dose = str(row.get("LOGS_CM_CMDOSE", ""))
        cm_freq = str(row.get("LOGS_CM_CMDOSFRQ", ""))
        cm_freq_oth = str(row.get("LOGS_CM_CMDOSFRQ_OTH", ""))
        cm_start = str(row.get("LOGS_CM_CMSTDAT", ""))
        cm_end = str(row.get("LOGS_CM_CMENDAT", ""))
        cm_ongoing = str(row.get("LOGS_CM_CMONGO", ""))
        
        trts = cm_trt.split('|')
        doses = cm_dose.split('|')
        freqs = cm_freq.split('|')
        freq_oths = cm_freq_oth.split('|')
        starts = cm_start.split('|')
        ends = cm_end.split('|')
        ongoings = cm_ongoing.split('|')
        
        # Loop Diuretics keywords
        loop_diuretics = ["torsemide", "furosemide", "bumetanide"]
        
        # Identify relevant meds indices
        relevant_indices = []
        for i, trt in enumerate(trts):
            trt_lower = trt.lower()
            for ld in loop_diuretics:
                if ld in trt_lower:
                    relevant_indices.append((i, ld.capitalize()))
                    break
        
        if not relevant_indices:
            # If no diuretics found, return a generic empty row
            return [["Diuretic Treatment (Daily Dose)"] + ["NA"] * len(visits_config)]
        
        
        # Dictionary to store meds data: { "DrugName": { "VisitName": { "sum": 0.0, "active": False } } }
        meds_data = {} 
        
        # Get list of visit names for initializing/ordering
        visit_names = [v_name for v_name, _ in visits_config]
        
        # Process each relevant med
        for idx, drug_name in relevant_indices:
            if drug_name not in meds_data:
                meds_data[drug_name] = {v: {"sum": 0.0, "active": False} for v in visit_names}
            
            # Get dates for this script
            start_date = self.parse_date(starts[idx]) if idx < len(starts) else None
            end_date = self.parse_date(ends[idx]) if idx < len(ends) else None
            is_ongoing = (ongoings[idx].lower() == 'checked') if idx < len(ongoings) else False
            
            # Get single dose and frequency
            single_dose_str = doses[idx].strip() if idx < len(doses) else ""
            freq_str = freqs[idx].strip() if idx < len(freqs) else ""
            freq_oth_str = freq_oths[idx].strip() if idx < len(freq_oths) else ""
            
            # Parse frequency and calculate daily dose
            multiplier, freq_note, override_dose = self.parse_frequency_multiplier(freq_str, freq_oth_str)
            
            # Calculate daily dose (numeric)
            daily_dose_num = 0.0
            try:
                single_dose = float(single_dose_str) if single_dose_str else None
                
                # Check for override (e.g. from explicit multi-dose string)
                if override_dose is not None:
                    daily_dose_num = float(override_dose)
                
                # Otherwise calculate using multiplier
                elif single_dose is not None and multiplier is not None:
                    daily_dose_num = single_dose * multiplier
                
                # If we have single dose but no multiplier/override, we can't add to sum reliably
            except (ValueError, TypeError):
                daily_dose_num = 0.0
            
            # Check overlap with each visit
            for visit_name, visit_conf in visits_config:
                visit_date_col = visit_conf.get("date_col")
                visit_date_val = row.get(visit_date_col)
                visit_date = self.parse_date(visit_date_val)
                
                if not visit_date:
                    continue
                
                # Check overlap: Start <= Visit <= End (or Ongoing)
                match = False
                ends_on_visit = False  # Track if this prescription ENDS on visit day
                starts_on_visit = False  # Track if this prescription STARTS on visit day
                
                if start_date and visit_date >= start_date:
                    if start_date == visit_date:
                        starts_on_visit = True
                    
                    if is_ongoing:
                        match = True
                    elif end_date and visit_date <= end_date:
                        match = True
                        if end_date == visit_date and not is_ongoing:
                            ends_on_visit = True
                    elif not end_date:
                         # Assume ongoing/single if no end date
                         match = True

                if match:
                    # Check if we should exclude this prescription due to same-day change
                    # Only exclude if: (1) ends on visit day, (2) doesn't start on visit day, 
                    # (3) another prescription for same drug starts on visit day
                    should_exclude = False
                    
                    if ends_on_visit and not starts_on_visit:
                        # Check if another prescription for this drug starts on visit day
                        for other_idx, other_drug in relevant_indices:
                            if other_drug == drug_name and other_idx != idx:
                                other_start = self.parse_date(starts[other_idx]) if other_idx < len(starts) else None
                                if other_start and other_start == visit_date:
                                    should_exclude = True
                                    break
                    
                    if not should_exclude:
                        # Found active med for this visit
                        meds_data[drug_name][visit_name]["active"] = True
                        meds_data[drug_name][visit_name]["sum"] += daily_dose_num
        
        # Generate result rows
        result_rows = []
        for drug in sorted(meds_data.keys()):
            values = []
            row_has_data = False
            for visit_name in visit_names:
                data = meds_data[drug][visit_name]
                if data["active"]:
                    val = data["sum"]
                    # Format nicely
                    if val > 0:
                        if val == int(val):
                            display_val = str(int(val))
                        else:
                            display_val = f"{val:.1f}"
                        values.append(f"{display_val} [mg/day]")
                    else:
                        values.append("Unknown")
                    row_has_data = True
                else:
                    values.append("NA")
            
            if row_has_data:
                result_rows.append([f"Diuretic Treatment ({drug}) [mg/day]"] + values)
                
        if not result_rows:
            return [["Diuretic Treatment (Daily Dose)"] + ["NA"] * len(visits_config)]
            
        return result_rows

    def generate_diuretic_timeline(self, patient_id, include_prn=False):
        """Generate a matplotlib figure showing diuretic dosage over time.
        
        Returns: matplotlib Figure object, or None if no data
        """
        rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if rows.empty:
            return None
        
        row = rows.iloc[0]
        
        # Get raw CM data
        cm_trt = str(row.get("LOGS_CM_CMTRT", ""))
        cm_dose = str(row.get("LOGS_CM_CMDOSE", ""))
        cm_freq = str(row.get("LOGS_CM_CMDOSFRQ", ""))
        cm_freq_oth = str(row.get("LOGS_CM_CMDOSFRQ_OTH", ""))
        cm_start = str(row.get("LOGS_CM_CMSTDAT", ""))
        cm_end = str(row.get("LOGS_CM_CMENDAT", ""))
        cm_ongoing = str(row.get("LOGS_CM_CMONGO", ""))
        cm_indication = str(row.get("LOGS_CM_CMINDI", ""))
        cm_ae_ref = str(row.get("LOGS_CM_CMREF_AE", ""))
        
        trts = cm_trt.split('|')
        doses = cm_dose.split('|')
        freqs = cm_freq.split('|')
        freq_oths = cm_freq_oth.split('|')
        starts = cm_start.split('|')
        ends = cm_end.split('|')
        ongoings = cm_ongoing.split('|')
        indications = cm_indication.split('|')
        ae_refs = cm_ae_ref.split('|')
        
        # Loop Diuretics keywords
        loop_diuretics = ["torsemide", "furosemide", "bumetanide"]
        
        # Collect prescription data
        prescriptions = []
        for i, trt in enumerate(trts):
            trt_clean = trt.strip()
            trt_lower = trt_clean.lower()
            drug_name = None
            
            if not trt_clean:
                continue
            
            # Fuzzy matching for loop diuretics
            matches = difflib.get_close_matches(trt_lower, loop_diuretics, n=1, cutoff=0.7)
            if matches:
                drug_name = matches[0].capitalize()
            # If no fuzzy match, check for substring
            elif any(ld in trt_lower for ld in loop_diuretics):
                 for ld in loop_diuretics:
                    if ld in trt_lower:
                        drug_name = ld.capitalize()
                        break
            
            if not drug_name:
                continue
            
            start_date = self.parse_date(starts[i]) if i < len(starts) else None
            end_date = self.parse_date(ends[i]) if i < len(ends) else None
            is_ongoing = (ongoings[i].lower() == 'checked') if i < len(ongoings) else False
            
            # If no start date, try to use screening date as fallback for pre-treatment meds
            if not start_date:
                screening_date_val = row.get("SBV_SV_SVSTDTC")
                start_date = self.parse_date(screening_date_val)
                if not start_date:
                    continue  # Skip if still no valid date
            
            # Calculate daily dose
            single_dose_str = doses[i].strip() if i < len(doses) else ""
            freq_str = freqs[i].strip() if i < len(freqs) else ""
            freq_oth_str = freq_oths[i].strip() if i < len(freq_oths) else ""
            multiplier, freq_note, override_dose = self.parse_frequency_multiplier(freq_str, freq_oth_str)
                
            # Check for PRN
            is_prn = False
            daily_dose = 0.0 # Initialize daily_dose
            if multiplier is None and (freq_note == "PRN" or freq_str.lower() == "as needed"):
                if include_prn:
                    is_prn = True
                    # For PRN, use single dose as visible height proxy (or 0 if missing)
                    single_dose = float(single_dose_str) if single_dose_str else None
                    daily_dose = float(single_dose_str) if single_dose_str else 0.0
                    # Ensure we don't accidentally multiply by None
                    multiplier = 0 
            else: 
                 # Standard calculation
                try:
                    single_dose = float(single_dose_str) if single_dose_str else None
                    if override_dose is not None:
                        daily_dose = float(override_dose)
                    elif single_dose is not None and multiplier is not None:
                        daily_dose = single_dose * multiplier
                except (ValueError, TypeError):
                    daily_dose = 0.0 # Ensure daily_dose is defined in case of exception
            
            if daily_dose <= 0 and not is_prn:
                continue
            
            # If ongoing, use today as end date for display
            if is_ongoing or not end_date:
                end_date = pd.Timestamp.today().normalize()
            
            # Get AE reference
            # Use specific AE reference for this row if available
            ae_ref = ""
            if i < len(ae_refs):
                current_ref = ae_refs[i].strip()
                if current_ref and current_ref.lower() not in ['nan', 'none', '']:
                    # Reformat like table: "#3 / Text" -> "#3. Text"
                    parts = current_ref.split(' / ', 1)
                    if len(parts) == 2:
                        ae_ref = f"{parts[0].replace('#', '')}. {parts[1]}"
                    else:
                        ae_ref = current_ref
            
            # Fallback check on indication if no explicit reference but indication mentions AE
            if not ae_ref and i < len(indications):
                ind = indications[i].strip().lower()
                ind = indications[i].strip().lower()
                if "adverse" in ind or ind == "ae":
                     ae_ref = "AE (No ref)"
            
            prescriptions.append({
                'drug': drug_name,
                'start': start_date,
                'end': end_date,
                'daily_dose': daily_dose,
                'is_prn': is_prn,
                'single_dose': single_dose,
                'multiplier': multiplier,
                'freq_str': freq_str,
                'freq_oth_str': freq_oth_str,
                'freq_note': freq_note,
                'ae_ref': ae_ref
            })
        
        if not prescriptions:
            return None
        
        # Get visit dates
        visit_dates = []
        for visit_key, visit_conf in FU_VISITS.items():
            date_col = visit_conf.get("date_col")
            date_val = row.get(date_col)
            parsed = self.parse_date(date_val)
            if parsed:
                visit_dates.append((visit_conf.get("label", visit_key), parsed))
        
        # Get Treatment visit date (separate from FU_VISITS)
        treatment_date = None
        treatment_date_val = row.get("TV_PR_SVDTC")
        if treatment_date_val:
            treatment_date = self.parse_date(treatment_date_val)
        
        # Sort prescriptions by start date
        prescriptions.sort(key=lambda x: x['start'])
        
        if not prescriptions:
            return None
            
        # --------------------------------------------------------------------------------
        # SEGMENT CALCULATION LOGIC (STACKING)
        # --------------------------------------------------------------------------------
        
        # Group by drug
        drugs_map = {}
        ae_markers = []
        for rx in prescriptions:
            d = rx['drug']
            if d not in drugs_map:
                drugs_map[d] = []
            drugs_map[d].append(rx)
            
            # Also collect AE markers here from raw prescriptions
            if rx.get('ae_ref'):
                ae_markers.append((rx['start'], rx['ae_ref'], d))

        all_segments = []
        
        for drug, rxs in drugs_map.items():
            # Collect all boundaries
            boundaries = set()
            for rx in rxs:
                boundaries.add(rx['start'])
                # Use end + 1 day for updated boundary (exclusive end) because bar plotting width is defined by interval
                boundaries.add(rx['end'] + timedelta(days=1))
            
            sorted_dates = sorted(list(boundaries))
            
            # Create segments
            for i in range(len(sorted_dates) - 1):
                seg_start = sorted_dates[i]
                seg_end = sorted_dates[i+1] # Exclusive
                
                # Find active prescriptions in this segment
                # Segment is [seg_start, seg_end)
                # RX is [rx_start, rx_end] (inclusive) -> [rx_start, rx_end + 1)
                
                active_rxs = []
                for rx in rxs:
                    rx_exclusive_end = rx['end'] + timedelta(days=1)
                    
                    # Check for overlap: max(starts) < min(ends)
                    latest_start = max(rx['start'], seg_start)
                    earliest_end = min(rx_exclusive_end, seg_end)
                    
                    if latest_start < earliest_end:
                        active_rxs.append(rx)
                
                if not active_rxs:
                    continue
                    
                # Calculate total dose
                total_dose = sum(rx['daily_dose'] for rx in active_rxs)
                
                # Generate label matching user request
                # "4 mg: 3mg everyday + 1 mg of qod"
                
                details = []
                has_prn = False
                all_prn = True
                
                for rx in active_rxs:
                    daily = rx['daily_dose']
                    mult = rx.get('multiplier')
                    is_rx_prn = rx.get('is_prn', False)
                    
                    if is_rx_prn:
                        has_prn = True
                        details.append(f"{daily:.1f} mg (PRN)")
                    else:
                        all_prn = False
                        if mult == 0.5:
                            details.append(f"{daily:.1f} mg (qod)")
                        else:
                            details.append(f"{daily:.1f} mg")
                
                if len(active_rxs) > 1:
                    detail_str = " + ".join(details)
                    final_label = f"{total_dose:.1f} mg: {detail_str}"
                else:
                     # Single RX
                     rx = active_rxs[0]
                     mult = rx.get('multiplier')
                     single = rx.get('single_dose')
                     is_rx_prn = rx.get('is_prn', False)
                     
                     if is_rx_prn:
                         final_label = f"{total_dose:.1f} mg\n(PRN)"
                         has_prn = True
                     elif mult == 0.5 and single:
                         # Format single dose for 0.5 multiplier
                         if single == int(single):
                             single_text = f"{int(single)}"
                         else:
                             single_text = f"{single:.1f}"
                         final_label = f"{single_text}\n(qod)"
                     else:
                         if total_dose == int(total_dose):
                            final_label = f"{int(total_dose)} mg/day"
                         else:
                            final_label = f"{total_dose:.1f} mg/day"

                all_segments.append({
                    'drug': drug,
                    'start': seg_start,
                    'end': seg_end, 
                    'daily_dose': total_dose,
                    'label': final_label,
                    'is_prn': has_prn,     # Flag if any PRN
                    'all_prn': all_prn,    # Flag if ALL PRN (for styling)
                    'ae_refs': [rx['ae_ref'] for rx in active_rxs if rx.get('ae_ref')]
                })

        # Calculate dynamic figure height based on SEGMENT max dose
        if all_segments:
            max_daily_dose = max((s['daily_dose'] for s in all_segments), default=10)
        else:
            max_daily_dose = 10
            
        # Define y_max here so it's available for label staggering
        y_max = max(max_daily_dose * 1.3, 20)
        
        # Base height 
        fig_height = 8 
        fig, ax = plt.subplots(figsize=(14, fig_height))
        
        # Colors for different drugs
        drug_colors = {
            'Torsemide': '#3498db',
            'Furosemide': '#e74c3c',
            'Bumetanide': '#2ecc71'
        }
        
        # Track label positions for staggering (prevent overlap)
        label_positions = []  # List of (x, y) positions already used
        drug_bars = {}

        # Debug logging loop removed for brevity, assuming standard logs OK
        # Or I can log segments to debug file?
        
        # Draw bars for each SEGMENT
        for seg in all_segments:
            drug = seg['drug']
            color = drug_colors.get(drug, '#95a5a6')
            
            daily_dose = seg['daily_dose']
            dose_label = seg['label']
            
            start_num = mdates.date2num(seg['start'])
            # Note: end is exclusive, but matplotlib width is simple diff
            end_num = mdates.date2num(seg['end'])
            width = end_num - start_num
            
            # Draw bar aligned to bottom (y=0)
            try:
                # Check for AE association
                is_ae_segment = any(ref and ref.strip() for ref in seg.get('ae_refs', []))
                is_all_prn = seg.get('all_prn', False)
                
                bar_kwargs = {
                    'width': width,
                    'align': 'edge',
                    'color': color,
                    'alpha': 0.7,
                    'edgecolor': 'black',
                    'linewidth': 0.5
                }
                
                # Apply highlight for AE segments
                if is_ae_segment:
                    bar_kwargs['hatch'] = '///'
                    bar_kwargs['edgecolor'] = 'red'
                    bar_kwargs['linewidth'] = 1.0
                    bar_kwargs['alpha'] = 0.8
                
                # Apply style for PRN segments
                if is_all_prn:
                    bar_kwargs['linestyle'] = '--' # Dashed border for PRN
                    bar_kwargs['alpha'] = 0.5      # Lighter transparency
                    if not is_ae_segment:          # Keep red border if AE
                        bar_kwargs['edgecolor'] = 'gray'
                        bar_kwargs['linewidth'] = 1.0
                
                bar = ax.bar(start_num, daily_dose, **bar_kwargs)
            except Exception as e:
                logger.error("Error plotting bar for %s at %s: %s", drug, start_num, e)
                continue
            
            # Position label above the bar, centered
            label_x = start_num + width / 2
            label_y = daily_dose
            
            # Stagger labels to prevent overlap
            try:
                for prev_x, prev_y in label_positions:
                    if abs(label_x - prev_x) < 0.5: 
                        label_y = max(label_y, prev_y + 0.1 * y_max)
            except Exception:
                pass # safe fallback
            
            label_positions.append((label_x, label_y))
            
            try:
                ax.text(label_x, label_y, dose_label, 
                       ha='center', va='bottom', fontsize=8, fontweight='bold')
            except Exception as e:
                logger.error("Error plotting label: %s", e)
            
            if drug not in drug_bars:
                drug_bars[drug] = color

        # Set Y-axis limit with headroom for headers
        y_max = max(max_daily_dose * 1.3, 20) # Ensure at least 20 height
        ax.set_ylim(0, y_max)
        
        # Header position (fixed at top)
        header_y_pos = y_max * 0.95
        
        # Add visit date markers (red) - position labels AT THE TOP
        all_visit_markers = []
        for label, vdate in visit_dates:
            try:
                x_pos = mdates.date2num(vdate)
                ax.axvline(x=x_pos, color='red', linestyle='--', alpha=0.5, linewidth=1)
                # Position label at top
                ax.text(x_pos, header_y_pos, label, rotation=90, ha='right', va='top',
                       fontsize=8, color='red')
                all_visit_markers.append((x_pos, vdate, 'red'))
            except Exception as e:
                logger.error("Error plotting visit marker %s: %s", label, e)
        
        # Add Treatment visit marker (green)
        if treatment_date:
            try:
                x_pos = mdates.date2num(treatment_date)
                ax.axvline(x=x_pos, color='green', linestyle='-', alpha=0.8, linewidth=2)
                ax.text(x_pos, header_y_pos, "Treatment", rotation=90, ha='right', va='top',
                       fontsize=9, color='green', fontweight='bold')
                all_visit_markers.append((x_pos, treatment_date, 'green'))
            except Exception as e:
                logger.error("Error plotting treatment marker: %s", e)
            
        # Add AE markers
        for ae_date, ae_ref, drug in ae_markers:
            try:
                x_pos = mdates.date2num(ae_date)
                # Plot line
                ax.axvline(x=x_pos, color='blue', linestyle=':', alpha=0.9, linewidth=1.5)
                
                ae_short = ae_ref[:15] + "..." if len(ae_ref) > 15 else ae_ref
                # Plot label
                ax.text(x_pos, header_y_pos * 0.8, f"AE: {ae_short}", rotation=90, ha='right', va='top',
                       fontsize=7, color='blue')
            except Exception as e:
                logger.error("Error plotting AE marker: %s", e)
        
        # Format x-axis as dates (regular monthly ticks)
        ax.xaxis.set_major_formatter(mdates.DateFormatter('%Y-%m'))
        ax.xaxis.set_major_locator(mdates.MonthLocator())
        plt.xticks(rotation=45, ha='right')
        
        # Add visit dates on the x-axis in a different color (blue)
        # Rotate at 45 degrees like regular axis labels to prevent overlap
        if all_visit_markers:
            # Sort markers by x position
            all_visit_markers.sort(key=lambda x: x[0])
            
            # Add text annotations at the bottom for each visit date (rotated like other labels)
            ax_bottom = ax.get_ylim()[0]
            for x_pos, vdate, color in all_visit_markers:
                date_str = vdate.strftime('%Y-%m-%d') if hasattr(vdate, 'strftime') else str(vdate)[:10]
                # Add date label below x-axis in blue, rotated 45 degrees like regular labels
                ax.annotate(date_str, xy=(x_pos, ax_bottom), xytext=(0, -20),
                           textcoords='offset points', ha='right', va='top',
                           fontsize=6, color='blue', fontweight='bold',
                           rotation=45, annotation_clip=False)
        
        # Set labels
        ax.set_xlabel('Date (YYYY-MM)')
        ax.set_ylabel('Daily Dose (mg)')
        ax.set_title(f'Loop Diuretic Dosage History - Patient {patient_id}')
        
        # Restore y-axis ticks (meaningful now)
        # ax.set_yticks([]) # Removed this line
        ax.grid(axis='y', linestyle='--', alpha=0.3)
        
        # Add legend
        legend_patches = [Patch(color=color, label=drug, alpha=0.7) 
                         for drug, color in drug_bars.items()]
        ax.legend(handles=legend_patches, loc='upper left') # Moved to top left to avoid conflict
        
        # Adjust layout
        plt.tight_layout()
        
        return fig

    def get_6mwt_value(self, row, prefix):
        """Get 6MWT distance - handles different column patterns for SBV vs FU visits."""
        # SBV uses: SBV_6MWT_FTORRES_DIS
        # FU visits use: FU1M_6MWT_FU_FTORRES_DIS
        
        # Try FU pattern first (for follow-up visits)
        col = f"{prefix}6MWT_FU_FTORRES_DIS"
        if col in row.index:
            val = row[col]
            if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                return str(val)
        
        # Try SBV pattern (for screening)
        col = f"{prefix}6MWT_FTORRES_DIS"
        if col in row.index:
            val = row[col]
            if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                return str(val)
        
        return "NA"
    
    def generate_highlights_table(self, patient_id, target_visit):
        """Generate highlights table for a specific FU visit.
        
        Returns tuple: (highlights_table_str, vitals_table_str)
        """
        rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if rows.empty:
            return None, None
        row = rows.iloc[0]
        
        # Get visit config
        visit_config = FU_VISITS.get(target_visit)
        if not visit_config:
            return None, None
        
        # Get all visits up to and including the target visit
        # Start from Screening, include all visits in order until target
        target_idx = VISIT_ORDER.index(target_visit) if target_visit in VISIT_ORDER else -1
        if target_idx < 0:
            return None, None
        
        visits_to_process = VISIT_ORDER[:target_idx + 1]  # Screening through target
        columns = [FU_VISITS[v]["label"] for v in visits_to_process]
        
        # Build highlights table
        highlights_rows = []
        for param_name, pattern, value_type in FU_HIGHLIGHT_PARAMS:
            
            # Special handling: Insert Diuretics dynamic rows before CO
            if "CO (based on Echo)" in param_name:
                # Prepare config list for get_diuretic_rows: [("Screening", config), ...]
                med_visits_config = []
                for v in visits_to_process:
                    med_visits_config.append((v, FU_VISITS[v]))
                
                # Get dynamic rows
                diuretic_rows = self.get_diuretic_rows(row, med_visits_config)
                highlights_rows.extend(diuretic_rows)

            row_data = [param_name]
            
            for visit in visits_to_process:
                v_config = FU_VISITS.get(visit)
                if not v_config:
                    row_data.append("NA")
                    continue
                
                prefix = v_config["prefix"]
                echo_prefix = v_config.get("echo_prefix", "")
                
                # Get value
                val = self.get_value(row, pattern, prefix, echo_prefix)
                
                # Simplify values
                if "NYHA" in param_name:
                    val = self.simplify_nyha(val)
                elif "CFS" in param_name:
                    val = self.simplify_cfs(val)
                elif "Edema" in param_name or "Ascites" in param_name:
                    val = self.simplify_grade(val)
                elif "6MWT" in param_name:
                    val = self.get_6mwt_value(row, prefix)
                
                row_data.append(val)
            
            highlights_rows.append(row_data)
        
        # Build vitals table - now with columns per visit like highlights
        vitals_rows = []
        for param_name, pattern, _ in VITAL_SIGNS_PARAMS:
            row_data = [param_name]
            
            for visit in visits_to_process:
                v_config = FU_VISITS.get(visit)
                if not v_config:
                    row_data.append("NA")
                    continue
                
                prefix = v_config["prefix"]
                val = self.get_value(row, pattern, prefix)
                row_data.append(val)
            
            vitals_rows.append(row_data)
        
        # Create DataFrames
        highlights_headers = ["Parameter"] + columns
        df_highlights = pd.DataFrame(highlights_rows, columns=highlights_headers)
        
        vitals_headers = ["Parameter"] + columns  # Same column headers as highlights
        df_vitals = pd.DataFrame(vitals_rows, columns=vitals_headers)
        
        # Get full diuretic history
        diuretic_history = self.get_diuretic_history(row)
        df_diuretic_history = pd.DataFrame(diuretic_history, columns=["Drug", "Single Dose", "Frequency", "Daily Dose", "Start Date", "End Date", "Indication"])
        
        return df_highlights, df_vitals, df_diuretic_history
    
    def get_diuretic_history(self, row):
        """Get full list of loop diuretic events sorted by date.
           Uses fuzzy matching to detect diuretic names with typos."""
        cm_trt = str(row.get("LOGS_CM_CMTRT", ""))
        cm_dose = str(row.get("LOGS_CM_CMDOSE", ""))
        cm_freq = str(row.get("LOGS_CM_CMDOSFRQ", ""))
        cm_freq_oth = str(row.get("LOGS_CM_CMDOSFRQ_OTH", ""))
        cm_start = str(row.get("LOGS_CM_CMSTDAT", ""))
        cm_end = str(row.get("LOGS_CM_CMENDAT", ""))
        cm_indc = str(row.get("LOGS_CM_CMINDC", ""))
        cm_ref_mh = str(row.get("LOGS_CM_CMREF_MH", ""))  # e.g. "#1 / Heart Failure"
        cm_ref_ae = str(row.get("LOGS_CM_CMREF_AE", ""))  # e.g. "#3 / Acute kidney injury"
        
        trts = cm_trt.split('|')
        doses = cm_dose.split('|')
        freqs = cm_freq.split('|')
        freq_oths = cm_freq_oth.split('|')
        starts = cm_start.split('|')
        ends = cm_end.split('|')
        indcs = cm_indc.split('|')
        ref_mhs = cm_ref_mh.split('|')
        ref_aes = cm_ref_ae.split('|')
        
        loop_diuretics = ["torsemide", "furosemide", "bumetanide"]
        events = []
        
        for i, trt in enumerate(trts):
            trt_clean = trt.strip()
            trt_lower = trt_clean.lower()
            if not trt_clean:
                continue
                
            # Identify drug (Fuzzy Matching)
            drug_name = trt_clean
            is_loop = False
            
            # Try fuzzy match first
            matches = difflib.get_close_matches(trt_lower, loop_diuretics, n=1, cutoff=0.7)
            if matches:
                matched_name = matches[0].capitalize()
                # If we have a callback, ask for confirmation
                if self.fuzzy_confirm_callback:
                    # Only ask if not an exact match (ignoring case)
                    if trt_lower == matched_name.lower():
                        is_loop = True
                        drug_name = matched_name
                    elif self.fuzzy_confirm_callback(drug_name, matched_name):
                        is_loop = True
                        drug_name = matched_name
                    else:
                        is_loop = False # Rejected by user
                else:
                    # No callback, auto-accept (default behavior)
                    is_loop = True
                    drug_name = matched_name
            # Fallback check for exact substrings if fuzzy match fails (for compound names)
            elif any(ld in trt_lower for ld in loop_diuretics):
                for ld in loop_diuretics:
                    if ld in trt_lower:
                        is_loop = True
                        drug_name = ld.capitalize()
                        break
            
            if is_loop:
                single_dose = doses[i] if i < len(doses) else "Unknown"
                freq_str = freqs[i] if i < len(freqs) else ""
                freq_oth_str = freq_oths[i] if i < len(freq_oths) else ""
                start = starts[i] if i < len(starts) else ""
                end = ends[i] if i < len(ends) else ""
                indc = indcs[i] if i < len(indcs) else ""
                
                # Careful with references - they might be sparse or aligned differently? 
                # Assuming pipe-alignment matches treatments
                ref_mh = ref_mhs[i].strip() if i < len(ref_mhs) else ""
                ref_ae = ref_aes[i].strip() if i < len(ref_aes) else ""
                
                # Build rich indication display
                # Priority: Use MH/AE reference with term if available
                if ref_mh and ref_mh.lower() not in ['nan', 'none', '']:
                    # Format: "#1 / Chronic Heart Failure" -> "MH#1. Chronic Heart Failure"
                    parts = ref_mh.split(' / ', 1)
                    if len(parts) == 2:
                        mh_num = parts[0].replace('#', '').strip()
                        mh_term = parts[1].strip()
                        indc_display = f"MH#{mh_num}. {mh_term}"
                    else:
                        indc_display = f"MH: {ref_mh}"
                elif ref_ae and ref_ae.lower() not in ['nan', 'none', '']:
                    # Format: "#3 / Acute kidney injury" -> "AE#3. Acute kidney injury"
                    parts = ref_ae.split(' / ', 1)
                    if len(parts) == 2:
                        ae_num = parts[0].replace('#', '').strip()
                        ae_term = parts[1].strip()
                        indc_display = f"AE#{ae_num}. {ae_term}"
                    else:
                        indc_display = f"AE: {ref_ae}"
                elif indc and indc.lower() not in ['nan', 'none', '']:
                    # Fallback to raw indication (e.g., "Prophylaxis")
                    indc_display = indc
                else:
                    indc_display = ""
                
                # Get frequency display and calculate daily dose
                multiplier, freq_note, override_dose = self.parse_frequency_multiplier(freq_str, freq_oth_str)
                
                # Format frequency for display
                if freq_str and freq_str.lower() != 'nan':
                    if freq_str.lower() == "other" and freq_oth_str:
                        freq_display = freq_oth_str
                    else:
                        freq_display = freq_str
                else:
                    freq_display = ""
                
                # Calculate daily dose
                try:
                    dose_num = float(single_dose) if single_dose else None
                    if override_dose is not None:
                         # Explicit override (e.g. from multi-dose text)
                         daily_dose = str(int(override_dose)) if override_dose == int(override_dose) else f"{override_dose:.1f}"
                    elif dose_num is not None and multiplier is not None:
                        daily = dose_num * multiplier
                        daily_dose = str(int(daily)) if daily == int(daily) else f"{daily:.1f}"
                    elif dose_num is not None:
                        # Can't calculate, show single dose
                        daily_dose = f"{single_dose} {freq_note}" if freq_note else single_dose
                    else:
                        daily_dose = "Unknown"
                except (ValueError, TypeError):
                    daily_dose = single_dose
                
                # Format dates for display
                start_dt = self.parse_date(start)
                end_dt = self.parse_date(end)
                start_str = start_dt.strftime('%Y-%m-%d') if start_dt else start
                end_str = end_dt.strftime('%Y-%m-%d') if end_dt else end
                
                events.append({
                    "Drug": drug_name,
                    "Single Dose": single_dose,
                    "Frequency": freq_display,
                    "Daily Dose": daily_dose,
                    "Start Date": start_str,
                    "End Date": end_str,
                    "Indication": indc_display,
                    "_dt": start_dt or datetime.datetime.min  # For sorting
                })
        
        # Sort by start date
        events.sort(key=lambda x: x["_dt"])
        
        # Convert to list of lists for DataFrame
        result = []
        for e in events:
            result.append([e["Drug"], e["Single Dose"], e["Frequency"], e["Daily Dose"], 
                          e["Start Date"], e["End Date"], e["Indication"]])
            
        return result
    
    def get_available_fu_visits(self, patient_id):
        """Get list of FU visits that have data for this patient."""
        rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if rows.empty:
            return []
        row = rows.iloc[0]
        
        available = []
        for visit_name, config in FU_VISITS.items():
            if visit_name in ["Screening", "Discharge"]:
                continue  # Skip baseline visits
            
            # Check if visit has any data (check visit date or any field)
            prefix = config["prefix"]
            date_col = f"{prefix}SV_SVSTDTC"
            if date_col in row.index:
                val = row[date_col]
                if pd.notna(val) and str(val).strip() not in ['', 'nan', 'NaN']:
                    available.append(visit_name)
        
        return available
