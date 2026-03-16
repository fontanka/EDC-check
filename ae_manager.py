import logging
import pandas as pd
import re
from datetime import datetime
from typing import List, Dict, Tuple, Optional

logger = logging.getLogger("ClinicalViewer.AEManager")

class AEManager:
    """
    Manages Adverse Event (AE) data parsing, filtering, and statistics.
    """
    def __init__(self, df_main: pd.DataFrame, df_ae: pd.DataFrame):
        self.df_main = df_main
        self.df_ae = df_ae
        
        # Column Mappings (aligned with clinical_viewer1.py)
        self.col_mapping = {
            'AE #': ['Template number', 'AE #', 'AE Number', 'AESEQ', 'LOGS_AE_AESEQ', 'LOGS_AE_AE #'],
            'SAE?': ['LOGS_AE_AESER', 'Is the event SAE?', 'AESER', 'SAE'],
            'AE Term': ['LOGS_AE_AETERM', 'adverse event / term', 'AETERM', 'Term'],
            'Severity': ['LOGS_AE_AESEV', 'Severity', 'AESEV'],
            'Interval': ['LOGS_AE_AEINT', 'Interval', 'AEINT'],
            'Onset Date': ['LOGS_AE_AESTDTC', 'Date of event onset', 'AESTDTC', 'Start Date'],
            'Resolution Date': ['LOGS_AE_AEENDTC', 'Date resolved', 'AEENDTC', 'End Date'],
            'Ongoing': ['LOGS_AE_AEONGO', 'Ongoing', 'AEONGO'],
            'Outcome': ['LOGS_AE_AEOUT', 'Outcome', 'AEOUT'],
            'Rel. PKG Trillium': ['LOGS_AE_AEREL1', 'relationship / PKG Trillium', 'AEREL1', 'Rel Trillium'],
            'Rel. Delivery System': ['LOGS_AE_AEREL2', 'relationship / PKG Delivery System', 'AEREL2', 'Rel Delivery'],
            'Rel. Handle': ['LOGS_AE_AEREL3', 'relationship / PKG Handle', 'AEREL3', 'Rel Handle'],
            'Rel. Index Procedure': ['LOGS_AE_AEREL4', 'relationship / index procedure', 'AEREL4', 'Rel Procedure'],
            'AE Description': ['LOGS_AE_AETERM_COMM', 'AE and sequelae / description', 'AETERM_COMM'],
            'SAE Description': ['LOGS_AE_AETERM_COMM1', 'SAE and sequelae / description', 'AETERM_COMM1'],
            'Hospitalization': ['LOGS_AE_AESHOSP', 'Hospitalization', 'AESHOSP'],
            'Life Threatening': ['LOGS_AE_AESLIFE', 'Life Threatening', 'AESLIFE'],
            'Death': ['LOGS_AE_AESDTH', 'Death', 'AESDTH'],
            'Disability': ['LOGS_AE_AESDISAB', 'Disability', 'AESDISAB'],
            'Other Medical Event': ['LOGS_AE_AESMIE', 'Other', 'AESMIE'],
            'AE Report Date': ['LOGS_AE_AEREPDAT', 'AE Report Date', 'AEREPDAT'],
            # Death form fields (from LOGS prefix)
            'Death Date': ['LOGS_DTH_DDDTC', 'Death Date', 'DDDTC'],
            'Death Category': ['LOGS_DTH_DDRESCAT', 'Mortality classification', 'DDRESCAT'],
            'Death Reason': ['LOGS_DTH_DDORRES', 'Reason of death', 'DDORRES'],
        }
        
        # Cache for procedure dates
        self._procedure_dates = {}
        # Cache for screen failure patient IDs
        self._screen_failures = None

    def get_screen_failures(self) -> List[str]:
        """Return list of patient IDs who are screen failures."""
        if self._screen_failures is not None:
            return self._screen_failures
        if self.df_main is None or 'Status' not in self.df_main.columns:
            self._screen_failures = []
            return self._screen_failures
        statuses = self.df_main['Status'].astype(str).str.strip().str.lower()
        mask = statuses.str.contains('screen', na=False) & statuses.str.contains('fail', na=False)
        self._screen_failures = self.df_main.loc[mask, 'Screening #'].astype(str).str.strip().tolist()
        return self._screen_failures

    def get_patient_ae_data(self, patient_id: str, filters: Dict = None) -> List[Dict]:
        """
        Get structured AE lists for a specific patient, applying filters.
        
        filters: {
            'sae_only': bool,
            'exclude_pre_proc': bool,
            'device_related_only': bool  # Implies SAE Only + Relationship check
        }
        """
        if self.df_ae is None or self.df_ae.empty:
            return []
            
        filters = filters or {}
        
        # Filter raw DF for patient (exact match to avoid substring leakage)
        pat_aes = self.df_ae[self.df_ae['Screening #'].astype(str).str.strip() == patient_id.strip()].copy()
        
        if pat_aes.empty:
            return []
            
        # Resolve column names
        available_cols = {}
        for display_name, possible_names in self.col_mapping.items():
            for pn in possible_names:
                if pn in pat_aes.columns:
                    available_cols[display_name] = pn
                    break
        
        # Get procedure date if needed
        proc_date = None
        if filters.get('exclude_pre_proc'):
            proc_date = self._get_procedure_date(patient_id)

        # Group by AE # to handle duplicates/overflow rows
        # Strategy: For each AE #, pick the row with the most populated fields
        # First, ensure we have the AE # column resolved
        ae_num_col = self._find_col('AE #')
        if ae_num_col and ae_num_col in pat_aes.columns:
            # Helper to count non-empty values
            def count_vals(row):
                return sum(1 for v in row if isinstance(v, str) and v.strip())

            # Prioritize rows with non-empty AE term, then most populated fields
            term_col_name = self._find_col('AE Term')
            pat_aes['__pop_count'] = pat_aes.apply(count_vals, axis=1)
            if term_col_name and term_col_name in pat_aes.columns:
                pat_aes['__has_term'] = pat_aes[term_col_name].apply(
                    lambda x: 0 if (str(x).strip().lower() in ('nan', '', 'none')) else 1
                )
                pat_aes = pat_aes.sort_values(['__has_term', '__pop_count'], ascending=[False, False])
                pat_aes = pat_aes.drop_duplicates(subset=[ae_num_col], keep='first')
                pat_aes = pat_aes.drop(columns=['__has_term', '__pop_count'])
            else:
                pat_aes = pat_aes.sort_values('__pop_count', ascending=False)
                pat_aes = pat_aes.drop_duplicates(subset=[ae_num_col], keep='first')
                pat_aes = pat_aes.drop(columns=['__pop_count'])
            
            # Re-sort by AE # numerically if possible
            try:
                pat_aes['__ae_sort'] = pd.to_numeric(pat_aes[ae_num_col], errors='coerce')
                pat_aes = pat_aes.sort_values('__ae_sort')
                pat_aes = pat_aes.drop(columns=['__ae_sort'])
            except (ValueError, KeyError, TypeError):
                pass

        ae_data = []
        for _, ae_row in pat_aes.iterrows():
            row_data = {}
            ongoing_value = False
            
            # Extract basic data
            for display, source in available_cols.items():
                val = ae_row.get(source, '')
                if pd.isna(val) or str(val).lower() == 'nan':
                    val = ''
                else:
                    val = str(val).strip()
                    
                    # Date cleaning
                    if 'Date' in display:
                        val = self._clean_date(val)
                    
                    # SAE Normalization
                    if display == 'SAE?':
                        val = self._normalize_boolean(val)
                        
                    # Ongoing check
                    if display == 'Ongoing':
                        ongoing_value = self._is_checked(val)
                
                row_data[display] = val
            
            # Post-processing
            if ongoing_value:
                row_data['Resolution Date'] = 'Ongoing'
                
            # --- FILTERING ---
            
            # 1. SAE Only / Device Related (Device related implies SAE only per user request)
            is_sae = row_data.get('SAE?', 'No') == 'Yes'
            if filters.get('sae_only') and not is_sae:
                continue
                
            if filters.get('device_related_only'):
                # Check relationships
                is_related = False
                rel_keys = ['Rel. PKG Trillium', 'Rel. Delivery System', 'Rel. Handle', 'Rel. Index Procedure']
                for key in rel_keys:
                    rel_val = row_data.get(key, 'Not Related')
                    if rel_val and rel_val.lower() != 'not related':
                        is_related = True
                        break
                if not is_related:
                    continue

            # 2. Pre-Procedure
            if filters.get('exclude_pre_proc') and proc_date:
                ae_date_str = row_data.get('Onset Date', '')
                ae_date = self._parse_date_obj(ae_date_str)
                if ae_date and ae_date < proc_date:
                    continue
                    
            # 3. Onset Date Cutoff
            if filters.get('onset_cutoff'):
                cutoff = self._parse_date_obj(filters['onset_cutoff'])
                if cutoff:
                    onset_str = row_data.get('Onset Date', '')
                    # Onset might be partial or empty. If empty, include? User said "started after ...". If unknown, probably exclude or include?
                    # Let's check if we can parse it.
                    onset_date = self._parse_date_obj(onset_str)
                    if not onset_date or onset_date < cutoff:
                        continue

            # 4. Report Date Cutoff
            if filters.get('report_cutoff'):
                report_cutoff = self._parse_date_obj(filters['report_cutoff'])
                if report_cutoff:
                    report_str = row_data.get('AE Report Date', '')
                    report_date = self._parse_date_obj(report_str)
                    # If report date is unknown, we can't be sure it's after cutoff, so exclude? 
                    # User: "see only events that were ADDED ... AFTER the cutoff date"
                    if not report_date or report_date < report_cutoff:
                        continue
                        
            ae_data.append(row_data)
            
        return ae_data

    def get_dataset_ae_data(self, filters: Dict = None) -> List[Dict]:
        """
        Get filtered AE data for ALL patients in the dataset.
        Useful for bulk export.
        """
        if self.df_ae is None:
            return []
            
        all_data = []
        # Get all patients
        patients = self.df_ae['Screening #'].unique().tolist()
        
        for pat_id in patients:
            pat_data = self.get_patient_ae_data(str(pat_id), filters)
            # Add Patient ID to each row for context if not already present
            for row in pat_data:
                row['Patient ID'] = str(pat_id)
                all_data.append(row)
                
        return all_data

    def get_summary_stats(self, excluded_patients: List[str] = None, exclude_pre_proc: bool = False, exclude_screen_failures: bool = False) -> Dict:
        """
        Calculate AE statistics (Total, SAE, Fatal, by Site, by Patient).
        """
        stats = {
            'total_aes': 0,
            'total_saes': 0,
            'fatal_cases': 0,
            'patients_with_aes': 0,
            'ongoing_aes': 0,
            'outcome_dist': {},
            'top_terms': {},
            'sae_criteria': {},
            'by_site': {},
            'by_patient': {},          # Dict for charts
            'per_patient_details': []  # List of strings for text view
        }
        
        if self.df_ae is None or self.df_ae.empty:
            return stats
            
        df = self.df_ae.copy()
        
        # Filter patients
        if excluded_patients:
            # Normalize to strings just in case
            excluded_patients = [str(p) for p in excluded_patients]
            df = df[~df['Screening #'].astype(str).isin(excluded_patients)]

        # Exclude screen failures
        if exclude_screen_failures:
            sf_list = self.get_screen_failures()
            if sf_list:
                df = df[~df['Screening #'].astype(str).str.strip().isin(sf_list)]

        if df.empty:
            return stats
            
        # Deduplicate overflow rows (same Patient + AE # = continuation rows for long text)
        pre_dedup_count = len(df)
        ae_num_col = self._find_col('AE #')
        term_col = self._find_col('AE Term')
        if ae_num_col and ae_num_col in df.columns:
             df['__pop_count'] = df.apply(lambda row: sum(1 for v in row if isinstance(v, str) and v.strip()), axis=1)
             # Prioritize rows with non-empty AE term, then most populated
             if term_col and term_col in df.columns:
                 df['__has_term'] = df[term_col].apply(
                     lambda x: 0 if (str(x).strip().lower() in ('nan', '', 'none')) else 1
                 )
                 df = df.sort_values(['Screening #', ae_num_col, '__has_term', '__pop_count'],
                                     ascending=[True, True, False, False])
                 df = df.drop_duplicates(subset=['Screening #', ae_num_col], keep='first')
                 df = df.drop(columns=['__has_term', '__pop_count'], errors='ignore')
             else:
                 df = df.sort_values(['Screening #', ae_num_col, '__pop_count'], ascending=[True, True, False])
                 df = df.drop_duplicates(subset=['Screening #', ae_num_col], keep='first')
                 df = df.drop(columns=['__pop_count'], errors='ignore')
        logger.debug("AE dedup: %d -> %d rows (removed %d overflow rows)",
                      pre_dedup_count, len(df), pre_dedup_count - len(df))

        # Filter Pre-Procedure (Done AFTER deduplication to ensure we check the valid row)
        if exclude_pre_proc:
            try:
                onset_col = self._find_col('Onset Date')
                if onset_col:
                    pre_filter_count = len(df)
                    unique_pats = df['Screening #'].unique()
                    proc_map = {p: self._get_procedure_date(str(p)) for p in unique_pats}
                    logger.debug("Pre-proc filter: %d patients, proc dates found: %d",
                                 len(unique_pats),
                                 sum(1 for v in proc_map.values() if v is not None))

                    def is_keep(row):
                         pid = row['Screening #']
                         proc_date = proc_map.get(pid)
                         if not proc_date: return True

                         onset_val = row.get(onset_col)
                         onset_date = self._parse_date_obj(onset_val)
                         if onset_date and onset_date < proc_date:
                             return False
                         return True

                    df = df[df.apply(is_keep, axis=1)]
                    logger.debug("Pre-proc filter: %d -> %d rows", pre_filter_count, len(df))
                else:
                    logger.warning("Pre-proc filter: Onset Date column not found")
            except Exception as e:
                logger.warning("Error in Pre-Procedure Filter: %s", e)
                import traceback
                traceback.print_exc()
        
        # Identify columns
        sae_col = self._find_col('SAE?')
        outcome_col = self._find_col('Outcome')
        term_col = self._find_col('AE Term')
        ongoing_col = self._find_col('Ongoing')
        
        # Relationship columns
        rel_device_col = self._find_col('Rel. Delivery System')
        rel_handle_col = self._find_col('Rel. Handle')
        rel_trillium_col = self._find_col('Rel. PKG Trillium')
        rel_proc_col = self._find_col('Rel. Index Procedure')
        
        # SAE Criteria columns
        hosp_col = self._find_col('Hospitalization')
        life_col = self._find_col('Life Threatening')
        death_col = self._find_col('Death')
        disab_col = self._find_col('Disability')
        other_col = self._find_col('Other Medical Event')
        
        # --- CALCULATION ---
        
        stats['total_aes'] = len(df)
        stats['patients_with_aes'] = df['Screening #'].nunique()
        
        # SAE and Outcome
        if sae_col:
            df['__is_sae'] = df[sae_col].astype(str).apply(lambda x: self._normalize_boolean(x) == 'Yes')
            stats['total_saes'] = df['__is_sae'].sum()
        else:
            df['__is_sae'] = False

        if outcome_col:
            outcomes = df[outcome_col].astype(str).str.strip().value_counts().to_dict()
            # Clean up 'nan' keys
            stats['outcome_dist'] = {k: v for k, v in outcomes.items() if k.lower() != 'nan' and k}
            
            # Fatal count from outcome
            is_fatal = df[outcome_col].astype(str).str.strip().str.lower() == 'fatal'
            stats['fatal_cases'] = is_fatal.sum()
            
        if ongoing_col:
            is_marked_ongoing = df[ongoing_col].apply(self._is_checked)
            
            # Revised Implied Ongoing Logic (Based on debugging)
            # Only count as implied if:
            # 1. End Date is EMPTY
            # 2. Outcome is NOT Fatal or Recovered
            # 3. Term is NOT Empty (valid AE)
            
            non_ongoing_mask = ~is_marked_ongoing
            
            # Helper to check emptiness
            def is_empty_date(val):
                s = str(val).lower().strip()
                return s == '' or s == 'nan' or s == 'nat' or s == 'none'
            
            # Get End Date column if available
            end_date_col = self._find_col('Resolution Date')
            
            if end_date_col:
                no_end_date = df[end_date_col].apply(is_empty_date)
            else:
                no_end_date = pd.Series([False] * len(df), index=df.index)
                
            # Check Outcome (exclude Fatal)
            if outcome_col:
                def is_not_fatal_or_recovered(val):
                    s = str(val).lower().strip()
                    return 'fatal' not in s and 'recovered' not in s and 'resolved' not in s
                valid_outcome = df[outcome_col].apply(is_not_fatal_or_recovered)
            else:
                valid_outcome = pd.Series([True] * len(df), index=df.index)
                
            # Check Term (exclude empty)
            if term_col:
                valid_term = df[term_col].apply(lambda x: bool(str(x).strip() and str(x).lower() != 'nan'))
            else:
                valid_term = pd.Series([True] * len(df), index=df.index)
                
            is_implied = non_ongoing_mask & no_end_date & valid_outcome & valid_term
            
            is_ongoing = is_marked_ongoing | is_implied
            
            stats['ongoing_aes'] = is_ongoing.sum()
            df['__is_ongoing'] = is_ongoing
            
            # Debug/Log if needed (could log is_implied.sum())
        else:
            df['__is_ongoing'] = False

        # Top Terms (case-insensitive grouping, preserving most common casing)
        if term_col:
            _terms = df[term_col].astype(str).str.strip()
            _terms_lower = _terms.str.lower()
            _lower_counts = _terms_lower.value_counts()
            # For each lowercase term, find the most common original casing
            _best_case = {}
            for orig_term in _terms.unique():
                lo = orig_term.lower()
                if lo not in _best_case:
                    _best_case[lo] = orig_term
            stats['top_terms'] = {_best_case.get(k, k): v for k, v in _lower_counts.head(10).items()}
            
        # SAE Criteria
        criteria_counts = {
            'Hospitalization': 0, 'Life-threatening': 0, 'Death': 0, 
            'Disability': 0, 'Other Med/Surg': 0
        }
        if hosp_col: criteria_counts['Hospitalization'] = df[hosp_col].apply(self._is_checked).sum()
        if life_col: criteria_counts['Life-threatening'] = df[life_col].apply(self._is_checked).sum()
        if death_col: criteria_counts['Death'] = df[death_col].apply(self._is_checked).sum()
        if disab_col: criteria_counts['Disability'] = df[disab_col].apply(self._is_checked).sum()
        if other_col: criteria_counts['Other Med/Surg'] = df[other_col].apply(self._is_checked).sum()
        
        stats['sae_criteria'] = criteria_counts
            
        # By Site
        df['Site'] = df['Screening #'].astype(str).str.split('-').str[0]
        stats['by_site'] = df['Site'].value_counts().to_dict()
        stats['by_patient'] = df['Screening #'].value_counts().to_dict()
        
        # Per Patient Details String
        # Group by patient
        grouped = df.groupby('Screening #')
        details = []
        
        for pid, group in grouped:
            n_aes = len(group)
            n_saes = group['__is_sae'].sum()
            n_ongoing = group['__is_ongoing'].sum()
            
            # Relationships
            def is_rel(row, col):
                if not col: return False
                val = row.get(col, '')
                return str(val).lower() not in ['not related', 'nan', '', 'none']

            n_device = 0
            n_proc = 0
            n_poss_proc = 0 # "Possibly Related" logic?
            
            for _, row in group.iterrows():
                # Device
                dev_rel = False
                for c in [rel_trillium_col, rel_device_col, rel_handle_col]:
                    if is_rel(row, c):
                        dev_rel = True
                        break
                if dev_rel: n_device += 1
                
                # Procedure
                proc_val = str(row.get(rel_proc_col, '')).lower()
                if 'possibly' in proc_val:
                    n_poss_proc += 1
                elif proc_val not in ['not related', 'nan', '', 'none']:
                    n_proc += 1
            
            line = (f"{pid}: {n_aes} AEs; including {n_saes} SAEs; {n_device} device-related; "
                    f"{n_proc} procedure-related; {n_poss_proc} possibly procedure-related; "
                    f"{n_ongoing} ongoing")
            details.append(line)
            
        stats['per_patient_details'] = sorted(details)
        
        # --- RELATEDNESS TABLE ---
        # Rows: Device, Delivery System, Handle, Procedure
        # Cols: Related, Probably Related, Possibly Related, Not Related, Unknown/Blank
        
        rel_map = {
            'Device': rel_trillium_col,
            'Delivery System': rel_device_col,
            'Handle': rel_handle_col,
            'Procedure': rel_proc_col
        }
        
        rel_table = {}
        
        for row_label, col_name in rel_map.items():
            counts = {
                'Related': 0, 'Probably Related': 0, 'Possibly Related': 0,
                'Not Related': 0, 'Unknown/Blank': 0, 'Related+Probably': 0
            }
            if col_name and col_name in df.columns:
                # Get series
                vals = df[col_name].astype(str).str.strip()
                
                for v in vals:
                    v_lower = v.lower()
                    if v_lower == 'related':
                        counts['Related'] += 1
                        counts['Related+Probably'] += 1
                    elif 'probably' in v_lower:
                        counts['Probably Related'] += 1
                        counts['Related+Probably'] += 1
                    elif 'possibly' in v_lower:
                        counts['Possibly Related'] += 1
                    elif 'not related' in v_lower:
                        counts['Not Related'] += 1
                    else:
                        # Blank, nan, None, Unknown
                        counts['Unknown/Blank'] += 1
            else:
                # If column missing, everything is blank? Or just 0s. 
                # Let's count all as Unknown/Blank if we have 0s logic, but better to leave 0 if no data?
                # Actually, if col is missing, likely input file issue. Keep 0s.
                pass
                
            rel_table[row_label] = counts
            
        stats['relatedness_table'] = rel_table

        # --- DEATH / MORTALITY DETAILS ---
        # Pull death form data from df_main for patients who have AEs
        death_details = []
        if self.df_main is not None:
            death_date_col = None
            death_cat_col = None
            death_reason_col = None
            for c in self.df_main.columns:
                cs = str(c)
                if 'DTH_DDDTC' in cs and death_date_col is None:
                    death_date_col = c
                if 'DTH_DDRESCAT' in cs and death_cat_col is None:
                    death_cat_col = c
                if 'DTH_DDORRES' in cs and death_reason_col is None:
                    death_reason_col = c

            ae_patients = df['Screening #'].unique()
            for pid in ae_patients:
                pat_main = self.df_main[self.df_main['Screening #'].astype(str).str.strip() == str(pid).strip()]
                if pat_main.empty:
                    continue
                pat_row = pat_main.iloc[0]
                d_date = str(pat_row.get(death_date_col, '')).strip() if death_date_col else ''
                if d_date and d_date.lower() not in ('nan', '', 'none', 'nat'):
                    d_cat = str(pat_row.get(death_cat_col, '')).strip() if death_cat_col else ''
                    d_reason = str(pat_row.get(death_reason_col, '')).strip() if death_reason_col else ''
                    if d_cat.lower() in ('nan', 'none', 'nat'):
                        d_cat = ''
                    if d_reason.lower() in ('nan', 'none', 'nat'):
                        d_reason = ''
                    death_details.append({
                        'patient_id': str(pid),
                        'death_date': self._clean_date(d_date) if d_date else '',
                        'mortality_classification': d_cat,
                        'cause_of_death': d_reason,
                    })

        stats['death_details'] = death_details

        return stats

    def _find_col(self, display_name):
        possible = self.col_mapping.get(display_name, [])
        for p in possible:
            if p in self.df_ae.columns:
                return p
        return None

    def _clean_date(self, val):
        if 'T' in val:
            val = val.split('T')[0]
        elif ' ' in val and any(c.isdigit() for c in val.split(' ')[-1]):
             # If there's a time portion after space (e.g., "2025-02-05 12:30")
             parts = val.split(' ')
             if len(parts) > 1 and ':' in parts[-1]:
                 val = ' '.join(parts[:-1])  # Remove time portion
        
        # Remove "time unknown"
        val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
        return val

    def _normalize_boolean(self, val):
        if str(val).lower() in ['yes', 'y', '1', 'true']:
            return 'Yes'
        elif str(val).lower() in ['no', 'n', '0', 'false']:
            return 'No'
        return val # Keep original if unknown or empty

    def _is_checked(self, val):
        return str(val).lower() in ['yes', 'y', '1', 'true', 'checked']

    def _parse_date_obj(self, date_str):
        if not date_str: return None
        try:
            return pd.to_datetime(date_str).date()
        except (ValueError, TypeError, OverflowError):
            return None

    def _get_procedure_date(self, patient_id):
        """Get extraction date from df_main (Treatment Date)."""
        if patient_id in self._procedure_dates:
            return self._procedure_dates[patient_id]
            
        # Find row in main
        if self.df_main is None: return None
        
        row = self.df_main[self.df_main['Screening #'].astype(str).str.strip() == patient_id.strip()]
        if row.empty: return None
        row = row.iloc[0]
        
        # Try to find treatment date columns
        # Priority: TV_PR_PRSTDTC (implant procedure date), TV_PR_SVDTC (treatment visit date)

        date_candidates = [
            'TV_PR_PRSTDTC', 'TV_PR_SVDTC',
        ]
        
        # Find mapped columns
        found_date = None
        for col_part in date_candidates:
             # Find actual column name in df_main that contains this part
             matches = [c for c in self.df_main.columns if col_part in str(c)]
             if matches:
                 val = row.get(matches[0])
                 if pd.notna(val) and str(val).strip():
                     found_date = self._parse_date_obj(str(val))
                     if found_date: break
        
        self._procedure_dates[patient_id] = found_date
        return found_date
