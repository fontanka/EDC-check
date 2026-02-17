import pandas as pd
import numpy as np
from collections import defaultdict
from sdv_manager import SDVManager

class DashboardManager:
    """
    Manages statistics and data retrieval for the SDV & Data Gap Dashboard.
    aggregates data from the SDVManager.
    """
    def __init__(self, sdv_manager: SDVManager):
        self.sdv_mgr = sdv_manager
        
        # Data structures to store aggregated stats
        # Keys: 'V', '!', 'NS', 'GAP'
        self.stats = {
            'study': defaultdict(int),
            'site': defaultdict(lambda: defaultdict(int)),
            'patient': defaultdict(lambda: defaultdict(int)),
            'form': defaultdict(lambda: defaultdict(int)) # Key: (PatientID, FormName) -> stats
        }
        
        self.details_df = None
        self.is_calculated = False

    def set_labels(self, labels: dict):
        """
        Sets the label dictionary (Code -> Label) for Field mapping.
        Generates common variations to handle SDV vs Project file discrepancies.
        """
        self.labels = labels.copy() if labels else {}
        
        # Visit prefixes used in data exports (must match all visits in config.py VISIT_MAP)
        visit_prefixes = [
            'SBV_', 'TV_', 'DV_',
            'FU1M_', 'FU3M_', 'FU6M_',
            'FU1Y_', 'FU2Y_', 'FU3Y_', 'FU4Y_', 'FU5Y_',
        ]
        
        # Generate variations for robust matching
        if labels:
            for key, label in labels.items():
                k = str(key).strip()
                parts = k.split('_')
                if len(parts) > 2:
                    # Variation 1: First + Last (SBV_DM_AGE -> SBV_AGE)
                    v1 = f"{parts[0]}_{parts[-1]}"
                    if v1 not in self.labels:
                        self.labels[v1] = label
                        
                    # Variation 2: Remove second part (SBV_SV_SVSTDTC -> SBV_SVSTDTC)
                    v2 = "_".join([parts[0]] + parts[2:])
                    if v2 not in self.labels:
                        self.labels[v2] = label
                
                # Cross-visit prefix variations (SBV_FAORRES_X -> TV_FAORRES_X, etc.)
                for prefix in visit_prefixes:
                    if k.startswith(prefix):
                        suffix = k[len(prefix):]  # Everything after prefix
                        for alt_prefix in visit_prefixes:
                            if alt_prefix != prefix:
                                alt_key = alt_prefix + suffix
                                if alt_key not in self.labels:
                                    self.labels[alt_key] = label
                        break
            
            # Build suffix index for fallback matching (last 2 parts, e.g., FAORRES_HR)
            # This helps with naming differences like SBV_ECHO_FAORRES_HR vs TV_FAORRES_PRE_HR
            self.suffix_labels = {}
            for key, label in self.labels.items():
                parts = key.split('_')
                if len(parts) >= 2:
                    suffix2 = '_'.join(parts[-2:])  # Last 2 parts
                    if suffix2 not in self.suffix_labels:
                        self.suffix_labels[suffix2] = label
                        
        # Trigger calculation to refresh fields with new labels
        self.calculate_stats(excluded_patients=None)

    def clean_label(self, label: str) -> str:
        """Clean label text similar to Clinical Viewer."""
        txt = str(label).strip()
        txt = txt.replace("_x0009_", "")
        for prefix in ["Sponsor/", "Sponsor ", "Core Lab/", "Core Lab "]:
             if txt.startswith(prefix): txt = txt[len(prefix):]
        
        # Simple mapping for common ugly labels
        lower = txt.lower()
        if "post-treatment hospitalizations" in lower and "status" in lower: return "Hospitalization Occurred?"
        if "reason for hospitalization" in lower: return "Reason"
        if "occurrence of heart failure" in lower: return "HF Hospitalization?"
        
        return txt

    def _preprocess_data(self, excluded_patients=None):
        """Copy modular data, add helper columns, filter exclusions.

        Returns pre-processed DataFrame or None if data unavailable.
        """
        df = self.sdv_mgr.modular_data
        if df is None:
            return None

        if excluded_patients:
            df = df[~df['Subject Screening #'].isin(set(excluded_patients))].copy()
        else:
            df = df.copy()

        df['Patient'] = df['Subject Screening #'].astype(str)
        df['Value'] = df['Variable Value'].astype(str).str.strip()
        empty_mask = df['Value'].str.lower().isin(['nan', 'none', '<na>', ''])
        df.loc[empty_mask, 'Value'] = ""
        df['HasValue'] = ~empty_mask
        df['Hidden'] = pd.to_numeric(df['Hidden'], errors='coerce').fillna(0).astype(int)
        df['CRA_STATUS'] = pd.to_numeric(df['CRA_CONTROL_STATUS'], errors='coerce').fillna(0).astype(int)
        df['Site'] = df['Patient'].str.split('-').str[0]
        return df

    def _map_labels_and_aggregate(self, df_metric):
        """Map variable names to human-readable labels and aggregate stats.

        Populates self.details_df and self.stats.
        """
        df_metric['FieldID'] = df_metric['Variable name']

        if hasattr(self, 'labels') and self.labels:
            df_metric['VarClean'] = df_metric['Variable name'].astype(str).str.strip()
            df_metric['RawLabel'] = df_metric['VarClean'].map(self.labels)

            if hasattr(self, 'suffix_labels') and self.suffix_labels:
                _needs_fallback = df_metric['RawLabel'].isna()
                if _needs_fallback.any():
                    _fb = df_metric.loc[_needs_fallback, 'VarClean'].astype(str)
                    _suffix = _fb.str.rsplit('_', n=2).str[-2:].str.join('_')
                    _mapped = _suffix.map(self.suffix_labels)
                    df_metric.loc[_needs_fallback, 'RawLabel'] = _mapped.fillna(
                        df_metric.loc[_needs_fallback, 'Variable name'])
            else:
                df_metric['RawLabel'] = df_metric['RawLabel'].fillna(df_metric['Variable name'])

            df_metric['Field'] = df_metric['RawLabel'].apply(self.clean_label)
        else:
            df_metric['Field'] = df_metric['Variable name']

        if 'Date' not in df_metric.columns:
            df_metric['Date'] = ""

        self.details_df = df_metric[
            ['Patient', 'Site', 'VisitName', 'FormName', 'Field', 'FieldID',
             'Value', 'Metric', 'Date']]

        # Aggregation
        self.stats['study'] = df_metric['Metric'].value_counts().to_dict()

        site_counts = df_metric.groupby(['Site', 'Metric']).size()
        for (site, metric), count in site_counts.items():
            self.stats['site'][site][metric] = count

        pat_counts = df_metric.groupby(['Patient', 'Metric']).size()
        for (pat, metric), count in pat_counts.items():
            self.stats['patient'][pat][metric] = count

        form_counts = df_metric.groupby(['Patient', 'FormName', 'Metric']).size()
        for (pat, form, metric), count in form_counts.items():
            self.stats['form'][(pat, form)][metric] = count

        self.is_calculated = True

    def calculate_stats(self, excluded_patients: list = None) -> None:
        """
        Calculates statistics from raw SDVManager data (Global Scope).
        Uses self.labels to map Variable Names to Human-Readable Fields.
        """
        if not self.sdv_mgr.is_loaded():
            return

        # Reset simple stats structures
        self.stats = {
            'study': defaultdict(int),
            'site': defaultdict(lambda: defaultdict(int)),
            'patient': defaultdict(lambda: defaultdict(int)),
            'form': defaultdict(lambda: defaultdict(int))
        }
        self.details_df = None
        self.is_calculated = False

        df = self._preprocess_data(excluded_patients)
        if df is None:
            return

        # 2. STATUS DETERMINATION (V, !, GAP)
        cond_ver = df['CRA_STATUS'].isin([2, 4])
        # Pending: Status 1 (Changed) or 3 (Query Answered) are implicitly pending.
        # Status 0 (Created) is only pending if Value Exists (data was entered).
        cond_pending = (df['CRA_STATUS'].isin([1, 3])) | ( (df['CRA_STATUS'] == 0) & (df['HasValue']) & (df['Hidden'] != 1) )
        
        # 3. FAST FORM STATUS LOOKUP (for NS and Verified Forms)
        not_sent_keys = set()
        verified_form_keys = set()  # NEW: Track verified forms
        for key, status_tuple in self.sdv_mgr.form_entry_status.items():
            form_status, ver_status = status_tuple[0], status_tuple[1]
            if form_status == 'Created' and ver_status in ['Blank', 'nan', 'None', '']:
                parts = key.split('|')
                if len(parts) >= 4:
                    not_sent_keys.add( (parts[0].lower(), parts[2].lower(), parts[1].lower(), parts[3]) )
            # NEW: Identify verified forms (green/blue checkmarks)
            if ver_status in ['Verified', 'SDV Verified', 'DMR Verified']:
                parts = key.split('|')
                if len(parts) >= 4:
                    verified_form_keys.add( (parts[0].lower(), parts[2].lower(), parts[1].lower(), parts[3]) )

        if 'Form name' not in df.columns:
             df['FormName'] = df.get('Form Name', df['Form Code']).astype(str).str.strip()
        else:
             df['FormName'] = df['Form name'].astype(str).str.strip()
             
        if 'Visit name' not in df.columns:
             df['VisitName'] = df.get('Folder Name', pd.Series([""]*len(df))).astype(str).str.strip()
        else:
             df['VisitName'] = df['Visit name'].astype(str).str.strip()
             
        df['TableRow'] = pd.to_numeric(df['Table row #'], errors='coerce').fillna(0).astype(int).astype(str)
        
        # Create Key in DF
        df['LookupKey'] = list(zip(
            df['Patient'].str.lower(), 
            df['FormName'].str.lower(), 
            df['VisitName'].str.lower(),
            df['TableRow']
        ))
        
        df['IsNS'] = df['LookupKey'].isin(not_sent_keys)
        df['FormIsVerified'] = df['LookupKey'].isin(verified_form_keys)
        
        # 3a. FILTER OUT FUTURE/UNSTARTED VISITS
        # Forms that have NO data at all are likely future visits - should not be counted as Gaps
        # Group by FormKey and check if any field in the form has a value
        forms_with_data = set()
        
        # Use FormKey (without TableRow) to check at form/visit level
        df['FormKey'] = list(zip(
            df['Patient'].str.lower(), 
            df['FormName'].str.lower(), 
            df['VisitName'].str.lower()
        ))
        
        # Find all forms that have at least one filled field
        filled_forms = df[df['HasValue']].groupby(['Patient', 'FormName', 'VisitName']).size()
        for (patient, form, visit), _ in filled_forms.items():
            forms_with_data.add((str(patient).lower(), str(form).lower(), str(visit).lower()))
            
        # 3a-bis. ECHO SISTER FORM LOGIC
        # If a site Echo form is empty, but its Core Lab sister has data, it's a Gap.
        # Define strict pairs based on analysis of actual data
        core_to_site_map = {
            'Echocardiography - Core lab': 'Echocardiography',
            'Echocardiography – Core lab': 'Echocardiography',
            'Echocardiography – 1 day prior the procedure - Core lab': 'Echocardiography – 1 day prior the procedure',
            'Echocardiography – 1-day post procedure - Core lab': 'Echocardiography – 1-day post procedure',
            'Echocardiography – Pre and Post procedure - Core lab': 'Echocardiography – Pre and Post procedure'
        }
        
        is_core = df['FormName'].str.contains('Core lab', case=False, na=False)
        core_with_data = df[is_core & df['HasValue']][['Patient', 'VisitName', 'FormName']].drop_duplicates()
        
        # Build a set of (Patient, Visit, SiteFormName) that should be considered "Started" via proxy
        proxy_started_forms = set()
        
        if not core_with_data.empty:
            core_keys = list(zip(
                core_with_data['Patient'].astype(str),
                core_with_data['VisitName'].astype(str),
                core_with_data['FormName'].astype(str)
            ))
            for pat, vis, core_name in core_keys:
                core_name_stripped = core_name.strip()
                site_name = core_to_site_map.get(core_name_stripped)
                if not site_name and 'Echocardiography' in core_name_stripped:
                    normalized = core_name_stripped.replace(' - Core lab', '').replace(' \u2013 Core lab', '').strip()
                    if normalized != core_name_stripped:
                        site_name = normalized
                if site_name:
                    proxy_started_forms.add((pat.lower(), vis.lower(), site_name.lower()))
        
        # Now add these to forms_with_data if the site form actually exists in the data for that patient/visit
        # (It should, as an empty form row, but we want to be safe)
        # We don't strictly need to check existence if we rely on "if it's not in the data, it's not a gap anyway" logic of CondGap?
        # Actually CondGap requires the row to exist.
        # But here we are setting 'FormHasAnyData' which is a prerequisite for being a Gap (usually).
        # So we just add the key.
        for key in proxy_started_forms:
            # key is (patient, visit, site_form_name) -> matches FormKey structure? 
            # FormKey structure in line 169 is (Patient, FormName, VisitName)
            # So we need to reorder: (Patient, SiteFormName, VisitName)
            pat, vis, form = key
            forms_with_data.add((pat, form, vis))
        
        df['FormHasAnyData'] = df['FormKey'].isin(forms_with_data)
        
        # 3b. ECG-SPECIFIC CHECKBOX LOGIC
        # If an ECG form has EGORRES_RHYTHM field filled, treat checkbox fields as Pending (!), not Gap
        ecg_forms_with_rhythm = set()
        
        # Find ECG forms where EGORRES_RHYTHM has data
        ecg_mask = df['FormName'].str.contains('ECG', case=False, na=False)
        rhythm_mask = df['Variable name'].str.contains('EGORRES_RHYTHM', case=False, na=False)
        ecg_rhythm_filled = df[ecg_mask & rhythm_mask & df['HasValue']]
        
        if not ecg_rhythm_filled.empty:
            ecg_forms_with_rhythm = set(zip(
                ecg_rhythm_filled['Patient'].str.lower(),
                ecg_rhythm_filled['FormName'].str.lower(),
                ecg_rhythm_filled['VisitName'].str.lower()
            ))
        
        # Create lookup key for ECG forms (without TableRow for form-level matching)
        df['FormKey'] = list(zip(
            df['Patient'].str.lower(), 
            df['FormName'].str.lower(), 
            df['VisitName'].str.lower()
        ))
        
        # Identify ECG checkbox fields (_ABN pattern)
        df['IsECGCheckbox'] = (
            df['FormKey'].isin(ecg_forms_with_rhythm) & 
            df['Variable name'].str.contains('_ABN|_EGORRES_', case=False, na=False, regex=True)
        )
        
        # Update pending condition to include ECG checkbox fields
        cond_pending = (
            (df['CRA_STATUS'].isin([1, 3])) | 
            ( (df['CRA_STATUS'] == 0) & (df['HasValue']) & (df['Hidden'] != 1) ) |
            (df['IsECGCheckbox'] & (~df['HasValue']))  # ECG checkboxes without value are Pending
        )
        
        # Identify "Not Done/Not Recorded" checkbox fields by checking if their label contains these patterns
        # Build a set of variable names whose labels contain "not done" or "not recorded"
        not_done_vars = set()
        if hasattr(self, 'labels') and self.labels:
            for var_name, label in self.labels.items():
                if isinstance(label, str):
                    label_lower = label.lower()
                    if 'not done' in label_lower or 'not recorded' in label_lower:
                        not_done_vars.add(var_name.lower())
        
        df['IsNotDoneField'] = df['Variable name'].str.lower().isin(not_done_vars)
        
        # 3c. LBSTAT FIELD LOGIC
        # LBSTAT fields are "Not Done" checkboxes for lab results
        # If the corresponding result field has data, the LBSTAT checkbox being empty is intentional
        lbstat_mask = df['Variable name'].str.contains('LBSTAT', case=False, na=False)
        
        # For each form, check if any non-LBSTAT field has data (indicating test was done)
        forms_with_results = set()
        result_mask = ~lbstat_mask & df['HasValue']  # Non-LBSTAT fields with values
        result_forms = df[result_mask].groupby(['Patient', 'FormName', 'VisitName']).size()
        for (patient, form, visit), _ in result_forms.items():
            forms_with_results.add((str(patient).lower(), str(form).lower(), str(visit).lower()))
        
        # LBSTAT fields in forms that have result data are not gaps
        df['IsLBSTATWithResult'] = lbstat_mask & df['FormKey'].isin(forms_with_results)
        
        # 3d. FASTAT FIELD LOGIC (similar to LBSTAT, for ALL Echocardiography forms)
        # FASTAT fields are "Not Done" status fields for Echo measurements
        # If the form has corresponding result data, the FASTAT checkbox being empty is intentional
        fastat_mask = df['Variable name'].str.contains('FASTAT', case=False, na=False)
        any_echo_form_mask = df['FormName'].str.contains('Echocardiography', case=False, na=False)  # ALL Echo forms
        
        # For each Echo form, check if any non-FASTAT field has data
        echo_forms_with_results = set()
        echo_result_mask = any_echo_form_mask & ~fastat_mask & df['HasValue']
        echo_result_forms = df[echo_result_mask].groupby(['Patient', 'FormName', 'VisitName']).size()
        for (patient, form, visit), _ in echo_result_forms.items():
            echo_forms_with_results.add((str(patient).lower(), str(form).lower(), str(visit).lower()))
        
        # FASTAT fields in ANY Echo forms that have result data are not gaps
        df['IsFASTATWithResult'] = fastat_mask & df['FormKey'].isin(echo_forms_with_results)
        
        # 3e. ECHOCARDIOGRAPHY FAORRES FIELDS (Core Lab specific)
        # FAORRES fields (result fields) in Echocardiography Core Lab that do NOT end with _SP are excluded
        # These are measurement fields that can legitimately be empty
        echo_core_form_mask = df['FormName'].str.contains('Echocardiography', case=False, na=False) & df['FormName'].str.contains('Core', case=False, na=False)
        faorres_mask = df['Variable name'].str.contains('FAORRES', case=False, na=False)
        not_sp_mask = ~df['Variable name'].str.endswith('_SP', na=False)
        df['IsEchoFAORRES'] = echo_core_form_mask & faorres_mask & not_sp_mask
        
        
        # 3f. ECHO "REASON WHY NOT PERFORMED" COMMENT FIELD (ALL Echo forms)
        # If the form has any result/date data (implying assessment was performed), 
        # then "reason why not performed" (PRREASND/REASON) is irrelevant.
        # We can reuse 'echo_forms_with_results' which we calculated for FASTAT logic above
        # (It contains forms with non-FASTAT data).
        # We should ensure that 'reason' fields themselves don't count as "results" for this check
        # But 'echo_result_mask' was: any_echo_form_mask & ~fastat_mask & df['HasValue']
        # This might include PRREASND itself if it has value.
        # But if PRREASND has value, IsIrrelevantComment doesn't matter (it's not a gap).
        # The Gap logic checks ~HasValue.
        # So crucial part is: Does the form have OTHER data?
        # 'echo_forms_with_results' includes forms with specific 'fastat' excluded.
        # Let's refine 'echo_forms_with_results' or make a new set that looks for "Positive" data.
        
        # Let's trust 'echo_forms_with_results' for now as it captures forms with Data.
        # (If only PRREASND is filled, then PRREASND is not a Gap anyway).
        # The issue is when PRREASND is EMPTY. If form has other data -> Irrelevant.
        
        reason_comment_mask = df['Variable name'].str.contains('REASND|REASON', case=False, na=False, regex=True)
        df['IsIrrelevantComment'] = any_echo_form_mask & reason_comment_mask & df['FormKey'].isin(echo_forms_with_results)
        
        # 3g. TEST PARAMS ROW COMMENTS (optional comment fields)
        df['IsTestParamsComment'] = df['Variable name'].str.contains('TestParamsRowComments', case=False, na=False)
        
        # 3h. PRE-PROCEDURE CHECKLIST COMMENTS (PRCOMM fields)
        # If the form has data (checklist is filled), comment fields can be empty
        prcomm_mask = df['Variable name'].str.contains('PRCOMM', case=False, na=False)
        df['IsPRCOMMWithData'] = prcomm_mask & df['FormHasAnyData']
        
        # 3i. PARTIAL FIELDS (PARTIAL_CHECKBOX or _PARTIAL)
        # These are "Full date unknown" checkboxes. If the corresponding date field has a value,
        # the partial checkbox is irrelevant (the full date is known)
        partial_mask = df['Variable name'].str.contains('PARTIAL', case=False, na=False)
        
        # For each PARTIAL field, find the corresponding date field
        # e.g., SBV_HOSTDTC_PARTIAL_CHECKBOX -> SBV_HOSTDTC
        # e.g., SBV_BRTHDAT_PARTIAL -> SBV_BRTHDAT
        df['DateFieldBase'] = df['Variable name'].str.replace('_PARTIAL_CHECKBOX', '', case=False, regex=False).str.replace('_PARTIAL', '', case=False, regex=False)
        
        # Build set of all fields with values using vectorized zip (no iterrows)
        _has_val = df[df['HasValue']]
        date_fields_with_value = set(zip(
            _has_val['Patient'].str.lower(),
            _has_val['FormName'].str.lower(),
            _has_val['VisitName'].str.lower(),
            _has_val['TableRow'],
            _has_val['Variable name'].str.lower()
        ))
        
        # Vectorized: build lookup key for partial fields and check against set
        _partial_lookup = list(zip(
            df['Patient'].str.lower(),
            df['FormName'].str.lower(),
            df['VisitName'].str.lower(),
            df['TableRow'],
            df['DateFieldBase'].str.lower()
        ))
        df['IsPartialWithDate'] = partial_mask & pd.Series(_partial_lookup, index=df.index).isin(date_fields_with_value)
        
        # 3j. GENERAL COMMENT FIELDS (ending with COMM or _COMM)
        # Comment fields are optional if form has other data
        general_comment_mask = df['Variable name'].str.match(r'.*_?COMM$', case=False, na=False)
        df['IsGeneralComment'] = general_comment_mask & df['FormHasAnyData']

        # 3l. PESTAT FIELDS (Physical Exam Status)
        # Exclude PESTAT fields from Gaps if their corresponding PEORRES field has data.
        # e.g., SBV_PESTAT_CARD -> SBV_PEORRES_CARD
        # Special case: SBV_PESTAT -> SBV_PEORRES
        pestat_mask = df['Variable name'].str.contains('PESTAT', case=False, na=False)
        
        # Calculate Base Name: replace PESTAT with PEORRES
        # Note: double underscores __NEUR need handling if strictly mapped, but usually standard replacement works
        # If SBV_PESTAT__NEUR -> SBV_PEORRES__NEUR (if data matches this, great. If data uses _NEUR, we might miss it)
        # Based on script output, results are SBV_PEORRES_NEUR (single underscore).
        # But PESTAT is SBV_PESTAT__NEUR (double underscore).
        # So we should normalize both to single underscore or handle the specific double->single replacement.
        # Let's clean the base name: replace PESTAT with PEORRES, then replace __ with _ just in case.
        
        df['PEORRESBase'] = df['Variable name'].str.replace('PESTAT', 'PEORRES', case=False, regex=False).str.replace('__', '_', regex=False)
        
        # We can reuse the `date_fields_with_value` set if we rename it to more generic `fields_with_value`
        # OR just iterate again for clarity (negligible performance difference for this size).
        # Let's reuse `date_fields_with_value` as it contains ALL fields with values?
        # Line 348 iterrows over `df[df['HasValue']]`, so yes, it contains ALL variables with values, not just Dates.
        # So I can just check against `date_fields_with_value`.
        
        # Vectorized: build PEORRES lookup key and check against set
        _peorres_lookup = list(zip(
            df['Patient'].str.lower(),
            df['FormName'].str.lower(),
            df['VisitName'].str.lower(),
            df['TableRow'],
            df['PEORRESBase'].str.lower()
        ))
        df['IsPESTATWithData'] = pestat_mask & pd.Series(_peorres_lookup, index=df.index).isin(date_fields_with_value)
        
        # 3k. TIME UNKNOWN FIELDS (TIMUNC)
        # These are "Time unknown" checkboxes. If the corresponding time field has a value,
        # the unknown checkbox is irrelevant
        timunc_mask = df['Variable name'].str.contains('TIMUNC', case=False, na=False)
        
        # Determine base time field: Replace TIMUNC with TIM
        # e.g., SBV_LBTIMUNC_BM -> SBV_LBTIM_BM
        df['TimeFieldBase'] = df['Variable name'].str.replace('TIMUNC', 'TIM', case=False, regex=False)
        
        # Vectorized: build time lookup key and check against set
        _time_lookup = list(zip(
            df['Patient'].str.lower(),
            df['FormName'].str.lower(),
            df['VisitName'].str.lower(),
            df['TableRow'],
            df['TimeFieldBase'].str.lower()
        ))
        df['IsTimeUnknownWithTime'] = timunc_mask & pd.Series(_time_lookup, index=df.index).isin(date_fields_with_value)

        # 3m. AE CHECKBOX GROUPS (Action Taken & Seriousness)
        # Groups: LOGS_AEACN_... and LOGS_AES... (excluding AESEV/AESER if needed, but AESER usually key)
        # Detailed logic:
        # AEACN Group: Any variable starting with LOGS_AEACN_
        # AES Group: Any variable starting with LOGS_AES... (including AESHOSP, AESDTH, etc.)
        # If ANY field in the group has a value, the empty ones are not gaps (just unchecked options).
        
        aeacn_mask = df['Variable name'].str.startswith('LOGS_AEACN_', na=False)
        aes_mask = df['Variable name'].str.contains('LOGS_AES|LOGS_AESHOSP', case=False, na=False) 
        # Note: AESHOSP is part of the seriousness criteria group but might not start with AES if naming varies, 
        # but typically it is LOGS_AESHOSP. The regex covers it.
        # Exclude AESER (Serious Event Y/N) and AESEV (Severity) from the "Group Check" if we want strictly checkboxes?
        # Actually, if AESER is "Yes", then one of the AES... boxes should be checked.
        # If AESER is "No", then AES... boxes are empty.
        # If we just say "If any AES... box is checked, the others are not gaps", that handles the "Selected" case.
        # What if NONE are checked? Then they are all Gaps (if AESER=Yes).
        # But if AESER=No, they are all empty and should NOT be Gaps (if logic allows).
        # However, FormHasAnyData is True. So they show as Gaps.
        # If AESER=No, we should trigger "irrelevant".
        # But let's stick to the "Checkbox Group" logic: If >0 checked, rest are implied unchecked.
        # If 0 checked, they remain Gaps (which is correct if one MUST be checked).
        # But for 'Action Taken', maybe 'None' is an option? Yes, LOGS_AEACN_NONE.
        
        # We need to find "Groups with at least one value".
        # We can do this by aggregating HasValue over (Patient, Visit, Form, Row) for the group.
        
        # Helper to mark groups with data
        # Vectorized: build group data masks using zip instead of iterrows
        def get_group_data_keys(group_mask):
            group_df = df[group_mask & df['HasValue']]
            if group_df.empty:
                return set()
            return set(zip(
                group_df['Patient'].str.lower(),
                group_df['VisitName'].str.lower(),
                group_df['FormName'].str.lower(),
                group_df['TableRow']
            ))

        aeacn_keys_with_data = get_group_data_keys(aeacn_mask)
        aes_keys_with_data = get_group_data_keys(aes_mask)
        
        # Vectorized: build row-level key and check group membership
        _ae_row_key = list(zip(
            df['Patient'].str.lower(),
            df['VisitName'].str.lower(),
            df['FormName'].str.lower(),
            df['TableRow']
        ))
        _ae_key_series = pd.Series(_ae_row_key, index=df.index)
        df['IsAECheckboxGroupWithData'] = (
            (aeacn_mask & _ae_key_series.isin(aeacn_keys_with_data)) |
            (aes_mask & _ae_key_series.isin(aes_keys_with_data))
        )

        # 3n. GENERIC STAT FIELDS (Echo cleanup etc.)
        # Exclude fields ending in STAT, STAT_..., _STAT from Gaps.
        # Rationale: If the main result is missing, that is the Gap. The Status field is secondary.
        # Regex: Ends with STAT, or STAT_something, or _STAT
        # Examples: TV_PRSTAT_PRE_ECHO, TV_FASTAT_PRE_HR, SBV_QSSTAT_MNA
        # Fix: Previous regex r'(_STAT$|_STAT_)' missed PRSTAT_PRE where underscore is NOT before S.
        # identifying STAT followed by _ or end, or preceded by _.
        
        stat_mask = df['Variable name'].str.contains(r'(STAT_|STAT$|_STAT)', case=False, na=False, regex=True)
        df['IsGenericStatusField'] = stat_mask

        # 3o. LAB METADATA (Ranges, Units, Reasons, Comments)
        # Exclude: LBORNRLO, LBORNRHI, LBORRESUN, LBORRESU, REASND, LBCOMMENT
        # Also PRSCAT (Procedure Category), SUPPPR (Supplemental Procedure Questions)
        # LOGS_LBREF (Additional laboratory test reference in AE logs)
        # REASND is generic "Reason Not Done", usually secondary to Result Gap or Status=Not Done.
        # If Result is missing, that's the Gap.
        lab_meta_mask = df['Variable name'].str.contains(
            r'(?:LBORNRLO|LBORNRHI|LBORRESUN|LBORRESU|REASND|PRSCAT|SUPPPR|LOGS_LBREF|LBCOMMENT)', 
            case=False, na=False, regex=True
        )
        df['IsLabMetadata'] = lab_meta_mask

        # 3p. DATE WITH RESULT (EGDTC, DTC empty but ORRES has data)
        # If the Date field is empty but the form has Result data, the Date is secondary.
        # Flag as ! instead of GAP.
        date_mask = df['Variable name'].str.contains(r'(?:EGDTC|_DTC$|_DTC_|LBDTC)', case=False, na=False, regex=True)
        
        # Find forms (Patient + Visit + Form) that have ORRES data
        orres_mask = df['Variable name'].str.contains(r'ORRES', case=False, na=False)
        orres_with_value = df[orres_mask & df['HasValue']]
        forms_with_orres = set(orres_with_value[['Patient', 'VisitName', 'FormName']].apply(tuple, axis=1))
        
        # Vectorized: Create form key column and check membership
        df['_FormKey'] = list(zip(df['Patient'], df['VisitName'], df['FormName']))
        df['IsDateWithResult'] = date_mask & (~df['HasValue']) & df['_FormKey'].isin(forms_with_orres)
        df.drop(columns=['_FormKey'], inplace=True)

        # 3q. PGA COMMENTS (If PGA form has data, blank comment fields are not gaps)
        pga_form_mask = df['FormName'].str.contains('Physician Global Assessment', case=False, na=False)
        pga_comment_mask = df['Variable name'].str.contains(r'COMM|PGA', case=False, na=False) & pga_form_mask
        # Find PGA forms with any data
        pga_with_data = df[pga_form_mask & df['HasValue']]
        pga_forms_with_data = set(pga_with_data[['Patient', 'VisitName', 'FormName']].apply(tuple, axis=1))
        df['_FormKey'] = list(zip(df['Patient'], df['VisitName'], df['FormName']))
        df['IsPGACommentWithData'] = pga_comment_mask & (~df['HasValue']) & df['_FormKey'].isin(pga_forms_with_data)
        
        # 3r. AE ONGOING (If AE End Date is present, blank Ongoing is not a gap)
        ae_form_mask = df['FormName'].str.contains('Adverse Event', case=False, na=False)
        ae_ongoing_mask = df['Variable name'].str.contains(r'AEONGO|AONGO|_ONGO', case=False, na=False) & ae_form_mask
        # Find AEs with end date (AEENDTC or similar)
        ae_enddate_mask = df['Variable name'].str.contains(r'AEENDTC|AEEN', case=False, na=False) & ae_form_mask
        ae_with_enddate = df[ae_enddate_mask & df['HasValue']]
        # Key by Patient + Row (need to link AE row, using FormKey + approx row via Patient+Visit+Form for now)
        ae_forms_with_enddate = set(ae_with_enddate[['Patient', 'VisitName', 'FormName']].apply(tuple, axis=1))
        df['IsAEOngoingWithEndDate'] = ae_ongoing_mask & (~df['HasValue']) & df['_FormKey'].isin(ae_forms_with_enddate)
        
        # 3s. SAE COMMENTS (If event is NOT SAE, blank SAE/sequelae comments are not gaps)
        sae_comment_mask = df['Variable name'].str.contains(r'AETERM_COMM|SEQUELAE|SAE.*COMM', case=False, na=False, regex=True) & ae_form_mask
        # Find AEs where SAE indicator is "N" or "No"
        # This is complex - for now, mark all SAE comments in AE form as excluded if form has data
        ae_with_data = df[ae_form_mask & df['HasValue']]
        ae_forms_with_data = set(ae_with_data[['Patient', 'VisitName', 'FormName']].apply(tuple, axis=1))
        df['IsSAECommentWithData'] = sae_comment_mask & (~df['HasValue']) & df['_FormKey'].isin(ae_forms_with_data)
        
        # 3t. CONMED ONGOING (If drug End Date is present, blank Ongoing is not a gap)
        cm_form_mask = df['FormName'].str.contains('Concomitant Medications', case=False, na=False)
        cm_ongoing_mask = df['Variable name'].str.contains(r'CMONGO|_ONGO', case=False, na=False) & cm_form_mask
        cm_enddate_mask = df['Variable name'].str.contains(r'CMENDTC|CMEN', case=False, na=False) & cm_form_mask
        cm_with_enddate = df[cm_enddate_mask & df['HasValue']]
        cm_forms_with_enddate = set(cm_with_enddate[['Patient', 'VisitName', 'FormName']].apply(tuple, axis=1))
        df['IsCMOngoingWithEndDate'] = cm_ongoing_mask & (~df['HasValue']) & df['_FormKey'].isin(cm_forms_with_enddate)
        
        # 3u. MEDICAL HISTORY ONGOING (If MH End Date is present, blank Ongoing is not a gap)
        mh_form_mask = df['FormName'].str.contains('Medical History', case=False, na=False)
        mh_ongoing_mask = df['Variable name'].str.contains(r'MHONGO', case=False, na=False) & mh_form_mask
        mh_enddate_mask = df['Variable name'].str.contains(r'MHENDTC|MHEN', case=False, na=False) & mh_form_mask
        mh_with_enddate = df[mh_enddate_mask & df['HasValue']]
        mh_forms_with_enddate = set(mh_with_enddate[['Patient', 'VisitName', 'FormName']].apply(tuple, axis=1))
        df['IsMHOngoingWithEndDate'] = mh_ongoing_mask & (~df['HasValue']) & df['_FormKey'].isin(mh_forms_with_enddate)
        
        df.drop(columns=['_FormKey'], inplace=True)

        # 3l. AGGREGATE EXCLUDED GAPS
        # User wants these empty "irrelevant" fields to be listed as "Pending" (!) for verification
        # instead of being hidden.
        df['IsExcludedGap'] = (
            df['IsNotDoneField'] |
            df['IsLBSTATWithResult'] |
            df['IsFASTATWithResult'] |
            df['IsEchoFAORRES'] |
            df['IsIrrelevantComment'] |
            df['IsTestParamsComment'] |
            df['IsPRCOMMWithData'] |
            df['IsPartialWithDate'] | 
            df['IsGeneralComment'] |
            df['IsTimeUnknownWithTime'] |
            df['IsPESTATWithData'] |
            df['IsAECheckboxGroupWithData'] |
            df['IsGenericStatusField'] |
            df['IsLabMetadata'] |
            df['IsDateWithResult'] |
            df['IsPGACommentWithData'] |
            df['IsAEOngoingWithEndDate'] |
            df['IsSAECommentWithData'] |
            df['IsCMOngoingWithEndDate'] |
            df['IsMHOngoingWithEndDate']
        )
        # Note: IsECGCheckbox is already handled in cond_pending
        
        # GAP: Empty AND not hidden AND form is NOT verified AND not ECG checkbox 
        # AND not Excluded Gap
        # Forms with NO data at all are future visits - not gaps
        cond_gap = (~df['HasValue']) & (df['Hidden'] == 0) & (~df['FormIsVerified']) & (~df['IsECGCheckbox']) & (df['FormHasAnyData']) & (~df['IsExcludedGap'])

        # Update pending condition to include Excluded Gaps (empty fields)
        # We need to redefine cond_pending here because we added logic after its simplified definition above
        # (Actually, cond_pending was defined way up. We should update it here using boolean OR)
        
        # Re-eval cond_pending with new logic:
        # Original: (Status=1/3) OR (Status=0 & HasValue) OR (ECGCheckbox & ~HasValue)
        # New: Original OR (IsExcludedGap & ~HasValue & Hidden==0 & ~FormIsVerified)
        # Verify strictness: If form is verified, they are V (handled by cond_ver).
        
        cond_pending = cond_pending | (
            df['IsExcludedGap'] & (~df['HasValue']) & (df['Hidden'] == 0) & (~df['FormIsVerified'])
        )

        # 4. FINAL METRIC ASSIGNMENT
        # Priority: NS > V > ! > GAP
        conditions = [
            df['IsNS'],
            cond_ver & (~df['IsNS']), 
            cond_pending & (~df['IsNS']),
            cond_gap & (~df['IsNS'])
             # & ~cond_ver & ~cond_pending implied by order usually, 
             # but np.select executes in order. 
             # Check: If row is V, first match is V.
             # If row is GAP (empty, status 0), match GAP.
        ]
        choices = ['NS', 'V', '!', 'GAP']
        
        df['Metric'] = np.select(conditions, choices, default='')
        
        df_metric = df[df['Metric'] != ''].copy()

        # 5-6. LABEL MAPPING + AGGREGATION (extracted to helper)
        self._map_labels_and_aggregate(df_metric)


    def get_summary(self, level: str, level_id: str = None) -> dict:
        if not self.is_calculated:
            self.calculate_stats()
            
        if level == 'study':
            return self.stats['study']
        elif level == 'site':
            return self.stats['site'].get(level_id, {}) if level_id else self.stats['site']
        elif level == 'patient':
            return self.stats['patient'].get(level_id, {}) if level_id else self.stats['patient']
        elif level == 'form':
            return self.stats['form'].get(level_id, {}) if level_id else self.stats['form']
        return {}
    
    def get_details(self, level: str, level_id: str, metric: str) -> list:
        """Returns a list of dicts for the specified drill-down level/metric."""
        if not self.is_calculated or self.details_df is None:
            self.calculate_stats()
            
        if self.details_df is None or self.details_df.empty:
            return []

        # Filter DataFrame based on request
        mask = (self.details_df['Metric'] == metric)
        
        if level == 'study':
            pass # All items with that metric
        elif level == 'site':
            mask &= (self.details_df['Site'].astype(str) == str(level_id))
        elif level == 'patient':
            mask &= (self.details_df['Patient'] == str(level_id))
        elif level == 'form':
            # level_id is (Patient, Form)
            if isinstance(level_id, (list, tuple)) and len(level_id) == 2:
                pat, form = level_id
                mask &= (self.details_df['Patient'] == str(pat)) & (self.details_df['FormName'] == str(form))
        
        result_df = self.details_df[mask].copy()
        
        if metric == 'NS':
            # Aggregate to Form Level
            result_df = result_df.drop_duplicates(subset=['Patient', 'VisitName', 'FormName']).copy()
            result_df['Variable name'] = ""
            result_df['Field'] = ""
            result_df['FieldID'] = ""
        
        # Ensure 'Code' exists (for UI)
        if 'Code' not in result_df.columns:
            if 'FieldID' in result_df.columns:
                result_df['Code'] = result_df['FieldID']
            elif 'Variable name' in result_df.columns:
                result_df['Code'] = result_df['Variable name']
        
        # Ensure 'Field' exists (Label)
        if 'Field' not in result_df.columns:
            if 'Variable name' in result_df.columns:
                result_df['Field'] = result_df['Variable name']
            elif 'FieldID' in result_df.columns:
                result_df['Field'] = result_df['FieldID']

        # Handle Status collision
        if 'Status' in result_df.columns:
            result_df.rename(columns={'Status': 'RawStatus'}, inplace=True)

        rename_map = {
            'VisitName': 'Visit',
            'FormName': 'Form',
            'Metric': 'Status'
        }
        records = result_df.rename(columns=rename_map).to_dict('records')
        
        # For Verified items, add verification metadata (User, Date)
        if metric == 'V' and hasattr(self.sdv_mgr, 'get_verification_details'):
            for rec in records:
                patient = rec.get('Patient', '')
                form = rec.get('Form', '')
                visit = rec.get('Visit', '')
                field_id = rec.get('Code', '') or rec.get('FieldID', '')
                details = self.sdv_mgr.get_verification_details(patient, form, visit, field_id=field_id)
                if details:
                    rec['VerifiedBy'] = details.get('user', '')
                    rec['Date'] = details.get('date', '')
                else:
                    rec['VerifiedBy'] = ''
                    rec['Date'] = ''
        
        return records

    def get_top_counts(self, level, metric, top_n=10):
        if not self.is_calculated:
            self.calculate_stats()
        
        data_source = self.stats[level]
        sorted_items = sorted(data_source.items(), 
                              key=lambda x: x[1].get(metric, 0), 
                              reverse=True)
        return sorted_items[:top_n]

    def get_cra_activity(self, start_date=None, end_date=None, user_filter=None):
        """Wrapper for SDVManager get_cra_performance."""
        return self.sdv_mgr.get_cra_performance(start_date, end_date, user_filter)

    def get_cra_kpi(self, start_date=None, end_date=None, user_filter=None):
        """Wrapper for SDVManager get_cra_kpi."""
        return self.sdv_mgr.get_cra_kpi(start_date, end_date, user_filter)

