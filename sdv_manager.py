"""
SDV Manager - Remastered for Modular Export File
=================================================
Loads SDV status from the Modular export file and provides direct field lookup.

CRA_CONTROL_STATUS mapping:
- 0 (Blank) + Hidden=1: "hidden" (display ~)
- 0 (Blank) + Hidden=0: "pending" (display red !)
- 2 (Verified): "verified" (display green ✓)
- 3 (AwaitingReVerification): "awaiting" (display yellow ?)
- 4 (AutoVerified): "verified" (display green ✓)
"""

import os
import logging
import pandas as pd
from typing import Optional, Tuple, Dict, Callable

# Configure logging with file output
logger = logging.getLogger("SDVManager")
logger.setLevel(logging.DEBUG)

# File handler for detailed debugging
log_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), "sdv_debug.log")
file_handler = logging.FileHandler(log_file, mode='a', encoding='utf-8')
file_handler.setLevel(logging.DEBUG)
file_handler.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))
logger.addHandler(file_handler)

# Status constants
STATUS_HIDDEN = "hidden"          # ~ (debug mark for hidden fields)
STATUS_NOT_SENT = "not_sent"      # Not sent (Empty in EDC)
STATUS_NOT_CHECKED = "not_checked" # Red ! (Filled but not verified)
STATUS_VERIFIED = "verified"      # Green ✓
STATUS_AUTO_VERIFIED = "auto_verified" # Blue check ☑️
STATUS_AWAITING = "awaiting"      # Yellow ? (awaiting re-verification)
STATUS_NONE = None



class SDVManager:
    """Manages SDV status lookup using the Modular export file."""
    
    def __init__(self):
        self.modular_data: Optional[pd.DataFrame] = None
        self.patient_index: Dict[str, pd.DataFrame] = {}  # Pre-indexed by patient
        self.form_entry_status: Dict[str, str] = {}  # patient_form -> most recent Data Entry Status
        self.verification_metadata: Dict[str, tuple] = {} # Key -> (User, Date)
        self.all_history_df: Optional[pd.DataFrame] = None
        self.file_path: Optional[str] = None
        logger.info("SDVManager initialized (Modular mode)")
    
    def load_modular_file(self, filepath: str, progress_callback: Optional[Callable] = None) -> bool:
        """Load the Modular export file and build patient index.
        
        Args:
            filepath: Path to the Modular Excel file
            progress_callback: Optional callable(stage_text) for progress updates
        """
        def update_progress(stage):
            if progress_callback:
                progress_callback(stage)
        
        try:
            logger.info(f"Loading Modular file: {filepath}")
            update_progress("Reading file...")
            
            # Load Export Data sheet using calamine engine (much faster)
            self.modular_data = pd.read_excel(
                filepath, 
                sheet_name='Export Data',
                engine='calamine',  # Fast Rust-based Excel reader
                dtype={
                    'Subject Screening #': str,
                    'Variable name': str,
                    'Variable Value': str,  # Loaded to check if field is filled
                    'CRA_CONTROL_STATUS': int,
                    'Hidden': int,
                    'Table row #': 'Int64'  # Nullable int
                }
            )
            self.file_path = filepath
            
            update_progress("Processing data...")
            
            # Normalize patient IDs
            self.modular_data['Subject Screening #'] = (
                self.modular_data['Subject Screening #']
                .astype(str)
                .str.strip()
                .str.replace('.0', '', regex=False)
            )
            
            update_progress("Building index...")
            
            # Build patient index for fast lookups
            self._build_patient_index()
            
            logger.info(f"Loaded {len(self.modular_data)} rows for {len(self.patient_index)} patients")
            return True
            
        except Exception as e:
            logger.error(f"Error loading Modular file: {e}")
            return False
    
    def load_crf_status_file(self, filepath: str, progress_callback: Optional[Callable] = None) -> bool:
        """Load CrfStatusHistory file to track form submission status.
        
        Forms with most recent Data Entry Status = 'Created' are not yet submitted to CRA.
        """
        def update_progress(stage):
            if progress_callback:
                progress_callback(stage)
        
        try:
            logger.info(f"Loading CrfStatusHistory: {filepath}")
            update_progress("Loading form status...")
            
            # Load Export sheet - First try to find the header row
            # The file may have metadata rows at the top
            
            # 1. Read first few rows without header to scan
            df_scan = pd.read_excel(filepath, sheet_name='Export', header=None, nrows=50, engine='openpyxl')
            
            header_row_idx = 0
            found_header = False
            
            # Scan for a row that contains our target columns
            target_cols = ['Scr #', 'Subject', 'Subject Screening #', 'Activity', 'Event', 'Visit']
            
            for i, row in df_scan.iterrows():
                row_vals = [str(v).strip() for v in row.values]
                # Check if this row looks like a header
                matches = [col for col in target_cols if col in row_vals]
                if len(matches) >= 2: # At least 2 matches to be sure
                    header_row_idx = i
                    found_header = True
                    logger.info(f"Found header at row {i}: {matches}")
                    break
            
            if not found_header:
                logger.warning("Could not identify header row in CrfStatusHistory, trying default (0)")
                header_row_idx = 0

            # 2. Reload with correct header
            df = pd.read_excel(filepath, sheet_name='Export', header=header_row_idx, engine='openpyxl')
            
            # Rename 'Subject Screening #' or 'Subject' to 'Scr #' if needed
            if 'Scr #' not in df.columns:
                if 'Subject Screening #' in df.columns:
                    df = df.rename(columns={'Subject Screening #': 'Scr #'})
                elif 'Subject' in df.columns:
                    df = df.rename(columns={'Subject': 'Scr #'})

            # Normalize patient IDs, Form, and Activity
            if 'Scr #' in df.columns:
                df['Scr #'] = df['Scr #'].astype(str).str.strip()
            
            if 'Form' in df.columns:
                df['Form'] = df['Form'].astype(str).str.strip()
            
            if 'Activity' in df.columns:
                df['Activity'] = df['Activity'].astype(str).str.strip()
            elif 'Visit' in df.columns:
                df['Activity'] = df['Visit'].astype(str).str.strip()
            
            # Create datetime for sorting - handle various date/time formats
            try:
                df['DateTime'] = pd.to_datetime(
                    df['Date'].astype(str) + ' ' + df['Time'].astype(str), 
                    format='%d-%b-%Y %H:%M:%S (UTC)',
                    errors='coerce'
                )
            except Exception:
                df['DateTime'] = pd.to_datetime(
                    df['Date'].astype(str) + ' ' + df['Time'].astype(str), 
                    errors='coerce'
                )
            
            # Drop rows with invalid datetime and get most recent per patient+activity+form+repeat
            df_valid = df.dropna(subset=['DateTime'])
            self.all_history_df = df_valid.copy()
            
            # Normalize repeat number (empty/NaN becomes '0')
            df_valid['Repeat'] = df_valid['Repeatable form #'].fillna('0').astype(str).str.strip()
            df_valid['Repeat'] = df_valid['Repeat'].replace(['', 'nan', 'None'], '0')
            
            # Build form status index AND verification metadata index
            self.form_entry_status = {}
            self.verification_metadata = {} # Stores (User, Date) of the specific Verification action
            
            # Group by key for processing
            grouped = df_valid.groupby(['Scr #', 'Activity', 'Form', 'Repeat'])
            
            for name, group in grouped:
                # name is tuple: (Scr #, Activity, Form, Repeat)
                # 1. Current Status: Most recent row
                last_row = group.loc[group['DateTime'].idxmax()]
                
                repeat_num = str(name[3]) if name[3] else '0'
                if repeat_num.endswith('.0'): repeat_num = repeat_num[:-2]
                key = f"{name[0]}|{name[1]}|{name[2]}|{repeat_num}"
                
                ver_status = str(last_row['Verification Status']).strip()
                user = str(last_row.get('User', '')).strip()
                date_time = str(last_row['DateTime'])
                self.form_entry_status[key] = (last_row['Data Entry Status'], ver_status, user, date_time)
                
                # 2. Verification Metadata: Specific Verification Action
                # Find row where Ver Status is Verified/Re-verified BUT Appr Status is NOT Approved
                # This isolates the Verification event from subsequent Approval events
                ver_keywords = ['Verified', 'Verified by a single action', 'Re-verified', 'Re-verified by a single action']
                appr_keywords = ['Approved', 'AutoApproved', 'Approved by a single action', 'Locked']
                
                def is_ver(s): 
                    s = str(s)
                    if 'NotYetVerified' in s: return False
                    return any(k in s for k in ver_keywords)
                def is_appr(s): return any(k in str(s) for k in appr_keywords)
                
                # Filter group using state-change detection
                # Sort by DateTime to track sequence
                group = group.sort_values('DateTime')
                
                ver_rows = []
                prev_is_ver = False
                
                for _, row in group.iterrows():
                    curr_status = str(row['Verification Status'])
                    curr_is_ver = is_ver(curr_status)
                    
                    # Identify transition to Verified state
                    if curr_is_ver and not prev_is_ver:
                        ver_rows.append(row)
                    
                    # Update previous state
                    # Note: We track the boolean state. 
                    # If status becomes Unverified (False), we record that, so next Verified is a new event.
                    prev_is_ver = curr_is_ver
                
                if ver_rows:
                    # Take the most recent verification transition
                    ver_row = ver_rows[-1]
                    v_user = str(ver_row.get('User', '')).strip()
                    v_date = str(ver_row['DateTime'])
                    self.verification_metadata[key] = (v_user, v_date)
            
            logger.info(f"Loaded form status for {len(self.form_entry_status)} patient-form combinations")
            
            # Count forms with 'Created' status (not submitted)
            # Only count as 'Created' if Verification Status is Blank
            created_count = sum(1 for s, v in self.form_entry_status.values() 
                              if s == 'Created' and v in ['Blank', 'nan', 'None', ''])
            logger.info(f"Forms not yet submitted (Created + Blank Verification): {created_count}")
            
            return True
            
        except Exception as e:
            logger.error(f"Error loading CrfStatusHistory: {e}")
            return False
    
    def get_form_status(self, patient_id: str, form_name: str, visit_name: str = None, repeat_number: str = None) -> str:
        """Get form-level Data Entry Status.
        
        Args:
            patient_id: Patient ID
            form_name: Form name (exact match required)
            visit_name: Optional visit/activity name (fuzzy match allowed)
            repeat_number: Optional repeat number for repeating forms (e.g. "10" for AE #10)
            
        Returns 'Created' ONLY if form is Created AND Verification Status is Blank.
        Otherwise return 'EntryCompleted'.
        """
        patient_id = str(patient_id).strip()
        form_name_lower = str(form_name).strip().lower()
        visit_name_lower = str(visit_name).strip().lower() if visit_name else ""
        repeat_str = str(repeat_number).strip() if repeat_number else "0"
        
        # Iterate through keys to find match
        # Key format: patient|visit|form|repeat
        prefix = f"{patient_id}|"
        
        for full_key, status_tuple in self.form_entry_status.items():
            status = status_tuple[0]
            ver_status = status_tuple[1]
            if full_key.startswith(prefix):
                parts = full_key.split('|')
                if len(parts) >= 4:
                    k_visit = parts[1]
                    k_form = parts[2]
                    k_repeat = parts[3]
                    
                    # Check form match (Exact or Substring - Bidirectional)
                    form_match = False
                    if k_form.lower() == form_name_lower:
                        form_match = True
                    elif form_name_lower in k_form.lower():
                        form_match = True
                    elif k_form.lower() in form_name_lower:
                        form_match = True
                    
                    if form_match:
                        # Check repeat number match (if not provided, match any)
                        repeat_match = (k_repeat == repeat_str or repeat_str == "0")
                        if not repeat_match:
                            continue
                        # Check visit match if provided
                        # Visit matching needs to be flexible (e.g. "Screening/Baseline" vs "Screening")
                        visit_match = True
                        if visit_name:
                            # 1. Exact match (case insensitive)
                            if k_visit.lower() == visit_name_lower:
                                visit_match = True
                            # 2. Key contains search visit (e.g. k="Screening/Baseline", search="Screening")
                            elif visit_name_lower in k_visit.lower():
                                visit_match = True
                            # 3. Search visit contains key
                            elif k_visit.lower() in visit_name_lower:
                                visit_match = True
                            else:
                                visit_match = False
                        
                        if visit_match:
                            # Logic: Return 'Created' only if truly not sent (ver_status is Blank)
                            # if status == 'Created' and ver_status not in ['Blank', 'nan', 'None', '']:
                            #     return status

                            
                            # If we are strictly matching, we should return whatever status we found
                            # But if verification logic applies, override it
                            if status == 'Created' and ver_status in ['Blank', 'nan', 'None', '']:
                                return status
        
    def get_verification_details(self, patient_id: str, form_name: str, visit_name: str = None, repeat_number: str = None) -> Optional[dict]:
        """Get verification metadata (User, Date) if available."""
        if not self.form_entry_status:
            return None
            
        patient_id = str(patient_id).strip()
        form_name = str(form_name).strip()
        
        # Reuse key matching logic from get_form_status
        form_name_lower = form_name.lower()
        visit_name_lower = str(visit_name).strip().lower() if visit_name else ""
        repeat_str = str(repeat_number).strip() if repeat_number else "0"
        
        prefix = f"{patient_id}|"
        
        for full_key, status_tuple in self.form_entry_status.items():
            if full_key.startswith(prefix):
                parts = full_key.split('|')
                if len(parts) >= 4:
                    k_visit = parts[1]
                    k_form = parts[2]
                    k_repeat = parts[3]
                    
                    # Check form match (Exact or Substring - Bidirectional)
                    form_match = False
                    if k_form.lower() == form_name_lower:
                        form_match = True
                    elif form_name_lower in k_form.lower():
                        form_match = True
                    elif k_form.lower() in form_name_lower:
                        form_match = True
                    
                    if form_match:
                        # Check repeat match
                        repeat_match = (k_repeat == repeat_str or repeat_str == "0")
                        if not repeat_match: continue
                        
                        visit_match = True
                        if visit_name:
                            # 1. Exact match (case insensitive)
                            if k_visit.lower() == visit_name_lower:
                                visit_match = True
                            # 2. Key contains search visit
                            elif visit_name_lower in k_visit.lower():
                                visit_match = True
                            # 3. Search visit contains key
                            elif k_visit.lower() in visit_name_lower:
                                visit_match = True
                            else:
                                visit_match = False
                        
                        if visit_match:
                            # Found match
                            # Check verification metadata for more accurate user/date (isolates Verification from Approval)
                            ver_meta = self.verification_metadata.get(full_key)
                            if ver_meta:
                                 return {
                                     'user': ver_meta[0],
                                     'date': ver_meta[1],
                                     'status': status_tuple[1]
                                 }
                            
                            # Fallback to current status row - ONLY if it looks verified
                            # Otherwise we risk showing the Data Entry user as the Verifier
                            current_ver_status = status_tuple[1]
                            strict_ver_keywords = ['Verified', 'SDV Verified', 'DMR Verified']
                            is_strictly_verified = any(k == current_ver_status for k in strict_ver_keywords)
                            # Or check containment if status is like "Verified by..."
                            if not is_strictly_verified and 'Verified' in current_ver_status and 'NotYetVerified' not in current_ver_status:
                                is_strictly_verified = True
                                
                            if len(status_tuple) >= 4 and is_strictly_verified:
                                return {
                                    'user': status_tuple[2],
                                    'date': status_tuple[3],
                                    'status': status_tuple[1]
                                }
                            return None
        return None


    
    def _build_patient_index(self):
        """Pre-index data by patient using triple approach: Variable name + Constructed Key + Field Key."""
        self.patient_index = {}
        
        df = self.modular_data.copy()
        
        # Prepare all columns
        df['var_name'] = df['Variable name'].fillna('').astype(str).str.strip()
        df['visit_code'] = df['Visit Code'].fillna('').astype(str).str.strip()
        df['form_code'] = df['Form Code'].fillna('').astype(str).str.strip()
        df['field_key'] = df['Field Key'].fillna('').astype(str).str.strip()
        df['status_code'] = df['CRA_CONTROL_STATUS'].fillna(0).astype(int)
        df['hidden'] = df['Hidden'].fillna(0).astype(int)
        df['table_row'] = df['Table row #'].fillna(0).astype(str).str.strip()
        # Use 'Repeatable form #' as fallback when 'Table row #' is empty (e.g., for AEs)
        df['repeat_form'] = df['Repeatable form #'].fillna('').astype(str).str.strip()
        # Remove .0 suffix from repeat numbers
        df['repeat_form'] = df['repeat_form'].apply(lambda x: x[:-2] if x.endswith('.0') else x)
        # Combine: use table_row if present, otherwise repeat_form
        df['effective_row'] = df.apply(lambda r: r['table_row'] if r['table_row'] not in ['', '0', 0] else r['repeat_form'], axis=1)
        df['has_value'] = df['Variable Value'].notnull() & (df['Variable Value'].astype(str).str.strip() != '')
        
        # Filter out rows without variable names
        df = df[df['var_name'] != '']
        
        # Build field suffix: strip visit prefix from var_name
        def get_field_suffix(row):
            if row['visit_code'] and row['var_name'].startswith(row['visit_code'] + '_'):
                return row['var_name'][len(row['visit_code']) + 1:]
            return row['var_name']
        
        df['field_suffix'] = df.apply(get_field_suffix, axis=1)
        
        # Build Constructed Key: Visit_Form_Suffix (matches TreeView ID logic)
        def build_constructed_key(row):
            if row['visit_code'] and row['form_code']:
                return f"{row['visit_code']}_{row['form_code']}_{row['field_suffix']}"
            elif row['visit_code']:
                return f"{row['visit_code']}_{row['field_suffix']}"
            return row['var_name']
            
        df['constructed_key'] = df.apply(build_constructed_key, axis=1)
        
        # Group by patient and create dictionaries with triple indexing
        for patient_id, group in df.groupby('Subject Screening #'):
            patient_dict = {}
            
            for _, row in group.iterrows():
                status_tuple = (row['status_code'], row['hidden'], row['has_value'])
                
                # 1. Index by Variable name (e.g., SBV_PEDTC, SBV_LBORRESUN_ALB)
                # Helper to update index with priority logic
                def update_index(key, tuple_val):
                    if key not in patient_dict:
                        patient_dict[key] = tuple_val
                        return
                    
                    # Priority logic:
                    # 1. Non-hidden over hidden
                    # 2. Has value over no value
                    old_stat, old_hidden, old_has_val = patient_dict[key]
                    new_stat, new_hidden, new_has_val = tuple_val
                    
                    if old_hidden == 1 and new_hidden == 0:
                        patient_dict[key] = tuple_val
                    elif old_hidden == new_hidden: # Both same visibility
                        if not old_has_val and new_has_val:
                            patient_dict[key] = tuple_val
                        elif old_has_val == new_has_val: # Both have or don't have value
                            # Keep most recent (overwrite)
                            patient_dict[key] = tuple_val

                if row['var_name']:
                    update_index(row['var_name'], status_tuple)
                    
                if row['constructed_key'] and row['constructed_key'] != row['var_name']:
                    update_index(row['constructed_key'], status_tuple)
                
                if row['field_key'] and '#' in row['field_key']:
                    update_index(row['field_key'], status_tuple)
            
            self.patient_index[str(patient_id)] = patient_dict


    
    def get_field_status(self, patient_id: str, field_id: str, table_row: int = None, form_name: str = None, visit_name: str = None) -> Optional[str]:
        """
        Get SDV status for a specific field.
        
        Uses form-level check first, then dual lookup strategy:
        1. Check form-level Data Entry Status (Created = not submitted)
        2. Try Variable name directly (for most fields)
        3. Try Field Key format for table rows (SBV/MH/MHTERM#1)
        
        Args:
            patient_id: Patient screening number
            field_id: Variable name from TreeView (e.g., SBV_MH_MHTERM or SBV_PE_PEDTC)
            table_row: Optional table row number for repeating forms
            form_name: Form name for form-level status check (e.g., "Vital signs")
            visit_name: Visit name for more precise form status check (e.g., "Screening")
            
        Returns:
            STATUS_NOT_SENT, STATUS_NOT_CHECKED, STATUS_VERIFIED, STATUS_AUTO_VERIFIED, STATUS_AWAITING, STATUS_HIDDEN, or STATUS_NONE
        """
        patient_id = str(patient_id).strip()
        field_id = str(field_id).strip()
        
        # Check form-level status first (if form_name provided)
        if form_name and self.form_entry_status:
            # For repeating forms like AE, pass table_row as repeat number
            repeat_num = str(table_row) if table_row else None
            form_status = self.get_form_status(patient_id, form_name, visit_name, repeat_number=repeat_num)
            if form_status == 'Created':
                # Form not yet submitted to CRA - all fields are "Not Sent"
                return STATUS_NOT_SENT
        
        if patient_id not in self.patient_index:
            return STATUS_NONE
        
        patient_data = self.patient_index[patient_id]
        
        # 0. If table row is provided, try that specific lookup FIRST
        if table_row is not None:
            # Format A: Suffix style (Standard) -> KEY#ROW
            # Convert SBV_MH_MHTERM -> SBV/MH/MHTERM
            field_key = field_id.replace('_', '/')
            key_with_row = f"{field_key}#{table_row}"
            
            if key_with_row in patient_data:
                return self._map_status(*patient_data[key_with_row], field_name=field_id)
            
            # Format B: Infix style (AE style) -> KEY_PART#ROW/KEY_PART
            # We don't know where the row number is, so we scan keys containing #{row}
            row_marker = f"#{table_row}"
            for key in patient_data:
                if row_marker in key:
                    # Remove the row marker and normalize slashes/underscores to comparing
                    # Example: LOGS/AE#1/AETERM -> LOGS/AE/AETERM
                    key_clean = key.replace(row_marker, "")
                    # Example: LOGS_AE_AETERM -> LOGS/AE/AETERM
                    field_key_clean = field_key # already slashed
                    
                    # Also handle if row marker replaced a slash or was inserted?
                    # Usually it's inserted like LOGS/AE#1/AETERM -> LOGS/AE/AETERM (Implicit slash handling)
                    # Let's try direct comparison
                    if key_clean == field_key:
                        return self._map_status(*patient_data[key], field_name=field_id)
                    
                    # Fallback: Suffix Match (for complex keys like LOGS/LB_PR_OTH/LBTEST_OTH#1 matching LOGS_LBTEST_OTH)
                    # Try to match the last part of the variable name
                    # Remove common prefixes from field_id: LOGS_, SBV_, etc.
                    parts = field_id.split('_')
                    if len(parts) > 1:
                        suffix = parts[-1] # Try exact suffix match e.g. "LBTEST_OTH" ? No, split splits all.
                        # Actually "LBTEST_OTH" splits to "LBTEST", "OTH".
                        # Let's try matching the stripped field_id against the end of key_clean (normalized)
                        
                        # Normalize key_clean to underscore for comparison?
                        key_norm = key_clean.replace('/', '_')
                        if field_id in key_norm or key_norm.endswith(field_id) or (len(parts) >= 2 and key_norm.endswith('_'.join(parts[1:]))):
                             return self._map_status(*patient_data[key], field_name=field_id)
        
        # 1. Try direct Variable name match (works for non-table fields or generic loopkup)
        if field_id in patient_data:
            return self._map_status(*patient_data[field_id], field_name=field_id)
        
        # 2. For table fields without specific row, try Field Key format
        field_key = field_id.replace('_', '/')
        
        # Try finding any matching table row (when row not specified)
        for key in patient_data:
            if key.startswith(field_key + '#'):
                return self._map_status(*patient_data[key], field_name=field_id)
        
        return STATUS_NONE


    
    def get_ae_repeat_number(self, patient_id: str, ae_term: str, match_index: int = 0) -> Optional[str]:
        """Find the Repeatable form # for a given AE term string.
        
        Args:
            patient_id: Patient ID
            ae_term: The Adverse Event term as displayed in the UI
            match_index: 0-based index if multiple AEs have the same term
            
        Returns:
            The actual Repeatable form # (e.g. "10") or None if not found.
        """
        if self.modular_data is None:
            return None
            
        try:
            ae_term = str(ae_term).strip()
            df = self.modular_data
            
            # Filter efficiently
            mask = (df['Subject Screening #'] == str(patient_id)) & \
                   (df['Form Code'] == 'AE') & \
                   (df['Variable name'].str.contains('AETERM', na=False)) & \
                   (~df['Variable name'].str.contains('COMM', na=False)) & \
                   (df['Variable Value'].astype(str).str.strip() == ae_term)
                   
            rows = df[mask]
            
            if not rows.empty and len(rows) > match_index:
                # Use Repeatable form # column
                # Note: We verified this column contains the headers like 1, 10, etc.
                val = str(rows.iloc[match_index]['Repeatable form #'])
                # Clean .0 float suffix if present
                if val.endswith('.0'): 
                    val = val[:-2]
                return val.strip()
                
        except Exception as e:
            logger.error(f"Error resolving AE repeat number: {e}")
            
        return None

    def get_lab_row_number(self, patient_id: str, lab_test: str, match_index: int = 0) -> Optional[str]:
        """Find the Table row # for a given Lab Test (Additional Labs).
        
        Args:
            patient_id: Patient ID
            lab_test: The Lab Test name
            match_index: 0-based index if multiple tests have the same name
            
        Returns:
            The actual Table row # or None.
        """
        if self.modular_data is None:
            return None
            
        try:
            lab_test = str(lab_test).strip()
            df = self.modular_data
            
            # Filter for LB_PR_OTH form
            mask = (df['Subject Screening #'] == str(patient_id)) & \
                   (df['Form Code'] == 'LB_PR_OTH') & \
                   (df['Variable name'].str.contains('TEST', na=False)) & \
                   (df['Variable Value'].astype(str).str.strip() == lab_test)
                   
            rows = df[mask]
            
            if not rows.empty and len(rows) > match_index:
                # Use Table row # column for Labs
                val = str(rows.iloc[match_index]['Table row #'])
                # Clean .0 float suffix
                if val.endswith('.0'): 
                    val = val[:-2]
                return val.strip()
                
        except Exception as e:
            logger.error(f"Error resolving Lab row number: {e}")
            
        return None



    def _map_status(self, status_code: int, hidden: int, has_value: bool = True, field_name: str = "") -> str:
        """Map CRA_CONTROL_STATUS code to status string."""
        # Checkbox patterns - empty value means "unchecked", not "not sent"
        # Substring patterns for common checkbox fields
        checkbox_substr = ["ONGO", "OCCUR", "AEACN", "AESAE", "YN"]
        # Suffix patterns — use endswith to avoid false positives
        # (e.g. "_LT" would match "ALT" lab field, "_PR" matches many non-checkbox fields)
        checkbox_suffix = ["_LTFL", "_PRFL"]
        is_checkbox = (
            any(pat in field_name for pat in checkbox_substr) or
            any(field_name.endswith(pat) for pat in checkbox_suffix)
        )
        
        if status_code == 0:  # Blank
            if hidden == 1:
                return STATUS_HIDDEN
            elif not has_value and not is_checkbox:
                return STATUS_NOT_SENT      # Value missing = Not Sent (but not for checkboxes)
            else:
                return STATUS_NOT_CHECKED   # Value exists or checkbox = Not Checked (Red !)
        elif status_code == 2:  # Verified (Manual)
            return STATUS_VERIFIED
        elif status_code == 3:  # AwaitingReVerification
            return STATUS_AWAITING
        elif status_code == 4:  # Auto Verified (Blue checkmark in EDC)
            return STATUS_AUTO_VERIFIED
        else:
            return STATUS_NONE
    
    def get_patient_stats(self, patient_id: str) -> Tuple[int, int, int, int]:
        """
        Get SDV statistics for a patient.
        
        Returns:
            Tuple of (verified_count, pending_count, awaiting_count, hidden_count)
        """
        patient_id = str(patient_id).strip()
        
        if patient_id not in self.patient_index:
            return (0, 0, 0, 0)
        
        verified = 0
        pending = 0
        awaiting = 0
        hidden = 0
        
        for status_data in self.patient_index[patient_id].values():
            status = self._map_status(*status_data)
            if status in [STATUS_VERIFIED, STATUS_AUTO_VERIFIED]:
                verified += 1
            elif status in [STATUS_NOT_CHECKED, STATUS_NOT_SENT]:
                pending += 1
            elif status == STATUS_AWAITING:
                awaiting += 1
            elif status == STATUS_HIDDEN:
                hidden += 1
        
        return (verified, pending, awaiting, hidden)
    
    def get_total_stats(self) -> Tuple[int, int, int, int]:
        """Get total SDV statistics across all patients."""
        total_verified = 0
        total_pending = 0
        total_awaiting = 0
        total_hidden = 0
        
        for patient_id in self.patient_index:
            v, p, a, h = self.get_patient_stats(patient_id)
            total_verified += v
            total_pending += p
            total_awaiting += a
            total_hidden += h
        
        return (total_verified, total_pending, total_awaiting, total_hidden)
    
    def is_loaded(self) -> bool:
        """Check if Modular file is loaded."""
        return self.modular_data is not None and len(self.patient_index) > 0

    def get_cra_performance(self, start_date=None, end_date=None, user_filter=None):
        """Analyze CRA verification activity within a date range."""
        if self.all_history_df is None:
            return pd.DataFrame()
            
        df = self.all_history_df.copy()
        
        # 1. Filter for Verification Events
        ver_keywords = ['Verified', 'Verified by a single action', 'Re-verified', 'Re-verified by a single action']
        df = df[df['Verification Status'].isin(ver_keywords)]
        
        # 2. Date Filtering
        if start_date:
            try:
                s_dt = pd.to_datetime(start_date)
                df = df[df['DateTime'] >= s_dt]
            except: pass
        if end_date:
            try:
                e_dt = pd.to_datetime(end_date) + pd.Timedelta(days=1)
                df = df[df['DateTime'] < e_dt]
            except: pass
            
        # 3. User Filtering
        if user_filter and user_filter != "All":
            df = df[df['User'] == user_filter]
            
        if df.empty:
            return pd.DataFrame()
            
        # 4. Group by User, Day, Site, Patient, Activity, Form
        df['Day'] = df['DateTime'].dt.strftime('%Y-%m-%d')
        site_col = 'Site #' if 'Site #' in df.columns else 'Site'
        
        # Unique Pages (Forms) verified per Patient/Visit/Day
        performance = df.groupby(['User', 'Day', site_col, 'Scr #', 'Activity', 'Form']).size().reset_index()
        summary = performance.groupby(['User', 'Day', site_col, 'Scr #', 'Activity']).size().reset_index()
        summary.columns = ['User', 'Date', 'Site', 'Patient', 'Visit', 'Pages Verified']
        
        return summary



# Test code
if __name__ == "__main__":
    import glob
    
    # Find latest Modular file
    modular_files = glob.glob('verified/*Modular*.xlsx')
    if modular_files:
        latest = max(modular_files, key=os.path.getmtime)
        print(f"Testing with: {latest}")
        
        manager = SDVManager()
        if manager.load_modular_file(latest):
            print(f"\nPatients loaded: {len(manager.patient_index)}")
            
            # Test patient 206-06
            patient = '206-06'
            stats = manager.get_patient_stats(patient)
            print(f"\nPatient {patient} stats:")
            print(f"  Verified: {stats[0]}")
            print(f"  Pending: {stats[1]}")
            print(f"  Awaiting: {stats[2]}")
            print(f"  Hidden: {stats[3]}")
            
            # Test field lookup
            test_fields = ['SBV_PE_PEDTC', 'SBV_DM_BRTHDAT', 'SBV_VS_VSDAT']
            print(f"\nField status tests:")
            for field in test_fields:
                status = manager.get_field_status(patient, field)
                print(f"  {field}: {status}")
            
            # Total stats
            total = manager.get_total_stats()
            print(f"\nTotal stats:")
            print(f"  Verified: {total[0]}")
            print(f"  Pending: {total[1]}")
            print(f"  Awaiting: {total[2]}")
            print(f"  Hidden: {total[3]}")
    else:
        print("No Modular file found in verified/ folder")
