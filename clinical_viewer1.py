
import tkinter as tk
import threading
import logging
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import numpy as np
import glob
import os
import re
from datetime import datetime
from dashboard_manager import DashboardManager
from dashboard_ui import DashboardWindow

import math
import time
from echo_export import EchoExporter, VISIT_ORDER as ECHO_VISITS
from labs_export import LabsExporter
from fu_highlights_export import FUHighlightsExporter
from cvc_export import CVCExporter
from procedure_timing_export import ProcedureTimingExporter
from sdv_manager import SDVManager
from hf_hospitalization_manager import HFHospitalizationManager
from ae_manager import AEManager
from ae_ui import AEWindow
from hf_ui import HFWindow
from assessment_data_table import AssessmentDataExtractor, ASSESSMENT_CATEGORIES
from assessment_table_ui import AssessmentTableWindow
from procedure_timing_ui import ProcedureTimingWindow
from export_dialogs_ui import EchoExportDialog, CVCExportDialog, LabsExportDialog, FUHighlightsDialog
from data_loader import (
    detect_latest_project_file, load_project_file, parse_cutoff_from_filename,
    validate_cross_form,
)

# Module-level logger
logger = logging.getLogger("ClinicalViewer")


# Domain rules imported from centralized config

from config import VISIT_MAP, ASSESSMENT_RULES, CONDITIONAL_SKIPS, VISIT_SCHEDULE
import batch_export
import data_comparator
from data_sources import DataSourceManager, DataSourcesWindow
from patient_timeline import PatientTimelineWindow
from gap_analysis import DataGapsWindow
from visit_schedule_ui import VisitScheduleWindow
from matrix_display import MatrixDisplay
from view_builder import ViewBuilder
from toolbar_setup import setup_toolbar



class ClinicalDataMasterV30:
    def __init__(self, root):
        self.root = root
        self.root.title("Clinical Data Master v30 (AE Column)")
        self.root.geometry("1600x900")

        self.df_main = None
        self.df_ae = None
        self.df_cm = None
        self.df_cvh = None
        self.df_act = None
        self.labels = {}
        self.ae_lookup = {}
        self.current_file_path = None
        self.current_patient_gaps = []
        self.current_tree_data = {}
        
        # SDV (Source Data Verification) Manager
        self.sdv_manager = None
        self.dashboard_manager = None
        self.hf_manager = None  # HF Hospitalization tracking
        self.sdv_verified_fields = set()  # Set of verified field IDs for current patient
        
        # View Builder
        self.view_builder = ViewBuilder(self)

        # Matrix Display (specialized table windows)
        self.matrix_display = MatrixDisplay(self)

        # Data Sources manager
        self.data_source_manager = DataSourceManager(os.path.dirname(os.path.abspath(__file__)))
        
        self._setup_ui()
        
        # Auto-load latest file
        self.root.after(500, self.find_and_load_latest)

    @staticmethod
    def _clean_id(val):
        """Safely strip trailing .0 from numeric IDs loaded from Excel.
        
        Handles edge cases like '10.05' that naive str.replace('.0','') would corrupt.
        """
        s = str(val).strip()
        if s.endswith('.0'):
            return s[:-2]
        return s

    def _setup_ui(self):
        """Setup UI using external module."""
        setup_toolbar(self, self.root)
        
        # --- Treeview Setup ---
        # Container frame for tree and scrollbars
        tree_frame = tk.Frame(self.root)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        columns = ("Value", "Status", "User", "Date", "Code")
        self.tree = ttk.Treeview(tree_frame, columns=columns, selectmode="browse")
        
        self.tree.heading("#0", text="  Hierarchy / Parameter", anchor="w")
        self.tree.heading("Value", text="Result / Response", anchor="w")
        self.tree.heading("Status", text="SDV Status", anchor="center")
        self.tree.heading("User", text="Verified By", anchor="w")
        self.tree.heading("Date", text="Date", anchor="w")
        self.tree.heading("Code", text="Variable / Code", anchor="w")

        self.tree.column("#0", width=400)
        self.tree.column("Value", width=300)
        self.tree.column("Status", width=100)
        self.tree.column("User", width=120)
        self.tree.column("Date", width=120)
        self.tree.column("Code", width=200)

        # Scrollbars
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        # Grid layout for tree and scrollbars
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Events
        self.tree.bind("<Double-1>", self.on_double_click)
        self.tree.bind("<Button-3>", self.show_context_menu)
        
        # Tags for SDV coloring
        self.tree.tag_configure('verified', foreground='#27ae60') # Green
        self.tree.tag_configure('pending', foreground='#e67e22')  # Orange
        self.tree.tag_configure('patient', font=("Segoe UI", 9, "bold"), background="#ecf0f1")

    def find_and_load_latest(self):
        """Find the most recent project file and load it."""
        try:
            result = detect_latest_project_file(os.getcwd())
            if result:
                full_path, cutoff_time = result
                self.load_data(full_path, cutoff_time)
        except Exception as e:
            logger.error("Auto-load failed: %s", e)

    def load_data(self, path, cutoff_time=None):
        """Load data from specific path using data_loader module."""
        try:
            self.root.config(cursor="watch")
            self.root.update()

            # Delegate to data_loader (pure data, no UI)
            result = load_project_file(path, cutoff_time)

            # Assign data
            self.current_file_path = result.file_path
            self.df_main = result.df_main
            self.df_ae = result.df_ae
            self.df_cm = result.df_cm
            self.df_cvh = result.df_cvh
            self.df_act = result.df_act
            self.labels = result.labels

            # Update UI labels
            filename = os.path.basename(path)
            self.file_info_var.set(f"Loaded: {filename}")
            if result.cutoff_time:
                self.cutoff_var.set(f"Cutoff: {result.cutoff_time.strftime('%Y-%m-%d %H:%M:%S')}")
            else:
                self.cutoff_var.set("")

            # Log any load warnings
            for w in result.warnings:
                logger.warning(w)

            # Run cross-form consistency checks
            xform_issues = validate_cross_form(result)
            for issue in xform_issues:
                logger.warning("Cross-form: %s", issue)

            self._populate_filters()
            self.view_builder.clear_cache()
            self.generate_view()
            self.root.config(cursor="")

            # Auto-load SDV data after main data is loaded
            self.load_sdv_data()

            # Register with Data Sources manager
            self.data_source_manager.register_loaded_file("project", path)

            # Initialize HF Hospitalization Manager
            self.hf_manager = HFHospitalizationManager(self.df_main, self.df_ae)

            # Initialize AE Manager
            self.ae_manager = AEManager(self.df_main, self.df_ae)

        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror("Error", f"Failed to load file: {str(e)}")

    # _load_extra_sheets() has been moved to data_loader.py


    def _add_filter(self, parent, text):
        tk.Label(parent, text=text).pack(side=tk.LEFT, padx=5)
        cb = ttk.Combobox(parent, state="readonly", width=15)
        cb.pack(side=tk.LEFT, padx=5)
        return cb

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if not path: return
        self.load_data(path)

    def update_patients(self, event):
        s = self.cb_site.get()
        if self.df_main is not None:
            subset = self.df_main[self.df_main['Site #'].astype(str).apply(self._clean_id) == s]
            self.cb_pat['values'] = sorted(list(subset['Screening #'].dropna().unique()))
            if self.cb_pat['values']: self.cb_pat.current(0)

    def identify_column(self, col_name):
        visit_name, visit_prefix = None, None
        for code, human in VISIT_MAP.items():
            if col_name.startswith(code + "_"):
                visit_name, visit_prefix = human, code
                break

        if not visit_name: return None, None, None, None

        cat, assess = "Uncategorized", "Uncategorized"
        for pattern, category, name in ASSESSMENT_RULES:
            if re.search(pattern, col_name):
                cat, assess = category, name
                break 
        
        dp_key = col_name.replace(visit_prefix + "_", "")
        return visit_name, cat, assess, dp_key

    def _is_matrix_supported_col(self, col_name):
        """Check if a column should be included in the longitudinal Data Matrix."""
        # Generic Result patterns
        is_result = any(x in col_name for x in ["ORRES", "_RES", "_VAL", "PRORRES", "eGFR"])
        
        # Excluded meta-data patterns
        is_excluded = any(x in col_name for x in ["ORRESU", "ORRESSU", "STAT", "PERF", "NAM", "COMM"])
        
        # Specialized Clinical History/Event patterns
        # Include Procedure Timings in the matrix
        is_clinical_event = any(x in col_name for x in ["_PTHME_", "_HFH_", "_PR_TIM_"])
        
        # ACT Lab columns are handled specially (loaded from separate sheet)
        if "_LB_ACT_" in col_name:
            return True
            
        return (is_result or is_clinical_event) and not is_excluded

    def annotate_procedure_timing(self, label, col_name):
        """Add pre/post-procedure annotations for Procedure Timing fields."""
        if "_PR_TIM_" not in col_name:
            return label
            
        if col_name.endswith("_POST"):
            return f"{label} (post-procedure)"
        elif col_name.endswith("_NR"):
            return label
        else:
            # If it doesn't end in POST but there exists a POST version, it's pre
            post_col = col_name + "_POST"
            if self.df_main is not None and post_col in self.df_main.columns:
                return f"{label} (pre-procedure)"
        return label

    # _validate_schema() has been moved to data_loader.validate_schema()

    def _populate_filters(self):
        """Populate filter dropdowns based on loaded data."""
        if self.df_main is not None and not self.df_main.empty:
            # Populate Site Combobox
            if 'Site #' in self.df_main.columns:
                sites = sorted(list(self.df_main['Site #'].astype(str).apply(self._clean_id).dropna().unique()))
                self.cb_site['values'] = sites
                if sites:
                    self.cb_site.current(0)
                    self.update_patients(None)

    def clean_label(self, label):
        txt = str(label).strip()
        # Remove _x0009_ prefix (tab character encoding) from Echo parameters
        txt = txt.replace("_x0009_", "")
        for prefix in ["Sponsor/", "Sponsor ", "Core Lab/", "Core Lab "]:
            if txt.startswith(prefix): txt = txt[len(prefix):]
            
        # Specific PTHME / HFH mapping for matrix/tree clarity
        pthme_map = {
            "post-treatment hospitalizations and medical event / status": "Hospitalization Occurred?",
            "date of hospitalization/event": "Event Date",
            "reason for hospitalization/event": "Reason",
            "hospitalization / source of report": "Source of Report",
            "cardivascular /details": "Details (CV)",
            "non-cardiovascular / describe": "Details (Non-CV)",
            "source documents / status": "Source Docs Available?",
            "occurrence of heart failure hospitalization": "HF Hospitalization?",
        }
        low_txt = txt.lower()
        if low_txt in pthme_map:
            return pthme_map[low_txt]
            
        return txt

    def is_not_done_column(self, col_name, val, label):
        val = str(val).lower().strip()
        col_name = str(col_name).lower()
        label = str(label).lower()
        if val in ["not done", "not performed"]: return True
        if val in ["no", "n", "0"]:
            if any(x in col_name for x in ["perf", "compl", "done", "prfrm", "stat"]):
                return True
        if val in ["true", "yes", "checked", "1", "y"]:
            if "not" in col_name and "done" in col_name: return True
            if "not done" in label or "not performed" in label: return True
            # Physical Exam sub-statuses: SBV_PE_PESTAT_HEAD=True means Not Done
            # We exclude the main 'PESTAT' (Performed=Yes/No) by checking for underscore suffix
            if "pestat_" in col_name: return True
        return False

    def build_ae_lookup(self, patient_id):
        """Build AE lookup: maps (test_row_num, ref_type) -> (ae_num, ae_term)"""
        self.ae_lookup = {}
        if self.df_ae is None: return
        
        pat_aes = self.df_ae[self.df_ae['Screening #'].str.contains(patient_id.replace('-', '-'), na=False)]
        
        for _, ae_row in pat_aes.iterrows():
            ae_num = ae_row.get('Template number', '')
            ae_term = ae_row.get('LOGS_AE_AETERM', '')
            
            if pd.isna(ae_term) or not str(ae_term).strip() or str(ae_term) == 'nan':
                ae_term = None
            
            lbref = ae_row.get('LOGS_AE_LBREF', '')
            prref = ae_row.get('LOGS_AE_PRREF', '')
            
            for ref_val, ref_type in [(lbref, 'LB'), (prref, 'PR')]:
                if pd.isna(ref_val) or not ref_val: continue
                ref_str = str(ref_val).strip()
                if ref_str.startswith('#') and '/' in ref_str:
                    match = re.match(r'#(\d+)', ref_str)
                    if match:
                        test_row_num = match.group(1)
                        parent_ae = pat_aes[(pat_aes['Template number'] == ae_num) & 
                                           (pat_aes['LOGS_AE_AETERM'].notna()) & 
                                           (pat_aes['LOGS_AE_AETERM'] != 'nan')]
                        if not parent_ae.empty:
                            parent_term = parent_ae.iloc[0].get('LOGS_AE_AETERM', '')
                            key = (test_row_num, ref_type)
                            if key not in self.ae_lookup:
                                self.ae_lookup[key] = (ae_num, parent_term)

    def get_ae_info(self, test_row_num, ref_type='PR'):
        key = (str(test_row_num), ref_type)
        return self.ae_lookup.get(key, (None, None))

    def parse_timeline_entry(self, entry_str):
        if not entry_str or "/" not in entry_str: return None
        parts = [p.strip() for p in entry_str.split('/')]
        if len(parts) < 4: return None
        row_num = parts[0].replace('#', '') if parts[0].startswith("#") else ""
        date_str = parts[1]
        test_name = parts[2]
        val = parts[3]
        unit = parts[4] if len(parts) > 4 else ""
        date_clean = date_str.replace("T", " ")
        if "," in date_clean: date_clean = date_clean.split(",")[0]
        full_val = f"{val} {unit}".strip()
        return row_num, date_clean, test_name, full_val

    def _get_all_descendants(self, item):
        """Recursively get all descendant items (leaves) of a tree item."""
        children = self.tree.get_children(item)
        descendants = list(children)
        for child in children:
            descendants.extend(self._get_all_descendants(child))
        return descendants

    def show_data_matrix(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select at least one row.")
            return

        site, pat = self.cb_site.get(), self.cb_pat.get()
        if not site or not pat or self.df_main is None: 
            messagebox.showwarning("Warning", "Patient data not loaded.")
            return

        mask = (self.df_main['Site #'].astype(str).apply(self._clean_id) == site) & \
               (self.df_main['Screening #'].astype(str).apply(self._clean_id) == pat)
        rows = self.df_main[mask]
        if rows.empty: return
        row = rows.iloc[0]
        
        matrix_data = [] 
        processed_cols = set()
        
        # Fetch Procedure Date for "Treat. Day" calculations
        proc_date = None
        if 'TV_PR_SVDTC' in row and pd.notna(row['TV_PR_SVDTC']):
            proc_date_str = str(row['TV_PR_SVDTC']).split('T')[0]
            try:
                proc_date = datetime.strptime(proc_date_str, '%Y-%m-%d')
            except Exception:
                pass

        for item in selected_items:
            # Recursively collect all descendants to handle folder selections (L1 or L2)
            # This ensures we find the actual data nodes (L3) even when selecting a top-level folder
            descendants = self._get_all_descendants(item)
            items_to_process = [item] + descendants if descendants else [item]
            
            for proc_item in items_to_process:
                vals = self.tree.item(proc_item, "values")
                if not vals or len(vals) < 5: continue 
                
                col_name = vals[4]  # DB Variable is 5th column (index 4)
                if not col_name or col_name in processed_cols: continue
                
                # Check if column exists in dataframe
                if col_name not in self.df_main.columns:
                    # Try to find without the suffix variations
                    continue
                
                is_ae_ref_col = "LBREF" in col_name or "PRREF" in col_name
                
                # Check for AE-related columns
                is_ae_col = any(x in col_name for x in ["AETERM", "AESTDTC", "AEENDTC", "AESER", "AESEV", "AEOUT", "AEREL", "AEDECOD"])
                
                # Check for CM (Concomitant Medications) columns
                is_cm_col = any(x in col_name for x in ["CMTRT", "CMDOSE", "CMDOSFRQ", "CMSTDTC", "CMSTDAT", "CMENDTC", "CMENDAT", "CMINDC", "CMROUTE", "CMDOSU", "CMONGO", "CMREF"])
                
                # Check for MH (Medical History) columns
                is_mh_col = "_MH_" in col_name and any(x in col_name for x in ["MHTERM", "MHSTDTC", "MHENDTC", "MHONGO", "MHCAT", "MHBODSYS", "MHOCCUR"])
                
                # Check for HFH (Heart Failure History) columns
                is_hfh_col = "_HFH_" in col_name and any(x in col_name for x in ["HOSTDTC", "HOTERM", "HONUM", "HOOCCUR", "HODESC"])
                
                # Check for HMEH (Hospitalization and Medical Events History) columns
                is_hmeh_col = "_HMEH_" in col_name or "HMEH" in col_name
                
                # Check for CVC (Cardiac and Venous Catheterization) columns
                is_cvc_col = "_CVC_" in col_name and any(x in col_name for x in ["CVORRES", "FAORRES", "PRSTDTC"])
                
                # Check for CVH (Cardiovascular History) columns
                is_cvh_col = "_CVH_" in col_name and any(x in col_name for x in ["PRSTDTC", "PRCAT", "PRTRT", "PROCCUR"])
                
                # Check for ACT (ACT Lab Results) columns - Heparin and ACT measurements
                is_act_col = "_LB_ACT_" in col_name and any(x in col_name for x in ["LBORRES", "LBTIM", "CMTIM", "CMDOS", "CMSTAT", "LBSTAT"])
                
                if not is_ae_ref_col and not is_ae_col and not is_cm_col and not is_mh_col and not is_hfh_col and not is_hmeh_col and not is_cvc_col and not is_cvh_col and not is_act_col:
                    if not self._is_matrix_supported_col(col_name):
                        continue

                processed_cols.add(col_name)
                
                val_in_db = row[col_name]
                if pd.isna(val_in_db): continue
                val_str = str(val_in_db)
                
                if "#" in val_str and "/" in val_str and " / " in val_str:
                    entries = [e.strip() for e in val_str.split('|')] if '|' in val_str else [val_str]
                    for e in entries:
                        if e.startswith("#") and "/" in e:
                            parsed = self.parse_timeline_entry(e)
                            if parsed:
                                row_num, d, p, v = parsed
                                # AE Reference - only for LOGS columns, as separate field
                                ae_ref = ""
                                is_logs_col = "LOGS_" in col_name
                                if is_logs_col:
                                    # Check for PRORRES specifically since 'LB' appears in all LOGS columns
                                    ref_type = 'PR' if 'PRORRES' in col_name else 'LB'
                                    ae_num, ae_term = self.get_ae_info(row_num, ref_type)
                                    if ae_term:
                                        ae_ref = f"AE#{ae_num}"
                                matrix_data.append({'Time': d, 'Param': p, 'Value': v, 'AE_Ref': ae_ref})
                else:
                    # Check if this is an AE column - handle differently
                    if is_ae_col:
                        # For AE data, use the dedicated AE sheet (df_ae) if available
                        # This will be handled separately after the loop
                        if not hasattr(self, '_ae_data_requested'):
                            self._ae_data_requested = True
                        continue  # Skip regular processing, handle AE data from df_ae
                    elif is_cm_col:
                        # For CM data, use the dedicated CM sheet (df_cm) if available
                        # This will be handled separately after the loop
                        if not hasattr(self, '_cm_data_requested'):
                            self._cm_data_requested = True
                        continue  # Skip regular processing, handle CM data from df_cm
                    elif is_mh_col:
                        # For MH data, use the main dataframe MH columns
                        # This will be handled separately after the loop
                        if not hasattr(self, '_mh_data_requested'):
                            self._mh_data_requested = True
                        continue  # Skip regular processing, handle MH data separately
                    elif is_hfh_col:
                        # For HFH (Heart Failure History) data
                        if not hasattr(self, '_hfh_data_requested'):
                            self._hfh_data_requested = True
                        continue
                    elif is_hmeh_col:
                        # For HMEH (Hospitalization and Medical Events History) data
                        if not hasattr(self, '_hmeh_data_requested'):
                            self._hmeh_data_requested = True
                        continue
                    elif is_cvc_col:
                        # For CVC (Cardiac and Venous Catheterization) data
                        if not hasattr(self, '_cvc_data_requested'):
                            self._cvc_data_requested = True
                        continue
                    elif is_cvh_col:
                        # For CVH (Cardiovascular History) data - use CVH_TABLE sheet
                        if not hasattr(self, '_cvh_data_requested'):
                            self._cvh_data_requested = True
                        continue
                    elif is_act_col:
                        # For ACT (ACT Lab Results) data - use LB_ACT sheet
                        if not hasattr(self, '_act_data_requested'):
                            self._act_data_requested = True
                        continue
                    else:
                        # Special handling for Treatment Visit (TV_) lab columns
                        # These have shared date columns (e.g., TV_LB_BMP_DV_LBDAT_BMP) 
                        # and test-specific result columns (e.g., TV_LB_BMP_DV_LBORRES_BUN)
                        is_tv_lab = col_name.startswith("TV_") and "_LB_" in col_name and "_DV_" in col_name
                        
                        date_col = None
                        unit_col = None
                        test_name_col = None
                        
                        if is_tv_lab:
                            # Extract lab type from column name (e.g., "BMP" from "TV_LB_BMP_DV_LBORRES_BUN")
                            # Pattern: TV_LB_{LABTYPE}_DV_LBORRES_{TEST} or TV_LB_{LABTYPE}_DV_eGFR
                            tv_parts = col_name.split("_DV_")
                            if len(tv_parts) >= 2:
                                lab_prefix = tv_parts[0]  # e.g., "TV_LB_BMP"
                                suffix_part = tv_parts[1]  # e.g., "LBORRES_BUN" or "eGFR"
                                
                                # Extract lab type (BMP, CBC, etc.) for finding shared date column
                                lab_match = re.search(r'TV_LB_(\w+)', lab_prefix)
                                lab_type = lab_match.group(1) if lab_match else None
                                
                                # Find shared date column: TV_LB_{LABTYPE}_DV_LBDAT_{LABTYPE}
                                if lab_type:
                                    shared_date_col = f"{lab_prefix}_DV_LBDAT_{lab_type}"
                                    if shared_date_col in self.df_main.columns:
                                        date_col = shared_date_col
                                
                                # Find unit column: replace LBORRES with LBORRESU (keep same test suffix)
                                # Also look for "other units" column for when unit value is "Other"
                                other_unit_col = None
                                if "LBORRES_" in suffix_part:
                                    test_suffix = suffix_part.split("LBORRES_")[1]  # e.g., "BUN"
                                    unit_col_name = f"{lab_prefix}_DV_LBORRESU_{test_suffix}"
                                    if unit_col_name in self.df_main.columns:
                                        unit_col = unit_col_name
                                    # Look for "other units" column
                                    other_unit_col_name = f"{lab_prefix}_DV_LBORRESU_OTH_{test_suffix}"
                                    if other_unit_col_name in self.df_main.columns:
                                        other_unit_col = other_unit_col_name
                                # Special case: eGFR has no unit column (unit is in the label)
                        else:
                            # Standard parallel arrays handling for non-TV lab/proc data
                            prefix = None
                            if "_LBORRES" in col_name: prefix = col_name.split("_LBORRES")[0]
                            elif "_ORRES" in col_name: prefix = col_name.split("_ORRES")[0]
                            elif "_PRORRES" in col_name: prefix = col_name.split("_PRORRES")[0]
                            
                            if prefix:
                                is_lab_col = ("_LBORRES" in col_name or "_ORRES" in col_name) and "_PRORRES" not in col_name
                                
                                for cand in self.df_main.columns:
                                    if cand.startswith(prefix):
                                        if is_lab_col:
                                            if "LBDTC" in cand or "LBDAT" in cand:
                                                date_col = cand
                                            if "LBTEST" in cand:
                                                test_name_col = cand
                                            if "LBORRESU" in cand and "OTH" not in cand:
                                                unit_col = cand
                                            if "LBORRESU_OTH" in cand:
                                                other_unit_col = cand
                                        else:
                                            if "PRDTC" in cand or "PRDAT" in cand:
                                                date_col = cand
                                            if "PRTEST" in cand:
                                                test_name_col = cand
                                
                                # Fallback: try to find shared date column like TV_ labs have
                                # Pattern: {prefix}_LB_{type}_LBDAT_{type} or {prefix}_LB_{type}_LBDTC
                                if not date_col:
                                    # Extract visit prefix and lab type from column name
                                    # e.g., SBV_LB_BMP_LBORRES_NTP -> visit=SBV, lab_type=BMP or BM
                                    lab_match = re.match(r'^(\w+)_LB_(\w+?)P?_', col_name)
                                    if lab_match:
                                        visit_prefix = lab_match.group(1)  # e.g., "SBV"
                                        lab_type_base = lab_match.group(2)  # e.g., "BM" from "BMP"
                                        
                                        # Try shared date column patterns
                                        shared_date_patterns = [
                                            f"{visit_prefix}_LB_{lab_type_base}_LBDAT_{lab_type_base}",  # SBV_LB_BM_LBDAT_BM
                                            f"{visit_prefix}_LB_{lab_type_base}_LBDTC_{lab_type_base}",
                                            f"{visit_prefix}_LB_{lab_type_base}_LBDTC",
                                        ]
                                        for pattern in shared_date_patterns:
                                            if pattern in self.df_main.columns:
                                                date_col = pattern
                                                break

                        r_vals = [v.strip() for v in val_str.split('|')]
                        param_name_default = self.clean_label(self.labels.get(col_name, col_name))
                        # Add pre/post annotation for Procedure Timing
                        param_name_default = self.annotate_procedure_timing(param_name_default, col_name)
                        
                        # Strip '/result' suffix from parameter name for cleaner display
                        if param_name_default.lower().endswith('/result'):
                            param_name_default = param_name_default[:-7]  # Remove '/result'
                        
                        n_vals = []
                        if test_name_col and pd.notna(row[test_name_col]):
                            n_vals = [n.strip() for n in str(row[test_name_col]).split('|')]
                        
                        d_vals = []
                        if date_col and pd.notna(row[date_col]):
                            d_vals = [d.strip() for d in str(row[date_col]).split('|')]
                        
                        u_vals = []
                        if unit_col and pd.notna(row[unit_col]):
                            u_vals = [u.strip() for u in str(row[unit_col]).split('|')]
                        
                        # Get "other units" values if available (for when unit is "Other")
                        other_u_vals = []
                        if is_tv_lab and 'other_unit_col' in dir() and other_unit_col and pd.notna(row.get(other_unit_col, None)):
                            other_u_vals = [u.strip() for u in str(row[other_unit_col]).split('|')]
                        
                        for i, val in enumerate(r_vals):
                            if not val: continue
                            curr_param = n_vals[i] if n_vals and i < len(n_vals) and n_vals[i] else param_name_default
                            # Also strip '/result' from individual test names
                            if curr_param.lower().endswith('/result'):
                                curr_param = curr_param[:-7]
                            
                            # Determine time label: prefer date, fallback to visit date from schedule
                            curr_date = "Unknown"
                            visit_label = ""
                            
                            # 1. Identify visit prefix and label
                            for prefix, label in VISIT_MAP.items():
                                if col_name.startswith(prefix + "_"):
                                    visit_label = label
                                    break
                            
                            # 2. Get specific date if available
                            specific_date = None
                            if d_vals and i < len(d_vals) and d_vals[i]:
                                specific_date_str = str(d_vals[i]).split('T')[0]
                                specific_date = specific_date_str
                            
                            # 3. Fallback to VISIT_SCHEDULE if specific_date missing
                            if not specific_date and visit_label:
                                for date_col_name, sv_label in VISIT_SCHEDULE:
                                    if date_col_name.startswith(prefix + "_"):
                                        visit_date_val = row.get(date_col_name)
                                        if pd.notna(visit_date_val):
                                            specific_date = str(visit_date_val).split('T')[0]
                                        break
                                        
                            # 4. Refine visit_label for Treatment if multi-day
                            header_visit = visit_label
                            if visit_label == "Treatment" and proc_date and specific_date:
                                try:
                                    curr_dt = datetime.strptime(specific_date, '%Y-%m-%d')
                                    delta = (curr_dt - proc_date).days
                                    if delta > 0:
                                        header_visit = f"Treat. Day +{delta}"
                                    elif delta < 0:
                                        header_visit = f"Treat. Day {delta}"
                                    # Day 0 remains "Treatment"
                                except Exception:
                                    pass
                                    
                            # 5. Build final header
                            if header_visit and specific_date:
                                curr_date = f"{header_visit} ({specific_date})"
                            elif header_visit:
                                curr_date = header_visit
                            elif specific_date:
                                curr_date = specific_date
                            
                            # Keep value separate - don't combine with unit
                            display_val = val
                            
                            # For unit columns (LBORRESU), if value is "Other", substitute with actual unit
                            is_unit_col = "LBORRESU" in col_name and "OTH" not in col_name
                            if is_unit_col and val.lower() == "other":
                                # Look for corresponding OTH column
                                oth_col_name = col_name.replace("LBORRESU_", "LBORRESU_OTH_")
                                if oth_col_name in self.df_main.columns:
                                    oth_val = row.get(oth_col_name, None)
                                    if pd.notna(oth_val):
                                        oth_vals = str(oth_val).split('|')
                                        if i < len(oth_vals) and oth_vals[i].strip():
                                            display_val = oth_vals[i].strip()
                            
                            # Append units for specific lab panels (Biomarkers, Enz, CBC, BMP, Coag, LFP)
                            # Only for result columns, not unit columns themselves
                            lab_panels = ["_LB_BM_", "_LB_ENZ_", "_LB_CBC_", "_LB_BMP_", "_LB_COA", "_LB_LFP_"]
                            is_result_col = ("_LBORRES_" in col_name or "_ORRES" in col_name) and "LBORRESU" not in col_name
                            if any(p in col_name for p in lab_panels) and is_result_col:
                                curr_unit = ""
                                if u_vals and i < len(u_vals) and u_vals[i] and u_vals[i].lower() not in ['nan', 'none', '']:
                                    curr_unit = u_vals[i]
                                
                                # Handle "Other" unit substitution from parallel array if applicable
                                if curr_unit.lower() == "other" and other_u_vals and i < len(other_u_vals):
                                    if other_u_vals[i] and other_u_vals[i].lower() not in ['nan', '']:
                                        curr_unit = other_u_vals[i]
                                
                                if curr_unit:
                                    display_val = f"{display_val} {curr_unit}"
                            
                            # AE Reference - only for LOGS additional tests columns, and as separate field
                            ae_ref = ""
                            is_logs_col = "LOGS_" in col_name
                            if is_logs_col:
                                # Check for PRORRES specifically since 'LB' appears in all LOGS columns
                                ref_type = 'PR' if 'PRORRES' in col_name else 'LB'
                                ae_num, ae_term = self.get_ae_info(str(i+1), ref_type)
                                if ae_term:
                                    ae_ref = f"AE#{ae_num}"
                            
                            # Angiography Column Splitting (Pre/Post Procedure)
                            # Split by modifying the 'Time' header (column) and cleaning the 'Param' (row)
                            # Use count > 1 because "_AG_" usually implies "_PRE_POST_" prefix is present
                            if "_AG_" in col_name:
                                if col_name.count("_PRE_") > 1:
                                    curr_date = f"{curr_date} (Pre-Procedure)"
                                    if curr_param.startswith("Pre-procedure / "):
                                        curr_param = curr_param.replace("Pre-procedure / ", "")
                                    elif curr_param.startswith("Pre-procedure /"):
                                        curr_param = curr_param.replace("Pre-procedure /", "")
                                elif col_name.count("_POST_") > 1:
                                    curr_date = f"{curr_date} (Post-Procedure)"
                                    if curr_param.startswith("Post-procedure / "):
                                        curr_param = curr_param.replace("Post-procedure / ", "")
                                    elif curr_param.startswith("Post-procedure /"):
                                        curr_param = curr_param.replace("Post-procedure /", "")
                            
                            matrix_data.append({'Time': curr_date, 'Param': curr_param, 'Value': display_val, 'AE_Ref': ae_ref})

        # Handle AE data separately if requested - show structured AE table from df_ae
        if hasattr(self, '_ae_data_requested') and self._ae_data_requested:
            del self._ae_data_requested  # Reset the flag
            
            if self.df_ae is not None and not self.df_ae.empty:
                # Filter AE data for this patient
                pat_aes = self.df_ae[self.df_ae['Screening #'].astype(str).str.contains(pat.replace('-', '-'), na=False)]
                
                if not pat_aes.empty:
                    # Show AE-specific matrix window
                    self.matrix_display.show_ae_matrix(pat_aes, pat)
                    return
                else:
                    messagebox.showinfo("Info", "No adverse events found for this patient in AE sheet.")
                    return
            else:
                messagebox.showinfo("Info", "AE sheet not loaded or empty.")
                return

        # Handle CM data separately if requested - extract from Main sheet LOGS_CM columns
        if hasattr(self, '_cm_data_requested') and self._cm_data_requested:
            del self._cm_data_requested  # Reset the flag
            
            # First try dedicated CM sheet, otherwise parse from Main sheet
            if self.df_cm is not None and not self.df_cm.empty:
                # Filter CM data for this patient
                pat_cms = self.df_cm[self.df_cm['Screening #'].astype(str).str.contains(pat.replace('-', '-'), na=False)]
                
                if not pat_cms.empty:
                    # Show CM-specific matrix window
                    self.matrix_display.show_cm_matrix(pat_cms, pat)
                    return
            
            # Parse CM data from Main sheet LOGS_CM columns
            cm_cols = {
                'CMTRT': 'Medication',
                'CMDOSE': 'Dose', 
                'CMDOSU': 'Dose Unit',
                'CMROUTE': 'Route',
                'CMINDC': 'Indication',
                'CMSTDTC': 'Start Date',
                'CMENDTC': 'End Date',
                'CMENDAT': 'End Date',
                'CMONGO': 'Ongoing',
                'CMDOSFRQ': 'Frequency',
                'CMDOSFRQ_OTH': 'Frequency (Other)',
            }
            
            # Find all LOGS_CM columns in main sheet
            logs_cm_cols = {}
            for col in self.df_main.columns:
                col_str = str(col)
                if 'LOGS_CM_' in col_str or (col_str.startswith('LOGS_') and '_CM_' in col_str):
                    for cm_key, display_name in cm_cols.items():
                        if cm_key in col_str:
                            logs_cm_cols[display_name] = col_str
                            break
            
            if not logs_cm_cols:
                messagebox.showinfo("Info", "No CM columns found in data.")
                return
            
            # Get medication names - this is the key column
            med_col = logs_cm_cols.get('Medication')
            if not med_col or pd.isna(row.get(med_col)):
                messagebox.showinfo("Info", "No medications found for this patient.")
                return
            
            # Parse pipe-delimited values
            med_vals = [m.strip() for m in str(row[med_col]).split('|') if m.strip() and m.strip().lower() != 'nan']
            
            if not med_vals:
                messagebox.showinfo("Info", "No medications found for this patient.")
                return
            
            # Build CM data records
            cm_data = []
            for i, med in enumerate(med_vals):
                record = {'CM #': str(i + 1), 'Medication': med}
                
                for display_name, col_name in logs_cm_cols.items():
                    if display_name == 'Medication':
                        continue
                    
                    col_val = row.get(col_name, '')
                    if pd.notna(col_val):
                        vals = [v.strip() for v in str(col_val).split('|')]
                        val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                        
                        # Clean up date values
                        if 'Date' in display_name and val:
                            if 'T' in val:
                                val = val.split('T')[0]
                            val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
                        
                        record[display_name] = val
                
                # Handle Ongoing flag
                if record.get('Ongoing', '').lower() in ['yes', 'y', '1', 'true', 'checked']:
                    record['End Date'] = 'Ongoing'
                
                # Handle "Other" frequency - substitute with actual value from CMDOSFRQ_OTH
                if record.get('Frequency', '').lower() == 'other':
                    freq_other = record.get('Frequency (Other)', '')
                    if freq_other:
                        record['Frequency'] = freq_other
                
                # Calculate Daily Dose
                try:
                    dose_str = record.get('Dose', '')
                    freq_str = record.get('Frequency', '')
                    freq_oth = record.get('Frequency (Other)', '')
                    unit_str = record.get('Dose Unit', '')
                    
                    if dose_str and dose_str.lower() not in ['nan', 'none', '']:
                        single_dose = float(dose_str)
                        multiplier, freq_note, override_dose = self.matrix_display.parse_frequency_multiplier(freq_str, freq_oth)
                        
                        if override_dose is not None:
                            daily = override_dose
                        elif multiplier is not None:
                            daily = single_dose * multiplier
                        else:
                            daily = None
                        
                        if daily is not None:
                            if daily == int(daily):
                                daily_str = str(int(daily))
                            else:
                                daily_str = f"{daily:.1f}"
                            
                            # Add unit
                            if unit_str and unit_str.lower() not in ['nan', 'none', '']:
                                if 'milligram' in unit_str.lower():
                                    unit_str = 'mg'
                                daily_str += f" {unit_str}/day"
                            else:
                                daily_str += "/day"
                            record['Daily Dose'] = daily_str
                        elif freq_note:
                            record['Daily Dose'] = f"{int(single_dose) if single_dose == int(single_dose) else single_dose} {freq_note}"
                except (ValueError, TypeError):
                    pass
                
                # Remove the internal 'Frequency (Other)' key from final display
                record.pop('Frequency (Other)', None)
                
                cm_data.append(record)
            
            # Show CM matrix from parsed data
            self.matrix_display.show_cm_matrix_from_data(cm_data, pat)
            return

        # Handle MH (Medical History) data separately if requested
        if hasattr(self, '_mh_data_requested') and self._mh_data_requested:
            del self._mh_data_requested  # Reset the flag
            
            # Medical History column mappings
            mh_cols = {
                'MHTERM': 'Condition',
                'MHBODSYS': 'Body System',
                'MHCAT': 'Category',
                'MHSTDTC': 'Start Date',
                'MHENDTC': 'End Date',
                'MHONGO': 'Ongoing',
                # Note: MHOCCUR (Status) excluded - always "Yes" for reported conditions
            }
            
            # Find all MH columns in main sheet (look for _MH_ pattern)
            visit_prefixes = ['SBV', 'SCR']  # Medical History is typically at Screening/Baseline
            mh_columns = {}
            
            for col in self.df_main.columns:
                col_str = str(col)
                if '_MH_' in col_str:
                    for mh_key, display_name in mh_cols.items():
                        if mh_key in col_str:
                            if display_name not in mh_columns:
                                mh_columns[display_name] = col_str
                            break
            
            if not mh_columns:
                messagebox.showinfo("Info", "No Medical History columns found in data.")
                return
            
            # Get the condition/term column - this is the key column
            term_col = mh_columns.get('Condition')
            if not term_col or pd.isna(row.get(term_col)):
                messagebox.showinfo("Info", "No medical history conditions found for this patient.")
                return
            
            # Parse pipe-delimited values for conditions
            term_vals = [t.strip() for t in str(row[term_col]).split('|') if t.strip() and t.strip().lower() != 'nan']
            
            if not term_vals:
                messagebox.showinfo("Info", "No medical history conditions found for this patient.")
                return
            
            # Build MH data records
            mh_data = []
            for i, term in enumerate(term_vals):
                record = {'MH #': str(i + 1), 'Condition': term}
                
                for display_name, col_name in mh_columns.items():
                    if display_name == 'Condition':
                        continue
                    
                    col_val = row.get(col_name, '')
                    if pd.notna(col_val):
                        vals = [v.strip() for v in str(col_val).split('|')]
                        val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                        
                        # Clean up date values
                        if 'Date' in display_name and val:
                            if 'T' in val:
                                val = val.split('T')[0]
                            val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
                            # Handle date unknown values
                            if val.lower() in ['date unknown', 'unknown date', 'unknown']:
                                val = 'Date Unknown'
                        
                        record[display_name] = val
                
                # Handle Ongoing flag - modify End Date display
                if record.get('Ongoing', '').lower() in ['yes', 'y', '1', 'true', 'checked']:
                    record['End Date'] = 'Ongoing'
                
                mh_data.append(record)
            
            # Show MH matrix
            self.matrix_display.show_mh_matrix(mh_data, pat)
            return

        # Handle HFH (Heart Failure History) data separately if requested
        if hasattr(self, '_hfh_data_requested') and self._hfh_data_requested:
            del self._hfh_data_requested  # Reset the flag
            
            # HFH column mappings
            hfh_cols = {
                'HOSTDTC': 'Hospitalization Date',
                'HOTERM': 'Details',
                'HONUM': 'Number of Hospitalizations',
                # Note: HOOCCUR excluded - always "Yes" for reported items
            }
            
            # Find all HFH columns in main sheet
            hfh_columns = {}
            for col in self.df_main.columns:
                col_str = str(col)
                if '_HFH_' in col_str:
                    for hfh_key, display_name in hfh_cols.items():
                        if hfh_key in col_str:
                            if display_name not in hfh_columns:
                                hfh_columns[display_name] = col_str
                            break
            
            if hfh_columns:
                # Build HFH data records
                hfh_data = []
                # Get the primary date column values
                date_col = hfh_columns.get('Hospitalization Date')
                if date_col and pd.notna(row.get(date_col)):
                    date_vals = [d.strip() for d in str(row[date_col]).split('|') if d.strip() and d.strip().lower() != 'nan']
                    
                    for i, date_val in enumerate(date_vals):
                        record = {'HFH #': str(i + 1), 'Hospitalization Date': date_val}
                        
                        for display_name, col_name in hfh_columns.items():
                            if display_name == 'Hospitalization Date':
                                continue
                            col_val = row.get(col_name, '')
                            if pd.notna(col_val):
                                vals = [v.strip() for v in str(col_val).split('|')]
                                val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                                record[display_name] = val
                        
                        hfh_data.append(record)
                    
                    if hfh_data:
                        self.matrix_display.show_hfh_matrix(hfh_data, pat)
                        return
            
            messagebox.showinfo("Info", "No Heart Failure History data found for this patient.")
            return

        # Handle HMEH (Hospitalization and Medical Events History) data separately if requested
        if hasattr(self, '_hmeh_data_requested') and self._hmeh_data_requested:
            del self._hmeh_data_requested  # Reset the flag
            
            # HMEH column mappings
            hmeh_cols = {
                'HOSTDTC': 'Event Date',
                'HOTERM': 'Event Details',
            }
            
            # Find all HMEH columns in main sheet
            hmeh_columns = {}
            for col in self.df_main.columns:
                col_str = str(col)
                if '_HMEH_' in col_str or 'HMEH' in col_str:
                    for hmeh_key, display_name in hmeh_cols.items():
                        if hmeh_key in col_str:
                            if display_name not in hmeh_columns:
                                hmeh_columns[display_name] = col_str
                            break
            
            if hmeh_columns:
                # Build HMEH data records
                hmeh_data = []
                # Get the primary date/term column values
                date_col = hmeh_columns.get('Event Date')
                term_col = hmeh_columns.get('Event Details')
                
                primary_col = date_col if date_col and pd.notna(row.get(date_col)) else term_col
                if primary_col and pd.notna(row.get(primary_col)):
                    primary_vals = [v.strip() for v in str(row[primary_col]).split('|') if v.strip() and v.strip().lower() != 'nan']
                    
                    for i, pval in enumerate(primary_vals):
                        record = {'HMEH #': str(i + 1)}
                        
                        for display_name, col_name in hmeh_columns.items():
                            col_val = row.get(col_name, '')
                            if pd.notna(col_val):
                                vals = [v.strip() for v in str(col_val).split('|')]
                                val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                                # Clean up dates
                                if 'Date' in display_name and val:
                                    if 'T' in val:
                                        val = val.split('T')[0]
                                record[display_name] = val
                        
                        hmeh_data.append(record)
                    
                    if hmeh_data:
                        self.matrix_display.show_hmeh_matrix(hmeh_data, pat)
                        return
            
            messagebox.showinfo("Info", "No Hospitalization/Medical Events History data found for this patient.")
            return

        # Handle CVC (Cardiac and Venous Catheterization) data separately if requested
        if hasattr(self, '_cvc_data_requested') and self._cvc_data_requested:
            del self._cvc_data_requested  # Reset the flag
            
            # Use CVCExporter to generate formatted tables
            exporter = CVCExporter(self.df_main)
            
            # Determine which visit types are present based on selected columns
            is_screening = any("SBV_CVC_" in str(col) for col in processed_cols)
            is_treatment = any("TV_CVC_" in str(col) for col in processed_cols)
            
            # If selecting Cardiac and Venous Catheterization folder, detect all CVC types
            # Check if we have both Screening and Treatment data
            has_any_cvc = is_screening or is_treatment
            
            if not has_any_cvc:
                # User may have selected the parent folder - try to show both tables
                is_screening = True
                is_treatment = True
            
            tables_shown = 0
            
            # Show Screening table
            if is_screening:
                screening_df = exporter.generate_screening_table(pat)
                if screening_df is not None:
                    self.matrix_display.show_cvc_matrix(screening_df, pat, "Screening")
                    tables_shown += 1
            
            # Show Treatment (Hemodynamic/Pre-Post) table
            if is_treatment:
                hemo_df = exporter.generate_hemodynamic_table(pat)
                if hemo_df is not None:
                    self.matrix_display.show_cvc_matrix(hemo_df, pat, "Hemodynamic Effect (Pre/Post Procedure)")
                    tables_shown += 1
            
            if tables_shown == 0:
                messagebox.showinfo("Info", "No CVC data found for this patient.")
            return

        # Handle CVH (Cardiovascular History) data separately if requested
        if hasattr(self, '_cvh_data_requested') and self._cvh_data_requested:
            del self._cvh_data_requested  # Reset the flag
            
            if self.df_cvh is not None and not self.df_cvh.empty:
                # Filter CVH data for this patient
                pat_cvh = self.df_cvh[self.df_cvh['Screening #'].astype(str).str.contains(pat.replace('-', '-'), na=False)]
                
                if not pat_cvh.empty:
                    # Build CVH data records
                    cvh_data = []
                    for _, cvh_row in pat_cvh.iterrows():
                        # Get date - prefer full date, fallback to partial date
                        full_date = cvh_row.get('SBV_CVH_PRSTDTC', '')
                        partial_date = cvh_row.get('SBV_CVH_PRSTDTC_PARTIAL', '')
                        is_partial = str(cvh_row.get('SBV_CVH_PRSTDTC_PARTIAL_CHECKBOX', '')).lower() in ['yes', 'checked', 'true', '1']
                        
                        if pd.notna(full_date) and str(full_date).strip() and str(full_date).strip().lower() not in ['nan', 'nat']:
                            date_str = str(full_date).split('T')[0] if 'T' in str(full_date) else str(full_date)
                        elif pd.notna(partial_date) and str(partial_date).strip():
                            date_str = f"{partial_date} (partial)"
                        else:
                            date_str = "Unknown"
                        
                        # Get intervention type
                        int_type = cvh_row.get('SBV_CVH_PRCAT', '')
                        int_type_str = str(int_type).strip() if pd.notna(int_type) and str(int_type).strip().lower() not in ['nan', ''] else ""
                        
                        # Get intervention term
                        int_term = cvh_row.get('SBV_CVH_PRTRT', '')
                        int_term_str = str(int_term).strip() if pd.notna(int_term) and str(int_term).strip().lower() not in ['nan', ''] else ""
                        
                        # Get "other" category if type is Other
                        if int_type_str.lower() == 'other':
                            int_type_oth = cvh_row.get('SBV_CVH_PRCAT_OTH', '')
                            if pd.notna(int_type_oth) and str(int_type_oth).strip():
                                int_type_str = f"Other: {int_type_oth}"
                        
                        # Get "other" term if intervention is Other
                        int_term_oth = cvh_row.get('SBV_CVH_PRTRT_OTHCAT', '')
                        if int_term_str.lower() == 'other' and pd.notna(int_term_oth) and str(int_term_oth).strip():
                            int_term_str = f"Other: {int_term_oth}"
                        
                        # Only add if we have meaningful data
                        if int_term_str or int_type_str or date_str != "Unknown":
                            cvh_data.append({
                                'Date': date_str,
                                'Type of Intervention': int_type_str,
                                'Intervention': int_term_str
                            })
                    
                    if cvh_data:
                        self.matrix_display.show_cvh_matrix(cvh_data, pat)
                        return
                    else:
                        messagebox.showinfo("Info", "No Cardiovascular History interventions found for this patient.")
                        return
                else:
                    messagebox.showinfo("Info", "No Cardiovascular History data found for this patient in CVH sheet.")
                    return
            else:
                messagebox.showinfo("Info", "CVH sheet not loaded or empty.")
                return

        # Handle ACT (ACT Lab Results) data separately if requested
        if hasattr(self, '_act_data_requested') and self._act_data_requested:
            del self._act_data_requested  # Reset the flag
            
            if self.df_act is not None and not self.df_act.empty:
                # Filter ACT data for this patient
                # Find Screening column robustly
                scr_col = next((c for c in self.df_act.columns if "Screening" in str(c) and "#" in str(c)), None)
                if not scr_col:
                    scr_col = next((c for c in self.df_act.columns if "Screening" in str(c)), None)
                
                if scr_col:
                    pat_clean = str(pat).strip()
                    logger.debug("Matrix Filter. Pat raw='%s' clean='%s'", pat, pat_clean)
                    logger.debug("Using Scr Col='%s'", scr_col)
                    # Filter with strip and regex=False
                    pat_act = self.df_act[self.df_act[scr_col].astype(str).str.strip().str.contains(pat_clean, regex=False, na=False)]
                    logger.debug("Matrix Rows Found: %d", len(pat_act))
                else:
                    pat_act = pd.DataFrame()
                
                if not pat_act.empty:
                    # Build ACT data records - combine ACT measurements and Heparin administrations
                    act_events = []
                    
                    for _, act_row in pat_act.iterrows():
                        # Get ACT measurement data (Try TV then UV)
                        act_time = act_row.get('TV_LB_ACT_LBTIM_ACT', act_row.get('UV_LB_ACT_LBTIM_ACT', ''))
                        act_level = act_row.get('TV_LB_ACT_LBORRES_ACT', act_row.get('UV_LB_ACT_LBORRES_ACT', ''))
                        act_stat = act_row.get('TV_LB_ACT_LBSTAT_ACT', act_row.get('UV_LB_ACT_LBSTAT_ACT', ''))
                        
                        # Add ACT measurement if present
                        if pd.notna(act_time) and str(act_time).strip() and str(act_time).strip().lower() not in ['nan', '']:
                            act_level_str = str(act_level).strip() if pd.notna(act_level) else ""
                            act_events.append({
                                'Time': str(act_time).strip(),
                                'Event': "ACT Level",
                                'Value': f"{act_level_str} sec" if act_level_str else "",
                                'Type': 'ACT',
                                'Status': 'OK' if act_level_str else 'GAP'
                            })
                        elif pd.notna(act_stat) and str(act_stat).strip().lower() in ['not done', 'not performed']:
                            act_events.append({
                                'Time': '',
                                'Event': "ACT Level",
                                'Value': "Not Done",
                                'Type': 'ACT',
                                'Status': 'Confirmed'
                            })
                        
                        # Get Heparin administration data
                        hep_time = act_row.get('TV_LB_ACT_CMTIM_HEP', act_row.get('UV_LB_ACT_CMTIM_HEP', ''))
                        hep_dose = act_row.get('TV_LB_ACT_CMDOS_HEP', act_row.get('UV_LB_ACT_CMDOS_HEP', ''))
                        hep_stat = act_row.get('TV_LB_ACT_CMSTAT_HEP', act_row.get('UV_LB_ACT_CMSTAT_HEP', ''))
                        
                        # Add Heparin administration if present
                        if pd.notna(hep_time) and str(hep_time).strip() and str(hep_time).strip().lower() not in ['nan', '']:
                            hep_dose_str = str(hep_dose).strip() if pd.notna(hep_dose) else ""
                            act_events.append({
                                'Time': str(hep_time).strip(),
                                'Event': "Heparin",
                                'Value': f"{hep_dose_str} Units" if hep_dose_str else "",
                                'Type': 'HEP',
                                'Status': 'OK' if hep_dose_str else 'GAP'
                            })
                        elif pd.notna(hep_stat) and str(hep_stat).strip().lower() in ['not done', 'not performed']:
                            act_events.append({
                                'Time': '',
                                'Event': "Heparin",
                                'Value': "Not Done",
                                'Type': 'HEP',
                                'Status': 'Confirmed'
                            })
                    
                    if act_events:
                        # Sort chronologically by time
                        def parse_time(t):
                            try:
                                parts = str(t).split(':')
                                if len(parts) >= 2:
                                    return int(parts[0]) * 60 + int(parts[1])
                                return 9999
                            except Exception:
                                return 9999
                        
                        act_events.sort(key=lambda x: parse_time(x['Time']))
                        self.matrix_display.show_act_matrix(act_events, pat)
                        return
                    else:
                        messagebox.showinfo("Info", "No ACT/Heparin data found for this patient.")
                        return
                else:
                    messagebox.showinfo("Info", "No ACT data found for this patient in ACT sheet.")
                    return
            else:
                messagebox.showinfo("Info", "ACT sheet not loaded or empty.")
                return

        if not matrix_data:
            messagebox.showinfo("Info", "No data found.")
            return

        df_matrix = pd.DataFrame(matrix_data)
        
        # Filter out rows where param ends with '/' (these are partial unit duplicates like 'nt-probnp/')
        df_matrix = df_matrix[~df_matrix['Param'].str.strip().str.endswith('/')]
        # Also filter out empty param names
        df_matrix = df_matrix[df_matrix['Param'].str.strip() != '']
        
        # Create a composite row key combining Param and AE_Ref for grouping
        # This ensures entries with same param but different AE associations are in separate rows
        df_matrix['AE_Ref'] = df_matrix['AE_Ref'].fillna('')
        df_matrix['Row_Key'] = df_matrix.apply(
            lambda x: f"{x['Param']}||{x['AE_Ref']}" if x['AE_Ref'] else x['Param'], 
            axis=1
        )
        
        df_matrix['Time_Unique'] = df_matrix.groupby(['Row_Key', 'Time']).cumcount()
        df_matrix['Time_Label'] = df_matrix.apply(
            lambda x: f"{x['Time']} ({x['Time_Unique'] + 1})" if x['Time_Unique'] > 0 else x['Time'], axis=1
        )
        
        df_pivot = df_matrix.pivot_table(index='Row_Key', columns='Time_Label', values='Value', aggfunc='first') 
        
        win = tk.Toplevel(self.root)
        win.title(f"Data Matrix - Patient {pat}")
        win.geometry("1200x600")
        
        # Store pivot data for export
        self.data_matrix_df = df_pivot
        self.data_matrix_patient = pat
        
        # Top toolbar with export buttons
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self.export_data_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self.export_data_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        tk.Label(toolbar, text="  |", bg="#f4f4f4", fg="#666").pack(side=tk.LEFT, padx=5)
        
        # Hide Units toggle button
        hide_units_var = tk.BooleanVar(value=False)
        
        def toggle_units():
            """Toggle visibility of unit rows in the tree view."""
            hide = hide_units_var.get()
            # Rebuild tree content based on filter
            for item in tree.get_children():
                tree.delete(item)
            
            for row_key in df_pivot.index:
                # Parse composite Row_Key to extract Param and AE Ref
                if "||" in str(row_key):
                    param_name, ae_ref = str(row_key).split("||", 1)
                else:
                    param_name, ae_ref = str(row_key), ""
                
                # Check if this is a unit row: ends with /units, /, or contains 'unit' after last slash
                param_lower = param_name.lower().strip()
                is_unit_row = (param_lower.endswith('/units') or 
                               param_lower.endswith('/') or
                               param_lower.endswith('units') or
                               (('/' in param_lower) and ('unit' in param_lower.split('/')[-1])))
                
                if hide and is_unit_row:
                    continue
                    
                row_vals = [param_name]
                for tc in time_cols:
                    val = df_pivot.at[row_key, tc]
                    row_vals.append(val if pd.notna(val) else "")
                if has_ae_refs:
                    row_vals.append(ae_ref)
                tree.insert("", "end", values=row_vals)
        
        tk.Checkbutton(toolbar, text="Hide Unit Rows", variable=hide_units_var, 
                       command=toggle_units, bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=5)
        
        tk.Label(toolbar, text="  |  Tip: Drag column borders to resize", bg="#f4f4f4", fg="#666", 
                 font=("Segoe UI", 8, "italic")).pack(side=tk.LEFT, padx=10)
        
        # Tree container
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        tree = ttk.Treeview(tree_frame)
        
        def try_parse_date(d_str):
            try:
                base_d = re.sub(r" \(\d+\)$", "", d_str)
                return datetime.fromisoformat(base_d) if len(base_d) > 10 else datetime.max
            except Exception:
                return datetime.max

        time_cols = sorted(df_pivot.columns.tolist(), key=try_parse_date)
        
        # Build AE_Ref lookup from original data using (param, time) as key
        ae_ref_lookup = {}
        for row_data in df_matrix.itertuples(index=False):
            param = getattr(row_data, 'Param', '')
            time_val = getattr(row_data, 'Time', '')
            ae_ref = getattr(row_data, 'AE_Ref', '')
            if param and time_val:
                # Match time to time_cols (may need to handle unique suffixes)
                for tc in [t for t in df_matrix['Time_Label'].unique() if str(t).startswith(str(time_val)[:10])]:
                    ae_ref_lookup[(param, tc)] = ae_ref
                # Also store with raw time
                ae_ref_lookup[(param, time_val)] = ae_ref
        
        # Check if there's any AE data
        has_ae_refs = any(ae_ref_lookup.values())
        
        if has_ae_refs:
            tree["columns"] = ["Parameter"] + time_cols + ["AE Ref"]
        else:
            tree["columns"] = ["Parameter"] + time_cols
        tree.heading("#0", text="") 
        tree.column("#0", width=0, stretch=tk.NO)
        tree.heading("Parameter", text="Parameter")
        tree.column("Parameter", width=300, anchor="w", minwidth=100)
        
        for tc in time_cols:
            tree.heading(tc, text=tc)
            tree.column(tc, width=150, anchor="center", minwidth=80)
        
        if has_ae_refs:
            tree.heading("AE Ref", text="AE Ref")
            tree.column("AE Ref", width=80, anchor="center", minwidth=60)
            
        for row_key in df_pivot.index:
            # Parse composite Row_Key to extract Param and AE Ref
            if "||" in str(row_key):
                param_name, ae_ref = str(row_key).split("||", 1)
            else:
                param_name, ae_ref = str(row_key), ""
            
            row_vals = [param_name]
            for tc in time_cols:
                val = df_pivot.at[row_key, tc]
                row_vals.append(val if pd.notna(val) else "")
            if has_ae_refs:
                row_vals.append(ae_ref)
            tree.insert("", "end", values=row_vals)
            
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)


    def show_visit_schedule(self):
        """Display a visit schedule matrix (delegated to visit_schedule_ui.py)."""
        VisitScheduleWindow(self).show()

    def show_data_gaps(self):
        """Display all missing data (gaps) per patient, organized by visit."""
        DataGapsWindow(self).show()

    def generate_view(self, *args):
        """Wrapper for ViewBuilder.generate_view."""
        self.view_builder.generate_view(*args)


    def load_sdv_data(self):
        """Async Load SDV status from Modular export file."""
        
        # Initialize SDV manager if needed
        if self.sdv_manager is None:
            self.sdv_manager = SDVManager()
        
        # Auto-detect latest Modular file in 'verified' folder
        verified_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), "verified")
        modular_file = None
        
        if os.path.isdir(verified_dir):
            modular_files = [f for f in glob.glob(os.path.join(verified_dir, "*Modular*.xlsx"))
                             if not os.path.basename(f).startswith("~$")]
            if modular_files:
                modular_file = max(modular_files, key=os.path.getmtime)
        
        if not modular_file:
            modular_file = filedialog.askopenfilename(
                title="Select Modular Export File",
                filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
                initialdir=verified_dir if os.path.isdir(verified_dir) else "."
            )
        
        if not modular_file:
            return
        
        # Show loading message and disable button
        self.sdv_btn.config(text="Loading...", bg="#f39c12", state="disabled")
        self.root.update_idletasks()
        
        # Start background thread
        threading.Thread(target=self._load_sdv_thread, args=(modular_file,), daemon=True).start()

    def _load_sdv_thread(self, filepath):
        """Background thread for loading SDV files."""
        
        def progress_callback(stage_text):
            # Schedule button update on main thread
            self.root.after(0, lambda: self.sdv_btn.config(text=stage_text))
        
        try:
            # Load Modular file (field-level status)
            success = self.sdv_manager.load_modular_file(filepath, progress_callback=progress_callback)
            if not success:
                result = (False, "Failed to load Modular file.")
                self.root.after(0, self._on_sdv_loaded, result)
                return
            
            # Auto-detect and load CrfStatusHistory file (form-level status)
            verified_dir = os.path.dirname(filepath)
            crf_files = [f for f in glob.glob(os.path.join(verified_dir, "*CrfStatusHistory*.xlsx")) 
                         if not os.path.basename(f).startswith("~$")]
            if crf_files:
                crf_file = max(crf_files, key=os.path.getmtime)
                self.sdv_manager.load_crf_status_file(crf_file, progress_callback=progress_callback)
            
            result = (True, filepath)
        except Exception as e:
            result = (False, str(e))
            
        # Schedule UI update on main thread
        self.root.after(0, self._on_sdv_loaded, result)

    def _on_sdv_loaded(self, result):
        """Callback when SDV loading completes."""
        self.view_builder.clear_cache()  # Invalidate cache — SDV status changed
        success, data = result
        
        # Re-enable button
        self.sdv_btn.config(state="normal")
        
        if not success:
            messagebox.showerror("SDV Error", f"Error loading SDV: {data}")
            self.sdv_btn.config(text="📋 SDV Check", bg="#27ae60")
            return
            
        modular_file = data
        # Initialize Dashboard Manager
        if self.sdv_manager:
            self.dashboard_manager = DashboardManager(self.sdv_manager)
        
        # Register loaded files with Data Sources manager
        self.data_source_manager.register_loaded_file("modular", modular_file)
        if self.sdv_manager.form_entry_status:
            import glob
            verified_dir = os.path.dirname(modular_file)
            crf_files = [f for f in glob.glob(os.path.join(verified_dir, "*CrfStatusHistory*.xlsx"))
                         if not os.path.basename(f).startswith("~$")]
            if crf_files:
                crf_file = max(crf_files, key=os.path.getmtime)
                self.data_source_manager.register_loaded_file("crf_status", crf_file)
        
        # Get statistics for current patient
        current_patient = self.cb_pat.get()
        if current_patient:
            pat_verified, pat_pending, pat_awaiting, pat_hidden = self.sdv_manager.get_patient_stats(current_patient)
            
            # Refresh tree to show SDV marks
            self.generate_view()
            
            # Show summary
            messagebox.showinfo(
                "SDV Loaded",
                f"SDV data loaded!\n\n"
                f"File: {os.path.basename(modular_file)}\n\n"
                f"Patient {current_patient}:\n"
                f"  ✓ Verified: {pat_verified}\n"
                f"  ! Pending: {pat_pending}\n"
                f"  ? Awaiting: {pat_awaiting}"
            )
            
            self.sdv_btn.config(text=f"SDV ✓{pat_verified} !{pat_pending}", bg="#2ecc71")
        else:
            total_verified, total_pending, _, _ = self.sdv_manager.get_total_stats()
            messagebox.showinfo(
                "SDV Loaded",
                f"SDV data loaded!\n\nTotal: ✓{total_verified} !{total_pending}\n\n"
                "Select a patient to see SDV status."
            )
            self.sdv_btn.config(text="📋 SDV ✓", bg="#2ecc71")




    def export_view(self):
        if not self.tree.get_children():
            messagebox.showinfo("Export", "Nothing to export!")
            return
            
        path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV", "*.csv")])
        if not path: return
        
        data = []
        
        def recurse(parent_id, hierarchy_path):
            for child in self.tree.get_children(parent_id):
                text = self.tree.item(child, "text")
                vals = self.tree.item(child, "values")
                current_path = hierarchy_path + [text]
                
                if vals and len(vals) >= 6: 
                    row = {
                        "Level 1": current_path[0] if len(current_path) > 0 else "",
                        "Level 2": current_path[1] if len(current_path) > 1 else "",
                        "Level 3": current_path[2] if len(current_path) > 2 else "",
                        "Value": vals[0],
                        "Status": vals[1],
                        "SDV": vals[2],
                        "User": vals[3],
                        "Date": vals[4],
                        "DB Variable": vals[5]
                    }
                    data.append(row)

                
                recurse(child, current_path)

        recurse("", [])
        
        if data:
            pd.DataFrame(data).to_csv(path, index=False)
            messagebox.showinfo("Success", f"Exported {len(data)} rows to {path}")

    def export_data_matrix(self, format_type='xlsx'):
        """Export Data Matrix data to XLSX or CSV."""
        if not hasattr(self, 'data_matrix_df') or self.data_matrix_df is None:
            messagebox.showwarning("Export", "No Data Matrix data available.")
            return
        
        patient = getattr(self, 'data_matrix_patient', 'unknown')
        
        if format_type == 'xlsx':
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx", 
                filetypes=[("Excel Files", "*.xlsx")],
                initialfile=f"data_matrix_{patient}.xlsx"
            )
            if not path: return
            try:
                # Reset index to make Parameter a column
                df_export = self.data_matrix_df.reset_index()
                df_export.rename(columns={'index': 'Parameter'}, inplace=True)
                df_export.to_excel(path, index=False, engine='openpyxl')
                messagebox.showinfo("Success", f"Exported to {path}")
            except ImportError:
                messagebox.showerror("Error", "openpyxl is required for XLSX export. Install with: pip install openpyxl")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {e}")
        else:
            path = filedialog.asksaveasfilename(
                defaultextension=".csv", 
                filetypes=[("CSV Files", "*.csv")],
                initialfile=f"data_matrix_{patient}.csv"
            )
            if not path: return
            try:
                df_export = self.data_matrix_df.reset_index()
                df_export.rename(columns={'index': 'Parameter'}, inplace=True)
                df_export.to_csv(path, index=False)
                messagebox.showinfo("Success", f"Exported to {path}")
            except Exception as e:
                df_export.to_csv(path, index=False)
                messagebox.showinfo("Success", f"Exported to {path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to export: {e}")

    # --- Echo Export Feature (delegated to export_dialogs_ui.py) ---
    def show_echo_export(self):
        """Show configuration dialog for Echo Export."""
        EchoExportDialog(self).show()

    def _is_screen_failure(self, patient_id):
        """Check if a patient is a screen failure (has no treatment date)."""
        pat_rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if pat_rows.empty:
            return True
        row = pat_rows.iloc[0]
        # Check for treatment/procedure date
        treatment_date_col = "TV_PR_SVDTC"
        if treatment_date_col in row.index:
            val = row[treatment_date_col]
            if pd.notna(val) and str(val).strip():
                return False
        return True


    # --- CVC Export Feature (delegated to export_dialogs_ui.py) ---
    def show_cvc_export(self):
        """Show CVC Export dialog."""
        CVCExportDialog(self).show()

    # --- Labs Export Feature (delegated to export_dialogs_ui.py) ---
    def show_labs_export(self):
        """Show Labs Export dialog."""
        LabsExportDialog(self).show()

    def show_batch_export(self):
        """Show configuration dialog for Batch Export."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        batch_export.BatchExportDialog(self.root, self)

    def show_data_comparison(self):
        """Show Data Comparison Dialog."""
        data_comparator.DataComparatorDialog(self.root)

    def show_data_sources(self):
        """Show Data Sources management window."""
        DataSourcesWindow(self.root, self.data_source_manager,
                          reload_callback=self._reload_data_source)

    def show_patient_timeline(self):
        """Show Patient Timeline window."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        PatientTimelineWindow(self.root, self.df_main,
                              get_screen_failures_fn=self.get_screen_failures)

    def _reload_data_source(self, source_type: str, filepath: str):
        """Reload a specific data source by type."""
        try:
            if source_type == "project":
                self.load_data(filepath)
                self.data_source_manager.register_loaded_file("project", filepath)
            elif source_type == "modular":
                if self.sdv_manager is None:
                    self.sdv_manager = SDVManager()
                self.sdv_btn.config(text="Loading...", bg="#f39c12", state="disabled")
                self.root.update_idletasks()
                threading.Thread(target=self._load_sdv_thread, args=(filepath,), daemon=True).start()
            elif source_type == "crf_status":
                if self.sdv_manager is None:
                    self.sdv_manager = SDVManager()
                self.sdv_manager.load_crf_status_file(filepath)
                self.data_source_manager.register_loaded_file("crf_status", filepath)
            else:
                # Custom source — just track it, user decides what to do
                self.data_source_manager.register_loaded_file(source_type, filepath)
        except Exception as e:
            self.data_source_manager.register_error(source_type, str(e))
            messagebox.showerror("Load Error", f"Failed to load {source_type}: {e}")


    # --- FU Highlights Feature (delegated to export_dialogs_ui.py) ---
    def show_fu_highlights(self):
        """Show FU Highlights Generator dialog."""
        FUHighlightsDialog(self).show()

    def show_procedure_timing(self, master=None):
        """Show Procedure Timing matrix (delegated to procedure_timing_ui.py)."""
        ProcedureTimingWindow(self).show()


    def on_tree_select(self, event):
        """Handle tree selection to show SDV details."""
        if not self.sdv_manager: return
        sel = self.tree.selection()
        if not sel: return
        
        item_id = sel[0]
        # values tuple: (val, status, sdv, code)
        vals = self.tree.item(item_id, "values")
        if not vals or len(vals) < 4: return
        
        val, status, sdv, code = vals
        pat = self.cb_pat.get().strip()
        
        # Determine Form/Visit from parents navigation
        parent_id = self.tree.parent(item_id)
        if not parent_id: return
        u_parent_id = self.tree.parent(parent_id)
        if not u_parent_id: return
        
        # Remove grid icon from text
        l1 = self.tree.item(u_parent_id, "text").replace(" ▦", "")
        l2 = self.tree.item(parent_id, "text").replace(" ▦", "")
        
        visit_name = ""
        form_name = ""
        
        if self.view_mode.get() == "assess":
             # L1=Form, L2=Field
             form_name = l1
             visit_name = self.tree.item(item_id, "text") 
        else:
             # Visit mode: L1=Visit, L2=Form
             visit_name = l1
             form_name = l2
        
        # Clean names
        visit_name = visit_name.strip()
        form_name = form_name.strip()
        
        # Resolve Repeat/Row
        repeat_num = "0"
        code_str = str(code)
        val_str = str(val)
        
        if "LOGS" in code_str:
             # Try resolving row for Logs/Repeatable forms
             if "AE" in code_str and "TERM" in code_str:
                  res = self.sdv_manager.get_ae_repeat_number(pat, val_str)
                  if res: repeat_num = res
             elif ("LI_PR" in code_str or "OTH" in code_str) and ("TEST" in code_str or "NAM" in code_str):
                  res = self.sdv_manager.get_lab_row_number(pat, val_str)
                  if res: repeat_num = res
        
        # Lookup details
        details = self.sdv_manager.get_verification_details(pat, form_name, visit_name, repeat_num)
        
        if details:
             logger.info(
                 "SDV Info | Variable: %s | Value: %s | Form: %s (Row #%s) | "
                 "Status: %s | User: %s | Date: %s",
                 code, val, form_name, repeat_num,
                 details['status'], details['user'], details['date'],
             )

    def get_screen_failures(self):
        """Return list of patient IDs who are screen failures."""
        if self.df_main is None:
            return []
            
        statuses = self.df_main['Status'].astype(str).str.strip().str.lower()
        mask = statuses.str.contains('screen', na=False) & statuses.str.contains('fail', na=False)
        return self.df_main.loc[mask, 'Screening #'].astype(str).str.strip().tolist()

    def open_dashboard(self):
        """Open the SDV & Data Gap Dashboard."""
        if not self.sdv_manager or not self.sdv_manager.is_loaded():
            messagebox.showwarning("No Data", "Please load SDV data first (usually loaded automatically with project file).")
            return
            
        # PASS LABELS FOR MAPPING (Global Scope)
        if hasattr(self, 'labels') and self.labels:
            self.dashboard_manager.set_labels(self.labels)

        DashboardWindow(self.root, self.dashboard_manager, 
                        get_screen_failures_callback=self.get_screen_failures)

    # -------------------------------------------------------------------------
    # AE Module
    # -------------------------------------------------------------------------
    def show_ae_module(self):
        """Open the AE Module Window."""
        if not hasattr(self, 'ae_manager') or self.ae_manager is None:
             # Try to init if data loaded but not manager
             if self.df_main is not None and self.df_ae is not None:
                 self.ae_manager = AEManager(self.df_main, self.df_ae)
             else:
                messagebox.showwarning("Warning", "No data loaded. Please load an Excel file first.")
                return
        
        AEWindow(self.root, self.ae_manager)

    # -------------------------------------------------------------------------
    # HF Hospitalizations Module (delegated to hf_ui.py)
    # -------------------------------------------------------------------------

    def show_hf_hospitalizations(self):
        """Show HF Hospitalizations summary window."""
        HFWindow(self).show()

    # =========================================================================
    # Assessment Data Table Feature
    # =========================================================================
    
    def show_assessment_data_table(self):
        """Display Assessment Data Table (delegated to assessment_table_ui.py)."""
        AssessmentTableWindow(self).show()


    # =========================================================================
    # Event Handlers (Restored)
    # =========================================================================
    
    def on_double_click(self, event):
        """Handle double click on tree item."""
        item = self.tree.selection()
        if not item: return
        item_id = item[0]
        
        # If item has children, toggle expand/collapse
        if self.tree.get_children(item_id):
            if self.tree.item(item_id, "open"):
                self.tree.item(item_id, open=False)
            else:
                self.tree.item(item_id, open=True)

    def show_context_menu(self, event):
        """Show context menu on right click."""
        # Select item under cursor
        full_row = self.tree.identify_row(event.y)
        if full_row:
            self.tree.selection_set(full_row)
            
            menu = tk.Menu(self.root, tearoff=0)
            menu.add_command(label="Copy Value", command=self.copy_selected_value)
            
            # Check if SDV manager is loaded
            if self.sdv_manager and self.sdv_manager.is_loaded():
                 menu.add_command(label="Verify (SDV)", command=self.verify_selected_item)
            
            menu.tk_popup(event.x_root, event.y_root)
            
    def copy_selected_value(self):
        """Copy selected item value to clipboard."""
        item = self.tree.selection()
        if item:
            val = self.tree.item(item[0], "values")
            if val:
                self.root.clipboard_clear()
                self.root.clipboard_append(str(val[0]))
                
    def verify_selected_item(self):
        """Verify selected item (placeholder)."""
        messagebox.showinfo("SDV", "Verification functionality needs to be re-connected.")



if __name__ == "__main__":
    root = tk.Tk()
    app = ClinicalDataMasterV30(root)
    root.mainloop()