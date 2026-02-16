
import tkinter as tk
import threading
import logging
from tkinter import filedialog, ttk, messagebox
import pandas as pd
import numpy as np
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
from assessment_data_table import AssessmentDataExtractor, ASSESSMENT_CATEGORIES

# Module-level logger
logger = logging.getLogger("ClinicalViewer")


# Domain rules imported from centralized config

from config import VISIT_MAP, ASSESSMENT_RULES, CONDITIONAL_SKIPS, VISIT_SCHEDULE
import batch_export
import data_comparator
from data_sources import DataSourceManager, DataSourcesWindow
from patient_timeline import PatientTimelineWindow
from gap_analysis import DataGapsWindow
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
        self._ae_data_requested = False
        self.current_patient_gaps = []
        self.current_tree_data = {}
        
        # SDV (Source Data Verification) Manager
        self.sdv_manager = None
        self.dashboard_manager = None
        self.hf_manager = None  # HF Hospitalization tracking
        self.sdv_verified_fields = set()  # Set of verified field IDs for current patient
        
        # View Builder
        self.view_builder = ViewBuilder(self)
        
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
            cwd = os.getcwd()
            files = [f for f in os.listdir(cwd) if f.startswith("Innoventric_CLD-048_DM_ProjectToOneFile") and f.endswith(".xlsx")]
            
            if not files:
                return

            # Determine latest by parsing timestamp in filename
            # Format: Innoventric_CLD-048_DM_ProjectToOneFile_04-01-2026_09-45_03_(UTC).xlsx
            latest_file = None
            latest_time = None
            
            for f in files:
                # Extract datetime part. Filename format: ..._DD-MM-YYYY_HH-MM_SS_(UTC).xlsx
                # Supports both dash and underscore separators for flexibility
                match = re.search(r'_(\d{2}-\d{2}-\d{4}_\d{2}-\d{2}[-_]\d{2})_', f)
                if match:
                    dt_str = match.group(1)
                    try:
                        # Try parsing with underscore for seconds (as seen in user file)
                        try:
                            dt = datetime.strptime(dt_str, "%d-%m-%Y_%H-%M_%S")
                        except ValueError:
                            # Fallback to standard dashes
                            dt = datetime.strptime(dt_str, "%d-%m-%Y_%H-%M-%S")
                            
                        if latest_time is None or dt > latest_time:
                            latest_time = dt
                            latest_file = f
                    except ValueError:
                        continue
            
            if latest_file:
                full_path = os.path.join(cwd, latest_file)
                self.load_data(full_path, latest_time)
                
        except Exception as e:
            print(f"Auto-load failed: {e}")

    def load_data(self, path, cutoff_time=None):
        """Load data from specific path."""
        try:
            self.root.config(cursor="watch")
            self.root.update()
            
            self.current_file_path = path
            
            # Update labels
            filename = os.path.basename(path)
            self.file_info_var.set(f"Loaded: {filename}")
            if cutoff_time:
                self.cutoff_var.set(f"Cutoff: {cutoff_time.strftime('%Y-%m-%d %H:%M:%S')}")
            else:
                # Try to parse from filename if not provided
                match = re.search(r'_(\d{2}-\d{2}-\d{4}_\d{2}-\d{2}-\d{2})_', filename)
                if match:
                    try:
                        dt = datetime.strptime(match.group(1), "%d-%m-%Y_%H-%M-%S")
                        self.cutoff_var.set(f"Cutoff: {dt.strftime('%Y-%m-%d %H:%M:%S')}")
                    except Exception:
                        self.cutoff_var.set("")
                else:
                    self.cutoff_var.set("")

            xls = pd.read_excel(path, sheet_name=None, header=None, dtype=str, keep_default_na=False)
            target = next((n for n in xls.keys() if "main" in n.lower()), None)
            if not target:
                messagebox.showerror("Error", "Could not find 'Main' sheet.")
                self.root.config(cursor="")
                return
            
            raw = xls[target]
            # Robustly clean codes and labels (strip whitespace)
            codes = [str(c).strip() for c in raw.iloc[0].tolist()]
            labels = [str(l).strip() for l in raw.iloc[1].tolist()]
            
            self.df_main = raw.iloc[2:].copy()
            self.df_main.columns = codes
            self.labels = dict(zip(codes, labels))
            
            # Load other sheets (AE, CM, etc)
            self._load_extra_sheets(xls)
            
            self._populate_filters()
            self.view_builder.clear_cache()  # Invalidate view cache on new data
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

    def _load_extra_sheets(self, xls):
        """Helper to load auxiliary sheets."""
        # AE Sheet
        ae_sheet = next((n for n in xls.keys() if n.startswith("AE_")), None)
        if ae_sheet:
            try:
                raw_ae = xls[ae_sheet]
                if not raw_ae.empty:
                    new_cols = raw_ae.iloc[0].tolist()
                    self.df_ae = raw_ae.iloc[1:].copy()
                    self.df_ae.columns = new_cols
                    self.df_ae.columns = [str(c).replace('\xa0', ' ') for c in self.df_ae.columns]
                    self.df_ae.reset_index(drop=True, inplace=True)
                else:
                    self.df_ae = None
            except Exception as e:
                print(f"Error loading AE sheet: {e}")
                self.df_ae = None
        
        # CM Sheet
        cm_sheet = next((n for n in xls.keys() if n.startswith("CMTAB")), None)
        if cm_sheet:
            try:
                raw_cm = xls[cm_sheet]
                if not raw_cm.empty:
                    new_cols = raw_cm.iloc[0].tolist()
                    self.df_cm = raw_cm.iloc[1:].copy()
                    self.df_cm.columns = new_cols
                    self.df_cm.columns = [str(c).replace('\xa0', ' ') for c in self.df_cm.columns]
                    self.df_cm.reset_index(drop=True, inplace=True)
                else:
                    self.df_cm = None
            except Exception as e:
                print(f"Error loading CM sheet: {e}")
                self.df_cm = None

        # CVH Sheet (Cardiovascular History)
        cvh_sheet = next((n for n in xls.keys() if n.startswith("CVH_TABLE")), None)
        if cvh_sheet:
            try:
                raw_cvh = xls[cvh_sheet]
                if not raw_cvh.empty:
                    new_cols = raw_cvh.iloc[0].tolist()
                    self.df_cvh = raw_cvh.iloc[1:].copy()
                    self.df_cvh.columns = new_cols
                    self.df_cvh.columns = [str(c).replace('\xa0', ' ') for c in self.df_cvh.columns]
                    self.df_cvh.reset_index(drop=True, inplace=True)
                else:
                    self.df_cvh = None
            except Exception as e:
                print(f"Error loading CVH sheet: {e}")
                self.df_cvh = None
        else:
            self.df_cvh = None

        # ACT Sheet (ACT Lab Results - Heparin and ACT measurements)
        # ACT Sheets (Unscheduled LB_ACT and Scheduled Group* 717)
        act_sheets_names = [n for n in xls.keys() if n.startswith("LB_ACT") or ("Group" in n and "717" in n)]
        
        act_dfs = []
        for sheet_name in act_sheets_names:
            try:
                raw_act = xls[sheet_name]
                if not raw_act.empty:
                    # FIX: Extract headers from Row 0 explicitly
                    headers = raw_act.iloc[0].tolist()
                    cleaned_headers = [str(c).replace('\xa0', ' ').strip() for c in headers]
                    
                    # Create df from Row 1+ (Data starts after header)
                    df_clean = raw_act.iloc[1:].copy()
                    df_clean.columns = cleaned_headers
                    df_clean.reset_index(drop=True, inplace=True)
                    
                    act_dfs.append(df_clean)
                    logger.debug(f"Loaded ACT sheet '{sheet_name}' with {len(df_clean)} rows")
            except Exception as e:
                logger.error(f"Error loading ACT sheet {sheet_name}: {e}")

        if act_dfs:
            self.df_act = pd.concat(act_dfs, ignore_index=True)
            logger.debug(f"Total merged ACT rows: {len(self.df_act)}")
            logger.debug(f"ACT Columns: {list(self.df_act.columns[:20])}...")
        else:
            logger.debug("No ACT sheets loaded")
            self.df_act = None


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

    def add_gap(self, visit, form, field, column, collected_gaps):
        """Standardized way to record a missing value in the gaps report."""
        collected_gaps.append({
            'Visit': visit,
            'Form': form,
            'Field': field,
            'DB Column': column
        })

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
                    self._show_ae_matrix(pat_aes, pat)
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
                    self._show_cm_matrix(pat_cms, pat)
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
                        multiplier, freq_note, override_dose = self._parse_frequency_multiplier(freq_str, freq_oth)
                        
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
            self._show_cm_matrix_from_data(cm_data, pat)
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
            self._show_mh_matrix(mh_data, pat)
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
                        self._show_hfh_matrix(hfh_data, pat)
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
                        self._show_hmeh_matrix(hmeh_data, pat)
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
                    self._show_cvc_matrix(screening_df, pat, "Screening")
                    tables_shown += 1
            
            # Show Treatment (Hemodynamic/Pre-Post) table
            if is_treatment:
                hemo_df = exporter.generate_hemodynamic_table(pat)
                if hemo_df is not None:
                    self._show_cvc_matrix(hemo_df, pat, "Hemodynamic Effect (Pre/Post Procedure)")
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
                        self._show_cvh_matrix(cvh_data, pat)
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
                    print(f"DEBUG: Matrix Filter. Pat raw='{pat}' clean='{pat_clean}'")
                    print(f"DEBUG: Using Scr Col='{scr_col}'")
                    # Filter with strip and regex=False
                    pat_act = self.df_act[self.df_act[scr_col].astype(str).str.strip().str.contains(pat_clean, regex=False, na=False)]
                    print(f"DEBUG: Matrix Rows Found: {len(pat_act)}")
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
                        self._show_act_matrix(act_events, pat)
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

    def _show_cvh_matrix(self, cvh_data, pat):
        """Display Cardiovascular History data as a structured table."""
        import tkinter.ttk as ttk
        from tkinter import messagebox
        
        # Create window
        win = tk.Toplevel(self.root)
        win.title(f"Cardiovascular History - Patient {pat}")
        win.geometry("800x400")
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Cardiovascular History", bg="#f4f4f4", 
                 font=("Segoe UI", 11, "bold"), fg="#8b0000").pack(side=tk.LEFT, padx=10)
        
        # Create treeview
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ["#", "Date", "Type of Intervention", "Intervention"]
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        tree.heading("#", text="#")
        tree.column("#", width=40, anchor="center")
        tree.heading("Date", text="Date")
        tree.column("Date", width=150, anchor="center")
        tree.heading("Type of Intervention", text="Type of Intervention")
        tree.column("Type of Intervention", width=150, anchor="center")
        tree.heading("Intervention", text="Intervention")
        tree.column("Intervention", width=300, anchor="w")
        
        # Insert data
        for i, record in enumerate(cvh_data, 1):
            tree.insert("", "end", values=(
                i,
                record.get('Date', ''),
                record.get('Type of Intervention', ''),
                record.get('Intervention', '')
            ))
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _show_act_matrix(self, act_events, pat):
        """Display ACT/Heparin data as a chronological table."""
        import tkinter.ttk as ttk
        
        # Create window
        win = tk.Toplevel(self.root)
        win.title(f"ACT Lab Results - Patient {pat}")
        win.geometry("600x400")
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="ACT Lab Results (Chronological)", bg="#f4f4f4", 
                 font=("Segoe UI", 11, "bold"), fg="#2c3e50").pack(side=tk.LEFT, padx=10)
        
        # Create treeview
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        columns = ["#", "Event", "Value", "Time"]
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        tree.heading("#", text="#")
        tree.column("#", width=40, anchor="center")
        tree.heading("Event", text="Event")
        tree.column("Event", width=150, anchor="w")
        tree.heading("Value", text="Value")
        tree.column("Value", width=150, anchor="w")
        tree.heading("Time", text="Time")
        tree.column("Time", width=100, anchor="center")
        
        # Configure tags for coloring
        tree.tag_configure('gap', foreground='red', font=('Segoe UI', 9, 'bold'))
        tree.tag_configure('not_done', foreground='blue', font=('Segoe UI', 9, 'italic'))
        tree.tag_configure('ok', foreground='black')
        
        # Insert data
        for i, event in enumerate(act_events, 1):
            status = event.get('Status', 'OK')
            if status == 'GAP':
                tag = 'gap'
            elif status == 'Confirmed':
                tag = 'not_done'
            else:
                tag = 'ok'
            
            tree.insert("", "end", values=(
                i,
                event.get('Event', ''),
                event.get('Value', ''),
                event.get('Time', '')
            ), tags=(tag,))
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _show_ae_matrix(self, pat_aes, pat):
        """Display AE data as a structured table with proper columns."""
        # Debug: Write all column names to a file
        import os
        debug_path = os.path.join(os.path.dirname(self.current_file_path), "debug_ae_cols.txt")
        with open(debug_path, 'w') as f:
            f.write("=== AE Sheet Column Names ===\n")
            for i, col in enumerate(pat_aes.columns):
                f.write(f"  {i}: '{col}'\n")
            f.write("=============================\n")
        
        print(f"DEBUG: AE columns written to {debug_path}")
        
        # Define column mappings - map from expected column names to possible variations in the data
        col_mapping = {
            'AE #': ['Template number', 'AE #', 'AE Number', 'AESEQ'],
            'SAE?': ['LOGS_AE_AESER', 'Is the event SAE?', 'AESER', 'SAE'],
            'AE Term': ['LOGS_AE_AETERM', 'adverse event / term', 'AETERM', 'Term'],
            'Severity': ['LOGS_AE_AESEV', 'Severity', 'AESEV'],
            'Interval': ['LOGS_AE_AEINT', 'Interval', 'AEINT'],
            'Onset Date': ['LOGS_AE_AESTDTC', 'Date of event onset', 'AESTDTC', 'Start Date'],
            'Resolution Date': ['LOGS_AE_AEENDTC', 'Date resolved', 'AEENDTC', 'End Date'],
            'Ongoing': ['LOGS_AE_AEONGO', 'Ongoing', 'AEONGO'],  # Hidden - used to modify Resolution Date
            'Outcome': ['LOGS_AE_AEOUT', 'Outcome', 'AEOUT'],
            'Rel. PKG Trillium': ['LOGS_AE_AEREL1', 'relationship / PKG Trillium', 'AEREL1', 'Rel Trillium'],
            'Rel. Delivery System': ['LOGS_AE_AEREL2', 'relationship / PKG Delivery System', 'AEREL2', 'Rel Delivery'],
            'Rel. Handle': ['LOGS_AE_AEREL3', 'relationship / PKG Handle', 'AEREL3', 'Rel Handle'],
            'Rel. Index Procedure': ['LOGS_AE_AEREL4', 'relationship / index procedure', 'AEREL4', 'Rel Procedure'],
            'AE Description': ['LOGS_AE_AETERM_COMM', 'AE and sequelae / description', 'AETERM_COMM'],
            'SAE Description': ['LOGS_AE_AETERM_COMM1', 'SAE and sequelae / description', 'AETERM_COMM1'],
        }
        
        # Find available columns
        available_cols = {}
        for display_name, possible_names in col_mapping.items():
            for pn in possible_names:
                if pn in pat_aes.columns:
                    available_cols[display_name] = pn
                    break
        
        print("=== Matched Columns ===")
        for display, source in available_cols.items():
            print(f"  {display}: {source}")
        print("======================")
        
        # Build the data for display
        ae_data = []
        for _, ae_row in pat_aes.iterrows():
            row_data = {}
            ongoing_value = False
            
            # First pass: get all values including Ongoing flag
            for display_name, source_col in available_cols.items():
                val = ae_row.get(source_col, '')
                if pd.isna(val) or str(val).lower() == 'nan':
                    val = ''
                else:
                    val = str(val).strip()
                    
                    # Clean up date values - strip time, keep only date
                    if 'Date' in display_name:
                        # Handle various date formats
                        if 'T' in val:
                            val = val.split('T')[0]  # Take only date part before 'T'
                        elif ' ' in val and any(c.isdigit() for c in val.split(' ')[-1]):
                            # If there's a time portion after space (e.g., "2025-02-05 12:30")
                            parts = val.split(' ')
                            if len(parts) > 1 and ':' in parts[-1]:
                                val = ' '.join(parts[:-1])  # Remove time portion
                        # Remove any "time unknown" or similar text
                        val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
                    
                    # Clean up SAE values
                    if display_name == 'SAE?':
                        if val.lower() in ['yes', 'y', '1', 'true']:
                            val = 'Yes'
                        elif val.lower() in ['no', 'n', '0', 'false']:
                            val = 'No'
                    
                    # Check Ongoing flag
                    if display_name == 'Ongoing':
                        ongoing_value = val.lower() in ['yes', 'y', '1', 'true', 'checked']
                
                row_data[display_name] = val
            
            # Apply Ongoing flag to Resolution Date
            if ongoing_value and 'Resolution Date' in row_data:
                row_data['Resolution Date'] = 'Ongoing'
            
            # Only include rows that have at least an AE term
            if row_data.get('AE Term', '').strip():
                ae_data.append(row_data)
        
        if not ae_data:
            messagebox.showinfo("Info", "No valid adverse event terms found.")
            return
        
        # Create window
        win = tk.Toplevel(self.root)
        win.title(f"Adverse Events - Patient {pat}")
        win.geometry("1400x600")
        
        # Store for export
        self.ae_matrix_df = pd.DataFrame(ae_data)
        self.ae_matrix_patient = pat
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self._export_ae_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self._export_ae_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        tk.Label(toolbar, text=f"  |  {len(ae_data)} adverse event(s) found", bg="#f4f4f4", fg="#666", 
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=10)
        
        # Interval filter checkboxes
        tk.Label(toolbar, text="  |  Exclude:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        exclude_screening_var = tk.BooleanVar(value=False)
        exclude_prior_var = tk.BooleanVar(value=False)
        
        # Tree container
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Define display columns - exclude 'Ongoing' since it's now integrated into Resolution Date
        display_columns = [col for col in available_cols.keys() if col != 'Ongoing']
        
        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')
        
        # Column widths based on content type
        col_widths = {
            'AE #': 60, 'SAE?': 50, 'AE Term': 180, 'Severity': 70,
            'Interval': 120, 'Onset Date': 90, 'Resolution Date': 90, 
            'Outcome': 100,
            'Rel. PKG Trillium': 100, 'Rel. Delivery System': 100, 
            'Rel. Handle': 80, 'Rel. Index Procedure': 100, 
            'AE Description': 200, 'SAE Description': 200
        }
        
        for col in display_columns:
            tree.heading(col, text=col)
            width = col_widths.get(col, 100)
            tree.column(col, width=width, anchor="w" if col in ['AE Term', 'AE Description', 'SAE Description'] else "center", minwidth=50)
        
        def refresh_tree():
            """Rebuild tree with filtered data based on interval exclusions."""
            # Clear existing items
            for item in tree.get_children():
                tree.delete(item)
            
            filtered_data = []
            for ae_record in ae_data:
                interval_val = ae_record.get('Interval', '').lower()
                
                # Check if should be excluded
                if exclude_screening_var.get() and 'screening' in interval_val:
                    continue
                # Match "prior to implant" specifically, not "prior to discharge"
                if exclude_prior_var.get() and 'prior to implant' in interval_val:
                    continue
                
                filtered_data.append(ae_record)
            
            # Add filtered rows
            for ae_record in filtered_data:
                values = [ae_record.get(col, '') for col in display_columns]
                tree.insert("", "end", values=values)
            
            # Update count label
            count_label.config(text=f"  |  {len(filtered_data)} adverse event(s) shown")
        
        # Add filter checkboxes (placed after tree is created so refresh_tree can reference it)
        tk.Checkbutton(toolbar, text="During Screening", variable=exclude_screening_var, 
                       command=refresh_tree, bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=2)
        tk.Checkbutton(toolbar, text="Prior to Implant", variable=exclude_prior_var, 
                       command=refresh_tree, bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=2)
        
        # Count label (will be updated by refresh_tree)
        count_label = tk.Label(toolbar, text=f"  |  {len(ae_data)} adverse event(s) shown", bg="#f4f4f4", fg="#666", 
                               font=("Segoe UI", 9))
        count_label.pack(side=tk.LEFT, padx=5)
        
        # Initial population
        refresh_tree()
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _export_ae_matrix(self, fmt):
        """Export AE matrix data to file."""
        if not hasattr(self, 'ae_matrix_df') or self.ae_matrix_df is None:
            messagebox.showerror("Error", "No AE data to export.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"AE_Matrix_{self.ae_matrix_patient}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self.ae_matrix_df.to_excel(path, index=False)
            else:
                self.ae_matrix_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"AE data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _parse_frequency_multiplier(self, freq_str, freq_other_str=""):
        """Parse frequency string and return (multiplier, display_note, override_daily_dose).
        
        Mirrors the logic from FUHighlightsExporter for consistency.
        """
        import re
        
        if not freq_str or str(freq_str).lower() in ['nan', 'none', '']:
            return 1, "", None
        
        freq = str(freq_str).strip().lower()
        
        # Standard frequencies
        if freq in ["once a day", "qd", "od"]:
            return 1, "", None
        elif freq in ["twice a day", "bid"]:
            return 2, "", None
        elif freq in ["3 times a day", "tid"]:
            return 3, "", None
        elif freq in ["4 times a day", "qid"]:
            return 4, "", None
        elif freq in ["every other day", "qod"]:
            return 0.5, "(every 48h)", None
        elif freq == "as needed":
            return None, "PRN", None
        elif freq == "once":
            return 1, "(single dose)", None
        elif freq == "other":
            # Parse the "other" field
            if freq_other_str and str(freq_other_str).lower() not in ['nan', 'none', '']:
                other = str(freq_other_str).strip().lower()
                
                # Check for explicit multiple "mg" dosages
                mg_matches = re.findall(r'(\d+(?:\.\d+)?)\s*mg', other)
                if len(mg_matches) > 1:
                    total_dose = sum(float(m) for m in mg_matches)
                    return None, f"({freq_other_str})", total_dose
                    
                # Every other day check
                if "every other day" in other or "qod" in other:
                     return 0.5, "(every 48h)", None

                # Check for q{N}h pattern
                match = re.match(r'q\s*(\d+)\s*h', other)
                if match:
                    interval_hours = int(match.group(1))
                    if interval_hours > 0:
                        doses_per_day = 24 // interval_hours
                        return doses_per_day, f"(q{interval_hours}h→{doses_per_day}x/d)", None
                
                # Continuous infusion
                if "continuous" in other:
                    return None, "(continuous)", None
                
                return 1, f"({str(freq_other_str).strip()})", None
            return 1, "", None
        
        return 1, "", None

    def _show_cm_matrix(self, pat_cms, pat):
        """Display CM (Concomitant Medications) data as a structured table with proper columns."""
        # Debug: Write all column names to a file
        import os
        debug_path = os.path.join(os.path.dirname(self.current_file_path), "debug_cm_cols.txt")
        with open(debug_path, 'w') as f:
            f.write("=== CM Sheet Column Names ===\n")
            for i, col in enumerate(pat_cms.columns):
                f.write(f"  {i}: '{col}'\n")
            f.write("=============================\n")
        
        print(f"DEBUG: CM columns written to {debug_path}")
        
        # Define header mapping
        header_map = {
             'LOGS_CM_CMTRT': 'Medication',
             'LOGS_CM_CMINDC': 'Indication',
             'LOGS_CM_CMREF_MH': 'MH Reference',
             'LOGS_CM_CMREF_AE': 'AE Reference',
             'LOGS_CM_CMINDC_OTH': 'Indication (Other)',
             'LOGS_CM_CMSTDAT': 'Start Date',
             'LOGS_CM_CMSTDTC': 'Start Date',
             'LOGS_CM_CMENDAT': 'End Date',
             'LOGS_CM_CMENDTC': 'End Date',
             'LOGS_CM_CMDOSE': 'Dose',
             'LOGS_CM_CMDOSU': 'Unit',
             'LOGS_CM_CMROUTE': 'Route',
             'LOGS_CM_CMDOSFRQ': 'Frequency',
             'LOGS_CM_CMDOSFRQ_OTH': 'Frequency (Other)',
             'Screening #': 'Subject',
             'Randomization #': 'Rand #',
             'Initials': 'Initials',
             'Site #': 'Site',
             'Status': 'Status'
        }
        
        # Use all actual columns from the sheet (exclude system columns)
        exclude_cols = ['Row number', 'Form name', 'Form SN']
        
        # Identify special columns - case insensitive
        ongoing_col = next((c for c in pat_cms.columns if 'CMONGO' in c.upper() or 'ONGOING' in c.upper()), None)
        end_date_col = next((c for c in pat_cms.columns if 'CMENDTC' in c.upper() or 'CMENDAT' in c.upper() or 'END DATE' in c.upper()), None)
        
        # Filter display columns - exclude Ongoing column
        display_columns = [col for col in pat_cms.columns 
                          if col not in exclude_cols 
                          and not col.startswith('_')
                          and (col != ongoing_col if ongoing_col else True)]
        
        print(f"=== Using {len(display_columns)} columns from CM sheet ===")
        
        # Build the data for display
        cm_data = []
        # generated_keys tracks used display names to ensure uniqueness
        
        for _, cm_row in pat_cms.iterrows():
            row_data = {}
            is_ongoing = False
            
            # Check ongoing status first
            if ongoing_col:
                ongoing_val = str(cm_row.get(ongoing_col, '')).lower()
                if ongoing_val in ['yes', 'y', '1', 'true', 'checked']:
                    is_ongoing = True
            
            for col in display_columns:
                val = cm_row.get(col, '')
                if pd.isna(val) or str(val).lower() == 'nan':
                    val = ''
                else:
                    val = str(val).strip()
                    
                    # Clean up date values - strip time, keep only date
                    if 'date' in col.lower() or 'dtc' in col.lower() or 'dat' in col.lower():
                        if 'T' in val:
                            val = val.split('T')[0]
                        val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
                
                # Override End Date if ongoing
                if is_ongoing and end_date_col and col == end_date_col:
                    val = "Ongoing"
                
                # Use mapped key
                display_key = header_map.get(col, col)
                
                # Check for duplicates? Since we process row by row, the keys must be consistent across rows
                # We'll rely on generating final_columns list before the loop to resolve duplicates
                row_data[col] = val # Store using original key first, re-map later?
                # Actually, simpler to store by original key then map when inserting to tree?
                # But exporting wants friendly names.
                # Let's use simpler approach: Map keys here. If multiple raw map to same 'Start Date', last one wins in this dict.
                # To prevent data loss, we'll verify unique keys below.
                row_data[display_key] = val
            
            # Only include rows with some data
            if any(v.strip() for v in row_data.values()):
                cm_data.append(row_data)
        
        if not cm_data:
            messagebox.showinfo("Info", "No valid medications found.")
            return

        # Determine final unique column names for Treeview
        final_columns = []
        seen_cols = set()
        for col in display_columns:
            base_name = header_map.get(col, col)
            final_name = base_name
            counter = 2
            while final_name in seen_cols:
                final_name = f"{base_name} ({counter})"
                counter += 1
            seen_cols.add(final_name)
            final_columns.append(final_name)
        
        # Add Daily Dose column (calculated field)
        final_columns.append("Daily Dose")
        
        # Identify columns needed for Daily Dose calculation
        dose_col = next((c for c in pat_cms.columns if 'CMDOSE' in c.upper() and 'DOSU' not in c.upper()), None)
        freq_col = next((c for c in pat_cms.columns if 'CMDOSFRQ' in c.upper() and 'OTH' not in c.upper()), None)
        freq_oth_col = next((c for c in pat_cms.columns if 'CMDOSFRQ_OTH' in c.upper() or 'CMDOSFRQ_OTHER' in c.upper()), None)
        unit_col = next((c for c in pat_cms.columns if 'CMDOSU' in c.upper() or 'UNIT' in c.upper()), None)
            
        # Re-map row data to match unique final columns
        # This is needed because row_data construction above might have overwritten keys if they weren't unique
        # We need to rebuild data with correct unique keys
        
        final_cm_data = []
        for _, cm_row in pat_cms.iterrows():
            row_data = {}
            is_ongoing = False
            
            if ongoing_col:
                ongoing_val = str(cm_row.get(ongoing_col, '')).lower()
                if ongoing_val in ['yes', 'y', '1', 'true', 'checked']:
                    is_ongoing = True
            
            # Iterate using index tracking to match final_columns (excluding Daily Dose which is last)
            for i, col in enumerate(display_columns):
                val = cm_row.get(col, '')
                if pd.isna(val) or str(val).lower() == 'nan':
                    val = ''
                else:
                    val = str(val).strip()
                    if 'date' in col.lower() or 'dtc' in col.lower() or 'dat' in col.lower():
                        if 'T' in val:
                            val = val.split('T')[0]
                        val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
                
                if is_ongoing and end_date_col and col == end_date_col:
                    val = "Ongoing"
                
                # Use the unique key we determined
                final_key = final_columns[i]
                row_data[final_key] = val
            
            # Calculate Daily Dose
            daily_dose_str = ""
            try:
                dose_val = cm_row.get(dose_col, '') if dose_col else ''
                freq_val = cm_row.get(freq_col, '') if freq_col else ''
                freq_oth_val = cm_row.get(freq_oth_col, '') if freq_oth_col else ''
                unit_val = cm_row.get(unit_col, '') if unit_col else ''
                
                if dose_val and not pd.isna(dose_val) and str(dose_val).lower() not in ['nan', 'none', '']:
                    single_dose = float(str(dose_val).strip())
                    multiplier, freq_note, override_dose = self._parse_frequency_multiplier(freq_val, freq_oth_val)
                    
                    if override_dose is not None:
                        # Explicit override from parsed frequency text
                        daily = override_dose
                    elif multiplier is not None:
                        daily = single_dose * multiplier
                    else:
                        daily = None
                    
                    if daily is not None:
                        if daily == int(daily):
                            daily_dose_str = str(int(daily))
                        else:
                            daily_dose_str = f"{daily:.1f}"
                        
                        # Add unit if available
                        if unit_val and not pd.isna(unit_val) and str(unit_val).lower() not in ['nan', 'none', '']:
                            unit_str = str(unit_val).strip()
                            # Normalize common unit variations
                            if 'milligram' in unit_str.lower():
                                unit_str = 'mg'
                            daily_dose_str += f" {unit_str}/day"
                        else:
                            daily_dose_str += "/day"
                    elif freq_note:
                        # PRN or other non-calculable
                        daily_dose_str = f"{int(single_dose) if single_dose == int(single_dose) else single_dose} {freq_note}"
            except (ValueError, TypeError):
                daily_dose_str = ""
            
            row_data["Daily Dose"] = daily_dose_str
            
            if any(v.strip() for v in row_data.values()):
                final_cm_data.append(row_data)

        # Create window
        win = tk.Toplevel(self.root)
        win.title(f"Concomitant Medications - Patient {pat}")
        win.geometry("1400x600")
        
        # Store for export
        self.cm_matrix_df = pd.DataFrame(final_cm_data)
        self.cm_matrix_patient = pat
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self._export_cm_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self._export_cm_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        tk.Label(toolbar, text=f"  |  {len(final_cm_data)} medication(s) found", bg="#f4f4f4", fg="#666", 
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=10)
        
        # Tree container
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Filter out columns that are completely empty
        non_empty_columns = []
        for col in final_columns:
            has_data = any(str(record.get(col, '')).strip() for record in final_cm_data)
            if has_data:
                non_empty_columns.append(col)
        
        tree = ttk.Treeview(tree_frame, columns=non_empty_columns, show='headings')
        
        # Compact column widths to fit on screen
        compact_widths = {
            'Subject': 60, 'Rand #': 50, 'Initials': 50, 'Site': 40, 'Status': 55,
            'Medication': 120, 'Indication': 100, 'MH Reference': 150, 'AE Reference': 100,
            'Indication (Other)': 120, 'Start Date': 75, 'End Date': 70, 
            'Dose': 45, 'Unit': 60, 'Frequency': 80, 'Frequency (Other)': 120,
            'Daily Dose': 80, 'Route': 60
        }
        
        for col in non_empty_columns:
            tree.heading(col, text=col)
            # Use compact width if defined, otherwise auto-size
            width = compact_widths.get(col, min(max(len(col) * 8, 60), 150))
            tree.column(col, width=width, anchor="w", minwidth=40)
        
        # Add rows
        for cm_record in final_cm_data:
            values = [cm_record.get(col, '') for col in non_empty_columns]
            tree.insert("", "end", values=values)
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _export_cm_matrix(self, fmt):
        """Export CM matrix data to file."""
        if not hasattr(self, 'cm_matrix_df') or self.cm_matrix_df is None:
            messagebox.showerror("Error", "No CM data to export.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"CM_Matrix_{self.cm_matrix_patient}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self.cm_matrix_df.to_excel(path, index=False)
            else:
                self.cm_matrix_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"CM data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _show_cm_matrix_from_data(self, cm_data, pat):
        """Display CM data from parsed Main sheet columns as a structured table."""
        if not cm_data:
            messagebox.showinfo("Info", "No valid medications found.")
            return
        
        # Create window
        win = tk.Toplevel(self.root)
        win.title(f"Concomitant Medications - Patient {pat}")
        win.geometry("1400x600")
        
        # Store for export
        self.cm_matrix_df = pd.DataFrame(cm_data)
        self.cm_matrix_patient = pat
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self._export_cm_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self._export_cm_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        tk.Label(toolbar, text=f"  |  {len(cm_data)} medication(s) found", bg="#f4f4f4", fg="#666", 
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=10)
        
        # Tree container
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Define display columns from the first record keys (excluding Ongoing)
        all_cols = ['CM #', 'Medication', 'Dose', 'Dose Unit', 'Frequency', 'Daily Dose', 'Route', 'Indication', 
                    'Start Date', 'End Date']
        
        # Filter to columns that exist in data AND have at least one non-empty value
        display_columns = []
        for col in all_cols:
            exists_in_data = any(col in record for record in cm_data)
            has_data = any(str(record.get(col, '')).strip() for record in cm_data)
            if exists_in_data and has_data:
                display_columns.append(col)
        
        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')
        
        # Compact column widths
        col_widths = {
            'CM #': 40, 'Medication': 140, 'Dose': 50, 'Dose Unit': 70,
            'Route': 80, 'Indication': 140, 'Start Date': 80, 'End Date': 70,
            'Frequency': 90, 'Daily Dose': 80
        }
        
        for col in display_columns:
            tree.heading(col, text=col)
            width = col_widths.get(col, 80)
            tree.column(col, width=width, anchor="w" if col in ['Medication', 'Indication'] else "center", minwidth=40)
        
        # Add rows
        for cm_record in cm_data:
            values = [cm_record.get(col, '') for col in display_columns]
            tree.insert("", "end", values=values)
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _show_mh_matrix(self, mh_data, pat):
        """Display Medical History data as a structured table."""
        if not mh_data:
            messagebox.showinfo("Info", "No valid medical history conditions found.")
            return
        
        # Create window
        win = tk.Toplevel(self.root)
        win.title(f"Medical History - Patient {pat}")
        win.geometry("1100x500")
        
        # Store for export
        self.mh_matrix_df = pd.DataFrame(mh_data)
        self.mh_matrix_patient = pat
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self._export_mh_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self._export_mh_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Tree view
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Define column order (Status excluded - always "Yes" for reported conditions)
        column_order = ['MH #', 'Condition', 'Body System', 'Category', 'Start Date', 'End Date']
        display_columns = [c for c in column_order if c in mh_data[0] or any(c in r for r in mh_data)]
        
        # Ensure all columns from data are included
        for r in mh_data:
            for k in r.keys():
                if k not in display_columns and k != 'Ongoing':  # Exclude Ongoing (internal use)
                    display_columns.append(k)
        
        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')
        
        # Set column widths
        widths = {'MH #': 50, 'Condition': 200, 'Body System': 150, 'Category': 120, 
                  'Start Date': 100, 'End Date': 100}
        for col in display_columns:
            tree.heading(col, text=col)
            width = widths.get(col, 100)
            tree.column(col, width=width, anchor="w", minwidth=50)
        
        # Add rows
        for mh_record in mh_data:
            values = [mh_record.get(col, '') for col in display_columns]
            tree.insert("", "end", values=values)
        
        # Scrollbars
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _export_mh_matrix(self, fmt):
        """Export MH matrix data to file."""
        if not hasattr(self, 'mh_matrix_df') or self.mh_matrix_df is None:
            messagebox.showerror("Error", "No MH data to export.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"MedicalHistory_{self.mh_matrix_patient}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self.mh_matrix_df.to_excel(path, index=False)
            else:
                self.mh_matrix_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Medical History data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _show_hfh_matrix(self, hfh_data, pat):
        """Display Heart Failure History data as a structured table."""
        if not hfh_data:
            messagebox.showinfo("Info", "No valid Heart Failure History data found.")
            return
        
        win = tk.Toplevel(self.root)
        win.title(f"Heart Failure History - Patient {pat}")
        win.geometry("900x400")
        
        self.hfh_matrix_df = pd.DataFrame(hfh_data)
        self.hfh_matrix_patient = pat
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self._export_hfh_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self._export_hfh_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Tree view
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        column_order = ['HFH #', 'Hospitalization Date', 'Details', 'Number of Hospitalizations']
        display_columns = [c for c in column_order if any(c in r for r in hfh_data)]
        
        for r in hfh_data:
            for k in r.keys():
                if k not in display_columns:
                    display_columns.append(k)
        
        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')
        
        widths = {'HFH #': 50, 'Hospitalization Date': 150, 'Details': 300, 'Number of Hospitalizations': 80}
        for col in display_columns:
            tree.heading(col, text=col)
            width = widths.get(col, 120)
            tree.column(col, width=width, anchor="w", minwidth=50)
        
        for record in hfh_data:
            values = [record.get(col, '') for col in display_columns]
            tree.insert("", "end", values=values)
        
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _export_hfh_matrix(self, fmt):
        """Export HFH matrix data to file."""
        if not hasattr(self, 'hfh_matrix_df') or self.hfh_matrix_df is None:
            messagebox.showerror("Error", "No HFH data to export.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"HeartFailureHistory_{self.hfh_matrix_patient}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self.hfh_matrix_df.to_excel(path, index=False)
            else:
                self.hfh_matrix_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Heart Failure History data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _show_hmeh_matrix(self, hmeh_data, pat):
        """Display Hospitalization and Medical Events History data as a structured table."""
        if not hmeh_data:
            messagebox.showinfo("Info", "No valid Hospitalization/Medical Events History data found.")
            return
        
        win = tk.Toplevel(self.root)
        win.title(f"Hospitalization & Medical Events History - Patient {pat}")
        win.geometry("900x450")
        
        self.hmeh_matrix_df = pd.DataFrame(hmeh_data)
        self.hmeh_matrix_patient = pat
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self._export_hmeh_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self._export_hmeh_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Tree view
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        column_order = ['HMEH #', 'Event Date', 'Event Details']
        display_columns = [c for c in column_order if any(c in r for r in hmeh_data)]
        
        for r in hmeh_data:
            for k in r.keys():
                if k not in display_columns:
                    display_columns.append(k)
        
        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')
        
        widths = {'HMEH #': 60, 'Event Date': 120, 'Event Details': 400}
        for col in display_columns:
            tree.heading(col, text=col)
            width = widths.get(col, 120)
            tree.column(col, width=width, anchor="w", minwidth=50)
        
        for record in hmeh_data:
            values = [record.get(col, '') for col in display_columns]
            tree.insert("", "end", values=values)
        
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _export_hmeh_matrix(self, fmt):
        """Export HMEH matrix data to file."""
        if not hasattr(self, 'hmeh_matrix_df') or self.hmeh_matrix_df is None:
            messagebox.showerror("Error", "No HMEH data to export.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"HospMedEventsHistory_{self.hmeh_matrix_patient}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self.hmeh_matrix_df.to_excel(path, index=False)
            else:
                self.hmeh_matrix_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Hospitalization/Medical Events History data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _show_cvc_matrix(self, df, pat, table_type):
        """Display CVC (Cardiac and Venous Catheterization) data as a structured table."""
        if df is None or df.empty:
            messagebox.showinfo("Info", f"No CVC {table_type} data found.")
            return
        
        win = tk.Toplevel(self.root)
        win.title(f"CVC {table_type} - Patient {pat}")
        win.geometry("1000x350")
        
        self.cvc_matrix_df = df
        self.cvc_matrix_patient = pat
        self.cvc_matrix_type = table_type
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(toolbar, text=f"CVC {table_type}", bg="#f4f4f4", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(10, 20))
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX", command=lambda: self._export_cvc_matrix('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV", command=lambda: self._export_cvc_matrix('csv'),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Tree view
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        display_columns = list(df.columns)
        
        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')
        
        # Set column widths based on content
        for col in display_columns:
            tree.heading(col, text=col)
            # Auto-size based on header length
            width = max(len(str(col)) * 10, 80)
            tree.column(col, width=width, anchor="center", minwidth=60)
        
        for values in df[display_columns].fillna('').values:
            tree.insert("", "end", values=list(values))

        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    def _export_cvc_matrix(self, fmt):
        """Export CVC matrix data to file."""
        if not hasattr(self, 'cvc_matrix_df') or self.cvc_matrix_df is None:
            messagebox.showerror("Error", "No CVC data to export.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        table_type = getattr(self, 'cvc_matrix_type', 'CVC')
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"CVC_{table_type}_{self.cvc_matrix_patient}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self.cvc_matrix_df.to_excel(path, index=False)
            else:
                self.cvc_matrix_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"CVC {table_type} data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def show_visit_schedule(self):
        """Display a visit schedule matrix: Patients (rows) x Visit Types (columns) with dates."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        
        # VISIT_SCHEDULE is now defined at module level
        
        # Get all patients - filter to enrolled only (exclude screen failures)
        if 'Status' in self.df_main.columns:
            enrolled_mask = self.df_main['Status'].astype(str).str.lower().isin(['enrolled', 'early withdrawal'])
            patients_df = self.df_main[enrolled_mask]
        else:
            patients_df = self.df_main
        
        patients = patients_df['Screening #'].dropna().unique()
        
        # Build matrix data
        matrix_data = []
        
        for pat_id in sorted(patients):
            row_data = {"Patient": str(pat_id)}
            
            # Get patient row
            pat_rows = self.df_main[self.df_main['Screening #'] == pat_id]
            if pat_rows.empty:
                continue
            pat_row = pat_rows.iloc[0]
            
            # Check if patient died - use LOGS_DTH_DDDTC only
            patient_died = False
            death_date = None
            if 'LOGS_DTH_DDDTC' in self.df_main.columns:
                death_val = pat_row.get('LOGS_DTH_DDDTC')
                if pd.notna(death_val):
                    # Clean value - remove pipes and check if actual date
                    clean_val = str(death_val).replace('|', '').strip()
                    if clean_val and clean_val.lower() not in ['nan', '']:
                        patient_died = True
                        death_date = str(death_val).split('T')[0] if 'T' in str(death_val) else str(death_val)[:10]
            
            # Check for early withdrawal (for future cases where it's not death)
            patient_early_withdrawal = False
            if 'Status' in self.df_main.columns:
                status = str(pat_row.get('Status', '')).lower()
                if 'early withdrawal' in status or 'early-withdrawal' in status:
                    patient_early_withdrawal = True
            
            # Determine end status: Death takes priority over Early Withdrawal
            if patient_died:
                end_status = "Death"
            elif patient_early_withdrawal:
                end_status = "Withdrawn"
            else:
                end_status = None
            
            # Get visit dates
            for date_col, visit_label in VISIT_SCHEDULE:
                if date_col in self.df_main.columns:
                    date_val = pat_row.get(date_col)
                    if pd.notna(date_val) and str(date_val).strip() not in ['', 'nan']:
                        # Format date
                        date_str = str(date_val)
                        if 'T' in date_str:
                            date_str = date_str.split('T')[0]
                        elif len(date_str) > 10:
                            date_str = date_str[:10]
                        row_data[visit_label] = date_str
                    elif end_status:
                        # Show end status (Death or Withdrawn) for missing visits
                        row_data[visit_label] = end_status
                    else:
                        row_data[visit_label] = "Pending"
                else:
                    # Column doesn't exist
                    if end_status:
                        row_data[visit_label] = end_status
                    else:
                        row_data[visit_label] = "Pending"
            
            matrix_data.append(row_data)
        
        if not matrix_data:
            messagebox.showinfo("No Data", "No patient visit data found.")
            return
        
        # Create DataFrame for export
        self.visit_schedule_df = pd.DataFrame(matrix_data)
        
        # Create popup window
        win = tk.Toplevel(self.root)
        win.title("Visit Schedule Matrix")
        win.geometry("1400x700")
        
        # Header
        header = tk.Frame(win, bg="#16a085", padx=10, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="Visit Schedule Matrix - All Patients", 
                 font=("Segoe UI", 14, "bold"), bg="#16a085", fg="white").pack(side=tk.LEFT)
        
        # Export buttons
        btn_frame = tk.Frame(header, bg="#16a085")
        btn_frame.pack(side=tk.RIGHT)
        tk.Button(btn_frame, text="Export XLSX", command=lambda: self._export_visit_schedule('xlsx'),
                  bg="white", fg="#16a085", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="Export CSV", command=lambda: self._export_visit_schedule('csv'),
                  bg="white", fg="#16a085", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=3)
        
        # Summary
        summary_frame = tk.Frame(win, padx=10, pady=5)
        summary_frame.pack(fill=tk.X)
        tk.Label(summary_frame, text=f"Patients: {len(matrix_data)} | Visits: {len(VISIT_SCHEDULE)}",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)
        
        # Tree container with canvas for scrolling
        container = tk.Frame(win)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Build column list
        columns = ["Patient"] + [v[1] for v in VISIT_SCHEDULE]
        
        # Create canvas with scrollbars for cell-level coloring
        canvas = tk.Canvas(container, bg="white")
        v_scroll = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        h_scroll = ttk.Scrollbar(container, orient="horizontal", command=canvas.xview)
        
        scrollable_frame = tk.Frame(canvas, bg="white")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        # Header row
        for col_idx, col_name in enumerate(columns):
            lbl = tk.Label(scrollable_frame, text=col_name, font=("Segoe UI", 9, "bold"),
                          bg="#16a085", fg="white", padx=5, pady=3, width=12, anchor="center")
            lbl.grid(row=0, column=col_idx, sticky="nsew", padx=1, pady=1)
        
        # Data rows with cell-level coloring
        for row_idx, row_data in enumerate(matrix_data, start=1):
            for col_idx, col_name in enumerate(columns):
                value = row_data.get(col_name, "")
                
                # Determine cell color based on value
                if value == "Death":
                    fg_color = "red"
                    bg_color = "#ffe6e6"  # Light red background
                elif value == "Withdrawn":
                    fg_color = "#0066CC"  # Blue
                    bg_color = "#e6f0ff"  # Light blue background
                elif value == "Pending":
                    fg_color = "#CC8800"  # Orange
                    bg_color = "#fff8e6"  # Light yellow background
                else:
                    fg_color = "#228B22"  # Green
                    bg_color = "#e6ffe6"  # Light green background
                
                # Patient column (first column) - neutral colors
                if col_idx == 0:
                    fg_color = "black"
                    bg_color = "#f5f5f5"
                
                lbl = tk.Label(scrollable_frame, text=value, font=("Segoe UI", 9),
                              fg=fg_color, bg=bg_color, padx=5, pady=2, width=12, anchor="center")
                lbl.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)
        
        # Grid layout
        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)
        
        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _export_visit_schedule(self, fmt):
        """Export visit schedule data to file."""
        if not hasattr(self, 'visit_schedule_df') or self.visit_schedule_df is None:
            messagebox.showwarning("No Data", "No visit schedule data to export.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
            title=f"Export Visit Schedule as {ext.upper()}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self.visit_schedule_df.to_excel(path, index=False)
            else:
                self.visit_schedule_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Visit schedule exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def show_data_gaps(self):
        """Display all missing data (gaps) per patient, organized by visit."""
        DataGapsWindow(self).show()

    def generate_view(self, *args):
        """Wrapper for ViewBuilder.generate_view."""
        self.view_builder.generate_view(*args)


    def load_sdv_data(self):
        """Async Load SDV status from Modular export file."""
        import os
        import glob
        
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
        import os
        import glob
        
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
        import os
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

    # --- Echo Export Feature ---
    def show_echo_export(self):
        """Show configuration dialog for Echo Export."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.root)
        win.title("Export Echo Data (Sponsor)")
        win.geometry("600x850")
        win.configure(bg="#f4f4f4")
        
        # Store window reference for helper functions
        self._echo_win = win
        
        # 1. Template Section
        tpl_frame = tk.LabelFrame(win, text=" Template ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        tpl_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        inner_tpl = tk.Frame(tpl_frame, bg="#f4f4f4")
        inner_tpl.pack(fill=tk.X)
        
        # Default to echo_tmpl.xlsx in app folder
        default_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "echo_tmpl.xlsx")
        self.tpl_path_var = tk.StringVar()
        if os.path.exists(default_template):
            self.tpl_path_var.set(default_template)
        
        tk.Entry(inner_tpl, textvariable=self.tpl_path_var, width=50, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(inner_tpl, text="Browse...", command=self.browse_template,
                  bg="#2c3e50", fg="white", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        
        # 2. Patient Filter Section
        pat_filter_frame = tk.LabelFrame(win, text=" Patient Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_filter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Filter options row
        filter_row = tk.Frame(pat_filter_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 8))
        
        self.exclude_screen_failures_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self.exclude_screen_failures_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=self._update_patient_list).pack(side=tk.LEFT)
        
        # Select All / None buttons
        btn_row = tk.Frame(filter_row, bg="#f4f4f4")
        btn_row.pack(side=tk.RIGHT)
        tk.Button(btn_row, text="Select All", command=lambda: self._select_all_patients(True),
                  bg="#27ae60", fg="white", font=("Segoe UI", 8, "bold"), padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Select None", command=lambda: self._select_all_patients(False),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 8, "bold"), padx=8).pack(side=tk.LEFT, padx=2)
        
        # Patient list with scrollbar
        list_frame = tk.Frame(pat_filter_frame, bd=1, relief="sunken", bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(list_frame, bg="white", highlightthickness=0)
        sb = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self._pat_scroll_frame = tk.Frame(canvas, bg="white")
        
        self._pat_scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._pat_scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Store patient data and populate
        self.pat_chk_vars = {}
        self._all_patients = sorted(self.df_main['Screening #'].dropna().unique())
        self._update_patient_list()
        
        # 3. Visit Selection Section  
        vis_frame = tk.LabelFrame(win, text=" Visit Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        vis_frame.pack(fill=tk.X, padx=10, pady=5)
        
        # Create grid for visits (2 columns)
        vis_inner = tk.Frame(vis_frame, bg="#f4f4f4")
        vis_inner.pack(fill=tk.X)
        
        self.vis_chk_vars = {}
        for i, v in enumerate(ECHO_VISITS):
            var = tk.BooleanVar(value=True)
            self.vis_chk_vars[v] = var
            row, col = divmod(i, 2)
            chk = tk.Checkbutton(vis_inner, text=v, variable=var, bg="#f4f4f4", font=("Segoe UI", 9), anchor="w")
            chk.grid(row=row, column=col, sticky="w", padx=(0, 20))
        
        # 4. Output Options Section
        opt_frame = tk.LabelFrame(win, text=" Output Options ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        opt_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(opt_frame, text="For visits without data:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        
        self.delete_empty_rows_var = tk.BooleanVar(value=True)
        tk.Radiobutton(opt_frame, text="Delete row from output (recommended)", variable=self.delete_empty_rows_var, 
                       value=True, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)
        tk.Radiobutton(opt_frame, text="Leave blank row in output", variable=self.delete_empty_rows_var, 
                       value=False, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)
        
        # 5. Generate Button
        tk.Button(win, text="Generate Echo Reports (ZIP)", command=lambda: self.generate_echo_export(win),
                  bg="#2980b9", fg="white", font=("Segoe UI", 11, "bold"), pady=12, cursor="hand2").pack(fill=tk.X, padx=10, pady=15)
        
        # Focus the window
        win.focus_set()
        win.grab_set()

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

    def _update_patient_list(self):
        """Update patient checkbox list based on filter settings."""
        # Clear existing checkboxes
        for widget in self._pat_scroll_frame.winfo_children():
            widget.destroy()
        
        # Determine which patients to show
        exclude_sf = self.exclude_screen_failures_var.get()
        
        for p in self._all_patients:
            if not p:
                continue
            if exclude_sf and self._is_screen_failure(p):
                continue
            
            # Preserve previous selection state if exists
            if p in self.pat_chk_vars:
                var = self.pat_chk_vars[p]
            else:
                var = tk.BooleanVar(value=True)
                self.pat_chk_vars[p] = var
            
            tk.Checkbutton(self._pat_scroll_frame, text=p, variable=var, 
                          bg="white", font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)

    def _select_all_patients(self, select):
        """Select or deselect all visible patients."""
        for widget in self._pat_scroll_frame.winfo_children():
            if isinstance(widget, tk.Checkbutton):
                # Get the patient ID from the text
                pat_id = widget.cget("text")
                if pat_id in self.pat_chk_vars:
                    self.pat_chk_vars[pat_id].set(select)
                   
    def browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.tpl_path_var.set(path)

    def generate_echo_export(self, dialog):
        tpl_path = self.tpl_path_var.get()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showerror("Error", "Please select a valid template file.")
            return
            
        # Get selected patients (only those currently visible)
        selected_pats = []
        for p in self._all_patients:
            if p in self.pat_chk_vars and self.pat_chk_vars[p].get():
                # Also check if visible (not filtered out)
                exclude_sf = self.exclude_screen_failures_var.get()
                if exclude_sf and self._is_screen_failure(p):
                    continue
                selected_pats.append(p)
        
        selected_visits = [v for v, var in self.vis_chk_vars.items() if var.get()]
        delete_empty = self.delete_empty_rows_var.get()
        
        if not selected_pats:
            messagebox.showwarning("Warning", "No patients selected.")
            return

        try:
            dialog.config(cursor="watch")
            dialog.update()
            
            exporter = EchoExporter(self.df_main, tpl_path, self.labels)
            export_data, extension, patient_id = exporter.generate_export(selected_pats, selected_visits, delete_empty)
            
            if not export_data:
                messagebox.showinfo("Info", "No data found for selected criteria.")
                return
            
            # Set default filename based on single or multiple patients
            if extension == 'xlsx':
                default_filename = f"{patient_id}.xlsx"
                filetypes = [("Excel Files", "*.xlsx")]
                title_suffix = "Excel File"
            else:
                default_filename = "echo_reports.zip"
                filetypes = [("ZIP Files", "*.zip")]
                title_suffix = "ZIP Archive"
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=f".{extension}",
                filetypes=filetypes,
                initialfile=default_filename,
                title=f"Save Echo Reports as {title_suffix}"
            )
            
            if save_path:
                with open(save_path, "wb") as f:
                    f.write(export_data)
                messagebox.showinfo("Success", f"Echo reports saved to:\n{save_path}\n\n{len(selected_pats)} patient(s) exported.")
                dialog.destroy()
                
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
        finally:
            try:
                dialog.config(cursor="")
            except Exception:
                pass

    # --- CVC Export Feature ---
    def show_cvc_export(self):
        """Show configuration dialog for CVC (Cardiac and Venous Catheterization) Export."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.root)
        win.title("Export CVC Data")
        win.geometry("500x600")
        win.configure(bg="#f4f4f4")
        
        # 1. Patient Selection Section
        pat_frame = tk.LabelFrame(win, text=" Patient Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))
        
        # Filter options row
        filter_row = tk.Frame(pat_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 8))
        
        self._cvc_exclude_sf_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self._cvc_exclude_sf_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=lambda: self._update_cvc_patient_list(scroll_frame, all_patients)).pack(side=tk.LEFT)
        
        # Patient list with scrollbar
        list_frame = tk.Frame(pat_frame, bd=1, relief="sunken", bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(list_frame, bg="white", highlightthickness=0)
        sb = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas, bg="white")
        
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            try:
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Populate patient list
        self._cvc_pat_vars = {}
        self._cvc_all_patients = sorted(self.df_main['Screening #'].dropna().unique())
        all_patients = self._cvc_all_patients
        for pid in all_patients:
            var = tk.BooleanVar(value=True)
            self._cvc_pat_vars[pid] = var
            tk.Checkbutton(scroll_frame, text=str(pid), variable=var, bg="white", 
                          font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)
        
        # Select All / None buttons
        btn_row = tk.Frame(pat_frame, bg="#f4f4f4")
        btn_row.pack(fill=tk.X, pady=(5, 0))
        tk.Button(btn_row, text="Select All", command=lambda: [v.set(True) for v in self._cvc_pat_vars.values()],
                  bg="#27ae60", fg="white", font=("Segoe UI", 8, "bold"), padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Select None", command=lambda: [v.set(False) for v in self._cvc_pat_vars.values()],
                  bg="#e74c3c", fg="white", font=("Segoe UI", 8, "bold"), padx=8).pack(side=tk.LEFT, padx=2)
        
        # 2. Table Type Section
        table_frame = tk.LabelFrame(win, text=" Table Type ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        table_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self._cvc_screening_var = tk.BooleanVar(value=True)
        self._cvc_hemodynamic_var = tk.BooleanVar(value=True)
        
        tk.Checkbutton(table_frame, text="Screening Table (Right Heart Catheterization)", 
                      variable=self._cvc_screening_var, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Checkbutton(table_frame, text="Hemodynamic Effect Table (Pre/Post Trillium)", 
                      variable=self._cvc_hemodynamic_var, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        
        # 3. Export Format Section
        format_frame = tk.LabelFrame(win, text=" Export Format ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        format_frame.pack(fill=tk.X, padx=10, pady=5)
        
        self._cvc_format_var = tk.StringVar(value="xlsx")
        tk.Radiobutton(format_frame, text="Excel (.xlsx)", variable=self._cvc_format_var, 
                       value="xlsx", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Radiobutton(format_frame, text="CSV (.csv)", variable=self._cvc_format_var, 
                       value="csv", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        
        # 4. Generate Button
        tk.Button(win, text="Generate CVC Export", command=lambda: self.generate_cvc_export(win),
                  bg="#e67e22", fg="white", font=("Segoe UI", 11, "bold"), pady=12, cursor="hand2").pack(fill=tk.X, padx=10, pady=15)
        
        win.focus_set()
        win.grab_set()

    def _update_cvc_patient_list(self, scroll_frame, all_patients):
        """Update CVC patient list based on screen failure filter."""
        # Clear existing checkboxes
        for widget in scroll_frame.winfo_children():
            widget.destroy()
        
        # Rebuild with filter
        exclude_sf = self._cvc_exclude_sf_var.get()
        for pid in all_patients:
            if exclude_sf and self._is_screen_failure(pid):
                continue
            if pid not in self._cvc_pat_vars:
                self._cvc_pat_vars[pid] = tk.BooleanVar(value=True)
            tk.Checkbutton(scroll_frame, text=str(pid), variable=self._cvc_pat_vars[pid], 
                          bg="white", font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)

    def generate_cvc_export(self, dialog):
        """Generate and save CVC export files."""
        # Get selected patients, filtering screen failures if checkbox is checked
        exclude_sf = getattr(self, '_cvc_exclude_sf_var', None) and self._cvc_exclude_sf_var.get()
        selected_pats = []
        for pid, var in self._cvc_pat_vars.items():
            if var.get():
                if exclude_sf and self._is_screen_failure(pid):
                    continue
                selected_pats.append(pid)
        
        if not selected_pats:
            messagebox.showwarning("No Selection", "Please select at least one patient.")
            return
        
        include_screening = self._cvc_screening_var.get()
        include_hemodynamic = self._cvc_hemodynamic_var.get()
        
        if not include_screening and not include_hemodynamic:
            messagebox.showwarning("No Table", "Please select at least one table type.")
            return
        
        export_format = self._cvc_format_var.get()
        
        try:
            dialog.config(cursor="wait")
            dialog.update()
            
            exporter = CVCExporter(self.df_main)
            
            if export_format == "xlsx":
                # Export to Excel - each patient in separate sheets or files
                if len(selected_pats) == 1:
                    # Single patient - single file
                    pid = selected_pats[0]
                    data = exporter.export_to_excel(pid, include_screening, include_hemodynamic)
                    
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".xlsx",
                        filetypes=[("Excel files", "*.xlsx")],
                        initialfile=f"CVC_{pid}.xlsx",
                        title="Save CVC Export"
                    )
                    
                    if save_path:
                        with open(save_path, "wb") as f:
                            f.write(data)
                        messagebox.showinfo("Success", f"CVC data exported to:\n{save_path}")
                        dialog.destroy()
                else:
                    # Multiple patients - ZIP file
                    import zipfile
                    from io import BytesIO
                    
                    zip_buffer = BytesIO()
                    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
                        for pid in selected_pats:
                            data = exporter.export_to_excel(pid, include_screening, include_hemodynamic)
                            if data:
                                zf.writestr(f"CVC_{pid}.xlsx", data)
                    
                    save_path = filedialog.asksaveasfilename(
                        defaultextension=".zip",
                        filetypes=[("ZIP files", "*.zip")],
                        initialfile="CVC_Export.zip",
                        title="Save CVC Export (ZIP)"
                    )
                    
                    if save_path:
                        with open(save_path, "wb") as f:
                            f.write(zip_buffer.getvalue())
                        messagebox.showinfo("Success", f"CVC data exported to:\n{save_path}\n\n{len(selected_pats)} patient(s) exported.")
                        dialog.destroy()
            
            else:
                # CSV export - separate files per patient and table type
                save_dir = filedialog.askdirectory(title="Select folder for CSV exports")
                
                if save_dir:
                    count = 0
                    for pid in selected_pats:
                        if include_screening:
                            csv_data = exporter.export_to_csv(pid, "screening")
                            if csv_data:
                                with open(os.path.join(save_dir, f"CVC_Screening_{pid}.csv"), "w", encoding="utf-8") as f:
                                    f.write(csv_data)
                                count += 1
                        
                        if include_hemodynamic:
                            csv_data = exporter.export_to_csv(pid, "hemodynamic")
                            if csv_data:
                                with open(os.path.join(save_dir, f"CVC_Hemodynamic_{pid}.csv"), "w", encoding="utf-8") as f:
                                    f.write(csv_data)
                                count += 1
                    
                    messagebox.showinfo("Success", f"CVC data exported to:\n{save_dir}\n\n{count} file(s) created.")
                    dialog.destroy()
                    
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
        finally:
            try:
                dialog.config(cursor="")
            except Exception:
                pass

    def show_labs_export(self):
        """Show configuration dialog for Labs Export."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.root)
        win.title("Export Labs Data")
        win.geometry("600x650")
        win.configure(bg="#f4f4f4")
        
        # Store window reference
        self._labs_win = win
        
        # 1. Template Section
        tpl_frame = tk.LabelFrame(win, text=" Template ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        tpl_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        inner_tpl = tk.Frame(tpl_frame, bg="#f4f4f4")
        inner_tpl.pack(fill=tk.X)
        
        self.labs_tpl_path_var = tk.StringVar()
        # Pre-select labs_tmpl.xlsx if exists
        cwd = os.getcwd()
        default_tpl = os.path.join(cwd, "labs_tmpl.xlsx")
        if os.path.exists(default_tpl):
            self.labs_tpl_path_var.set(default_tpl)
            
        tk.Entry(inner_tpl, textvariable=self.labs_tpl_path_var, width=50, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(inner_tpl, text="Browse...", command=self.browse_labs_template,
                  bg="#2c3e50", fg="white", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        
        # 2. Patient Filter Section
        pat_filter_frame = tk.LabelFrame(win, text=" Patient Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_filter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Filter options row
        filter_row = tk.Frame(pat_filter_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 8))
        
        self.labs_exclude_screen_failures_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self.labs_exclude_screen_failures_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=self._update_labs_patient_list).pack(side=tk.LEFT)
        
        # Select All / None buttons
        btn_row = tk.Frame(filter_row, bg="#f4f4f4")
        btn_row.pack(side=tk.RIGHT)
        tk.Button(btn_row, text="Select All", command=lambda: self._select_all_labs_patients(True),
                  bg="#27ae60", fg="white", font=("Segoe UI", 8, "bold"), padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Select None", command=lambda: self._select_all_labs_patients(False),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 8, "bold"), padx=8).pack(side=tk.LEFT, padx=2)
        
        # Patient list with scrollbar
        list_frame = tk.Frame(pat_filter_frame, bd=1, relief="sunken", bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        canvas = tk.Canvas(list_frame, bg="white", highlightthickness=0)
        sb = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self._labs_pat_scroll_frame = tk.Frame(canvas, bg="white")
        
        self._labs_pat_scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._labs_pat_scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Enable mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Store patient data and populate
        self.labs_pat_chk_vars = {}
        self._labs_all_patients = sorted(self.df_main['Screening #'].dropna().unique())
        self._update_labs_patient_list()
        
        # 3. Output Options Section
        opt_frame = tk.LabelFrame(win, text=" Output Options ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        opt_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(opt_frame, text="For days without data:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        
        self.labs_delete_empty_cols_var = tk.BooleanVar(value=True)
        tk.Radiobutton(opt_frame, text="Delete column from output (recommended)", variable=self.labs_delete_empty_cols_var, 
                       value=True, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)
        tk.Radiobutton(opt_frame, text="Leave blank column in output", variable=self.labs_delete_empty_cols_var, 
                       value=False, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)
        
        # Add separator
        tk.Frame(opt_frame, height=8, bg="#f4f4f4").pack(fill=tk.X)
        
        # Highlight out-of-range values option
        self.labs_highlight_out_of_range_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_frame, text="Highlight out-of-range values in red", 
                       variable=self.labs_highlight_out_of_range_var,
                       bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        
        # 4. Buttons: Preview and Generate
        btn_frame = tk.Frame(win, bg="#f4f4f4")
        btn_frame.pack(fill=tk.X, padx=10, pady=15)
        
        tk.Button(btn_frame, text="Preview (Single Patient)", command=lambda: self.preview_labs_export(win),
                  bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"), pady=10, cursor="hand2").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))
        
        tk.Button(btn_frame, text="Generate Labs Reports", command=lambda: self.generate_labs_export(win),
                  bg="#9b59b6", fg="white", font=("Segoe UI", 10, "bold"), pady=10, cursor="hand2").pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(5, 0))
        
        # Focus the window
        win.focus_set()
        win.grab_set()

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

    def _update_labs_patient_list(self):
        """Update patient checkbox list for labs export."""
        for widget in self._labs_pat_scroll_frame.winfo_children():
            widget.destroy()
        
        exclude_sf = self.labs_exclude_screen_failures_var.get()
        
        for p in self._labs_all_patients:
            if not p:
                continue
            if exclude_sf and self._is_screen_failure(p):
                continue
            
            if p in self.labs_pat_chk_vars:
                var = self.labs_pat_chk_vars[p]
            else:
                var = tk.BooleanVar(value=True)
                self.labs_pat_chk_vars[p] = var
            
            tk.Checkbutton(self._labs_pat_scroll_frame, text=p, variable=var, 
                          bg="white", font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)

    def _select_all_labs_patients(self, select):
        """Select or deselect all visible patients for labs export."""
        for widget in self._labs_pat_scroll_frame.winfo_children():
            if isinstance(widget, tk.Checkbutton):
                pat_id = widget.cget("text")
                if pat_id in self.labs_pat_chk_vars:
                    self.labs_pat_chk_vars[pat_id].set(select)
                   
    def browse_labs_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self.labs_tpl_path_var.set(path)

    def ask_unit_resolution(self, param_name, found_units, patient_id):
        """Callback to resolve unit conflicts."""
        # Use a variable to store result from dialog
        self.resolved_unit = found_units[0] if found_units else None
        
        # Create custom dialog
        top = tk.Toplevel(self.root)
        top.title("Unit Conflict Detected")
        top.geometry("400x250")
        top.transient(self.root)
        top.grab_set()
        
        tk.Label(top, text="Unit Conflict Detected", font=("Segoe UI", 11, "bold"), fg="#e74c3c").pack(pady=10)
        
        msg = f"Parameter: {param_name}\nPatient: {patient_id}\n\nExisting units found: {', '.join(found_units)}"
        tk.Label(top, text=msg, justify=tk.LEFT, padx=20).pack(fill=tk.X)
        
        tk.Label(top, text="Select target unit (all values will be converted):", font=("Segoe UI", 9, "bold")).pack(pady=(15, 5))
        
        cmb = ttk.Combobox(top, values=found_units, state="readonly", width=20)
        cmb.set(found_units[0])
        cmb.pack(pady=5)
        
        def on_ok():
            self.resolved_unit = cmb.get()
            top.destroy()
            
        tk.Button(top, text="Convert / Proceed", command=on_ok, bg="#2c3e50", fg="white", width=20).pack(pady=20)
        
        # Center dialog
        self.root.update_idletasks()
        try:
            x = self.root.winfo_x() + (self.root.winfo_width() // 2) - 200
            y = self.root.winfo_y() + (self.root.winfo_height() // 2) - 125
            top.geometry(f"+{x}+{y}")
        except Exception:
            pass
        
        top.wait_window()
        return self.resolved_unit

    def preview_labs_export(self, dialog):
        """Generate a temp file for a single selected patient and open in Excel for preview."""
        import tempfile
        import subprocess
        
        tpl_path = self.labs_tpl_path_var.get()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showerror("Error", "Please select a valid template file.")
            return
        
        # Get first selected patient only
        selected_pat = None
        for p in self._labs_all_patients:
            if p in self.labs_pat_chk_vars and self.labs_pat_chk_vars[p].get():
                exclude_sf = self.labs_exclude_screen_failures_var.get()
                if exclude_sf and self._is_screen_failure(p):
                    continue
                selected_pat = p
                break
        
        if not selected_pat:
            messagebox.showwarning("Warning", "Please select at least one patient for preview.")
            return
        
        delete_empty = self.labs_delete_empty_cols_var.get()
        highlight_oor = self.labs_highlight_out_of_range_var.get()
        
        try:
            dialog.config(cursor="watch")
            dialog.update()
            
            exporter = LabsExporter(self.df_main, tpl_path, self.labels, 
                                    unit_callback=self.ask_unit_resolution,
                                    highlight_out_of_range=highlight_oor)
            export_data = exporter.process_patient(selected_pat, delete_empty)
            
            if not export_data:
                messagebox.showinfo("Info", f"No data found for patient {selected_pat}.")
                return
            
            # Save to temp file
            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"{selected_pat}_labs_preview.xlsx")
            
            with open(temp_path, "wb") as f:
                f.write(export_data)
            
            # Open in default application (Excel)
            try:
                os.startfile(temp_path)  # Windows
            except AttributeError:
                subprocess.run(["open", temp_path])  # macOS
            except Exception:
                subprocess.run(["xdg-open", temp_path])  # Linux
            
            messagebox.showinfo("Preview", f"Preview opened for patient: {selected_pat}\n\nFile: {temp_path}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Preview failed: {e}")
        finally:
            try:
                dialog.config(cursor="")
            except Exception:
                pass


    def generate_labs_export(self, dialog):
        tpl_path = self.labs_tpl_path_var.get()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showerror("Error", "Please select a valid template file.")
            return
            
        # Get selected patients
        selected_pats = []
        for p in self._labs_all_patients:
            if p in self.labs_pat_chk_vars and self.labs_pat_chk_vars[p].get():
                exclude_sf = self.labs_exclude_screen_failures_var.get()
                if exclude_sf and self._is_screen_failure(p):
                    continue
                selected_pats.append(p)
        
        delete_empty = self.labs_delete_empty_cols_var.get()
        highlight_oor = self.labs_highlight_out_of_range_var.get()
        
        if not selected_pats:
            messagebox.showwarning("Warning", "No patients selected.")
            return

        try:
            dialog.config(cursor="watch")
            dialog.update()
            
            exporter = LabsExporter(self.df_main, tpl_path, self.labels, 
                                    unit_callback=self.ask_unit_resolution,
                                    highlight_out_of_range=highlight_oor)
            export_data, extension, patient_id = exporter.generate_export(selected_pats, delete_empty)
            
            if not export_data:
                messagebox.showinfo("Info", "No data found for selected criteria.")
                return
            
            # Set default filename
            if extension == 'xlsx':
                default_filename = f"{patient_id}_labs.xlsx"
                filetypes = [("Excel Files", "*.xlsx")]
                title_suffix = "Excel File"
            else:
                default_filename = "labs_reports.zip"
                filetypes = [("ZIP Files", "*.zip")]
                title_suffix = "ZIP Archive"
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=f".{extension}",
                filetypes=filetypes,
                initialfile=default_filename,
                title=f"Save Labs Reports as {title_suffix}"
            )
            
            if save_path:
                with open(save_path, "wb") as f:
                    f.write(export_data)
                messagebox.showinfo("Success", f"Labs reports saved to:\n{save_path}\n\n{len(selected_pats)} patient(s) exported.")
                dialog.destroy()
                
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
        finally:
            try:
                dialog.config(cursor="")
            except:
                pass

    def show_fu_highlights(self):
        """Show dialog to generate FU Highlights tables for copy-paste."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        
        win = tk.Toplevel(self.root)
        win.title("FU Highlights Generator")
        win.geometry("700x600")
        win.configure(bg="#f4f4f4")
        
        # 1. Patient Selection
        pat_frame = tk.LabelFrame(win, text=" Patient ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        tk.Label(pat_frame, text="Select Patient:", bg="#f4f4f4").pack(side=tk.LEFT, padx=5)
        all_patients = sorted(self.df_main['Screening #'].dropna().unique())
        
        # Filter out screen failures if checkbox is checked
        self.fu_exclude_sf_var = tk.BooleanVar(value=False)
        
        def update_patient_list():
            if self.fu_exclude_sf_var.get():
                filtered = [p for p in all_patients if not self._is_screen_failure(p)]
            else:
                filtered = all_patients
            pat_combo['values'] = filtered
            if filtered and (not self.fu_patient_var.get() or self.fu_patient_var.get() not in filtered):
                self.fu_patient_var.set(filtered[0])
        
        self.fu_patient_var = tk.StringVar()
        if all_patients:
            self.fu_patient_var.set(all_patients[0])
        pat_combo = ttk.Combobox(pat_frame, textvariable=self.fu_patient_var, values=all_patients, state="readonly", width=20)
        pat_combo.pack(side=tk.LEFT, padx=5)
        
        tk.Checkbutton(pat_frame, text="Exclude Screen Failures", variable=self.fu_exclude_sf_var, 
                       bg="#f4f4f4", font=("Segoe UI", 9), command=update_patient_list).pack(side=tk.LEFT, padx=15)
        
        # 2. Visit Selection
        vis_frame = tk.LabelFrame(win, text=" Follow-up Visit ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        vis_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(vis_frame, text="Select Visit:", bg="#f4f4f4").pack(side=tk.LEFT, padx=5)
        fu_visits = ["30D", "6M", "1Y", "2Y", "4Y"]
        self.fu_visit_var = tk.StringVar(value="30D")
        vis_combo = ttk.Combobox(vis_frame, textvariable=self.fu_visit_var, values=fu_visits, state="readonly", width=15)
        vis_combo.pack(side=tk.LEFT, padx=5)
        
        tk.Button(vis_frame, text="Generate Tables", command=lambda: self._generate_fu_tables(win),
                  bg="#1abc9c", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=20)
        
        # 3. Clinical Parameters Table
        params_frame = tk.LabelFrame(win, text=" Clinical Parameters ", padx=5, pady=5, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        params_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Create Treeview for clinical parameters
        params_cols = ("Parameter", "Baseline", "Discharge", "FU")
        self.fu_params_tree = ttk.Treeview(params_frame, columns=params_cols, show="headings", height=11)
        for col in params_cols:
            self.fu_params_tree.heading(col, text=col)
            self.fu_params_tree.column(col, width=150 if col == "Parameter" else 100, anchor="center")
        self.fu_params_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        params_scroll = ttk.Scrollbar(params_frame, orient=tk.VERTICAL, command=self.fu_params_tree.yview)
        params_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.fu_params_tree.config(yscrollcommand=params_scroll.set)
        
        # 4. Vital Signs Table
        vitals_frame = tk.LabelFrame(win, text=" Vital Signs ", padx=5, pady=5, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        vitals_frame.pack(fill=tk.X, padx=10, pady=5)
        
        vitals_cols = ("Parameter", "Value")
        self.fu_vitals_tree = ttk.Treeview(vitals_frame, columns=vitals_cols, show="headings", height=5)
        for col in vitals_cols:
            self.fu_vitals_tree.heading(col, text=col)
            self.fu_vitals_tree.column(col, width=200 if col == "Parameter" else 150, anchor="center")
        self.fu_vitals_tree.pack(fill=tk.X)
        
        # 5. Diuretic History Table
        meds_frame = tk.LabelFrame(win, text=" Loop Diuretic History ", padx=5, pady=5, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        meds_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        meds_cols = ("Drug", "Single Dose", "Frequency", "Daily Dose", "Start Date", "End Date", "Indication")
        self.fu_meds_tree = ttk.Treeview(meds_frame, columns=meds_cols, show="headings", height=5)
        for col in meds_cols:
            self.fu_meds_tree.heading(col, text=col)
            if col == "Drug":
                self.fu_meds_tree.column(col, width=100, anchor="center")
            elif col == "Indication":
                self.fu_meds_tree.column(col, width=200, anchor="w")
            elif col in ("Start Date", "End Date"):
                self.fu_meds_tree.column(col, width=90, anchor="center")
            else:
                self.fu_meds_tree.column(col, width=80, anchor="center")
        
        meds_scroll = ttk.Scrollbar(meds_frame, orient=tk.VERTICAL, command=self.fu_meds_tree.yview)
        meds_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self.fu_meds_tree.config(yscrollcommand=meds_scroll.set)
        self.fu_meds_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        # Hidden text for clipboard (stores tab-separated data)
        self.fu_clipboard_data = ""
        
        # Copy button
        btn_frame = tk.Frame(win, bg="#f4f4f4")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Button(btn_frame, text="Copy All to Clipboard", command=self._copy_fu_output,
                  bg="#2c3e50", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_frame, text="Export XLSX", command=self._export_fu_xlsx,
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Export CSV", command=self._export_fu_csv,
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        self.fu_include_prn = tk.BooleanVar(value=False)
        tk.Checkbutton(btn_frame, text="Include PRN", variable=self.fu_include_prn,
                       bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="📊 View Timeline", command=self._show_diuretic_timeline,
                  bg="#9b59b6", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
                  
        tk.Button(btn_frame, text="Close", command=win.destroy,
                  bg="#e74c3c", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)
    
    
    def _ask_confirm_fuzzy(self, original_name, matched_name):
        """Ask user to confirm a fuzzy match."""
        return messagebox.askyesno(
            "Confirm Medication Match", 
            f"Fuzzy match detected:\n\nOriginal: '{original_name}'\nMatched: '{matched_name}'\n\nIs this correct?",
            parent=self.root # Ensure it appears on top
        )
    
    def _generate_fu_tables(self, win):
        """Generate FU Highlights tables using DataFrames."""
        patient_id = self.fu_patient_var.get()
        visit = self.fu_visit_var.get()
        
        if not patient_id:
            messagebox.showwarning("No Patient", "Please select a patient.")
            return
        
        # Initialize exporter with fuzzy confirmation callback
        exporter = FUHighlightsExporter(self.df_main, fuzzy_confirm_callback=self._ask_confirm_fuzzy)
        # Now returns DataFrames
        df_highlights, df_vitals, df_diuretics = exporter.generate_highlights_table(patient_id, visit)
        
        if df_highlights is None:
            messagebox.showwarning("No Data", f"No data found for patient {patient_id}")
            return
        
        # Store for export
        self.fu_highlights_df = df_highlights
        self.fu_vitals_df = df_vitals
        self.fu_diuretics_df = df_diuretics
        self.fu_visit_label = visit
        self.fu_patient_label = patient_id
        
        # --- Populate Clinical Parameters Tree ---
        headers = list(df_highlights.columns)
        self.fu_params_tree["columns"] = headers
        for col in headers:
            self.fu_params_tree.heading(col, text=col)
            self.fu_params_tree.column(col, width=200 if col == "Parameter" else 100, anchor="center")
        
        self.fu_params_tree.delete(*self.fu_params_tree.get_children())
        for values in df_highlights.values:
            self.fu_params_tree.insert("", "end", values=list(values))
        
        # --- Populate Vital Signs Tree (dynamic columns like params) ---
        vitals_headers = list(df_vitals.columns)
        self.fu_vitals_tree["columns"] = vitals_headers
        for col in vitals_headers:
            self.fu_vitals_tree.heading(col, text=col)
            self.fu_vitals_tree.column(col, width=200 if col == "Parameter" else 100, anchor="center")
        
        self.fu_vitals_tree.delete(*self.fu_vitals_tree.get_children())
        for values in df_vitals.values:
            self.fu_vitals_tree.insert("", "end", values=list(values))

        # --- Populate Diuretic History Tree ---
        self.fu_meds_tree.delete(*self.fu_meds_tree.get_children())
        for values in df_diuretics.values:
            self.fu_meds_tree.insert("", "end", values=list(values))
            
    def _copy_fu_output(self):
        """Copy FU output to clipboard from stored DataFrames."""
        if not hasattr(self, 'fu_highlights_df') or self.fu_highlights_df is None:
            messagebox.showwarning("No Data", "Generate tables first before copying.")
            return
            
        try:
            # Build string output
            output = f"=== {self.fu_visit_label} FU Highlights - {self.fu_patient_label} ===\n\n"
            
            output += "--- Clinical Parameters ---\n"
            output += self.fu_highlights_df.to_csv(sep="\t", index=False)
            
            output += "\n--- Vital Signs ---\n"
            output += self.fu_vitals_df.to_csv(sep="\t", index=False)
            
            output += "\n--- Loop Diuretic History (All Events) ---\n"
            output += self.fu_diuretics_df.to_csv(sep="\t", index=False)
            
            self.root.clipboard_clear()
            self.root.clipboard_append(output)
            messagebox.showinfo("Copied", "Content copied to clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy: {str(e)}")

    def _export_fu_xlsx(self):
        """Export FU data to Excel."""
        if not hasattr(self, 'fu_highlights_df') or self.fu_highlights_df is None:
            messagebox.showwarning("No Data", "Generate tables first.")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"FU_Highlights_{self.fu_patient_label}_{self.fu_visit_label}.xlsx",
            title="Export to Excel"
        )
        if not filename:
            return
            
        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                self.fu_highlights_df.to_excel(writer, sheet_name="Clinical Params", index=False)
                self.fu_vitals_df.to_excel(writer, sheet_name="Vital Signs", index=False)
                self.fu_diuretics_df.to_excel(writer, sheet_name="Diuretic History", index=False)
            messagebox.showinfo("Success", f"Exported to {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

    def _export_fu_csv(self):
        """Export FU data to CSV (Combined)."""
        if not hasattr(self, 'fu_highlights_df') or self.fu_highlights_df is None:
            messagebox.showwarning("No Data", "Generate tables first.")
            return
            
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=f"FU_Highlights_{self.fu_patient_label}_{self.fu_visit_label}.csv",
            title="Export to CSV"
        )
        if not filename:
            return
            
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                f.write(f"=== {self.fu_visit_label} FU Highlights - {self.fu_patient_label} ===\n")
                f.write("\n--- Clinical Parameters ---\n")
                self.fu_highlights_df.to_csv(f, index=False)
                f.write("\n--- Vital Signs ---\n")
                self.fu_vitals_df.to_csv(f, index=False)
                f.write("\n--- Loop Diuretic History (All Events) ---\n")
                self.fu_diuretics_df.to_csv(f, index=False)
            messagebox.showinfo("Success", f"Exported to {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

    def _show_diuretic_timeline(self):
        """Show diuretic dosage timeline chart."""
        if not hasattr(self, 'fu_patient_var') or not self.fu_patient_var.get():
            messagebox.showwarning("No Patient", "Select a patient first.")
            return
        
        patient_id = self.fu_patient_var.get()
        
        try:
            from fu_highlights_export import FUHighlightsExporter
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            
            exporter = FUHighlightsExporter(self.df_main)
            include_prn = self.fu_include_prn.get()
            fig = exporter.generate_diuretic_timeline(patient_id, include_prn=include_prn)
            
            if fig is None:
                messagebox.showinfo("No Data", "No loop diuretic data found for this patient.")
                return
            
            # Create popup window for chart - size based on figure dimensions
            chart_win = tk.Toplevel(self.root)
            chart_win.title(f"Diuretic Timeline - Patient {patient_id}")
            # Get figure size and scale window accordingly
            fig_width, fig_height = fig.get_size_inches()
            win_width = int(fig_width * 80)  # Scale to pixels
            win_height = int(fig_height * 80) + 50  # Extra space for toolbar
            chart_win.geometry(f"{win_width}x{win_height}")
            
            # Embed matplotlib figure in tkinter
            canvas = FigureCanvasTkAgg(fig, master=chart_win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)
            
            # Add toolbar
            from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
            toolbar_frame = tk.Frame(chart_win)
            toolbar_frame.pack(fill=tk.X)
            toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
            toolbar.update()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate timeline: {str(e)}")

    def show_procedure_timing(self, master=None):
        """Show Procedure Timing matrix with adjustable row order and multi-patient selection."""
        # Use existing master if provided (e.g. for refresh), else self.root
        parent = master if master else self.root
        
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        
        win = tk.Toplevel(parent)
        win.title("Procedure Timing Matrix")
        win.geometry("1400x800")
        win.configure(bg="#f4f4f4")
        
        # Get all non-screen-failure patients
        all_patients = sorted(self.df_main['Screening #'].dropna().unique())
        non_sf_patients = [p for p in all_patients if not self._is_screen_failure(p)]
        
        # Initialize exporter
        exporter = ProcedureTimingExporter(self.df_main, self.labels)
        self._proc_timing_fields = exporter.get_field_order()
        self._proc_timing_exporter = exporter # Store for reuse
        
        # --- Layout ---
        # 1. Top Toolbar (Exports)
        top_bar = tk.Frame(win, bg="#ecf0f1", pady=10, highlightthickness=1, highlightbackground="#bdc3c7")
        top_bar.pack(fill=tk.X, side=tk.TOP)
        
        tk.Label(top_bar, text="Procedure Timing", bg="#ecf0f1", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=20)
        
        tk.Button(top_bar, text="Export XLSX", command=lambda: self._export_proc_timing('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=10)
        tk.Button(top_bar, text="Copy to Clipboard", command=self._copy_proc_timing,
                  bg="#9b59b6", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=10)
        
        # Main content area
        main_content = tk.Frame(win, bg="#f4f4f4")
        main_content.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Left Panel (Configurations)
        config_frame = tk.Frame(main_content, bg="#f4f4f4", width=350)
        config_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        config_frame.pack_propagate(False)
        
        # --- Section 1: Patients ---
        # Allow expanding to fill vertical space
        pat_group = tk.LabelFrame(config_frame, text=" 1. Select Patients ", bg="white", font=("Segoe UI", 10, "bold"), padx=5, pady=5)
        pat_group.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Buttons for Select All/None
        pat_btn_frame = tk.Frame(pat_group, bg="white")
        pat_btn_frame.pack(fill=tk.X, pady=(0, 5))
        
        self.proc_pat_vars = {} # Dictionary to store vars
        
        def toggle_all_pats(state):
            for pid, var in self.proc_pat_vars.items():
                var.set(state)
            self._generate_proc_timing_matrix()

        tk.Button(pat_btn_frame, text="All", command=lambda: toggle_all_pats(True), 
                  font=("Segoe UI", 8), width=5).pack(side=tk.LEFT, padx=2)
        tk.Button(pat_btn_frame, text="None", command=lambda: toggle_all_pats(False), 
                  font=("Segoe UI", 8), width=5).pack(side=tk.LEFT, padx=2)
        
        # Scrollable patient list
        # Frame to hold canvas and scrollbar
        list_container = tk.Frame(pat_group, bg="white")
        list_container.pack(fill=tk.BOTH, expand=True)
        
        pat_canvas = tk.Canvas(list_container, bg="white", highlightthickness=0)
        pat_scroll = ttk.Scrollbar(list_container, orient="vertical", command=pat_canvas.yview)
        pat_scrollable = tk.Frame(pat_canvas, bg="white")
        
        pat_scrollable.bind(
            "<Configure>",
            lambda e: pat_canvas.configure(scrollregion=pat_canvas.bbox("all"))
        )
        # Mousewheel scrolling
        def _on_mousewheel(event):
            try:
                pat_canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            except Exception:
                pass
            
        pat_canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        pat_canvas.create_window((0, 0), window=pat_scrollable, anchor="nw")
        pat_canvas.configure(yscrollcommand=pat_scroll.set)
        
        pat_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pat_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Populate patients
        for pid in non_sf_patients:
            var = tk.BooleanVar(value=False)
            self.proc_pat_vars[pid] = var
            cb = tk.Checkbutton(pat_scrollable, text=str(pid), variable=var, bg="white", anchor="w",
                                command=self._generate_proc_timing_matrix, font=("Segoe UI", 9))
            cb.pack(fill=tk.X)
            
        # Select first patient by default
        if non_sf_patients:
            self.proc_pat_vars[non_sf_patients[0]].set(True)

        # --- Section 2: Row Order ---
        # Fixed height for row order relative to window, or let it share space?
        # User wants patient list scrollable, so let's give patient list more weight
        # or fixed height for row order? 
        # Make row order fixed height (e.g. 300px) so patient list takes remaining space
        order_group = tk.LabelFrame(config_frame, text=" 2. Row Order ", bg="white", font=("Segoe UI", 10, "bold"), padx=5, pady=5)
        order_group.pack(fill=tk.X, ipady=5) # No expand, just fill width
        
        # Listbox for ordering
        self._proc_order_listbox = tk.Listbox(order_group, bg="white", height=15, font=("Segoe UI", 9),
                                               selectmode=tk.SINGLE, exportselection=False)
        self._proc_order_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        order_scroll = ttk.Scrollbar(order_group, orient=tk.VERTICAL, command=self._proc_order_listbox.yview)
        order_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self._proc_order_listbox.config(yscrollcommand=order_scroll.set)
        
        # Populate with field labels
        for col_name, label in self._proc_timing_fields:
            self._proc_order_listbox.insert(tk.END, label)
            
        # Up/Down buttons
        btn_frame = tk.Frame(order_group, bg="white")
        btn_frame.pack(side=tk.LEFT, padx=5, fill=tk.Y)
        
        tk.Button(btn_frame, text="↑", command=self._move_proc_field_up,
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"), width=3).pack(pady=2)
        tk.Button(btn_frame, text="↓", command=self._move_proc_field_down,
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"), width=3).pack(pady=2)
        tk.Button(btn_frame, text="R", command=lambda: self._reset_proc_order(exporter),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 9, "bold"), width=3).pack(pady=10)

        # Right Panel (Matrix Preview)
        preview_group = tk.LabelFrame(main_content, text=" Matrix Preview (Auto-Pivoted) ", bg="#f4f4f4", font=("Segoe UI", 10, "bold"), padx=5, pady=5)
        preview_group.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        tree_container = tk.Frame(preview_group)
        tree_container.pack(fill=tk.BOTH, expand=True)
        
        self._proc_timing_tree = ttk.Treeview(tree_container, show='headings')
        
        h_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self._proc_timing_tree.xview)
        v_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self._proc_timing_tree.yview)
        self._proc_timing_tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)
        
        self._proc_timing_tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        # Styles for wrapping and BOLD headers
        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        # Initial Display
        self._generate_proc_timing_matrix()

    def _move_proc_field_up(self):
        """Move selected field up in order."""
        sel = self._proc_order_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self._proc_timing_fields[idx], self._proc_timing_fields[idx-1] = \
            self._proc_timing_fields[idx-1], self._proc_timing_fields[idx]
        text = self._proc_order_listbox.get(idx)
        self._proc_order_listbox.delete(idx)
        self._proc_order_listbox.insert(idx-1, text)
        self._proc_order_listbox.selection_set(idx-1)
        self._generate_proc_timing_matrix() # Refresh matrix

    def _move_proc_field_down(self):
        """Move selected field down in order."""
        sel = self._proc_order_listbox.curselection()
        if not sel or sel[0] >= len(self._proc_timing_fields) - 1:
            return
        idx = sel[0]
        self._proc_timing_fields[idx], self._proc_timing_fields[idx+1] = \
            self._proc_timing_fields[idx+1], self._proc_timing_fields[idx]
        text = self._proc_order_listbox.get(idx)
        self._proc_order_listbox.delete(idx)
        self._proc_order_listbox.insert(idx+1, text)
        self._proc_order_listbox.selection_set(idx+1)
        self._generate_proc_timing_matrix() # Refresh matrix

    def _reset_proc_order(self, exporter):
        """Reset field order to default."""
        self._proc_timing_fields = exporter.get_field_order()
        self._proc_order_listbox.delete(0, tk.END)
        for col_name, label in self._proc_timing_fields:
            self._proc_order_listbox.insert(tk.END, label)
        self._generate_proc_timing_matrix()

    def _generate_proc_timing_matrix(self):
        """Generate/Update the procedure timing matrix (Pivoted)."""
        # Get selected patients
        selected_pats = [pid for pid, var in self.proc_pat_vars.items() if var.get()]
        
        if not selected_pats:
            # Clear tree
            self._proc_timing_tree.delete(*self._proc_timing_tree.get_children())
            self._proc_timing_tree["columns"] = []
            return

        self._proc_timing_exporter.set_field_order(self._proc_timing_fields)
        df_flat = self._proc_timing_exporter.generate_matrix(selected_pats)
        
        if df_flat is None or df_flat.empty:
            return
            
        # PIVOT: Transpose the DataFrame
        # Set 'Patient' as index, then transpose
        df_t = df_flat.set_index('Patient').T
        
        # Reset index to make "Field Names" the first column
        df_final = df_t.reset_index()
        df_final.rename(columns={'index': 'Procedure Step'}, inplace=True)
        
        # Store for export
        self._proc_timing_df = df_final
        
        # Update treeview
        # Columns: 'Procedure Step' + Patient IDs
        # Patient IDs might be integers, ensure strings for column names
        columns = list(df_final.columns)
        st_columns = [str(c) for c in columns]
        
        self._proc_timing_tree["columns"] = st_columns
        
        for col in st_columns:
            self._proc_timing_tree.heading(col, text=col)
            # Layout adjustment: First column (Step) wider, others narrower
            width = 300 if col == 'Procedure Step' else 100
            self._proc_timing_tree.column(col, width=width, anchor="w" if col == 'Procedure Step' else "center")
        
        self._proc_timing_tree.delete(*self._proc_timing_tree.get_children())
        for values in df_final[columns].fillna('').values:
            self._proc_timing_tree.insert("", "end", values=list(values))

    def _export_proc_timing(self, fmt):
        """Export procedure timing matrix."""
        if not hasattr(self, '_proc_timing_df') or self._proc_timing_df is None:
            messagebox.showwarning("No Data", "Generate matrix first.")
            return
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"Procedure_Timing_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self._proc_timing_df.to_excel(path, index=False)
            else:
                self._proc_timing_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _copy_proc_timing(self):
        """Copy procedure timing matrix to clipboard."""
        if not hasattr(self, '_proc_timing_df') or self._proc_timing_df is None:
            messagebox.showwarning("No Data", "Generate matrix first.")
            return
        
        try:
            # Tab-separated for easy paste into Excel
            output = self._proc_timing_df.to_csv(sep="\t", index=False)
            self.root.clipboard_clear()
            self.root.clipboard_append(output)
            messagebox.showinfo("Copied", "Matrix copied to clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"Copy failed: {e}")

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
             print("-" * 40)
             print(f"SDV INFO FOR SELECTED ITEM")
             print(f"Variable: {code}")
             print(f"Value:    {val}")
             print(f"Form:     {form_name} (Row #{repeat_num})")
             print(f"Status:   {details['status']}")
             print(f"User:     {details['user']}")
             print(f"Date:     {details['date']}")
             print("-" * 40)

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
    # HF Hospitalizations Module
    # -------------------------------------------------------------------------
    
    def show_hf_hospitalizations(self):
        """Show HF Hospitalizations summary window."""
        if self.hf_manager is None:
            messagebox.showwarning("Warning", "No data loaded. Please load an Excel file first.")
            return
        
        # Create summary window
        win = tk.Toplevel(self.root)
        win.title("HF Hospitalizations - Summary")
        win.geometry("900x600")
        win.transient(self.root)
        
        # Header
        header_frame = tk.Frame(win, bg="#e74c3c", pady=10)
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text="Heart Failure Hospitalizations", 
                 bg="#e74c3c", fg="white", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT, padx=20)
        
        # Info label
        info_frame = tk.Frame(win, bg="#f4f4f4", pady=5)
        info_frame.pack(fill=tk.X)
        tk.Label(info_frame, text="Pre-Treatment: 6 months before | Post-Treatment: 6 months after (based on AE symptom onset)", 
                 bg="#f4f4f4", fg="#666", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=10)
        
        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X)
        
        tk.Button(toolbar, text="Refresh", command=lambda: self._refresh_hf_summary(tree, exclude_sf_var.get()),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(toolbar, text="Export Summary", command=lambda: self._export_hf_summary(exclude_sf_var.get()),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Tuning Keywords", command=self._show_hf_tuning_dialog,
                  bg="#f39c12", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        
        # Exclude Screen Failures checkbox
        exclude_sf_var = tk.BooleanVar(value=True)  # Default to exclude
        exclude_sf_cb = tk.Checkbutton(toolbar, text="Exclude Screen Failures", 
                                        variable=exclude_sf_var, bg="#f4f4f4",
                                        command=lambda: self._refresh_hf_summary(tree, exclude_sf_var.get()))
        exclude_sf_cb.pack(side=tk.LEFT, padx=20)
        
        # Summary Treeview
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        columns = ("patient", "treatment_date", "pre_count", "pre_count_1y", "post_count", "post_count_1y")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        tree.heading("patient", text="Patient ID")
        tree.heading("treatment_date", text="Treatment Date")
        tree.heading("pre_count", text="Pre-6M")
        tree.heading("pre_count_1y", text="Pre-1Y")
        tree.heading("post_count", text="Post-6M")
        tree.heading("post_count_1y", text="Post-1Y")
        
        tree.column("patient", width=120, anchor="center")
        tree.column("treatment_date", width=120, anchor="center")
        tree.column("pre_count", width=100, anchor="center")
        tree.column("pre_count_1y", width=100, anchor="center")
        tree.column("post_count", width=100, anchor="center")
        tree.column("post_count_1y", width=100, anchor="center")
        
        # Scrollbars
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)
        
        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Tags for styling
        tree.tag_configure('has_events', background='#fff3cd')
        tree.tag_configure('no_events', background='#ffffff')
        
        # Double-click to drill down
        tree.bind("<Double-1>", lambda e: self._show_hf_detail_for_selected(tree))
        
        # Store references for refresh
        self.hf_summary_tree = tree
        self.hf_exclude_sf_var = exclude_sf_var
        
        # Populate data
        self._refresh_hf_summary(tree, exclude_sf_var.get())
        
        # Instructions
        tk.Label(win, text="Double-click on a patient to view/edit event details", 
                 fg="#888", font=("Segoe UI", 9, "italic")).pack(pady=5)
    
    def _refresh_hf_summary(self, tree, exclude_sf=True):
        """Refresh the HF summary tree with current data."""
        # Clear existing items
        for item in tree.get_children():
            tree.delete(item)
        
        # Get screen failures list if excluding
        screen_failures = set()
        if exclude_sf:
            try:
                screen_failures = set(self.get_screen_failures())
            except Exception as e:
                print(f"Error getting screen failures: {e}")
        
        # Get all patients' summaries
        try:
            summaries = self.hf_manager.get_all_patients_summary()
        except Exception as e:
            print(f"Error getting HF summaries: {e}")
            return
        
        # Sort by patient ID
        summaries.sort(key=lambda x: x['patient_id'])
        
        for summary in summaries:
            # Skip screen failures if checkbox is checked
            if exclude_sf and summary['patient_id'] in screen_failures:
                continue
                
            has_events = summary['pre_count'] > 0 or summary['post_count'] > 0
            tag = 'has_events' if has_events else 'no_events'
            
            tree.insert("", "end", iid=summary['patient_id'], values=(
                summary['patient_id'],
                summary['treatment_date'],
                summary['pre_count_6m'],
                summary['pre_count_1y'],
                summary['post_count_6m'],
                summary['post_count_1y']
            ), tags=(tag,))
    
    def _show_hf_detail_for_selected(self, tree):
        """Show detail window for selected patient."""
        selection = tree.selection()
        if not selection:
            return
        
        patient_id = selection[0]
        self._show_hf_detail_window(patient_id)
    
    def _show_hf_detail_window(self, patient_id):
        """Show detailed HF events for a patient with editable lists."""
        summary = self.hf_manager.get_patient_summary(patient_id)
        
        # Create detail window
        win = tk.Toplevel(self.root)
        win.title(f"HF Hospitalizations - Patient {patient_id}")
        win.geometry("1100x650")
        win.transient(self.root)
        
        # Header
        header_frame = tk.Frame(win, bg="#e74c3c", pady=10)
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text=f"Patient: {patient_id}", 
                 bg="#e74c3c", fg="white", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT, padx=20)
        tk.Label(header_frame, text=f"Treatment Date: {summary['treatment_date']}", 
                 bg="#e74c3c", fg="white", font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=20)
        
        # Summary stats
        stats_frame = tk.Frame(win, bg="#f4f4f4", pady=10)
        stats_frame.pack(fill=tk.X)
        tk.Label(stats_frame, text=f"Pre-Treatment (1Y): {summary['pre_count_1y']} events (6M: {summary['pre_count_6m']})", 
                 bg="#f4f4f4", fg="#c0392b", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=20)
        tk.Label(stats_frame, text=f"Post-Treatment (1Y): {summary['post_count_1y']} events (6M: {summary['post_count_6m']})", 
                 bg="#f4f4f4", fg="#27ae60", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=20)
        
        # Notebook for Pre/Post tabs
        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Pre-Treatment Tab
        pre_frame = tk.Frame(notebook)
        notebook.add(pre_frame, text=f"Pre-Treatment ({summary['pre_count_1y']})")
        pre_tree = self._create_hf_events_tree(pre_frame, summary['pre_events'], patient_id, "pre")
        
        # Post-Treatment Tab
        post_frame = tk.Frame(notebook)
        notebook.add(post_frame, text=f"Post-Treatment ({summary['post_count_1y']})")
        post_tree = self._create_hf_events_tree(post_frame, summary['post_events'], patient_id, "post")
        
        # Store references for save
        self._hf_detail_patient = patient_id
        self._hf_detail_pre_tree = pre_tree
        self._hf_detail_post_tree = post_tree
        
        # Button frame
        btn_frame = tk.Frame(win, pady=10)
        btn_frame.pack(fill=tk.X)
        
        tk.Button(btn_frame, text="Save Changes", command=lambda: self._save_hf_changes(win),
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=20)
        tk.Button(btn_frame, text="Add Manual Event", command=lambda: self._add_manual_hf_event(patient_id, notebook),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="Export Details", command=lambda: self._export_hf_details(patient_id),
                  bg="#9b59b6", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)
    
    def _create_hf_events_tree(self, parent, events, patient_id, period):
        """Create a treeview for HF events with edit controls."""
        # Toolbar
        toolbar = tk.Frame(parent, pady=5)
        toolbar.pack(fill=tk.X)
        
        tk.Button(toolbar, text="Toggle Included/Excluded", command=lambda: toggle_include(tree),
                  bg="#e67e22", fg="white", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=10)
        
        columns = ("include", "date", "source", "term", "matched", "confidence", "type")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)
        
        tree.heading("include", text="Include")
        tree.heading("date", text="Date")
        tree.heading("source", text="Source")
        tree.heading("term", text="Original Term")
        tree.heading("matched", text="Matched Synonym")
        tree.heading("confidence", text="Conf.")
        tree.heading("type", text="Match Type")
        
        tree.column("include", width=60, anchor="center")
        tree.column("date", width=100, anchor="center")
        tree.column("source", width=60, anchor="center")
        tree.column("term", width=300, anchor="w")
        tree.column("matched", width=200, anchor="w")
        tree.column("confidence", width=60, anchor="center")
        tree.column("type", width=80, anchor="center")
        
        # Tags for styling
        tree.tag_configure('included', background='#d4edda')
        tree.tag_configure('excluded', background='#f8d7da', foreground='#888')
        tree.tag_configure('manual', background='#cce5ff')
        
        # Populate events
        for event in events:
            include_text = "✓" if event.is_included else "✗"
            tag = 'manual' if event.is_manual else ('included' if event.is_included else 'excluded')
            
            tree.insert("", "end", iid=event.event_id, values=(
                include_text,
                event.date,
                event.source_form,
                event.original_term[:50] + "..." if len(event.original_term) > 50 else event.original_term,
                event.matched_synonym,
                f"{event.confidence:.0%}",
                event.match_type
            ), tags=(tag,))
        
        # Scrollbars
        tree_frame = tk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=v_scroll.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Context menu for toggle
        def toggle_include(event_widget=tree):
            selection = event_widget.selection()
            for item_id in selection:
                current = event_widget.item(item_id, "values")
                new_include = "✗" if current[0] == "✓" else "✓"
                new_tag = 'included' if new_include == "✓" else 'excluded'
                event_widget.item(item_id, values=(new_include,) + current[1:], tags=(new_tag,))
        
        tree.bind("<Double-1>", lambda e: toggle_include())
        
        # Store events reference on tree
        tree.events = events
        tree.period = period
        
        tk.Label(parent, text="Double-click to toggle Include/Exclude", 
                 fg="#888", font=("Segoe UI", 8, "italic")).pack()
        
        return tree
    
    def _save_hf_changes(self, win):
        """Save changes made in the detail window."""
        from hf_hospitalization_manager import HFEvent
        
        patient_id = self._hf_detail_patient
        
        # Process both trees
        for tree in [self._hf_detail_pre_tree, self._hf_detail_post_tree]:
            for item_id in tree.get_children():
                values = tree.item(item_id, "values")
                is_included = values[0] == "✓"
                
                # Find original event
                original_event = None
                for e in tree.events:
                    if e.event_id == item_id:
                        original_event = e
                        break
                
                if original_event and original_event.is_included != is_included:
                    # Create updated event
                    updated_event = HFEvent(
                        event_id=original_event.event_id,
                        date=original_event.date,
                        source_form=original_event.source_form,
                        source_row=original_event.source_row,
                        original_term=original_event.original_term,
                        matched_synonym=original_event.matched_synonym,
                        match_type=original_event.match_type,
                        confidence=original_event.confidence,
                        is_included=is_included,
                        is_manual=original_event.is_manual,
                        notes=original_event.notes
                    )
                    self.hf_manager.update_event(patient_id, updated_event)
        
        messagebox.showinfo("Saved", "Changes saved successfully.")
        
        # Refresh summary if open
        if hasattr(self, 'hf_summary_tree'):
            self._refresh_hf_summary(self.hf_summary_tree)
    
    def _add_manual_hf_event(self, patient_id, notebook):
        """Add a manual HF event."""
        from hf_hospitalization_manager import HFEvent
        
        # Determine which tab is active
        tab_index = notebook.index(notebook.select())
        is_pre = tab_index == 0
        
        # Simple dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Add Manual HF Event")
        dialog.geometry("400x250")
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text="Date (YYYY-MM-DD):", font=("Segoe UI", 10)).grid(row=0, column=0, padx=10, pady=10, sticky="e")
        date_entry = tk.Entry(dialog, width=20)
        date_entry.grid(row=0, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Description:", font=("Segoe UI", 10)).grid(row=1, column=0, padx=10, pady=10, sticky="e")
        desc_entry = tk.Entry(dialog, width=30)
        desc_entry.grid(row=1, column=1, padx=10, pady=10)
        
        tk.Label(dialog, text="Period:", font=("Segoe UI", 10)).grid(row=2, column=0, padx=10, pady=10, sticky="e")
        period_var = tk.StringVar(value="pre" if is_pre else "post")
        tk.Radiobutton(dialog, text="Pre-Treatment", variable=period_var, value="pre").grid(row=2, column=1, sticky="w")
        tk.Radiobutton(dialog, text="Post-Treatment", variable=period_var, value="post").grid(row=3, column=1, sticky="w")
        
        def save_manual():
            date_str = date_entry.get().strip()
            desc = desc_entry.get().strip()
            period = period_var.get()
            
            if not date_str or not desc:
                messagebox.showwarning("Warning", "Please fill in all fields.")
                return
            
            # Create manual event
            event_id = f"MANUAL_{patient_id}_{len(self.hf_manager.manual_edits.get(patient_id, []))}"
            source = f"MANUAL_{'PRE' if period == 'pre' else 'POST'}"
            
            event = HFEvent(
                event_id=event_id,
                date=date_str,
                source_form=source,
                source_row=0,
                original_term=desc,
                matched_synonym="Manual Entry",
                match_type="manual",
                confidence=1.0,
                is_included=True,
                is_manual=True
            )
            
            self.hf_manager.update_event(patient_id, event)
            dialog.destroy()
            messagebox.showinfo("Added", "Manual event added. Please refresh the detail view.")
        
        tk.Button(dialog, text="Add Event", command=save_manual,
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).grid(row=4, column=0, columnspan=2, pady=20)
    
    def _export_hf_summary(self, exclude_sf=True):
        """Export HF summary to Excel."""
        if not hasattr(self, 'hf_manager') or self.hf_manager is None:
            return
        
        try:
            summaries = self.hf_manager.get_all_patients_summary()
            
            # Filter screen failures if enabled
            if exclude_sf:
                screen_failures = set(self.get_screen_failures())
                summaries = [s for s in summaries if s['patient_id'] not in screen_failures]
            
            df = pd.DataFrame([{
                'Patient ID': s['patient_id'],
                'Treatment Date': s['treatment_date'],
                'Pre-6M HF Hosps': s.get('pre_count_6m', s['pre_count']),
                'Pre-1Y HF Hosps': s.get('pre_count_1y', 0),
                'Post-6M HF Hosps': s.get('post_count_6m', s['post_count']),
                'Post-1Y HF Hosps': s.get('post_count_1y', 0)
            } for s in summaries])
            
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                initialfile="hf_hospitalizations_summary.xlsx"
            )
            
            if path:
                if path.endswith('.csv'):
                    df.to_csv(path, index=False)
                else:
                    df.to_excel(path, index=False)
                messagebox.showinfo("Exported", f"Summary exported to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
    
    def _export_hf_details(self, patient_id):
        """Export detailed HF events for a patient."""
        try:
            summary = self.hf_manager.get_patient_summary(patient_id)
            
            # Combine pre and post events
            all_events = []
            for event in summary['pre_events']:
                all_events.append({
                    'Patient ID': patient_id,
                    'Period': 'Pre-Treatment',
                    'Date': event.date,
                    'Source': event.source_form,
                    'Term': event.original_term,
                    'Matched': event.matched_synonym,
                    'Confidence': f"{event.confidence:.0%}",
                    'Type': event.match_type,
                    'Included': 'Yes' if event.is_included else 'No',
                    'Manual': 'Yes' if event.is_manual else 'No'
                })
            for event in summary['post_events']:
                all_events.append({
                    'Patient ID': patient_id,
                    'Period': 'Post-Treatment',
                    'Date': event.date,
                    'Source': event.source_form,
                    'Term': event.original_term,
                    'Matched': event.matched_synonym,
                    'Confidence': f"{event.confidence:.0%}",
                    'Type': event.match_type,
                    'Included': 'Yes' if event.is_included else 'No',
                    'Manual': 'Yes' if event.is_manual else 'No'
                })
            
            df = pd.DataFrame(all_events)
            
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                initialfile=f"hf_details_{patient_id}.xlsx"
            )
            
            if path:
                if path.endswith('.csv'):
                    df.to_csv(path, index=False)
                else:
                    df.to_excel(path, index=False)
                messagebox.showinfo("Exported", f"Details exported to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    # =========================================================================
    # Assessment Data Table Feature
    # =========================================================================
    
    def show_assessment_data_table(self):
        """Display Assessment Data Table window with dropdowns for assessment, parameter, and visits."""
        if self.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return
        
        # Create popup window
        win = tk.Toplevel(self.root)
        win.title("Assessment Data Table")
        win.geometry("1200x800")
        win.configure(bg="#f4f4f4")
        
        # Store window reference
        self._assess_win = win
        self._assess_extractor = AssessmentDataExtractor(self.df_main, self.labels)
        
        # Header
        header = tk.Frame(win, bg="#3498db", padx=10, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="Assessment Data Table", font=("Segoe UI", 14, "bold"), 
                 bg="#3498db", fg="white").pack(side=tk.LEFT)
        
        # 1. Assessment Selection Section
        assess_frame = tk.LabelFrame(win, text=" Assessment & Parameter Selection ", 
                                     padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        assess_frame.pack(fill=tk.X, padx=10, pady=(10, 5))
        
        row1 = tk.Frame(assess_frame, bg="#f4f4f4")
        row1.pack(fill=tk.X, pady=5)
        
        # Assessment Type Dropdown
        tk.Label(row1, text="Assessment:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 5))
        self._assess_type_var = tk.StringVar()
        assess_types = list(ASSESSMENT_CATEGORIES.keys())
        self._assess_type_cb = ttk.Combobox(row1, textvariable=self._assess_type_var, 
                                           values=assess_types, state="readonly", width=20)
        self._assess_type_cb.pack(side=tk.LEFT, padx=(0, 15))
        self._assess_type_cb.bind("<<ComboboxSelected>>", self._on_assessment_type_changed)
        if assess_types:
            self._assess_type_cb.current(0)
        
        # Parameter Dropdown (dynamic)
        tk.Label(row1, text="Parameter:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 5))
        self._assess_param_var = tk.StringVar()
        self._assess_param_cb = ttk.Combobox(row1, textvariable=self._assess_param_var, 
                                            state="readonly", width=25)
        self._assess_param_cb.pack(side=tk.LEFT)
        
        # 2. Visit Selection Section
        visit_frame = tk.LabelFrame(win, text=" Visit Types ", 
                                   padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        visit_frame.pack(fill=tk.X, padx=10, pady=5)
        
        visit_inner = tk.Frame(visit_frame, bg="#f4f4f4")
        visit_inner.pack(fill=tk.X)
        
        self._assess_visit_vars = {}
        self._assess_visit_chks = {}
        
        # Visit checkboxes (populated dynamically based on assessment type)
        self._assess_visit_frame = visit_inner
        
        # Select All / None for visits
        visit_btn_row = tk.Frame(visit_frame, bg="#f4f4f4")
        visit_btn_row.pack(fill=tk.X, pady=(5, 0))
        tk.Button(visit_btn_row, text="Select All Visits", command=lambda: self._select_all_visits(True),
                  bg="#27ae60", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        tk.Button(visit_btn_row, text="Clear Visits", command=lambda: self._select_all_visits(False),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        
        # 3. Patient Selection Section
        pat_frame = tk.LabelFrame(win, text=" Patient Selection ", 
                                 padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        # Filter row
        filter_row = tk.Frame(pat_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 5))
        
        self._assess_exclude_sf_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self._assess_exclude_sf_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=self._update_assess_patient_list).pack(side=tk.LEFT)
        
        # Patient list buttons
        btn_row = tk.Frame(filter_row, bg="#f4f4f4")
        btn_row.pack(side=tk.RIGHT)
        tk.Button(btn_row, text="Select All", command=lambda: self._select_all_assess_patients(True),
                  bg="#27ae60", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Select None", command=lambda: self._select_all_assess_patients(False),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        
        # Reorder buttons
        tk.Button(btn_row, text="▲ Up", command=self._move_patient_up,
                  bg="#95a5a6", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(10, 2))
        tk.Button(btn_row, text="▼ Down", command=self._move_patient_down,
                  bg="#95a5a6", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="💾 Save Order", command=self._save_assess_patient_order,
                  bg="#9b59b6", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(10, 2))
        
        # Patient listbox with scrollbar
        list_container = tk.Frame(pat_frame, bg="#f4f4f4")
        list_container.pack(fill=tk.BOTH, expand=True)
        
        self._assess_pat_listbox = tk.Listbox(list_container, selectmode=tk.EXTENDED, 
                                              font=("Segoe UI", 9), bg="white", height=8)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self._assess_pat_listbox.yview)
        self._assess_pat_listbox.configure(yscrollcommand=scrollbar.set)
        
        self._assess_pat_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Store patient order (for reordering)
        self._assess_patient_order = []
        self._assess_all_patients = sorted(self.df_main['Screening #'].dropna().unique())
        self._assess_saved_order = self._load_assess_patient_order()  # Load saved order
        
        # 4. Generate and Export Buttons
        btn_frame = tk.Frame(win, bg="#f4f4f4")
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Button(btn_frame, text="Generate Table", command=self._generate_assessment_table,
                  bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"), padx=20, pady=8).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Export XLSX", command=lambda: self._export_assessment_table('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold"), padx=15, pady=8).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Export CSV", command=lambda: self._export_assessment_table('csv'),
                  bg="#e67e22", fg="white", font=("Segoe UI", 10, "bold"), padx=15, pady=8).pack(side=tk.LEFT, padx=5)
        
        # 5. Results Table
        result_frame = tk.LabelFrame(win, text=" Results ", 
                                    padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))
        
        # Treeview for results
        self._assess_tree = ttk.Treeview(result_frame, show="headings")
        self._assess_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        scrollbar_v = ttk.Scrollbar(result_frame, orient="vertical", command=self._assess_tree.yview)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        self._assess_tree.configure(yscrollcommand=scrollbar_v.set)
        
        scrollbar_h = ttk.Scrollbar(result_frame, orient="horizontal", command=self._assess_tree.xview)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        self._assess_tree.configure(xscrollcommand=scrollbar_h.set)
        
        # Store result DataFrame
        self._assess_result_df = None
        
        # Initialize dropdowns
        self._on_assessment_type_changed(None)
        self._update_assess_patient_list()
        
        win.focus_set()
    
    def _on_assessment_type_changed(self, event):
        """Handle assessment type change - update parameters and visits."""
        assess_type = self._assess_type_var.get()
        if not assess_type or assess_type not in ASSESSMENT_CATEGORIES:
            return
        
        category = ASSESSMENT_CATEGORIES[assess_type]
        
        # Update parameter dropdown
        params = category.get("params", [])
        param_display = [p[0] for p in params]  # First element is display name
        self._assess_param_cb['values'] = param_display
        if param_display:
            self._assess_param_cb.current(0)
        
        # Update visit checkboxes
        for widget in self._assess_visit_frame.winfo_children():
            widget.destroy()
        
        self._assess_visit_vars = {}
        visits = category.get("visits", {})
        
        col = 0
        for visit_name in visits.keys():
            var = tk.BooleanVar(value=True)
            self._assess_visit_vars[visit_name] = var
            chk = tk.Checkbutton(self._assess_visit_frame, text=visit_name, variable=var,
                                bg="#f4f4f4", font=("Segoe UI", 9))
            chk.grid(row=0, column=col, padx=5, sticky="w")
            col += 1
    
    def _select_all_visits(self, select):
        """Select or deselect all visit checkboxes."""
        for var in self._assess_visit_vars.values():
            var.set(select)
    
    def _update_assess_patient_list(self):
        """Update patient listbox based on filter settings, respecting saved order."""
        self._assess_pat_listbox.delete(0, tk.END)
        
        exclude_sf = self._assess_exclude_sf_var.get()
        
        # Get available patients (filtered)
        available = []
        for p in self._assess_all_patients:
            if not p:
                continue
            if exclude_sf and self._is_screen_failure(p):
                continue
            available.append(p)
        
        # Apply saved order if exists
        if self._assess_saved_order:
            # First add patients in saved order (if they exist in available)
            ordered = []
            for p in self._assess_saved_order:
                if p in available:
                    ordered.append(p)
                    available.remove(p)
            # Then add any remaining patients not in saved order
            ordered.extend(available)
            self._assess_patient_order = ordered
        else:
            self._assess_patient_order = available
        
        for p in self._assess_patient_order:
            self._assess_pat_listbox.insert(tk.END, p)
    
    def _save_assess_patient_order(self):
        """Save patient order to JSON file for persistence."""
        import json
        order_file = os.path.join(os.path.dirname(__file__), 'assess_patient_order.json')
        try:
            with open(order_file, 'w') as f:
                json.dump(self._assess_patient_order, f, indent=2)
            self._assess_saved_order = self._assess_patient_order[:]  # Update cached order
            messagebox.showinfo("Saved", f"Patient order saved successfully.\n({len(self._assess_patient_order)} patients)")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save patient order: {e}")
    
    def _load_assess_patient_order(self):
        """Load patient order from JSON file."""
        import json
        order_file = os.path.join(os.path.dirname(__file__), 'assess_patient_order.json')
        try:
            if os.path.exists(order_file):
                with open(order_file, 'r') as f:
                    return json.load(f)
        except Exception:
            pass
        return []
    
    def _select_all_assess_patients(self, select):
        """Select or deselect all patients in the listbox."""
        if select:
            self._assess_pat_listbox.select_set(0, tk.END)
        else:
            self._assess_pat_listbox.select_clear(0, tk.END)
    
    def _move_patient_up(self):
        """Move selected patient(s) up in the list."""
        selected = list(self._assess_pat_listbox.curselection())
        if not selected or selected[0] == 0:
            return
        
        for idx in selected:
            if idx > 0:
                # Swap in order list
                self._assess_patient_order[idx-1], self._assess_patient_order[idx] = \
                    self._assess_patient_order[idx], self._assess_patient_order[idx-1]
        
        # Refresh listbox
        self._assess_pat_listbox.delete(0, tk.END)
        for p in self._assess_patient_order:
            self._assess_pat_listbox.insert(tk.END, p)
        
        # Re-select moved items
        for idx in selected:
            self._assess_pat_listbox.select_set(idx - 1)
    
    def _move_patient_down(self):
        """Move selected patient(s) down in the list."""
        selected = list(self._assess_pat_listbox.curselection())
        if not selected or selected[-1] >= len(self._assess_patient_order) - 1:
            return
        
        for idx in reversed(selected):
            if idx < len(self._assess_patient_order) - 1:
                # Swap in order list
                self._assess_patient_order[idx], self._assess_patient_order[idx+1] = \
                    self._assess_patient_order[idx+1], self._assess_patient_order[idx]
        
        # Refresh listbox
        self._assess_pat_listbox.delete(0, tk.END)
        for p in self._assess_patient_order:
            self._assess_pat_listbox.insert(tk.END, p)
        
        # Re-select moved items
        for idx in selected:
            self._assess_pat_listbox.select_set(idx + 1)
    
    def _generate_assessment_table(self):
        """Generate the assessment data table based on selections."""
        assess_type = self._assess_type_var.get()
        param_display = self._assess_param_var.get()
        
        if not assess_type or not param_display:
            messagebox.showwarning("Selection Required", "Please select an assessment type and parameter.")
            return
        
        # Get selected visits
        selected_visits = [v for v, var in self._assess_visit_vars.items() if var.get()]
        if not selected_visits:
            messagebox.showwarning("No Visits", "Please select at least one visit type.")
            return
        
        # Get selected patients (use listbox selection or all if none selected)
        selected_indices = self._assess_pat_listbox.curselection()
        if selected_indices:
            selected_patients = [self._assess_patient_order[i] for i in selected_indices]
        else:
            selected_patients = self._assess_patient_order[:]
        
        if not selected_patients:
            messagebox.showwarning("No Patients", "Please select at least one patient.")
            return
        
        # Get parameter code from display name
        category = ASSESSMENT_CATEGORIES[assess_type]
        params = category.get("params", [])
        param_code = None
        for p in params:
            if p[0] == param_display:
                param_code = p[1]
                break
        
        if not param_code:
            messagebox.showwarning("Error", "Could not find parameter code.")
            return
        
        # Generate table
        try:
            df = self._assess_extractor.generate_table(
                selected_patients, assess_type, param_code, selected_visits
            )
            self._assess_result_df = df
            
            # Update treeview
            self._assess_tree.delete(*self._assess_tree.get_children())
            
            # Configure columns
            columns = list(df.columns)
            self._assess_tree['columns'] = columns
            for col in columns:
                self._assess_tree.heading(col, text=col)
                self._assess_tree.column(col, width=100, anchor="center")
            # First column (Patient) wider
            self._assess_tree.column(columns[0], width=120, anchor="w")
            
            # Add data rows
            for values in df[columns].values:
                self._assess_tree.insert("", tk.END, values=list(values))
            
            messagebox.showinfo("Success", f"Generated table with {len(df)} patients × {len(selected_visits)} visits")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate table: {e}")
    
    def _export_assessment_table(self, fmt):
        """Export assessment table to file."""
        if self._assess_result_df is None or self._assess_result_df.empty:
            messagebox.showwarning("No Data", "Please generate a table first.")
            return
        
        assess_type = self._assess_type_var.get().replace(" ", "_").replace("-", "_")
        param = self._assess_param_var.get().replace(" ", "_").replace("/", "_")
        
        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
            initialfile=f"Assessment_{assess_type}_{param}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        
        if not path:
            return
        
        try:
            if fmt == 'xlsx':
                self._assess_result_df.to_excel(path, index=False)
            else:
                self._assess_result_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

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

    def _show_hf_tuning_dialog(self):
        """Management dialog for HF tuning keywords (Include/Exclude)."""
        if self.hf_manager is None: return
        
        dialog = tk.Toplevel(self.root)
        dialog.title("HF Tuning Keywords")
        dialog.geometry("600x550")
        dialog.transient(self.root)
        
        main_frame = tk.Frame(dialog, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        tk.Label(main_frame, text="Global Inclusion/Exclusion Keywords", font=("Segoe UI", 12, "bold")).pack(pady=5)
        tk.Label(main_frame, text="These keywords affect autodetected events globally across all patients.", 
                 fg="#666", font=("Segoe UI", 9, "italic")).pack()
        
        # Two columns: Include and Exclude
        lists_frame = tk.Frame(main_frame)
        lists_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Include List
        inc_frame = tk.LabelFrame(lists_frame, text="Include Keywords (Hard Match)", padx=5, pady=5)
        inc_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        inc_list = tk.Listbox(inc_frame, height=10)
        inc_list.pack(fill=tk.BOTH, expand=True)
        for kw in self.hf_manager.custom_includes:
            inc_list.insert(tk.END, kw)
            
        inc_ctrl = tk.Frame(inc_frame)
        inc_ctrl.pack(fill=tk.X, pady=5)
        inc_entry = tk.Entry(inc_ctrl)
        inc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def add_include():
            kw = inc_entry.get().strip().lower()
            if kw and kw not in self.hf_manager.custom_includes:
                self.hf_manager.custom_includes.append(kw)
                inc_list.insert(tk.END, kw)
                inc_entry.delete(0, tk.END)
                
        def del_include():
            sel = inc_list.curselection()
            if sel:
                kw = inc_list.get(sel[0])
                self.hf_manager.custom_includes.remove(kw)
                inc_list.delete(sel[0])
        
        tk.Button(inc_ctrl, text="+", command=add_include, width=3).pack(side=tk.LEFT, padx=2)
        tk.Button(inc_ctrl, text="-", command=del_include, width=3).pack(side=tk.LEFT, padx=2)
        
        # Exclude List
        excl_frame = tk.LabelFrame(lists_frame, text="Exclude Keywords (Ignore Match)", padx=5, pady=5)
        excl_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)
        
        excl_list = tk.Listbox(excl_frame, height=10)
        excl_list.pack(fill=tk.BOTH, expand=True)
        for kw in self.hf_manager.custom_excludes:
            excl_list.insert(tk.END, kw)
            
        excl_ctrl = tk.Frame(excl_frame)
        excl_ctrl.pack(fill=tk.X, pady=5)
        excl_entry = tk.Entry(excl_ctrl)
        excl_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        def add_exclude():
            kw = excl_entry.get().strip().lower()
            if kw and kw not in self.hf_manager.custom_excludes:
                self.hf_manager.custom_excludes.append(kw)
                excl_list.insert(tk.END, kw)
                excl_entry.delete(0, tk.END)
                
        def del_exclude():
            sel = excl_list.curselection()
            if sel:
                kw = excl_list.get(sel[0])
                self.hf_manager.custom_excludes.remove(kw)
                excl_list.delete(sel[0])
                
        tk.Button(excl_ctrl, text="+", command=add_exclude, width=3).pack(side=tk.LEFT, padx=2)
        tk.Button(excl_ctrl, text="-", command=del_exclude, width=3).pack(side=tk.LEFT, padx=2)
        
        # Save Button
        def save_and_close():
            self.hf_manager.save_tuning_config()
            dialog.destroy()
            messagebox.showinfo("Saved", "Tuning keywords saved. Please refresh summary to apply.")
            
        tk.Button(main_frame, text="Save Global Keywords", command=save_and_close,
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).pack(pady=10)


if __name__ == "__main__":
    root = tk.Tk()
    app = ClinicalDataMasterV30(root)
    root.mainloop()