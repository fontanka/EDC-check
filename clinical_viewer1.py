
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
from data_matrix_builder import show_data_matrix as _show_data_matrix
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
        """Build and display a data matrix (delegated to data_matrix_builder)."""
        _show_data_matrix(self)

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