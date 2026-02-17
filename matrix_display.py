"""
Matrix Display Module
=====================
Extracted from clinical_viewer1.py — handles all specialized matrix/table
display windows for clinical data forms (AE, CM, MH, HFH, HMEH, CVC, CVH, ACT).
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import re
import logging
from datetime import datetime

logger = logging.getLogger(__name__)


class MatrixDisplay:
    """Manages specialized matrix/table display windows.

    Args:
        app: ClinicalDataMasterV30 instance — provides root, df_main, labels, etc.
    """

    def __init__(self, app):
        self.app = app

    # ------------------------------------------------------------------
    # Generic export helper — replaces 6 near-identical export methods
    # ------------------------------------------------------------------

    def _export_matrix(self, fmt, df, patient, prefix):
        """Export a matrix DataFrame to XLSX or CSV."""
        if df is None or df.empty:
            messagebox.showerror("Error", f"No {prefix} data to export.")
            return

        ext = 'xlsx' if fmt == 'xlsx' else 'csv'
        path = filedialog.asksaveasfilename(
            defaultextension=f".{ext}",
            filetypes=[(f"{ext.upper()} Files", f"*.{ext}")],
            initialfile=f"{prefix}_{patient}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.{ext}"
        )
        if not path:
            return

        try:
            if fmt == 'xlsx':
                df.to_excel(path, index=False)
            else:
                df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"{prefix} data exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    # ------------------------------------------------------------------
    # Generic simple matrix — shared by MH, HFH, HMEH, CM-from-data
    # ------------------------------------------------------------------

    def _show_simple_matrix(self, data, pat, *, title, prefix, geometry,
                            column_order, col_widths, exclude_keys=('Ongoing',)):
        """Display a list-of-dicts as a Treeview table with export toolbar.

        Args:
            data: list of dicts (one per row)
            pat: patient identifier
            title: window title label
            prefix: filename prefix for exports
            geometry: window geometry string
            column_order: preferred column order (extras auto-appended)
            col_widths: {column_name: pixel_width} mapping
            exclude_keys: keys to hide from the display (default: 'Ongoing')
        Returns:
            (win, tree, df) tuple for callers that need further customization
        """
        if not data:
            messagebox.showinfo("Info", f"No valid {title} data found.")
            return None, None, None

        win = tk.Toplevel(self.app.root)
        win.title(f"{title} - Patient {pat}")
        win.geometry(geometry)
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

        df = pd.DataFrame(data)

        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)

        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX",
                  command=lambda: self._export_matrix('xlsx', df, pat, prefix),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV",
                  command=lambda: self._export_matrix('csv', df, pat, prefix),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        # Determine display columns: preferred order first, then any extras
        display_columns = [c for c in column_order if any(c in r for r in data)]
        for r in data:
            for k in r.keys():
                if k not in display_columns and k not in exclude_keys:
                    display_columns.append(k)

        # Filter out entirely empty columns
        display_columns = [
            c for c in display_columns
            if any(str(record.get(c, '')).strip() for record in data)
        ]

        # Tree view
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')

        for col in display_columns:
            tree.heading(col, text=col)
            width = col_widths.get(col, 120)
            tree.column(col, width=width, anchor="w", minwidth=50)

        for record in data:
            values = [record.get(col, '') for col in display_columns]
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

        return win, tree, df

    # ------------------------------------------------------------------
    # AE (Adverse Events)
    # ------------------------------------------------------------------

    def show_ae_matrix(self, pat_aes, pat):
        """Display AE data as a structured table with proper columns."""
        # Define column mappings
        col_mapping = {
            'AE #': ['Template number', 'AE #', 'AE Number', 'AESEQ'],
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
        }

        # Find available columns
        available_cols = {}
        for display_name, possible_names in col_mapping.items():
            for pn in possible_names:
                if pn in pat_aes.columns:
                    available_cols[display_name] = pn
                    break

        # Build the data for display
        ae_data = []
        for _, ae_row in pat_aes.iterrows():
            row_data = {}
            ongoing_value = False

            for display_name, source_col in available_cols.items():
                val = ae_row.get(source_col, '')
                if pd.isna(val) or str(val).lower() == 'nan':
                    val = ''
                else:
                    val = str(val).strip()

                    # Clean up date values
                    if 'Date' in display_name:
                        if 'T' in val:
                            val = val.split('T')[0]
                        elif ' ' in val and any(c.isdigit() for c in val.split(' ')[-1]):
                            parts = val.split(' ')
                            if len(parts) > 1 and ':' in parts[-1]:
                                val = ' '.join(parts[:-1])
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

            if row_data.get('AE Term', '').strip():
                ae_data.append(row_data)

        if not ae_data:
            messagebox.showinfo("Info", "No valid adverse event terms found.")
            return

        # Create window
        win = tk.Toplevel(self.app.root)
        win.title(f"Adverse Events - Patient {pat}")
        win.geometry("1400x600")
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

        # Store for export
        self._ae_matrix_df = pd.DataFrame(ae_data)
        self._ae_matrix_patient = pat

        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)

        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX",
                  command=lambda: self._export_matrix('xlsx', self._ae_matrix_df, self._ae_matrix_patient, "AE_Matrix"),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV",
                  command=lambda: self._export_matrix('csv', self._ae_matrix_df, self._ae_matrix_patient, "AE_Matrix"),
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

        # Define display columns — exclude 'Ongoing' since it's integrated into Resolution Date
        display_columns = [col for col in available_cols.keys() if col != 'Ongoing']

        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')

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
            tree.column(col, width=width,
                        anchor="w" if col in ['AE Term', 'AE Description', 'SAE Description'] else "center",
                        minwidth=50)

        def refresh_tree():
            """Rebuild tree with filtered data based on interval exclusions."""
            for item in tree.get_children():
                tree.delete(item)

            filtered_data = []
            for ae_record in ae_data:
                interval_val = ae_record.get('Interval', '').lower()
                if exclude_screening_var.get() and 'screening' in interval_val:
                    continue
                if exclude_prior_var.get() and 'prior to implant' in interval_val:
                    continue
                filtered_data.append(ae_record)

            for ae_record in filtered_data:
                values = [ae_record.get(col, '') for col in display_columns]
                tree.insert("", "end", values=values)

            count_label.config(text=f"  |  {len(filtered_data)} adverse event(s) shown")

        tk.Checkbutton(toolbar, text="During Screening", variable=exclude_screening_var,
                       command=refresh_tree, bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=2)
        tk.Checkbutton(toolbar, text="Prior to Implant", variable=exclude_prior_var,
                       command=refresh_tree, bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=2)

        count_label = tk.Label(toolbar, text=f"  |  {len(ae_data)} adverse event(s) shown",
                               bg="#f4f4f4", fg="#666", font=("Segoe UI", 9))
        count_label.pack(side=tk.LEFT, padx=5)

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

    # ------------------------------------------------------------------
    # CM (Concomitant Medications) — from CM sheet
    # ------------------------------------------------------------------

    def show_cm_matrix(self, pat_cms, pat):
        """Display CM data from dedicated CM sheet as a structured table."""
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

        exclude_cols = ['Row number', 'Form name', 'Form SN']

        # Identify special columns
        ongoing_col = next((c for c in pat_cms.columns if 'CMONGO' in c.upper() or 'ONGOING' in c.upper()), None)
        end_date_col = next((c for c in pat_cms.columns if 'CMENDTC' in c.upper() or 'CMENDAT' in c.upper() or 'END DATE' in c.upper()), None)

        display_columns = [col for col in pat_cms.columns
                           if col not in exclude_cols
                           and not col.startswith('_')
                           and (col != ongoing_col if ongoing_col else True)]

        logger.debug("Using %d columns from CM sheet", len(display_columns))

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

        # Build rows with unique keys
        final_cm_data = []
        for _, cm_row in pat_cms.iterrows():
            row_data = {}
            is_ongoing = False

            if ongoing_col:
                ongoing_val = str(cm_row.get(ongoing_col, '')).lower()
                if ongoing_val in ['yes', 'y', '1', 'true', 'checked']:
                    is_ongoing = True

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
                    multiplier, freq_note, override_dose = self.parse_frequency_multiplier(freq_val, freq_oth_val)

                    if override_dose is not None:
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

                        if unit_val and not pd.isna(unit_val) and str(unit_val).lower() not in ['nan', 'none', '']:
                            unit_str = str(unit_val).strip()
                            if 'milligram' in unit_str.lower():
                                unit_str = 'mg'
                            daily_dose_str += f" {unit_str}/day"
                        else:
                            daily_dose_str += "/day"
                    elif freq_note:
                        daily_dose_str = f"{int(single_dose) if single_dose == int(single_dose) else single_dose} {freq_note}"
            except (ValueError, TypeError):
                daily_dose_str = ""

            row_data["Daily Dose"] = daily_dose_str

            if any(v.strip() for v in row_data.values()):
                final_cm_data.append(row_data)

        if not final_cm_data:
            messagebox.showinfo("Info", "No valid medications found.")
            return

        # Create window
        win = tk.Toplevel(self.app.root)
        win.title(f"Concomitant Medications - Patient {pat}")
        win.geometry("1400x600")
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

        # Store for export
        self._cm_matrix_df = pd.DataFrame(final_cm_data)
        self._cm_matrix_patient = pat

        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)

        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX",
                  command=lambda: self._export_matrix('xlsx', self._cm_matrix_df, self._cm_matrix_patient, "CM_Matrix"),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV",
                  command=lambda: self._export_matrix('csv', self._cm_matrix_df, self._cm_matrix_patient, "CM_Matrix"),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        tk.Label(toolbar, text=f"  |  {len(final_cm_data)} medication(s) found", bg="#f4f4f4", fg="#666",
                 font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=10)

        # Tree container
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True)

        # Filter out completely empty columns
        non_empty_columns = []
        for col in final_columns:
            has_data = any(str(record.get(col, '')).strip() for record in final_cm_data)
            if has_data:
                non_empty_columns.append(col)

        tree = ttk.Treeview(tree_frame, columns=non_empty_columns, show='headings')

        compact_widths = {
            'Subject': 60, 'Rand #': 50, 'Initials': 50, 'Site': 40, 'Status': 55,
            'Medication': 120, 'Indication': 100, 'MH Reference': 150, 'AE Reference': 100,
            'Indication (Other)': 120, 'Start Date': 75, 'End Date': 70,
            'Dose': 45, 'Unit': 60, 'Frequency': 80, 'Frequency (Other)': 120,
            'Daily Dose': 80, 'Route': 60
        }

        for col in non_empty_columns:
            tree.heading(col, text=col)
            width = compact_widths.get(col, min(max(len(col) * 8, 60), 150))
            tree.column(col, width=width, anchor="w", minwidth=40)

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

    # ------------------------------------------------------------------
    # CM (Concomitant Medications) — from parsed Main sheet data
    # ------------------------------------------------------------------

    def show_cm_matrix_from_data(self, cm_data, pat):
        """Display CM data from parsed Main sheet columns as a structured table."""
        self._show_simple_matrix(
            cm_data, pat,
            title="Concomitant Medications", prefix="CM_Matrix", geometry="1400x600",
            column_order=['CM #', 'Medication', 'Dose', 'Dose Unit', 'Frequency',
                          'Daily Dose', 'Route', 'Indication', 'Start Date', 'End Date'],
            col_widths={'CM #': 40, 'Medication': 140, 'Dose': 50, 'Dose Unit': 70,
                        'Route': 80, 'Indication': 140, 'Start Date': 80, 'End Date': 70,
                        'Frequency': 90, 'Daily Dose': 80},
        )

    # ------------------------------------------------------------------
    # MH (Medical History)
    # ------------------------------------------------------------------

    def show_mh_matrix(self, mh_data, pat):
        """Display Medical History data as a structured table."""
        self._show_simple_matrix(
            mh_data, pat,
            title="Medical History", prefix="MedicalHistory", geometry="1100x500",
            column_order=['MH #', 'Condition', 'Body System', 'Category', 'Start Date', 'End Date'],
            col_widths={'MH #': 50, 'Condition': 200, 'Body System': 150, 'Category': 120,
                        'Start Date': 100, 'End Date': 100},
        )

    # ------------------------------------------------------------------
    # HFH (Heart Failure History)
    # ------------------------------------------------------------------

    def show_hfh_matrix(self, hfh_data, pat):
        """Display Heart Failure History data as a structured table."""
        self._show_simple_matrix(
            hfh_data, pat,
            title="Heart Failure History", prefix="HeartFailureHistory", geometry="900x400",
            column_order=['HFH #', 'Hospitalization Date', 'Details', 'Number of Hospitalizations'],
            col_widths={'HFH #': 50, 'Hospitalization Date': 150, 'Details': 300,
                        'Number of Hospitalizations': 80},
            exclude_keys=(),
        )

    # ------------------------------------------------------------------
    # HMEH (Hospitalization and Medical Events History)
    # ------------------------------------------------------------------

    def show_hmeh_matrix(self, hmeh_data, pat):
        """Display Hospitalization and Medical Events History data."""
        self._show_simple_matrix(
            hmeh_data, pat,
            title="Hospitalization & Medical Events History",
            prefix="HospMedEventsHistory", geometry="900x450",
            column_order=['HMEH #', 'Event Date', 'Event Details'],
            col_widths={'HMEH #': 60, 'Event Date': 120, 'Event Details': 400},
            exclude_keys=(),
        )

    # ------------------------------------------------------------------
    # CVC (Cardiac and Venous Catheterization)
    # ------------------------------------------------------------------

    def show_cvc_matrix(self, df, pat, table_type):
        """Display CVC data as a structured table."""
        if df is None or df.empty:
            messagebox.showinfo("Info", f"No CVC {table_type} data found.")
            return

        win = tk.Toplevel(self.app.root)
        win.title(f"CVC {table_type} - Patient {pat}")
        win.geometry("1000x350")
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

        self._cvc_matrix_df = df
        self._cvc_matrix_patient = pat
        self._cvc_matrix_type = table_type

        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X, side=tk.TOP)

        tk.Label(toolbar, text=f"CVC {table_type}", bg="#f4f4f4", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=(10, 20))
        tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
        tk.Button(toolbar, text="Export XLSX",
                  command=lambda: self._export_matrix('xlsx', self._cvc_matrix_df, self._cvc_matrix_patient,
                                                      f"CVC_{self._cvc_matrix_type}"),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Export CSV",
                  command=lambda: self._export_matrix('csv', self._cvc_matrix_df, self._cvc_matrix_patient,
                                                      f"CVC_{self._cvc_matrix_type}"),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        # Tree view
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        display_columns = list(df.columns)

        tree = ttk.Treeview(tree_frame, columns=display_columns, show='headings')

        for col in display_columns:
            tree.heading(col, text=col)
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

    # ------------------------------------------------------------------
    # CVH (Cardiovascular History)
    # ------------------------------------------------------------------

    def show_cvh_matrix(self, cvh_data, pat):
        """Display Cardiovascular History data as a structured table."""
        win = tk.Toplevel(self.app.root)
        win.title(f"Cardiovascular History - Patient {pat}")
        win.geometry("800x400")
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

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

        for i, record in enumerate(cvh_data, 1):
            tree.insert("", "end", values=(
                i,
                record.get('Date', ''),
                record.get('Type of Intervention', ''),
                record.get('Intervention', '')
            ))

        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    # ACT (ACT Lab Results / Heparin)
    # ------------------------------------------------------------------

    def show_act_matrix(self, act_events, pat):
        """Display ACT/Heparin data as a chronological table."""
        win = tk.Toplevel(self.app.root)
        win.title(f"ACT Lab Results - Patient {pat}")
        win.geometry("600x400")
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

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

        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

    # ------------------------------------------------------------------
    # Frequency parsing helper (used by CM displays)
    # ------------------------------------------------------------------

    def parse_frequency_multiplier(self, freq_str, freq_other_str=""):
        """Parse frequency string and return (multiplier, display_note, override_daily_dose).

        Mirrors the logic from FUHighlightsExporter for consistency.
        """
        if not freq_str or str(freq_str).lower() in ['nan', 'none', '']:
            return 1, "", None

        freq = str(freq_str).strip().lower()

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
            if freq_other_str and str(freq_other_str).lower() not in ['nan', 'none', '']:
                other = str(freq_other_str).strip().lower()

                mg_matches = re.findall(r'(\d+(?:\.\d+)?)\s*mg', other)
                if len(mg_matches) > 1:
                    total_dose = sum(float(m) for m in mg_matches)
                    return None, f"({freq_other_str})", total_dose

                if "every other day" in other or "qod" in other:
                    return 0.5, "(every 48h)", None

                match = re.match(r'q\s*(\d+)\s*h', other)
                if match:
                    interval_hours = int(match.group(1))
                    if interval_hours > 0:
                        doses_per_day = 24 // interval_hours
                        return doses_per_day, f"(q{interval_hours}h->{doses_per_day}x/d)", None

                if "continuous" in other:
                    return None, "(continuous)", None

                return 1, f"({str(freq_other_str).strip()})", None
            return 1, "", None

        return 1, "", None
