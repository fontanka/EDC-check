"""
Data Gaps Analysis Module
=========================
Extracted from clinical_viewer1.py — provides the Data Gaps Report window.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import logging

from config import VISIT_MAP, ASSESSMENT_RULES, VISIT_SCHEDULE

logger = logging.getLogger("GapAnalysis")


class DataGapsWindow:
    """Manages the Data Gaps Report popup window.
    
    Args:
        app: ClinicalDataMasterV30 instance — provides df_main, cb_site,
             cb_pat, generate_view, current_patient_gaps, root.
    """

    def __init__(self, app):
        self.app = app

    def show(self):
        """Display all missing data (gaps) per patient, organized by visit."""
        app = self.app

        if app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        # Create popup window
        win = tk.Toplevel(app.root)
        win.title("Data Gaps Report")
        win.geometry("1400x800")
        win.transient(app.root)
        win.lift()
        win.focus_force()

        # Header
        header = tk.Frame(win, bg="#c0392b", padx=10, pady=10)
        header.pack(fill=tk.X)
        tk.Label(header, text="Data Gaps Report - Missing Values by Patient/Visit",
                 font=("Segoe UI", 14, "bold"), bg="#c0392b", fg="white").pack(side=tk.LEFT)

        # Filter and export controls
        ctrl_frame = tk.Frame(win, padx=10, pady=5)
        ctrl_frame.pack(fill=tk.X)

        # Patient filter dropdown
        tk.Label(ctrl_frame, text="Patient:", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(5, 2))

        # Get list of patients
        patient_list = ["All Patients"]
        for _, row in app.df_main.iterrows():
            pat_id = row.get('Screening #')
            status = str(row.get('Status', '')).strip()
            if pd.notna(pat_id):
                pat_str = str(pat_id).replace('.0', '')
                # Skip screen failures for default list
                if 'screen' in status.lower() and 'fail' in status.lower():
                    continue
                if pat_str not in patient_list:
                    patient_list.append(pat_str)

        patient_var = tk.StringVar(value="All Patients")
        patient_combo = ttk.Combobox(ctrl_frame, textvariable=patient_var, values=patient_list,
                                      width=15, state="readonly")
        patient_combo.pack(side=tk.LEFT, padx=5)

        # Screen failure filter
        exclude_sf_var = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl_frame, text="Exclude Screen Failures", variable=exclude_sf_var,
                      font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=10)

        # Hide future visits filter
        hide_future_var = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl_frame, text="Hide Future Visits", variable=hide_future_var,
                      font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=10)

        # Second row for additional filters
        ctrl_frame2 = tk.Frame(win, padx=10, pady=2)
        ctrl_frame2.pack(fill=tk.X)

        # Visit filter dropdown
        tk.Label(ctrl_frame2, text="Visit:", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(5, 2))
        visit_list = ["All Visits"] + list(VISIT_MAP.values())
        visit_var = tk.StringVar(value="All Visits")
        visit_combo = ttk.Combobox(ctrl_frame2, textvariable=visit_var, values=visit_list,
                                    width=20, state="readonly")
        visit_combo.pack(side=tk.LEFT, padx=5)

        # Form filter dropdown
        tk.Label(ctrl_frame2, text="Form:", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(15, 2))
        form_list = ["All Forms"] + sorted(list(set([r[2] for r in ASSESSMENT_RULES])))
        form_var = tk.StringVar(value="All Forms")
        form_combo = ttk.Combobox(ctrl_frame2, textvariable=form_var, values=form_list,
                                   width=35, state="readonly")
        form_combo.pack(side=tk.LEFT, padx=5)

        # Export buttons
        btn_frame = tk.Frame(ctrl_frame)
        btn_frame.pack(side=tk.RIGHT)

        # Summary label (will be updated)
        summary_label = tk.Label(ctrl_frame2, text="", font=("Segoe UI", 10))
        summary_label.pack(side=tk.LEFT, padx=20)

        # Tree container
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Create tree with columns
        columns = ("Patient", "Status", "Visit", "Form", "Field", "DB Column")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="extended")

        # Configure columns
        tree.column("Patient", width=80, anchor="center")
        tree.column("Status", width=120, anchor="center")
        tree.column("Visit", width=150, anchor="w")
        tree.column("Form", width=200, anchor="w")
        tree.column("Field", width=300, anchor="w")
        tree.column("DB Column", width=300, anchor="w")

        for col in columns:
            tree.heading(col, text=col)

        # Configure tags
        tree.tag_configure('gap', foreground='red')
        tree.tag_configure('enrolled', foreground='black')
        tree.tag_configure('screenfail', foreground='gray')

        # Progress bar for gap scanning
        progress_frame = tk.Frame(win)
        progress_frame.pack(fill=tk.X, padx=10)
        progress_bar = ttk.Progressbar(progress_frame, mode='determinate', length=400)
        progress_bar.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        progress_text = tk.Label(progress_frame, text="", font=("Segoe UI", 9))
        progress_text.pack(side=tk.LEFT)
        cancel_btn = tk.Button(progress_frame, text="Cancel", state="disabled",
                               bg="#e74c3c", fg="white", font=("Segoe UI", 8, "bold"))
        cancel_btn.pack(side=tk.LEFT, padx=5)

        app._gaps_cancel = False

        def refresh_gaps():
            """Build and display gaps data — non-blocking chunked approach.

            Uses cached gap data from ViewBuilder when available for
            single-patient mode, avoiding a full rescan.
            """
            for i in tree.get_children():
                tree.delete(i)

            exclude_sf = exclude_sf_var.get()
            hide_future = hide_future_var.get()
            selected_patient = patient_var.get()
            selected_visit = visit_var.get()
            selected_form = form_var.get()

            # Fast path: use cached gaps for single-patient mode
            cached_gaps = getattr(app, 'current_patient_gaps', [])
            if (selected_patient != "All Patients"
                    and cached_gaps
                    and hasattr(app, 'cb_pat')
                    and app.cb_pat.get() == selected_patient):
                gaps_data = []
                for gap in cached_gaps:
                    g_visit = gap.get('visit', '')
                    g_form = gap.get('form', '')
                    if selected_visit != "All Visits" and g_visit != selected_visit:
                        continue
                    if selected_form != "All Forms" and g_form != selected_form:
                        continue
                    # Look up patient status from df_main
                    row_match = app.df_main[app.df_main['Screening #'] == selected_patient]
                    status = str(row_match.iloc[0].get('Status', '')) if not row_match.empty else ''
                    gaps_data.append({
                        'Patient': selected_patient,
                        'Status': status,
                        'Visit': g_visit,
                        'Form': g_form,
                        'Field': gap.get('field', ''),
                        'DB Column': gap.get('variable', ''),
                    })
                for g in gaps_data:
                    tree.insert("", tk.END, values=(
                        g['Patient'], g['Status'], g['Visit'],
                        g['Form'], g['Field'], g['DB Column']
                    ), tags=('gap',))
                app.gaps_data_df = pd.DataFrame(gaps_data)
                summary_label.config(
                    text=f"Done (cached) — Gaps: {len(gaps_data)} | Patient: {selected_patient}"
                )
                progress_bar['value'] = 1
                progress_bar['maximum'] = 1
                return

            # Save current combo selections
            orig_site = app.cb_site.get()
            orig_pat = app.cb_pat.get()

            # Pre-collect patient info from DataFrame (fast, no UI)
            patients_to_scan = []
            for _, row in app.df_main.iterrows():
                pat_id = row.get('Screening #')
                site_id = row.get('Site #')
                if pd.isna(pat_id) or pd.isna(site_id):
                    continue
                pat_id = str(pat_id).replace('.0', '')
                site_id = str(site_id).replace('.0', '')

                if selected_patient != "All Patients" and pat_id != selected_patient:
                    continue

                status = str(row.get('Status', '')).strip()
                is_screenfail = 'screen' in status.lower() and 'fail' in status.lower()

                if exclude_sf and is_screenfail:
                    continue

                # Determine which visits have occurred
                visits_occurred = set()
                if hide_future:
                    for date_col, visit_label in VISIT_SCHEDULE:
                        if date_col in app.df_main.columns:
                            date_val = row.get(date_col)
                            if pd.notna(date_val) and str(date_val).strip() not in ['', 'nan']:
                                visit_prefix = date_col.split('_')[0]
                                if visit_prefix in VISIT_MAP:
                                    visits_occurred.add(VISIT_MAP[visit_prefix])

                patients_to_scan.append((pat_id, site_id, status, visits_occurred))

            total = len(patients_to_scan)
            if total == 0:
                summary_label.config(text="No patients to scan.")
                return

            progress_bar['maximum'] = total
            progress_bar['value'] = 0
            app._gaps_cancel = False
            cancel_btn.config(state="normal", command=lambda: setattr(app, '_gaps_cancel', True))

            gaps_data = []
            patients_with_gaps = set()
            selected_visit = visit_var.get()
            selected_form = form_var.get()
            idx = [0]  # mutable counter for closure

            def process_next():
                """Process one patient per tick, keeping UI responsive."""
                if app._gaps_cancel or idx[0] >= total:
                    # Finished or cancelled — finalize
                    app.cb_site.set(orig_site)
                    app.cb_pat.set(orig_pat)
                    if orig_site and orig_pat:
                        app.generate_view()

                    for gap in gaps_data:
                        tree.insert("", tk.END, values=(
                            gap['Patient'], gap['Status'], gap['Visit'],
                            gap['Form'], gap['Field'], gap['DB Column']
                        ), tags=('gap',))

                    total_gaps = len(gaps_data)
                    label = "Cancelled" if app._gaps_cancel else "Done"
                    summary_label.config(
                        text=f"{label} — Gaps: {total_gaps} | Patients with gaps: {len(patients_with_gaps)}"
                    )
                    app.gaps_data_df = pd.DataFrame(gaps_data)
                    progress_bar['value'] = total if not app._gaps_cancel else idx[0]
                    progress_text.config(text="")
                    cancel_btn.config(state="disabled")
                    return

                pat_id, site_id, status, visits_occurred = patients_to_scan[idx[0]]

                # Update progress
                progress_bar['value'] = idx[0]
                progress_text.config(text=f"{idx[0]+1}/{total}: {pat_id}")

                # Extract gaps directly from data (no UI rendering needed)
                row = app.df_main[app.df_main['Screening #'] == pat_id]
                if row.empty:
                    idx[0] += 1
                    app.root.after(1, process_next)
                    return
                row = row.iloc[0]

                # Collect gaps: find required fields that are empty
                for col in app.df_main.columns:
                    val = row[col]
                    if pd.notna(val) and str(val).strip() not in ("", "nan"):
                        continue  # Has data — not a gap

                    # Identify which visit/form this column belongs to
                    info = app.view_builder._identify_column(col) if hasattr(app, 'view_builder') else None
                    if not info:
                        continue
                    visit, form, category = info

                    if hide_future and visit not in visits_occurred:
                        continue
                    if selected_visit != "All Visits" and visit != selected_visit:
                        continue
                    if selected_form != "All Forms" and form != selected_form:
                        continue

                    field_label = app.labels.get(col, col)
                    gap = {
                        'Patient': pat_id,
                        'Status': status,
                        'Visit': visit,
                        'Form': form,
                        'Field': field_label,
                        'DB Column': col,
                    }
                    gaps_data.append(gap)
                    patients_with_gaps.add(pat_id)

                idx[0] += 1
                # Schedule next patient — yields to event loop between patients
                app.root.after(1, process_next)

            summary_label.config(text="Scanning...")
            process_next()

        def export_gaps(fmt):
            if not hasattr(app, 'gaps_data_df') or app.gaps_data_df is None or app.gaps_data_df.empty:
                messagebox.showwarning("No Data", "No gaps data to export.")
                return
            ext = 'xlsx' if fmt == 'xlsx' else 'csv'
            path = filedialog.asksaveasfilename(
                defaultextension=f".{ext}",
                filetypes=[(f"{ext.upper()} files", f"*.{ext}")],
                title=f"Export Data Gaps as {ext.upper()}"
            )
            if not path:
                return
            try:
                if fmt == 'xlsx':
                    app.gaps_data_df.to_excel(path, index=False)
                else:
                    app.gaps_data_df.to_csv(path, index=False)
                messagebox.showinfo("Success", f"Data gaps exported to:\n{path}")
            except Exception as e:
                messagebox.showerror("Error", f"Export failed: {e}")

        # Add export buttons
        tk.Button(btn_frame, text="Export XLSX", command=lambda: export_gaps('xlsx'),
                  bg="white", fg="#c0392b", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="Export CSV", command=lambda: export_gaps('csv'),
                  bg="white", fg="#c0392b", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=3)

        # Refresh button
        tk.Button(ctrl_frame, text="Refresh", command=refresh_gaps,
                  bg="#c0392b", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        # Scrollbars
        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        # Initial load
        refresh_gaps()
