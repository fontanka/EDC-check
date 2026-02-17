"""
Visit Schedule UI Module
========================
Extracted from clinical_viewer1.py — provides the Visit Schedule Matrix window.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import logging

from config import VISIT_SCHEDULE

logger = logging.getLogger(__name__)


class VisitScheduleWindow:
    """Manages the Visit Schedule Matrix popup window.

    Args:
        app: ClinicalDataMasterV30 instance — provides df_main, root.
    """

    def __init__(self, app):
        self.app = app

    def show(self):
        """Display a visit schedule matrix: Patients (rows) x Visit Types (columns) with dates."""
        if self.app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        # Get all patients - filter to enrolled only (exclude screen failures)
        if 'Status' in self.app.df_main.columns:
            enrolled_mask = self.app.df_main['Status'].astype(str).str.lower().isin(['enrolled', 'early withdrawal'])
            patients_df = self.app.df_main[enrolled_mask]
        else:
            patients_df = self.app.df_main

        patients = patients_df['Screening #'].dropna().unique()

        # Build matrix data
        matrix_data = []

        for pat_id in sorted(patients):
            row_data = {"Patient": str(pat_id)}

            pat_rows = self.app.df_main[self.app.df_main['Screening #'] == pat_id]
            if pat_rows.empty:
                continue
            pat_row = pat_rows.iloc[0]

            # Check if patient died
            patient_died = False
            death_date = None
            if 'LOGS_DTH_DDDTC' in self.app.df_main.columns:
                death_val = pat_row.get('LOGS_DTH_DDDTC')
                if pd.notna(death_val):
                    clean_val = str(death_val).replace('|', '').strip()
                    if clean_val and clean_val.lower() not in ['nan', '']:
                        patient_died = True
                        death_date = str(death_val).split('T')[0] if 'T' in str(death_val) else str(death_val)[:10]

            # Check for early withdrawal
            patient_early_withdrawal = False
            if 'Status' in self.app.df_main.columns:
                status = str(pat_row.get('Status', '')).lower()
                if 'early withdrawal' in status or 'early-withdrawal' in status:
                    patient_early_withdrawal = True

            if patient_died:
                end_status = "Death"
            elif patient_early_withdrawal:
                end_status = "Withdrawn"
            else:
                end_status = None

            # Get visit dates
            for date_col, visit_label in VISIT_SCHEDULE:
                if date_col in self.app.df_main.columns:
                    date_val = pat_row.get(date_col)
                    if pd.notna(date_val) and str(date_val).strip() not in ['', 'nan']:
                        date_str = str(date_val)
                        if 'T' in date_str:
                            date_str = date_str.split('T')[0]
                        elif len(date_str) > 10:
                            date_str = date_str[:10]
                        row_data[visit_label] = date_str
                    elif end_status:
                        row_data[visit_label] = end_status
                    else:
                        row_data[visit_label] = "Pending"
                else:
                    if end_status:
                        row_data[visit_label] = end_status
                    else:
                        row_data[visit_label] = "Pending"

            matrix_data.append(row_data)

        if not matrix_data:
            messagebox.showinfo("No Data", "No patient visit data found.")
            return

        # Store for export
        self._schedule_df = pd.DataFrame(matrix_data)

        # Create popup window
        win = tk.Toplevel(self.app.root)
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
        tk.Button(btn_frame, text="Export XLSX", command=lambda: self._export('xlsx'),
                  bg="white", fg="#16a085", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="Export CSV", command=lambda: self._export('csv'),
                  bg="white", fg="#16a085", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=3)

        # Summary
        summary_frame = tk.Frame(win, padx=10, pady=5)
        summary_frame.pack(fill=tk.X)
        tk.Label(summary_frame, text=f"Patients: {len(matrix_data)} | Visits: {len(VISIT_SCHEDULE)}",
                 font=("Segoe UI", 10)).pack(side=tk.LEFT)

        # Tree container with canvas for scrolling
        container = tk.Frame(win)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        columns = ["Patient"] + [v[1] for v in VISIT_SCHEDULE]

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

                if value == "Death":
                    fg_color = "red"
                    bg_color = "#ffe6e6"
                elif value == "Withdrawn":
                    fg_color = "#0066CC"
                    bg_color = "#e6f0ff"
                elif value == "Pending":
                    fg_color = "#CC8800"
                    bg_color = "#fff8e6"
                else:
                    fg_color = "#228B22"
                    bg_color = "#e6ffe6"

                if col_idx == 0:
                    fg_color = "black"
                    bg_color = "#f5f5f5"

                lbl = tk.Label(scrollable_frame, text=value, font=("Segoe UI", 9),
                              fg=fg_color, bg=bg_color, padx=5, pady=2, width=12, anchor="center")
                lbl.grid(row=row_idx, column=col_idx, sticky="nsew", padx=1, pady=1)

        canvas.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        container.grid_rowconfigure(0, weight=1)
        container.grid_columnconfigure(0, weight=1)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

    def _export(self, fmt):
        """Export visit schedule data to file."""
        if self._schedule_df is None or self._schedule_df.empty:
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
                self._schedule_df.to_excel(path, index=False)
            else:
                self._schedule_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Visit schedule exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
