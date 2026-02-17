"""FU Highlights Dialog — configuration UI for follow-up highlights export."""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import logging

from fu_highlights_export import FUHighlightsExporter

logger = logging.getLogger(__name__)


class FUHighlightsDialog:
    """Follow-up Highlights export configuration dialog."""

    def __init__(self, app):
        self.app = app

    def show(self):
        """Show dialog to generate FU Highlights tables for copy-paste."""
        if self.app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.app.root)
        win.title("FU Highlights Generator")
        win.geometry("700x600")
        win.configure(bg="#f4f4f4")

        # 1. Patient Selection
        pat_frame = tk.LabelFrame(win, text=" Patient ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        tk.Label(pat_frame, text="Select Patient:", bg="#f4f4f4").pack(side=tk.LEFT, padx=5)
        all_patients = sorted(self.app.df_main['Screening #'].dropna().unique())

        self._exclude_sf_var = tk.BooleanVar(value=False)

        self._patient_var = tk.StringVar()
        if all_patients:
            self._patient_var.set(all_patients[0])

        pat_combo = ttk.Combobox(pat_frame, textvariable=self._patient_var, values=all_patients, state="readonly", width=20)
        pat_combo.pack(side=tk.LEFT, padx=5)

        def update_patient_list():
            if self._exclude_sf_var.get():
                filtered = [p for p in all_patients if not self.app._is_screen_failure(p)]
            else:
                filtered = all_patients
            pat_combo['values'] = filtered
            if filtered and (not self._patient_var.get() or self._patient_var.get() not in filtered):
                self._patient_var.set(filtered[0])

        tk.Checkbutton(pat_frame, text="Exclude Screen Failures", variable=self._exclude_sf_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=update_patient_list).pack(side=tk.LEFT, padx=15)

        # 2. Visit Selection
        vis_frame = tk.LabelFrame(win, text=" Follow-up Visit ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        vis_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(vis_frame, text="Select Visit:", bg="#f4f4f4").pack(side=tk.LEFT, padx=5)
        fu_visits = ["30D", "6M", "1Y", "2Y", "4Y"]
        self._visit_var = tk.StringVar(value="30D")
        vis_combo = ttk.Combobox(vis_frame, textvariable=self._visit_var, values=fu_visits, state="readonly", width=15)
        vis_combo.pack(side=tk.LEFT, padx=5)

        tk.Button(vis_frame, text="Generate Tables", command=lambda: self._generate_tables(win),
                  bg="#1abc9c", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=20)

        # 3. Clinical Parameters Table
        params_frame = tk.LabelFrame(win, text=" Clinical Parameters ", padx=5, pady=5, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        params_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        params_cols = ("Parameter", "Baseline", "Discharge", "FU")
        self._params_tree = ttk.Treeview(params_frame, columns=params_cols, show="headings", height=11)
        for col in params_cols:
            self._params_tree.heading(col, text=col)
            self._params_tree.column(col, width=150 if col == "Parameter" else 100, anchor="center")
        self._params_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        params_scroll = ttk.Scrollbar(params_frame, orient=tk.VERTICAL, command=self._params_tree.yview)
        params_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._params_tree.config(yscrollcommand=params_scroll.set)

        # 4. Vital Signs Table
        vitals_frame = tk.LabelFrame(win, text=" Vital Signs ", padx=5, pady=5, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        vitals_frame.pack(fill=tk.X, padx=10, pady=5)

        vitals_cols = ("Parameter", "Value")
        self._vitals_tree = ttk.Treeview(vitals_frame, columns=vitals_cols, show="headings", height=5)
        for col in vitals_cols:
            self._vitals_tree.heading(col, text=col)
            self._vitals_tree.column(col, width=200 if col == "Parameter" else 150, anchor="center")
        self._vitals_tree.pack(fill=tk.X)

        # 5. Diuretic History Table
        meds_frame = tk.LabelFrame(win, text=" Loop Diuretic History ", padx=5, pady=5, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        meds_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        meds_cols = ("Drug", "Single Dose", "Frequency", "Daily Dose", "Start Date", "End Date", "Indication")
        self._meds_tree = ttk.Treeview(meds_frame, columns=meds_cols, show="headings", height=5)
        for col in meds_cols:
            self._meds_tree.heading(col, text=col)
            if col == "Drug":
                self._meds_tree.column(col, width=100, anchor="center")
            elif col == "Indication":
                self._meds_tree.column(col, width=200, anchor="w")
            elif col in ("Start Date", "End Date"):
                self._meds_tree.column(col, width=90, anchor="center")
            else:
                self._meds_tree.column(col, width=80, anchor="center")

        meds_scroll = ttk.Scrollbar(meds_frame, orient=tk.VERTICAL, command=self._meds_tree.yview)
        meds_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        self._meds_tree.config(yscrollcommand=meds_scroll.set)
        self._meds_tree.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        # Storage
        self._highlights_df = None
        self._vitals_df = None
        self._diuretics_df = None
        self._visit_label = ""
        self._patient_label = ""

        # Buttons
        btn_frame = tk.Frame(win, bg="#f4f4f4")
        btn_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Button(btn_frame, text="Copy All to Clipboard", command=self._copy_output,
                  bg="#2c3e50", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Export XLSX", command=self._export_xlsx,
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(btn_frame, text="Export CSV", command=self._export_csv,
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        self._include_prn = tk.BooleanVar(value=False)
        tk.Checkbutton(btn_frame, text="Include PRN", variable=self._include_prn,
                       bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="View Timeline", command=self._show_diuretic_timeline,
                  bg="#9b59b6", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Close", command=win.destroy,
                  bg="#e74c3c", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)

    def _ask_confirm_fuzzy(self, original_name, matched_name):
        return messagebox.askyesno(
            "Confirm Medication Match",
            f"Fuzzy match detected:\n\nOriginal: '{original_name}'\nMatched: '{matched_name}'\n\nIs this correct?",
            parent=self.app.root
        )

    def _generate_tables(self, win):
        patient_id = self._patient_var.get()
        visit = self._visit_var.get()

        if not patient_id:
            messagebox.showwarning("No Patient", "Please select a patient.")
            return

        exporter = FUHighlightsExporter(self.app.df_main, fuzzy_confirm_callback=self._ask_confirm_fuzzy)
        df_highlights, df_vitals, df_diuretics = exporter.generate_highlights_table(patient_id, visit)

        if df_highlights is None:
            messagebox.showwarning("No Data", f"No data found for patient {patient_id}")
            return

        self._highlights_df = df_highlights
        self._vitals_df = df_vitals
        self._diuretics_df = df_diuretics
        self._visit_label = visit
        self._patient_label = patient_id

        # Populate Clinical Parameters Tree
        headers = list(df_highlights.columns)
        self._params_tree["columns"] = headers
        for col in headers:
            self._params_tree.heading(col, text=col)
            self._params_tree.column(col, width=200 if col == "Parameter" else 100, anchor="center")

        self._params_tree.delete(*self._params_tree.get_children())
        for values in df_highlights.values:
            self._params_tree.insert("", "end", values=list(values))

        # Populate Vital Signs Tree
        vitals_headers = list(df_vitals.columns)
        self._vitals_tree["columns"] = vitals_headers
        for col in vitals_headers:
            self._vitals_tree.heading(col, text=col)
            self._vitals_tree.column(col, width=200 if col == "Parameter" else 100, anchor="center")

        self._vitals_tree.delete(*self._vitals_tree.get_children())
        for values in df_vitals.values:
            self._vitals_tree.insert("", "end", values=list(values))

        # Populate Diuretic History Tree
        self._meds_tree.delete(*self._meds_tree.get_children())
        for values in df_diuretics.values:
            self._meds_tree.insert("", "end", values=list(values))

    def _copy_output(self):
        if self._highlights_df is None:
            messagebox.showwarning("No Data", "Generate tables first before copying.")
            return

        try:
            output = f"=== {self._visit_label} FU Highlights - {self._patient_label} ===\n\n"
            output += "--- Clinical Parameters ---\n"
            output += self._highlights_df.to_csv(sep="\t", index=False)
            output += "\n--- Vital Signs ---\n"
            output += self._vitals_df.to_csv(sep="\t", index=False)
            output += "\n--- Loop Diuretic History (All Events) ---\n"
            output += self._diuretics_df.to_csv(sep="\t", index=False)

            self.app.root.clipboard_clear()
            self.app.root.clipboard_append(output)
            messagebox.showinfo("Copied", "Content copied to clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to copy: {str(e)}")

    def _export_xlsx(self):
        if self._highlights_df is None:
            messagebox.showwarning("No Data", "Generate tables first.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            initialfile=f"FU_Highlights_{self._patient_label}_{self._visit_label}.xlsx",
            title="Export to Excel"
        )
        if not filename:
            return

        try:
            with pd.ExcelWriter(filename, engine='openpyxl') as writer:
                self._highlights_df.to_excel(writer, sheet_name="Clinical Params", index=False)
                self._vitals_df.to_excel(writer, sheet_name="Vital Signs", index=False)
                self._diuretics_df.to_excel(writer, sheet_name="Diuretic History", index=False)
            messagebox.showinfo("Success", f"Exported to {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

    def _export_csv(self):
        if self._highlights_df is None:
            messagebox.showwarning("No Data", "Generate tables first.")
            return

        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=f"FU_Highlights_{self._patient_label}_{self._visit_label}.csv",
            title="Export to CSV"
        )
        if not filename:
            return

        try:
            with open(filename, 'w', newline='', encoding='utf-8') as f:
                f.write(f"=== {self._visit_label} FU Highlights - {self._patient_label} ===\n")
                f.write("\n--- Clinical Parameters ---\n")
                self._highlights_df.to_csv(f, index=False)
                f.write("\n--- Vital Signs ---\n")
                self._vitals_df.to_csv(f, index=False)
                f.write("\n--- Loop Diuretic History (All Events) ---\n")
                self._diuretics_df.to_csv(f, index=False)
            messagebox.showinfo("Success", f"Exported to {os.path.basename(filename)}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {str(e)}")

    def _show_diuretic_timeline(self):
        if not self._patient_var.get():
            messagebox.showwarning("No Patient", "Select a patient first.")
            return

        patient_id = self._patient_var.get()

        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            exporter = FUHighlightsExporter(self.app.df_main)
            include_prn = self._include_prn.get()
            fig = exporter.generate_diuretic_timeline(patient_id, include_prn=include_prn)

            if fig is None:
                messagebox.showinfo("No Data", "No loop diuretic data found for this patient.")
                return

            chart_win = tk.Toplevel(self.app.root)
            chart_win.title(f"Diuretic Timeline - Patient {patient_id}")
            fig_width, fig_height = fig.get_size_inches()
            win_width = int(fig_width * 80)
            win_height = int(fig_height * 80) + 50
            chart_win.geometry(f"{win_width}x{win_height}")

            canvas = FigureCanvasTkAgg(fig, master=chart_win)
            canvas.draw()
            canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

            from matplotlib.backends.backend_tkagg import NavigationToolbar2Tk
            toolbar_frame = tk.Frame(chart_win)
            toolbar_frame.pack(fill=tk.X)
            toolbar = NavigationToolbar2Tk(canvas, toolbar_frame)
            toolbar.update()

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate timeline: {str(e)}")
