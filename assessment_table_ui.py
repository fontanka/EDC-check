"""
Assessment Table UI Module
===========================
Extracted from clinical_viewer1.py — provides the Assessment Data Table window.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import json
import logging
from datetime import datetime

from assessment_data_table import AssessmentDataExtractor, ASSESSMENT_CATEGORIES

logger = logging.getLogger(__name__)


class AssessmentTableWindow:
    """Manages the Assessment Data Table popup window.

    Args:
        app: ClinicalDataMasterV30 instance — provides df_main, labels, root.
    """

    def __init__(self, app):
        self.app = app

    def show(self):
        """Display Assessment Data Table window with dropdowns for assessment, parameter, and visits."""
        if self.app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        # Create popup window
        win = tk.Toplevel(self.app.root)
        win.title("Assessment Data Table")
        win.geometry("1200x800")
        win.configure(bg="#f4f4f4")
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

        self._win = win
        self._extractor = AssessmentDataExtractor(self.app.df_main, self.app.labels)

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
        self._type_var = tk.StringVar()
        assess_types = list(ASSESSMENT_CATEGORIES.keys())
        self._type_cb = ttk.Combobox(row1, textvariable=self._type_var,
                                     values=assess_types, state="readonly", width=20)
        self._type_cb.pack(side=tk.LEFT, padx=(0, 15))
        self._type_cb.bind("<<ComboboxSelected>>", self._on_type_changed)
        if assess_types:
            self._type_cb.current(0)

        # Parameter Dropdown (dynamic)
        tk.Label(row1, text="Parameter:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 5))
        self._param_var = tk.StringVar()
        self._param_cb = ttk.Combobox(row1, textvariable=self._param_var,
                                      state="readonly", width=25)
        self._param_cb.pack(side=tk.LEFT)

        # 2. Visit Selection Section
        visit_frame = tk.LabelFrame(win, text=" Visit Types ",
                                    padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        visit_frame.pack(fill=tk.X, padx=10, pady=5)

        visit_inner = tk.Frame(visit_frame, bg="#f4f4f4")
        visit_inner.pack(fill=tk.X)

        self._visit_vars = {}
        self._visit_frame = visit_inner

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

        self._exclude_sf_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self._exclude_sf_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=self._update_patient_list).pack(side=tk.LEFT)

        # Patient list buttons
        btn_row = tk.Frame(filter_row, bg="#f4f4f4")
        btn_row.pack(side=tk.RIGHT)
        tk.Button(btn_row, text="Select All", command=lambda: self._select_all_patients(True),
                  bg="#27ae60", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Select None", command=lambda: self._select_all_patients(False),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)

        # Reorder buttons
        tk.Button(btn_row, text="\u25b2 Up", command=self._move_patient_up,
                  bg="#95a5a6", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(10, 2))
        tk.Button(btn_row, text="\u25bc Down", command=self._move_patient_down,
                  bg="#95a5a6", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Save Order", command=self._save_patient_order,
                  bg="#9b59b6", fg="white", font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=(10, 2))

        # Patient listbox with scrollbar
        list_container = tk.Frame(pat_frame, bg="#f4f4f4")
        list_container.pack(fill=tk.BOTH, expand=True)

        self._pat_listbox = tk.Listbox(list_container, selectmode=tk.EXTENDED,
                                       font=("Segoe UI", 9), bg="white", height=8)
        scrollbar = ttk.Scrollbar(list_container, orient="vertical", command=self._pat_listbox.yview)
        self._pat_listbox.configure(yscrollcommand=scrollbar.set)

        self._pat_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Store patient order (for reordering)
        self._patient_order = []
        self._all_patients = sorted(self.app.df_main['Screening #'].dropna().unique())
        self._saved_order = self._load_patient_order()

        # 4. Generate and Export Buttons
        action_frame = tk.Frame(win, bg="#f4f4f4")
        action_frame.pack(fill=tk.X, padx=10, pady=10)

        tk.Button(action_frame, text="Generate Table", command=self._generate_table,
                  bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"), padx=20, pady=8).pack(side=tk.LEFT, padx=5)
        tk.Button(action_frame, text="Export XLSX", command=lambda: self._export_table('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold"), padx=15, pady=8).pack(side=tk.LEFT, padx=5)
        tk.Button(action_frame, text="Export CSV", command=lambda: self._export_table('csv'),
                  bg="#e67e22", fg="white", font=("Segoe UI", 10, "bold"), padx=15, pady=8).pack(side=tk.LEFT, padx=5)

        # 5. Results Table
        result_frame = tk.LabelFrame(win, text=" Results ",
                                     padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        result_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(5, 10))

        self._tree = ttk.Treeview(result_frame, show="headings")
        self._tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar_v = ttk.Scrollbar(result_frame, orient="vertical", command=self._tree.yview)
        scrollbar_v.pack(side=tk.RIGHT, fill=tk.Y)
        self._tree.configure(yscrollcommand=scrollbar_v.set)

        scrollbar_h = ttk.Scrollbar(result_frame, orient="horizontal", command=self._tree.xview)
        scrollbar_h.pack(side=tk.BOTTOM, fill=tk.X)
        self._tree.configure(xscrollcommand=scrollbar_h.set)

        self._result_df = None

        # Initialize dropdowns
        self._on_type_changed(None)
        self._update_patient_list()

        win.focus_set()

    def _on_type_changed(self, event):
        """Handle assessment type change — update parameters and visits."""
        assess_type = self._type_var.get()
        if not assess_type or assess_type not in ASSESSMENT_CATEGORIES:
            return

        category = ASSESSMENT_CATEGORIES[assess_type]

        # Update parameter dropdown
        params = category.get("params", [])
        param_display = [p[0] for p in params]
        self._param_cb['values'] = param_display
        if param_display:
            self._param_cb.current(0)

        # Update visit checkboxes
        for widget in self._visit_frame.winfo_children():
            widget.destroy()

        self._visit_vars = {}
        visits = category.get("visits", {})

        col = 0
        for visit_name in visits.keys():
            var = tk.BooleanVar(value=True)
            self._visit_vars[visit_name] = var
            chk = tk.Checkbutton(self._visit_frame, text=visit_name, variable=var,
                                 bg="#f4f4f4", font=("Segoe UI", 9))
            chk.grid(row=0, column=col, padx=5, sticky="w")
            col += 1

    def _select_all_visits(self, select):
        """Select or deselect all visit checkboxes."""
        for var in self._visit_vars.values():
            var.set(select)

    def _update_patient_list(self):
        """Update patient listbox based on filter settings, respecting saved order."""
        self._pat_listbox.delete(0, tk.END)

        exclude_sf = self._exclude_sf_var.get()

        available = []
        for p in self._all_patients:
            if not p:
                continue
            if exclude_sf and self.app._is_screen_failure(p):
                continue
            available.append(p)

        # Apply saved order if exists
        if self._saved_order:
            ordered = []
            for p in self._saved_order:
                if p in available:
                    ordered.append(p)
                    available.remove(p)
            ordered.extend(available)
            self._patient_order = ordered
        else:
            self._patient_order = available

        for p in self._patient_order:
            self._pat_listbox.insert(tk.END, p)

    def _save_patient_order(self):
        """Save patient order to JSON file for persistence."""
        order_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assess_patient_order.json')
        try:
            with open(order_file, 'w') as f:
                json.dump(self._patient_order, f, indent=2)
            self._saved_order = self._patient_order[:]
            messagebox.showinfo("Saved", f"Patient order saved successfully.\n({len(self._patient_order)} patients)")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save patient order: {e}")

    def _load_patient_order(self):
        """Load patient order from JSON file."""
        order_file = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assess_patient_order.json')
        try:
            if os.path.exists(order_file):
                with open(order_file, 'r') as f:
                    return json.load(f)
        except Exception:
            pass
        return []

    def _select_all_patients(self, select):
        """Select or deselect all patients in the listbox."""
        if select:
            self._pat_listbox.select_set(0, tk.END)
        else:
            self._pat_listbox.select_clear(0, tk.END)

    def _move_patient_up(self):
        """Move selected patient(s) up in the list."""
        selected = list(self._pat_listbox.curselection())
        if not selected or selected[0] == 0:
            return

        for idx in selected:
            if idx > 0:
                self._patient_order[idx - 1], self._patient_order[idx] = \
                    self._patient_order[idx], self._patient_order[idx - 1]

        self._pat_listbox.delete(0, tk.END)
        for p in self._patient_order:
            self._pat_listbox.insert(tk.END, p)

        for idx in selected:
            self._pat_listbox.select_set(idx - 1)

    def _move_patient_down(self):
        """Move selected patient(s) down in the list."""
        selected = list(self._pat_listbox.curselection())
        if not selected or selected[-1] >= len(self._patient_order) - 1:
            return

        for idx in reversed(selected):
            if idx < len(self._patient_order) - 1:
                self._patient_order[idx], self._patient_order[idx + 1] = \
                    self._patient_order[idx + 1], self._patient_order[idx]

        self._pat_listbox.delete(0, tk.END)
        for p in self._patient_order:
            self._pat_listbox.insert(tk.END, p)

        for idx in selected:
            self._pat_listbox.select_set(idx + 1)

    def _generate_table(self):
        """Generate the assessment data table based on selections."""
        assess_type = self._type_var.get()
        param_display = self._param_var.get()

        if not assess_type or not param_display:
            messagebox.showwarning("Selection Required", "Please select an assessment type and parameter.")
            return

        selected_visits = [v for v, var in self._visit_vars.items() if var.get()]
        if not selected_visits:
            messagebox.showwarning("No Visits", "Please select at least one visit type.")
            return

        selected_indices = self._pat_listbox.curselection()
        if selected_indices:
            selected_patients = [self._patient_order[i] for i in selected_indices]
        else:
            selected_patients = self._patient_order[:]

        if not selected_patients:
            messagebox.showwarning("No Patients", "Please select at least one patient.")
            return

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

        try:
            df = self._extractor.generate_table(
                selected_patients, assess_type, param_code, selected_visits
            )
            self._result_df = df

            self._tree.delete(*self._tree.get_children())

            columns = list(df.columns)
            self._tree['columns'] = columns
            for col in columns:
                self._tree.heading(col, text=col)
                self._tree.column(col, width=100, anchor="center")
            self._tree.column(columns[0], width=120, anchor="w")

            for values in df[columns].values:
                self._tree.insert("", tk.END, values=list(values))

            messagebox.showinfo("Success", f"Generated table with {len(df)} patients \u00d7 {len(selected_visits)} visits")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate table: {e}")

    def _export_table(self, fmt):
        """Export assessment table to file."""
        if self._result_df is None or self._result_df.empty:
            messagebox.showwarning("No Data", "Please generate a table first.")
            return

        assess_type = self._type_var.get().replace(" ", "_").replace("-", "_")
        param = self._param_var.get().replace(" ", "_").replace("/", "_")

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
                self._result_df.to_excel(path, index=False)
            else:
                self._result_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")
