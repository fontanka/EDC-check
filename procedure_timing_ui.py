"""
Procedure Timing UI Module
============================
Extracted from clinical_viewer1.py — provides the Procedure Timing Matrix window.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import logging
from datetime import datetime

from procedure_timing_export import ProcedureTimingExporter

logger = logging.getLogger(__name__)


class ProcedureTimingWindow:
    """Manages the Procedure Timing Matrix popup window.

    Args:
        app: ClinicalDataMasterV30 instance — provides df_main, labels, root.
    """

    def __init__(self, app):
        self.app = app

    def show(self):
        """Show Procedure Timing matrix with adjustable row order and multi-patient selection."""
        if self.app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.app.root)
        win.title("Procedure Timing Matrix")
        win.geometry("1400x800")
        win.configure(bg="#f4f4f4")

        # Get all non-screen-failure patients
        all_patients = sorted(self.app.df_main['Screening #'].dropna().unique())
        non_sf_patients = [p for p in all_patients if not self.app._is_screen_failure(p)]

        # Initialize exporter
        exporter = ProcedureTimingExporter(self.app.df_main, self.app.labels)
        self._fields = exporter.get_field_order()
        self._exporter = exporter

        # --- Layout ---
        # 1. Top Toolbar (Exports)
        top_bar = tk.Frame(win, bg="#ecf0f1", pady=10, highlightthickness=1, highlightbackground="#bdc3c7")
        top_bar.pack(fill=tk.X, side=tk.TOP)

        tk.Label(top_bar, text="Procedure Timing", bg="#ecf0f1", font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=20)

        tk.Button(top_bar, text="Export XLSX", command=lambda: self._export('xlsx'),
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=10)
        tk.Button(top_bar, text="Copy to Clipboard", command=self._copy_to_clipboard,
                  bg="#9b59b6", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=10)

        # Main content area
        main_content = tk.Frame(win, bg="#f4f4f4")
        main_content.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

        # Left Panel (Configurations)
        config_frame = tk.Frame(main_content, bg="#f4f4f4", width=350)
        config_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))
        config_frame.pack_propagate(False)

        # --- Section 1: Patients ---
        pat_group = tk.LabelFrame(config_frame, text=" 1. Select Patients ", bg="white",
                                  font=("Segoe UI", 10, "bold"), padx=5, pady=5)
        pat_group.pack(fill=tk.BOTH, expand=True, pady=(0, 15))

        pat_btn_frame = tk.Frame(pat_group, bg="white")
        pat_btn_frame.pack(fill=tk.X, pady=(0, 5))

        self._pat_vars = {}

        def toggle_all_pats(state):
            for var in self._pat_vars.values():
                var.set(state)
            self._generate_matrix()

        tk.Button(pat_btn_frame, text="All", command=lambda: toggle_all_pats(True),
                  font=("Segoe UI", 8), width=5).pack(side=tk.LEFT, padx=2)
        tk.Button(pat_btn_frame, text="None", command=lambda: toggle_all_pats(False),
                  font=("Segoe UI", 8), width=5).pack(side=tk.LEFT, padx=2)

        # Scrollable patient list
        list_container = tk.Frame(pat_group, bg="white")
        list_container.pack(fill=tk.BOTH, expand=True)

        pat_canvas = tk.Canvas(list_container, bg="white", highlightthickness=0)
        pat_scroll = ttk.Scrollbar(list_container, orient="vertical", command=pat_canvas.yview)
        pat_scrollable = tk.Frame(pat_canvas, bg="white")

        pat_scrollable.bind(
            "<Configure>",
            lambda e: pat_canvas.configure(scrollregion=pat_canvas.bbox("all"))
        )

        def _on_mousewheel(event):
            try:
                pat_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
            except Exception:
                pass

        pat_canvas.bind_all("<MouseWheel>", _on_mousewheel)

        pat_canvas.create_window((0, 0), window=pat_scrollable, anchor="nw")
        pat_canvas.configure(yscrollcommand=pat_scroll.set)

        pat_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        pat_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        for pid in non_sf_patients:
            var = tk.BooleanVar(value=False)
            self._pat_vars[pid] = var
            cb = tk.Checkbutton(pat_scrollable, text=str(pid), variable=var, bg="white", anchor="w",
                                command=self._generate_matrix, font=("Segoe UI", 9))
            cb.pack(fill=tk.X)

        # Select first patient by default
        if non_sf_patients:
            self._pat_vars[non_sf_patients[0]].set(True)

        # --- Section 2: Row Order ---
        order_group = tk.LabelFrame(config_frame, text=" 2. Row Order ", bg="white",
                                    font=("Segoe UI", 10, "bold"), padx=5, pady=5)
        order_group.pack(fill=tk.X, ipady=5)

        self._order_listbox = tk.Listbox(order_group, bg="white", height=15, font=("Segoe UI", 9),
                                         selectmode=tk.SINGLE, exportselection=False)
        self._order_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        order_scroll = ttk.Scrollbar(order_group, orient=tk.VERTICAL, command=self._order_listbox.yview)
        order_scroll.pack(side=tk.LEFT, fill=tk.Y)
        self._order_listbox.config(yscrollcommand=order_scroll.set)

        for col_name, label in self._fields:
            self._order_listbox.insert(tk.END, label)

        # Up/Down/Reset buttons
        btn_frame = tk.Frame(order_group, bg="white")
        btn_frame.pack(side=tk.LEFT, padx=5, fill=tk.Y)

        tk.Button(btn_frame, text="\u2191", command=self._move_field_up,
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"), width=3).pack(pady=2)
        tk.Button(btn_frame, text="\u2193", command=self._move_field_down,
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold"), width=3).pack(pady=2)
        tk.Button(btn_frame, text="R", command=lambda: self._reset_order(exporter),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 9, "bold"), width=3).pack(pady=10)

        # Right Panel (Matrix Preview)
        preview_group = tk.LabelFrame(main_content, text=" Matrix Preview (Auto-Pivoted) ", bg="#f4f4f4",
                                      font=("Segoe UI", 10, "bold"), padx=5, pady=5)
        preview_group.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        tree_container = tk.Frame(preview_group)
        tree_container.pack(fill=tk.BOTH, expand=True)

        self._tree = ttk.Treeview(tree_container, show='headings')

        h_scroll = ttk.Scrollbar(tree_container, orient="horizontal", command=self._tree.xview)
        v_scroll = ttk.Scrollbar(tree_container, orient="vertical", command=self._tree.yview)
        self._tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

        self._tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        tree_container.grid_rowconfigure(0, weight=1)
        tree_container.grid_columnconfigure(0, weight=1)

        style = ttk.Style()
        style.configure("Treeview", rowheight=25)
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

        # Initial display
        self._generate_matrix()

    def _move_field_up(self):
        """Move selected field up in order."""
        sel = self._order_listbox.curselection()
        if not sel or sel[0] == 0:
            return
        idx = sel[0]
        self._fields[idx], self._fields[idx - 1] = self._fields[idx - 1], self._fields[idx]
        text = self._order_listbox.get(idx)
        self._order_listbox.delete(idx)
        self._order_listbox.insert(idx - 1, text)
        self._order_listbox.selection_set(idx - 1)
        self._generate_matrix()

    def _move_field_down(self):
        """Move selected field down in order."""
        sel = self._order_listbox.curselection()
        if not sel or sel[0] >= len(self._fields) - 1:
            return
        idx = sel[0]
        self._fields[idx], self._fields[idx + 1] = self._fields[idx + 1], self._fields[idx]
        text = self._order_listbox.get(idx)
        self._order_listbox.delete(idx)
        self._order_listbox.insert(idx + 1, text)
        self._order_listbox.selection_set(idx + 1)
        self._generate_matrix()

    def _reset_order(self, exporter):
        """Reset field order to default."""
        self._fields = exporter.get_field_order()
        self._order_listbox.delete(0, tk.END)
        for col_name, label in self._fields:
            self._order_listbox.insert(tk.END, label)
        self._generate_matrix()

    def _generate_matrix(self):
        """Generate/update the procedure timing matrix (pivoted)."""
        selected_pats = [pid for pid, var in self._pat_vars.items() if var.get()]

        if not selected_pats:
            self._tree.delete(*self._tree.get_children())
            self._tree["columns"] = []
            return

        self._exporter.set_field_order(self._fields)
        df_flat = self._exporter.generate_matrix(selected_pats)

        if df_flat is None or df_flat.empty:
            return

        # PIVOT: Transpose the DataFrame
        df_t = df_flat.set_index('Patient').T

        df_final = df_t.reset_index()
        df_final.rename(columns={'index': 'Procedure Step'}, inplace=True)

        # Store for export
        self._result_df = df_final

        columns = list(df_final.columns)
        st_columns = [str(c) for c in columns]

        self._tree["columns"] = st_columns

        for col in st_columns:
            self._tree.heading(col, text=col)
            width = 300 if col == 'Procedure Step' else 100
            self._tree.column(col, width=width, anchor="w" if col == 'Procedure Step' else "center")

        self._tree.delete(*self._tree.get_children())
        for values in df_final[columns].fillna('').values:
            self._tree.insert("", "end", values=list(values))

    def _export(self, fmt):
        """Export procedure timing matrix."""
        if not hasattr(self, '_result_df') or self._result_df is None:
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
                self._result_df.to_excel(path, index=False)
            else:
                self._result_df.to_csv(path, index=False)
            messagebox.showinfo("Success", f"Exported to:\n{path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _copy_to_clipboard(self):
        """Copy procedure timing matrix to clipboard."""
        if not hasattr(self, '_result_df') or self._result_df is None:
            messagebox.showwarning("No Data", "Generate matrix first.")
            return

        try:
            output = self._result_df.to_csv(sep="\t", index=False)
            self.app.root.clipboard_clear()
            self.app.root.clipboard_append(output)
            messagebox.showinfo("Copied", "Matrix copied to clipboard!")
        except Exception as e:
            messagebox.showerror("Error", f"Copy failed: {e}")
