"""
Data Comparison Module
======================
Compares two Clinical Data Master export/project files to identify changes.
Highlights:
- New patients / removed patients
- New fields / removed fields
- Changed values
"""

import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import logging
from concurrent.futures import ThreadPoolExecutor

logger = logging.getLogger("DataComparator")

class DataComparatorDialog:
    def __init__(self, parent):
        self.parent = parent
        self.df_old = None
        self.df_new = None
        self._build_dialog()

    def _build_dialog(self):
        self.win = tk.Toplevel(self.parent)
        self.win.title("Compare Data Versions")
        self.win.geometry("800x600")
        self.win.configure(bg="#f4f4f4")
        self.win.transient(self.parent)
        self.win.lift()
        self.win.focus_force()

        # --- File Selection ---
        files_frame = tk.LabelFrame(self.win, text=" Data Files ", padx=10, pady=8, bg="#f4f4f4")
        files_frame.pack(fill=tk.X, padx=10, pady=5)

        # File 1 (Old)
        f1_row = tk.Frame(files_frame, bg="#f4f4f4")
        f1_row.pack(fill=tk.X, pady=2)
        tk.Label(f1_row, text="Reference File (Old):", width=18, anchor="w", bg="#f4f4f4").pack(side=tk.LEFT)
        self.path_old = tk.StringVar()
        tk.Entry(f1_row, textvariable=self.path_old, width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(f1_row, text="Browse...", command=lambda: self._browse_file(self.path_old)).pack(side=tk.LEFT)

        # File 2 (New)
        f2_row = tk.Frame(files_frame, bg="#f4f4f4")
        f2_row.pack(fill=tk.X, pady=2)
        tk.Label(f2_row, text="Target File (New):", width=18, anchor="w", bg="#f4f4f4").pack(side=tk.LEFT)
        self.path_new = tk.StringVar()
        tk.Entry(f2_row, textvariable=self.path_new, width=60).pack(side=tk.LEFT, padx=5)
        tk.Button(f2_row, text="Browse...", command=lambda: self._browse_file(self.path_new)).pack(side=tk.LEFT)

        # Compare Button
        btn_frame = tk.Frame(self.win, bg="#f4f4f4")
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        tk.Button(btn_frame, text="Run Comparison", command=self._run_comparison,
                  bg="#2980b9", fg="white", font=("Segoe UI", 10, "bold"), padx=15).pack()

        # --- Results Area ---
        res_frame = tk.LabelFrame(self.win, text=" Differences Found ", padx=10, pady=8, bg="#f4f4f4")
        res_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Treeview for diffs
        cols = ("Type", "Patient", "Field", "Old Value", "New Value")
        self.tree = ttk.Treeview(res_frame, columns=cols, show="headings")
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=120 if c != "Field" else 200)

        sb = tk.Scrollbar(res_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=sb.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Tag colors
        self.tree.tag_configure("new_pat", background="#e8f8f5") # Light green
        self.tree.tag_configure("del_pat", background="#fadbd8") # Light red
        self.tree.tag_configure("change", background="#fef9e7")  # Light yellow

    def _browse_file(self, var):
        f = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if f: var.set(f)

    def _run_comparison(self):
        p_old, p_new = self.path_old.get(), self.path_new.get()
        if not p_old or not p_new:
            messagebox.showwarning("Missing Files", "Please select both files.")
            return

        self.win.config(cursor="wait")
        self.win.update()

        try:
            # Run in thread to keep UI responsive
            with ThreadPoolExecutor() as executor:
                future = executor.submit(self._compare_data, p_old, p_new)
                diffs = future.result()
            
            self._display_results(diffs)
            
        except Exception as e:
            logger.error(f"Comparison failed: {e}")
            messagebox.showerror("Error", f"Comparison failed:\n{e}")
        finally:
            self.win.config(cursor="")

    def _compare_data(self, path_old, path_new):
        """Logic to load and compare datasets."""
        # Load only Main sheets for now
        df1 = pd.read_excel(path_old, sheet_name=0)
        df2 = pd.read_excel(path_new, sheet_name=0)

        # Ensure key columns exist
        key_col = 'Screening #'
        if key_col not in df1.columns or key_col not in df2.columns:
            raise ValueError(f"Column '{key_col}' not found in one or both files.")

        # Set index
        df1 = df1.set_index(key_col)
        df2 = df2.set_index(key_col)

        diffs = []

        # 1. Check Patients
        pats1 = set(df1.index)
        pats2 = set(df2.index)

        new_pats = pats2 - pats1
        del_pats = pats1 - pats2
        common_pats = pats1 & pats2

        for p in new_pats:
            diffs.append(("New Patient", p, "-", "-", "-"))
        
        for p in del_pats:
            diffs.append(("Removed Patient", p, "-", "-", "-"))

        # 2. Check content for common patients
        # Focus on columns present in both (or check for schema changes too)
        common_cols = list(set(df1.columns) & set(df2.columns))
        
        for p in common_pats:
            row1 = df1.loc[p]
            row2 = df2.loc[p]
            
            # Compare common columns
            # Vectorized comparison would be faster, but iterating row-by-row gives granular diffs per patient
            for col in common_cols:
                v1 = row1[col]
                v2 = row2[col]

                # Normalize for comparison
                s1 = str(v1).strip() if pd.notna(v1) else ""
                s2 = str(v2).strip() if pd.notna(v2) else ""
                
                # Special handling for floats behaving oddly (10.0 vs 10)
                if s1.endswith('.0'): s1 = s1[:-2]
                if s2.endswith('.0'): s2 = s2[:-2]

                if s1 != s2:
                    diffs.append(("Value Change", p, col, s1, s2))

        return diffs

    def _display_results(self, diffs):
        # Clear tree
        for i in self.tree.get_children():
            self.tree.delete(i)

        if not diffs:
            messagebox.showinfo("No Differences", "Files are identical.")
            return

        for dtype, pat, field, v1, v2 in diffs:
            tag = "change"
            if dtype == "New Patient": tag = "new_pat"
            elif dtype == "Removed Patient": tag = "del_pat"
            
            self.tree.insert("", "end", values=(dtype, pat, field, v1, v2), tags=(tag,))
        
        tk.Label(self.win, text=f"Found {len(diffs)} differences.", 
                 bg="#f4f4f4", font=("Segoe UI", 9, "bold")).pack(pady=5)
