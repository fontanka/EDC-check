"""
Export Dialogs UI Module
=========================
Extracted from clinical_viewer1.py — provides configuration dialogs for
Echo, CVC, Labs, and FU Highlights exports.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import logging

from echo_export import EchoExporter, VISIT_ORDER as ECHO_VISITS
from labs_export import LabsExporter
from fu_highlights_export import FUHighlightsExporter
from cvc_export import CVCExporter

logger = logging.getLogger(__name__)


class EchoExportDialog:
    """Echo data export configuration dialog."""

    def __init__(self, app):
        self.app = app

    def show(self):
        """Show configuration dialog for Echo Export."""
        if self.app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.app.root)
        win.title("Export Echo Data (Sponsor)")
        win.geometry("600x850")
        win.configure(bg="#f4f4f4")

        self._win = win

        # 1. Template Section
        tpl_frame = tk.LabelFrame(win, text=" Template ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        tpl_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        inner_tpl = tk.Frame(tpl_frame, bg="#f4f4f4")
        inner_tpl.pack(fill=tk.X)

        default_template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "echo_tmpl.xlsx")
        self._tpl_path_var = tk.StringVar()
        if os.path.exists(default_template):
            self._tpl_path_var.set(default_template)

        tk.Entry(inner_tpl, textvariable=self._tpl_path_var, width=50, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(inner_tpl, text="Browse...", command=self._browse_template,
                  bg="#2c3e50", fg="white", font=("Segoe UI", 9)).pack(side=tk.LEFT)

        # 2. Patient Filter Section
        pat_filter_frame = tk.LabelFrame(win, text=" Patient Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_filter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        filter_row = tk.Frame(pat_filter_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 8))

        self._exclude_sf_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self._exclude_sf_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=self._update_patient_list).pack(side=tk.LEFT)

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

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self._pat_chk_vars = {}
        self._all_patients = sorted(self.app.df_main['Screening #'].dropna().unique())
        self._update_patient_list()

        # 3. Visit Selection Section
        vis_frame = tk.LabelFrame(win, text=" Visit Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        vis_frame.pack(fill=tk.X, padx=10, pady=5)

        vis_inner = tk.Frame(vis_frame, bg="#f4f4f4")
        vis_inner.pack(fill=tk.X)

        self._vis_chk_vars = {}
        for i, v in enumerate(ECHO_VISITS):
            var = tk.BooleanVar(value=True)
            self._vis_chk_vars[v] = var
            row, col = divmod(i, 2)
            chk = tk.Checkbutton(vis_inner, text=v, variable=var, bg="#f4f4f4", font=("Segoe UI", 9), anchor="w")
            chk.grid(row=row, column=col, sticky="w", padx=(0, 20))

        # 4. Output Options Section
        opt_frame = tk.LabelFrame(win, text=" Output Options ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        opt_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(opt_frame, text="For visits without data:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")

        self._delete_empty_rows_var = tk.BooleanVar(value=True)
        tk.Radiobutton(opt_frame, text="Delete row from output (recommended)", variable=self._delete_empty_rows_var,
                       value=True, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)
        tk.Radiobutton(opt_frame, text="Leave blank row in output", variable=self._delete_empty_rows_var,
                       value=False, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)

        # 5. Generate Button
        tk.Button(win, text="Generate Echo Reports (ZIP)", command=lambda: self._generate(win),
                  bg="#2980b9", fg="white", font=("Segoe UI", 11, "bold"), pady=12, cursor="hand2").pack(fill=tk.X, padx=10, pady=15)

        win.focus_set()
        win.grab_set()

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self._tpl_path_var.set(path)

    def _update_patient_list(self):
        for widget in self._pat_scroll_frame.winfo_children():
            widget.destroy()

        exclude_sf = self._exclude_sf_var.get()

        for p in self._all_patients:
            if not p:
                continue
            if exclude_sf and self.app._is_screen_failure(p):
                continue

            if p in self._pat_chk_vars:
                var = self._pat_chk_vars[p]
            else:
                var = tk.BooleanVar(value=True)
                self._pat_chk_vars[p] = var

            tk.Checkbutton(self._pat_scroll_frame, text=p, variable=var,
                          bg="white", font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)

    def _select_all_patients(self, select):
        for widget in self._pat_scroll_frame.winfo_children():
            if isinstance(widget, tk.Checkbutton):
                pat_id = widget.cget("text")
                if pat_id in self._pat_chk_vars:
                    self._pat_chk_vars[pat_id].set(select)

    def _generate(self, dialog):
        tpl_path = self._tpl_path_var.get()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showerror("Error", "Please select a valid template file.")
            return

        exclude_sf = self._exclude_sf_var.get()
        selected_pats = []
        for p in self._all_patients:
            if p in self._pat_chk_vars and self._pat_chk_vars[p].get():
                if exclude_sf and self.app._is_screen_failure(p):
                    continue
                selected_pats.append(p)

        selected_visits = [v for v, var in self._vis_chk_vars.items() if var.get()]
        delete_empty = self._delete_empty_rows_var.get()

        if not selected_pats:
            messagebox.showwarning("Warning", "No patients selected.")
            return

        try:
            dialog.config(cursor="watch")
            dialog.update()

            exporter = EchoExporter(self.app.df_main, tpl_path, self.app.labels)
            export_data, extension, patient_id = exporter.generate_export(selected_pats, selected_visits, delete_empty)

            if not export_data:
                messagebox.showinfo("Info", "No data found for selected criteria.")
                return

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


class CVCExportDialog:
    """CVC data export configuration dialog."""

    def __init__(self, app):
        self.app = app

    def show(self):
        """Show configuration dialog for CVC Export."""
        if self.app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.app.root)
        win.title("Export CVC Data")
        win.geometry("500x600")
        win.configure(bg="#f4f4f4")

        # 1. Patient Selection Section
        pat_frame = tk.LabelFrame(win, text=" Patient Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(10, 5))

        filter_row = tk.Frame(pat_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 8))

        all_patients = sorted(self.app.df_main['Screening #'].dropna().unique())
        self._pat_vars = {}
        self._exclude_sf_var = tk.BooleanVar(value=False)

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

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self._exclude_sf_var,
                       bg="#f4f4f4", font=("Segoe UI", 9),
                       command=lambda: self._update_patient_list(scroll_frame, all_patients)).pack(side=tk.LEFT)

        # Populate patients
        for pid in all_patients:
            self._pat_vars[pid] = tk.BooleanVar(value=True)
            tk.Checkbutton(scroll_frame, text=str(pid), variable=self._pat_vars[pid],
                          bg="white", font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)

        # 2. Table Type Selection
        table_frame = tk.LabelFrame(win, text=" Table Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        table_frame.pack(fill=tk.X, padx=10, pady=5)

        self._screening_var = tk.BooleanVar(value=True)
        self._hemodynamic_var = tk.BooleanVar(value=True)

        tk.Checkbutton(table_frame, text="Screening Data", variable=self._screening_var,
                       bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Checkbutton(table_frame, text="Hemodynamic Assessment", variable=self._hemodynamic_var,
                       bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")

        # 3. Output Format
        fmt_frame = tk.LabelFrame(win, text=" Output Format ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        fmt_frame.pack(fill=tk.X, padx=10, pady=5)

        self._format_var = tk.StringVar(value="xlsx")
        tk.Radiobutton(fmt_frame, text="Excel (.xlsx)", variable=self._format_var,
                       value="xlsx", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")
        tk.Radiobutton(fmt_frame, text="CSV (.csv)", variable=self._format_var,
                       value="csv", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")

        # 4. Generate Button
        tk.Button(win, text="Generate CVC Export", command=lambda: self._generate(win),
                  bg="#2980b9", fg="white", font=("Segoe UI", 11, "bold"), pady=12, cursor="hand2").pack(fill=tk.X, padx=10, pady=15)

        win.focus_set()
        win.grab_set()

    def _update_patient_list(self, scroll_frame, all_patients):
        for widget in scroll_frame.winfo_children():
            widget.destroy()

        exclude_sf = self._exclude_sf_var.get()
        for pid in all_patients:
            if exclude_sf and self.app._is_screen_failure(pid):
                continue
            if pid not in self._pat_vars:
                self._pat_vars[pid] = tk.BooleanVar(value=True)
            tk.Checkbutton(scroll_frame, text=str(pid), variable=self._pat_vars[pid],
                          bg="white", font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)

    def _generate(self, dialog):
        exclude_sf = self._exclude_sf_var.get()
        selected_pats = []
        for pid, var in self._pat_vars.items():
            if var.get():
                if exclude_sf and self.app._is_screen_failure(pid):
                    continue
                selected_pats.append(pid)

        if not selected_pats:
            messagebox.showwarning("No Selection", "Please select at least one patient.")
            return

        include_screening = self._screening_var.get()
        include_hemodynamic = self._hemodynamic_var.get()

        if not include_screening and not include_hemodynamic:
            messagebox.showwarning("No Table", "Please select at least one table type.")
            return

        export_format = self._format_var.get()

        try:
            dialog.config(cursor="wait")
            dialog.update()

            exporter = CVCExporter(self.app.df_main)

            if export_format == "xlsx":
                if len(selected_pats) == 1:
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


class LabsExportDialog:
    """Labs data export configuration dialog."""

    def __init__(self, app):
        self.app = app

    def show(self):
        """Show configuration dialog for Labs Export."""
        if self.app.df_main is None:
            messagebox.showwarning("No Data", "Please load an Excel file first.")
            return

        win = tk.Toplevel(self.app.root)
        win.title("Export Labs Data")
        win.geometry("600x650")
        win.configure(bg="#f4f4f4")

        # 1. Template Section
        tpl_frame = tk.LabelFrame(win, text=" Template ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        tpl_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        inner_tpl = tk.Frame(tpl_frame, bg="#f4f4f4")
        inner_tpl.pack(fill=tk.X)

        self._tpl_path_var = tk.StringVar()
        default_tpl = os.path.join(os.getcwd(), "labs_tmpl.xlsx")
        if os.path.exists(default_tpl):
            self._tpl_path_var.set(default_tpl)

        tk.Entry(inner_tpl, textvariable=self._tpl_path_var, width=50, font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 5))
        tk.Button(inner_tpl, text="Browse...", command=self._browse_template,
                  bg="#2c3e50", fg="white", font=("Segoe UI", 9)).pack(side=tk.LEFT)

        # 2. Patient Filter Section
        pat_filter_frame = tk.LabelFrame(win, text=" Patient Selection ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_filter_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        filter_row = tk.Frame(pat_filter_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 8))

        self._exclude_sf_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures", variable=self._exclude_sf_var,
                       bg="#f4f4f4", font=("Segoe UI", 9), command=self._update_patient_list).pack(side=tk.LEFT)

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

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        self._pat_chk_vars = {}
        self._all_patients = sorted(self.app.df_main['Screening #'].dropna().unique())
        self._update_patient_list()

        # 3. Output Options Section
        opt_frame = tk.LabelFrame(win, text=" Output Options ", padx=10, pady=8, bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        opt_frame.pack(fill=tk.X, padx=10, pady=5)

        tk.Label(opt_frame, text="For days without data:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")

        self._delete_empty_cols_var = tk.BooleanVar(value=True)
        tk.Radiobutton(opt_frame, text="Delete column from output (recommended)", variable=self._delete_empty_cols_var,
                       value=True, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)
        tk.Radiobutton(opt_frame, text="Leave blank column in output", variable=self._delete_empty_cols_var,
                       value=False, bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", padx=10)

        tk.Frame(opt_frame, height=8, bg="#f4f4f4").pack(fill=tk.X)

        self._highlight_oor_var = tk.BooleanVar(value=False)
        tk.Checkbutton(opt_frame, text="Highlight out-of-range values in red",
                       variable=self._highlight_oor_var,
                       bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")

        # 4. Buttons: Preview and Generate
        btn_frame = tk.Frame(win, bg="#f4f4f4")
        btn_frame.pack(fill=tk.X, padx=10, pady=15)

        tk.Button(btn_frame, text="Preview (Single Patient)", command=lambda: self._preview(win),
                  bg="#3498db", fg="white", font=("Segoe UI", 10, "bold"), pady=10, cursor="hand2").pack(side=tk.LEFT, expand=True, fill=tk.X, padx=(0, 5))

        tk.Button(btn_frame, text="Generate Labs Reports", command=lambda: self._generate(win),
                  bg="#9b59b6", fg="white", font=("Segoe UI", 10, "bold"), pady=10, cursor="hand2").pack(side=tk.RIGHT, expand=True, fill=tk.X, padx=(5, 0))

        win.focus_set()
        win.grab_set()

    def _browse_template(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if path:
            self._tpl_path_var.set(path)

    def _update_patient_list(self):
        for widget in self._pat_scroll_frame.winfo_children():
            widget.destroy()

        exclude_sf = self._exclude_sf_var.get()

        for p in self._all_patients:
            if not p:
                continue
            if exclude_sf and self.app._is_screen_failure(p):
                continue

            if p in self._pat_chk_vars:
                var = self._pat_chk_vars[p]
            else:
                var = tk.BooleanVar(value=True)
                self._pat_chk_vars[p] = var

            tk.Checkbutton(self._pat_scroll_frame, text=p, variable=var,
                          bg="white", font=("Segoe UI", 9), anchor="w").pack(fill=tk.X, padx=5)

    def _select_all_patients(self, select):
        for widget in self._pat_scroll_frame.winfo_children():
            if isinstance(widget, tk.Checkbutton):
                pat_id = widget.cget("text")
                if pat_id in self._pat_chk_vars:
                    self._pat_chk_vars[pat_id].set(select)

    def _ask_unit_resolution(self, param_name, found_units, patient_id):
        """Callback to resolve unit conflicts."""
        resolved_unit = found_units[0] if found_units else None

        top = tk.Toplevel(self.app.root)
        top.title("Unit Conflict Detected")
        top.geometry("400x250")
        top.transient(self.app.root)
        top.grab_set()

        tk.Label(top, text="Unit Conflict Detected", font=("Segoe UI", 11, "bold"), fg="#e74c3c").pack(pady=10)

        msg = f"Parameter: {param_name}\nPatient: {patient_id}\n\nExisting units found: {', '.join(found_units)}"
        tk.Label(top, text=msg, justify=tk.LEFT, padx=20).pack(fill=tk.X)

        tk.Label(top, text="Select target unit (all values will be converted):", font=("Segoe UI", 9, "bold")).pack(pady=(15, 5))

        cmb = ttk.Combobox(top, values=found_units, state="readonly", width=20)
        cmb.set(found_units[0])
        cmb.pack(pady=5)

        result = [resolved_unit]

        def on_ok():
            result[0] = cmb.get()
            top.destroy()

        tk.Button(top, text="Convert / Proceed", command=on_ok, bg="#2c3e50", fg="white", width=20).pack(pady=20)

        try:
            x = self.app.root.winfo_x() + (self.app.root.winfo_width() // 2) - 200
            y = self.app.root.winfo_y() + (self.app.root.winfo_height() // 2) - 125
            top.geometry(f"+{x}+{y}")
        except Exception:
            pass

        top.wait_window()
        return result[0]

    def _preview(self, dialog):
        """Generate a temp file for a single selected patient and open in default viewer."""
        import tempfile
        import subprocess

        tpl_path = self._tpl_path_var.get()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showerror("Error", "Please select a valid template file.")
            return

        selected_pat = None
        for p in self._all_patients:
            if p in self._pat_chk_vars and self._pat_chk_vars[p].get():
                exclude_sf = self._exclude_sf_var.get()
                if exclude_sf and self.app._is_screen_failure(p):
                    continue
                selected_pat = p
                break

        if not selected_pat:
            messagebox.showwarning("Warning", "Please select at least one patient for preview.")
            return

        delete_empty = self._delete_empty_cols_var.get()
        highlight_oor = self._highlight_oor_var.get()

        try:
            dialog.config(cursor="watch")
            dialog.update()

            exporter = LabsExporter(self.app.df_main, tpl_path, self.app.labels,
                                    unit_callback=self._ask_unit_resolution,
                                    highlight_out_of_range=highlight_oor)
            export_data = exporter.process_patient(selected_pat, delete_empty)

            if not export_data:
                messagebox.showinfo("Info", f"No data found for patient {selected_pat}.")
                return

            temp_dir = tempfile.gettempdir()
            temp_path = os.path.join(temp_dir, f"{selected_pat}_labs_preview.xlsx")

            with open(temp_path, "wb") as f:
                f.write(export_data)

            try:
                os.startfile(temp_path)
            except AttributeError:
                subprocess.run(["open", temp_path])
            except Exception:
                subprocess.run(["xdg-open", temp_path])

            messagebox.showinfo("Preview", f"Preview opened for patient: {selected_pat}\n\nFile: {temp_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Preview failed: {e}")
        finally:
            try:
                dialog.config(cursor="")
            except Exception:
                pass

    def _generate(self, dialog):
        tpl_path = self._tpl_path_var.get()
        if not tpl_path or not os.path.exists(tpl_path):
            messagebox.showerror("Error", "Please select a valid template file.")
            return

        selected_pats = []
        for p in self._all_patients:
            if p in self._pat_chk_vars and self._pat_chk_vars[p].get():
                exclude_sf = self._exclude_sf_var.get()
                if exclude_sf and self.app._is_screen_failure(p):
                    continue
                selected_pats.append(p)

        delete_empty = self._delete_empty_cols_var.get()
        highlight_oor = self._highlight_oor_var.get()

        if not selected_pats:
            messagebox.showwarning("Warning", "No patients selected.")
            return

        try:
            dialog.config(cursor="watch")
            dialog.update()

            exporter = LabsExporter(self.app.df_main, tpl_path, self.app.labels,
                                    unit_callback=self._ask_unit_resolution,
                                    highlight_out_of_range=highlight_oor)
            export_data, extension, patient_id = exporter.generate_export(selected_pats, delete_empty)

            if not export_data:
                messagebox.showinfo("Info", "No data found for selected criteria.")
                return

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
            except Exception:
                pass


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
