"""Labs Export Dialog — configuration UI for laboratory data export."""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import os
import logging

from labs_export import LabsExporter

logger = logging.getLogger(__name__)


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
