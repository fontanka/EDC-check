"""CVC Export Dialog — configuration UI for cardiac catheterization data export."""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import logging

from cvc_export import CVCExporter

logger = logging.getLogger(__name__)


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
