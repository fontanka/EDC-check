"""Echo Export Dialog — configuration UI for echocardiography data export."""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import logging

from echo_export import EchoExporter, VISIT_ORDER as ECHO_VISITS

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
