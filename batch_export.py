"""
Batch Export Module
===================
Orchestrates exporting Labs + Echo + CVC + FU Highlights + Procedure Timing
for multiple patients in one operation. Produces a ZIP file with all outputs.
"""

import os
import logging
import zipfile
from io import BytesIO
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

logger = logging.getLogger("BatchExport")


class BatchExportDialog:
    """Dialog for configuring and running batch exports."""

    EXPORT_TYPES = {
        "labs": "Labs Data",
        "echo": "Echo Data (Sponsor)",
        "cvc": "CVC (Catheterization)",
        "fu_highlights": "FU Highlights",
        "procedure_timing": "Procedure Timing",
    }

    def __init__(self, parent, app):
        """
        Args:
            parent: Parent tk widget
            app: ClinicalDataMasterV30 instance (provides df_main, labels, etc.)
        """
        self.app = app
        self.parent = parent
        self._build_dialog()

    def _build_dialog(self):
        self.win = tk.Toplevel(self.parent)
        self.win.title("Batch Export — All Reports")
        self.win.geometry("550x600")
        self.win.configure(bg="#f4f4f4")
        self.win.transient(self.parent)
        self.win.lift()
        self.win.focus_force()

        # --- 1. Export Type Selection ---
        type_frame = tk.LabelFrame(self.win, text=" Export Types ", padx=10, pady=8,
                                   bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        type_frame.pack(fill=tk.X, padx=10, pady=(10, 5))

        self.type_vars = {}
        for key, label in self.EXPORT_TYPES.items():
            var = tk.BooleanVar(value=True)
            self.type_vars[key] = var
            tk.Checkbutton(type_frame, text=label, variable=var,
                           bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w")

        # --- 2. Patient Selection ---
        pat_frame = tk.LabelFrame(self.win, text=" Patient Selection ", padx=10, pady=8,
                                  bg="#f4f4f4", font=("Segoe UI", 10, "bold"))
        pat_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Filter row
        filter_row = tk.Frame(pat_frame, bg="#f4f4f4")
        filter_row.pack(fill=tk.X, pady=(0, 8))

        self.exclude_sf_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_row, text="Exclude Screen Failures",
                       variable=self.exclude_sf_var, bg="#f4f4f4",
                       font=("Segoe UI", 9),
                       command=self._refresh_patients).pack(side=tk.LEFT)

        btn_row = tk.Frame(filter_row, bg="#f4f4f4")
        btn_row.pack(side=tk.RIGHT)
        tk.Button(btn_row, text="Select All",
                  command=lambda: self._set_all_patients(True),
                  bg="#27ae60", fg="white", font=("Segoe UI", 8, "bold"),
                  padx=8).pack(side=tk.LEFT, padx=2)
        tk.Button(btn_row, text="Select None",
                  command=lambda: self._set_all_patients(False),
                  bg="#e74c3c", fg="white", font=("Segoe UI", 8, "bold"),
                  padx=8).pack(side=tk.LEFT, padx=2)

        # Scrollable patient list
        list_frame = tk.Frame(pat_frame, bd=1, relief="sunken", bg="white")
        list_frame.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(list_frame, bg="white", highlightthickness=0)
        sb = tk.Scrollbar(list_frame, orient="vertical", command=canvas.yview)
        self._scroll_frame = tk.Frame(canvas, bg="white")
        self._scroll_frame.bind("<Configure>",
                                lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=self._scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=sb.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Scoped mousewheel: only scroll when mouse is over the canvas
        def _on_mousewheel(e):
            try:
                canvas.yview_scroll(int(-1 * (e.delta / 120)), "units")
            except tk.TclError:
                pass

        def _bind_wheel(e):
            canvas.bind_all("<MouseWheel>", _on_mousewheel)

        def _unbind_wheel(e):
            canvas.unbind_all("<MouseWheel>")

        canvas.bind("<Enter>", _bind_wheel)
        canvas.bind("<Leave>", _unbind_wheel)
        self.win.bind("<Destroy>", lambda e: canvas.unbind_all("<MouseWheel>"))

        self.pat_vars = {}
        self._all_patients = sorted(
            self.app.df_main['Screening #'].dropna().unique()
        )
        self._refresh_patients()

        # --- 3. Export Button ---
        btn_frame = tk.Frame(self.win, bg="#f4f4f4")
        btn_frame.pack(fill=tk.X, padx=10, pady=15)

        self.status_var = tk.StringVar(value="Select export types and patients, then click Export.")
        tk.Label(btn_frame, textvariable=self.status_var, bg="#f4f4f4",
                 font=("Segoe UI", 9, "italic"), fg="#555").pack(side=tk.LEFT)

        tk.Button(btn_frame, text="Export All ►", command=self._run_export,
                  bg="#2c3e50", fg="white", font=("Segoe UI", 11, "bold"),
                  padx=20, pady=6).pack(side=tk.RIGHT)

    # ------------------------------------------------------------------
    # Patient list helpers
    # ------------------------------------------------------------------
    def _refresh_patients(self):
        for w in self._scroll_frame.winfo_children():
            w.destroy()
        self.pat_vars.clear()

        for pid in self._all_patients:
            if self.exclude_sf_var.get():
                mask = self.app.df_main['Screening #'] == pid
                row = self.app.df_main[mask]
                if not row.empty and 'SBV_ELIG_IEORRES_CONF5' in row.columns:
                    val = str(row.iloc[0].get('SBV_ELIG_IEORRES_CONF5', '')).strip().lower()
                    if val in ('screen failure', 'not eligible', 'no'):
                        continue

            var = tk.BooleanVar(value=True)
            self.pat_vars[pid] = var
            tk.Checkbutton(self._scroll_frame, text=str(pid), variable=var,
                           bg="white", font=("Segoe UI", 9), anchor="w").pack(
                fill=tk.X, padx=5)

    def _set_all_patients(self, state: bool):
        for var in self.pat_vars.values():
            var.set(state)

    # ------------------------------------------------------------------
    # Export orchestration
    # ------------------------------------------------------------------
    def _run_export(self):
        selected_types = [k for k, v in self.type_vars.items() if v.get()]
        selected_patients = [p for p, v in self.pat_vars.items() if v.get()]

        if not selected_types:
            messagebox.showwarning("No Types", "Please select at least one export type.")
            return
        if not selected_patients:
            messagebox.showwarning("No Patients", "Please select at least one patient.")
            return

        # Ask for output directory
        out_dir = filedialog.askdirectory(title="Select Output Folder")
        if not out_dir:
            return

        self.win.config(cursor="wait")
        self.status_var.set(f"Exporting {len(selected_types)} types × {len(selected_patients)} patients...")
        self.win.update_idletasks()

        try:
            results = self._generate_all(selected_types, selected_patients)
            # Write ZIP
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            zip_name = f"batch_export_{timestamp}.zip"
            zip_path = os.path.join(out_dir, zip_name)

            with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zf:
                for folder, filename, data in results:
                    arc_path = f"{folder}/{filename}" if folder else filename
                    zf.writestr(arc_path, data)

            n_files = len(results)
            self.status_var.set(f"✓ Exported {n_files} files → {zip_name}")
            messagebox.showinfo("Export Complete",
                                f"Batch export saved to:\n{zip_path}\n\n{n_files} files generated.")
            logger.info(f"Batch export: {n_files} files → {zip_path}")

        except Exception as e:
            logger.error(f"Batch export failed: {e}")
            messagebox.showerror("Export Error", f"Export failed:\n{e}")
            self.status_var.set(f"✗ Export failed: {e}")
        finally:
            try:
                self.win.config(cursor="")
            except Exception:
                pass

    def _generate_all(self, types, patients):
        """Generate all selected exports. Returns list of (folder, filename, bytes)."""
        results = []

        if "labs" in types:
            results.extend(self._export_labs(patients))
        if "echo" in types:
            results.extend(self._export_echo(patients))
        if "cvc" in types:
            results.extend(self._export_cvc(patients))
        if "fu_highlights" in types:
            results.extend(self._export_fu_highlights(patients))
        if "procedure_timing" in types:
            results.extend(self._export_procedure_timing(patients))

        return results

    def _export_labs(self, patients):
        """Export labs for each patient."""
        results = []
        template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "labs_tmpl.xlsx")
        if not os.path.exists(template):
            logger.warning("Labs template not found, skipping labs export")
            return results

        try:
            from labs_export import LabsExporter
            exporter = LabsExporter(self.app.df_main, template, self.app.labels)
            data, ext = exporter.generate_export(patients, delete_empty_cols=True)
            if data:
                results.append(("Labs", f"labs_export.{ext}", data))
        except Exception as e:
            logger.error(f"Labs export failed: {e}")
        return results

    def _export_echo(self, patients):
        """Export echo data for each patient."""
        results = []
        template = os.path.join(os.path.dirname(os.path.abspath(__file__)), "echo_tmpl.xlsx")
        if not os.path.exists(template):
            logger.warning("Echo template not found, skipping echo export")
            return results

        try:
            from echo_export import EchoExporter, VISIT_ORDER
            exporter = EchoExporter(self.app.df_main, template, self.app.labels)
            data, ext = exporter.generate_export(patients, VISIT_ORDER, delete_empty_rows=True)
            if data:
                results.append(("Echo", f"echo_export.{ext}", data))
        except Exception as e:
            logger.error(f"Echo export failed: {e}")
        return results

    def _export_cvc(self, patients):
        """Export CVC data for each patient."""
        results = []
        try:
            from cvc_export import CVCExporter
            exporter = CVCExporter(self.app.df_main)
            for pid in patients:
                try:
                    data = exporter.export_to_excel(pid)
                    if data:
                        safe_pid = str(pid).replace("/", "-")
                        results.append(("CVC", f"cvc_{safe_pid}.xlsx", data.getvalue()))
                except Exception as e:
                    logger.warning(f"CVC export for {pid} failed: {e}")
        except Exception as e:
            logger.error(f"CVC export init failed: {e}")
        return results

    def _export_fu_highlights(self, patients):
        """Export FU highlights for each patient."""
        results = []
        try:
            from fu_highlights_export import FUHighlightsExporter
            exporter = FUHighlightsExporter(self.app.df_main)
            for pid in patients:
                try:
                    visits = exporter.get_available_fu_visits(pid)
                    for visit in visits:
                        highlights, vitals = exporter.generate_highlights_table(pid, visit)
                        if highlights:
                            safe_pid = str(pid).replace("/", "-")
                            content = f"=== FU Highlights: {pid} — {visit} ===\n\n"
                            content += highlights + "\n\n" + vitals
                            results.append(("FU_Highlights",
                                            f"fu_{safe_pid}_{visit}.txt",
                                            content.encode('utf-8')))
                except Exception as e:
                    logger.warning(f"FU export for {pid} failed: {e}")
        except Exception as e:
            logger.error(f"FU Highlights export init failed: {e}")
        return results

    def _export_procedure_timing(self, patients):
        """Export procedure timing matrix."""
        results = []
        try:
            from procedure_timing_export import ProcedureTimingExporter
            exporter = ProcedureTimingExporter(self.app.df_main, self.app.labels)
            data = exporter.generate_export(patients)
            if data:
                results.append(("Procedure_Timing", "procedure_timing.xlsx", data))
        except Exception as e:
            logger.error(f"Procedure timing export failed: {e}")
        return results
