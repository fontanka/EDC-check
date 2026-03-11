"""
HF Hospitalizations UI Module
==============================
Extracted from clinical_viewer1.py — provides the HF Hospitalizations
summary, detail, tuning, and export windows.
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import logging
from hf_hospitalization_manager import HFEvent

logger = logging.getLogger(__name__)


class HFWindow:
    """Manages all HF Hospitalizations UI windows.

    Args:
        app: ClinicalDataMasterV30 instance — provides hf_manager,
             df_main, root, get_screen_failures().
    """

    def __init__(self, app):
        self.app = app

    # ------------------------------------------------------------------
    # Summary window
    # ------------------------------------------------------------------
    def show(self):
        """Display the HF Hospitalizations summary window."""
        if self.app.hf_manager is None:
            messagebox.showwarning("Warning", "No data loaded. Please load an Excel file first.")
            return

        win = tk.Toplevel(self.app.root)
        win.title("HF Hospitalizations - Summary")
        win.geometry("900x600")
        win.transient(self.app.root)
        win.lift()
        win.focus_force()

        # Header
        header_frame = tk.Frame(win, bg="#e74c3c", pady=10)
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text="Heart Failure Hospitalizations",
                 bg="#e74c3c", fg="white", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT, padx=20)

        # Info label
        info_frame = tk.Frame(win, bg="#f4f4f4", pady=5)
        info_frame.pack(fill=tk.X)
        tk.Label(info_frame,
                 text="Pre-Treatment: 6 months before | Post-Treatment: 6 months after (based on AE symptom onset)",
                 bg="#f4f4f4", fg="#666", font=("Segoe UI", 9, "italic")).pack(side=tk.LEFT, padx=10)

        # Toolbar
        toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
        toolbar.pack(fill=tk.X)

        exclude_sf_var = tk.BooleanVar(value=True)

        tk.Button(toolbar, text="Refresh",
                  command=lambda: self._refresh_summary(tree, exclude_sf_var.get()),
                  bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=10)
        tk.Button(toolbar, text="Export Summary",
                  command=lambda: self._export_summary(exclude_sf_var.get()),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
        tk.Button(toolbar, text="Tuning Keywords",
                  command=self._show_tuning_dialog,
                  bg="#f39c12", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

        tk.Checkbutton(toolbar, text="Exclude Screen Failures",
                       variable=exclude_sf_var, bg="#f4f4f4",
                       command=lambda: self._refresh_summary(tree, exclude_sf_var.get())).pack(side=tk.LEFT, padx=20)

        # Summary Treeview
        tree_frame = tk.Frame(win)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        columns = ("patient", "treatment_date", "pre_count", "pre_count_1y", "post_count", "post_count_1y")
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")

        tree.heading("patient", text="Patient ID")
        tree.heading("treatment_date", text="Treatment Date")
        tree.heading("pre_count", text="Pre-6M")
        tree.heading("pre_count_1y", text="Pre-1Y")
        tree.heading("post_count", text="Post-6M")
        tree.heading("post_count_1y", text="Post-1Y")

        tree.column("patient", width=120, anchor="center")
        tree.column("treatment_date", width=120, anchor="center")
        tree.column("pre_count", width=100, anchor="center")
        tree.column("pre_count_1y", width=100, anchor="center")
        tree.column("post_count", width=100, anchor="center")
        tree.column("post_count_1y", width=100, anchor="center")

        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=v_scroll.set, xscrollcommand=h_scroll.set)

        tree.grid(row=0, column=0, sticky="nsew")
        v_scroll.grid(row=0, column=1, sticky="ns")
        h_scroll.grid(row=1, column=0, sticky="ew")

        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)

        tree.tag_configure('has_events', background='#fff3cd')
        tree.tag_configure('no_events', background='#ffffff')

        tree.bind("<Double-1>", lambda e: self._show_detail_for_selected(tree))

        # Store references for refresh
        self._summary_tree = tree
        self._exclude_sf_var = exclude_sf_var

        # Populate data
        self._refresh_summary(tree, exclude_sf_var.get())

        tk.Label(win, text="Double-click on a patient to view/edit event details",
                 fg="#888", font=("Segoe UI", 9, "italic")).pack(pady=5)

    # ------------------------------------------------------------------
    # Summary helpers
    # ------------------------------------------------------------------
    def _refresh_summary(self, tree, exclude_sf=True):
        """Refresh the HF summary tree with current data."""
        for item in tree.get_children():
            tree.delete(item)

        screen_failures = set()
        if exclude_sf:
            try:
                screen_failures = set(self.app.get_screen_failures())
            except Exception as e:
                logger.error("Error getting screen failures: %s", e)

        try:
            summaries = self.app.hf_manager.get_all_patients_summary()
        except Exception as e:
            logger.error("Error getting HF summaries: %s", e)
            return

        summaries.sort(key=lambda x: x['patient_id'])

        for summary in summaries:
            if exclude_sf and summary['patient_id'] in screen_failures:
                continue

            has_events = summary['pre_count'] > 0 or summary['post_count'] > 0
            tag = 'has_events' if has_events else 'no_events'

            tree.insert("", "end", iid=summary['patient_id'], values=(
                summary['patient_id'],
                summary['treatment_date'],
                summary['pre_count_6m'],
                summary['pre_count_1y'],
                summary['post_count_6m'],
                summary['post_count_1y']
            ), tags=(tag,))

    def _show_detail_for_selected(self, tree):
        """Show detail window for selected patient."""
        selection = tree.selection()
        if not selection:
            return
        self._show_detail_window(selection[0])

    # ------------------------------------------------------------------
    # Detail window
    # ------------------------------------------------------------------
    def _show_detail_window(self, patient_id):
        """Show detailed HF events for a patient with editable lists."""
        summary = self.app.hf_manager.get_patient_summary(patient_id)

        win = tk.Toplevel(self.app.root)
        win.title(f"HF Hospitalizations - Patient {patient_id}")
        win.geometry("1100x650")
        win.transient(self.app.root)

        # Header
        header_frame = tk.Frame(win, bg="#e74c3c", pady=10)
        header_frame.pack(fill=tk.X)
        tk.Label(header_frame, text=f"Patient: {patient_id}",
                 bg="#e74c3c", fg="white", font=("Segoe UI", 14, "bold")).pack(side=tk.LEFT, padx=20)
        tk.Label(header_frame, text=f"Treatment Date: {summary['treatment_date']}",
                 bg="#e74c3c", fg="white", font=("Segoe UI", 11)).pack(side=tk.LEFT, padx=20)

        # Summary stats
        stats_frame = tk.Frame(win, bg="#f4f4f4", pady=10)
        stats_frame.pack(fill=tk.X)
        tk.Label(stats_frame,
                 text=f"Pre-Treatment (1Y): {summary['pre_count_1y']} events (6M: {summary['pre_count_6m']})",
                 bg="#f4f4f4", fg="#c0392b", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=20)
        tk.Label(stats_frame,
                 text=f"Post-Treatment (1Y): {summary['post_count_1y']} events (6M: {summary['post_count_6m']})",
                 bg="#f4f4f4", fg="#27ae60", font=("Segoe UI", 11, "bold")).pack(side=tk.LEFT, padx=20)

        # Notebook for Pre/Post tabs
        notebook = ttk.Notebook(win)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        pre_frame = tk.Frame(notebook)
        notebook.add(pre_frame, text=f"Pre-Treatment ({summary['pre_count_1y']})")
        pre_tree = self._create_events_tree(pre_frame, summary['pre_events'], patient_id, "pre")

        post_frame = tk.Frame(notebook)
        notebook.add(post_frame, text=f"Post-Treatment ({summary['post_count_1y']})")
        post_tree = self._create_events_tree(post_frame, summary['post_events'], patient_id, "post")

        # Store references for save
        self._detail_patient = patient_id
        self._detail_pre_tree = pre_tree
        self._detail_post_tree = post_tree

        # Button frame
        btn_frame = tk.Frame(win, pady=10)
        btn_frame.pack(fill=tk.X)

        tk.Button(btn_frame, text="Save Changes", command=lambda: self._save_changes(win),
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.RIGHT, padx=20)
        tk.Button(btn_frame, text="Add Manual Event",
                  command=lambda: self._add_manual_event(patient_id, notebook),
                  bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)
        tk.Button(btn_frame, text="Export Details",
                  command=lambda: self._export_details(patient_id),
                  bg="#9b59b6", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.RIGHT, padx=5)

    # ------------------------------------------------------------------
    # Events tree (used within detail window)
    # ------------------------------------------------------------------
    def _create_events_tree(self, parent, events, patient_id, period):
        """Create a treeview for HF events with edit controls."""
        toolbar = tk.Frame(parent, pady=5)
        toolbar.pack(fill=tk.X)

        tk.Button(toolbar, text="Toggle Included/Excluded", command=lambda: toggle_include(tree),
                  bg="#e67e22", fg="white", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=10)

        columns = ("include", "date", "source", "term", "matched", "confidence", "type")
        tree = ttk.Treeview(parent, columns=columns, show="headings", height=12)

        tree.heading("include", text="Include")
        tree.heading("date", text="Date")
        tree.heading("source", text="Source")
        tree.heading("term", text="Original Term")
        tree.heading("matched", text="Matched Synonym")
        tree.heading("confidence", text="Conf.")
        tree.heading("type", text="Match Type")

        tree.column("include", width=60, anchor="center")
        tree.column("date", width=100, anchor="center")
        tree.column("source", width=60, anchor="center")
        tree.column("term", width=300, anchor="w")
        tree.column("matched", width=200, anchor="w")
        tree.column("confidence", width=60, anchor="center")
        tree.column("type", width=80, anchor="center")

        tree.tag_configure('included', background='#d4edda')
        tree.tag_configure('excluded', background='#f8d7da', foreground='#888')
        tree.tag_configure('manual', background='#cce5ff')

        for event in events:
            include_text = "\u2713" if event.is_included else "\u2717"
            tag = 'manual' if event.is_manual else ('included' if event.is_included else 'excluded')

            tree.insert("", "end", iid=event.event_id, values=(
                include_text,
                event.date,
                event.source_form,
                event.original_term[:50] + "..." if len(event.original_term) > 50 else event.original_term,
                event.matched_synonym,
                f"{event.confidence:.0%}",
                event.match_type
            ), tags=(tag,))

        tree_frame = tk.Frame(parent)
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)

        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=v_scroll.set)

        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        v_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        def toggle_include(event_widget=tree):
            selection = event_widget.selection()
            for item_id in selection:
                current = event_widget.item(item_id, "values")
                new_include = "\u2717" if current[0] == "\u2713" else "\u2713"
                new_tag = 'included' if new_include == "\u2713" else 'excluded'
                event_widget.item(item_id, values=(new_include,) + current[1:], tags=(new_tag,))

        tree.bind("<Double-1>", lambda e: toggle_include())

        tree.events = events
        tree.period = period

        tk.Label(parent, text="Double-click to toggle Include/Exclude",
                 fg="#888", font=("Segoe UI", 8, "italic")).pack()

        return tree

    # ------------------------------------------------------------------
    # Save / manual add / export
    # ------------------------------------------------------------------
    def _save_changes(self, win):
        """Save changes made in the detail window."""
        patient_id = self._detail_patient

        for tree in [self._detail_pre_tree, self._detail_post_tree]:
            for item_id in tree.get_children():
                values = tree.item(item_id, "values")
                is_included = values[0] == "\u2713"

                original_event = None
                for e in tree.events:
                    if e.event_id == item_id:
                        original_event = e
                        break

                if original_event and original_event.is_included != is_included:
                    updated_event = HFEvent(
                        event_id=original_event.event_id,
                        date=original_event.date,
                        source_form=original_event.source_form,
                        source_row=original_event.source_row,
                        original_term=original_event.original_term,
                        matched_synonym=original_event.matched_synonym,
                        match_type=original_event.match_type,
                        confidence=original_event.confidence,
                        is_included=is_included,
                        is_manual=original_event.is_manual,
                        notes=original_event.notes
                    )
                    self.app.hf_manager.update_event(patient_id, updated_event)

        messagebox.showinfo("Saved", "Changes saved successfully.")

        if hasattr(self, '_summary_tree'):
            self._refresh_summary(self._summary_tree)

    def _add_manual_event(self, patient_id, notebook):
        """Add a manual HF event."""
        tab_index = notebook.index(notebook.select())
        is_pre = tab_index == 0

        dialog = tk.Toplevel(self.app.root)
        dialog.title("Add Manual HF Event")
        dialog.geometry("400x250")
        dialog.transient(self.app.root)
        dialog.grab_set()

        tk.Label(dialog, text="Date (YYYY-MM-DD):", font=("Segoe UI", 10)).grid(
            row=0, column=0, padx=10, pady=10, sticky="e")
        date_entry = tk.Entry(dialog, width=20)
        date_entry.grid(row=0, column=1, padx=10, pady=10)

        tk.Label(dialog, text="Description:", font=("Segoe UI", 10)).grid(
            row=1, column=0, padx=10, pady=10, sticky="e")
        desc_entry = tk.Entry(dialog, width=30)
        desc_entry.grid(row=1, column=1, padx=10, pady=10)

        tk.Label(dialog, text="Period:", font=("Segoe UI", 10)).grid(
            row=2, column=0, padx=10, pady=10, sticky="e")
        period_var = tk.StringVar(value="pre" if is_pre else "post")
        tk.Radiobutton(dialog, text="Pre-Treatment", variable=period_var, value="pre").grid(row=2, column=1, sticky="w")
        tk.Radiobutton(dialog, text="Post-Treatment", variable=period_var, value="post").grid(row=3, column=1, sticky="w")

        def save_manual():
            date_str = date_entry.get().strip()
            desc = desc_entry.get().strip()
            period = period_var.get()

            if not date_str or not desc:
                messagebox.showwarning("Warning", "Please fill in all fields.")
                return

            event_id = f"MANUAL_{patient_id}_{len(self.app.hf_manager.manual_edits.get(patient_id, []))}"
            source = f"MANUAL_{'PRE' if period == 'pre' else 'POST'}"

            event = HFEvent(
                event_id=event_id,
                date=date_str,
                source_form=source,
                source_row=0,
                original_term=desc,
                matched_synonym="Manual Entry",
                match_type="manual",
                confidence=1.0,
                is_included=True,
                is_manual=True
            )

            self.app.hf_manager.update_event(patient_id, event)
            dialog.destroy()
            messagebox.showinfo("Added", "Manual event added. Please refresh the detail view.")

        tk.Button(dialog, text="Add Event", command=save_manual,
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).grid(
            row=4, column=0, columnspan=2, pady=20)

    def _export_summary(self, exclude_sf=True):
        """Export HF summary to Excel."""
        if self.app.hf_manager is None:
            return

        try:
            summaries = self.app.hf_manager.get_all_patients_summary()

            if exclude_sf:
                screen_failures = set(self.app.get_screen_failures())
                summaries = [s for s in summaries if s['patient_id'] not in screen_failures]

            df = pd.DataFrame([{
                'Patient ID': s['patient_id'],
                'Treatment Date': s['treatment_date'],
                'Pre-6M HF Hosps': s.get('pre_count_6m', s['pre_count']),
                'Pre-1Y HF Hosps': s.get('pre_count_1y', 0),
                'Post-6M HF Hosps': s.get('post_count_6m', s['post_count']),
                'Post-1Y HF Hosps': s.get('post_count_1y', 0)
            } for s in summaries])

            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                initialfile="hf_hospitalizations_summary.xlsx"
            )

            if path:
                if path.endswith('.csv'):
                    df.to_csv(path, index=False)
                else:
                    df.to_excel(path, index=False)
                messagebox.showinfo("Exported", f"Summary exported to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    def _export_details(self, patient_id):
        """Export detailed HF events for a patient."""
        try:
            summary = self.app.hf_manager.get_patient_summary(patient_id)

            all_events = []
            for event in summary['pre_events']:
                all_events.append({
                    'Patient ID': patient_id,
                    'Period': 'Pre-Treatment',
                    'Date': event.date,
                    'Source': event.source_form,
                    'Term': event.original_term,
                    'Matched': event.matched_synonym,
                    'Confidence': f"{event.confidence:.0%}",
                    'Type': event.match_type,
                    'Included': 'Yes' if event.is_included else 'No',
                    'Manual': 'Yes' if event.is_manual else 'No'
                })
            for event in summary['post_events']:
                all_events.append({
                    'Patient ID': patient_id,
                    'Period': 'Post-Treatment',
                    'Date': event.date,
                    'Source': event.source_form,
                    'Term': event.original_term,
                    'Matched': event.matched_synonym,
                    'Confidence': f"{event.confidence:.0%}",
                    'Type': event.match_type,
                    'Included': 'Yes' if event.is_included else 'No',
                    'Manual': 'Yes' if event.is_manual else 'No'
                })

            df = pd.DataFrame(all_events)

            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
                initialfile=f"hf_details_{patient_id}.xlsx"
            )

            if path:
                if path.endswith('.csv'):
                    df.to_csv(path, index=False)
                else:
                    df.to_excel(path, index=False)
                messagebox.showinfo("Exported", f"Details exported to {path}")
        except Exception as e:
            messagebox.showerror("Error", f"Export failed: {e}")

    # ------------------------------------------------------------------
    # Tuning dialog
    # ------------------------------------------------------------------
    def _show_tuning_dialog(self):
        """Management dialog for HF tuning keywords (Include/Exclude)."""
        if self.app.hf_manager is None:
            return

        dialog = tk.Toplevel(self.app.root)
        dialog.title("HF Tuning Keywords")
        dialog.geometry("600x550")
        dialog.transient(self.app.root)

        main_frame = tk.Frame(dialog, padx=10, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="Global Inclusion/Exclusion Keywords",
                 font=("Segoe UI", 12, "bold")).pack(pady=5)
        tk.Label(main_frame, text="These keywords affect autodetected events globally across all patients.",
                 fg="#666", font=("Segoe UI", 9, "italic")).pack()

        lists_frame = tk.Frame(main_frame)
        lists_frame.pack(fill=tk.BOTH, expand=True, pady=10)

        # Include List
        inc_frame = tk.LabelFrame(lists_frame, text="Include Keywords (Hard Match)", padx=5, pady=5)
        inc_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        inc_list = tk.Listbox(inc_frame, height=10)
        inc_list.pack(fill=tk.BOTH, expand=True)
        for kw in self.app.hf_manager.custom_includes:
            inc_list.insert(tk.END, kw)

        inc_ctrl = tk.Frame(inc_frame)
        inc_ctrl.pack(fill=tk.X, pady=5)
        inc_entry = tk.Entry(inc_ctrl)
        inc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def add_include():
            kw = inc_entry.get().strip().lower()
            if kw and kw not in self.app.hf_manager.custom_includes:
                self.app.hf_manager.custom_includes.append(kw)
                inc_list.insert(tk.END, kw)
                inc_entry.delete(0, tk.END)

        def del_include():
            sel = inc_list.curselection()
            if sel:
                kw = inc_list.get(sel[0])
                self.app.hf_manager.custom_includes.remove(kw)
                inc_list.delete(sel[0])

        tk.Button(inc_ctrl, text="+", command=add_include, width=3).pack(side=tk.LEFT, padx=2)
        tk.Button(inc_ctrl, text="-", command=del_include, width=3).pack(side=tk.LEFT, padx=2)

        # Exclude List
        excl_frame = tk.LabelFrame(lists_frame, text="Exclude Keywords (Ignore Match)", padx=5, pady=5)
        excl_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=5)

        excl_list = tk.Listbox(excl_frame, height=10)
        excl_list.pack(fill=tk.BOTH, expand=True)
        for kw in self.app.hf_manager.custom_excludes:
            excl_list.insert(tk.END, kw)

        excl_ctrl = tk.Frame(excl_frame)
        excl_ctrl.pack(fill=tk.X, pady=5)
        excl_entry = tk.Entry(excl_ctrl)
        excl_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)

        def add_exclude():
            kw = excl_entry.get().strip().lower()
            if kw and kw not in self.app.hf_manager.custom_excludes:
                self.app.hf_manager.custom_excludes.append(kw)
                excl_list.insert(tk.END, kw)
                excl_entry.delete(0, tk.END)

        def del_exclude():
            sel = excl_list.curselection()
            if sel:
                kw = excl_list.get(sel[0])
                self.app.hf_manager.custom_excludes.remove(kw)
                excl_list.delete(sel[0])

        tk.Button(excl_ctrl, text="+", command=add_exclude, width=3).pack(side=tk.LEFT, padx=2)
        tk.Button(excl_ctrl, text="-", command=del_exclude, width=3).pack(side=tk.LEFT, padx=2)

        def save_and_close():
            self.app.hf_manager.save_tuning_config()
            dialog.destroy()
            messagebox.showinfo("Saved", "Tuning keywords saved. Please refresh summary to apply.")

        tk.Button(main_frame, text="Save Global Keywords", command=save_and_close,
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).pack(pady=10)
