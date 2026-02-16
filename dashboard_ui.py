import logging
import threading
import datetime
import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd

logger = logging.getLogger("ClinicalViewer.DashboardUI")

class DashboardWindow:
    def __init__(self, root, dashboard_manager, get_screen_failures_callback=None):
        self.root = root
        self.mgr = dashboard_manager
        self.get_sf_callback = get_screen_failures_callback
        
        self.window = tk.Toplevel(root)
        self.window.title("SDV & Data Gap Dashboard")
        self.window.geometry("1200x800")
        
        self._destroyed = False
        self.window.bind("<Destroy>", self._on_window_destroy)
        
        # Style Configuration
        style = ttk.Style(self.window)
        try:
             style.theme_use('clam')
        except:
             pass
        
        style.configure("Treeview", rowheight=25, font=('Segoe UI', 10))
        style.configure("Treeview.Heading", font=('Segoe UI', 10, 'bold'))
        style.map("Treeview", background=[('selected', '#0078D7')])
        
        # Toolbar
        toolbar = tk.Frame(self.window, pady=5, padx=5)

        toolbar.pack(side=tk.TOP, fill=tk.X)
        
        self.exclude_sf_var = tk.BooleanVar(value=False)
        self.manual_sf_overrides = set()  # Store manually added screen failure patient IDs
        
        if self.get_sf_callback:
            chk = tk.Checkbutton(toolbar, text="Exclude Screen Failures", 
                                 variable=self.exclude_sf_var, command=self._calculate_and_load)
            chk.pack(side=tk.LEFT)
            
            # Button to manage manual overrides
            btn_override = tk.Button(toolbar, text="Manage SF Overrides", command=self._open_sf_override_dialog)
            btn_override.pack(side=tk.LEFT, padx=5)


        
        # Calculate stats first
        self.status_var = tk.StringVar(value="Calculating statistics...")
        self.status_label = tk.Label(self.window, textvariable=self.status_var, fg="blue")
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)
        
        # Run calc in background ideally, but for now synchronous
        self.window.after(100, self._calculate_and_load)
        
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Create Tabs
        self.tab_study = ttk.Frame(self.notebook)
        self.tab_site = ttk.Frame(self.notebook)
        self.tab_patient = ttk.Frame(self.notebook)
        self.tab_form = ttk.Frame(self.notebook)
        self.tab_performance = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_study, text="Study Level")
        self.notebook.add(self.tab_site, text="Site Level")
        self.notebook.add(self.tab_patient, text="Patient Level")
        self.notebook.add(self.tab_form, text="Form Level")
        self.notebook.add(self.tab_performance, text="CRA Performance")
        
        self.trees = {}

    def _calculate_and_load(self):
        self.status_var.set("Calculating statistics... (this may take a moment)")
        self.window.update_idletasks()
        
        # Disable checkbox during calc
        # (Optional: disable toolbar items)
        
        excluded = []
        if self.exclude_sf_var.get() and self.get_sf_callback:
            auto_sf = self.get_sf_callback()
            # Combine auto-detected and manual overrides
            excluded = list(set(auto_sf) | self.manual_sf_overrides)
            
        # Run in thread
        threading.Thread(target=self._calculate_thread, args=(excluded,), daemon=True).start()

    def _calculate_thread(self, excluded):
        try:
            self.mgr.calculate_stats(excluded_patients=excluded)
            # Schedule UI update on main thread (only if window still exists)
            if not self._destroyed:
                try:
                    self.window.after(0, self._on_stats_ready)
                except (tk.TclError, RuntimeError):
                    pass  # Window destroyed between check and call
        except Exception as e:
            err_msg = str(e)  # Capture before lambda
            if not self._destroyed:
                try:
                    self.window.after(0, lambda msg=err_msg: self._on_calc_error(msg))
                except (tk.TclError, RuntimeError):
                    pass  # Window destroyed between check and call

    def _on_window_destroy(self, event):
        """Mark window as destroyed to prevent thread callbacks on dead widget."""
        if event.widget is self.window:
            self._destroyed = True

    def _on_stats_ready(self):
        self.status_var.set("Statistics ready.")
        self._build_study_tab()
        self._build_site_tab()
        self._build_patient_tab()
        self._build_form_tab()
        self._build_performance_tab()
        
    def _on_calc_error(self, error_msg):
        self.status_var.set(f"Error: {error_msg}")
        messagebox.showerror("Error", f"Failed to calculate stats: {error_msg}")
    
    def _open_sf_override_dialog(self):
        """Open dialog to manage manual screen failure overrides."""
        dialog = tk.Toplevel(self.window)
        dialog.title("Manage Screen Failure Overrides")
        dialog.geometry("400x350")
        dialog.transient(self.window)
        dialog.grab_set()
        
        # Instructions
        tk.Label(dialog, text="Add patient IDs to manually mark as Screen Failures:",
                 wraplength=380, justify=tk.LEFT).pack(pady=10, padx=10)
        
        # Current overrides list
        list_frame = tk.Frame(dialog)
        list_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        tk.Label(list_frame, text="Current Manual Overrides:").pack(anchor=tk.W)
        
        listbox = tk.Listbox(list_frame, selectmode=tk.SINGLE, height=8)
        listbox.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)
        
        scrollbar = tk.Scrollbar(list_frame, orient=tk.VERTICAL, command=listbox.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.configure(yscrollcommand=scrollbar.set)
        
        # Populate listbox
        for pid in sorted(self.manual_sf_overrides):
            listbox.insert(tk.END, pid)
        
        # Add/Remove controls
        ctrl_frame = tk.Frame(dialog)
        ctrl_frame.pack(fill=tk.X, padx=10, pady=5)
        
        tk.Label(ctrl_frame, text="Patient ID:").pack(side=tk.LEFT)
        entry_var = tk.StringVar()
        entry = tk.Entry(ctrl_frame, textvariable=entry_var, width=15)
        entry.pack(side=tk.LEFT, padx=5)
        
        def add_patient():
            pid = entry_var.get().strip()
            if pid and pid not in self.manual_sf_overrides:
                self.manual_sf_overrides.add(pid)
                listbox.insert(tk.END, pid)
                entry_var.set("")
        
        def remove_selected():
            selection = listbox.curselection()
            if selection:
                pid = listbox.get(selection[0])
                self.manual_sf_overrides.discard(pid)
                listbox.delete(selection[0])
        
        tk.Button(ctrl_frame, text="Add", command=add_patient).pack(side=tk.LEFT, padx=2)
        tk.Button(ctrl_frame, text="Remove Selected", command=remove_selected).pack(side=tk.LEFT, padx=2)
        
        # Close button
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=10, pady=10)
        
        def close_and_refresh():
            dialog.destroy()
            # Recalculate if exclusion is enabled
            if self.exclude_sf_var.get():
                self._calculate_and_load()
        
        tk.Button(btn_frame, text="Close & Apply", command=close_and_refresh).pack(side=tk.RIGHT)

    def _create_tree(self, parent, columns, show_cols, key):
        if key in self.trees:
            tree = self.trees[key]
            tree.delete(*tree.get_children())
            return tree
            
        tree = ttk.Treeview(parent, columns=columns, show='headings')
        for col in columns:
            tree.heading(col, text=col, command=lambda c=col: self._sort_tree(tree, c, False))
            width = 100 if col != "Name" else 200
            tree.column(col, width=width, anchor='center' if col != "Name" else 'w')
        
        scrollbar = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Bind double click or single click to drill down?
        # User asked for hyperlinks. In Treeview, we simulate this by clicking the cell.
        tree.bind('<ButtonRelease-1>', lambda event: self._on_tree_click(event, tree))
        
        self.trees[key] = tree
        return tree

    def _build_study_tab(self):
        cols = ("Metric", "Count", "Percentage")
        tree = self._create_tree(self.tab_study, cols, 'headings', 'study')
        
        stats = self.mgr.get_summary('study')
        total = sum(stats.values())
        
        for metric in ['V', '!', 'NS', 'GAP']:
            count = stats.get(metric, 0)
            pct = (count / total * 100) if total > 0 else 0
            tree.insert("", "end", values=(metric, count, f"{pct:.1f}%"), tags=(metric,))

    def _build_site_tab(self):
        cols = ("Site", "V", "!", "NS", "GAP", "Total")
        tree = self._create_tree(self.tab_site, cols, 'headings', 'site')
        
        sites = self.mgr.get_summary('site') # returns dict of site_id -> stats
        
        for site_id, metrics in sites.items():
            total = sum(metrics.values())
            vals = [site_id]
            for m in ['V', '!', 'NS', 'GAP']:
                c = metrics.get(m, 0)
                p = (c / total * 100) if total > 0 else 0
                vals.append(f"{c} ({p:.1f}%)")
            vals.append(total)
            tree.insert("", "end", values=vals, tags=(site_id,))

    def _build_patient_tab(self):
        cols = ("Patient", "V", "!", "NS", "GAP", "Total")
        tree = self._create_tree(self.tab_patient, cols, 'headings', 'patient')
        
        patients = self.mgr.get_summary('patient')
        for pat_id, metrics in patients.items():
            total = sum(metrics.values())
            vals = [pat_id]
            for m in ['V', '!', 'NS', 'GAP']:
                c = metrics.get(m, 0)
                p = (c / total * 100) if total > 0 else 0
                vals.append(f"{c} ({p:.1f}%)")
            vals.append(total)
            tree.insert("", "end", values=vals, tags=(pat_id,))

    def _build_form_tab(self):
        cols = ("Patient", "Form", "V", "!", "NS", "GAP", "Total")
        tree = self._create_tree(self.tab_form, cols, 'headings', 'form')
        
        forms = self.mgr.get_summary('form')
        for (pat_id, form_name), metrics in forms.items():
            total = sum(metrics.values())
            vals = [pat_id, form_name]
            for m in ['V', '!', 'NS', 'GAP']:
                c = metrics.get(m, 0)
                p = (c / total * 100) if total > 0 else 0
                vals.append(f"{c} ({p:.1f}%)")
            vals.append(total)
            tree.insert("", "end", values=vals, tags=((pat_id, form_name),))

    def _build_performance_tab(self):
        # Container for controls and tree
        main_frame = tk.Frame(self.tab_performance)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Filter Controls
        ctrl_frame = tk.LabelFrame(main_frame, text="Monitoring Visit Period (History Data)", padx=10, pady=10)
        ctrl_frame.pack(side=tk.TOP, fill=tk.X, pady=5)
        
        default_end = datetime.datetime.now().strftime('%Y-%m-%d')
        default_start = (datetime.datetime.now() - datetime.timedelta(days=7)).strftime('%Y-%m-%d')
        
        tk.Label(ctrl_frame, text="Start (YYYY-MM-DD):").pack(side=tk.LEFT)
        self.perf_start_var = tk.StringVar(value=default_start)
        tk.Entry(ctrl_frame, textvariable=self.perf_start_var, width=12).pack(side=tk.LEFT, padx=5)
        
        tk.Label(ctrl_frame, text="End (YYYY-MM-DD):").pack(side=tk.LEFT)
        self.perf_end_var = tk.StringVar(value=default_end)
        tk.Entry(ctrl_frame, textvariable=self.perf_end_var, width=12).pack(side=tk.LEFT, padx=5)
        
        tk.Label(ctrl_frame, text="CRA:").pack(side=tk.LEFT, padx=(10, 0))
        self.perf_user_var = tk.StringVar(value="All")
        
        # Extract unique users from history
        users = ["All"]
        if self.mgr.sdv_mgr.all_history_df is not None:
             hist_users = sorted(self.mgr.sdv_mgr.all_history_df['User'].unique().tolist())
             users.extend([str(u) for u in hist_users if str(u).strip()])
        
        user_combo = ttk.Combobox(ctrl_frame, textvariable=self.perf_user_var, values=users, width=20)
        user_combo.pack(side=tk.LEFT, padx=5)
        
        btn_run = ttk.Button(ctrl_frame, text="Analyze Performance", command=self._refresh_performance)
        btn_run.pack(side=tk.LEFT, padx=10)
        
        # Tree
        cols = ("User", "Date", "Site", "Patient", "Visit", "Pages Verified")
        tree_frame = tk.Frame(main_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree_perf = self._create_tree(tree_frame, cols, 'headings', 'performance')
        
    def _refresh_performance(self):
        start = self.perf_start_var.get().strip()
        end = self.perf_end_var.get().strip()
        user = self.perf_user_var.get()
        
        # Validate dates briefly
        try:
            if start: datetime.datetime.strptime(start, '%Y-%m-%d')
            if end: datetime.datetime.strptime(end, '%Y-%m-%d')
        except ValueError:
            messagebox.showerror("Error", "Invalid date format. Use YYYY-MM-DD")
            return

        df = self.mgr.get_cra_activity(start, end, user_filter=user)
        
        # Clear tree
        self.tree_perf.delete(*self.tree_perf.get_children())
        
        if df.empty:
             messagebox.showinfo("CRA Performance", "No activity found for the selected period/CRA.")
             return
             
        # Populate
        total_pages = 0
        for _, row in df.iterrows():
            vals = (row['User'], row['Date'], row.get('Site',''), row['Patient'], row['Visit'], row['Pages Verified'])
            self.tree_perf.insert("", "end", values=vals)
            total_pages += int(row['Pages Verified'])
            
        messagebox.showinfo("CRA Performance", f"Analysis complete.\nTotal pages verified in period: {total_pages}")

    def _sort_tree(self, tree, col, reverse):
        l = [(tree.set(k, col), k) for k in tree.get_children('')]
        try:
            l.sort(key=lambda t: float(t[0].split()[0]), reverse=reverse) # Try numeric sort first
        except ValueError:
            l.sort(reverse=reverse)

        for index, (val, k) in enumerate(l):
            tree.move(k, '', index)

        tree.heading(col, command=lambda: self._sort_tree(tree, col, not reverse))

    def _on_tree_click(self, event, tree):
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
            
        col = tree.identify_column(event.x)
        row_id = tree.identify_row(event.y)
        
        if not row_id: 
            return
            
        col_idx = int(col.replace('#', '')) - 1
        col_name = tree['columns'][col_idx]
        
        # Check if clicked column is a metric column (V, !, NS, GAP)
        # Handle "Metric" column in Study tab specifically
        
        level = None
        level_id = None
        metric = None
        
        # Determine tab context
        current_tab = self.notebook.select()
        tab_index = self.notebook.index(current_tab)
        
        if tab_index == 0: # Study
            # In Study tab, rows are the metrics
            item = tree.item(row_id)
            vals = item['values']
            metric = vals[0] # "V", "!", etc.
            # If they clicked Count (col 1), show details.
            if col_idx == 1:
                level = 'study'
                level_id = 'all'
                
        elif tab_index == 1: # Site
            if col_name in ['V', '!', 'NS', 'GAP']:
                metric = col_name
                item = tree.item(row_id)
                level_id = item['values'][0] # Site Name
                level = 'site'
                
        elif tab_index == 2: # Patient
            if col_name in ['V', '!', 'NS', 'GAP']:
                metric = col_name
                item = tree.item(row_id)
                level_id = item['values'][0] # Patient ID
                level = 'patient'

        elif tab_index == 3: # Form
             if col_name in ['V', '!', 'NS', 'GAP']:
                metric = col_name
                item = tree.item(row_id)
                pat = item['values'][0]
                frm = item['values'][1]
                level_id = (pat, frm)
                level = 'form'

        if level and metric and level_id:
            self._show_details(level, level_id, metric)

    def _show_details(self, level, level_id, metric):
        details = self.mgr.get_details(level, level_id, metric)
        if not details:
            messagebox.showinfo("Info", "No records found.")
            return
            
        DetailWindow(self.window, details, f"{level} {level_id} - {metric}")

class DetailWindow:
    def __init__(self, parent, data, title):
        self.win = tk.Toplevel(parent)
        self.win.title(f"Details: {title}")
        self.win.geometry("1100x600")
        self.data = data # List of dicts
        
        # Search Frame
        frame_top = tk.Frame(self.win)
        frame_top.pack(fill=tk.X, padx=5, pady=5)
        
        tk.Label(frame_top, text="Search:").pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_entry = tk.Entry(frame_top, textvariable=self.search_var)
        self.search_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.search_entry.bind("<KeyRelease>", self._filter_data)
        
        # Filter: Hide Empty Pending (! status with empty value)
        self.hide_empty_pending_var = tk.BooleanVar(value=True) # Default True
        chk_empty = tk.Checkbutton(frame_top, text="Hide Empty Pending",
                                   variable=self.hide_empty_pending_var, command=self._filter_data)
        chk_empty.pack(side=tk.LEFT, padx=10)
        
        # Export Button
        btn_export = ttk.Button(frame_top, text="Export to Excel", command=self._export_to_excel)
        btn_export.pack(side=tk.RIGHT)
        
        # Tree
        cols = ("Patient", "Visit", "Form", "Field", "Field ID", "Status", "VerifiedBy", "Date")
        
        self.tree = ttk.Treeview(self.win, columns=cols, show='headings')
        for c in cols:
            self.tree.heading(c, text=c, command=lambda _c=c: self._sort_tree(_c, False))
            # Adjust widths
            width = 120
            if c == "Field": width = 200
            elif c == "Field ID": width = 150
            elif c == "Date": width = 120
            elif c == "VerifiedBy": width = 100
            elif c in ["Patient", "Status", "Visit", "Form"]: width = 100
            
            self.tree.column(c, width=width)
            
        scrollbar = ttk.Scrollbar(self.win, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Initial load
        self._filter_data()

    def _populate_tree(self, items):
        self.tree.delete(*self.tree.get_children())
        # print(f"DEBUG: _populate_tree called with {len(items)} items")
        count = 0
        try:
            for item in items:
                # Debug check for missing keys first time
                # if count == 0:
                #    print(f"DEBUG: First item keys: {item.keys()}")
                    
                vals = (item.get('Patient',''), item.get('Visit',''), item.get('Form',''), 
                        item.get('Field',''), item.get('FieldID', item.get('Code', '')), # Get 'FieldID' (data key)
                        item.get('Status',''), item.get('VerifiedBy', ''), item.get('Date',''))
                self.tree.insert("", "end", values=vals)
                count += 1
        except Exception as e:
            logger.error("Error in _populate_tree: %s", e)
            import traceback
            traceback.print_exc()

    def _filter_data(self, event=None):
        query = self.search_var.get().lower()
        hide_empty = self.hide_empty_pending_var.get()
        
        filtered = []
        for item in self.data:
            # Check Hide Empty Pending filter
            # Logic: If item is Pending ('!') AND Value is empty/None -> Hide
            status = str(item.get('Status', ''))
            val = str(item.get('Value', '')).strip()
            if val.lower() == 'none': val = ''
            
            if hide_empty and status == '!' and not val:
                continue

            # Search in all values
            if query:
                if not any(query in str(v).lower() for v in item.values()):
                    continue
            
            filtered.append(item)
        
        self.current_items = filtered # Track current items
        self._populate_tree(filtered)

    def _export_to_excel(self):
        try:
            from tkinter import filedialog
            
            # Determine what to export: filtered or all?
            # Let's export what is currently in view if filtered
            items_to_export = getattr(self, 'current_items', self.data)
            
            if not items_to_export:
                messagebox.showinfo("Info", "No data to export.")
                return

            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Export Drill-down Data"
            )
            
            if not filename:
                return
                
            # Create DataFrame
            df = pd.DataFrame(items_to_export)
            # Reorder/rename cols to match UI if raw data keys differ
            # self.data keys: Patient, Site, Visit, Form, Field, Code, Value, Status, SDV, Date, User
            # Desired export cols: Patient, Visit, Form, Field, Field ID (Code), Value, Status, SDV, Date, User
            
            # Create a clean export dataframe with renamed columns
            export_df = pd.DataFrame()
            cols_map = {
                'Patient': 'Patient',
                'Visit': 'Visit',
                'Form': 'Form',
                'Field': 'Field',
                'Code': 'Field ID',
                'Value': 'Value',
                'Status': 'Metric', # This is V/!/GAP status from tree calculation
                'SDV': 'SDV Mark',
                'Date': 'Verification Date',
                'User': 'Verifier'
            }
            
            for key, new_name in cols_map.items():
                if key in df.columns:
                    export_df[new_name] = df[key]
            
            # Fallback for metric/Status confusion (Tree "Status" vs SDV)
            # In clinical_viewer1.py dashboard_records: Status is GAP/OK/Confirmed. SDV is V/! etc.
            # User wants to troubleshoot dashboard numbers, so let's include both clearly
            
            export_df.to_excel(filename, index=False)
            messagebox.showinfo("Success", f"Data exported successfully to:\n{filename}")
            
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export data:\n{str(e)}")
