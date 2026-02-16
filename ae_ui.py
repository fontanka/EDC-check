import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
from ae_manager import AEManager
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

class AEWindow:
    def __init__(self, root, ae_manager: AEManager):
        self.root = root
        self.mgr = ae_manager
        
        self.window = tk.Toplevel(root)
        self.window.title("Adverse Event Module")
        self.window.geometry("1200x800")
        
        # Style
        style = ttk.Style(self.window)
        style.configure("AE.TLabel", font=('Segoe UI', 10))
        style.configure("AE.TButton", font=('Segoe UI', 10))
        style.configure("Card.TFrame", background="white", relief="raised")
        
        # Notebook for Tabs
        self.notebook = ttk.Notebook(self.window)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        self.tab_dashboard = ttk.Frame(self.notebook)
        self.tab_browser = ttk.Frame(self.notebook)
        
        self.notebook.add(self.tab_dashboard, text="AE Dashboard")
        self.notebook.add(self.tab_browser, text="AE Browser")
        
        self._build_dashboard_tab()
        self._build_browser_tab()

    def _build_dashboard_tab(self):
        # Stats
        # We need to get exclude list from entry if it exists, but it doesn't exist yet on first build.
        # So first build is empty exclusion.
        
        if hasattr(self, 'exclude_entry'):
             self.excluded_patients_str = self.exclude_entry.get()
             
        excluded_list = [p.strip() for p in self.excluded_patients_str.split(',')] if hasattr(self, 'excluded_patients_str') and self.excluded_patients_str else []
        
        if not hasattr(self, 'exclude_pre_proc_var'):
            self.exclude_pre_proc_var = tk.BooleanVar(value=False)
        
        stats = self.mgr.get_summary_stats(excluded_patients=excluded_list, exclude_pre_proc=self.exclude_pre_proc_var.get())
    
        # Scrollable Main Frame
        canvas = tk.Canvas(self.tab_dashboard, bg="#ecf0f1")
        scrollbar = ttk.Scrollbar(self.tab_dashboard, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#ecf0f1")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        create_window_id = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        
        # Mousewheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        # Resize internal frame to match canvas width
        def _configure_canvas(event):
            canvas.itemconfig(create_window_id, width=event.width)
        canvas.bind("<Configure>", _configure_canvas)
        
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # --- ROW 0: FILTERS / CONTROLS ---
        # Add a frame at the top of scrollable_frame for controls
        controls_frame = tk.Frame(scrollable_frame, bg="#ecf0f1", pady=10)
        controls_frame.pack(fill=tk.X, padx=10)
        
        tk.Label(controls_frame, text="Exclude Patients (comma-sep):", font=("Segoe UI", 10), bg="#ecf0f1").pack(side=tk.LEFT, padx=5)
        self.exclude_entry = tk.Entry(controls_frame, font=("Segoe UI", 10), width=30)
        self.exclude_entry.pack(side=tk.LEFT, padx=5)
        
        # Add current excluded patients to entry if any
        # Need to store this state in window?
        if not hasattr(self, 'excluded_patients_str'):
            self.excluded_patients_str = ""
        self.exclude_entry.insert(0, self.excluded_patients_str)
        
        tk.Checkbutton(controls_frame, text="Exclude Pre-Proc AEs", variable=self.exclude_pre_proc_var,
                       bg="#ecf0f1", font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=10)
        
        tk.Button(controls_frame, text="Apply Filter", command=self._refresh_dashboard,
                  bg="#34495e", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=10)
        
        # --- ROW 1: METRICS ---
        metrics_frame = tk.Frame(scrollable_frame, bg="#ecf0f1", pady=10)
        metrics_frame.pack(fill=tk.X, padx=10)
        
        # 5 Cards
        self._create_card(metrics_frame, "Total Patients", stats.get('patients_with_aes', 0), "#2c3e50")
        self._create_card(metrics_frame, "Total AEs", stats.get('total_aes', 0), "#3498db")
        self._create_card(metrics_frame, "Total SAEs", stats.get('total_saes', 0), "#e67e22")
        self._create_card(metrics_frame, "Ongoing AEs", stats.get('ongoing_aes', 0), "#f1c40f")
        self._create_card(metrics_frame, "Fatal Cases", stats.get('fatal_cases', 0), "#c0392b")

        # --- ROW 1.5: TEXT SUMMARIES ---
        summary_frame = tk.Frame(scrollable_frame, bg="#ecf0f1", pady=5)
        summary_frame.pack(fill=tk.X, padx=20)
        
        # Outcome Text
        outcomes = stats.get('outcome_dist', {})
        if outcomes:
            out_str = " | ".join([f"{k}: {v}" for k, v in outcomes.items()])
            tk.Label(summary_frame, text=f"Outcome Distribution: {out_str}", 
                     font=("Segoe UI", 11), bg="#ecf0f1", fg="#2c3e50", justify="left").pack(anchor="w")
                     
        # SAE Criteria Text
        criteria = stats.get('sae_criteria', {})
        if criteria:
            crit_str = " | ".join([f"{k}: {v}" for k, v in criteria.items() if v > 0])
            if crit_str:
                tk.Label(summary_frame, text=f"SAE Criteria: {crit_str}", 
                         font=("Segoe UI", 11), bg="#ecf0f1", fg="#c0392b", justify="left").pack(anchor="w")

        # --- ROW 2: CHARTS ---
        charts_frame = tk.Frame(scrollable_frame, bg="#ecf0f1", pady=10)
        charts_frame.pack(fill=tk.X, padx=10)
        
        # Figure 1: Outcome Distribution (Pie/Bar) and SAE Criteria (Bar)
        # We'll use subplots
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(10, 4))
        fig.patch.set_facecolor('#ecf0f1')
        
        # Outcome
        if outcomes:
            labels = list(outcomes.keys())
            sizes = list(outcomes.values())
            # Simple bar for readability if many text labels
            y_pos = range(len(labels))
            bars1 = ax1.barh(y_pos, sizes, color='#1abc9c')
            ax1.bar_label(bars1, padding=3) # Add numbers
            ax1.set_yticks(y_pos)
            ax1.set_yticklabels(labels)
            ax1.invert_yaxis()
            ax1.set_title('Outcome Distribution')
        else:
            ax1.text(0.5, 0.5, 'No Data', ha='center')
            
        # SAE Criteria
        # Filter zero values? No, show zeros if requested to track
        c_labels = list(criteria.keys())
        c_vals = list(criteria.values())
        y_pos2 = range(len(c_labels))
        bars2 = ax2.barh(y_pos2, c_vals, color='#e74c3c')
        ax2.bar_label(bars2, padding=3) # Add numbers
        ax2.set_yticks(y_pos2)
        ax2.set_yticklabels(c_labels)
        ax2.invert_yaxis()
        ax2.set_title('SAE Criteria Counts')
        
        plt.tight_layout()
        canvas_charts = FigureCanvasTkAgg(fig, master=charts_frame)
        canvas_charts.draw()
        canvas_charts.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- ROW 3: TOP TERMS & SITES ---
        row3_frame = tk.Frame(scrollable_frame, bg="#ecf0f1", pady=10)
        row3_frame.pack(fill=tk.X, padx=10)
        
        fig2, (ax3, ax4) = plt.subplots(1, 2, figsize=(10, 4))
        fig2.patch.set_facecolor('#ecf0f1')
        
        # Top Terms
        terms = stats.get('top_terms', {})
        if terms:
            t_labels = list(terms.keys())
            t_vals = list(terms.values())
            # Truncate long labels
            t_labels = [(t[:20] + '...') if len(t) > 20 else t for t in t_labels]
            
            y_pos3 = range(len(t_labels))
            bars3 = ax3.barh(y_pos3, t_vals, color='#9b59b6')
            ax3.bar_label(bars3, padding=3) # Add numbers
            ax3.set_yticks(y_pos3)
            ax3.set_yticklabels(t_labels)
            ax3.invert_yaxis()
            ax3.set_title('Top 10 AE Terms')
        else:
            ax3.text(0.5, 0.5, 'No Data', ha='center')
            
        # Site Data
        site_data = stats.get('by_site', {})
        if site_data:
            s_labels = list(site_data.keys())
            s_vals = list(site_data.values())
            bars4 = ax4.bar(s_labels, s_vals, color='#34495e')
            ax4.bar_label(bars4, padding=3) # Add numbers
            ax4.set_title('AEs by Site')
        else:
            ax4.text(0.5, 0.5, 'No Data', ha='center')
            
        plt.tight_layout()
        canvas_charts2 = FigureCanvasTkAgg(fig2, master=row3_frame)
        canvas_charts2.draw()
        canvas_charts2.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # --- ROW 3.5: RELATEDNESS STATISTICS TABLE ---
        rel_frame = tk.LabelFrame(scrollable_frame, text="Relatedness Statistics", font=("Segoe UI", 12, "bold"), bg="#ecf0f1", padx=10, pady=10)
        rel_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Grid Layout
        # Headers
        headers = ["Category", "Related", "Probably Related", "Possibly Related", "Not Related", "Unknown/Blank", "Related+\nProbably Related"]
        # Colors matching user screenshot feel (approx) - Deep Red headers?
        header_bg = "#900C3F"
        header_fg = "white"
        
        for col, text in enumerate(headers):
            lbl = tk.Label(rel_frame, text=text, font=("Segoe UI", 10, "bold"), bg=header_bg, fg=header_fg, padx=5, pady=5, borderwidth=1, relief="solid")
            lbl.grid(row=0, column=col, sticky="nsew")
        
        # Data Rows
        rel_stats = stats.get('relatedness_table', {})
        rows = ['Device', 'Delivery System', 'Handle', 'Procedure']
        
        # Map table keys to mapped keys in stats
        # They should match exactly as created in manager
        
        for r_idx, row_name in enumerate(rows):
            row_data = rel_stats.get(row_name, {})
            
            # Row Header
            tk.Label(rel_frame, text=row_name, font=("Segoe UI", 10), bg="#e8e8e8", padx=5, pady=5, borderwidth=1, relief="solid", anchor="w").grid(row=r_idx+1, column=0, sticky="nsew")
            
            # Values
            # Match header order
            keys = ["Related", "Probably Related", "Possibly Related", "Not Related", "Unknown/Blank", "Related+Probably"]
            
            for c_idx, key in enumerate(keys):
                val = row_data.get(key, 0)
                bg_color = "white"
                if (r_idx % 2) == 1: bg_color = "#f2f2f2" # Zebra striping
                
                tk.Label(rel_frame, text=str(val), font=("Segoe UI", 10), bg=bg_color, padx=5, pady=5, borderwidth=1, relief="solid").grid(row=r_idx+1, column=c_idx+1, sticky="nsew")

        # Configure grid weights
        for c in range(len(headers)):
            rel_frame.columnconfigure(c, weight=1)

        # --- ROW 4: PER-PATIENT DETAILS ---
        details_frame = tk.LabelFrame(scrollable_frame, text="Per-Patient Details", font=("Segoe UI", 12, "bold"), bg="#ecf0f1", padx=10, pady=10)
        details_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=20)
        
        details_text = tk.Text(details_frame, height=15, width=100, font=("Consolas", 10))
        details_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        dt_scroll = ttk.Scrollbar(details_frame, orient="vertical", command=details_text.yview)
        dt_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        details_text.configure(yscrollcommand=dt_scroll.set)
        
        # Insert Data
        lines = stats.get('per_patient_details', [])
        details_text.insert(tk.END, "\n".join(lines))
        details_text.config(state="disabled") # Read-only

    def _create_card(self, parent, title, value, color):
        card = tk.Frame(parent, bg="white", highlightbackground=color, highlightthickness=2, padx=20, pady=15)
        card.pack(side=tk.LEFT, expand=True, padx=20)
        
        tk.Label(card, text=title, font=("Segoe UI", 12), bg="white", fg="#7f8c8d").pack()
        tk.Label(card, text=str(value), font=("Segoe UI", 24, "bold"), bg="white", fg=color).pack()

    def _build_browser_tab(self):
        # Left Sidebar: Controls
        sidebar = tk.Frame(self.tab_browser, width=300, bg="#f4f4f4", padx=10, pady=10)
        sidebar.pack(side=tk.LEFT, fill=tk.Y)
        sidebar.pack_propagate(False)
        
        # Patient Selector
        tk.Label(sidebar, text="Select Patient:", font=("Segoe UI", 10, "bold"), bg="#f4f4f4").pack(anchor="w")
        
        # Get patient list from AE data
        if self.mgr.df_ae is not None and not self.mgr.df_ae.empty:
            patients = sorted(self.mgr.df_ae['Screening #'].unique().tolist())
        else:
            patients = []
            
        self.pat_var = tk.StringVar()
        self.pat_combo = ttk.Combobox(sidebar, textvariable=self.pat_var, values=patients, state="readonly")
        self.pat_combo.pack(fill=tk.X, pady=(0, 20))
        self.pat_combo.bind("<<ComboboxSelected>>", self._refresh_ae_table)
        
        # Filters Group
        filter_group = tk.LabelFrame(sidebar, text="Filters", font=("Segoe UI", 10, "bold"), bg="#f4f4f4", padx=10, pady=10)
        filter_group.pack(fill=tk.X)
        
        # Pre-Procedure Checkbox
        self.exclude_pre_proc_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_group, text="Exclude Pre-Procedure AEs", variable=self.exclude_pre_proc_var, 
                       bg="#f4f4f4", command=self._refresh_ae_table).pack(anchor="w", pady=2)
                       
        # SAE Filter (Radio)
        self.sae_mode_var = tk.StringVar(value="all")
        tk.Radiobutton(filter_group, text="All AEs", variable=self.sae_mode_var, value="all", 
                       bg="#f4f4f4", command=self._refresh_ae_table).pack(anchor="w", pady=2)
        tk.Radiobutton(filter_group, text="SAEs Only", variable=self.sae_mode_var, value="sae_only", 
                       bg="#f4f4f4", command=self._refresh_ae_table).pack(anchor="w", pady=2)
                       
        # Relationship Filter (Device/Proc Only)
        self.device_rel_var = tk.BooleanVar(value=False)
        tk.Checkbutton(filter_group, text="Device/Procedure Related Only", variable=self.device_rel_var, 
                       bg="#f4f4f4", command=self._on_device_rel_change, justify="left").pack(anchor="w", pady=(10, 2))
        
        # Date Cutoffs
        tk.Label(filter_group, text="Onset Date Cutoff (YYYY-MM-DD):", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", pady=(10, 0))
        self.onset_cutoff_var = tk.StringVar()
        self.onset_cutoff_entry = tk.Entry(filter_group, textvariable=self.onset_cutoff_var)
        self.onset_cutoff_entry.pack(fill=tk.X, pady=2)
        self.onset_cutoff_entry.bind('<Return>', self._refresh_ae_table)
        
        tk.Label(filter_group, text="Report Date Cutoff (YYYY-MM-DD):", bg="#f4f4f4", font=("Segoe UI", 9)).pack(anchor="w", pady=(5, 0))
        self.report_cutoff_var = tk.StringVar()
        self.report_cutoff_entry = tk.Entry(filter_group, textvariable=self.report_cutoff_var)
        self.report_cutoff_entry.pack(fill=tk.X, pady=2)
        self.report_cutoff_entry.bind('<Return>', self._refresh_ae_table)
        
        tk.Button(filter_group, text="Apply Filters", command=self._refresh_ae_table, bg="#f4f4f4").pack(fill=tk.X, pady=5)

        # Export Button
        tk.Button(sidebar, text="Export Filtered Data (All Patients)", command=self._export_data, 
                  bg="#27ae60", fg="white", font=("Segoe UI", 10, "bold")).pack(fill=tk.X, pady=20)
                       
        # Main Table Area
        self.table_frame = tk.Frame(self.tab_browser)
        self.table_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)
        
        # Treeview
        cols = ['AE #', 'SAE?', 'AE Term', 'Outcome', 'Onset Date', 'Resolution Date', 'Severity', 'Rel. PKG Trillium', 'Rel. Delivery System', 'Rel. Handle', 'Rel. Index Procedure']
        self.tree = ttk.Treeview(self.table_frame, columns=cols, show="headings")
        
        for c in cols:
            self.tree.heading(c, text=c)
            self.tree.column(c, width=100)
        
        # Scrollbars
        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(self.table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)
        hsb.pack(side=tk.BOTTOM, fill=tk.X)

    def _on_device_rel_change(self):
        # Independent filter now
        self._refresh_ae_table()

    def _export_data(self):
        # 1. Get Filters
        filters = {
            'sae_only': self.sae_mode_var.get() == "sae_only",
            'exclude_pre_proc': self.exclude_pre_proc_var.get(),
            'device_related_only': self.device_rel_var.get(),
            'onset_cutoff': self.onset_cutoff_var.get(),
            'report_cutoff': self.report_cutoff_var.get()
        }
        
        # 2. Get Data
        data = self.mgr.get_dataset_ae_data(filters)
        if not data:
            messagebox.showinfo("Export", "No data found matching filters.")
            return

        # 3. Convert to DF
        df = pd.DataFrame(data)
        
        # 4. Reorder Columns
        # Patient ID first
        cols = ['Patient ID', 'AE #', 'SAE?', 'AE Term', 'Outcome', 'Onset Date', 'Resolution Date', 'Severity', 
                'Rel. PKG Trillium', 'Rel. Delivery System', 'Rel. Handle', 'Rel. Index Procedure', 
                'Ongoing', 'Interval', 'AE Description', 'SAE Description']
        
        # Only include cols that exist in df
        final_cols = [c for c in cols if c in df.columns]
        # Add any remaining cols
        remaining = [c for c in df.columns if c not in final_cols]
        final_cols.extend(remaining)
        
        df = df[final_cols]
        
        # 5. Save
        timestamp = pd.Timestamp.now().strftime("%Y%m%d_%H%M")
        filename = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
            initialfile=f"AE_Export_Filtered_{timestamp}.xlsx",
            title="Export Filtered AE Data"
        )
        
        if filename:
            try:
                if filename.endswith('.csv'):
                    df.to_csv(filename, index=False)
                else:
                    df.to_excel(filename, index=False)
                messagebox.showinfo("Success", f"Data exported to {filename}")
            except Exception as e:
                messagebox.showerror("Error", f"Could not save file: {e}")

    def _refresh_ae_table(self, event=None):
        patient_id = self.pat_var.get()
        if not patient_id:
            return
            
        # Build Filters
        # Build Filters
        filters = {
            'sae_only': self.sae_mode_var.get() == "sae_only",
            'exclude_pre_proc': self.exclude_pre_proc_var.get(),
            'device_related_only': self.device_rel_var.get(),
            'onset_cutoff': self.onset_cutoff_var.get(),
            'report_cutoff': self.report_cutoff_var.get()
        }
        
        # Get Data
        ae_data = self.mgr.get_patient_ae_data(patient_id, filters)
        
        # Clear Tree
        self.tree.delete(*self.tree.get_children())
        
        if not ae_data:
            return
            
        # Insert Data
        cols = ['AE #', 'SAE?', 'AE Term', 'Outcome', 'Onset Date', 'Resolution Date', 'Severity', 'Rel. PKG Trillium', 'Rel. Delivery System', 'Rel. Handle', 'Rel. Index Procedure']
        for row in ae_data:
            vals = [row.get(c, '') for c in cols]
            self.tree.insert("", "end", values=vals)

    def _refresh_dashboard(self):
        # Save current exclusion string
        self.excluded_patients_str = self.exclude_entry.get()
        
        # Clear dashboard tab content
        for widget in self.tab_dashboard.winfo_children():
            widget.destroy()
            
        # Rebuild
        self._build_dashboard_tab()
