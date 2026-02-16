import tkinter as tk
from tkinter import ttk

def setup_toolbar(app, root):
    """
    Setup the toolbar and filter UI elements.
    Args:
        app: The ClinicalDataMasterV30 instance.
        root: The root window or parent frame.
    """
    # Top Bar: File Loading & info
    top = tk.Frame(root, bg="#f4f4f4", pady=10)
    top.pack(fill=tk.X)
    
    tk.Button(top, text="Load Excel (Main Sheet)", command=app.load_excel, 
                bg="#2c3e50", fg="white", font=("Segoe UI", 10, "bold")).pack(side=tk.LEFT, padx=15)
    
    # File info label with cutoff time
    app.file_info_var = tk.StringVar(value="No file loaded")
    app.lbl_status = tk.Label(top, textvariable=app.file_info_var, bg="#f4f4f4", fg="#555", font=("Segoe UI", 9))
    app.lbl_status.pack(side=tk.LEFT)
    
    # Cutoff info specific label
    app.cutoff_var = tk.StringVar(value="")
    tk.Label(top, textvariable=app.cutoff_var, bg="#f4f4f4", fg="#c0392b", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=10)

    # Filter Frame
    flt = tk.Frame(root)
    flt.pack(fill=tk.X, padx=10, pady=5)
    
    # Helper for adding filters
    def add_filter(parent, label_text):
        frame = tk.Frame(parent)
        frame.pack(side=tk.LEFT, padx=10)
        tk.Label(frame, text=label_text, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        cb = ttk.Combobox(frame, state="readonly", width=25)
        cb.pack(side=tk.LEFT, padx=5)
        return cb

    app.cb_site = add_filter(flt, "Site:")
    app.cb_site.bind("<<ComboboxSelected>>", app.update_patients)
    app.cb_pat = add_filter(flt, "Patient:")
    # Binding for patient selection to refresh view
    app.cb_pat.bind("<<ComboboxSelected>>", lambda e: app.generate_view())

    # View Options Frame
    view_frame = tk.LabelFrame(root, text=" View Options ", padx=10, pady=5)
    view_frame.pack(fill=tk.X, padx=10, pady=5)

    # Row 1: View controls (radio buttons, checkboxes, search)
    ctrl_row = tk.Frame(view_frame)
    ctrl_row.pack(fill=tk.X, pady=(0, 5))

    app.view_mode = tk.StringVar(value="assess")
    tk.Radiobutton(ctrl_row, text="Assessment -> Visit", variable=app.view_mode, value="assess", command=app.generate_view).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(ctrl_row, text="Visit Phase -> Assessment", variable=app.view_mode, value="visit", command=app.generate_view).pack(side=tk.LEFT, padx=5)

    tk.Frame(ctrl_row, width=20).pack(side=tk.LEFT)
    app.chk_hide_dup = tk.BooleanVar(value=True)
    tk.Checkbutton(ctrl_row, text="Smart Dedupe", variable=app.chk_hide_dup, command=app.generate_view).pack(side=tk.LEFT, padx=5)
    app.chk_hide_future = tk.BooleanVar(value=True)
    tk.Checkbutton(ctrl_row, text="Hide Unstarted Visits", variable=app.chk_hide_future, command=app.generate_view).pack(side=tk.LEFT, padx=5)

    tk.Label(ctrl_row, text="Search:").pack(side=tk.LEFT, padx=(20, 5))
    app.search_var = tk.StringVar()
    app._search_debounce_id = None
    def _on_search_change(*_args):
        if app._search_debounce_id:
            app.root.after_cancel(app._search_debounce_id)
        app._search_debounce_id = app.root.after(300, app.generate_view)
    app.search_var.trace("w", _on_search_change)
    tk.Entry(ctrl_row, textvariable=app.search_var, width=20).pack(side=tk.LEFT, padx=5)

    # Row 2: All action buttons
    btn_row = tk.Frame(view_frame)
    btn_row.pack(fill=tk.X)

    btn_font = ("Segoe UI", 8, "bold")

    tk.Button(btn_row, text="Refresh Tree", command=app.generate_view, 
                bg="#27ae60", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Export View", command=app.export_view, 
                bg="#e67e22", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Data Matrix", command=app.show_data_matrix, 
                bg="#8e44ad", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Visit Sched", command=app.show_visit_schedule, 
                bg="#16a085", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Dashboard", command=app.open_dashboard, 
                bg="#c0392b", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)

    app.sdv_btn = tk.Button(btn_row, text="📋 SDV", command=app.load_sdv_data, 
                bg="#27ae60", fg="white", font=btn_font)
    app.sdv_btn.pack(side=tk.LEFT, padx=2)

    tk.Button(btn_row, text="Echo", command=app.show_echo_export, 
                bg="#2980b9", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="CVC", command=app.show_cvc_export, 
                bg="#e67e22", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Labs", command=app.show_labs_export, 
                bg="#9b59b6", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="FU Highlights", command=app.show_fu_highlights, 
                bg="#1abc9c", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Proc Timing", command=app.show_procedure_timing, 
                bg="#34495e", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="HF Hosps", command=app.show_hf_hospitalizations, 
                bg="#e74c3c", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="AE Module", command=app.show_ae_module, 
                bg="#d35400", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Assess.", command=app.show_assessment_data_table, 
                bg="#3498db", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Batch Export", command=app.show_batch_export, 
                bg="#2c3e50", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="Compare", command=app.show_data_comparison, 
                bg="#7f8c8d", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="📂 Sources", command=app.show_data_sources, 
                bg="#6c5ce7", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
    tk.Button(btn_row, text="📅 Timeline", command=app.show_patient_timeline, 
                bg="#e17055", fg="white", font=btn_font).pack(side=tk.LEFT, padx=2)
