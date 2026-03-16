import tkinter as tk
from tkinter import ttk

# Modern light color palette
_BG = "#f8f9fa"          # Main background
_CARD_BG = "#ffffff"      # Card/panel background
_ACCENT = "#4361ee"       # Primary accent (blue)
_ACCENT_HOVER = "#3a56d4"
_TEXT = "#212529"          # Primary text
_TEXT_MUTED = "#6c757d"   # Secondary text
_BORDER = "#dee2e6"       # Subtle borders
_SUCCESS = "#2d9f5e"      # Green
_WARN = "#e8590c"         # Orange
_DANGER = "#dc3545"       # Red


def _make_btn(parent, text, command, bg=_ACCENT, fg="white", **kw):
    """Create a modern flat button."""
    btn = tk.Button(parent, text=text, command=command,
                    bg=bg, fg=fg, activebackground=_ACCENT_HOVER, activeforeground="white",
                    font=("Segoe UI", 9), relief="flat", bd=0, padx=10, pady=4,
                    cursor="hand2", **kw)
    return btn


def setup_toolbar(app, root):
    """
    Setup the toolbar and filter UI elements.
    Args:
        app: The ClinicalDataMasterV30 instance.
        root: The root window or parent frame.
    """
    root.configure(bg=_BG)

    # --- Top Bar: File Loading & info ---
    top = tk.Frame(root, bg=_CARD_BG, pady=8, padx=12, highlightbackground=_BORDER, highlightthickness=1)
    top.pack(fill=tk.X, padx=8, pady=(8, 0))

    _make_btn(top, "Load Excel", app.load_excel, bg="#2c3e50").pack(side=tk.LEFT, padx=(0, 12))

    app.file_info_var = tk.StringVar(value="No file loaded")
    app.lbl_status = tk.Label(top, textvariable=app.file_info_var, bg=_CARD_BG,
                              fg=_TEXT_MUTED, font=("Segoe UI", 9))
    app.lbl_status.pack(side=tk.LEFT)

    app.cutoff_var = tk.StringVar(value="")
    tk.Label(top, textvariable=app.cutoff_var, bg=_CARD_BG,
             fg=_DANGER, font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=10)

    # --- Filter Row ---
    flt = tk.Frame(root, bg=_CARD_BG, pady=6, padx=12, highlightbackground=_BORDER, highlightthickness=1)
    flt.pack(fill=tk.X, padx=8, pady=(4, 0))

    def add_filter(parent, label_text):
        frame = tk.Frame(parent, bg=_CARD_BG)
        frame.pack(side=tk.LEFT, padx=(0, 16))
        tk.Label(frame, text=label_text, font=("Segoe UI", 9, "bold"),
                 bg=_CARD_BG, fg=_TEXT).pack(side=tk.LEFT, padx=(0, 4))
        cb = ttk.Combobox(frame, state="readonly", width=25)
        cb.pack(side=tk.LEFT)
        return cb

    app.cb_site = add_filter(flt, "Site:")
    app.cb_site.bind("<<ComboboxSelected>>", app.update_patients)
    app.cb_pat = add_filter(flt, "Patient:")
    app.cb_pat.bind("<<ComboboxSelected>>", lambda e: app.generate_view())

    # --- View Options Panel ---
    view_frame = tk.Frame(root, bg=_CARD_BG, pady=6, padx=12,
                          highlightbackground=_BORDER, highlightthickness=1)
    view_frame.pack(fill=tk.X, padx=8, pady=(4, 0))

    # Row 1: View controls
    ctrl_row = tk.Frame(view_frame, bg=_CARD_BG)
    ctrl_row.pack(fill=tk.X, pady=(0, 4))

    app.view_mode = tk.StringVar(value="assess")
    tk.Radiobutton(ctrl_row, text="Assessment -> Visit", variable=app.view_mode,
                   value="assess", command=app.generate_view,
                   bg=_CARD_BG, fg=_TEXT, font=("Segoe UI", 9),
                   activebackground=_CARD_BG, selectcolor=_CARD_BG).pack(side=tk.LEFT, padx=5)
    tk.Radiobutton(ctrl_row, text="Visit -> Assessment", variable=app.view_mode,
                   value="visit", command=app.generate_view,
                   bg=_CARD_BG, fg=_TEXT, font=("Segoe UI", 9),
                   activebackground=_CARD_BG, selectcolor=_CARD_BG).pack(side=tk.LEFT, padx=5)

    tk.Frame(ctrl_row, width=16, bg=_CARD_BG).pack(side=tk.LEFT)
    app.chk_hide_dup = tk.BooleanVar(value=True)  # Kept for cache key compatibility
    app.chk_hide_future = tk.BooleanVar(value=True)
    tk.Checkbutton(ctrl_row, text="Hide Unstarted Visits", variable=app.chk_hide_future,
                   command=app.generate_view, bg=_CARD_BG, fg=_TEXT,
                   font=("Segoe UI", 9), activebackground=_CARD_BG,
                   selectcolor=_CARD_BG).pack(side=tk.LEFT, padx=5)

    tk.Label(ctrl_row, text="Search:", bg=_CARD_BG, fg=_TEXT,
             font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(20, 4))
    app.search_var = tk.StringVar()
    app._search_debounce_id = None
    def _on_search_change(*_args):
        if app._search_debounce_id:
            app.root.after_cancel(app._search_debounce_id)
        app._search_debounce_id = app.root.after(300, app.generate_view)
    app.search_var.trace("w", _on_search_change)
    search_entry = tk.Entry(ctrl_row, textvariable=app.search_var, width=22,
                            font=("Segoe UI", 9), relief="solid", bd=1)
    search_entry.pack(side=tk.LEFT, padx=4)

    # Row 2: Action buttons - grouped by function
    btn_row = tk.Frame(view_frame, bg=_CARD_BG)
    btn_row.pack(fill=tk.X, pady=(2, 0))

    # Core actions
    _make_btn(btn_row, "Refresh", app.generate_view, bg=_SUCCESS).pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Export View", app.export_view, bg=_TEXT_MUTED).pack(side=tk.LEFT, padx=2)

    # Separator
    tk.Frame(btn_row, width=2, bg=_BORDER).pack(side=tk.LEFT, padx=6, fill=tk.Y, pady=2)

    # Data views
    _make_btn(btn_row, "Data Matrix", app.show_data_matrix, bg="#6c5ce7").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Visit Sched", app.show_visit_schedule, bg="#00b894").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Assess.", app.show_assessment_data_table, bg="#0984e3").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Proc Timing", app.show_procedure_timing, bg="#636e72").pack(side=tk.LEFT, padx=2)

    tk.Frame(btn_row, width=2, bg=_BORDER).pack(side=tk.LEFT, padx=6, fill=tk.Y, pady=2)

    # SDV & Dashboard
    app.sdv_btn = _make_btn(btn_row, "SDV", app.load_sdv_data, bg=_SUCCESS)
    app.sdv_btn.pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Dashboard", app.open_dashboard, bg=_DANGER).pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Gaps", app.show_data_gaps, bg=_WARN).pack(side=tk.LEFT, padx=2)

    tk.Frame(btn_row, width=2, bg=_BORDER).pack(side=tk.LEFT, padx=6, fill=tk.Y, pady=2)

    # Clinical modules
    _make_btn(btn_row, "AE Module", app.show_ae_module, bg="#d63031").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "HF Hosps", app.show_hf_hospitalizations, bg="#e17055").pack(side=tk.LEFT, padx=2)

    tk.Frame(btn_row, width=2, bg=_BORDER).pack(side=tk.LEFT, padx=6, fill=tk.Y, pady=2)

    # Exports
    _make_btn(btn_row, "Echo", app.show_echo_export, bg="#74b9ff").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "CVC", app.show_cvc_export, bg="#fdcb6e").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Labs", app.show_labs_export, bg="#a29bfe").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "FU Highlights", app.show_fu_highlights, bg="#55efc4").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Batch Export", app.show_batch_export, bg="#2d3436").pack(side=tk.LEFT, padx=2)

    tk.Frame(btn_row, width=2, bg=_BORDER).pack(side=tk.LEFT, padx=6, fill=tk.Y, pady=2)

    # Utilities
    _make_btn(btn_row, "Compare", app.show_data_comparison, bg="#b2bec3").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Sources", app.show_data_sources, bg="#6c5ce7").pack(side=tk.LEFT, padx=2)
    _make_btn(btn_row, "Timeline", app.show_patient_timeline, bg="#e17055").pack(side=tk.LEFT, padx=2)
