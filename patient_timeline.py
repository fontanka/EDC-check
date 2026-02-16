"""
Patient Timeline
================
Configurable multi-view timeline showing patient progression through study milestones.
Milestones are fully editable, addable, deletable, and persist across sessions.
"""

import os
import json
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime
from typing import Optional, Dict, List, Tuple
import pandas as pd
import numpy as np
import logging

logger = logging.getLogger("PatientTimeline")

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "milestone_config.json")

# Default milestones auto-populated from data
DEFAULT_MILESTONES = [
    {"id": "screening", "name": "Screening", "column_pattern": "SBV_SV_SVSTDTC", "color": "#3498db", "type": "data"},
    {"id": "consent", "name": "Consent", "column_pattern": "SBV_DSSTDTC", "color": "#2ecc71", "type": "data"},
    {"id": "treatment", "name": "Treatment", "column_pattern": "TV_PR_SVDTC", "color": "#e74c3c", "type": "data"},
    {"id": "fu_1m", "name": "1-Month FU", "column_pattern": "FU1M_SV_SVSTDTC", "color": "#f39c12", "type": "data"},
    {"id": "fu_6m", "name": "6-Month FU", "column_pattern": "FU6M_SV_SVSTDTC", "color": "#9b59b6", "type": "data"},
    {"id": "fu_1y", "name": "1-Year FU", "column_pattern": "FU1Y_SV_SVSTDTC", "color": "#1abc9c", "type": "data"},
]


class MilestoneConfig:
    """Manages milestone definitions with persistence."""

    def __init__(self):
        self.milestones: List[dict] = []
        self.manual_dates: Dict[str, Dict[str, str]] = {}  # patient_id -> {milestone_id: date_str}
        self._load()

    def _load(self):
        """Load config from JSON, falling back to defaults."""
        if os.path.isfile(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                self.milestones = data.get("milestones", DEFAULT_MILESTONES.copy())
                self.manual_dates = data.get("manual_dates", {})
                return
            except Exception as e:
                logger.warning(f"Failed to load milestone config: {e}")
        self.milestones = [m.copy() for m in DEFAULT_MILESTONES]

    def save(self):
        """Persist config to JSON."""
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({
                    "milestones": self.milestones,
                    "manual_dates": self.manual_dates,
                }, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save milestone config: {e}")

    def add_milestone(self, milestone_id: str, name: str, milestone_type: str = "manual",
                      column_pattern: str = "", color: str = "#95a5a6"):
        """Add a new milestone."""
        self.milestones.append({
            "id": milestone_id,
            "name": name,
            "column_pattern": column_pattern,
            "color": color,
            "type": milestone_type,
        })
        self.save()

    def remove_milestone(self, milestone_id: str):
        """Remove a milestone by ID."""
        self.milestones = [m for m in self.milestones if m["id"] != milestone_id]
        self.save()

    def move_milestone(self, milestone_id: str, direction: int):
        """Move a milestone up (-1) or down (+1) in order."""
        idx = next((i for i, m in enumerate(self.milestones) if m["id"] == milestone_id), None)
        if idx is None:
            return
        new_idx = idx + direction
        if 0 <= new_idx < len(self.milestones):
            self.milestones[idx], self.milestones[new_idx] = \
                self.milestones[new_idx], self.milestones[idx]
            self.save()

    def set_manual_date(self, patient_id: str, milestone_id: str, date_str: str):
        """Set a manual date for a patient milestone."""
        if patient_id not in self.manual_dates:
            self.manual_dates[patient_id] = {}
        self.manual_dates[patient_id][milestone_id] = date_str
        self.save()

    def get_manual_date(self, patient_id: str, milestone_id: str) -> Optional[str]:
        """Get a manual date if set."""
        return self.manual_dates.get(patient_id, {}).get(milestone_id)


def _parse_date(val) -> Optional[datetime]:
    """Try to parse a date value from various formats."""
    if pd.isna(val) or val is None:
        return None
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'none', 'nat', ''):
        return None
    for fmt in ("%Y-%m-%d", "%d-%b-%Y", "%d/%m/%Y", "%m/%d/%Y",
                "%Y-%m-%d %H:%M:%S", "%d-%b-%Y %H:%M:%S"):
        try:
            return datetime.strptime(s, fmt)
        except ValueError:
            continue
    try:
        return pd.to_datetime(s).to_pydatetime()
    except Exception:
        return None


def _find_column(df: pd.DataFrame, pattern: str) -> Optional[str]:
    """Find a column matching pattern (case-insensitive, partial)."""
    pattern_lower = pattern.lower()
    # Exact match first
    for col in df.columns:
        if col.lower() == pattern_lower:
            return col
    # Partial match
    for col in df.columns:
        if pattern_lower in col.lower():
            return col
    return None


class PatientTimelineWindow:
    """Main timeline window with tabs for Table, Gantt, and Summary views."""

    def __init__(self, parent: tk.Tk, df_main: pd.DataFrame,
                 get_screen_failures_fn=None):
        self.parent = parent
        self.df_main = df_main
        self.get_screen_failures = get_screen_failures_fn
        self.config = MilestoneConfig()

        self.win = tk.Toplevel(parent)
        self.win.title("Patient Timeline")
        self.win.geometry("1200x700")
        self.win.configure(bg="#1e1e2e")
        self.win.resizable(True, True)

        self._build_ui()
        self._refresh_all()

        self.win.transient(parent)
        self.win.focus_force()

    def _build_ui(self):
        # Header with controls
        header = tk.Frame(self.win, bg="#1e1e2e")
        header.pack(fill=tk.X, padx=10, pady=(10, 5))

        tk.Label(header, text="📅 Patient Timeline", font=("Segoe UI", 14, "bold"),
                 bg="#1e1e2e", fg="#cdd6f4").pack(side=tk.LEFT)

        btn_style = {"font": ("Segoe UI", 9, "bold"), "cursor": "hand2", "relief": "flat",
                      "padx": 10, "pady": 4}

        tk.Button(header, text="⚙ Milestones", bg="#585b70", fg="white",
                  command=self._edit_milestones, **btn_style).pack(side=tk.RIGHT, padx=5)
        tk.Button(header, text="🔄 Refresh", bg="#585b70", fg="white",
                  command=self._refresh_all, **btn_style).pack(side=tk.RIGHT, padx=5)

        self.exclude_sf_var = tk.BooleanVar(value=True)
        tk.Checkbutton(header, text="Exclude Screen Failures",
                       variable=self.exclude_sf_var, command=self._refresh_all,
                       bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                       activebackground="#1e1e2e", activeforeground="#cdd6f4",
                       font=("Segoe UI", 9)).pack(side=tk.RIGHT, padx=10)

        # Notebook for views
        self.notebook = ttk.Notebook(self.win)
        self.notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)

        # Table View tab
        self.table_frame = tk.Frame(self.notebook, bg="#1e1e2e")
        self.notebook.add(self.table_frame, text="  Table View  ")

        # Gantt View tab
        self.gantt_frame = tk.Frame(self.notebook, bg="#1e1e2e")
        self.notebook.add(self.gantt_frame, text="  Gantt View  ")

        # Summary View tab
        self.summary_frame = tk.Frame(self.notebook, bg="#1e1e2e")
        self.notebook.add(self.summary_frame, text="  Summary  ")

    def _get_patients(self) -> List[Tuple[str, pd.Series]]:
        """Get patient list, optionally excluding screen failures."""
        patients = []
        screen_failures = set()
        if self.exclude_sf_var.get() and self.get_screen_failures:
            screen_failures = set(self.get_screen_failures())

        for _, row in self.df_main.iterrows():
            pat_id = str(row.get('Screening #', '')).strip()
            if pat_id.endswith('.0'):
                pat_id = pat_id[:-2]
            if not pat_id or pat_id in screen_failures:
                continue
            patients.append((pat_id, row))
        return patients

    def _get_milestone_date(self, row: pd.Series, milestone: dict,
                             patient_id: str) -> Optional[datetime]:
        """Get the date for a milestone from the data row or manual entry."""
        # Manual date takes priority
        manual = self.config.get_manual_date(patient_id, milestone["id"])
        if manual:
            return _parse_date(manual)

        if milestone["type"] == "manual":
            return None

        # Data-bound: find column and extract date
        col = _find_column(self.df_main, milestone["column_pattern"])
        if col and col in row.index:
            return _parse_date(row[col])
        return None

    def _build_timeline_data(self) -> pd.DataFrame:
        """Build a DataFrame: patients (rows) × milestones (columns) with dates."""
        patients = self._get_patients()
        milestones = self.config.milestones

        rows = []
        for pat_id, row in patients:
            entry = {"Patient": pat_id}
            for ms in milestones:
                dt = self._get_milestone_date(row, ms, pat_id)
                entry[ms["name"]] = dt
            rows.append(entry)

        return pd.DataFrame(rows) if rows else pd.DataFrame()

    def _refresh_all(self):
        """Refresh all views."""
        self._refresh_table_view()
        self._refresh_gantt_view()
        self._refresh_summary_view()

    # ── TABLE VIEW ──────────────────────────────────────────────────────

    def _refresh_table_view(self):
        for w in self.table_frame.winfo_children():
            w.destroy()

        df = self._build_timeline_data()
        if df.empty:
            tk.Label(self.table_frame, text="No data available",
                     bg="#1e1e2e", fg="#cdd6f4", font=("Segoe UI", 12)).pack(pady=50)
            return

        milestones = self.config.milestones
        cols = ["Patient"] + [m["name"] for m in milestones]
        tree = ttk.Treeview(self.table_frame, columns=cols, show="headings",
                            height=25, selectmode="browse")

        style = ttk.Style()
        style.configure("TL.Treeview", background="#313244", foreground="#cdd6f4",
                         fieldbackground="#313244", font=("Segoe UI", 10), rowheight=26)
        style.configure("TL.Treeview.Heading", background="#45475a",
                         foreground="#cdd6f4", font=("Segoe UI", 9, "bold"))
        tree.configure(style="TL.Treeview")

        tree.heading("Patient", text="Patient", anchor="w")
        tree.column("Patient", width=100, minwidth=80)

        for ms in milestones:
            tree.heading(ms["name"], text=ms["name"], anchor="center")
            tree.column(ms["name"], width=120, minwidth=80, anchor="center")

        tree.tag_configure("complete", foreground="#a6e3a1")
        tree.tag_configure("partial", foreground="#fab387")
        tree.tag_configure("none", foreground="#f38ba8")

        for _, row in df.iterrows():
            values = [row["Patient"]]
            filled = 0
            for ms in milestones:
                dt = row.get(ms["name"])
                if pd.notna(dt) and dt is not None:
                    values.append(dt.strftime("%Y-%m-%d"))
                    filled += 1
                else:
                    values.append("—")

            ratio = filled / len(milestones) if milestones else 0
            tag = "complete" if ratio >= 0.8 else "partial" if ratio > 0 else "none"
            tree.insert("", "end", values=values, tags=(tag,))

        vsb = ttk.Scrollbar(self.table_frame, orient="vertical", command=tree.yview)
        tree.configure(yscrollcommand=vsb.set)
        tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        vsb.pack(side=tk.RIGHT, fill=tk.Y)

    # ── GANTT VIEW ──────────────────────────────────────────────────────

    def _refresh_gantt_view(self):
        for w in self.gantt_frame.winfo_children():
            w.destroy()

        df = self._build_timeline_data()
        if df.empty:
            tk.Label(self.gantt_frame, text="No data available",
                     bg="#1e1e2e", fg="#cdd6f4", font=("Segoe UI", 12)).pack(pady=50)
            return

        milestones = self.config.milestones
        if not milestones:
            return

        try:
            import matplotlib
            matplotlib.use("Agg")
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            import matplotlib.dates as mdates
        except ImportError:
            tk.Label(self.gantt_frame, text="matplotlib required for Gantt view",
                     bg="#1e1e2e", fg="#f38ba8", font=("Segoe UI", 12)).pack(pady=50)
            return

        patients = df["Patient"].tolist()
        n_patients = len(patients)

        # Determine date range
        all_dates = []
        for ms in milestones:
            col_dates = df[ms["name"]].dropna()
            all_dates.extend(col_dates.tolist())

        if not all_dates:
            tk.Label(self.gantt_frame, text="No dates found in data",
                     bg="#1e1e2e", fg="#fab387", font=("Segoe UI", 12)).pack(pady=50)
            return

        date_min = min(all_dates)
        date_max = max(all_dates)

        # Add padding
        from datetime import timedelta
        date_range = (date_max - date_min).days
        padding = max(timedelta(days=15), timedelta(days=int(date_range * 0.05)))
        date_min -= padding
        date_max += padding

        # Create figure
        fig_height = max(4, min(20, n_patients * 0.35 + 2))
        fig, ax = plt.subplots(figsize=(12, fig_height))
        fig.patch.set_facecolor("#1e1e2e")
        ax.set_facecolor("#313244")

        # Plot each patient as a row
        y_positions = list(range(n_patients))

        for ms in milestones:
            dates = []
            y_vals = []
            for i, (_, row) in enumerate(df.iterrows()):
                dt = row.get(ms["name"])
                if pd.notna(dt) and dt is not None:
                    dates.append(dt)
                    y_vals.append(i)

            if dates:
                ax.scatter(dates, y_vals, c=ms["color"], s=80, zorder=3,
                           label=ms["name"], edgecolors="white", linewidths=0.5)

        # Draw connecting lines between milestones per patient
        for i, (_, row) in enumerate(df.iterrows()):
            pat_dates = []
            for ms in milestones:
                dt = row.get(ms["name"])
                if pd.notna(dt) and dt is not None:
                    pat_dates.append(dt)
            if len(pat_dates) >= 2:
                pat_dates.sort()
                ax.plot([pat_dates[0], pat_dates[-1]], [i, i],
                         color="#585b70", linewidth=1.5, zorder=1)

        # Formatting
        ax.set_yticks(y_positions)
        ax.set_yticklabels(patients, fontsize=8, color="#cdd6f4")
        ax.invert_yaxis()
        ax.xaxis.set_major_formatter(mdates.DateFormatter("%b %Y"))
        ax.xaxis.set_major_locator(mdates.MonthLocator(interval=max(1, date_range // 180)))
        plt.setp(ax.xaxis.get_majorticklabels(), rotation=45, ha="right",
                 fontsize=8, color="#cdd6f4")
        ax.tick_params(axis="y", colors="#cdd6f4")
        ax.tick_params(axis="x", colors="#cdd6f4")

        ax.set_xlim(date_min, date_max)
        ax.set_title("Patient Timeline — Gantt View", fontsize=12,
                      color="#cdd6f4", pad=15)
        ax.legend(loc="upper right", fontsize=8, facecolor="#45475a",
                  edgecolor="#585b70", labelcolor="#cdd6f4")

        for spine in ax.spines.values():
            spine.set_color("#585b70")

        ax.grid(axis="x", color="#45475a", linewidth=0.5, alpha=0.5)

        fig.tight_layout()

        # Embed in tkinter
        canvas = FigureCanvasTkAgg(fig, self.gantt_frame)
        canvas.draw()
        canvas.get_tk_widget().pack(fill=tk.BOTH, expand=True)

    # ── SUMMARY VIEW ────────────────────────────────────────────────────

    def _refresh_summary_view(self):
        for w in self.summary_frame.winfo_children():
            w.destroy()

        df = self._build_timeline_data()
        if df.empty:
            tk.Label(self.summary_frame, text="No data available",
                     bg="#1e1e2e", fg="#cdd6f4", font=("Segoe UI", 12)).pack(pady=50)
            return

        milestones = self.config.milestones
        n_milestones = len(milestones)
        n_patients = len(df)

        # Stats frame
        stats_frame = tk.Frame(self.summary_frame, bg="#1e1e2e")
        stats_frame.pack(fill=tk.X, padx=15, pady=10)

        # Overall stats
        card_style = {"font": ("Segoe UI", 11), "padx": 20, "pady": 15, "relief": "flat"}

        total_cells = n_patients * n_milestones
        filled_cells = 0
        for ms in milestones:
            filled_cells += df[ms["name"]].notna().sum()

        completion_pct = (filled_cells / total_cells * 100) if total_cells > 0 else 0

        cards = [
            ("Patients", str(n_patients), "#3498db"),
            ("Milestones", str(n_milestones), "#2ecc71"),
            ("Completion", f"{completion_pct:.1f}%", "#e74c3c" if completion_pct < 50 else "#f39c12" if completion_pct < 80 else "#2ecc71"),
            ("Data Points", f"{filled_cells}/{total_cells}", "#9b59b6"),
        ]

        for text, value, color in cards:
            card = tk.Frame(stats_frame, bg="#313244", padx=20, pady=12)
            card.pack(side=tk.LEFT, padx=8, expand=True, fill=tk.X)
            tk.Label(card, text=value, font=("Segoe UI", 18, "bold"),
                     bg="#313244", fg=color).pack()
            tk.Label(card, text=text, font=("Segoe UI", 9),
                     bg="#313244", fg="#a6adc8").pack()

        # Per-milestone completion
        ms_frame = tk.LabelFrame(self.summary_frame, text=" Milestone Completion ",
                                 bg="#1e1e2e", fg="#cdd6f4", font=("Segoe UI", 10, "bold"),
                                 padx=10, pady=10)
        ms_frame.pack(fill=tk.X, padx=15, pady=5)

        for ms in milestones:
            filled = df[ms["name"]].notna().sum()
            pct = filled / n_patients * 100 if n_patients > 0 else 0

            row = tk.Frame(ms_frame, bg="#1e1e2e")
            row.pack(fill=tk.X, pady=2)

            tk.Label(row, text=ms["name"], width=15, anchor="w",
                     bg="#1e1e2e", fg="#cdd6f4", font=("Segoe UI", 9)).pack(side=tk.LEFT)

            # Progress bar using canvas
            bar_canvas = tk.Canvas(row, width=300, height=16, bg="#313244",
                                    highlightthickness=0)
            bar_canvas.pack(side=tk.LEFT, padx=10)
            bar_width = int(300 * pct / 100)
            bar_canvas.create_rectangle(0, 0, bar_width, 16, fill=ms["color"],
                                         outline="")

            tk.Label(row, text=f"{filled}/{n_patients} ({pct:.0f}%)",
                     bg="#1e1e2e", fg="#a6adc8", font=("Segoe UI", 9)).pack(side=tk.LEFT)

        # Overdue / incomplete patients
        incomplete_frame = tk.LabelFrame(self.summary_frame, text=" Incomplete Patients ",
                                         bg="#1e1e2e", fg="#cdd6f4",
                                         font=("Segoe UI", 10, "bold"), padx=10, pady=10)
        incomplete_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(5, 10))

        # Find patients with least completion
        df_copy = df.copy()
        df_copy["_filled"] = 0
        for ms in milestones:
            df_copy["_filled"] += df_copy[ms["name"]].notna().astype(int)
        df_copy["_pct"] = df_copy["_filled"] / n_milestones * 100 if n_milestones > 0 else 0

        incomplete = df_copy[df_copy["_pct"] < 100].sort_values("_pct")

        if incomplete.empty:
            tk.Label(incomplete_frame, text="✅ All patients are complete!",
                     bg="#1e1e2e", fg="#a6e3a1", font=("Segoe UI", 11)).pack(pady=10)
        else:
            cols = ("Patient", "Completed", "Missing")
            inc_tree = ttk.Treeview(incomplete_frame, columns=cols, show="headings", height=8)
            for col in cols:
                inc_tree.heading(col, text=col, anchor="w")
            inc_tree.column("Patient", width=100)
            inc_tree.column("Completed", width=100)
            inc_tree.column("Missing", width=400)

            for _, row in incomplete.iterrows():
                missing = [ms["name"] for ms in milestones if pd.isna(row.get(ms["name"]))]
                inc_tree.insert("", "end", values=(
                    row["Patient"],
                    f"{int(row['_filled'])}/{n_milestones}",
                    ", ".join(missing)
                ))

            inc_tree.pack(fill=tk.BOTH, expand=True)

    # ── MILESTONE EDITOR ────────────────────────────────────────────────

    def _edit_milestones(self):
        """Open milestone editor dialog."""
        editor = tk.Toplevel(self.win)
        editor.title("Edit Milestones")
        editor.geometry("550x450")
        editor.configure(bg="#1e1e2e")
        editor.transient(self.win)
        editor.grab_set()

        tk.Label(editor, text="⚙ Milestone Configuration",
                 font=("Segoe UI", 12, "bold"), bg="#1e1e2e", fg="#cdd6f4").pack(pady=10)

        # Listbox showing milestones
        list_frame = tk.Frame(editor, bg="#1e1e2e")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        cols = ("name", "type", "column", "color")
        ms_tree = ttk.Treeview(list_frame, columns=cols, show="headings", height=10)
        ms_tree.heading("name", text="Name", anchor="w")
        ms_tree.heading("type", text="Type", anchor="w")
        ms_tree.heading("column", text="Column Pattern", anchor="w")
        ms_tree.heading("color", text="Color", anchor="w")
        ms_tree.column("name", width=140)
        ms_tree.column("type", width=80)
        ms_tree.column("column", width=200)
        ms_tree.column("color", width=80)
        ms_tree.pack(fill=tk.BOTH, expand=True)

        def refresh_list():
            for item in ms_tree.get_children():
                ms_tree.delete(item)
            for ms in self.config.milestones:
                ms_tree.insert("", "end", iid=ms["id"],
                               values=(ms["name"], ms.get("type", "data"),
                                       ms.get("column_pattern", ""), ms.get("color", "#95a5a6")))

        refresh_list()

        # Buttons
        btn_frame = tk.Frame(editor, bg="#1e1e2e")
        btn_frame.pack(fill=tk.X, padx=15, pady=10)

        btn_style = {"font": ("Segoe UI", 9, "bold"), "cursor": "hand2", "relief": "flat",
                      "padx": 10, "pady": 5}

        def add_milestone():
            dialog = tk.Toplevel(editor)
            dialog.title("Add Milestone")
            dialog.geometry("400x300")
            dialog.configure(bg="#1e1e2e")
            dialog.transient(editor)
            dialog.grab_set()

            lbl_s = {"bg": "#1e1e2e", "fg": "#cdd6f4", "font": ("Segoe UI", 10)}
            ent_s = {"bg": "#313244", "fg": "#cdd6f4", "insertbackground": "#cdd6f4",
                      "font": ("Segoe UI", 10), "relief": "flat"}

            entries = {}
            for i, (label, default) in enumerate([
                ("Name:", ""),
                ("Column Pattern:", ""),
                ("Color:", "#95a5a6"),
            ]):
                tk.Label(dialog, text=label, **lbl_s).grid(
                    row=i, column=0, padx=15, pady=8, sticky="w")
                e = tk.Entry(dialog, width=25, **ent_s)
                e.insert(0, default)
                e.grid(row=i, column=1, padx=15, pady=8, sticky="ew")
                entries[label] = e

            type_var = tk.StringVar(value="data")
            tk.Label(dialog, text="Type:", **lbl_s).grid(
                row=3, column=0, padx=15, pady=8, sticky="w")
            type_frame = tk.Frame(dialog, bg="#1e1e2e")
            type_frame.grid(row=3, column=1, padx=15, pady=8, sticky="w")
            tk.Radiobutton(type_frame, text="Data-bound", variable=type_var, value="data",
                           bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#1e1e2e").pack(side=tk.LEFT)
            tk.Radiobutton(type_frame, text="Manual", variable=type_var, value="manual",
                           bg="#1e1e2e", fg="#cdd6f4", selectcolor="#313244",
                           activebackground="#1e1e2e").pack(side=tk.LEFT)

            dialog.columnconfigure(1, weight=1)

            def on_add():
                name = entries["Name:"].get().strip()
                if not name:
                    messagebox.showwarning("Missing", "Name is required.", parent=dialog)
                    return
                mid = name.lower().replace(" ", "_")
                col = entries["Column Pattern:"].get().strip()
                color = entries["Color:"].get().strip()
                self.config.add_milestone(mid, name, type_var.get(), col, color)
                dialog.destroy()
                refresh_list()

            tk.Button(dialog, text="Add", bg="#89b4fa", fg="#1e1e2e",
                      command=on_add, **btn_style).grid(row=4, column=1, pady=15, sticky="e")

        def remove_milestone():
            sel = ms_tree.selection()
            if not sel:
                return
            ms_id = sel[0]
            ms = next((m for m in self.config.milestones if m["id"] == ms_id), None)
            if ms and messagebox.askyesno("Remove", f"Remove '{ms['name']}'?", parent=editor):
                self.config.remove_milestone(ms_id)
                refresh_list()

        def move_up():
            sel = ms_tree.selection()
            if sel:
                self.config.move_milestone(sel[0], -1)
                refresh_list()
                ms_tree.selection_set(sel[0])

        def move_down():
            sel = ms_tree.selection()
            if sel:
                self.config.move_milestone(sel[0], 1)
                refresh_list()
                ms_tree.selection_set(sel[0])

        def reset_defaults():
            if messagebox.askyesno("Reset", "Reset to default milestones?", parent=editor):
                self.config.milestones = [m.copy() for m in DEFAULT_MILESTONES]
                self.config.save()
                refresh_list()

        tk.Button(btn_frame, text="➕ Add", bg="#89b4fa", fg="#1e1e2e",
                  command=add_milestone, **btn_style).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="🗑 Remove", bg="#f38ba8", fg="#1e1e2e",
                  command=remove_milestone, **btn_style).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="⬆", bg="#585b70", fg="white",
                  command=move_up, **btn_style).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="⬇", bg="#585b70", fg="white",
                  command=move_down, **btn_style).pack(side=tk.LEFT, padx=3)
        tk.Button(btn_frame, text="↺ Reset Defaults", bg="#585b70", fg="white",
                  command=reset_defaults, **btn_style).pack(side=tk.LEFT, padx=3)

        def on_close():
            editor.destroy()
            self._refresh_all()

        tk.Button(btn_frame, text="Done", bg="#45475a", fg="white",
                  command=on_close, **btn_style).pack(side=tk.RIGHT, padx=3)
