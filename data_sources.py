"""
Data Sources Manager
====================
Tracks all input files (Project, Modular, CRF Status) with metadata.
Provides a UI for viewing, reloading, and adding custom data sources.
"""

import os
import json
import glob
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from datetime import datetime
from typing import Optional, Dict, List, Callable
import logging

logger = logging.getLogger("DataSources")

CONFIG_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data_sources_config.json")


class DataSource:
    """Represents a single data source file."""

    def __init__(self, source_type: str, label: str, file_pattern: str,
                 search_dir: str = ".", filepath: Optional[str] = None,
                 is_builtin: bool = True):
        self.source_type = source_type      # Unique key e.g. "project", "modular", "crf_status"
        self.label = label                  # Display name
        self.file_pattern = file_pattern    # Glob pattern for auto-detection
        self.search_dir = search_dir        # Directory to search in
        self.filepath: Optional[str] = filepath
        self.is_builtin = is_builtin        # Built-in vs user-added
        self.is_loaded = False
        self.load_time: Optional[datetime] = None
        self.file_date: Optional[datetime] = None  # File modification time
        self.file_size: int = 0
        self.error: Optional[str] = None

    def detect_file(self) -> Optional[str]:
        """Auto-detect the latest matching file."""
        search_path = os.path.join(self.search_dir, self.file_pattern)
        candidates = [f for f in glob.glob(search_path)
                      if not os.path.basename(f).startswith("~$")]
        if candidates:
            latest = max(candidates, key=os.path.getmtime)
            return latest
        return None

    def update_metadata(self):
        """Update file metadata from the current filepath."""
        if self.filepath and os.path.isfile(self.filepath):
            stat = os.stat(self.filepath)
            self.file_date = datetime.fromtimestamp(stat.st_mtime)
            self.file_size = stat.st_size
        else:
            self.file_date = None
            self.file_size = 0

    def to_dict(self) -> dict:
        """Serialize for JSON config (custom sources only)."""
        return {
            "source_type": self.source_type,
            "label": self.label,
            "file_pattern": self.file_pattern,
            "search_dir": self.search_dir,
            "filepath": self.filepath,
            "is_builtin": self.is_builtin,
        }

    @classmethod
    def from_dict(cls, data: dict) -> "DataSource":
        return cls(
            source_type=data["source_type"],
            label=data["label"],
            file_pattern=data.get("file_pattern", "*"),
            search_dir=data.get("search_dir", "."),
            filepath=data.get("filepath"),
            is_builtin=data.get("is_builtin", False),
        )


class DataSourceManager:
    """Manages all data sources and their configuration."""

    def __init__(self, app_dir: str = "."):
        self.app_dir = app_dir
        self.sources: Dict[str, DataSource] = {}
        self._init_builtin_sources()
        self._load_custom_sources()

    def _init_builtin_sources(self):
        """Register the 3 built-in source types."""
        verified_dir = os.path.join(self.app_dir, "verified")

        self.sources["project"] = DataSource(
            source_type="project",
            label="Project File",
            file_pattern="Innoventric_CLD-048_DM_ProjectToOneFile*.xlsx",
            search_dir=self.app_dir,
        )
        self.sources["modular"] = DataSource(
            source_type="modular",
            label="Modular Export",
            file_pattern="*Modular*.xlsx",
            search_dir=verified_dir if os.path.isdir(verified_dir) else self.app_dir,
        )
        self.sources["crf_status"] = DataSource(
            source_type="crf_status",
            label="CRF Status History",
            file_pattern="*CrfStatusHistory*.xlsx",
            search_dir=verified_dir if os.path.isdir(verified_dir) else self.app_dir,
        )

    def _load_custom_sources(self):
        """Load user-added sources from config file."""
        if os.path.isfile(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                for item in data.get("custom_sources", []):
                    src = DataSource.from_dict(item)
                    if src.source_type not in self.sources:
                        self.sources[src.source_type] = src
            except Exception as e:
                logger.warning(f"Failed to load custom sources config: {e}")

    def _save_custom_sources(self):
        """Save user-added sources to config file."""
        custom = [s.to_dict() for s in self.sources.values() if not s.is_builtin]
        try:
            with open(CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump({"custom_sources": custom}, f, indent=2)
        except Exception as e:
            logger.warning(f"Failed to save custom sources config: {e}")

    def add_custom_source(self, source_type: str, label: str,
                          file_pattern: str, search_dir: str) -> DataSource:
        """Add a user-defined data source."""
        src = DataSource(
            source_type=source_type,
            label=label,
            file_pattern=file_pattern,
            search_dir=search_dir,
            is_builtin=False,
        )
        self.sources[source_type] = src
        self._save_custom_sources()
        return src

    def remove_custom_source(self, source_type: str):
        """Remove a user-defined source (cannot remove built-in)."""
        src = self.sources.get(source_type)
        if src and not src.is_builtin:
            del self.sources[source_type]
            self._save_custom_sources()

    def register_loaded_file(self, source_type: str, filepath: str):
        """Called by the app after successfully loading a file."""
        if source_type in self.sources:
            src = self.sources[source_type]
            src.filepath = filepath
            src.is_loaded = True
            src.load_time = datetime.now()
            src.error = None
            src.update_metadata()

    def register_error(self, source_type: str, error_msg: str):
        """Called by the app when a file fails to load."""
        if source_type in self.sources:
            self.sources[source_type].is_loaded = False
            self.sources[source_type].error = error_msg

    def detect_all(self):
        """Auto-detect files for all sources that don't have one set."""
        for src in self.sources.values():
            if not src.filepath:
                detected = src.detect_file()
                if detected:
                    src.filepath = detected
                    src.update_metadata()

    def get_all(self) -> List[DataSource]:
        """Return all sources in display order (built-in first)."""
        builtin = [s for s in self.sources.values() if s.is_builtin]
        custom = [s for s in self.sources.values() if not s.is_builtin]
        return builtin + custom


class DataSourcesWindow:
    """Toplevel window displaying all data sources with management controls."""

    def __init__(self, parent: tk.Tk, manager: DataSourceManager,
                 reload_callback: Optional[Callable] = None):
        self.manager = manager
        self.reload_callback = reload_callback

        self.win = tk.Toplevel(parent)
        self.win.title("Data Sources")
        self.win.geometry("820x450")
        self.win.configure(bg="#1e1e2e")
        self.win.resizable(True, True)

        self._build_ui()
        self._refresh_table()

        self.win.transient(parent)
        self.win.grab_set()
        self.win.focus_force()

    def _build_ui(self):
        # Header
        hdr = tk.Frame(self.win, bg="#1e1e2e")
        hdr.pack(fill=tk.X, padx=15, pady=(15, 5))
        tk.Label(hdr, text="📂 Data Sources", font=("Segoe UI", 14, "bold"),
                 bg="#1e1e2e", fg="#cdd6f4").pack(side=tk.LEFT)

        # Treeview
        tree_frame = tk.Frame(self.win, bg="#1e1e2e")
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=5)

        columns = ("type", "file", "date", "size", "status")
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings",
                                 height=8, selectmode="browse")

        self.tree.heading("type", text="Source Type", anchor="w")
        self.tree.heading("file", text="File", anchor="w")
        self.tree.heading("date", text="File Date", anchor="w")
        self.tree.heading("size", text="Size", anchor="w")
        self.tree.heading("status", text="Status", anchor="w")

        self.tree.column("type", width=140, minwidth=100)
        self.tree.column("file", width=300, minwidth=150)
        self.tree.column("date", width=140, minwidth=100)
        self.tree.column("size", width=80, minwidth=60)
        self.tree.column("status", width=100, minwidth=80)

        # Style
        style = ttk.Style()
        style.configure("DS.Treeview", background="#313244", foreground="#cdd6f4",
                         fieldbackground="#313244", font=("Segoe UI", 10),
                         rowheight=28)
        style.configure("DS.Treeview.Heading", background="#45475a",
                         foreground="#cdd6f4", font=("Segoe UI", 10, "bold"))
        self.tree.configure(style="DS.Treeview")

        scrollbar = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Tags for status colors
        self.tree.tag_configure("loaded", foreground="#a6e3a1")     # Green
        self.tree.tag_configure("not_loaded", foreground="#fab387") # Orange
        self.tree.tag_configure("error", foreground="#f38ba8")      # Red
        self.tree.tag_configure("custom", foreground="#89b4fa")     # Blue

        # Button bar
        btn_frame = tk.Frame(self.win, bg="#1e1e2e")
        btn_frame.pack(fill=tk.X, padx=15, pady=(5, 15))

        btn_style = {"font": ("Segoe UI", 10, "bold"), "cursor": "hand2",
                      "relief": "flat", "padx": 12, "pady": 6}

        tk.Button(btn_frame, text="📁 Browse", bg="#585b70", fg="white",
                  command=self._browse_file, **btn_style).pack(side=tk.LEFT, padx=(0, 5))

        tk.Button(btn_frame, text="🔄 Reload Selected", bg="#585b70", fg="white",
                  command=self._reload_selected, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="🔄 Reload All", bg="#585b70", fg="white",
                  command=self._reload_all, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="➕ Add Source", bg="#89b4fa", fg="#1e1e2e",
                  command=self._add_source, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="🗑 Remove", bg="#f38ba8", fg="#1e1e2e",
                  command=self._remove_source, **btn_style).pack(side=tk.LEFT, padx=5)

        tk.Button(btn_frame, text="Close", bg="#45475a", fg="white",
                  command=self.win.destroy, **btn_style).pack(side=tk.RIGHT)

    def _refresh_table(self):
        """Refresh the treeview with current source data."""
        for item in self.tree.get_children():
            self.tree.delete(item)

        for src in self.manager.get_all():
            src.update_metadata()

            filename = os.path.basename(src.filepath) if src.filepath else "—"
            date_str = src.file_date.strftime("%Y-%m-%d %H:%M") if src.file_date else "—"

            if src.file_size > 0:
                if src.file_size > 1_048_576:
                    size_str = f"{src.file_size / 1_048_576:.1f} MB"
                else:
                    size_str = f"{src.file_size / 1024:.0f} KB"
            else:
                size_str = "—"

            if src.error:
                status = "❌ Error"
                tag = "error"
            elif src.is_loaded:
                status = "✅ Loaded"
                tag = "loaded"
            else:
                status = "⬚ Not loaded"
                tag = "not_loaded"

            # Add custom tag for user sources
            tags = (tag,)
            if not src.is_builtin:
                tags = (tag, "custom")

            self.tree.insert("", "end", iid=src.source_type,
                             values=(src.label, filename, date_str, size_str, status),
                             tags=tags)

    def _get_selected_source(self) -> Optional[DataSource]:
        """Get the currently selected source."""
        sel = self.tree.selection()
        if sel:
            return self.manager.sources.get(sel[0])
        return None

    def _browse_file(self):
        """Browse for a file and assign it to the selected source."""
        src = self._get_selected_source()
        if not src:
            messagebox.showinfo("Select Source", "Please select a source type first.",
                                parent=self.win)
            return

        filepath = filedialog.askopenfilename(
            title=f"Select file for {src.label}",
            filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"),
                       ("All Files", "*.*")],
            initialdir=src.search_dir,
            parent=self.win
        )
        if filepath:
            src.filepath = filepath
            src.update_metadata()
            self._refresh_table()

            if self.reload_callback:
                self.reload_callback(src.source_type, filepath)

    def _reload_selected(self):
        """Reload the selected source."""
        src = self._get_selected_source()
        if not src:
            messagebox.showinfo("Select Source", "Please select a source to reload.",
                                parent=self.win)
            return

        if not src.filepath:
            detected = src.detect_file()
            if detected:
                src.filepath = detected
            else:
                messagebox.showwarning("No File", f"No file found for {src.label}.",
                                       parent=self.win)
                return

        if self.reload_callback:
            self.reload_callback(src.source_type, src.filepath)
        self._refresh_table()

    def _reload_all(self):
        """Reload all sources."""
        for src in self.manager.get_all():
            if not src.filepath:
                detected = src.detect_file()
                if detected:
                    src.filepath = detected

            if src.filepath and self.reload_callback:
                self.reload_callback(src.source_type, src.filepath)

        self._refresh_table()

    def _add_source(self):
        """Open dialog to add a custom data source."""
        dialog = tk.Toplevel(self.win)
        dialog.title("Add Data Source")
        dialog.geometry("420x280")
        dialog.configure(bg="#1e1e2e")
        dialog.transient(self.win)
        dialog.grab_set()

        lbl_style = {"bg": "#1e1e2e", "fg": "#cdd6f4", "font": ("Segoe UI", 10)}
        entry_style = {"bg": "#313244", "fg": "#cdd6f4", "insertbackground": "#cdd6f4",
                        "font": ("Segoe UI", 10), "relief": "flat"}

        fields = {}
        for i, (label, default) in enumerate([
            ("Source Name:", ""),
            ("File Pattern:", "*.xlsx"),
            ("Search Directory:", os.getcwd()),
        ]):
            tk.Label(dialog, text=label, **lbl_style).grid(
                row=i, column=0, padx=15, pady=8, sticky="w")
            entry = tk.Entry(dialog, width=30, **entry_style)
            entry.insert(0, default)
            entry.grid(row=i, column=1, padx=15, pady=8, sticky="ew")
            fields[label] = entry

        # Browse button for directory
        def browse_dir():
            d = filedialog.askdirectory(parent=dialog)
            if d:
                fields["Search Directory:"].delete(0, tk.END)
                fields["Search Directory:"].insert(0, d)

        tk.Button(dialog, text="📁", command=browse_dir,
                  bg="#585b70", fg="white", relief="flat").grid(
            row=2, column=2, padx=5, pady=8)

        dialog.columnconfigure(1, weight=1)

        def on_add():
            name = fields["Source Name:"].get().strip()
            pattern = fields["File Pattern:"].get().strip()
            search_dir = fields["Search Directory:"].get().strip()

            if not name:
                messagebox.showwarning("Missing Name", "Please enter a source name.",
                                       parent=dialog)
                return

            # Generate a safe key
            key = name.lower().replace(" ", "_")
            if key in self.manager.sources:
                messagebox.showwarning("Duplicate",
                                       f"Source '{name}' already exists.", parent=dialog)
                return

            self.manager.add_custom_source(key, name, pattern, search_dir)
            dialog.destroy()
            self._refresh_table()

        btn_frame = tk.Frame(dialog, bg="#1e1e2e")
        btn_frame.grid(row=3, column=0, columnspan=3, pady=20)

        tk.Button(btn_frame, text="Add", bg="#89b4fa", fg="#1e1e2e",
                  font=("Segoe UI", 10, "bold"), padx=20, pady=6,
                  relief="flat", command=on_add).pack(side=tk.LEFT, padx=10)
        tk.Button(btn_frame, text="Cancel", bg="#45475a", fg="white",
                  font=("Segoe UI", 10, "bold"), padx=20, pady=6,
                  relief="flat", command=dialog.destroy).pack(side=tk.LEFT, padx=10)

    def _remove_source(self):
        """Remove a custom data source."""
        src = self._get_selected_source()
        if not src:
            messagebox.showinfo("Select Source", "Please select a source to remove.",
                                parent=self.win)
            return

        if src.is_builtin:
            messagebox.showinfo("Built-in Source",
                                "Cannot remove built-in sources. You can only remove custom sources.",
                                parent=self.win)
            return

        if messagebox.askyesno("Confirm Remove",
                               f"Remove source '{src.label}'?", parent=self.win):
            self.manager.remove_custom_source(src.source_type)
            self._refresh_table()
