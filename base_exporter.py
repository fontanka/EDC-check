"""Base Exporter — shared helpers for all clinical data export modules.

Consolidates duplicated patterns across echo_export, labs_export, cvc_export:
  - Patient row lookup
  - Value validation / safe access
  - Date formatting
  - Numeric conversion
  - Template loading
  - Single/multi-patient export orchestration (xlsx or zip)
"""

import pandas as pd
import openpyxl
import logging
from io import BytesIO
import zipfile

logger = logging.getLogger(__name__)


class BaseExporter:
    """Shared export logic for all clinical data exporters."""

    def __init__(self, df_main, template_path=None, labels_map=None):
        self.df_main = df_main
        self.template_path = template_path
        self.labels_map = labels_map or {}

    # ------------------------------------------------------------------
    # Patient data access
    # ------------------------------------------------------------------

    def get_patient_row(self, patient_id):
        """Fetch single patient row from df_main by Screening #.

        Returns pandas Series or None.
        """
        rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if rows.empty:
            return None
        return rows.iloc[0]

    # ------------------------------------------------------------------
    # Value validation
    # ------------------------------------------------------------------

    @staticmethod
    def is_valid(val):
        """Return True if *val* is real data (not NaN / empty / placeholder)."""
        if pd.isna(val):
            return False
        s = str(val).strip().lower()
        return bool(s) and s not in ('', 'nan', 'none', 'n/a', 'na')

    @staticmethod
    def safe_str(row, col_name):
        """Get validated string value from *row*, or None."""
        if col_name not in row.index:
            return None
        val = row[col_name]
        if not BaseExporter.is_valid(val):
            return None
        return str(val)

    # ------------------------------------------------------------------
    # Numeric helpers
    # ------------------------------------------------------------------

    @staticmethod
    def to_number(val, try_int=True):
        """Convert *val* to int/float, return None on failure."""
        if val is None:
            return None
        val_str = str(val).strip()
        if not val_str or val_str.lower() in ('nan', 'none', 'not done'):
            return None
        try:
            if try_int and '.' not in val_str:
                return int(val_str)
            return float(val_str)
        except (ValueError, TypeError):
            return None

    @staticmethod
    def to_float(val):
        """Convert *val* to float, return None on failure."""
        if val is None or not BaseExporter.is_valid(val):
            return None
        try:
            return float(val)
        except (ValueError, TypeError):
            return None

    # ------------------------------------------------------------------
    # Date formatting
    # ------------------------------------------------------------------

    @staticmethod
    def format_date(date_val, fmt='%d-%b-%Y'):
        """Format date consistently; returns empty string on failure."""
        if not date_val or pd.isna(date_val):
            return ""
        try:
            date_str = str(date_val)
            if '|' in date_str:
                date_str = date_str.split('|')[0].strip()
            dt = pd.to_datetime(date_str)
            return dt.strftime(fmt)
        except (ValueError, TypeError):
            s = str(date_val)
            return s.split('T')[0] if 'T' in s else s

    # ------------------------------------------------------------------
    # Template loading
    # ------------------------------------------------------------------

    def load_template(self):
        """Load openpyxl Workbook from *self.template_path*.

        Returns Workbook or None on error.
        """
        if not self.template_path:
            return None
        try:
            return openpyxl.load_workbook(self.template_path)
        except Exception as e:
            logger.error("Error loading template %s: %s", self.template_path, e)
            return None

    # ------------------------------------------------------------------
    # Export orchestration
    # ------------------------------------------------------------------

    def generate_export(self, patient_ids, filename_fmt=None, **process_kwargs):
        """Generate export — single xlsx for one patient, ZIP for multiple.

        Subclasses must implement ``process_patient(patient_id, **kwargs)``
        returning bytes or None.

        *filename_fmt*: callable(pid) -> str filename inside ZIP.
                        Default: ``"{pid}.xlsx"``.

        Returns tuple ``(bytes_data, extension_str, first_pid_or_none)``.
        """
        if filename_fmt is None:
            filename_fmt = lambda pid: f"{pid}.xlsx"

        if len(patient_ids) == 1:
            data = self.process_patient(patient_ids[0], **process_kwargs)
            return (data, 'xlsx', patient_ids[0])

        zip_buffer = BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zf:
            for pid in patient_ids:
                data = self.process_patient(pid, **process_kwargs)
                if data:
                    zf.writestr(filename_fmt(pid), data)
        return (zip_buffer.getvalue(), 'zip', None)
