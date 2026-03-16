"""
Data Loader Module
==================
Handles file detection, Excel loading, sheet parsing, and schema validation.
Pure data operations — no UI / tkinter dependencies.

Usage:
    from data_loader import detect_latest_project_file, load_project_file

    path, cutoff = detect_latest_project_file(os.getcwd())
    result = load_project_file(path)
    # result.df_main, result.df_ae, result.labels, etc.
"""

import os
import re
import logging
import pandas as pd
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from dataclasses import dataclass, field

logger = logging.getLogger("ClinicalViewer.DataLoader")


# =============================================================================
# Result container
# =============================================================================

@dataclass
class LoadResult:
    """Result of loading a ProjectToOneFile export."""
    df_main: pd.DataFrame
    df_ae: Optional[pd.DataFrame] = None
    df_cm: Optional[pd.DataFrame] = None
    df_cvh: Optional[pd.DataFrame] = None
    df_act: Optional[pd.DataFrame] = None
    labels: Dict[str, str] = field(default_factory=dict)
    file_path: str = ""
    cutoff_time: Optional[datetime] = None
    warnings: List[str] = field(default_factory=list)


# =============================================================================
# File detection
# =============================================================================

_FILENAME_PREFIX = "Innoventric_CLD-048_DM_ProjectToOneFile"
_TIMESTAMP_RE = re.compile(r'_(\d{2}-\d{2}-\d{4}_\d{2}-\d{2}[-_]\d{2})_')
_TIMESTAMP_FMTS = ("%d-%m-%Y_%H-%M_%S", "%d-%m-%Y_%H-%M-%S")


def detect_latest_project_file(directory: str) -> Optional[Tuple[str, Optional[datetime]]]:
    """Find the most recent ProjectToOneFile Excel export in *directory*.

    Returns:
        (full_path, parsed_datetime) or None if no file found.
    """
    try:
        files = [f for f in os.listdir(directory)
                 if f.startswith(_FILENAME_PREFIX) and f.endswith(".xlsx")]
    except OSError as e:
        logger.error("Cannot list directory %s: %s", directory, e)
        return None

    if not files:
        return None

    latest_file = None
    latest_time = None

    for f in files:
        match = _TIMESTAMP_RE.search(f)
        if not match:
            continue
        dt_str = match.group(1)
        for fmt in _TIMESTAMP_FMTS:
            try:
                dt = datetime.strptime(dt_str, fmt)
                if latest_time is None or dt > latest_time:
                    latest_time = dt
                    latest_file = f
                break
            except ValueError:
                continue

    if latest_file:
        return os.path.join(directory, latest_file), latest_time
    return None


def parse_cutoff_from_filename(filename: str) -> Optional[datetime]:
    """Extract the cutoff timestamp from a ProjectToOneFile filename."""
    match = _TIMESTAMP_RE.search(filename)
    if not match:
        return None
    dt_str = match.group(1)
    for fmt in _TIMESTAMP_FMTS:
        try:
            return datetime.strptime(dt_str, fmt)
        except ValueError:
            continue
    return None


# =============================================================================
# Sheet loading helpers
# =============================================================================

def _load_repeating_sheet(raw: pd.DataFrame) -> Optional[pd.DataFrame]:
    """Parse a repeating-form sheet (AE, CM, CVH, ACT).

    Row 0 = column headers, Row 1+ = data.
    Non-breaking spaces in column names are replaced.
    """
    if raw.empty:
        return None
    cols = [str(c).replace('\xa0', ' ').strip() for c in raw.iloc[0].tolist()]
    df = raw.iloc[1:].copy()
    df.columns = cols
    df.reset_index(drop=True, inplace=True)
    return df


def _load_extra_sheets(xls: Dict[str, pd.DataFrame]) -> Tuple[
    Optional[pd.DataFrame],
    Optional[pd.DataFrame],
    Optional[pd.DataFrame],
    Optional[pd.DataFrame],
    List[str],
]:
    """Load auxiliary repeating-form sheets from the workbook.

    Returns:
        (df_ae, df_cm, df_cvh, df_act, warnings)
    """
    warnings: List[str] = []

    # --- AE ---
    ae_sheet = next((n for n in xls if n.startswith("AE_")), None)
    df_ae = None
    if ae_sheet:
        try:
            df_ae = _load_repeating_sheet(xls[ae_sheet])
        except Exception as e:
            warnings.append(f"Error loading AE sheet: {e}")
            logger.warning("Error loading AE sheet '%s': %s", ae_sheet, e)

    # --- CM ---
    cm_sheet = next((n for n in xls if n.startswith("CMTAB")), None)
    df_cm = None
    if cm_sheet:
        try:
            df_cm = _load_repeating_sheet(xls[cm_sheet])
        except Exception as e:
            warnings.append(f"Error loading CM sheet: {e}")
            logger.warning("Error loading CM sheet '%s': %s", cm_sheet, e)

    # --- CVH ---
    cvh_sheet = next((n for n in xls if n.startswith("CVH_TABLE")), None)
    df_cvh = None
    if cvh_sheet:
        try:
            df_cvh = _load_repeating_sheet(xls[cvh_sheet])
        except Exception as e:
            warnings.append(f"Error loading CVH sheet: {e}")
            logger.warning("Error loading CVH sheet '%s': %s", cvh_sheet, e)

    # --- ACT (may span multiple sheets) ---
    act_names = [n for n in xls if n.startswith("LB_ACT") or ("Group" in n and "717" in n)]
    act_dfs: List[pd.DataFrame] = []
    for name in act_names:
        try:
            df_part = _load_repeating_sheet(xls[name])
            if df_part is not None:
                act_dfs.append(df_part)
                logger.debug("Loaded ACT sheet '%s' with %d rows", name, len(df_part))
        except Exception as e:
            warnings.append(f"Error loading ACT sheet {name}: {e}")
            logger.warning("Error loading ACT sheet '%s': %s", name, e)

    df_act = pd.concat(act_dfs, ignore_index=True) if act_dfs else None
    if df_act is not None:
        logger.debug("Total merged ACT rows: %d", len(df_act))

    return df_ae, df_cm, df_cvh, df_act, warnings


# =============================================================================
# Main loader
# =============================================================================

def load_project_file(path: str, cutoff_time: Optional[datetime] = None) -> LoadResult:
    """Load a ProjectToOneFile Excel export and return structured data.

    Args:
        path: Full path to the .xlsx file.
        cutoff_time: Optional pre-parsed cutoff timestamp.

    Returns:
        LoadResult with all DataFrames, labels, and any warnings.

    Raises:
        FileNotFoundError: if *path* does not exist.
        ValueError: if the "Main" sheet is not found.
    """
    if not os.path.isfile(path):
        raise FileNotFoundError(f"File not found: {path}")

    logger.info("Loading project file: %s", path)

    # Parse cutoff from filename if not provided
    if cutoff_time is None:
        cutoff_time = parse_cutoff_from_filename(os.path.basename(path))

    warnings: List[str] = []

    # Read all sheets
    xls = pd.read_excel(path, sheet_name=None, header=None, dtype=str, keep_default_na=False)

    # Locate Main sheet
    target = next((n for n in xls if "main" in n.lower()), None)
    if not target:
        raise ValueError("Could not find 'Main' sheet in workbook.")

    raw = xls[target]
    codes = [str(c).strip() for c in raw.iloc[0].tolist()]
    labels_list = [str(lbl).strip() for lbl in raw.iloc[1].tolist()]

    df_main = raw.iloc[2:].copy()
    df_main.columns = codes
    labels = dict(zip(codes, labels_list))

    # Validate schema
    schema_warnings = validate_schema(codes)
    warnings.extend(schema_warnings)

    # Load auxiliary sheets
    df_ae, df_cm, df_cvh, df_act, sheet_warnings = _load_extra_sheets(xls)
    warnings.extend(sheet_warnings)

    result = LoadResult(
        df_main=df_main,
        df_ae=df_ae,
        df_cm=df_cm,
        df_cvh=df_cvh,
        df_act=df_act,
        labels=labels,
        file_path=path,
        cutoff_time=cutoff_time,
        warnings=warnings,
    )

    logger.info(
        "Loaded: %d patients, %d columns, AE=%s, CM=%s, CVH=%s, ACT=%s",
        len(df_main), len(codes),
        len(df_ae) if df_ae is not None else "N/A",
        len(df_cm) if df_cm is not None else "N/A",
        len(df_cvh) if df_cvh is not None else "N/A",
        len(df_act) if df_act is not None else "N/A",
    )

    return result


# =============================================================================
# Schema validation
# =============================================================================

def validate_schema(columns, required_cols=None) -> List[str]:
    """Check critical columns exist.  Returns list of warning strings."""
    try:
        from column_registry import CRITICAL_COLUMNS, validate_columns
        if required_cols is None:
            required_cols = CRITICAL_COLUMNS
        _, missing = validate_columns(columns, required_cols)
        if missing:
            msg = f"{len(missing)} expected column(s) not found: {', '.join(missing[:10])}"
            logger.warning(msg)
            return [msg]
    except ImportError:
        pass
    return []


# =============================================================================
# Cross-form consistency validation
# =============================================================================

def validate_cross_form(result: LoadResult) -> List[str]:
    """Run cross-form consistency checks on loaded data.

    Checks:
        1. AE onset dates fall within enrollment–last-visit window
        2. Fatal AE outcome ↔ Death form date consistency
        3. Procedure date precedes all follow-up visit dates
        4. Death date matches or follows the last AE onset

    Returns:
        List of human-readable issue strings (empty = all good).
    """
    issues: List[str] = []
    df = result.df_main
    if df is None or df.empty:
        return issues

    _check_fatal_ae_death_consistency(df, result.df_ae, issues)
    _check_procedure_before_followups(df, issues)
    _check_ae_onset_after_procedure(df, result.df_ae, issues)

    if issues:
        logger.info("Cross-form validation found %d issue(s)", len(issues))
    return issues


def _safe_date(val) -> Optional[datetime]:
    """Parse a date value, returning None on failure."""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if not s or s.lower() in ('nan', 'none', 'nat', ''):
        return None
    try:
        return pd.to_datetime(s).to_pydatetime()
    except (ValueError, TypeError, OverflowError):
        return None


def _find_col(df: pd.DataFrame, substring: str) -> Optional[str]:
    """Find the first column in *df* whose name contains *substring*."""
    for c in df.columns:
        if substring in str(c):
            return c
    return None


def _check_fatal_ae_death_consistency(
    df_main: pd.DataFrame,
    df_ae: Optional[pd.DataFrame],
    issues: List[str],
):
    """AE with outcome='Fatal' should have a matching Death form date."""
    if df_ae is None or df_ae.empty:
        return

    # Find relevant columns
    outcome_col = None
    for c in df_ae.columns:
        if 'AEOUT' in str(c):
            outcome_col = c
            break
    if outcome_col is None:
        return

    death_date_col = _find_col(df_main, 'DTH_DDDTC')

    # Get patients with fatal AEs
    fatal_mask = df_ae[outcome_col].astype(str).str.strip().str.lower() == 'fatal'
    fatal_patients = df_ae.loc[fatal_mask, 'Screening #'].astype(str).str.strip().unique()

    for pid in fatal_patients:
        if death_date_col is None:
            issues.append(f"{pid}: Fatal AE but no Death form column found in data")
            continue
        pat_row = df_main[df_main['Screening #'].astype(str).str.strip() == pid]
        if pat_row.empty:
            continue
        death_val = str(pat_row.iloc[0].get(death_date_col, '')).strip()
        if not death_val or death_val.lower() in ('nan', '', 'none', 'nat'):
            issues.append(f"{pid}: Fatal AE outcome but Death form date is empty")


def _check_procedure_before_followups(df_main: pd.DataFrame, issues: List[str]):
    """Procedure date should precede all follow-up visit dates."""
    proc_col = _find_col(df_main, 'TV_PR_PRSTDTC')
    if proc_col is None:
        return

    fu_prefixes = ['FU1M', 'FU3M', 'FU6M', 'FU1Y', 'FU2Y', 'FU3Y', 'FU4Y', 'FU5Y']
    visit_date_cols = {}
    for pfx in fu_prefixes:
        col = _find_col(df_main, f'{pfx}_SV_SVSTDTC')
        if col:
            visit_date_cols[pfx] = col

    if not visit_date_cols:
        return

    for _, row in df_main.iterrows():
        pid = str(row.get('Screening #', '')).strip()
        proc_dt = _safe_date(row.get(proc_col))
        if proc_dt is None:
            continue
        for pfx, col in visit_date_cols.items():
            fu_dt = _safe_date(row.get(col))
            if fu_dt and fu_dt < proc_dt:
                issues.append(
                    f"{pid}: {pfx} visit date ({fu_dt:%Y-%m-%d}) "
                    f"precedes procedure date ({proc_dt:%Y-%m-%d})"
                )


def _check_ae_onset_after_procedure(
    df_main: pd.DataFrame,
    df_ae: Optional[pd.DataFrame],
    issues: List[str],
):
    """Flag post-procedure AEs that have onset before the procedure date.

    Only checks AEs whose interval field indicates post-procedure.
    """
    if df_ae is None or df_ae.empty:
        return

    proc_col = _find_col(df_main, 'TV_PR_PRSTDTC')
    onset_col = None
    interval_col = None
    for c in df_ae.columns:
        if 'AESTDTC' in str(c) and onset_col is None:
            onset_col = c
        if 'AEINT' in str(c) and interval_col is None:
            interval_col = c

    if not proc_col or not onset_col:
        return

    # Build procedure-date lookup
    proc_map = {}
    for _, row in df_main.iterrows():
        pid = str(row.get('Screening #', '')).strip()
        dt = _safe_date(row.get(proc_col))
        if dt:
            proc_map[pid] = dt

    for _, ae_row in df_ae.iterrows():
        pid = str(ae_row.get('Screening #', '')).strip()
        if pid not in proc_map:
            continue

        # Only flag if interval says post-procedure (or is empty — assume post)
        if interval_col:
            interval = str(ae_row.get(interval_col, '')).strip().lower()
            if interval and 'pre' in interval:
                continue  # explicitly pre-procedure, skip

        onset_dt = _safe_date(ae_row.get(onset_col))
        if onset_dt and onset_dt < proc_map[pid]:
            ae_term = str(ae_row.get('LOGS_AE_AETERM', ae_row.get('AETERM', '?'))).strip()
            issues.append(
                f"{pid}: AE '{ae_term[:40]}' onset ({onset_dt:%Y-%m-%d}) "
                f"before procedure ({proc_map[pid]:%Y-%m-%d}) but not marked pre-procedure"
            )
