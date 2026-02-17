"""Data Matrix Builder — extracts column routing, type detection, and
repeating-form data parsing from clinical_viewer1.show_data_matrix().

The DataMatrixBuilder class handles:
  - Column type detection (AE, CM, MH, HFH, HMEH, CVC, CVH, ACT)
  - Repeating-form data extraction from Main sheet pipe-delimited fields
  - Generic lab/result matrix construction with visit-date resolution
  - Pivot table + Toplevel tree display
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import pandas as pd
import re
import logging
from datetime import datetime

from config import VISIT_MAP, VISIT_SCHEDULE
from cvc_export import CVCExporter

logger = logging.getLogger("ClinicalViewer")


# ---------------------------------------------------------------------------
# Column type classifier
# ---------------------------------------------------------------------------

_COL_TYPES = {
    'ae':   lambda c: any(x in c for x in [
                "AETERM", "AESTDTC", "AEENDTC", "AESER", "AESEV",
                "AEOUT", "AEREL", "AEDECOD"]),
    'cm':   lambda c: any(x in c for x in [
                "CMTRT", "CMDOSE", "CMDOSFRQ", "CMSTDTC", "CMSTDAT",
                "CMENDTC", "CMENDAT", "CMINDC", "CMROUTE", "CMDOSU",
                "CMONGO", "CMREF"]),
    'mh':   lambda c: "_MH_" in c and any(x in c for x in [
                "MHTERM", "MHSTDTC", "MHENDTC", "MHONGO", "MHCAT",
                "MHBODSYS", "MHOCCUR"]),
    'hfh':  lambda c: "_HFH_" in c and any(x in c for x in [
                "HOSTDTC", "HOTERM", "HONUM", "HOOCCUR", "HODESC"]),
    'hmeh': lambda c: "_HMEH_" in c or "HMEH" in c,
    'cvc':  lambda c: "_CVC_" in c and any(x in c for x in [
                "CVORRES", "FAORRES", "PRSTDTC"]),
    'cvh':  lambda c: "_CVH_" in c and any(x in c for x in [
                "PRSTDTC", "PRCAT", "PRTRT", "PROCCUR"]),
    'act':  lambda c: "_LB_ACT_" in c and any(x in c for x in [
                "LBORRES", "LBTIM", "CMTIM", "CMDOS", "CMSTAT", "LBSTAT"]),
}


def classify_column(col_name):
    """Return the type key ('ae', 'cm', …) or None for a CRF column name."""
    if "LBREF" in col_name or "PRREF" in col_name:
        return 'ae_ref'
    for key, test_fn in _COL_TYPES.items():
        if test_fn(col_name):
            return key
    return None


# ---------------------------------------------------------------------------
# Helpers (promoted from nested functions)
# ---------------------------------------------------------------------------

def parse_time_minutes(t):
    """Parse an HH:MM time string into minutes-since-midnight for sorting."""
    try:
        parts = str(t).split(':')
        if len(parts) >= 2:
            return int(parts[0]) * 60 + int(parts[1])
        return 9999
    except ValueError:
        return 9999


def try_parse_date(d_str):
    """Parse a date string for column sorting; returns datetime.max on failure."""
    try:
        base_d = re.sub(r" \(\d+\)$", "", d_str)
        return datetime.fromisoformat(base_d) if len(base_d) > 10 else datetime.max
    except ValueError:
        return datetime.max


# ---------------------------------------------------------------------------
# Repeating-form handlers
# ---------------------------------------------------------------------------

def _handle_ae(app, pat, row):
    """Show AE matrix from df_ae sheet."""
    if app.df_ae is not None and not app.df_ae.empty:
        pat_aes = app.df_ae[
            app.df_ae['Screening #'].astype(str).str.contains(
                pat.replace('-', '-'), na=False)]
        if not pat_aes.empty:
            app.matrix_display.show_ae_matrix(pat_aes, pat)
            return True
        messagebox.showinfo("Info", "No adverse events found for this patient in AE sheet.")
        return True
    messagebox.showinfo("Info", "AE sheet not loaded or empty.")
    return True


def _handle_cm(app, pat, row):
    """Show CM matrix — prefer dedicated sheet, fall back to Main sheet parsing."""
    if app.df_cm is not None and not app.df_cm.empty:
        pat_cms = app.df_cm[
            app.df_cm['Screening #'].astype(str).str.contains(
                pat.replace('-', '-'), na=False)]
        if not pat_cms.empty:
            app.matrix_display.show_cm_matrix(pat_cms, pat)
            return True

    # Parse from Main sheet LOGS_CM columns
    cm_cols = {
        'CMTRT': 'Medication', 'CMDOSE': 'Dose', 'CMDOSU': 'Dose Unit',
        'CMROUTE': 'Route', 'CMINDC': 'Indication', 'CMSTDTC': 'Start Date',
        'CMENDTC': 'End Date', 'CMENDAT': 'End Date', 'CMONGO': 'Ongoing',
        'CMDOSFRQ': 'Frequency', 'CMDOSFRQ_OTH': 'Frequency (Other)',
    }
    logs_cm_cols = {}
    for col in app.df_main.columns:
        col_str = str(col)
        if 'LOGS_CM_' in col_str or (col_str.startswith('LOGS_') and '_CM_' in col_str):
            for cm_key, display_name in cm_cols.items():
                if cm_key in col_str:
                    logs_cm_cols[display_name] = col_str
                    break

    if not logs_cm_cols:
        messagebox.showinfo("Info", "No CM columns found in data.")
        return True

    med_col = logs_cm_cols.get('Medication')
    if not med_col or pd.isna(row.get(med_col)):
        messagebox.showinfo("Info", "No medications found for this patient.")
        return True

    med_vals = [m.strip() for m in str(row[med_col]).split('|')
                if m.strip() and m.strip().lower() != 'nan']
    if not med_vals:
        messagebox.showinfo("Info", "No medications found for this patient.")
        return True

    cm_data = []
    for i, med in enumerate(med_vals):
        record = {'CM #': str(i + 1), 'Medication': med}
        for display_name, col_name in logs_cm_cols.items():
            if display_name == 'Medication':
                continue
            col_val = row.get(col_name, '')
            if pd.notna(col_val):
                vals = [v.strip() for v in str(col_val).split('|')]
                val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                if 'Date' in display_name and val:
                    if 'T' in val:
                        val = val.split('T')[0]
                    val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
                record[display_name] = val

        if record.get('Ongoing', '').lower() in ['yes', 'y', '1', 'true', 'checked']:
            record['End Date'] = 'Ongoing'
        if record.get('Frequency', '').lower() == 'other':
            freq_other = record.get('Frequency (Other)', '')
            if freq_other:
                record['Frequency'] = freq_other

        # Calculate Daily Dose
        try:
            dose_str = record.get('Dose', '')
            freq_str = record.get('Frequency', '')
            freq_oth = record.get('Frequency (Other)', '')
            unit_str = record.get('Dose Unit', '')
            if dose_str and dose_str.lower() not in ['nan', 'none', '']:
                single_dose = float(dose_str)
                multiplier, freq_note, override_dose = app.matrix_display.parse_frequency_multiplier(freq_str, freq_oth)
                if override_dose is not None:
                    daily = override_dose
                elif multiplier is not None:
                    daily = single_dose * multiplier
                else:
                    daily = None
                if daily is not None:
                    daily_str = str(int(daily)) if daily == int(daily) else f"{daily:.1f}"
                    if unit_str and unit_str.lower() not in ['nan', 'none', '']:
                        if 'milligram' in unit_str.lower():
                            unit_str = 'mg'
                        daily_str += f" {unit_str}/day"
                    else:
                        daily_str += "/day"
                    record['Daily Dose'] = daily_str
                elif freq_note:
                    record['Daily Dose'] = (
                        f"{int(single_dose) if single_dose == int(single_dose) else single_dose}"
                        f" {freq_note}")
        except (ValueError, TypeError):
            pass

        record.pop('Frequency (Other)', None)
        cm_data.append(record)

    app.matrix_display.show_cm_matrix_from_data(cm_data, pat)
    return True


def _handle_mh(app, pat, row):
    """Show Medical History matrix from Main sheet _MH_ columns."""
    mh_cols = {
        'MHTERM': 'Condition', 'MHBODSYS': 'Body System', 'MHCAT': 'Category',
        'MHSTDTC': 'Start Date', 'MHENDTC': 'End Date', 'MHONGO': 'Ongoing',
    }
    mh_columns = {}
    for col in app.df_main.columns:
        col_str = str(col)
        if '_MH_' in col_str:
            for mh_key, display_name in mh_cols.items():
                if mh_key in col_str:
                    if display_name not in mh_columns:
                        mh_columns[display_name] = col_str
                    break

    if not mh_columns:
        messagebox.showinfo("Info", "No Medical History columns found in data.")
        return True

    term_col = mh_columns.get('Condition')
    if not term_col or pd.isna(row.get(term_col)):
        messagebox.showinfo("Info", "No medical history conditions found for this patient.")
        return True

    term_vals = [t.strip() for t in str(row[term_col]).split('|')
                 if t.strip() and t.strip().lower() != 'nan']
    if not term_vals:
        messagebox.showinfo("Info", "No medical history conditions found for this patient.")
        return True

    mh_data = []
    for i, term in enumerate(term_vals):
        record = {'MH #': str(i + 1), 'Condition': term}
        for display_name, col_name in mh_columns.items():
            if display_name == 'Condition':
                continue
            col_val = row.get(col_name, '')
            if pd.notna(col_val):
                vals = [v.strip() for v in str(col_val).split('|')]
                val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                if 'Date' in display_name and val:
                    if 'T' in val:
                        val = val.split('T')[0]
                    val = re.sub(r',?\s*time\s*unknown', '', val, flags=re.IGNORECASE).strip()
                    if val.lower() in ['date unknown', 'unknown date', 'unknown']:
                        val = 'Date Unknown'
                record[display_name] = val
        if record.get('Ongoing', '').lower() in ['yes', 'y', '1', 'true', 'checked']:
            record['End Date'] = 'Ongoing'
        mh_data.append(record)

    app.matrix_display.show_mh_matrix(mh_data, pat)
    return True


def _handle_hfh(app, pat, row):
    """Show Heart Failure History matrix."""
    hfh_cols = {
        'HOSTDTC': 'Hospitalization Date', 'HOTERM': 'Details',
        'HONUM': 'Number of Hospitalizations',
    }
    hfh_columns = {}
    for col in app.df_main.columns:
        col_str = str(col)
        if '_HFH_' in col_str:
            for hfh_key, display_name in hfh_cols.items():
                if hfh_key in col_str:
                    if display_name not in hfh_columns:
                        hfh_columns[display_name] = col_str
                    break

    if hfh_columns:
        hfh_data = []
        date_col = hfh_columns.get('Hospitalization Date')
        if date_col and pd.notna(row.get(date_col)):
            date_vals = [d.strip() for d in str(row[date_col]).split('|')
                         if d.strip() and d.strip().lower() != 'nan']
            for i, date_val in enumerate(date_vals):
                record = {'HFH #': str(i + 1), 'Hospitalization Date': date_val}
                for display_name, col_name in hfh_columns.items():
                    if display_name == 'Hospitalization Date':
                        continue
                    col_val = row.get(col_name, '')
                    if pd.notna(col_val):
                        vals = [v.strip() for v in str(col_val).split('|')]
                        val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                        record[display_name] = val
                hfh_data.append(record)
            if hfh_data:
                app.matrix_display.show_hfh_matrix(hfh_data, pat)
                return True

    messagebox.showinfo("Info", "No Heart Failure History data found for this patient.")
    return True


def _handle_hmeh(app, pat, row):
    """Show Hospitalization and Medical Events History matrix."""
    hmeh_cols = {'HOSTDTC': 'Event Date', 'HOTERM': 'Event Details'}
    hmeh_columns = {}
    for col in app.df_main.columns:
        col_str = str(col)
        if '_HMEH_' in col_str or 'HMEH' in col_str:
            for hmeh_key, display_name in hmeh_cols.items():
                if hmeh_key in col_str:
                    if display_name not in hmeh_columns:
                        hmeh_columns[display_name] = col_str
                    break

    if hmeh_columns:
        hmeh_data = []
        date_col = hmeh_columns.get('Event Date')
        term_col = hmeh_columns.get('Event Details')
        primary_col = date_col if date_col and pd.notna(row.get(date_col)) else term_col
        if primary_col and pd.notna(row.get(primary_col)):
            primary_vals = [v.strip() for v in str(row[primary_col]).split('|')
                            if v.strip() and v.strip().lower() != 'nan']
            for i, _pval in enumerate(primary_vals):
                record = {'HMEH #': str(i + 1)}
                for display_name, col_name in hmeh_columns.items():
                    col_val = row.get(col_name, '')
                    if pd.notna(col_val):
                        vals = [v.strip() for v in str(col_val).split('|')]
                        val = vals[i] if i < len(vals) and vals[i].strip().lower() != 'nan' else ''
                        if 'Date' in display_name and val:
                            if 'T' in val:
                                val = val.split('T')[0]
                        record[display_name] = val
                hmeh_data.append(record)
            if hmeh_data:
                app.matrix_display.show_hmeh_matrix(hmeh_data, pat)
                return True

    messagebox.showinfo("Info", "No Hospitalization/Medical Events History data found for this patient.")
    return True


def _handle_cvc(app, pat, processed_cols):
    """Show Cardiac and Venous Catheterization matrix."""
    exporter = CVCExporter(app.df_main)
    is_screening = any("SBV_CVC_" in str(col) for col in processed_cols)
    is_treatment = any("TV_CVC_" in str(col) for col in processed_cols)
    if not is_screening and not is_treatment:
        is_screening = is_treatment = True

    tables_shown = 0
    if is_screening:
        screening_df = exporter.generate_screening_table(pat)
        if screening_df is not None:
            app.matrix_display.show_cvc_matrix(screening_df, pat, "Screening")
            tables_shown += 1
    if is_treatment:
        hemo_df = exporter.generate_hemodynamic_table(pat)
        if hemo_df is not None:
            app.matrix_display.show_cvc_matrix(hemo_df, pat, "Hemodynamic Effect (Pre/Post Procedure)")
            tables_shown += 1
    if tables_shown == 0:
        messagebox.showinfo("Info", "No CVC data found for this patient.")
    return True


def _handle_cvh(app, pat, row):
    """Show Cardiovascular History matrix from CVH_TABLE sheet."""
    if app.df_cvh is not None and not app.df_cvh.empty:
        pat_cvh = app.df_cvh[
            app.df_cvh['Screening #'].astype(str).str.contains(
                pat.replace('-', '-'), na=False)]
        if not pat_cvh.empty:
            cvh_data = []
            for _, cvh_row in pat_cvh.iterrows():
                full_date = cvh_row.get('SBV_CVH_PRSTDTC', '')
                partial_date = cvh_row.get('SBV_CVH_PRSTDTC_PARTIAL', '')
                is_partial = str(cvh_row.get('SBV_CVH_PRSTDTC_PARTIAL_CHECKBOX', '')).lower() in [
                    'yes', 'checked', 'true', '1']
                if pd.notna(full_date) and str(full_date).strip() and str(full_date).strip().lower() not in ['nan', 'nat']:
                    date_str = str(full_date).split('T')[0] if 'T' in str(full_date) else str(full_date)
                elif pd.notna(partial_date) and str(partial_date).strip():
                    date_str = f"{partial_date} (partial)"
                else:
                    date_str = "Unknown"

                int_type = cvh_row.get('SBV_CVH_PRCAT', '')
                int_type_str = str(int_type).strip() if pd.notna(int_type) and str(int_type).strip().lower() not in ['nan', ''] else ""
                int_term = cvh_row.get('SBV_CVH_PRTRT', '')
                int_term_str = str(int_term).strip() if pd.notna(int_term) and str(int_term).strip().lower() not in ['nan', ''] else ""

                if int_type_str.lower() == 'other':
                    int_type_oth = cvh_row.get('SBV_CVH_PRCAT_OTH', '')
                    if pd.notna(int_type_oth) and str(int_type_oth).strip():
                        int_type_str = f"Other: {int_type_oth}"
                int_term_oth = cvh_row.get('SBV_CVH_PRTRT_OTHCAT', '')
                if int_term_str.lower() == 'other' and pd.notna(int_term_oth) and str(int_term_oth).strip():
                    int_term_str = f"Other: {int_term_oth}"

                if int_term_str or int_type_str or date_str != "Unknown":
                    cvh_data.append({
                        'Date': date_str,
                        'Type of Intervention': int_type_str,
                        'Intervention': int_term_str,
                    })

            if cvh_data:
                app.matrix_display.show_cvh_matrix(cvh_data, pat)
                return True
            messagebox.showinfo("Info", "No Cardiovascular History interventions found for this patient.")
            return True
        messagebox.showinfo("Info", "No Cardiovascular History data found for this patient in CVH sheet.")
        return True
    messagebox.showinfo("Info", "CVH sheet not loaded or empty.")
    return True


def _handle_act(app, pat, row):
    """Show ACT Lab Results matrix from LB_ACT sheet."""
    if app.df_act is not None and not app.df_act.empty:
        scr_col = next((c for c in app.df_act.columns
                        if "Screening" in str(c) and "#" in str(c)), None)
        if not scr_col:
            scr_col = next((c for c in app.df_act.columns
                            if "Screening" in str(c)), None)
        if scr_col:
            pat_clean = str(pat).strip()
            pat_act = app.df_act[
                app.df_act[scr_col].astype(str).str.strip().str.contains(
                    pat_clean, regex=False, na=False)]
        else:
            pat_act = pd.DataFrame()

        if not pat_act.empty:
            act_events = []
            for _, act_row in pat_act.iterrows():
                act_time = act_row.get('TV_LB_ACT_LBTIM_ACT',
                           act_row.get('UV_LB_ACT_LBTIM_ACT', ''))
                act_level = act_row.get('TV_LB_ACT_LBORRES_ACT',
                             act_row.get('UV_LB_ACT_LBORRES_ACT', ''))
                act_stat = act_row.get('TV_LB_ACT_LBSTAT_ACT',
                            act_row.get('UV_LB_ACT_LBSTAT_ACT', ''))
                if pd.notna(act_time) and str(act_time).strip() and str(act_time).strip().lower() not in ['nan', '']:
                    act_level_str = str(act_level).strip() if pd.notna(act_level) else ""
                    act_events.append({
                        'Time': str(act_time).strip(), 'Event': "ACT Level",
                        'Value': f"{act_level_str} sec" if act_level_str else "",
                        'Type': 'ACT', 'Status': 'OK' if act_level_str else 'GAP',
                    })
                elif pd.notna(act_stat) and str(act_stat).strip().lower() in ['not done', 'not performed']:
                    act_events.append({
                        'Time': '', 'Event': "ACT Level", 'Value': "Not Done",
                        'Type': 'ACT', 'Status': 'Confirmed',
                    })

                hep_time = act_row.get('TV_LB_ACT_CMTIM_HEP',
                           act_row.get('UV_LB_ACT_CMTIM_HEP', ''))
                hep_dose = act_row.get('TV_LB_ACT_CMDOS_HEP',
                           act_row.get('UV_LB_ACT_CMDOS_HEP', ''))
                hep_stat = act_row.get('TV_LB_ACT_CMSTAT_HEP',
                           act_row.get('UV_LB_ACT_CMSTAT_HEP', ''))
                if pd.notna(hep_time) and str(hep_time).strip() and str(hep_time).strip().lower() not in ['nan', '']:
                    hep_dose_str = str(hep_dose).strip() if pd.notna(hep_dose) else ""
                    act_events.append({
                        'Time': str(hep_time).strip(), 'Event': "Heparin",
                        'Value': f"{hep_dose_str} Units" if hep_dose_str else "",
                        'Type': 'HEP', 'Status': 'OK' if hep_dose_str else 'GAP',
                    })
                elif pd.notna(hep_stat) and str(hep_stat).strip().lower() in ['not done', 'not performed']:
                    act_events.append({
                        'Time': '', 'Event': "Heparin", 'Value': "Not Done",
                        'Type': 'HEP', 'Status': 'Confirmed',
                    })

            if act_events:
                act_events.sort(key=lambda x: parse_time_minutes(x['Time']))
                app.matrix_display.show_act_matrix(act_events, pat)
                return True
            messagebox.showinfo("Info", "No ACT/Heparin data found for this patient.")
            return True
        messagebox.showinfo("Info", "No ACT data found for this patient in ACT sheet.")
        return True
    messagebox.showinfo("Info", "ACT sheet not loaded or empty.")
    return True


# Dispatch table: col_type -> handler(app, pat, row)
# Special cases: 'cvc' also needs processed_cols, 'ae_ref' is not dispatched
_FORM_HANDLERS = {
    'ae':   _handle_ae,
    'cm':   _handle_cm,
    'mh':   _handle_mh,
    'hfh':  _handle_hfh,
    'hmeh': _handle_hmeh,
    'cvh':  _handle_cvh,
    'act':  _handle_act,
}


# ---------------------------------------------------------------------------
# Lab/result column resolution
# ---------------------------------------------------------------------------

def _resolve_parallel_columns(app, col_name, row):
    """Find shared date, unit, test-name columns for a lab/result column.

    Returns (date_col, unit_col, other_unit_col, test_name_col).
    """
    date_col = unit_col = other_unit_col = test_name_col = None
    is_tv_lab = col_name.startswith("TV_") and "_LB_" in col_name and "_DV_" in col_name

    if is_tv_lab:
        tv_parts = col_name.split("_DV_")
        if len(tv_parts) >= 2:
            lab_prefix = tv_parts[0]
            suffix_part = tv_parts[1]
            lab_match = re.search(r'TV_LB_(\w+)', lab_prefix)
            lab_type = lab_match.group(1) if lab_match else None
            if lab_type:
                shared_date_col = f"{lab_prefix}_DV_LBDAT_{lab_type}"
                if shared_date_col in app.df_main.columns:
                    date_col = shared_date_col
            if "LBORRES_" in suffix_part:
                test_suffix = suffix_part.split("LBORRES_")[1]
                unit_col_name = f"{lab_prefix}_DV_LBORRESU_{test_suffix}"
                if unit_col_name in app.df_main.columns:
                    unit_col = unit_col_name
                other_unit_col_name = f"{lab_prefix}_DV_LBORRESU_OTH_{test_suffix}"
                if other_unit_col_name in app.df_main.columns:
                    other_unit_col = other_unit_col_name
    else:
        prefix = None
        if "_LBORRES" in col_name:
            prefix = col_name.split("_LBORRES")[0]
        elif "_ORRES" in col_name:
            prefix = col_name.split("_ORRES")[0]
        elif "_PRORRES" in col_name:
            prefix = col_name.split("_PRORRES")[0]

        if prefix:
            is_lab_col = ("_LBORRES" in col_name or "_ORRES" in col_name) and "_PRORRES" not in col_name
            for cand in app.df_main.columns:
                if cand.startswith(prefix):
                    if is_lab_col:
                        if "LBDTC" in cand or "LBDAT" in cand:
                            date_col = cand
                        if "LBTEST" in cand:
                            test_name_col = cand
                        if "LBORRESU" in cand and "OTH" not in cand:
                            unit_col = cand
                        if "LBORRESU_OTH" in cand:
                            other_unit_col = cand
                    else:
                        if "PRDTC" in cand or "PRDAT" in cand:
                            date_col = cand
                        if "PRTEST" in cand:
                            test_name_col = cand

            if not date_col:
                lab_match = re.match(r'^(\w+)_LB_(\w+?)P?_', col_name)
                if lab_match:
                    visit_pfx = lab_match.group(1)
                    lab_type_base = lab_match.group(2)
                    for pat_str in [
                        f"{visit_pfx}_LB_{lab_type_base}_LBDAT_{lab_type_base}",
                        f"{visit_pfx}_LB_{lab_type_base}_LBDTC_{lab_type_base}",
                        f"{visit_pfx}_LB_{lab_type_base}_LBDTC",
                    ]:
                        if pat_str in app.df_main.columns:
                            date_col = pat_str
                            break

    return date_col, unit_col, other_unit_col, test_name_col, is_tv_lab


# ---------------------------------------------------------------------------
# Matrix row builder for generic lab/result values
# ---------------------------------------------------------------------------

def _build_matrix_rows(app, col_name, row, proc_date):
    """Build matrix_data rows from a generic lab/result column.

    Returns list of dicts with keys: Time, Param, Value, AE_Ref.
    """
    val_str = str(row[col_name])
    date_col, unit_col, other_unit_col, test_name_col, is_tv_lab = (
        _resolve_parallel_columns(app, col_name, row))

    r_vals = [v.strip() for v in val_str.split('|')]
    param_name_default = app.clean_label(app.labels.get(col_name, col_name))
    param_name_default = app.annotate_procedure_timing(param_name_default, col_name)
    if param_name_default.lower().endswith('/result'):
        param_name_default = param_name_default[:-7]

    n_vals = []
    if test_name_col and pd.notna(row[test_name_col]):
        n_vals = [n.strip() for n in str(row[test_name_col]).split('|')]

    d_vals = []
    if date_col and pd.notna(row[date_col]):
        d_vals = [d.strip() for d in str(row[date_col]).split('|')]

    u_vals = []
    if unit_col and pd.notna(row[unit_col]):
        u_vals = [u.strip() for u in str(row[unit_col]).split('|')]

    other_u_vals = []
    if is_tv_lab and other_unit_col and pd.notna(row.get(other_unit_col, None)):
        other_u_vals = [u.strip() for u in str(row[other_unit_col]).split('|')]

    rows_out = []
    for i, val in enumerate(r_vals):
        if not val:
            continue
        curr_param = n_vals[i] if n_vals and i < len(n_vals) and n_vals[i] else param_name_default
        if curr_param.lower().endswith('/result'):
            curr_param = curr_param[:-7]

        # Resolve date / visit label
        curr_date = "Unknown"
        visit_label = ""
        prefix = ""
        for pfx, label in VISIT_MAP.items():
            if col_name.startswith(pfx + "_"):
                visit_label = label
                prefix = pfx
                break

        specific_date = None
        if d_vals and i < len(d_vals) and d_vals[i]:
            specific_date = str(d_vals[i]).split('T')[0]

        if not specific_date and visit_label:
            for date_col_name, _sv_label in VISIT_SCHEDULE:
                if date_col_name.startswith(prefix + "_"):
                    visit_date_val = row.get(date_col_name)
                    if pd.notna(visit_date_val):
                        specific_date = str(visit_date_val).split('T')[0]
                    break

        header_visit = visit_label
        if visit_label == "Treatment" and proc_date and specific_date:
            try:
                curr_dt = datetime.strptime(specific_date, '%Y-%m-%d')
                delta = (curr_dt - proc_date).days
                if delta > 0:
                    header_visit = f"Treat. Day +{delta}"
                elif delta < 0:
                    header_visit = f"Treat. Day {delta}"
            except ValueError:
                pass

        if header_visit and specific_date:
            curr_date = f"{header_visit} ({specific_date})"
        elif header_visit:
            curr_date = header_visit
        elif specific_date:
            curr_date = specific_date

        display_val = val

        # "Other" unit substitution
        is_unit_col = "LBORRESU" in col_name and "OTH" not in col_name
        if is_unit_col and val.lower() == "other":
            oth_col_name = col_name.replace("LBORRESU_", "LBORRESU_OTH_")
            if oth_col_name in app.df_main.columns:
                oth_val = row.get(oth_col_name, None)
                if pd.notna(oth_val):
                    oth_vals = str(oth_val).split('|')
                    if i < len(oth_vals) and oth_vals[i].strip():
                        display_val = oth_vals[i].strip()

        # Append units for lab panels
        lab_panels = ["_LB_BM_", "_LB_ENZ_", "_LB_CBC_", "_LB_BMP_", "_LB_COA", "_LB_LFP_"]
        is_result_col = ("_LBORRES_" in col_name or "_ORRES" in col_name) and "LBORRESU" not in col_name
        if any(p in col_name for p in lab_panels) and is_result_col:
            curr_unit = ""
            if u_vals and i < len(u_vals) and u_vals[i] and u_vals[i].lower() not in ['nan', 'none', '']:
                curr_unit = u_vals[i]
            if curr_unit.lower() == "other" and other_u_vals and i < len(other_u_vals):
                if other_u_vals[i] and other_u_vals[i].lower() not in ['nan', '']:
                    curr_unit = other_u_vals[i]
            if curr_unit:
                display_val = f"{display_val} {curr_unit}"

        # AE Reference
        ae_ref = ""
        if "LOGS_" in col_name:
            ref_type = 'PR' if 'PRORRES' in col_name else 'LB'
            ae_num, ae_term = app.get_ae_info(str(i + 1), ref_type)
            if ae_term:
                ae_ref = f"AE#{ae_num}"

        # Angiography split
        if "_AG_" in col_name:
            if col_name.count("_PRE_") > 1:
                curr_date = f"{curr_date} (Pre-Procedure)"
                for pfx in ["Pre-procedure / ", "Pre-procedure /"]:
                    if curr_param.startswith(pfx):
                        curr_param = curr_param.replace(pfx, "")
            elif col_name.count("_POST_") > 1:
                curr_date = f"{curr_date} (Post-Procedure)"
                for pfx in ["Post-procedure / ", "Post-procedure /"]:
                    if curr_param.startswith(pfx):
                        curr_param = curr_param.replace(pfx, "")

        rows_out.append({'Time': curr_date, 'Param': curr_param,
                         'Value': display_val, 'AE_Ref': ae_ref})
    return rows_out


# ---------------------------------------------------------------------------
# show_data_matrix (main entry point — called from clinical_viewer1)
# ---------------------------------------------------------------------------

def show_data_matrix(app):
    """Build and display a data matrix for the selected tree items.

    This replaces the monolithic ClinicalDataMasterV30.show_data_matrix().
    """
    selected_items = app.tree.selection()
    if not selected_items:
        messagebox.showwarning("Warning", "Please select at least one row.")
        return

    site, pat = app.cb_site.get(), app.cb_pat.get()
    if not site or not pat or app.df_main is None:
        messagebox.showwarning("Warning", "Patient data not loaded.")
        return

    mask = ((app.df_main['Site #'].astype(str).apply(app._clean_id) == site) &
            (app.df_main['Screening #'].astype(str).apply(app._clean_id) == pat))
    rows = app.df_main[mask]
    if rows.empty:
        return
    row = rows.iloc[0]

    matrix_data = []
    processed_cols = set()
    requested_types = set()

    # Procedure date for "Treat. Day" calculations
    proc_date = None
    if 'TV_PR_SVDTC' in row and pd.notna(row['TV_PR_SVDTC']):
        proc_date_str = str(row['TV_PR_SVDTC']).split('T')[0]
        try:
            proc_date = datetime.strptime(proc_date_str, '%Y-%m-%d')
        except ValueError:
            pass

    # --- Collect data from selected tree items ---
    for item in selected_items:
        descendants = app._get_all_descendants(item)
        items_to_process = [item] + descendants if descendants else [item]

        for proc_item in items_to_process:
            vals = app.tree.item(proc_item, "values")
            if not vals or len(vals) < 5:
                continue
            col_name = vals[4]
            if not col_name or col_name in processed_cols:
                continue
            if col_name not in app.df_main.columns:
                continue

            col_type = classify_column(col_name)

            if col_type == 'ae_ref':
                # AE reference columns are handled inline with timeline entries
                pass
            elif col_type in _FORM_HANDLERS:
                requested_types.add(col_type)
                processed_cols.add(col_name)
                continue
            else:
                if not app._is_matrix_supported_col(col_name):
                    continue

            processed_cols.add(col_name)
            val_in_db = row[col_name]
            if pd.isna(val_in_db):
                continue
            val_str = str(val_in_db)

            # Timeline entries (pipe-delimited #row/date/param/value)
            if "#" in val_str and "/" in val_str and " / " in val_str:
                entries = [e.strip() for e in val_str.split('|')] if '|' in val_str else [val_str]
                for e in entries:
                    if e.startswith("#") and "/" in e:
                        parsed = app.parse_timeline_entry(e)
                        if parsed:
                            row_num, d, p, v = parsed
                            ae_ref = ""
                            if "LOGS_" in col_name:
                                ref_type = 'PR' if 'PRORRES' in col_name else 'LB'
                                ae_num, ae_term = app.get_ae_info(row_num, ref_type)
                                if ae_term:
                                    ae_ref = f"AE#{ae_num}"
                            matrix_data.append({'Time': d, 'Param': p,
                                                'Value': v, 'AE_Ref': ae_ref})
            else:
                # Skip column types that will be dispatched to form handlers
                if col_type in _FORM_HANDLERS:
                    continue
                # Generic lab/result column
                matrix_data.extend(_build_matrix_rows(app, col_name, row, proc_date))

    # --- Dispatch repeating-form handlers ---
    for rtype in requested_types:
        if rtype == 'cvc':
            _handle_cvc(app, pat, processed_cols)
        else:
            handler = _FORM_HANDLERS[rtype]
            handler(app, pat, row)
        return  # form handlers take over the window

    # --- Build pivot display ---
    if not matrix_data:
        messagebox.showinfo("Info", "No data found.")
        return

    _show_pivot_matrix(app, matrix_data, pat)


# ---------------------------------------------------------------------------
# Pivot table + tree display window
# ---------------------------------------------------------------------------

def _show_pivot_matrix(app, matrix_data, pat):
    """Build a pivot table from matrix_data and display in a Toplevel window."""
    df_matrix = pd.DataFrame(matrix_data)

    # Filter artifacts
    df_matrix = df_matrix[~df_matrix['Param'].str.strip().str.endswith('/')]
    df_matrix = df_matrix[df_matrix['Param'].str.strip() != '']

    df_matrix['AE_Ref'] = df_matrix['AE_Ref'].fillna('')
    df_matrix['Row_Key'] = df_matrix.apply(
        lambda x: f"{x['Param']}||{x['AE_Ref']}" if x['AE_Ref'] else x['Param'],
        axis=1,
    )
    df_matrix['Time_Unique'] = df_matrix.groupby(['Row_Key', 'Time']).cumcount()
    df_matrix['Time_Label'] = df_matrix.apply(
        lambda x: f"{x['Time']} ({x['Time_Unique'] + 1})" if x['Time_Unique'] > 0 else x['Time'],
        axis=1,
    )

    df_pivot = df_matrix.pivot_table(
        index='Row_Key', columns='Time_Label', values='Value', aggfunc='first')

    win = tk.Toplevel(app.root)
    win.title(f"Data Matrix - Patient {pat}")
    win.geometry("1200x600")

    # Store for export
    app.data_matrix_df = df_pivot
    app.data_matrix_patient = pat

    # Toolbar
    toolbar = tk.Frame(win, bg="#f4f4f4", pady=5)
    toolbar.pack(fill=tk.X, side=tk.TOP)

    tk.Label(toolbar, text="Export:", bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(10, 5))
    tk.Button(toolbar, text="Export XLSX", command=lambda: app.export_data_matrix('xlsx'),
              bg="#27ae60", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)
    tk.Button(toolbar, text="Export CSV", command=lambda: app.export_data_matrix('csv'),
              bg="#3498db", fg="white", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT, padx=5)

    tk.Label(toolbar, text="  |", bg="#f4f4f4", fg="#666").pack(side=tk.LEFT, padx=5)

    # --- Toggle units ---
    hide_units_var = tk.BooleanVar(value=False)
    time_cols = sorted(df_pivot.columns.tolist(), key=try_parse_date)

    ae_ref_lookup = {}
    for row_data in df_matrix.itertuples(index=False):
        param = getattr(row_data, 'Param', '')
        time_val = getattr(row_data, 'Time', '')
        ae_ref = getattr(row_data, 'AE_Ref', '')
        if param and time_val:
            for tc in [t for t in df_matrix['Time_Label'].unique()
                       if str(t).startswith(str(time_val)[:10])]:
                ae_ref_lookup[(param, tc)] = ae_ref
            ae_ref_lookup[(param, time_val)] = ae_ref
    has_ae_refs = any(ae_ref_lookup.values())

    # Tree container
    tree_frame = tk.Frame(win)
    tree_frame.pack(fill=tk.BOTH, expand=True)

    tree = ttk.Treeview(tree_frame)

    if has_ae_refs:
        tree["columns"] = ["Parameter"] + time_cols + ["AE Ref"]
    else:
        tree["columns"] = ["Parameter"] + time_cols
    tree.heading("#0", text="")
    tree.column("#0", width=0, stretch=tk.NO)
    tree.heading("Parameter", text="Parameter")
    tree.column("Parameter", width=300, anchor="w", minwidth=100)
    for tc in time_cols:
        tree.heading(tc, text=tc)
        tree.column(tc, width=150, anchor="center", minwidth=80)
    if has_ae_refs:
        tree.heading("AE Ref", text="AE Ref")
        tree.column("AE Ref", width=80, anchor="center", minwidth=60)

    def toggle_units():
        for item in tree.get_children():
            tree.delete(item)
        for row_key in df_pivot.index:
            if "||" in str(row_key):
                param_name, ae_ref = str(row_key).split("||", 1)
            else:
                param_name, ae_ref = str(row_key), ""
            param_lower = param_name.lower().strip()
            is_unit_row = (param_lower.endswith('/units') or param_lower.endswith('/')
                           or param_lower.endswith('units')
                           or (('/' in param_lower) and ('unit' in param_lower.split('/')[-1])))
            if hide_units_var.get() and is_unit_row:
                continue
            row_vals = [param_name]
            for tc in time_cols:
                val = df_pivot.at[row_key, tc]
                row_vals.append(val if pd.notna(val) else "")
            if has_ae_refs:
                row_vals.append(ae_ref)
            tree.insert("", "end", values=row_vals)

    tk.Checkbutton(toolbar, text="Hide Unit Rows", variable=hide_units_var,
                   command=toggle_units, bg="#f4f4f4", font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=5)
    tk.Label(toolbar, text="  |  Tip: Drag column borders to resize", bg="#f4f4f4", fg="#666",
             font=("Segoe UI", 8, "italic")).pack(side=tk.LEFT, padx=10)

    # Populate initial rows
    for row_key in df_pivot.index:
        if "||" in str(row_key):
            param_name, ae_ref = str(row_key).split("||", 1)
        else:
            param_name, ae_ref = str(row_key), ""
        row_vals = [param_name]
        for tc in time_cols:
            val = df_pivot.at[row_key, tc]
            row_vals.append(val if pd.notna(val) else "")
        if has_ae_refs:
            row_vals.append(ae_ref)
        tree.insert("", "end", values=row_vals)

    h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
    v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
    tree.configure(xscrollcommand=h_scroll.set, yscrollcommand=v_scroll.set)

    tree.grid(row=0, column=0, sticky="nsew")
    v_scroll.grid(row=0, column=1, sticky="ns")
    h_scroll.grid(row=1, column=0, sticky="ew")

    tree_frame.grid_rowconfigure(0, weight=1)
    tree_frame.grid_columnconfigure(0, weight=1)
