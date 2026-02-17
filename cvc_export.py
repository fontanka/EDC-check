"""
CVC (Cardiac and Venous Catheterization) Export Module

Exports CVC data for:
- Screening visit
- Treatment visit (Pre-procedure and Post-procedure)

Generates tables in xlsx or csv format.
"""

import pandas as pd
import os
import logging
from io import BytesIO

from base_exporter import BaseExporter

logger = logging.getLogger(__name__)

# Field mappings: Display Name -> Column Suffix
CVC_FIELDS = {
    "cvp_mean": "CVPM",
    "cvp_vwave": "CVPV",
    "rap_mean": "RAPM",
    "rap_vwave": "RAPV",
    "rvedp": "RVEDP",
    "rvp_systolic": "SRVP",
    "pap_systolic": "SPAP",
    "pap_diastolic": "DPAP",
    "pap_mean": "MPAP",
    "pcwp_mean": "PCWPM",
    "co": "CO",
}

# Visit configuration
VISIT_CONFIG = {
    "Screening": {
        "prefix": "SBV_CVC_CVORRES_",
        "date_col": "SBV_CVC_PRSTDTC_CVC",
        "height_col": "SBV_VS_VSORRES_HEIGHT",  # Height in cm
        "weight_col": "SBV_VS_VSORRES_WEIGHT",  # Weight in kg
        "pvr_col": "SBV_CVC_FAORRES_ECHO_PVR",  # PVR stored separately
    },
    "Pre-procedure": {
        "prefix": "TV_CVC_PRE_POST_CVORRES_PRE_",
        "date_col": "TV_CVC_PRE_POST_PRSTDTC_PRE_CVC",
        "height_col": "TV_VS_VSORRES_HEIGHT",  # 1-day pre-procedure height
        "weight_col": "TV_VS_VSORRES_WEIGHT",  # 1-day pre-procedure weight
        "pvr_col": None,  # No PVR for treatment visits
    },
    "Post-procedure": {
        "prefix": "TV_CVC_PRE_POST_CVORRES_POST_",
        "date_col": "TV_CVC_PRE_POST_PRSTDTC_POST_CVC",
        "height_col": "TV_VS_VSORRES_HEIGHT",  # Same as pre-procedure
        "weight_col": "TV_VS_VSORRES_WEIGHT",
        "pvr_col": None,  # No PVR for treatment visits
    },
}


class CVCExporter(BaseExporter):
    def __init__(self, df_main):
        super().__init__(df_main)

    def get_value(self, row, col_name):
        """Get value from row, returning None if invalid."""
        return self.safe_str(row, col_name)

    def get_numeric(self, row, col_name):
        """Get numeric value from row."""
        val = self.get_value(row, col_name)
        return self.to_float(val)

    def get_integer(self, row, col_name):
        """Get integer value from row (for pressure values without decimals)."""
        val = self.get_numeric(row, col_name)
        if val is None:
            return None
        return int(round(val))
    
    def calculate_ci(self, co, bsa):
        """Calculate Cardiac Index: CI = CO / BSA"""
        if co is None or bsa is None:
            return None
        try:
            co_val = float(co)
            bsa_val = float(bsa)
            if bsa_val > 0:
                return round(co_val / bsa_val, 1)
        except (ValueError, TypeError):
            pass
        return None
    
    def calculate_bsa(self, height_cm, weight_kg):
        """Calculate Body Surface Area using Mosteller formula: BSA = sqrt((H × W) / 3600)"""
        if height_cm is None or weight_kg is None:
            return None
        try:
            h = float(height_cm)
            w = float(weight_kg)
            if h > 0 and w > 0:
                import math
                return round(math.sqrt((h * w) / 3600), 2)
        except (ValueError, TypeError):
            pass
        return None
    
    def get_visit_data(self, patient_id, visit_name):
        """Get CVC data for a specific visit."""
        rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if rows.empty:
            return None
        
        row = rows.iloc[0]
        config = VISIT_CONFIG.get(visit_name)
        if not config:
            return None
        
        prefix = config["prefix"]
        date_col = config["date_col"]
        height_col = config["height_col"]
        weight_col = config["weight_col"]
        
        # Get date - include time for Pre/Post procedure
        visit_date = self.get_value(row, date_col)
        if visit_date:
            try:
                dt = pd.to_datetime(visit_date)
                # Include time for Pre/Post procedure visits
                if visit_name in ["Pre-procedure", "Post-procedure"]:
                    visit_date = dt.strftime("%d-%b-%Y %H:%M")
                else:
                    visit_date = dt.strftime("%d-%b-%Y")
            except (ValueError, TypeError):
                visit_date = str(visit_date)
        
        # Calculate BSA from height and weight for CI calculation
        height = self.get_numeric(row, height_col)
        weight = self.get_numeric(row, weight_col)
        bsa = self.calculate_bsa(height, weight)
        
        # Get all CVC values - pressure values as integers, CO as float
        data = {"date": visit_date, "bsa": bsa}
        pressure_fields = ["cvp_mean", "cvp_vwave", "rap_mean", "rap_vwave", 
                           "rvedp", "rvp_systolic", "pap_systolic", "pap_diastolic", 
                           "pap_mean", "pcwp_mean"]
        for field_key, suffix in CVC_FIELDS.items():
            col_name = f"{prefix}{suffix}"
            if field_key in pressure_fields:
                data[field_key] = self.get_integer(row, col_name)
            else:
                data[field_key] = self.get_numeric(row, col_name)
        
        # Calculate CI = CO / BSA
        data["ci"] = self.calculate_ci(data.get("co"), bsa)
        
        # Get PVR from data column (not calculated)
        pvr_col = config.get("pvr_col")
        if pvr_col:
            data["pvr"] = self.get_numeric(row, pvr_col)
        else:
            data["pvr"] = None
        
        return data
    
    def generate_screening_table(self, patient_id):
        """Generate screening table (Right Heart Catheterization format)."""
        data = self.get_visit_data(patient_id, "Screening")
        if not data:
            return None
        
        # Build table matching the template format
        table_data = {
            "Date": [data.get("date", "")],
            "CVP Mean [mmHg]": [data.get("cvp_mean", "")],
            "CVP V-wave [mmHg]": [data.get("cvp_vwave", "")],
            "RAP Mean [mmHg]": [data.get("rap_mean", "")],
            "RAP V-wave [mmHg]": [data.get("rap_vwave", "")],
            "RVP RVEDP [mmHg]": [data.get("rvedp", "")],
            "RVP Systolic [mmHg]": [data.get("rvp_systolic", "")],
            "PAP Systolic [mmHg]": [data.get("pap_systolic", "")],
            "PAP Diastolic [mmHg]": [data.get("pap_diastolic", "")],
            "PAP Mean [mmHg]": [data.get("pap_mean", "")],
            "Mean PCWP [mmHg]": [data.get("pcwp_mean", "")],
            "CO [L/min]": [data.get("co", "")],
            "CI [L/min/m²]": [data.get("ci", "")],
            "PVR [WU]": [data.get("pvr", "")],
        }
        
        return pd.DataFrame(table_data)
    
    def generate_hemodynamic_table(self, patient_id):
        """Generate hemodynamic effect table (Pre/Post comparison)."""
        pre_data = self.get_visit_data(patient_id, "Pre-procedure")
        post_data = self.get_visit_data(patient_id, "Post-procedure")
        
        if not pre_data and not post_data:
            return None
        
        pre_data = pre_data or {}
        post_data = post_data or {}
        
        def format_combined(v_wave, mean, suffix=""):
            """Format as 'V,M' or with suffix like '(IVC)'"""
            parts = []
            if v_wave is not None:
                parts.append(str(v_wave))
            if mean is not None:
                parts.append(str(mean))
            result = ",".join(parts) if parts else ""
            return f"{result} {suffix}".strip() if suffix and result else result
        
        def format_pap(systolic, diastolic, mean):
            """Format PAP as 'S/D, M'"""
            parts = []
            if systolic is not None and diastolic is not None:
                parts.append(f"{systolic}/{diastolic}")
            if mean is not None:
                parts.append(str(mean))
            return ", ".join(parts) if parts else ""
        
        def format_rvp(systolic, rvedp):
            """Format RVP as 'S/D'"""
            if systolic is not None and rvedp is not None:
                return f"{systolic}/{rvedp}"
            return ""
        
        # Build comparison table
        table_data = {
            "Parameter": [
                "RAP (V,M) [mmHg]",
                "CVP (V,M) [mmHg]",
                "PAP (S/D,M) [mmHg]",
                "RVP (S/D) [mmHg]",
                "CO [L/min]",
                "CI [L/min/m²]",
            ],
            "Pre Trillium Implantation": [
                format_combined(pre_data.get("rap_vwave"), pre_data.get("rap_mean")),
                format_combined(pre_data.get("cvp_vwave"), pre_data.get("cvp_mean")),
                format_pap(pre_data.get("pap_systolic"), pre_data.get("pap_diastolic"), pre_data.get("pap_mean")),
                format_rvp(pre_data.get("rvp_systolic"), pre_data.get("rvedp")),
                pre_data.get("co", ""),
                pre_data.get("ci", ""),
            ],
            "Post Trillium Implantation": [
                format_combined(post_data.get("rap_vwave"), post_data.get("rap_mean")),
                format_combined(post_data.get("cvp_vwave"), post_data.get("cvp_mean")),
                format_pap(post_data.get("pap_systolic"), post_data.get("pap_diastolic"), post_data.get("pap_mean")),
                format_rvp(post_data.get("rvp_systolic"), post_data.get("rvedp")),
                post_data.get("co", ""),
                post_data.get("ci", ""),
            ],
        }
        
        return pd.DataFrame(table_data)
    
    def export_to_excel(self, patient_id, include_screening=True, include_hemodynamic=True):
        """Export tables to Excel format (BytesIO)."""
        output = BytesIO()
        
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if include_screening:
                screening_df = self.generate_screening_table(patient_id)
                if screening_df is not None:
                    screening_df.to_excel(writer, sheet_name='Screening CVC', index=False)
            
            if include_hemodynamic:
                hemo_df = self.generate_hemodynamic_table(patient_id)
                if hemo_df is not None:
                    hemo_df.to_excel(writer, sheet_name='Hemodynamic Effect', index=False)
        
        output.seek(0)
        return output.getvalue()
    
    def export_to_csv(self, patient_id, table_type="screening"):
        """Export single table to CSV format."""
        if table_type == "screening":
            df = self.generate_screening_table(patient_id)
        else:
            df = self.generate_hemodynamic_table(patient_id)
        
        if df is None:
            return None
        
        return df.to_csv(index=False)
