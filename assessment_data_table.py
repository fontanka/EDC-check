"""
Assessment Data Table Module

Provides data extraction for clinical assessments across visit types.
Used by the Assessment Data Table UI in clinical_viewer1.py.
"""

import pandas as pd
import re


# ============================================================================
# ASSESSMENT CONFIGURATIONS - Reusing mappings from existing export modules
# ============================================================================

# Labs parameters grouped by panel (from labs_export.py TEMPLATE_ROW_MAP)
LAB_PARAMS = {
    "CBC": [
        ("RBC", "RBC", "Red Blood Cells"),
        ("HGB", "HGB", "Hemoglobin"),
        ("HCT", "HCT", "Hematocrit"),
        ("PLT", "PLAT", "Platelets"),
        ("WBC", "WBC", "White Blood Cells"),
        ("MCV", "MCV", "Mean Corpuscular Volume"),
        ("MCH", "MCH", "Mean Corpuscular Hemoglobin"),
        ("MCHC", "MCHC", "Mean Corpuscular Hemoglobin Concentration"),
        ("RDW", "RDW", "Red Cell Distribution Width"),
        ("NEUT%", "NEUTP", "Neutrophils %"),
        ("NEUT abs", "NEUTA", "Neutrophils Absolute"),
        ("LYM%", "LYMPP", "Lymphocytes %"),
        ("LYM abs", "LYMPA", "Lymphocytes Absolute"),
        ("MONO%", "MONOP", "Monocytes %"),
        ("MONO abs", "MONOA", "Monocytes Absolute"),
        ("EOS%", "EOSP", "Eosinophils %"),
        ("EOS abs", "EOSA", "Eosinophils Absolute"),
        ("BASO%", "BASOP", "Basophils %"),
        ("BASO abs", "BASOA", "Basophils Absolute"),
    ],
    "BMP": [
        ("BUN", "BUN", "Blood Urea Nitrogen"),
        ("Glucose", "GLUC", "Glucose"),
        ("Creatinine", "CREA", "Creatinine"),
        ("eGFR", "eGFR", "Estimated GFR"),
        ("Sodium", "SODIUM", "Sodium"),
        ("Chloride", "CL", "Serum Chloride"),
        ("Potassium", "K", "Potassium"),
    ],
    "LFP": [
        ("Bilirubin Total", "BILI", "Total Bilirubin"),
        ("AST (GOT)", "AST", "Aspartate Aminotransferase"),
        ("ALT (GPT)", "ALT", "Alanine Aminotransferase"),
        ("GGT", "GGT", "Gamma-Glutamyl Transferase"),
        ("LDH", "LDH", "Lactate Dehydrogenase"),
    ],
    "ENZ": [
        ("Troponin T", "TROPONT", "Troponin T"),
        ("Troponin I", "TROPONI", "Troponin I"),
        ("NT-proBNP", "BNPPRO", "N-terminal pro-BNP"),
    ],
    "BM": [
        ("Procalcitonin", "PCT", "Procalcitonin"),
        ("CRP", "CRP", "C-Reactive Protein"),
    ],
    "COA": [
        ("PT-INR", "PTI", "Prothrombin Time INR"),
        ("PT", "PT", "Prothrombin Time"),
        ("APTT", "APTT", "Activated Partial Thromboplastin Time"),
    ],
}

# Lab visit configurations (from labs_export.py LABS_VISIT_CONFIG)
# Note: FU3M, FU3Y, FU5Y are remote/phone visits — no labs collected
LAB_VISIT_CONFIG = {
    "Screening": {"prefix": "SBV_LB_", "date_col": "SBV_SV_SVSTDTC"},
    "Baseline": {"prefix": "TV_LB_", "date_col": "TV_SV_SVSTDTC"},  # 1-day pre-procedure
    "Discharge": {"prefix": "DV_LB_", "date_col": "DV_SV_SVSTDTC"},
    "30D": {"prefix": "FU1M_LB_", "date_col": "FU1M_SV_SVSTDTC"},
    "6M": {"prefix": "FU6M_LB_", "date_col": "FU6M_SV_SVSTDTC"},
    "1Y": {"prefix": "FU1Y_LB_", "date_col": "FU1Y_SV_SVSTDTC"},
    "2Y": {"prefix": "FU2Y_LB_", "date_col": "FU2Y_SV_SVSTDTC"},
    "4Y": {"prefix": "FU4Y_LB_", "date_col": "FU4Y_SV_SVSTDTC"},
}

# CVC parameters (from cvc_export.py CVC_FIELDS + CVP)
CVC_PARAMS = [
    ("CVP Mean", "CVPM", "Central Venous Pressure Mean"),
    ("CVP V-wave", "CVPV", "Central Venous Pressure V-wave"),
    ("RAP Mean", "RAPM", "Right Atrial Pressure Mean"),
    ("RAP V-wave", "RAPV", "Right Atrial Pressure V-wave"),
    ("RVEDP", "RVEDP", "Right Ventricular End-Diastolic Pressure"),
    ("RVP Systolic", "SRVP", "Right Ventricular Pressure Systolic"),
    ("PAP Systolic", "SPAP", "Pulmonary Artery Pressure Systolic"),
    ("PAP Diastolic", "DPAP", "Pulmonary Artery Pressure Diastolic"),
    ("PAP Mean", "MPAP", "Pulmonary Artery Pressure Mean"),
    ("PCWP Mean", "PCWPM", "Pulmonary Capillary Wedge Pressure Mean"),
    ("CO", "CO", "Cardiac Output"),
]

# CVC visit configurations (from cvc_export.py VISIT_CONFIG)
CVC_VISIT_CONFIG = {
    "Screening": {"prefix": "SBV_CVC_CVORRES_", "date_col": "SBV_CVC_PRSTDTC_CVC"},
    "Pre-procedure": {"prefix": "TV_CVC_PRE_POST_CVORRES_PRE_", "date_col": "TV_CVC_PRE_POST_PRSTDTC_PRE_CVC"},
    "Post-procedure": {"prefix": "TV_CVC_PRE_POST_CVORRES_POST_", "date_col": "TV_CVC_PRE_POST_PRSTDTC_POST_CVC"},
}

# Echo parameters (from echo_export.py SEMANTIC_PATTERNS)
ECHO_PARAMS = [
    ("TR Color Grade", "tr_color_grade", "TR evaluated based on color doppler"),
    ("TR Hepatic Grade", "tr_hepatic_grade", "TR evaluated based on systolic hepatic backflow"),
    ("TR VCW Grade", "tr_vcw_grade", "Vena Contracta Width"),
    ("TR EROA Grade", "tr_eroa_grade", "TR evaluated based on EROA"),
    ("EROA", "eroa_numeric", "Effective Regurgitant Orifice Area"),
    ("Regurg Vol", "regurg_vol", "Regurgitant Volume"),
    ("RV Basal Diam", "rv_basal_diam", "RV Basal Diameter"),
    ("RV Mid Diam", "rv_mid_diam", "RV Mid Diameter"),
    ("RV Long Diam", "rv_long_diam", "RV Longitudinal Diameter"),
    ("RVFAC", "rvfac", "RV Fractional Area Change"),
    ("LVEF", "lvef", "Left Ventricular Ejection Fraction"),
    ("RVEF", "rvef", "Right Ventricular Ejection Fraction"),
    ("TAPSE", "tapse", "Tricuspid Annular Plane Systolic Excursion"),
    ("S Wave", "s_wave", "Pulsed Doppler S wave"),
    ("IVC Diam", "ivc_diam", "Inferior Vena Cava Diameter"),
    ("Cardiac Output", "cardiac_output", "Cardiac Output"),
    ("Delta Psys", "delta_psys", "ΔPsys on Tricuspid Valve"),
]

# Echo visit configurations (from echo_export.py VISIT_CONFIG)
ECHO_VISIT_CONFIG = {
    "Screening": {"prefix": "SBV_ECHO_SPONSOR_", "date_col": "SBV_ECHO_SPONSOR_EGDTC"},
    "1-day pre-procedure": {"prefix": "TV_ECHO_1DPP_SPONSOR_", "date_col": "TV_ECHO_1DPP_SPONSOR_EGDTC"},
    "Pre-procedure": {"prefix": "TV_ECHO_PRE_POST_SPONSOR_", "suffix_filter": "_PRE", "date_col": "TV_ECHO_PRE_POST_SPONSOR_EGDTC_PRE"},
    "Post-procedure": {"prefix": "TV_ECHO_PRE_POST_SPONSOR_", "suffix_filter": "_POST", "date_col": "TV_ECHO_PRE_POST_SPONSOR_EGDTC_POST_1"},
    "1-day post-procedure": {"prefix": "TV_ECHO_1D_SPONSOR_", "date_col": "TV_ECHO_1D_SPONSOR_EGDTC"},
    "Discharge": {"prefix": "DV_ECHO_1D_SPONSOR_", "date_col": "DV_ECHO_1D_SPONSOR_EGDTC"},
    "30-day": {"prefix": "FU1M_ECHO_1D_SPONSOR_", "date_col": "FU1M_ECHO_1D_SPONSOR_EGDTC"},
    "6M": {"prefix": "FU6M_ECHO_1D_SPONSOR_", "date_col": "FU6M_ECHO_1D_SPONSOR_EGDTC"},
    "1Y": {"prefix": "FU1Y_ECHO_1D_SPONSOR_", "date_col": "FU1Y_ECHO_1D_SPONSOR_EGDTC"},
    "2Y": {"prefix": "FU2Y_ECHO_1D_SPONSOR_", "date_col": "FU2Y_ECHO_1D_SPONSOR_EGDTC"},
    "4Y": {"prefix": "FU4Y_ECHO_1D_SPONSOR_", "date_col": "FU4Y_ECHO_1D_SPONSOR_EGDTC"},
}

# Echo semantic patterns for finding columns (from echo_export.py)
ECHO_SEMANTIC_PATTERNS = {
    "tr_hepatic_grade": ["TR evaluated based on systolic hepatic backflow"],
    "tr_color_grade": ["TR evaluated based on color doppler on the Trillium"],
    "tr_vcw_grade": ["Vena Contracta Width"],
    "tr_eroa_grade": ["TR evaluated based on EROA"],
    "eroa_numeric": ["EROA, cm2"],
    "regurg_vol": ["regurgitant volume, ml"],
    "rv_basal_diam": ["RV basal diameter"],
    "rv_mid_diam": ["RV mid diameter"],
    "rv_long_diam": ["RV longitudinal diameter"],
    "rvfac": ["RVFAC", "RV FAC", "RV Fractional Area Change"],
    "lvef": ["LVEF"],
    "rvef": ["RVEF"],
    "tapse": ["Tricuspid annular plane systolic excursion", "TAPSE"],
    "s_wave": ["Pulsed Doppler S wave"],
    "ivc_diam": ["IVC diameter"],
    "cardiac_output": ["Cardiac Output"],
    "delta_psys": ["ΔPsys on Tricuspid Valve"],
}

# 6MWD visit configuration
SIXMWD_VISIT_CONFIG = {
    "Screening": {"prefix": "SBV_", "date_col": "SBV_SV_SVSTDTC"},
    "30D": {"prefix": "FU1M_", "date_col": "FU1M_SV_SVSTDTC"},
    "6M": {"prefix": "FU6M_", "date_col": "FU6M_SV_SVSTDTC"},
    "1Y": {"prefix": "FU1Y_", "date_col": "FU1Y_SV_SVSTDTC"},
    "2Y": {"prefix": "FU2Y_", "date_col": "FU2Y_SV_SVSTDTC"},
    "4Y": {"prefix": "FU4Y_", "date_col": "FU4Y_SV_SVSTDTC"},
}

# KCCQ parameters — matches actual CRF fields (QSORRES_KCCQ_*)
KCCQ_PARAMS = [
    ("Overall Summary", "OVERALL", "KCCQ Overall Summary Score"),
    ("Clinical Summary", "CLINICAL", "KCCQ Clinical Summary Score"),
]

# KCCQ visit configuration (same as 6MWD)
KCCQ_VISIT_CONFIG = SIXMWD_VISIT_CONFIG.copy()


# ============================================================================
# ASSESSMENT CATEGORIES - Master list for UI dropdowns
# ============================================================================

ASSESSMENT_CATEGORIES = {
    "Labs - CBC": {"type": "lab", "panel": "CBC", "params": LAB_PARAMS["CBC"], "visits": LAB_VISIT_CONFIG},
    "Labs - BMP": {"type": "lab", "panel": "BMP", "params": LAB_PARAMS["BMP"], "visits": LAB_VISIT_CONFIG},
    "Labs - LFP": {"type": "lab", "panel": "LFP", "params": LAB_PARAMS["LFP"], "visits": LAB_VISIT_CONFIG},
    "Labs - Enzymes": {"type": "lab", "panel": "ENZ", "params": LAB_PARAMS["ENZ"], "visits": LAB_VISIT_CONFIG},
    "Labs - Biomarkers": {"type": "lab", "panel": "BM", "params": LAB_PARAMS["BM"], "visits": LAB_VISIT_CONFIG},
    "Labs - Coagulation": {"type": "lab", "panel": "COA", "params": LAB_PARAMS["COA"], "visits": LAB_VISIT_CONFIG},
    "6MWD": {"type": "6mwd", "params": [("Distance", "6MWD", "6 Minute Walk Distance")], "visits": SIXMWD_VISIT_CONFIG},
    "KCCQ": {"type": "kccq", "params": KCCQ_PARAMS, "visits": KCCQ_VISIT_CONFIG},
    "CVC": {"type": "cvc", "params": CVC_PARAMS, "visits": CVC_VISIT_CONFIG},
    "Echo (CoreLab)": {"type": "echo", "params": ECHO_PARAMS, "visits": ECHO_VISIT_CONFIG},
}


class AssessmentDataExtractor:
    """Extracts assessment data across patients and visits."""
    
    def __init__(self, df_main, labels_map=None):
        self.df_main = df_main
        self.labels_map = labels_map or {}
        self._echo_col_cache = {}
    
    def is_valid_value(self, val):
        """Check if value is valid (not NaN, not empty string)."""
        if pd.isna(val):
            return False
        val_str = str(val).strip().lower()
        return val_str not in ['', 'nan', 'none']
    
    def get_patient_row(self, patient_id):
        """Get the data row for a patient."""
        rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if rows.empty:
            return None
        return rows.iloc[0]
    
    # -------------------------------------------------------------------------
    # Lab Data Extraction
    # -------------------------------------------------------------------------
    
    def get_lab_value(self, row, visit_name, panel, param_code):
        """Get lab value for a specific visit and parameter."""
        config = LAB_VISIT_CONFIG.get(visit_name)
        if not config:
            return None
        
        prefix = config["prefix"]
        
        # Handle special column patterns for different visits
        if prefix == "TV_LB_":
            # Treatment visit (Baseline): TV_LB_{panel}_DV_LBORRES_{param_code}
            col_name = f"TV_LB_{panel}_DV_LBORRES_{param_code}"
            if param_code == "eGFR":
                col_name = f"TV_LB_{panel}_DV_eGFR"
        else:
            # Standard pattern: {prefix}{panel}_LBORRES_{param_code}
            col_name = f"{prefix}{panel}_LBORRES_{param_code}"
            if param_code == "eGFR":
                col_name = f"{prefix}{panel}_eGFR"
        
        if col_name not in row.index:
            return None
        
        val = row[col_name]
        if not self.is_valid_value(val):
            return None
        
        # For Baseline, values may be pipe-separated (multiple days) - take first value
        val_str = str(val).strip()
        if '|' in val_str:
            parts = val_str.split('|')
            val_str = parts[0].strip() if parts else ""
        
        return val_str if val_str and val_str.lower() not in ['nan', 'none'] else None
    
    # -------------------------------------------------------------------------
    # CVC Data Extraction
    # -------------------------------------------------------------------------
    
    def get_cvc_value(self, row, visit_name, param_code):
        """Get CVC value for a specific visit and parameter."""
        config = CVC_VISIT_CONFIG.get(visit_name)
        if not config:
            return None
        
        prefix = config["prefix"]
        col_name = f"{prefix}{param_code}"
        
        if col_name not in row.index:
            return None
        
        val = row[col_name]
        if not self.is_valid_value(val):
            return None
        
        # Return as integer for pressure values, float for CO
        try:
            num_val = float(val)
            if param_code == "CO":
                return str(round(num_val, 1))
            return str(int(round(num_val)))
        except (ValueError, TypeError):
            return str(val).strip()
    
    # -------------------------------------------------------------------------
    # Echo Data Extraction
    # -------------------------------------------------------------------------
    
    def find_echo_column(self, visit_name, semantic_key):
        """Find echo column for a visit and semantic key."""
        cache_key = (visit_name, semantic_key)
        if cache_key in self._echo_col_cache:
            return self._echo_col_cache[cache_key]
        
        config = ECHO_VISIT_CONFIG.get(visit_name)
        if not config:
            return None
        
        prefix = config["prefix"]
        suffix_filter = config.get("suffix_filter")
        patterns = ECHO_SEMANTIC_PATTERNS.get(semantic_key, [])
        
        candidates = []
        for col in self.df_main.columns:
            if not str(col).startswith(prefix):
                continue
            
            # Handle suffix filter for PRE/POST
            if suffix_filter:
                if suffix_filter == "_PRE":
                    if "FAORRES_PRE_" not in col and "_PRE_ECHO" not in col:
                        continue
                    if "FAORRES_POST_" in col or "_POST_ECHO" in col:
                        continue
                elif suffix_filter == "_POST":
                    if "FAORRES_POST_" not in col and "_POST_ECHO" not in col:
                        continue
            
            # Check label patterns
            label = str(self.labels_map.get(col, col)).lower()
            for p in patterns:
                if p.lower() in label:
                    candidates.append(col)
                    break
        
        # Prioritize _SP suffix columns
        sp_cols = [c for c in candidates if c.endswith('_SP')]
        result = sp_cols[0] if sp_cols else (candidates[0] if candidates else None)
        
        self._echo_col_cache[cache_key] = result
        return result
    
    def get_echo_value(self, row, visit_name, semantic_key):
        """Get echo value for a specific visit and parameter."""
        col_name = self.find_echo_column(visit_name, semantic_key)
        if not col_name or col_name not in row.index:
            return None
        
        val = row[col_name]
        if not self.is_valid_value(val):
            return None
        
        return str(val).strip()
    
    # -------------------------------------------------------------------------
    # 6MWD Data Extraction
    # -------------------------------------------------------------------------
    
    def get_6mwd_value(self, row, visit_name):
        """Get 6MWD value for a specific visit."""
        config = SIXMWD_VISIT_CONFIG.get(visit_name)
        if not config:
            return None
        
        prefix = config["prefix"]
        
        # Try different column patterns (FU visits use _FU_ infix)
        col_patterns = [
            f"{prefix}6MWT_FU_FTORRES_DIS",  # FU visit pattern
            f"{prefix}6MWT_FTORRES_DIS",      # Screening/standard pattern
            f"{prefix}6MWT_FTORRES_6MWD",     # Legacy pattern
        ]
        
        for col_name in col_patterns:
            if col_name in row.index:
                val = row[col_name]
                if self.is_valid_value(val):
                    return str(val).strip()
        
        return None
    
    # -------------------------------------------------------------------------
    # KCCQ Data Extraction
    # -------------------------------------------------------------------------
    
    def get_kccq_value(self, row, visit_name, param_code):
        """Get KCCQ value for a specific visit and parameter."""
        config = KCCQ_VISIT_CONFIG.get(visit_name)
        if not config:
            return None
        
        prefix = config["prefix"]
        col_name = f"{prefix}KCCQ_QSORRES_KCCQ_{param_code}"
        
        if col_name not in row.index:
            return None
        
        val = row[col_name]
        if not self.is_valid_value(val):
            return None
        
        return str(val).strip()
    
    # -------------------------------------------------------------------------
    # Generic Extraction Interface
    # -------------------------------------------------------------------------
    
    def get_value(self, patient_id, category_name, param_code, visit_name):
        """Get assessment value for any category/parameter/visit combination."""
        row = self.get_patient_row(patient_id)
        if row is None:
            return None
        
        category = ASSESSMENT_CATEGORIES.get(category_name)
        if not category:
            return None
        
        cat_type = category["type"]
        
        if cat_type == "lab":
            panel = category["panel"]
            return self.get_lab_value(row, visit_name, panel, param_code)
        elif cat_type == "cvc":
            return self.get_cvc_value(row, visit_name, param_code)
        elif cat_type == "echo":
            return self.get_echo_value(row, visit_name, param_code)
        elif cat_type == "6mwd":
            return self.get_6mwd_value(row, visit_name)
        elif cat_type == "kccq":
            return self.get_kccq_value(row, visit_name, param_code)
        
        return None
    
    def generate_table(self, patient_ids, category_name, param_code, visit_names):
        """
        Generate a table with patient IDs as rows and visit values as columns.
        
        Returns: pandas DataFrame with columns ['Patient'] + visit_names
        """
        data = []
        
        for patient_id in patient_ids:
            row_data = {"Patient": patient_id}
            for visit_name in visit_names:
                val = self.get_value(patient_id, category_name, param_code, visit_name)
                row_data[visit_name] = val if val else ""
            data.append(row_data)
        
        columns = ["Patient"] + list(visit_names)
        return pd.DataFrame(data, columns=columns)
