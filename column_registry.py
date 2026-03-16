"""
Centralized Column Registry
============================
Single source of truth for EDC column names, validated against the annotated CRF
(Innoventric_CLD-048, eCRF v2.0, exported 06-Nov-2025).

Usage:
    from column_registry import COL, VISITS, get_col

    # Direct access:
    procedure_date = row[COL.PROCEDURE_DATE]

    # Visit-prefixed access:
    col = get_col(VISITS.SCREENING, 'LB_CBC_LBORRES_HGB')
    # -> 'SBV_LB_CBC_LBORRES_HGB'
"""


# =============================================================================
# Visit Prefixes
# =============================================================================

class VISITS:
    """Visit prefix constants. Match actual ProjectToOneFile column prefixes."""
    SCREENING = "SBV"
    TREATMENT = "TV"
    DISCHARGE = "DV"
    FU_30D = "FU1M"
    FU_3M = "FU3M"   # Remote/phone visit — no labs or echo
    FU_6M = "FU6M"
    FU_1Y = "FU1Y"
    FU_2Y = "FU2Y"
    FU_3Y = "FU3Y"   # Remote/phone visit
    FU_4Y = "FU4Y"
    FU_5Y = "FU5Y"   # Remote/phone visit
    UNSCHEDULED = "UV"
    LOGS = "LOGS"

    # On-site visits that collect labs and echo
    ONSITE = [SCREENING, TREATMENT, DISCHARGE, FU_30D, FU_6M, FU_1Y, FU_2Y, FU_4Y]

    # Remote visits — only NYHA, CFS, KCCQ, basic status
    REMOTE = [FU_3M, FU_3Y, FU_5Y]

    # All visits (ordered chronologically)
    ALL = [SCREENING, TREATMENT, DISCHARGE, FU_30D, FU_3M, FU_6M,
           FU_1Y, FU_2Y, FU_3Y, FU_4Y, FU_5Y]


# =============================================================================
# Column Name Constants (CRF-validated)
# =============================================================================

class COL:
    """Column name constants, validated against annotated CRF and actual data export."""

    # --- Identifiers (no visit prefix) ---
    SCREENING_NUM = "Screening #"
    RANDOMIZATION_NUM = "Randomization #"
    INITIALS = "Initials"
    SITE_NUM = "Site #"
    STATUS = "Status"
    ROW_NUMBER = "Row number"
    TEMPLATE_NUMBER = "Template number"

    # --- Visit Date ---
    VISIT_DATE = "SV_SVSTDTC"  # Prefixed: {visit}_SV_SVSTDTC

    # --- Procedure (Treatment Visit) ---
    PROCEDURE_DATE = "TV_PR_PRSTDTC"         # Implant procedure date
    TREATMENT_VISIT_DATE = "TV_PR_SVDTC"     # Treatment visit date

    # --- Demographics (Screening/Baseline) ---
    AGE = "SBV_ELIG_AGE"
    SEX = "SBV_ELIG_SEX"
    BIRTH_DATE = "SBV_DM_BRTHDAT"

    # --- Eligibility ---
    ELIGIBILITY_DECISION = "SBV_ELIG_IEORRES_CONF5"  # Eligibility Committee decision
    PLANNED_PROCEDURE_DATE = "SBV_ELIG_PRSTDTC_PLAN"

    # --- ICF ---
    ICF_VERSION = "SBV_ICF_ICVERSION"
    ICF_DATE = "SBV_ICF_RFICDTC"

    # --- Adverse Events (LOGS prefix) ---
    AE_TERM = "LOGS_AE_AETERM"
    AE_START_DATE = "LOGS_AE_AESTDTC"
    AE_END_DATE = "LOGS_AE_AEENDTC"
    AE_ONGOING = "LOGS_AE_AEONGO"
    AE_INTERVAL = "LOGS_AE_AEINT"
    AE_OUTCOME = "LOGS_AE_AEOUT"
    AE_SEVERITY = "LOGS_AE_AESEV"
    AE_REL_PKG = "LOGS_AE_AEREL1"           # Relationship: PKG Trillium
    AE_REL_DELIVERY = "LOGS_AE_AEREL2"       # Relationship: PKG Delivery System
    AE_REL_HANDLE = "LOGS_AE_AEREL3"         # Relationship: PKG Handle
    AE_REL_PROCEDURE = "LOGS_AE_AEREL4"      # Relationship: index procedure
    AE_REL_OTHER = "LOGS_AE_AEREL5"          # Other relationship
    AE_REL_OTHER_SPEC = "LOGS_AE_AEREL5_OTH"
    AE_SERIOUS = "LOGS_AE_AESER"
    AE_DESCRIPTION = "LOGS_AE_AETERM_COMM"
    AE_SAE_DESCRIPTION = "LOGS_AE_AETERM_COMM1"
    # Seriousness criteria
    AE_CRIT_DEATH = "LOGS_AE_AESDTH"
    AE_CRIT_HOSP = "LOGS_AE_AESHOSP"
    AE_CRIT_LIFE = "LOGS_AE_AESLIFE"
    AE_CRIT_DISAB = "LOGS_AE_AESDISAB"
    AE_CRIT_INTERV = "LOGS_AE_AESMIE"
    AE_CRIT_CONGENIT = "LOGS_AE_AESCONG"
    AE_CRIT_OTHER = "LOGS_AE_AESMIE_OTH"
    # Actions taken
    AE_ACT_NONE = "LOGS_AE_AEACN_NONE"
    AE_ACT_MED = "LOGS_AE_AEACN_CM"
    AE_ACT_HOSP = "LOGS_AE_AEACN_HO"
    AE_ACT_SURG = "LOGS_AE_AEACN_SURG"
    AE_ACT_SUBINT = "LOGS_AE_AEACN_SUBI"
    AE_ACT_OTHER = "LOGS_AE_AEACN_OTH"
    AE_ACT_LAB = "LOGS_AE_AEACN_LT"
    AE_ACT_DIAG = "LOGS_AE_AEACN_PR"
    # Hospitalization within AE
    AE_HOSP_ADMIT = "LOGS_AE_HOSPSTDAT"
    AE_HOSP_DISCHARGE = "LOGS_AE_HOSPENDAT"
    AE_HOSP_ONGOING = "LOGS_AE_HOSP_ONGO"
    # Death within AE
    AE_DEATH_DATE = "LOGS_AE_DTHDAT"
    AE_DEATH_CAUSE = "LOGS_AE_DTHCAUSE"
    # SAE reporting
    AE_REPORT_TYPE = "LOGS_AE_AEREP"
    AE_REPORT_REF = "LOGS_AE_AEREPREF"
    AE_REPORT_DATE = "LOGS_AE_AEREPDAT"
    AE_AWARE_DATE = "LOGS_AE_AEAWDAT"
    # Additional references
    AE_LAB_REF = "LOGS_AE_LBREF"
    AE_DIAG_REF = "LOGS_AE_PRREF"

    # --- Death Form ---
    DEATH_DATE = "LOGS_DTH_DDDTC"
    DEATH_CATEGORY = "LOGS_DTH_DDRESCAT"     # Mortality classification
    DEATH_REASON = "LOGS_DTH_DDORRES"        # Reason of death

    # --- Concomitant Medications ---
    CM_DRUG = "LOGS_CM_CMTRT"
    CM_INDICATION = "LOGS_CM_CMINDC"
    CM_START = "LOGS_CM_CMSTDAT"
    CM_END = "LOGS_CM_CMENDAT"
    CM_ONGOING = "LOGS_CM_CMONGO"
    CM_DOSE = "LOGS_CM_CMDOSE"
    CM_UNITS = "LOGS_CM_CMDOSU"
    CM_ROUTE = "LOGS_CM_CMROUTE"
    CM_FREQUENCY = "LOGS_CM_CMDOSFRQ"

    # --- KCCQ (visit-prefixed) ---
    KCCQ_STATUS = "KCCQ_QSSTAT_KCCQ"
    KCCQ_DATE = "KCCQ_QSDTC_KCCQ"
    KCCQ_OVERALL = "KCCQ_QSORRES_KCCQ_OVERALL"
    KCCQ_CLINICAL = "KCCQ_QSORRES_KCCQ_CLINICAL"

    # --- NYHA (visit-prefixed) ---
    NYHA_STATUS = "FS_RSSTAT_FSNYHA"
    NYHA_RESULT = "FS_RSORRES_FSNYHA"

    # --- 6MWT (visit-prefixed) ---
    SIXMWT_STATUS = "6MWT_FTSTAT_SIXMW1"
    SIXMWT_DISTANCE = "6MWT_FTORRES_DIS"
    SIXMWT_TIME = "6MWT_FTORRES_TIM"

    # --- Clinical Frailty Scale (visit-prefixed) ---
    CFS_STATUS = "CFSS_RSSTAT_CFSS"
    CFS_RESULT = "CFSS_RSORRES_CFSS"

    # --- Vital Signs (visit-prefixed) ---
    VS_DATE = "VS_VSDTC"
    VS_SYSTOLIC = "VS_VSORRES_SYSBP"
    VS_DIASTOLIC = "VS_VSORRES_DIABP"
    VS_HR = "VS_VSORRES_HR"
    VS_WEIGHT = "VS_VSORRES_WEIGHT"
    VS_HEIGHT = "VS_VSORRES_HEIGHT"
    VS_BMI = "VS_VSORRES_BMI"
    VS_TEMP = "VS_VSORRES_TEMP"
    VS_RESP = "VS_VSORRES_RESP"
    VS_EDEMA = "VS_CVORRES_EDEMA"
    VS_ASCITES = "VS_CVORRES_ASCITIS"  # CRF typo preserved


# =============================================================================
# Helper Functions
# =============================================================================

def get_col(visit_prefix: str, field_suffix: str) -> str:
    """Build a visit-prefixed column name.

    Args:
        visit_prefix: e.g., 'SBV', 'TV', 'FU1M'
        field_suffix: e.g., 'LB_CBC_LBORRES_HGB', 'VS_VSORRES_HR'

    Returns:
        Full column name, e.g., 'SBV_LB_CBC_LBORRES_HGB'
    """
    return f"{visit_prefix}_{field_suffix}"


def validate_columns(df_columns, required_cols, logger=None):
    """Check that expected columns exist in a DataFrame.

    Args:
        df_columns: DataFrame column index (df.columns)
        required_cols: List of column names to check
        logger: Optional logger for warnings

    Returns:
        Tuple of (found_cols, missing_cols)
    """
    col_set = set(df_columns)
    found = [c for c in required_cols if c in col_set]
    missing = [c for c in required_cols if c not in col_set]
    if missing and logger:
        logger.warning(f"Missing {len(missing)} expected columns: {missing[:10]}...")
    return found, missing


# Critical columns that must exist for core functionality
CRITICAL_COLUMNS = [
    COL.SCREENING_NUM,
    COL.SITE_NUM,
    COL.STATUS,
    COL.PROCEDURE_DATE,
    COL.TREATMENT_VISIT_DATE,
    COL.AE_TERM,
    COL.AE_START_DATE,
    COL.AE_ONGOING,
    COL.AE_OUTCOME,
    COL.AE_SEVERITY,
    COL.AE_SERIOUS,
    COL.ELIGIBILITY_DECISION,
    COL.DEATH_DATE,
]
