"""
Clinical Data Master - Configuration Module
============================================
Centralized domain rules and mappings for the clinical viewer application.
Extracted from clinical_viewer1.py for maintainability.
"""

# --- 1. CONFIGURATION: VISIT PREFIXES ---
VISIT_MAP = {
    "SBV": "Baseline",
    "TV": "Treatment",
    "DV": "Discharge Visit",
    "FU1M": "30-Day Follow Up", "FU3M": "3-Month Follow Up (Remote)",
    "FU6M": "6-Month Follow Up", "FU1Y": "1-Year Follow Up",
    "FU2Y": "2-Year Follow Up", "FU3Y": "3-Year Follow Up (Remote)",
    "FU4Y": "4-Year Follow Up", "FU5Y": "5-Year Follow Up (Remote)",
    "UV": "Unscheduled", "LOGS": "Logs"
}

# --- 2. ASSESSMENT RULES ---
# Each tuple: (regex_pattern, category, form_name)
# Order matters — first match wins
ASSESSMENT_RULES = [
    # Additional Tests (uses specific pattern prefix)
    (r"LOGS_LB_PR_OTH_PRORRES", "Procedures", "Additional Laboratory / Diagnostic Tests"),
    (r"LOGS_LB_PR_OTH_ORRES", "Laboratory", "Additional Laboratory / Diagnostic Tests"),
    (r"LOGS_LB_PR_OTH_LBORRES", "Laboratory", "Additional Laboratory / Diagnostic Tests"),
    (r"LOGS_AE_LBREF", "Laboratory", "Additional Laboratory / Diagnostic Tests"),
    (r"LOGS_AE_PRREF", "Procedures", "Additional Laboratory / Diagnostic Tests"),
    
    # Admin
    (r"ELIG",  "Admin", "Eligibility Confirmation and Planned Procedure Date"),
    (r"IE",    "Admin", "Inclusion/Exclusion Criteria"),
    (r"ICF",   "Admin", "ICF procedure"),
    (r"_SV_",  "Admin", "Visit Date"), 

    # Procedures - Timing must be before ECG/CVC
    (r"_PR_TIM_", "Procedures", "Procedure form"),  # Procedure Timing is part of Procedure form
    (r"CVC.*PRE|CVC.*POST",   "Procedures", "Cardiac and Venous Catheterization – Pre- and Post-procedure"),
    (r"CVC",   "Procedures", "Cardiac and Venous Catheterization"),
    (r"TV_.*ECG.*POST", "Procedures", "Standard 12-lead ECG-Pre and Post procedure"),
    (r"TV_.*ECG.*PRE",  "Procedures", "Standard 12-lead ECG-Pre and Post procedure"),
    (r"ECG",            "Procedures", "Standard 12-lead ECG"), 
    (r"TRRI",  "Procedures", "Tricuspid Re-intervention"),
    (r"CVPHM", "Procedures", "CVP Hemodynamic Measurement"),
    (r"_PR_",  "Procedures", "Procedure form"), 
    
    # Imaging - Core Lab (with _SP or _CORE suffix)
    (r"TV_.*ECHO.*1DPP.*(_SP|_CORE)",        "Imaging (Core Lab)", "Echocardiography – 1 day prior the procedure - Core lab"),
    (r"TV_.*ECHO.*1D.*(_SP|_CORE)",          "Imaging (Core Lab)", "Echocardiography – 1-day post procedure - Core lab"),
    (r"TV_.*ECHO.*(PRE|POST).*(_SP|_CORE)",  "Imaging (Core Lab)", "Echocardiography – Pre and Post procedure - Core lab"),
    (r"TV_.*ECHO.*(_SP|_CORE)",              "Imaging (Core Lab)", "Echocardiography – Core lab"),
    (r"ECHO.*(_SP|_CORE)",                   "Imaging (Core Lab)", "Echocardiography – Core lab"),
    (r"ECHO.*SPONSOR",                       "Imaging (Core Lab)", "Echocardiography – Core lab"),

    # Imaging - Site
    (r"TV_.*ECHO.*1DPP",        "Imaging (Site)", "Echocardiography – 1 day prior the procedure"),
    (r"TV_.*ECHO.*1D",          "Imaging (Site)", "Echocardiography – 1-day post procedure"),
    (r"TV_.*ECHO.*(PRE|POST)",  "Imaging (Site)", "Echocardiography – Pre and Post procedure"),
    (r"TV_.*ECHO",              "Imaging (Site)", "Echocardiography"),
    (r"ECHO",                   "Imaging (Site)", "Echocardiography"),

    (r"_AG_",   "Imaging (Site)", "Angiography – Pre and Post procedure"),  # Angiography form
    (r"CMR",  "Imaging", "CMR Imaging"),
    (r"CCTA", "Imaging", "Cardiac CT Angiogram"),

    # Clinical Assessments
    (r"HE_GRADE|ENCEPH|LFP_HE|RS_EG", "Clinical Assessments", "Encephalopathy Grade"),
    (r"_VS",    "Clinical Assessments", "Vital signs"),
    (r"_PE",    "Clinical Assessments", "Physical Examination"),
    (r"6MWT",   "Clinical Assessments", "Exercise Tolerance (6MWT)"),
    (r"CFSS",   "Clinical Assessments", "Clinical Frailty Scale"),
    (r"_FS_",   "Clinical Assessments", "Functional Status (NYHA)"),
    (r"MNA",    "Clinical Assessments", "Mini Nutrition Assessment (MNA)"),
    (r"KCCQ",   "Questionnaires",       "Kansas City Cardiomyopathy Questionnaire (KCCQ)"),
    (r"RS_PGA", "Clinical Assessments", "Physician Global Assessment"),
    
    # Laboratory
    (r"LB_CBC",  "Laboratory", "CBC and platelets count"),
    (r"LB_BMP",  "Laboratory", "Basic metabolic panel and eGFR CKD-EPI (2021)"),
    (r"LB_LFP",  "Laboratory", "Liver function panel"),
    (r"LB_COA",  "Laboratory", "Coagulation study"),
    (r"LB_ENZ",  "Laboratory", "Blood enzymes"),
    (r"LB_PREG", "Laboratory", "Pregnancy test"),
    (r"LB_BM",   "Laboratory", "Biomarkers"),
    (r"LB_ACT",  "Laboratory", "ACT lab results"),
    (r"LB_ADD",  "Laboratory", "Additional Laboratory / Diagnostic Tests"),
    
    # History
    (r"_DM",  "History", "Demographics"),
    (r"_MH",  "History", "Medical History"),
    (r"_CVH", "History", "Cardiovascular History"),
    (r"_HFH", "History", "Heart Failure History"),
    (r"HMEH", "History", "Hospitalization and Medical Events History"),

    # Risk Scores
    (r"TRS",  "Risk Scores", "Trio Score for Tricuspid Regurgitation Risk"),
    (r"STSS", "Risk Scores", "Society of Thoracic Surgeons Score"),

    # Safety
    (r"_DDF",         "Safety", "Device Deficiency Form"),
    (r"_AE|AEACN",    "Safety", "Adverse Event"),
    (r"_CM",          "Safety", "Concomitant Medications"),
    (r"PTHME",        "Safety", "Post-Treatment Hospitalizations/Medical Events"),
    (r"DTF|DEATH|DTH", "Safety", "Death") 
]

# --- 3. CONDITIONAL SKIP RULES ---
# When a trigger field has the specified value, hide the target fields
CONDITIONAL_SKIPS = {
    "FTORRES_COMPL": {"trigger_value": "completed", "targets": ["REASNC", "REASND"]},
    "FTORRES_INC":   {"trigger_value": "yes",       "targets": ["INCD"]},
    "PESTAT":        {"trigger_value": "yes",       "targets": ["REASND"]},
    "VSSTAT":        {"trigger_value": "yes",       "targets": ["REASND"]},
    "RSSTAT":        {"trigger_value": "yes",       "targets": ["REASND"]},
    "QSSTAT":        {"trigger_value": "yes",       "targets": ["REASND"]},
    "PERF":          {"trigger_value": "yes",       "targets": ["REASND"]},
    # When full DOB is present, skip partial/year-only date fields
    "BRTHDAT":       {"trigger_value": "*ANY*",     "targets": ["BRTHYR", "AGE_YR", "DOB_YR", "PARTIAL", "BRTHDAT_YEAR", "BRTHDAT_PARTIAL"]},
    # Childbearing potential is only for females - skip if male
    "SEX":           {"trigger_value": "male",      "targets": ["CHILDPOT", "F_CHILDPOT", "NFFORRS_F"]},
    # When race is reported (including unknown/not reported), skip specific race entries
    "RACE":          {"trigger_value": "*ANY*",     "targets": ["RACE_AIAN", "RACE_ASIA", "RACE_BLAA", "RACE_NHPI", "RACE_WHIT", "RACE_OTH"]},
    # When ethnicity is reported, skip other ethnicity fields  
    "ETHNIC":        {"trigger_value": "*ANY*",     "targets": ["ETHNIC_OTH"]},
    # When vital signs result value exists, skip the "not done" status flag for that parameter
    "VSORRES":       {"trigger_value": "*ANY*",     "targets": ["VSSTAT"]},
    # Skip specific vital sign "not done" status when result exists
    "VSORRES_RISP":  {"trigger_value": "*ANY*",     "targets": ["VSSTAT_RISP"]},  # Respiratory rate
    "VSORRES_HR":    {"trigger_value": "*ANY*",     "targets": ["VSSTAT_HR"]},    # Heart rate
    "VSORRES_TEMP":  {"trigger_value": "*ANY*",     "targets": ["VSSTAT_TEMP"]},  # Temperature
    "VSORRES_DIABP": {"trigger_value": "*ANY*",     "targets": ["VSSTAT_DIABP"]}, # Diastolic BP
    "VSORRES_SYSBP": {"trigger_value": "*ANY*",     "targets": ["VSSTAT_SYSBP"]}, # Systolic BP
    "VSORRES_WEIGHT":{"trigger_value": "*ANY*",     "targets": ["VSSTAT_WEIGHT"]},# Weight
    "VSORRES_HEIGHT":{"trigger_value": "*ANY*",     "targets": ["VSSTAT_HEIGHT"]},# Height
}

# --- 4. VISIT SCHEDULE (for date tracking) ---
VISIT_SCHEDULE = [
    ("SBV_SV_SVSTDTC", "Baseline/Screening"),
    ("TV_PR_SVDTC", "Treatment"),
    ("DV_SV_SVSTDTC", "Discharge Visit"),
    ("FU1M_SV_SVSTDTC", "30-Day FU"),
    ("FU3M_SV_SVSTDTC", "3-Month FU"),
    ("FU6M_SV_SVSTDTC", "6-Month FU"),
    ("FU1Y_SV_SVSTDTC", "1-Year FU"),
    ("FU2Y_SV_SVSTDTC", "2-Year FU"),
    ("FU3Y_SV_SVSTDTC", "3-Year FU"),
    ("FU4Y_SV_SVSTDTC", "4-Year FU"),
    ("FU5Y_SV_SVSTDTC", "5-Year FU"),
]
