# Clinical Data Master — Developer Guide

## Overview
Desktop Python/Tkinter application for reviewing clinical trial data (Innoventric CLD-048 tricuspid valve study). Ingests MainEDC Excel exports and provides multi-window interface for data review, SDV tracking, AE analysis, HF hospitalization analysis, and data export.

## Setup
```bash
pip install pandas openpyxl matplotlib
python clinical_viewer1.py
```
Requires Python 3.9+. No build step.

## Data Files
The app expects these Excel exports from MainEDC:
- **ProjectToOneFile** (`*ProjectToOneFile*.xlsx`) — Main data: 8,758 columns, 1 row per patient on the "Main" sheet + repeating-form sheets (AE_732, CMTAB, PTHME_TABLE, etc.)
- **Modular** (`*Modular*.xlsx`) — SDV/CRF status data
- **CrfStatusHistory** (`*CrfStatusHistory*.xlsx`) — Historical CRF completion status

Column naming convention: `{VISIT}_{FORM}_{FIELD}` (e.g., `SBV_LB_CBC_LBORRES_HGB`)

### Visit Prefixes
| Prefix | Visit | Type |
|--------|-------|------|
| `SBV` | Screening/Baseline | On-site |
| `TV` | Treatment (procedure day) | On-site |
| `DV` | Discharge | On-site |
| `FU1M` | 30-Day Follow Up | On-site |
| `FU3M` | 3-Month Follow Up | **Remote/phone** (no labs/echo) |
| `FU6M` | 6-Month Follow Up | On-site |
| `FU1Y` | 1-Year Follow Up | On-site |
| `FU2Y` | 2-Year Follow Up | On-site |
| `FU3Y` | 3-Year Follow Up | **Remote/phone** |
| `FU4Y` | 4-Year Follow Up | On-site |
| `FU5Y` | 5-Year Follow Up | **Remote/phone** |
| `UV` | Unscheduled | On-site |
| `LOGS` | Logs (AE, CM, DTH, DDF) | Ongoing |

### Key CRF Field IDs
- Procedure date: `TV_PR_PRSTDTC` (implant procedure date)
- Treatment visit date: `TV_PR_SVDTC`
- Eligibility: `SBV_ELIG_IEORRES_CONF5` (Eligibility Committee decision)
- AE fields: `LOGS_AE_AETERM`, `LOGS_AE_AESTDTC`, `LOGS_AE_AEONGO`, etc.
- Death: `LOGS_DTH_DDDTC`, `LOGS_DTH_DDRESCAT`, `LOGS_DTH_DDORRES`
- KCCQ scores: `{prefix}KCCQ_QSORRES_KCCQ_OVERALL`, `{prefix}KCCQ_QSORRES_KCCQ_CLINICAL`
- Edema/Ascites: `{prefix}VS_CVORRES_EDEMA`, `{prefix}VS_CVORRES_ASCITIS` (note: "ASCITIS" typo is in CRF)

## Architecture
```
clinical_viewer1.py (5661 lines)  ← Main app, UI orchestration
├── config.py                     ← VISIT_MAP, ASSESSMENT_RULES, CONDITIONAL_SKIPS
├── column_registry.py            ← Centralized column name registry (CRF-validated)
├── view_builder.py               ← Tree view construction
├── toolbar_setup.py              ← Toolbar UI
├── data_sources.py               ← File detection and loading
├── data_comparator.py            ← File diff comparison
│
├── ae_manager.py + ae_ui.py      ← Adverse Event analysis
├── sdv_manager.py                ← Source Data Verification tracking
├── dashboard_manager.py + dashboard_ui.py  ← SDV/Gap dashboard
├── hf_hospitalization_manager.py ← HF event detection (fuzzy matching)
├── gap_analysis.py               ← Missing data report
├── assessment_data_table.py      ← Lab/assessment tabular view
├── patient_timeline.py           ← Patient timeline visualization
│
├── labs_export.py                ← Lab data Excel export
├── echo_export.py                ← Echocardiography Excel export
├── fu_highlights_export.py       ← Follow-up highlights + diuretic timeline
├── cvc_export.py                 ← Cardiac catheterization export
├── batch_export.py               ← Multi-patient export orchestrator
└── procedure_timing_export.py    ← Procedure timing matrix
```

## Key Patterns
- **All data loaded as `dtype=str`** — numeric parsing is manual throughout
- **SDV loading uses background thread** — other loading is synchronous
- **ViewBuilder caches views** keyed by (site, patient, view_mode, filters)
- **HF term matching uses `@lru_cache`** with word-boundary regex + `difflib.get_close_matches()`
- **Search has 300ms debounce** to avoid per-keystroke tree rebuilds

## Known Limitations
- No test suite — only debug scripts exist
- `clinical_viewer1.py` is a monolith (5661 lines, 131 methods in 1 class)
- Remaining `.iterrows()` calls at lines 468, 1286, 1364, 1765, 2057, 2131 (complex per-row logic)
- `add_gap()` method in ViewBuilder is defined but never called — gap detection needs wiring
