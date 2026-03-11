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
clinical_viewer1.py               ← Main app, UI orchestration (~885 lines)
├── data_loader.py                ← File detection, Excel loading, schema & cross-form validation
├── config.py                     ← VISIT_MAP, ASSESSMENT_RULES, CONDITIONAL_SKIPS
├── column_registry.py            ← Centralized column name registry (CRF-validated)
├── view_builder.py               ← Tree view construction
├── toolbar_setup.py              ← Toolbar UI
├── data_sources.py               ← Data source management UI
├── data_comparator.py            ← File diff comparison
│
├── data_matrix_builder.py        ← Column routing, type detection, matrix construction
├── matrix_display.py             ← Specialized matrix/table windows (AE, CM, MH, HFH, CVC, etc.)
├── assessment_table_ui.py        ← Assessment Data Table window (labs, vitals, etc.)
├── procedure_timing_ui.py        ← Procedure Timing Matrix window
│
├── echo_export_dialog.py         ← Echo export configuration dialog
├── cvc_export_dialog.py          ← CVC export configuration dialog
├── labs_export_dialog.py         ← Labs export configuration dialog
├── fu_highlights_dialog.py       ← FU Highlights export dialog
├── export_dialogs_ui.py          ← Re-export facade (backwards compat)
│
├── ae_manager.py + ae_ui.py      ← Adverse Event analysis + screen failure detection
├── sdv_manager.py                ← Source Data Verification tracking
├── dashboard_manager.py + dashboard_ui.py  ← SDV/Gap dashboard
├── hf_hospitalization_manager.py + hf_ui.py  ← HF event detection + UI
├── gap_analysis.py               ← Missing data report (uses cached gaps)
├── visit_schedule_ui.py          ← Visit schedule matrix window
├── assessment_data_table.py      ← Lab/assessment tabular view (data logic)
├── patient_timeline.py           ← Patient timeline visualization
│
├── base_exporter.py              ← Shared export base class (validation, formatting, ZIP)
├── labs_export.py                ← Lab data Excel export
├── echo_export.py                ← Echocardiography Excel export
├── fu_highlights_export.py       ← Follow-up highlights + diuretic timeline
├── cvc_export.py                 ← Cardiac catheterization export
├── batch_export.py               ← Multi-patient export orchestrator
├── procedure_timing_export.py    ← Procedure timing matrix
│
├── scripts/                      ← Debug/utility scripts (31 files)
│
└── tests/                        ← Unit test suite (268 tests)
    ├── test_ae_manager.py        ← AE column mapping, filters, stats, death details
    ├── test_hf_hospitalization_manager.py  ← HF term matching, boundaries, windows
    ├── test_data_loader.py       ← File detection, loading, cross-form validation
    ├── test_column_registry.py   ← Visit/column constants, get_col, validate_columns
    ├── test_gap_analysis.py      ← Gap detection, column mapping, gap count indexing
    ├── test_base_exporter.py     ← BaseExporter validation, formatting, export orchestration
    ├── test_data_matrix_builder.py ← Column classification, time/date parsing
    └── test_dashboard_manager.py ← Dashboard preprocessing, label mapping, aggregation
```

## Key Patterns
- **All data loaded as `dtype=str`** — numeric parsing is manual throughout
- **Data loading extracted to `data_loader.py`** — pure data, no UI dependencies
- **Cross-form validation runs on load** — checks fatal AE↔death form, procedure↔FU dates, AE onset timing
- **SDV loading uses background thread** — other loading is synchronous
- **ViewBuilder caches views** keyed by (site, patient, view_mode, filters)
- **HF term matching uses `@lru_cache`** with word-boundary regex + `difflib.get_close_matches()`
- **Search has 300ms debounce** to avoid per-keystroke tree rebuilds
- **Logging via `logging` module** — all debug/diagnostic output uses logger, not print()

## Running Tests
```bash
python -m unittest discover -s tests -v
```

## Key Patterns (continued)
- **UI extraction pattern** — extracted UI classes receive `app` reference, manage own Toplevel windows (e.g., `MatrixDisplay(self)`, `AssessmentTableWindow(self).show()`)
- **Generic matrix helper** — `MatrixDisplay._show_simple_matrix()` parameterizes window creation for MH, HFH, HMEH, CM-from-data displays
- **Generic export helper** — `MatrixDisplay._export_matrix(fmt, df, patient, prefix)` replaces 6 near-identical export methods
- **BaseExporter inheritance** — `echo_export`, `labs_export`, `cvc_export` share validation/formatting via `base_exporter.py`
- **Column type dispatch** — `data_matrix_builder.classify_column()` uses `_COL_TYPES` dict for clean repeating-form routing

## Known Limitations
- Some export dialog classes reference `self.app._is_screen_failure()` — screen failure detection remains on the main app class
