# Clinical Data Master - Handover Documentation

## Overview
This application ("Clinical Data Master" / `clinical_viewer1.py`) is a Python/Tkinter desktop application for viewing, analyzing, and auditing clinical trial data. It ingests complex Excel exports (specifically "ProjectToOneFile" or "SelectForms" formats) and provides a multi-tabbed interface for:
1.  **Clinical Viewer**: Hierarchical tree-view of all patient data.
2.  **AE Dashboard**: specialized analysis of Adverse Events.
3.  **SDV Dashboard**: Tracking Source Data Verification status.
4.  **HF Hospitalization**: Analysis of Heart Failure events.
5.  **Assessment Data**: Tabular view of specific lab/assessment values.

## Code Structure & Modules
The application has recently been modularized from a monolithic script.
*   **`clinical_viewer1.py`**: Entry point. Manages the main window, data loading, and tab orchestration.
    *   Key Class: `ClinicalViewerApp`.
*   **`ae_ui.py`**: UI logic for the "AE Dashboard" tab.
    *   Key Class: `AEWindow`.
    *   Handles: Dashboard visualization (matplotlib/cards) and Browser (treeview).
*   **`ae_manager.py`**: Business logic for AE data.
    *   Key Class: `AEManager`.
    *   Handles: Parsing, filtering (dates, pre-proc), statistics calculation (`get_summary_stats`).
    *   **CRITICAL**: Contains column mappings in `__init__`.
*   **`sdv_manager.py`**: Logic for Source Data Verification (SDV) tracking.
    *   Key Class: `SDVManager`.
    *   Handles: Reading `CrfStatusHistory` files to track who verified what and when.
*   **`dashboard_manager.py`**: Logic for the SDV Dashboard tab.
*   **`hf_hospitalization.py`**: specialized logic for HF event analysis.
*   **`assessment_data_table.py`**: Logic for the "Assessments" tab.
*   **`view_builder.py`**: Helper for constructing the main Clinical Viewer tree.
*   **`gap_analysis.py`**: Helper for identifying missing data/visits.

## Key Workflows

### 1. Data Loading
*   **Trigger**: Drag-and-drop or "Browse" in the main window.
*   **Logic**: `clinical_viewer1.py` -> `load_data`.
*   **Heuristics**:
    *   Identifies "Main" sheet by searching for "Main" or "Form" in sheet names.
    *   Identifies "AE" sheet by "AE" in name (and potentially "732").
    *   **Issue**: Recent crash caused by identifying wrong sheet. Fixed with `try-except` in `ae_manager` and better heuristics in user's workflow (loading correct file).

### 2. AE Analysis (Adverse Events)
*   **Dashboard**: Shows aggregate stats (Total, SAE, Device Related, etc.).
    *   **Filters**:
        *   exclude patients (text input).
        *   exclude Pre-Procedure AEs (checkbox, strictly `Onset < Procedure Date`).
*   **Browser**: List view of AEs.
    *   **Planned Features** (In Progress/Next):
        *   Split AE/SAE tables.
        *   Filter Screen Failures.
        *   Highlight Late AEs.
*   **Column Mappings**: Defined in `ae_manager.py`. Crucial for parsing.
    *   Date parser: `_parse_date_obj` (handles various formats).
    *   Procedure Date: `_get_procedure_date` (looks for `TV_IMP_IMPDAT` in main df).

### 3. SDV Tracking
*   Loads separate `CrfStatusHistory` file.
*   Matches forms (Subject, Folder, Form) to find "Verified" status.
*   Updates tree icons and provides dashboard metrics.

## Known Issues (Bugs & Bottlenecks)

### 1. Performance / Freezing
*   **Issue**: The app can freeze ("Not Responding") when loading large files or switching complex tabs.
*   **Cause**: Main thread doing heavy pandas processing. Tkinter UI runs on same thread.
*   **Mitigation**: Some optimizations done (vectorized operations), but fundamental single-thread architecture remains.

### 2. Data Inconsistency / Crashes
*   **Issue**: "Blank Screen" on Dashboard.
*   **Cause**: `get_summary_stats` encountering unhandled errors (e.g., missing columns, wrong sheet loaded).
*   **Fix**: Wrapped in `try-except` (Step 2842).
*   **Risk**: If header names change in the Excel export, mappings in `ae_manager.py` (and others) will break silent or loud.

### 3. Ongoing AE Count
*   **Logic**: Recently fixed to be strict: `AESER != 'Recovered/Fatal'` AND `End Date` is Empty.
*   **Prev**: Was just `Outcome` check, causing overcounting.

### 4. Excel File dependency
*   The app is tightly coupled to the specific structure (header row 1 codes, row 2 labels) of the "ProjectToOneFile" export format from the EDC system. "SelectForms" exports may work but are risky.

## Mappings
*   **AE Columns** (`ae_manager.py`):
    *   Onset: `LOGS_AE_AESTDTC`
    *   Outcome: `LOGS_AE_AEOUT`
    *   End Date: `LOGS_AE_AEENDTC`
    *   Relatedness: `LOGS_AE_AEREL1` (Trillium), `LOGS_AE_AEREL4` (Procedure)
*   **Procedure Date**:
    *   Looked up in `df_main` via `TV_IMP_IMPDAT` (Implantation Date).

## Next Steps for Development
1.  **AE Browser Overhaul**:
    *   Split into two tables (AEs vs SAEs).
    *   Add "Exclude Screen Failure" filter.
    *   Add "Highlight Late AEs" mode.
2.  **Refactoring**: 
    *   Further separate logic from UI (MVC pattern).
    *   Add async/threading for data loading to prevent freezing.

## How to Run
`python clinical_viewer1.py`
Requires: `pandas`, `openpyxl`, `tk` (standard lib).
