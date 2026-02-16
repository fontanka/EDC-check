import pandas as pd
import os
import sys

# Force stdout encoding
sys.stdout.reconfigure(encoding='utf-8')

cwd = os.getcwd()
files = [f for f in os.listdir(cwd) if f.startswith("Innoventric_CLD-048_DM_ProjectToOneFile") and f.endswith(".xlsx")]
if not files:
    print("No file found")
    exit()

latest_file = max(files, key=os.path.getctime)
print(f"Loading: {latest_file}")

try:
    xls = pd.read_excel(latest_file, sheet_name=None, header=None, dtype=str, keep_default_na=False)
    target = next((n for n in xls.keys() if "main" in n.lower()), None)

    if target:
        df = xls[target]
        header_row_idx = 0
        codes = [str(c).strip() for c in df.iloc[header_row_idx].tolist()]
        data = df.iloc[header_row_idx+2:].copy()
        data.columns = codes
        
        # Filter for Patient 206-06
        subset = data[data.astype(str).apply(lambda x: x.str.contains('206-06').any(), axis=1)]
        
        if subset.empty:
            print("Patient 206-06 not found.")
        else:
            print(f"\n--- ECG Columns for 206-06 ---")
            # Find ECG columns
            ecg_cols = [c for c in data.columns if 'SBV_ECG_EGORRES_ABN' in c]
            
            # Find CRA_CONTROL_STATUS column
            # In excel it might be named clearly? No, typically 'CRA_CONTROL_STATUS' or similar.
            # But wait, the raw excel has 'CRA_CONTROL_STATUS' as a generic column? 
            # OR is it per field?
            # Viewer implies it is per-row?
            # "Status" column in ACT sheet mentioned in debug? Main sheet logic is 1 row per record?
            # NO! Main sheet is 1 row per VISIT/FORM? 
            # SdvManager `modular_data` is melted.
            # inspecting the EXCEL file directly. The Excel usually has 1 row per 'Event' (Form instance).
            # Columns are Variables.
            # WHERE is the status stored?
            # In `read_data_matrix` or `load_data` logic?
            # `DashboardManager` uses `sdv_mgr.modular_data`.
            # `sdv_mgr.modular_data` comes from `load_sdv_data`.
            # `load_sdv_data` reads `SDV_Status.xlsx` (or similar)?
            # No, `clinical_viewer1.py` `load_data` loads MAIN sheet.
            # `DashboardManager` works on `sdv_mgr.modular_data`.
            
            pass 
            # I cannot inspect modular_data structure from raw excel easily if I don't know how it's built.
            # `modular_data` seems to be loaded from `SDV_Status.xlsx`?
            # User calls it `Innoventric_CLD-048_DM_ProjectToOneFile...`
            # Wait, `DashboardManager` uses `self.sdv_mgr.modular_data`.
            # `self.sdv_mgr` is `SDVManager`.
            # `SDVManager.load_data` reads... what?
            
            # Let's write a script that imports ClinicalViewer/SDVManager and inspects the `modular_data` DF directly!
            # Much more reliable.

except Exception as e:
    print(e)
