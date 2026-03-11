import pandas as pd
import os
import sys

# Find files
files = os.listdir('.')
# Try to find files in current or verified dir
stat_files = [f for f in files if "CrfStatusHistory" in f and f.endswith(".xlsx")]
if not stat_files:
    v_files = os.listdir('verified')
    stat_files = [os.path.join('verified', f) for f in v_files if "CrfStatusHistory" in f and f.endswith(".xlsx")]

if not stat_files:
    print("No status file found")
    sys.exit(1)

stat_file = stat_files[0]
print(f"Reading: {stat_file}")

# Load
try:
    # 1. Scan for header
    df_scan = pd.read_excel(stat_file, sheet_name='Export', header=None, nrows=50, engine='openpyxl')
    header_idx = 0
    for i, row in df_scan.iterrows():
        row_vals = [str(v).strip() for v in row.values]
        if 'Scr #' in row_vals or 'Subject' in row_vals or 'Subject Screening #' in row_vals:
            header_idx = i
            break
            
    df = pd.read_excel(stat_file, sheet_name='Export', header=header_idx, engine='openpyxl')
    
    # Normalize cols
    print(f"Columns found: {df.columns.tolist()}")
    
    if 'Form' in df.columns:
        df['Form'] = df['Form'].astype(str).str.strip()
    else:
        print("ERROR: 'Form' column not found!")
        sys.exit(1)
        
    # Search for ECG forms
    print("\n--- Listing ALL Unique Forms (First 50) ---")
    unique_forms = sorted(df['Form'].unique().tolist())
    for f in unique_forms[:50]:
        print(f"FORM: '{f}'")
        
    print(f"\nTotal Unique Forms: {len(unique_forms)}")
    
    ecg_forms = [f for f in unique_forms if 'ecg' in str(f).lower() or '12-lead' in str(f).lower()]
    
    for f in ecg_forms:
        print(f"DEBUG_FORM: '{f}'")
        
    print("\n--- checking specific patient 205-07 ---")
    sub_col = 'Subject Screening #' if 'Subject Screening #' in df.columns else 'Scr #'
    if sub_col not in df.columns and 'Subject' in df.columns: sub_col = 'Subject'
    
    if sub_col in df.columns:
        df_pat = df[df[sub_col].astype(str).str.contains('205-07')]
        if df_pat.empty:
            print("Patient 205-07 not found in this file.")
        else:
            pat_forms = df_pat['Form'].unique()
            print(f"Patient 205-07 Forms (ECG related):")
            found_ecg = False
            for f in pat_forms:
                if 'ecg' in str(f).lower() or '12-lead' in str(f).lower():
                     found_ecg = True
                     # Get status
                     row = df_pat[df_pat['Form'] == f].iloc[0]
                     status = row.get('Verification Status', 'N/A')
                     entry = row.get('Data Entry Status', 'N/A')
                     print(f"  - '{f}' : Entry={entry}, Ver={status}")
            if not found_ecg:
                print("No ECG forms found for this patient.")

except Exception as e:
    import traceback
    traceback.print_exc()
