import pandas as pd
import os
import sys

# Find files
files = os.listdir('.')
mod_files = [f for f in files if "Modular" in f and f.endswith(".xlsx")]
stat_files = [f for f in files if "CrfStatusHistory" in f and f.endswith(".xlsx")]

if not mod_files:
    v_files = os.listdir('verified')
    mod_files = [os.path.join('verified', f) for f in v_files if "Modular" in f and f.endswith(".xlsx")]
if not stat_files:
    v_files = os.listdir('verified')
    stat_files = [os.path.join('verified', f) for f in v_files if "CrfStatusHistory" in f and f.endswith(".xlsx")]

stat_file = stat_files[0]
print(f"Reading Status File: {stat_file}")

# Read status file - simpler load
try:
    # Check sheet names
    xl = pd.ExcelFile(stat_file)
    print("Sheet names:", xl.sheet_names)
    
    sheet_to_read = 'Export' if 'Export' in xl.sheet_names else xl.sheet_names[0]
    print(f"Reading sheet: {sheet_to_read}")

    # Re-read with no header to scan
    df = pd.read_excel(stat_file, sheet_name=sheet_to_read, header=None, nrows=100)
    print("Scanning first 100 rows for header...")
    
    header_idx = -1
    for i, row in df.iterrows():
        row_str = row.astype(str).tolist()
        if 'Scr #' in row_str:
            header_idx = i
            print(f"Found header at row {i}: {row_str}")
            break
    
    if header_idx != -1:
        df = pd.read_excel(stat_file, sheet_name=sheet_to_read, header=header_idx)
    else:
        print("Could not find 'Scr #' in first 100 rows")
        print("Rows 0-20 (since header not found):")
        print(df.iloc[0:21])
        sys.exit(1)

    print("Columns found:", df.columns.tolist())
    
    # Filter for patient 205-07
    # Determine Subject column
    sub_col = 'Subject Screening #' if 'Subject Screening #' in df.columns else 'Scr #'
    if sub_col not in df.columns:
         # Try finding any column with 'Scr' or 'Subject'
         candidates = [c for c in df.columns if "Scr" in str(c) or "Subject" in str(c)]
         if candidates:
             sub_col = candidates[0]
         else:
             print("Could not identify Subject column")
             sys.exit(1)

    df_pat = df[df[sub_col] == '205-07']

    print(f"Rows for 205-07: {len(df_pat)}")
    
    # Check for Biomarkers
    mask = df_pat.astype(str).apply(lambda x: x.str.contains('Biomarker', case=False)).any(axis=1)
    df_bio = df_pat[mask]
    
    print(f"Rows with 'Biomarker': {len(df_bio)}")
    
    ver_col = 'Verification Status' if 'Verification Status' in df.columns else df.columns[0] 
    user_col = 'User' if 'User' in df.columns else 'Status Changed By'
    visit_col = 'Activity' if 'Activity' in df.columns else ('Visit' if 'Visit' in df.columns else None)
    
    if not df_bio.empty:
        # Construct relevant columns dynamically
        cols_to_show = ['Form']
        if visit_col: cols_to_show.insert(0, visit_col)
        if ver_col in df.columns: cols_to_show.append(ver_col)
        if user_col in df.columns: cols_to_show.append(user_col)
        if 'Date' in df.columns: cols_to_show.append('Date')
        
        print("\nColumns available:", df.columns.tolist())
        print("\nSample Rows:")
        print(df_bio[cols_to_show].head().to_string())
        
        print("\nUnique Forms:")
        print(df_bio['Form'].unique())
        
        if visit_col:
            print(f"\nUnique {visit_col}:")
            print(df_bio[visit_col].unique())
            
            # Check Treatment Visit
            mask_tv = df_pat[visit_col].astype(str).str.contains('Treatment', case=False)
            df_tv = df_pat[mask_tv]
            print(f"\nRows with 'Treatment': {len(df_tv)}")
            if not df_tv.empty:
                print("Unique Forms in Treatment Visit:")
                print(df_tv['Form'].unique())

except Exception as e:
    import traceback
    traceback.print_exc()
