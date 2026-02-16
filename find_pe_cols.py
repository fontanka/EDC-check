import pandas as pd
import glob
import os

# Find the latest project file
files = glob.glob("Innoventric_CLD-048_DM_ProjectToOneFile*.xlsx")
if not files:
    print("No project files found!")
    exit()

latest_file = max(files, key=os.path.getmtime)
print(f"Reading {latest_file}...")

try:
    # Use openpyxl or calamine if installed, default to openpyxl usually
    xl = pd.ExcelFile(latest_file)
    print(f"Sheet names: {xl.sheet_names}")
    
    # Try to find the main data sheet. Usually "Export Data" or "Main" or similar.
    # checking clinical_viewer1.py might reveal the sheet name but let's guess "Main" based on file list having "ACT" sheets? 
    # Actually clinical_viewer1.py load_excel uses "Active sheet" or specific name?
    # Let's inspect all sheet headers
    
    sheet = 'Main'
    if sheet in xl.sheet_names:
        print(f"-- Parsing {sheet} --")
        df = xl.parse(sheet, nrows=1)
        cols = df.columns.tolist()
        
        # Search for suspects
        suspects = []
        for c in cols:
            c_upper = str(c).upper()
            # Broad search for Physical Exam terms
            if "PE_" in c_upper or "PHYS" in c_upper or "GEN_" in c_upper or "APPEAR" in c_upper or "SKIN" in c_upper or "HEAD" in c_upper or "NECK" in c_upper or "ABNORM" in c_upper:
                suspects.append(c)
        
        # Dump data for specific columns to understand values
        target_cols = [c for c in cols if "PE_" in str(c).upper() and ("HEAD" in str(c).upper() or "CARD" in str(c).upper())]
        target_cols = target_cols[:4] + [c for c in cols if c == "SBV_PE_PESTAT" or c == "SBV_PE_PEREASND"]
        
        if target_cols:
             print(f"Columns: {target_cols}")
             # Parse with header=None to see row indices 0 and 1 clearly
             df_raw = xl.parse(sheet, header=None, nrows=5)
             
             # Find indices of target columns
             indices = []
             for col_name in target_cols:
                 # Find where this column name is in row 0
                 matches = df_raw.iloc[0] == col_name
                 if matches.any():
                     indices.append(matches.idxmax())
             
             if indices:
                 subset = df_raw.iloc[:, indices]
                 print("\n-- Metadata (Codes & Labels) --")
                 print(subset.head(3).to_string())
             
             # Data Sample for 206-07
             df_data = xl.parse(sheet)
             # Find Screening # column
             pat_col = next((c for c in df_data.columns if "Screening" in str(c)), None)
             if pat_col:
                 subset_data = df_data[df_data[pat_col].astype(str) == '206-07']
                 if not subset_data.empty:
                     print(f"\n-- Data Sample for 206-07 -- (PESTAT={subset_data['SBV_PE_PESTAT'].iloc[0]})")
                     print(subset_data[target_cols].to_string())
                 else:
                     print("Patient 206-07 not found.")
             else:
                 print("Screening # column not found.")
        else:
             print("Target PE columns not found.")
            


except Exception as e:
    print(f"Error: {e}")
