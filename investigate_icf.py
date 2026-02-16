import pandas as pd
import os
import glob

verified_dir = r"c:\budgets\verified"
if os.path.isdir(verified_dir):
    modular_files = glob.glob(os.path.join(verified_dir, "*Modular*.xlsx"))
    if modular_files:
        filepath = max(modular_files, key=os.path.getmtime)
        print(f"File: {filepath}")
        
        # Load all data from 'Export Data'
        df = pd.read_excel(filepath, sheet_name='Export Data', engine='calamine')
        
        # Identify ICF rows. Try 'Form Code' if exists, otherwise search 'Form name'
        icf_mask = (df['Form name'].astype(str).str.contains('Informed Consent', case=False, na=False)) | \
                   (df['Form Code'].astype(str).str.contains('ICF', case=False, na=False) if 'Form Code' in df.columns else False)
        
        icf_df = df[icf_mask]
        
        if icf_df.empty:
            print("No ICF data found. Available forms:")
            print(df['Form name'].unique())
            exit()
            
        print(f"\nFound {len(icf_df)} ICF records.")
        
        # 1. List all variables for ICF
        print("\n--- ICF Variables ---")
        vars_df = icf_df[['Variable name', 'Variable label']].drop_duplicates()
        print(vars_df.to_string(index=False))
        
        # 2. Check a few rows for patient 208-07 to see actual values and statuses
        pat = '208-07'
        pat_icf = icf_df[icf_df['Subject Screening #'].astype(str).str.contains(pat, na=False)]
        
        if pat_icf.empty:
            pat = icf_df['Subject Screening #'].unique()[0]
            pat_icf = icf_df[icf_df['Subject Screening #'] == pat]
            
        print(f"\n--- Data for Patient {pat} ---")
        cols = ['Variable name', 'Variable label', 'Variable Value', 'CRA_CONTROL_STATUS', 'CRA_CONTROL_STATUS_NAME', 'Hidden']
        # Add any other interesting columns
        extra = [c for c in df.columns if any(x in c.upper() for x in ['STATUS', 'VERIFIED'])]
        for c in extra:
            if c not in cols: cols.append(c)
            
        print(pat_icf[cols].to_string(index=False))
        
        # 3. Look for differentiation indicators
        print("\n--- Status Distribution (Whole File) ---")
        print(df['CRA_CONTROL_STATUS_NAME'].value_counts())
