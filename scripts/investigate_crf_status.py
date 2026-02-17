import pandas as pd
import os
import glob
import shutil

verified_dir = r"c:\budgets\verified"
files = [f for f in glob.glob(os.path.join(verified_dir, "*CrfStatusHistory*.xlsx")) if not os.path.basename(f).startswith("~$")]
if files:
    src = max(files, key=os.path.getmtime)
    dst = "temp_crf_status.xlsx"
    shutil.copy2(src, dst)

    try:
        xl = pd.ExcelFile(dst)
        df = xl.parse('Export')
        
        print("Columns:", df.columns.tolist())
        print("\nData Entry Status values:")
        print(df['Data Entry Status'].value_counts())
        
        # Find most recent status for each patient + form + activity
        df['DateTime'] = pd.to_datetime(df['Date'].astype(str) + ' ' + df['Time'].astype(str), errors='coerce')
        
        # Group by Scr#, Form and get most recent entry
        most_recent = df.sort_values('DateTime').groupby(['Scr #', 'Form']).last().reset_index()
        
        print("\n--- Forms with 'Created' as most recent status ---")
        created_forms = most_recent[most_recent['Data Entry Status'] == 'Created']
        print(f"Total: {len(created_forms)}")
        print(created_forms[['Scr #', 'Form', 'Data Entry Status', 'Date', 'Time']].head(30).to_string())
        
        # Check 206-06 specifically
        print("\n--- 206-06 Form Statuses ---")
        pat_forms = most_recent[most_recent['Scr #'].astype(str).str.contains('206-06')]
        print(pat_forms[['Scr #', 'Form', 'Data Entry Status', 'Date', 'Time']].to_string())
        
    finally:
        if os.path.exists(dst):
            try: os.remove(dst)
            except: pass
