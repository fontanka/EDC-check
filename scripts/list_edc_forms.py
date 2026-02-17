import pandas as pd
import os
import glob
import shutil

verified_dir = r"c:\budgets\verified"

# Get unique form names from CrfStatusHistory
files = [f for f in glob.glob(os.path.join(verified_dir, "*CrfStatusHistory*.xlsx")) if not os.path.basename(f).startswith("~$")]
if files:
    src = max(files, key=os.path.getmtime)
    dst = "temp_crf.xlsx"
    shutil.copy2(src, dst)
    try:
        df = pd.read_excel(dst, sheet_name='Export')
        
        # Get unique form names
        forms = sorted(df['Form'].dropna().unique())
        print("=== All EDC Form Names ===")
        for form in forms:
            print(f"  '{form}'")
        print(f"\nTotal: {len(forms)} forms")
            
    finally:
        if os.path.exists(dst):
            try: os.remove(dst)
            except: pass
