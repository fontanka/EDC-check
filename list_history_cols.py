import pandas as pd
import os
import shutil
import glob

verified_dir = r"c:\budgets\verified"
files = [f for f in glob.glob(os.path.join(verified_dir, "*CrfStatusHistory*.xlsx")) if not os.path.basename(f).startswith("~$")]
if files:
    src = max(files, key=os.path.getmtime)
    dst = "temp_history_3.xlsx"
    shutil.copy2(src, dst)

    try:
        xl = pd.ExcelFile(dst)
        if 'Export' in xl.sheet_names:
            df = xl.parse('Export', nrows=1)
            print("\n--- Columns ---")
            for i, c in enumerate(df.columns):
                print(f"{i}: {c}")
    finally:
        if os.path.exists(dst):
            try: os.remove(dst)
            except: pass
