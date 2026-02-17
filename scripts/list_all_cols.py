import pandas as pd
import os
import glob

verified_dir = r"c:\budgets\verified"
if os.path.isdir(verified_dir):
    modular_files = glob.glob(os.path.join(verified_dir, "*Modular*.xlsx"))
    if modular_files:
        filepath = max(modular_files, key=os.path.getmtime)
        print(f"Analyzing: {filepath}")
        
        xl = pd.ExcelFile(filepath, engine='calamine')
        if 'Export Data' in xl.sheet_names:
            df = xl.parse('Export Data', nrows=1)
            print("\n--- Columns ---")
            for i, col in enumerate(df.columns):
                print(f"{i}: {col}")
