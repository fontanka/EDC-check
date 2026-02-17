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
    # Use exact params from clinical_viewer1.py
    xls = pd.read_excel(latest_file, sheet_name=None, header=None, dtype=str, keep_default_na=False)
    target = next((n for n in xls.keys() if "main" in n.lower()), None)

    if target:
        df = xls[target]
        print(f"\nShape: {df.shape}")
        
        # Search for "SBV_" in first 5 rows
        for i in range(min(5, len(df))):
            row = df.iloc[i].tolist()
            sbv_cols = [x for x in row if "SBV_" in str(x)]
            labels_cols = [x for x in row if "Date of Birth" in str(x) or "Weight" in str(x)]
            
            print(f"\n--- ROW {i} ---")
            print(f"Contains 'SBV_'? {len(sbv_cols)} items. Sample: {sbv_cols[:3]}")
            print(f"Contains Labels? {len(labels_cols)} items. Sample: {labels_cols[:3]}")
    else:
        print("Main sheet not found")

except Exception as e:
    print(f"Error: {e} \n(Note: openpyxl dependency mismatch?)")
