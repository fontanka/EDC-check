import pandas as pd
import os
import glob
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
    xls = pd.read_excel(latest_file, sheet_name=None, header=None)
    target = next((n for n in xls.keys() if "main" in n.lower()), None)

    if target:
        df = xls[target]
        print(f"\nShape: {df.shape}")
        
        for i in range(5):
            row_vals = df.iloc[i].tolist()
            # Filter empty strings/NaN to see content
            content = [x for x in row_vals if str(x).strip() not in ['nan', '', 'None', 'NaN']]
            print(f"\n--- ROW {i} CONTENT SAMPLE (First 5 of {len(content)}) ---")
            print(content[:5])
            print("Full Row first 5 raw:", row_vals[:5])
    else:
        print("Main sheet not found")

except Exception as e:
    print(f"Error: {e}")
