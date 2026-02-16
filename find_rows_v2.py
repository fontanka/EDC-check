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
    xls = pd.read_excel(latest_file, sheet_name=None, header=None, dtype=str, keep_default_na=False)
    target = next((n for n in xls.keys() if "main" in n.lower()), None)

    if target:
        df = xls[target]
        print(f"\nSearching ENTIRE sheet: {target} (First 20 rows)")
        
        found = False
        for i in range(min(20, len(df))):
            row = df.iloc[i].tolist()
            # Look for ANY SBV_ variable
            sbv_matches = [x for x in row if "SBV_" in str(x)]
            if sbv_matches:
                print(f"\n--- MATCH IN ROW {i} ---")
                print(f"SBV Matches: {sbv_matches[:3]}...")
                print(f"First 5 cols: {row[:5]}")
                found = True
        
        if not found:
            print("No 'SBV_' found in first 20 rows.")

    else:
        print("Main sheet not found")

except Exception as e:
    print(f"Error: {e}")
