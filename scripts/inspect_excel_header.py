import pandas as pd
import os
import glob

# Find latest file
cwd = os.getcwd()
files = [f for f in os.listdir(cwd) if f.startswith("Innoventric_CLD-048_DM_ProjectToOneFile") and f.endswith(".xlsx")]
if not files:
    print("No file found")
    exit()

latest_file = max(files, key=os.path.getctime)
print(f"Loading: {latest_file}")

xls = pd.read_excel(latest_file, sheet_name=None, header=None)
target = next((n for n in xls.keys() if "main" in n.lower()), None)

if target:
    df = xls[target]
    print("\n--- FIRST 5 ROWS ---")
    print(df.iloc[:5].to_string())
    
    print("\n--- ROW 0 (Codes?) ---")
    print(df.iloc[0].tolist()[:10])
    
    print("\n--- ROW 1 (Labels?) ---")
    print(df.iloc[1].tolist()[:10])
    
    print("\n--- ROW 2 (Data?) ---")
    print(df.iloc[2].tolist()[:10])
else:
    print("Main sheet not found")
