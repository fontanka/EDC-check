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
        print(f"\nSearching in sheet: {target}")
        
        found_code = False
        found_label = False
        
        target_code = "SBV_SVSTDTC"
        target_label_fragment = "Date" # "Visit Date" might be exact
        
        for i in range(min(10, len(df))):
            row = df.iloc[i].tolist()
            row_str = [str(x).strip() for x in row]
            
            if target_code in row_str:
                col_idx = row_str.index(target_code)
                print(f"FOUND CODE '{target_code}' at Row {i}, Col {col_idx}")
                found_code = True
                
                # Check corresponding label in other rows at same column
                print(f"Checking same column {col_idx} in other rows:")
                for j in range(min(5, len(df))):
                    print(f"  Row {j}: {df.iloc[j, col_idx]}")
                    
            if not found_code:
                # Fuzzy check
                matches = [x for x in row_str if target_code in x]
                if matches:
                     print(f"Partial match for code in Row {i}: {matches}")

    else:
        print("Main sheet not found")

except Exception as e:
    print(f"Error: {e}")
