import sys
sys.path.insert(0, 'c:/budgets')
import pandas as pd
import glob
import os

# Load the modular file
mod_files = [f for f in glob.glob(r"c:\budgets\verified\*Modular*.xlsx") if not os.path.basename(f).startswith("~$")]
mod_file = max(mod_files, key=os.path.getmtime)

df = pd.read_excel(mod_file, sheet_name='Export Data', dtype=str, keep_default_na=False)

patient = "208-07"
df_pat = df[df['Subject Screening #'] == patient]

# Find all unique form codes
form_codes = df_pat['Form Code'].unique()
print(f"=== All Form Codes for {patient} ===")
for fc in sorted([f for f in form_codes if f]):
    if 'AE' in fc or 'CM' in fc:
        count = len(df_pat[df_pat['Form Code'] == fc])
        print(f"{fc}: {count} rows")

# Try to find edema
print(f"\n=== Searching for 'edema' ===")
edema_rows = df_pat[df_pat['Variable Value'].str.contains('edema', case=False, na=False)]
if not edema_rows.empty:
    print(f"Found {len(edema_rows)} rows containing 'edema'")
    for idx, row in edema_rows.head(3).iterrows():
        print(f"  Form: {row['Form Code']}, Table Row: {row['Table row #']}, Var: {row['Variable name']}, Val: {row['Variable Value']}")
