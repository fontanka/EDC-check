import sys
sys.path.insert(0, 'c:/budgets')
import pandas as pd
import glob
import os

mod_files = [f for f in glob.glob(r"c:\budgets\verified\*Modular*.xlsx") if not os.path.basename(f).startswith("~$")]
mod_file = max(mod_files, key=os.path.getmtime)

df = pd.read_excel(mod_file, sheet_name='Export Data', dtype=str, keep_default_na=False)

patient = "208-07"
df_pat = df[df['Subject Screening #'] == patient]

# Get all AE form rows
ae_rows = df_pat[df_pat['Form Code'] == 'AE'].copy()

print(f"=== All AE Rows for {patient} ===\n")
print(f"Total AE rows: {len(ae_rows)}")

# Get AETERM entries to see which ones have/don't have row numbers
aeterm_rows = ae_rows[ae_rows['Variable name'].str.contains('AETERM', na=False) & 
                      ~ae_rows['Variable name'].str.contains('COMM', na=False)]

print(f"AETERM entries: {len(aeterm_rows)}\n")

for idx, row in aeterm_rows.iterrows():
    table_row = row['Table row #']
    val = row['Variable Value'][:50]
    status = row['CRA_CONTROL_STATUS']
    
    row_display = f"'{table_row}'" if table_row else "(EMPTY)"
    print(f"Row {row_display:8}: {val:50} Status={status}")

# Check if there's a pattern - maybe row numbers are assigned sequentially by data order
print(f"\n=== Checking if order matters ===")
print("If Table row # is empty, perhaps the display row is based on data order?")
