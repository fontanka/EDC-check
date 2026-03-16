import pandas as pd
import os
import glob

def find_file():
    files = glob.glob('verified/Innoventric_CLD-048_DM_CrfStatusHistory_*.xlsx')
    if not files:
        return None
    return sorted(files)[-1]

filepath = find_file()
if not filepath:
    print("No status history file found.")
    exit()

print(f"Reading: {filepath}")

# Load Export sheet - scan for header
df_scan = pd.read_excel(filepath, sheet_name='Export', header=None, nrows=50, engine='openpyxl')
header_row_idx = 0
for i, row in df_scan.iterrows():
    row_vals = [str(v).strip() for v in row.values]
    if 'Form' in row_vals and ('Scr #' in row_vals or 'Subject' in row_vals):
        header_row_idx = i
        break

df = pd.read_excel(filepath, sheet_name='Export', header=header_row_idx, engine='openpyxl')
df['Form'] = df['Form'].astype(str).str.strip()

sub_col = 'Subject Screening #' if 'Subject Screening #' in df.columns else 'Scr #'
if sub_col not in df.columns and 'Subject' in df.columns: sub_col = 'Subject'

df_pat = df[df[sub_col].astype(str).str.contains('206-07')]

print("\n--- Comparative View: Blood enzymes vs Liver function ---")
if not df_pat.empty:
    target_forms = ['Blood enzymes', 'Liver function panel']
    # Use str.contains with regex for target forms
    mask = df_pat['Form'].str.contains('Blood enzymes|Liver function', case=False)
    sub = df_pat[mask].copy()
    
    # Create DateTime correctly
    sub['DateTime'] = pd.to_datetime(sub['Date'].astype(str) + ' ' + sub['Time'].astype(str), errors='coerce')
    sub = sub.sort_values('DateTime')
    
    # Ensure Repeat column is visible
    repeat_col = 'Repeatable form #' if 'Repeatable form #' in sub.columns else 'Repeat'
    
    print(sub[['Activity', 'Form', repeat_col, 'Verification Status', 'User', 'DateTime']].to_string())
else:
    print("Patient 206-07 not found")
