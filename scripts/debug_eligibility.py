import pandas as pd
import os
import glob

def find_file():
    files = glob.glob('verified/Innoventric_CLD-048_DM_CrfStatusHistory_*.xlsx')
    if not files: return None
    return sorted(files)[-1]

filepath = find_file()
if not filepath: exit()

print(f"Reading: {filepath}")
df_scan = pd.read_excel(filepath, sheet_name='Export', header=None, nrows=50, engine='openpyxl')
h_idx = 0
for i, row in df_scan.iterrows():
    if 'Form' in [str(v).strip() for v in row.values]:
        h_idx = i
        break

df = pd.read_excel(filepath, sheet_name='Export', header=h_idx, engine='openpyxl')
all_forms = sorted([str(f).strip() for f in df['Form'].unique()])

print("\n--- ALL UNIQUE FORMS (Total: {}) ---".format(len(all_forms)))
for f in all_forms:
    if len(f) > 0:
        print(f"FORM: {f}")

print("\n--- Patient 206-07 All Entries ---")
sub_col = 'Subject Screening #' if 'Subject Screening #' in df.columns else 'Scr #'
if sub_col not in df.columns and 'Subject' in df.columns: sub_col = 'Subject'

df_pat = df[df[sub_col].astype(str).str.contains('206-07')]
if not df_pat.empty:
    for i, row in df_pat.iterrows():
        print(f"V: {row['Activity']} | F: {row['Form']} | S: {row['Verification Status']} | U: {row.get('User', 'N/A')}")
else:
    print("Patient 206-07 not found")
