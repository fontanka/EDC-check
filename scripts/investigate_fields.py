import pandas as pd

# Load SDV data for 206-06
df = pd.read_csv('verified/Innoventric_CLD-048_DM_GeneralEcrfHistory_14-01-2026_10-16_40_(UTC).csv', encoding='utf-8-sig')
patient_data = df[df['Scr_Number'].astype(str) == "206-06"]

# Load main data columns
df_main = pd.read_excel('Innoventric_CLD-048_DM_ProjectToOneFile_14-01-2026_09-05_42_(UTC).xlsx', sheet_name='Main', nrows=0)
main_cols = set(df_main.columns)

print("=== CVH (Cardiovascular History) Field IDs in SDV ===")
cvh = patient_data[patient_data['Form'].str.contains('Cardiovascular', na=False, case=False)]
cvh_fields = cvh[cvh['Field_Id'].notna()][['Field_Id', 'Action']].drop_duplicates()
for _, row in cvh_fields.head(30).iterrows():
    fid = row['Field_Id']
    # Check if matches in main columns
    matches = [c for c in main_cols if 'CVH' in c and fid.replace(' ', '') in c.replace(' ', '').replace('#', '')]
    print(f"  SDV: {fid:25} -> Main matches: {matches[:3]}")

print("\n=== Clinical Frailty Scale Field IDs in SDV ===")
cfs = patient_data[patient_data['Form'].str.contains('Clinical Frailty', na=False, case=False)]
cfs_fields = cfs[cfs['Field_Id'].notna()][['Activity', 'Field_Id', 'Action']].drop_duplicates()
for _, row in cfs_fields.iterrows():
    activity = row['Activity']
    fid = row['Field_Id']
    action = row['Action'][:50] if len(str(row['Action'])) > 50 else row['Action']
    # Check for matches
    matches = [c for c in main_cols if 'CFS' in c and fid in c]
    print(f"  {activity}: {fid:25} (action: {action})")
    if matches:
        print(f"    -> Matches: {matches[:3]}")

print("\n=== CFS columns in Main data ===")
cfs_cols = [c for c in main_cols if 'CFS' in c]
for c in sorted(cfs_cols)[:20]:
    print(f"  {c}")

print("\n=== Echocardiography Field IDs in SDV (verified only) ===")
echo = patient_data[patient_data['Form'].str.contains('Echocardiography', na=False, case=False)]
echo_verified = echo[echo['Action'].str.contains('verified', case=False, na=False)]
echo_fields = echo_verified[echo_verified['Field_Id'].notna()][['Activity', 'Form', 'Field_Id']].drop_duplicates()
for _, row in echo_fields.head(30).iterrows():
    activity = row['Activity']
    form = row['Form'][:30]
    fid = row['Field_Id']
    # Check for core lab
    is_core = 'Core' in str(row['Form'])
    suffix = '_SPONSOR' if is_core else ''
    matches = [c for c in main_cols if 'ECHO' in c and fid in c]
    print(f"  {activity}: {fid:25} (core={is_core}) -> {matches[:2] if matches else 'NO MATCH'}")

print("\n=== Echo columns sample in Main data ===")
echo_cols = [c for c in main_cols if 'ECHO_1D' in c]
for c in sorted(echo_cols)[:20]:
    print(f"  {c}")
