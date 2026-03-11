import pandas as pd

# Load SDV data for 206-06
df = pd.read_csv('verified/Innoventric_CLD-048_DM_GeneralEcrfHistory_14-01-2026_10-16_40_(UTC).csv', encoding='utf-8-sig')
patient_data = df[df['Scr_Number'].astype(str) == "206-06"]

print("=== Physical Examination (PE) ALL Entries ===")
pe = patient_data[patient_data['Form'].str.contains('Physical Examination', na=False, case=False)]
pe = pe[pe['Activity'] == 'Screening/Baseline']
print(f"Total PE entries: {len(pe)}")

# Show form-level entries (Field_Id is NaN)
print("\n--- Form-level entries (Field_Id is NaN): ---")
pe_form = pe[pe['Field_Id'].isna()]
for _, row in pe_form.iterrows():
    print(f"  Action: {row['Action']}")
    print(f"    Created: {row['Created_On']}")
    print()

# Show field-level entries with verification patterns
print("\n--- Field-level entries with 'verified' in Action: ---")
pe_field = pe[pe['Field_Id'].notna()]
pe_verified = pe_field[pe_field['Action'].str.contains('verified', case=False, na=False)]
for _, row in pe_verified.head(10).iterrows():
    print(f"  Field: {row['Field_Id']:30} Action: {row['Action'][:60]}")

print("\n\n=== Basic Metabolic Panel (BMP) ALL Entries ===")
bmp = patient_data[patient_data['Form'].str.contains('Basic metabolic', na=False, case=False)]
bmp = bmp[bmp['Activity'] == 'Screening/Baseline']
print(f"Total BMP entries: {len(bmp)}")

# Show form-level entries
print("\n--- Form-level entries (Field_Id is NaN): ---")
bmp_form = bmp[bmp['Field_Id'].isna()]
for _, row in bmp_form.iterrows():
    print(f"  Action: {row['Action']}")
    print(f"    Created: {row['Created_On']}")
    print()

# Show field-level verified entries
print("\n--- Field-level entries with 'verified' in Action: ---")
bmp_field = bmp[bmp['Field_Id'].notna()]
bmp_verified = bmp_field[bmp_field['Action'].str.contains('verified', case=False, na=False)]
for _, row in bmp_verified.head(10).iterrows():
    print(f"  Field: {row['Field_Id']:30} Action: {row['Action'][:60]}")

print("\n\n=== Biomarkers (BM) ALL Entries ===")
bm = patient_data[patient_data['Form'].str.contains('Biomarkers', na=False, case=False)]
bm = bm[bm['Activity'] == 'Screening/Baseline']
print(f"Total BM entries: {len(bm)}")

# Form-level
print("\n--- Form-level entries (Field_Id is NaN): ---")
bm_form = bm[bm['Field_Id'].isna()]
for _, row in bm_form.iterrows():
    print(f"  Action: {row['Action']}")
    print(f"    Created: {row['Created_On']}")
