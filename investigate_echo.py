import pandas as pd
from sdv_manager import SDVManager
import logging

# Configure basic logging
logging.basicConfig(level=logging.ERROR)

# Load SDV data
m = SDVManager()
m.load_sdv_file('verified/Innoventric_CLD-048_DM_GeneralEcrfHistory_14-01-2026_10-16_40_(UTC).csv')

# Build lookup for 208-07
m.build_verified_lookup('208-07')

print("=== ECHO Forms Status for 208-07 ===")
echo_forms = {k: v for k, v in m.form_status.get('208-07', {}).items() if 'ECHO' in k}
for f, s in echo_forms.items():
    print(f"  {f}: {s}")

print("\n=== ECHO Fields Status for 208-07 (Sample) ===")
echo_fields = {k: v for k, v in m.field_status.get('208-07', {}).items() if 'ECHO' in k}
for f in list(echo_fields.keys())[:10]:
    print(f"  {f}: {echo_fields[f]}")

print("\n=== Raw SDV Entries for 208-07 Screening Echo ===")
# Raw check for 208-07 Echo to see form names and actions
df = m.sdv_data
mask = (df['Scr_Number'].astype(str) == '208-07') & (df['Form'].str.contains('Echocardiography', na=False))
echo_data = df[mask]
print(echo_data[['Activity', 'Form', 'Field_Id', 'Action']].head(20).to_string())
