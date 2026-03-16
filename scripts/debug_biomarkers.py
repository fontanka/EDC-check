import sys
import os
import glob
import pandas as pd
from sdv_manager import SDVManager

# Setup
manager = SDVManager()

# Find files
files = os.listdir('.')
mod_files = [f for f in files if "Modular" in f and f.endswith(".xlsx")]
stat_files = [f for f in files if "CrfStatusHistory" in f and f.endswith(".xlsx")]

if not mod_files:
    v_files = os.listdir('verified')
    mod_files = [os.path.join('verified', f) for f in v_files if "Modular" in f and f.endswith(".xlsx")]
if not stat_files:
    v_files = os.listdir('verified')
    stat_files = [os.path.join('verified', f) for f in v_files if "CrfStatusHistory" in f and f.endswith(".xlsx")]

if not mod_files or not stat_files:
    print("Missing files!")
    sys.exit(1)

mod_file = mod_files[0]
stat_file = stat_files[0]

print(f"Loading Modular: {mod_file}")
manager.load_modular_file(mod_file)

print(f"Loading Status: {stat_file}")
manager.load_crf_status_file(stat_file)

# Test Target
patient = "205-07"
form = "Biomarkers" 
visit = "Treatment" 
field = "TV_LBDAT_BM"

print(f"\n--- Testing {patient} | {visit} | {form} ---")

# 1. Field Status
f_stat = manager.get_field_status(patient, field)
print(f"Field Status ({field}): {f_stat}")

# 2. Verification Details
details = manager.get_verification_details(patient, form, visit_name=visit)
print(f"Verification Details: {details}")

# 3. Inspect Form Entry Status Dump
print("\n--- Form Entry Status Matches ---")
found = False
prefix = f"{patient}|"
print(f"Prefix: {prefix}")
for key, val in manager.form_entry_status.items():
    if key.startswith(prefix):
        if "biomarker" in key.lower():
            print(f"MATCH Key: {key}")
            print(f"      Val: {val}")
            
            if key in manager.verification_metadata:
                 print(f"      Meta: {manager.verification_metadata[key]}")
            else:
                 print(f"      Meta: None")
