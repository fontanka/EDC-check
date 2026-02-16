from sdv_manager import SDVManager
import os
import glob

def find_file():
    files = glob.glob('verified/Innoventric_CLD-048_DM_CrfStatusHistory_*.xlsx')
    if not files: return None
    return sorted(files)[-1]

filepath = find_file()
if not filepath:
    print("No status history file found.")
    exit()

manager = SDVManager()
print(f"Loading: {filepath}")
manager.load_crf_status_file(filepath)

patient = '206-07'
print(f"\n--- Checking keys in verification_metadata for {patient} ---")
found = False
for key in manager.verification_metadata.keys():
    if patient in key:
        if 'Liver' in key or 'Blood' in key or 'Enzyme' in key:
             print(f"KEY: '{key}' -> {manager.verification_metadata[key]}")
             found = True

if not found:
    print(f"No Liver/Blood keys found for {patient} in results.")
    # Show ALL keys for this patient
    print("\n--- ALL keys for this patient ---")
    for key in manager.verification_metadata.keys():
        if patient in key:
            print(f"FULL_KEY: '{key}'")

# Test get_verification_details directly (Robustness tests)
print("\n--- Testing get_verification_details robustness ---")
l_details_exact = manager.get_verification_details(patient, "Liver function panel", "Screening/Baseline")
print(f"Liver (Exact): {l_details_exact}")

l_details_space = manager.get_verification_details(patient, "Liver function panel ", " Screening/Baseline ")
print(f"Liver (Spaces): {l_details_space}")

l_details_short = manager.get_verification_details(patient, "Liver function", "Screening")
print(f"Liver (Short Name): {l_details_short}")

l_details_long = manager.get_verification_details(patient, "Liver function panel (with extra data)", "Screening/Baseline")
print(f"Liver (Longer Name): {l_details_long}")

b_details = manager.get_verification_details(patient, "Blood enzymes", "Screening/Baseline")
print(f"Blood Details: {b_details}")
