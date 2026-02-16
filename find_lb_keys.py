import sys
sys.path.insert(0, 'c:/budgets')
from sdv_manager import SDVManager
import glob
import os

s = SDVManager()
mod_files = [f for f in glob.glob(r"c:\budgets\verified\*Modular*.xlsx") if not os.path.basename(f).startswith("~$")]
mod_file = max(mod_files, key=os.path.getmtime)
s.load_modular_file(mod_file)

patient = "208-07"

print(f"=== All Keys for {patient} containing 'LB' or 'ORRES' ===\n")

if patient in s.patient_index:
    p_data = s.patient_index[patient]
    
    matching_keys = [k for k in p_data.keys() if 'LB' in k or 'ORRES' in k]
    
    for k in sorted(matching_keys)[:20]:  # First 20
        status_code, hidden, has_val = p_data[k]
        mapped = s._map_status(status_code, hidden, has_val)
        print(f"{k:60} → Status={status_code}, Mapped='{mapped}'")
else:
    print(f"Patient {patient} not found")
