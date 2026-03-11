import pandas as pd
import os
import glob
from sdv_manager import SDVManager
from dashboard_manager import DashboardManager

def test_dashboard():
    print("Initializing SDV Manager...")
    sdv_mgr = SDVManager()
    
    # Load latest modular file
    verified_dir = "c:/budgets/verified"
    files = glob.glob(os.path.join(verified_dir, "*Modular*.xlsx"))
    if not files:
        print("No modular file found.")
        return
        
    latest_file = max(files, key=os.path.getmtime)
    print(f"Loading {latest_file}...")
    sdv_mgr.load_modular_file(latest_file)
    
    # Load CrfStatusHistory if available
    crf_files = glob.glob(os.path.join(verified_dir, "*CrfStatusHistory*.xlsx"))
    if crf_files:
        crf_file = max(crf_files, key=os.path.getmtime)
        print(f"Loading {crf_file}...")
        sdv_mgr.load_crf_status_file(crf_file)
    
    print("Initializing Dashboard Manager...")
    dash_mgr = DashboardManager(sdv_mgr)
    
    print("Calculating Stats...")
    dash_mgr.calculate_stats()
    
    print("\n=== Study Level Stats ===")
    study_stats = dash_mgr.get_summary('study')
    print(study_stats)
    
    print("\n=== Site Level Stats ===")
    site_stats = dash_mgr.get_summary('site')
    for site, stats in site_stats.items():
        print(f"Site {site}: {stats}")
        
    print("\n=== Top 5 Patients by Gaps ===")
    top_gaps = dash_mgr.get_top_counts('patient', 'GAP', 5)
    for pat_id, count in top_gaps:
        print(f"Patient {pat_id}: {count['GAP']} Gaps")
        # Show first 3 gaps
        gaps = dash_mgr.get_details('patient', pat_id, 'GAP')
        print(f"  First 3 gaps: {[g['Field'] for g in gaps[:3]]}")
        
    print("\n=== Top 5 Forms by Pending (!) ===")
    # form keys are tuples (pat, form)
    top_pending = dash_mgr.get_top_counts('form', '!', 5)
    for (pat, form), count in top_pending:
        print(f"Patient {pat} - {form}: {count['!']} Pending")

    print("\n=== Testing Exclusion (Screen Failure) ===")
    # Pick a patient to exclude
    pats = list(dash_mgr.stats['patient'].keys())
    if pats:
        exclude_pat = pats[0]
        print(f"Excluding Patient {exclude_pat}...")
        
        # Original count
        orig_study_count = dash_mgr.stats['study']['V']
        
        # Recalculate with exclusion
        dash_mgr.calculate_stats(excluded_patients=[exclude_pat])
        new_study_count = dash_mgr.stats['study']['V']
        
        print(f"Original Study V: {orig_study_count}")
        print(f"New Study V:      {new_study_count}")
        
        if new_study_count < orig_study_count:
            print("SUCCESS: Count decreased after exclusion.")
        else:
            print("FAILURE: Count did not decrease.")
            
        # Verify patient is gone
        if exclude_pat not in dash_mgr.stats['patient']:
             print(f"SUCCESS: Patient {exclude_pat} removed from stats.")
        else:
             print(f"FAILURE: Patient {exclude_pat} still in stats.")

if __name__ == "__main__":
    test_dashboard()
