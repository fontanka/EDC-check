import sys
import pandas as pd
import os
import glob
from unittest.mock import MagicMock

# Add current dir to path
sys.path.append('c:/budgets')

try:
    from config import VISIT_MAP
    from view_builder import ViewBuilder
    print("Imports successful")
    
    print(f"CWD: {os.getcwd()}")
    print("Files in CWD:")
    for f in os.listdir('.'):
        print(f" - {f}")

    # Find latest file
    files = glob.glob("*Innoventric*.xlsx")
    if not files:
        print("No data file found!")
        sys.exit(1)
        
    latest_file = max(files, key=os.path.getctime)
    print(f"Loading: {latest_file}")
    
    # Load Main sheet
    df = pd.read_excel(latest_file, sheet_name="Main", header=None, dtype=str, keep_default_na=False)
    codes = [str(c).strip() for c in df.iloc[0].tolist()]
    df.columns = codes
    df = df.iloc[2:] # Skip header/label rows
    
    print(f"Loaded DataFrame with {len(df)} rows and {len(df.columns)} columns.")
    
    # Initialize ViewBuilder
    app = MagicMock()
    app.labels = {}
    app.ae_lookup = {}
    app.df_ae = None
    app.chk_hide_dup.get.return_value = True
    app.chk_hide_future.get.return_value = False # Don't hide for debugging
    
    vb = ViewBuilder(app)
    
    # Check identifying columns
    print("\n--- Testing Column Identification ---")
    
    visit_counts = {}
    
    for col in df.columns:
        if col in ["Site #", "Screening #", "Status"]: continue
        
        info = vb._identify_column(col)
        if info:
            visit, form, cat = info
            visit_counts[visit] = visit_counts.get(visit, 0) + 1
            
            # Print samples for Baseline/Screening to verify
            if visit in ["Baseline", "Screening"] and visit_counts[visit] < 3:
                print(f"Col: {col} -> Visit: {visit}")
        else:
             pass # Failed identification
             
    print("\n--- Visit Counts ---")
    for v, count in visit_counts.items():
        print(f"{v}: {count} columns identified")
        
    # Check Data Content for a sample patient
    print("\n--- Checking Data Content for First Patient ---")
    if not df.empty:
        row = df.iloc[0]
        pat_id = row.get('Screening #', 'Unknown')
        print(f"Patient: {pat_id}")
        
        has_data = {v: False for v in VISIT_MAP.values()}
        
        for col in df.columns:
            val = row[col]
            if pd.isna(val) or str(val).strip() == "": continue
            
            info = vb._identify_column(col)
            if info:
                visit, _, _ = info
                if visit in has_data:
                    has_data[visit] = True
                    
        print("Visits with data (according to manual check):")
        for v, has in has_data.items():
            print(f"{v}: {has}")

except Exception as e:
    print(f"Error: {e}")
    import traceback
    traceback.print_exc()
