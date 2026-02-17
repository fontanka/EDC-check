import pandas as pd
import sys
import os
import re

# Add parent dir to path to import ae_manager
sys.path.append('c:/budgets')
from ae_manager import AEManager

# file_path = r"c:\budgets\verified\Innoventric_CLD-048_DM_SelectForms_16-02-2026_07-55_02_(UTC).xlsx"
file_path = r"c:\budgets\verified\Innoventric_CLD-048_DM_ProjectToOneFile_16-02-2026_07-55_40_(UTC).xlsx"

try:
    print("Loading data...")
    # Load all sheets
    xls = pd.read_excel(file_path, sheet_name=None, header=None, dtype=str, keep_default_na=False)
    
    print("Sheets found:", list(xls.keys()))
    for k, v in xls.items():
        if not v.empty:
            print(f"Sheet '{k}' Row 0: {v.iloc[0].tolist()[:5]}")
            
    # Check PR_ADD 732 specifically
    if 'PR_ADD 732 ' in xls:
        print("Checking PR_ADD 732 content...")
        pr_df = xls['PR_ADD 732 ']
        print(f"Columns: {pr_df.iloc[0].tolist()}")
            
    # 1. Main Sheet
    # Find sheet with 'TV_IMP_IMPDAT' (Implantation Date) in header
    main_sheet_name = None
    for name, df in xls.items():
        if not df.empty and any('IMP' in str(c).upper() for c in df.iloc[0].tolist()):
            main_sheet_name = name
            break
            
    if not main_sheet_name:
         print("Could not find sheet with Implantation Date. Trying '732'...")
         for name in xls.keys():
             if '732' in name and 'AE' not in name:
                 main_sheet_name = name
                 break
                 
    if not main_sheet_name:
         print("Still no main sheet. Searching for any 'Form'...")
         main_sheet_name = next((n for n in xls.keys() if "Form" in n), None)

    print(f"Main Sheet: {main_sheet_name}")
    raw_main = xls[main_sheet_name]
    codes = [str(c).strip() for c in raw_main.iloc[0].tolist()]
    df_main = raw_main.iloc[2:].copy()
    df_main.columns = codes
    
    # 2. AE Sheet
    # Find sheet with AE Term column
    ae_sheet_name = None
    for name, df in xls.items():
        if not df.empty and any('LOGS_AE_AETERM' in str(c) for c in df.iloc[0].tolist()):
            ae_sheet_name = name
            break
    
    if not ae_sheet_name:
        ae_sheet_name = next((n for n in xls.keys() if "AE" in n), None)
        
    print(f"AE Sheet: {ae_sheet_name}")
    raw_ae = xls[ae_sheet_name]
    ae_codes = [str(c).strip() for c in raw_ae.iloc[0].tolist()]
    df_ae = raw_ae.iloc[2:].copy() # AE also has 2 header rows usually? Viewer logic: self.df_main = raw.iloc[2:], but distinct for AE?
    # Viewer load_extra_sheets just grabs self.xls[ae_sheet]. 
    # But usually AE sheets follow same structure. Let's assume Row 0 codes, Row 1 labels, Row 2 data.
    # Check if 'Screening #' is in codes.
    if 'Screening #' not in ae_codes:
        # Maybe row 0 is title? Try row 1?
        # Let's inspect raw_ae in output if fails.
        pass
        
    df_ae.columns = ae_codes
    
    # Clean column names (viewer does strip loop?)
    # Viewer: self.df_ae.columns = [c.strip() for c in self.df_ae.columns]
    
    mgr = AEManager(df_main, df_ae)
    
    print("Running get_summary_stats(exclude_pre_proc=True)...")
    stats = mgr.get_summary_stats(exclude_pre_proc=True)
    print("Success!")
    print(f"Total AEs: {stats['total_aes']}")

except Exception as e:
    import traceback
    traceback.print_exc()
