import pandas as pd
import os
import re

# Use current directory or adjust if needed
file_path = r"c:\budgets\verified\Innoventric_CLD-048_DM_SelectForms_16-02-2026_07-55_02_(UTC).xlsx"
sheet_name = "AE_732"

try:
    if not os.path.exists(file_path):
        print(f"File not found: {file_path}")
        exit()
        
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    
    screening_col = 'Screening #'
    ae_num_col = 'Template number'
    ongoing_col = 'LOGS_AE_AEONGO'
    outcome_col = 'LOGS_AE_AEOUT'
    end_date_col = 'LOGS_AE_AEENDTC'
    
    # Filter out rows where 'AE #' is empty or NaN
    df_valid = df.dropna(subset=[ae_num_col])
    
    # Deduplication Logic
    df_valid['__pop_count'] = df_valid.apply(lambda r: sum(1 for v in r if isinstance(v, str) and v.strip()), axis=1)
    df_sorted = df_valid.sort_values([screening_col, ae_num_col, '__pop_count'], ascending=[True, True, False])
    df_unique = df_sorted.drop_duplicates(subset=[screening_col, ae_num_col], keep='first')
    
    # helper
    def is_checked(val):
        return str(val).lower() in ['yes', 'y', '1', 'true', 'checked', 'ongoing']

    # 1. Identify rows that are NOT marked ongoing
    not_marked_ongoing = df_unique[~df_unique[ongoing_col].apply(is_checked)]
    print(f"Rows NOT marked Ongoing: {len(not_marked_ongoing)}")
    
    # 2. Check Outcome
    print("\n--- Rows NOT marked Ongoing but with suspicious Outcome ---")
    suspicious_outcomes = ['not recovered', 'resolving', 'unknown', 'ongoing']
    
    for idx, row in not_marked_ongoing.iterrows():
        out = str(row.get(outcome_col, '')).lower()
        end = str(row.get(end_date_col, '')).lower()
        
        is_suspicious = any(s in out for s in suspicious_outcomes)
        if is_suspicious:
            print(f"Patient {row[screening_col]}, AE {row[ae_num_col]} -> Outcome: '{row.get(outcome_col)}', EndDate: '{row.get(end_date_col)}', OngoingCol: '{row.get(ongoing_col)}'")
            
    # 3. Check Empty End Date (and not fatal)
    print("\n--- Rows NOT marked Ongoing with Empty End Date (and not Fatal) ---")
    for idx, row in not_marked_ongoing.iterrows():
        out = str(row.get(outcome_col, '')).lower()
        end = str(row.get(end_date_col, '')).strip()
        
        if (end == '' or end == 'nan') and 'fatal' not in out and 'recovered' not in out:
             print(f"Patient {row[screening_col]}, AE {row[ae_num_col]} -> Outcome: '{row.get(outcome_col)}', EndDate: '{row.get(end_date_col)}'")

except Exception as e:
    import traceback
    traceback.print_exc()
