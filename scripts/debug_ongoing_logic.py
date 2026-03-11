import pandas as pd
import os
import re

file_path = r"c:\budgets\verified\Innoventric_CLD-048_DM_SelectForms_16-02-2026_07-55_02_(UTC).xlsx"
sheet_name = "AE_732"

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    df.columns = [re.sub(r'\s+', ' ', str(c)).strip() for c in df.columns]
    
    screening_col = 'Screening #'
    ae_num_col = 'Template number'
    ongoing_col = 'LOGS_AE_AEONGO'
    outcome_col = 'LOGS_AE_AEOUT'
    
    # Deduplication (Simplified version of Manager)
    df_valid = df.dropna(subset=[ae_num_col])
    df_valid['__pop_count'] = df_valid.apply(lambda r: sum(1 for v in r if isinstance(v, str) and v.strip()), axis=1)
    df_sorted = df_valid.sort_values([screening_col, ae_num_col, '__pop_count'], ascending=[True, True, False])
    df_unique = df_sorted.drop_duplicates(subset=[screening_col, ae_num_col], keep='first')
    
    def is_checked(val):
        return str(val).lower() in ['yes', 'y', '1', 'true', 'checked', 'ongoing']

    # 1. Base Count
    marked_ongoing = df_unique[df_unique[ongoing_col].apply(is_checked)]
    with open("debug_results.txt", "w", encoding="utf-8") as f:
        f.write(f"Explicitly Marked Ongoing: {len(marked_ongoing)}\n")
        
        f.write("\n--- Explicitly Marked Ongoing BUT has End Date ---\n")
        end_date_col = 'LOGS_AE_AEENDTC'
        
        for idx, row in marked_ongoing.iterrows():
            end = str(row.get(end_date_col, '')).lower().strip()
            has_end_date = (end != '' and end != 'nan' and end != 'nat')
            if has_end_date:
                f.write(f"Patient {row[screening_col]} AE {row[ae_num_col]}: EndDate='{row.get(end_date_col)}' Outcome='{row.get(outcome_col)}'\n")

        # 2. Analyze "Recovering/Resolving" and others
        not_marked = df_unique[~df_unique[ongoing_col].apply(is_checked)]
        
        f.write("\n--- NOT marked Ongoing, but with Suspicious Outcome ---\n")
        suspicious = ['recovering/resolving', 'not recovered/not resolved', 'not recovered', 'ongoing', 'unknown']
        
        for idx, row in not_marked.iterrows():
            out = str(row.get(outcome_col, '')).lower().strip()
            end = str(row.get(end_date_col, '')) # Keep raw string
            
            if any(s in out for s in suspicious):
                f.write(f"Patient {row[screening_col]} AE {row[ae_num_col]}: Outcome='{row.get(outcome_col)}' EndDateRaw='{end}'\n")

except Exception as e:
    with open("debug_results.txt", "w", encoding="utf-8") as f:
        f.write(str(e))
