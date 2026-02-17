"""
Procedure Timing Export Module

Generates a Procedure Timing matrix with adjustable row ordering.
Excludes Screen Failures by default.
Uses labels from the main application for proper display names.
"""

import pandas as pd
from datetime import datetime


class ProcedureTimingExporter:
    def __init__(self, df_main, labels=None):
        self.df_main = df_main
        self.labels = labels or {}
        self.timing_fields = []  # List of (col_name, label) tuples
        self._discover_timing_columns()
    
    def _discover_timing_columns(self):
        """Auto-discover TV_PR_TIM columns and their labels."""
        timing_cols = []
        seen_labels = {}
        
        for col in self.df_main.columns:
            col_str = str(col)
            # Procedure timing columns - exclude "_NR" (not recorded) columns
            if 'TV_PR_TIM' in col_str and '_NR' not in col_str:
                label = self.labels.get(col_str, col_str)
                # Clean up label - remove "/ not recorded" suffix, trailing slashes
                label = label.replace(' / not recorded', '').strip()
                if ' / time' in label:
                    label = label.replace(' / time', '').strip()
                if label.endswith('/'):
                    label = label[:-1].strip()
                
                # Disambiguate based on column suffix
                if '_POST' in col_str:
                    if '(post-procedure)' not in label.lower() and '(post)' not in label.lower():
                        label = f"{label} (Post-Procedure)"
                elif '_PRE' in col_str:
                     if '(pre-procedure)' not in label.lower() and '(pre)' not in label.lower():
                        label = f"{label} (Pre-Procedure)"
                elif '_CVC' in col_str and '_POST' not in col_str:
                    # Special case for CVC which has a POST counterpart with same label
                    if 'Cardiac and Venus' in label and '(pre' not in label.lower():
                         label = f"{label} (Pre-Procedure)"

                timing_cols.append((col_str, label))
        
        # Sort by column name to get consistent order
        timing_cols.sort(key=lambda x: x[0])
        
        # Final pass to ensure uniqueness
        final_cols = []
        existing_labels = set()
        for col, label in timing_cols:
            if label in existing_labels:
                # If label still duplicated, append suffix code
                suffix = col.replace('TV_PR_TIM_', '')
                label = f"{label} ({suffix})"
            existing_labels.add(label)
            final_cols.append((col, label))
            
        self.timing_fields = final_cols
    
    def set_field_order(self, new_order):
        """Set custom field order. new_order is list of (col_name, label) tuples."""
        self.timing_fields = new_order
    
    def get_field_order(self):
        """Get current field order."""
        return self.timing_fields
    
    def get_patient_timing(self, patient_id):
        """Get all timing data for a patient."""
        rows = self.df_main[self.df_main['Screening #'] == patient_id]
        if rows.empty:
            return None
        
        row = rows.iloc[0]
        data = {"Patient": patient_id}
        
        for col_name, label in self.timing_fields:
            if col_name in row.index:
                val = row[col_name]
                if pd.notna(val):
                    # Format time if it looks like datetime
                    val_str = str(val)
                    if 'T' in val_str:
                        # Extract just the time part from ISO datetime
                        try:
                            dt = pd.to_datetime(val)
                            data[label] = dt.strftime("%H:%M")
                        except (ValueError, TypeError):
                            data[label] = val_str.split('T')[-1][:5] if 'T' in val_str else val_str
                    elif ':' in val_str:
                        # Already a time, keep first 5 chars (HH:MM)
                        data[label] = val_str[:5]
                    else:
                        data[label] = val_str
                else:
                    data[label] = ""
            else:
                data[label] = ""
        
        return data
    
    def generate_matrix(self, patient_ids):
        """Generate timing matrix for multiple patients."""
        records = []
        for pid in patient_ids:
            timing = self.get_patient_timing(pid)
            if timing:
                records.append(timing)
        
        if not records:
            return None
        
        # Create DataFrame with columns in order
        columns = ["Patient"] + [label for _, label in self.timing_fields]
        df = pd.DataFrame(records, columns=columns)
        
        return df
    
    def export_to_excel(self, patient_ids):
        """Export timing matrix to Excel (bytes)."""
        from io import BytesIO
        
        df = self.generate_matrix(patient_ids)
        if df is None:
            return None
        
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='Procedure Timing', index=False)
        output.seek(0)
        return output.getvalue()
    
    def export_to_csv(self, patient_ids):
        """Export timing matrix to CSV."""
        df = self.generate_matrix(patient_ids)
        if df is None:
            return None
        return df.to_csv(index=False)
