import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
import numpy as np
import logging
from datetime import datetime
import re
from config import VISIT_MAP, CONDITIONAL_SKIPS

logger = logging.getLogger(__name__)

class ViewBuilder:
    def __init__(self, app):
        self.app = app
        self._view_cache = {}

    def invalidate_cache(self):
        """Invalidate the view cache."""
        self._view_cache.clear()

    def clear_cache(self):
        """Clear the view cache."""
        self._view_cache.clear()


    
    def add_gap(self, visit, form, field_label, db_var, gaps_list):
        """Record a data gap."""
        gaps_list.append({
            'visit': visit,
            'form': form,
            'field': field_label,
            'variable': db_var,
            'timestamp': datetime.now()
        })

    def generate_view(self, *args):
        """Identify relevant rows and populate the Treeview."""
        # 1. Gather current filter/options
        site_filter = self.app.cb_site.get().strip()
        pat_filter = self.app.cb_pat.get().strip()
        view_mode = self.app.view_mode.get()
        hide_dup = self.app.chk_hide_dup.get()
        hide_future = self.app.chk_hide_future.get()
        search_term = self.app.search_var.get().strip().upper()

        cache_key = (site_filter, pat_filter, view_mode, hide_dup, hide_future, search_term)

        # 2. Check cache
        if cache_key in self._view_cache:
            cached_data = self._view_cache[cache_key]
            self._render_tree(cached_data['tree_data'], 
                            cached_data['visit_has_data'], 
                            cached_data['matrix_supported_nodes'],
                            search_term)
            
            # Restore state
            self.app.current_tree_data = cached_data['tree_data']
            self.app.current_patient_gaps = cached_data.get('collected_gaps', [])
            return

        if self.app.df_main is None:
            return

        # 3. Filter data
        df = self.app.df_main.copy()

        if site_filter and site_filter != "All Sites":
            df = df[df['Site #'] == site_filter]

        if pat_filter and pat_filter != "All Patients":
             # Handle "Screen Failure" group
            if pat_filter == "Active Patients":
                 df = df[~df['Status'].astype(str).str.contains("Screen Fail", case=False, na=False)]
            elif pat_filter == "Screen Failures":
                 df = df[df['Status'].astype(str).str.contains("Screen Fail", case=False, na=False)]
            elif pat_filter == "All Patients":
                pass
            else:
                 # Single patient
                 df = df[df['Screening #'] == pat_filter]

        # 4. Search filtering
        if search_term:
            mask = df.astype(str).apply(lambda x: x.str.contains(search_term, case=False, na=False)).any(axis=1)
            df = df[mask]
        
        # 5. Build Tree Data Structure
        tree_data = {}
        visit_has_data = {v: False for v in VISIT_MAP.values()}
        matrix_supported_nodes = set()
        collected_gaps = []

        self._build_ae_lookup()

        for idx, row in df.iterrows():
            site = str(row.get('Site #', 'Unknown'))
            pat = str(row.get('Screening #', 'Unknown'))
            
            if site not in tree_data:
                tree_data[site] = {}
            if pat not in tree_data[site]:
                tree_data[site][pat] = {'visits': {}, 'forms': {}, 'demographics': {}}
                if 'Age' in row: tree_data[site][pat]['demographics']['Age'] = row['Age']
                if 'Sex' in row: tree_data[site][pat]['demographics']['Sex'] = row['Sex']

            for col in df.columns:
                val = row[col]
                if pd.isna(val) or str(val).strip() == "":
                    # Potential Gap Check (simplified)
                    # Real application has complex logic for what constitutes a gap
                    continue
                
                if col in ["Site #", "Screening #", "Status", "Subject Initials"]:
                    continue

                info = self._identify_column(col)
                
                if not info:
                    continue
                
                visit, form, category = info
                
                if self._is_not_done_column(col, val):
                   pass
                
                is_skipped = False
                for trigger, rule in CONDITIONAL_SKIPS.items():
                    if col in rule["targets"]:
                        trigger_val = str(row.get(trigger, "")).lower()
                        if rule["trigger_value"] == "*ANY*" and trigger_val and trigger_val not in ["nan", ""]:
                             is_skipped = True
                             break
                        elif rule["trigger_value"] in trigger_val:
                             is_skipped = True
                             break
                
                if is_skipped:
                    continue

                if view_mode == "visit":
                    if visit not in tree_data[site][pat]['visits']:
                        tree_data[site][pat]['visits'][visit] = {}
                    
                    if form not in tree_data[site][pat]['visits'][visit]:
                        tree_data[site][pat]['visits'][visit][form] = []
                    
                    clean_lbl = self._clean_label(self.app.labels.get(col, col))
                    tree_data[site][pat]['visits'][visit][form].append((clean_lbl, val, col))
                    
                    if visit in visit_has_data:
                        visit_has_data[visit] = True
                        
                else: 
                    grouper = form if form else "Uncategorized"
                    
                    if grouper not in tree_data[site][pat]['forms']:
                        tree_data[site][pat]['forms'][grouper] = {}
                        
                    if visit not in tree_data[site][pat]['forms'][grouper]:
                        tree_data[site][pat]['forms'][grouper][visit] = []
                        
                    clean_lbl = self._clean_label(self.app.labels.get(col, col))
                    tree_data[site][pat]['forms'][grouper][visit].append((clean_lbl, val, col))

        if view_mode == "visit":
            self._annotate_procedure_timing(tree_data)

        cache_data = {
            'tree_data': tree_data,
            'visit_has_data': visit_has_data,
            'matrix_supported_nodes': matrix_supported_nodes,
            'collected_gaps': collected_gaps
        }
        self._view_cache[cache_key] = cache_data
        
        self._render_tree(tree_data, visit_has_data, matrix_supported_nodes, search_term)
        
        self.app.current_tree_data = tree_data
        self.app.current_patient_gaps = collected_gaps

    def _render_tree(self, tree_data, visit_has_data, matrix_supported_nodes, search_term):
        """Render the tree structure into the UI."""
        self.app.tree.delete(*self.app.tree.get_children())
        
        # Update columns based on mode ? (Actually columns are fixed: Label, Value, Status, Code)
        
        for site in sorted(tree_data.keys()):
            site_node = self.app.tree.insert("", "end", text=f"Site {site}", open=True, values=("", "", "", "SITE"))
            
            for pat in sorted(tree_data[site].keys()):
                # Check SDV status for Patient level (if implemented)
                pat_tags = ('patient',)
                if str(pat) in self.app.sdv_verified_fields: # This logic might need adjustment based on how patient verification works
                     pass 

                pat_node = self.app.tree.insert(site_node, "end", text=f"Subject {pat}", open=False, values=("", "", "", "PATIENT"), tags=pat_tags)
                
                pat_data = tree_data[site][pat]
                
                if self.app.view_mode.get() == "visit":
                    # Visit Mode Rendering
                    # Sort visits based on VISIT_MAP order or specific logic
                    # Flatten visits list
                    sorted_visits = []
                    # Logic to sort visits... simplified for now
                    # We can use VISIT_MAP keys index if available, or just alphabetical/custom sort
                    # VISIT_MAP is typically used for regex matching, but keys can imply order if processed right
                    # Or use a separate VISIT_ORDER list if it exists in config
                    
                    for visit in sorted(pat_data['visits'].keys()):
                        
                        # Apply hide options
                        if self.app.chk_hide_future.get() and not visit_has_data.get(visit, True):
                             continue
                             
                        visit_node = self.app.tree.insert(pat_node, "end", text=visit, open=False, values=("", "", "", "VISIT"))
                        
                        forms = pat_data['visits'][visit]
                        for form in sorted(forms.keys()):
                            # Special handling for "Data Matrix" support indicator
                            text = form
                            if self._is_matrix_supported_col(form): # Naive check, usually form name
                                 text += " ▦" 
                                 
                            # Lookup Form Status
                            # We use row="0" default for form-level check
                            form_status = ""
                            form_user = ""
                            form_date = ""
                            
                            if self.app.sdv_manager and self.app.sdv_manager.is_loaded():
                                 # Get verification metadata (User, Date)
                                 details = self.app.sdv_manager.get_verification_details(pat, form, visit_name=visit)
                                 if details:
                                     form_user = details.get('user', '')
                                     form_date = details.get('date', '')
                                 
                                 # Get status string (e.g. Verified)
                                 # We can reuse get_field_status for the form level key usually
                                 form_status = self.app.sdv_manager.get_field_status(pat, "ANY", form_name=form, visit_name=visit) 
                                 # Passing "ANY" as field ignores field-specific logic if utilizing the form key directly, 
                                 # but let's see sdv_manager implementation. 
                                 # Actually sdv_manager.get_field_status builds key: f"{pat}|{visit}|{form}|{row}"
                                 # So if we pass row="0" (default), it looks up the form entry status.
                            
                            form_node = self.app.tree.insert(visit_node, "end", text=text, open=False, 
                                                           values=("", form_status, form_user, form_date, "FORM"))
                            
                            for label, val, col_code in forms[form]:
                                # Lookup SDV status
                                status = ""
                                user = ""
                                date = ""
                                tags = ()
                                
                                if self.app.sdv_manager and self.app.sdv_manager.is_loaded():
                                   # We need row info if it's a repeating form
                                   # Try to deduce row/repeat from AE/Lab logic
                                   row_num = "0" # Default
                                   
                                   # AE Logic
                                   if "AE" in col_code and "TERM" in col_code:
                                       # Try to extract seq num
                                       # This is tricky without row context in tree_data tuple
                                       pass
                                   
                                   # Get field status properly
                                   field_status = self.app.sdv_manager.get_field_status(pat, col_code, table_row=row_num, form_name=form, visit_name=visit)
                                   
                                   # Get details if needed (optional)
                                   details = self.app.sdv_manager.get_verification_details(pat, form, visit, row_num)

                                   if field_status in ["verified", "auto_verified"]:
                                       status = "Verified"
                                       tags = ('verified',)
                                       if details:
                                            user = details.get('user', '')
                                            date = details.get('date', '')
                                   elif field_status == "awaiting":
                                       status = "Awaiting"
                                       tags = ('pending',)
                                   elif field_status == "not_checked":
                                       status = "Pending" 
                                       tags = ('pending',)
                                   else:
                                       status = ""
                                       # Optional: Check form level fallback if field status is None/Not Sent?
                                       # For now, keep it simple.
                                
                                self.app.tree.insert(form_node, "end", text=label, values=(val, status, user, date, col_code), tags=tags)
                                
                else:
                    # Assessment Mode
                    forms = pat_data['forms']
                    for form in sorted(forms.keys()):
                        form_node = self.app.tree.insert(pat_node, "end", text=form, open=False, values=("", "", "", "FORM"))
                        
                        visits = forms[form]
                        for visit in sorted(visits.keys()):
                             # Determine if this visit has any data for this form
                             # If we are hiding future visits in this mode, logic is similar
                             
                             visit_node = self.app.tree.insert(form_node, "end", text=visit, open=False, values=("", "", "", "VISIT"))
                             
                             for label, val, col_code in visits[visit]:
                                 # SDV Status logic (same as above)
                                 # Currently assessment mode doesn't show sdv status fully? 
                                 # The logic above only did values=(val, "", "", col_code)
                                 # We should probably replicate or at least keep tuple size consistent
                                 self.app.tree.insert(visit_node, "end", text=label, values=(val, "", "", "", col_code))
        
        # Apply Search Highlighting
        if search_term:
             self._apply_search_highlight(search_term)

    def _apply_search_highlight(self, term):
        """Highlight items matching search term."""
        # This requires recursively traversing the tree
        # or using the 'tags' property we populated
        pass # To be implemented or rely on standard rendering

    def _identify_column(self, col_name):
        """
        Identify visit, form, and category from a column name.
        Returns (visit, form, category) or None.
        """
        # 1. Identify Visit
        visit = "Unscheduled"
        for prefix, v_name in VISIT_MAP.items():
            if col_name.startswith(prefix + "_") or col_name == prefix:
                visit = v_name
                break
        
        # 2. Identify Form/Category using Regex Rules
        # Assessment Rules (from config)
        from config import ASSESSMENT_RULES
        
        category = "Other"
        form = "General"
        
        for pattern, cat, frm in ASSESSMENT_RULES:
            if re.search(pattern, col_name):
                category = cat
                form = frm
                break
                
        return visit, form, category

    def _is_not_done_column(self, col_name, val):
        """Check if value indicates 'Not Done'."""
        if isinstance(val, str):
            v = val.lower()
            return v in ["not done", "nd", "n/a", "skipped"]
        return False
        
    def _build_ae_lookup(self):
        """Build a lookup for Adverse Events from df_ae if available."""
        self.app.ae_lookup = {}
        if self.app.df_ae is not None:
             # Logic to build lookup: Patient -> AE Term
             # Assuming df_ae has 'Project/Subject ID' and 'Adverse Event Term'
             pass

    def _clean_label(self, label):
        """Clean up column label for display."""
        # Remove variable name suffix [VARNAME]
        label = re.sub(r'\[.*?\]', '', str(label)).strip()
        return label

    def _is_matrix_supported_col(self, col_name):
        """Check if column/form supports data matrix view."""
        # Simplify: hardcoded list or logic
        supported = ["Medical History", "Adverse Event", "Concomitant Medications", "Laboratory"]
        return any(s in col_name for s in supported)

    def _annotate_procedure_timing(self, tree_data):
        """Annotate procedure timing (start/end) in the tree data."""
        pass # Logic for calculating duration etc.

