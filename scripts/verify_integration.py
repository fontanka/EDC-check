import sys
import tkinter as tk
import unittest
from unittest.mock import MagicMock

# Add current dir to path
sys.path.append('c:/budgets')

class TestIntegration(unittest.TestCase):
    def setUp(self):
        try:
            from clinical_viewer1 import ClinicalDataMasterV30
            self.root = tk.Tk()
            self.root.withdraw() # Hide window
            self.app = ClinicalDataMasterV30(self.root)
        except Exception as e:
            self.fail(f"Failed to initialize app: {e}")

    def tearDown(self):
        if hasattr(self, 'root'):
            self.root.destroy()

    def test_view_builder_dependencies(self):
        """Check attributes required by ViewBuilder."""
        required_attrs = [
            'cb_site', 'cb_pat', 'view_mode', 
            'chk_hide_dup', 'chk_hide_future', 'search_var',
            'df_main', 'labels', 'tree', 'sdv_manager',
            'current_tree_data', 'current_patient_gaps'
        ]
        
        missing = []
        for attr in required_attrs:
            if not hasattr(self.app, attr):
                missing.append(attr)
        
        if missing:
            self.fail(f"Missing attributes required by ViewBuilder: {missing}")
        else:
            print("✓ All ViewBuilder dependencies present")

    def test_toolbar_dependencies(self):
        """Check attributes created by ToolbarSetup."""
        required_attrs = [
            'file_info_var', 'lbl_status', 'cutoff_var',
            'sdv_btn'
        ]
        
        missing = []
        for attr in required_attrs:
            if not hasattr(self.app, attr):
                missing.append(attr)
        
        if missing:
            self.fail(f"Missing attributes from ToolbarSetup: {missing}")
        else:
            print("✓ All ToolbarSetup attributes present")

    def test_treeview_configuration(self):
        """Verify Treeview is correctly configured."""
        if not hasattr(self.app, 'tree'):
            self.fail("Treeview not initialized")
        
        # Check columns
        columns = self.app.tree['columns']
        expected_cols = ('Value', 'Status', 'User', 'Date', 'Code')
        if not all(col in columns for col in expected_cols):
             self.fail(f"Treeview columns mismatch. Expected {expected_cols}, got {columns}")
        
        print("✓ Treeview configured correctly")

if __name__ == '__main__':
    unittest.main()
