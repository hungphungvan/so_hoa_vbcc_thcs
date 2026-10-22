import unittest
import tkinter as tk
from src.gui.dialogs import DateFixDialog
from src.gui.custom_tools_dialog import CustomToolsDialog, CustomToolsView, ManageProfilesDialog
from src.gui.column_mapping_dialog import ColumnMappingDialog
from src.smart_merger import MappingProfileManager
from src.gui.export_dialog import ExportDialog
from src.gui.data_management_view import DataManagementView
from src.db_manager import DatabaseManager


class TestDialogs(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        try:
            cls.root = tk.Tk()
            cls.root.withdraw()
            cls.tk_available = True
        except Exception:
            cls.tk_available = False

    @classmethod
    def tearDownClass(cls):
        if getattr(cls, 'tk_available', False) and getattr(cls, 'root', None):
            try:
                cls.root.destroy()
            except Exception:
                pass

    def setUp(self):
        if not getattr(self, 'tk_available', False):
            self.skipTest("Tkinter display not available in current test environment")

    def test_date_fix_dialog_methods_and_init(self):
        # Kiểm tra đầy đủ các phương thức xử lý sự kiện
        self.assertTrue(hasattr(DateFixDialog, '_skip'))
        self.assertTrue(hasattr(DateFixDialog, '_keep'))
        self.assertTrue(hasattr(DateFixDialog, '_skip_all'))

        # Khởi tạo thử DateFixDialog
        dlg = DateFixDialog(self.root, {
            "name": "Nguyen Van A",
            "stt": 1,
            "dob": "99/99/9999",
            "context_rows": []
        })
        self.assertIsNotNone(dlg)
        dlg._skip()
        self.assertEqual(dlg.result, "")

    def test_data_management_view_init(self):
        """Kiểm tra khởi tạo widget DataManagementView không bị lỗi."""
        import tempfile, shutil
        from pathlib import Path
        tmp_dir = Path(tempfile.mkdtemp())
        try:
            db_mgr = DatabaseManager(tmp_dir / "test.db")
            view = DataManagementView(self.root, db_manager=db_mgr)
            self.assertIsNotNone(view)
            self.assertEqual(view.current_page, 1)
            self.assertEqual(view.total_records, 0)
            view.destroy()
        finally:
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def test_export_dialog_init(self):
        """Kiểm tra khởi tạo widget ExportDialog không bị lỗi."""
        import tempfile
        from pathlib import Path
        tmp_dir = Path(tempfile.mkdtemp())
        try:
            db_mgr = DatabaseManager(tmp_dir / "test.db")
            dlg = ExportDialog(self.root, db_manager=db_mgr)
            self.assertIsNotNone(dlg)
            self.assertTrue(hasattr(dlg, 'cohort_listbox'))
            self.assertTrue(hasattr(dlg, 'btn_start_export'))
            dlg.destroy()
        finally:
            import shutil
            shutil.rmtree(tmp_dir, ignore_errors=True)

    def test_custom_tools_view_and_dialog_init(self):
        """Kiểm tra khởi tạo CustomToolsView và CustomToolsDialog không bị lỗi."""
        view = CustomToolsView(self.root)
        self.assertIsNotNone(view)
        self.assertTrue(hasattr(view, 'split_file_var'))
        self.assertTrue(hasattr(view, 'btn_run_split'))
        self.assertTrue(hasattr(view, 'btn_run_merge'))
        self.assertTrue(hasattr(view, 'btn_run_smart_merge'))
        self.assertTrue(hasattr(view, 'smart_files_list'))
        self.assertTrue(hasattr(view, 'btn_run_convert'))
        self.assertTrue(hasattr(view, 'convert_files_list'))
        view.destroy()

        dlg = CustomToolsDialog(self.root)
        self.assertIsNotNone(dlg)
        self.assertTrue(hasattr(dlg, 'view'))
        dlg.destroy()

    def test_column_mapping_and_manage_profiles_dialog_init(self):
        """Kiểm tra khởi tạo ColumnMappingDialog và ManageProfilesDialog."""
        raw_headers = ["STT", "Họ và tên", "Ngày sinh", "CCCD"]
        formatted_headers = ["[A] STT", "[B] Họ và tên", "[C] Ngày sinh", "[D] CCCD"]
        suggested = {1: 0, 5: 1, 9: 2, 7: 3}

        dlg_map = ColumnMappingDialog(
            parent=self.root,
            file_name="test_school.xlsx",
            sheet_name="Sheet1",
            raw_headers=raw_headers,
            formatted_headers=formatted_headers,
            suggested_mapping=suggested
        )
        self.assertIsNotNone(dlg_map)
        self.assertTrue(hasattr(dlg_map, 'target_vars'))
        self.assertEqual(len(dlg_map.target_vars), 28)
        dlg_map.destroy()

        prof_mgr = MappingProfileManager()
        dlg_prof = ManageProfilesDialog(parent=self.root, profile_mgr=prof_mgr)
        self.assertIsNotNone(dlg_prof)
        self.assertTrue(hasattr(dlg_prof, 'tree'))
        dlg_prof.destroy()


if __name__ == '__main__':
    unittest.main()
