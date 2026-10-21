import unittest
import tkinter as tk
from src.gui.dialogs import DateFixDialog, SchoolBlockDialog, CohortBlockDialog, GradDateBlockDialog


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


if __name__ == '__main__':
    unittest.main()
