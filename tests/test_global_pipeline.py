import sys
import unittest
from pathlib import Path
import openpyxl

from src.lookup_manager import LookupManager, find_data_file
from src.excel_processor import ExcelProcessor

class TestGlobalPipeline(unittest.TestCase):
    def setUp(self):
        self.lookup_mgr = LookupManager()
        self.processor = ExcelProcessor(self.lookup_mgr)

    def test_pipeline_stages_order(self):
        # Tạo file Excel giả lập gồm 2 sheets với lỗi Ngày sinh, Tên trường chưa chuẩn, và Niên khóa
        test_file = Path("/tmp/test_pipeline_multi_sheets.xlsx")
        wb = openpyxl.Workbook()
        
        # Sheet 1: Dạng chuẩn 28 cột
        ws1 = wb.active
        ws1.title = "Lop 9A"
        
        # Header ở dòng 2
        header_row = ["STT", "Số hiệu", "Số vào sổ", "Tên văn bằng", "Họ và tên", "Mã", "CCCD", "Hộ chiếu",
                      "Ngày sinh", "Giới tính", "Dân tộc", "Quốc tịch", "Nơi sinh", "Tên trường", "Niên khóa",
                      "Xếp loại", "Hình thức", "Năm TN", "Ngày cấp", "Số QĐ", "Ngày QĐ", "Người ký", "CQ cấp",
                      "Ghi chú", "Trạng thái", "Lý do", "Mã ĐV", "Tên ĐV"]
        ws1.append(["Tiêu đề"])
        ws1.append(header_row)
        
        # Dòng 1: Ngày sinh lỗi format '31/02/2005'
        row1 = [1, "", "", "", "Nguyễn Văn An", "", "", "", "31/02/2005", "Nam", "Kinh", "Việt Nam", "Hà Nội", "THCS Chưa Chuẩn", "2016-2020", "", "", "", "15/06/2020"] + [""] * 9
        # Dòng 2: Ngày sinh chuẩn
        row2 = [2, "", "", "", "Trần Thị Bích", "", "", "", "10/05/2005", "Nữ", "Kinh", "Việt Nam", "Hà Nội", "THCS Chưa Chuẩn", "2016-2020", "", "", "", "15/06/2020"] + [""] * 9
        ws1.append(row1)
        ws1.append(row2)

        # Sheet 2
        ws2 = wb.create_sheet(title="Lop 9B")
        ws2.append(["Tiêu đề"])
        ws2.append(header_row)
        # Dòng 3: Ngày sinh sai logic '99/99/2005', trường trống, niên khóa chỉ có năm '2020'
        row3 = [1, "", "", "", "Lê Văn Cường", "", "", "", "99/99/2005", "Nam", "Kinh", "Việt Nam", "Hà Nội", "", "2020", "", "", "", "2020"] + [""] * 9
        ws2.append(row3)

        wb.save(test_file)
        wb.close()

        dialog_events = []
        def mock_dialog(d_type, *args):
            dialog_events.append(d_type)
            if d_type == "date_fix":
                return "15/05/2005", False
            elif d_type == "school_block":
                return "Trường THCS Chuẩn Mẫu", True
            elif d_type == "cohort_block":
                return "2016-2020", True
            return None, False

        unit_info = {
            "code": "TEST_001",
            "name": "Phòng GD&ĐT Thử Nghiệm",
            "schools": ["Trường THCS Chuẩn Mẫu"]
        }

        output_dir = Path("/tmp/test_out_pipeline")
        all_students, stats = self.processor.process_all_files(
            files=[test_file],
            unit_info=unit_info,
            output_dir=output_dir,
            dialog_callback=mock_dialog
        )

        # Kiểm tra thứ tự các dialog:
        # 1. Tất cả date_fix phải xuất hiện TRƯỚC
        # 2. Tiếp theo là school_block
        # 3. Cuối cùng là cohort_block
        print(f"Thứ tự dialog đã kích hoạt: {dialog_events}")
        
        # Phải có date_fix
        self.assertIn("date_fix", dialog_events)
        # Tìm vị trí xuất hiện cuối cùng của date_fix
        last_date_idx = max(i for i, d in enumerate(dialog_events) if d == "date_fix")
        # Tìm vị trí xuất hiện đầu tiên của school_block (nếu có)
        if "school_block" in dialog_events:
            first_school_idx = min(i for i, d in enumerate(dialog_events) if d == "school_block")
            self.assertGreater(first_school_idx, last_date_idx, "School dialog phải xuất hiện SAU DateFix dialog!")
        
        if "cohort_block" in dialog_events and "school_block" in dialog_events:
            last_school_idx = max(i for i, d in enumerate(dialog_events) if d == "school_block")
            first_cohort_idx = min(i for i, d in enumerate(dialog_events) if d == "cohort_block")
            self.assertGreater(first_cohort_idx, last_school_idx, "Cohort dialog phải xuất hiện SAU School dialog!")

        # Kiểm tra kết quả học sinh
        self.assertEqual(len(all_students), 3)
        for st in all_students:
            self.assertEqual(st["school"], "Trường THCS Chuẩn Mẫu")
            self.assertEqual(st["cohort"], "2016-2020")
            # Cột S: ngày cấp bằng được tự động điền chuẩn từ nhóm (Trường, Niên khóa)
            self.assertEqual(st["grad_date"], "15/06/2020")
            # Cột Y: Đã số hóa đầy đủ, chính xác (do ngày sinh đã được sửa thành 15/05/2005)
            self.assertEqual(st["final_row"][24], "Đã số hóa đầy đủ, chính xác")

        print("--> Test Global Pipeline PASSED successfully!")

if __name__ == "__main__":
    unittest.main()
