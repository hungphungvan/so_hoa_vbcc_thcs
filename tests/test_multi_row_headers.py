import unittest
from pathlib import Path
import openpyxl

from src.custom_excel_manager import (
    detect_header_and_columns,
    is_header_duplicate_row,
    ExcelMerger
)
from src.smart_merger import SmartMerger, auto_suggest_mapping


class TestMultiRowHeaders(unittest.TestCase):
    """Bộ kiểm thử cho tính năng nhận diện dòng tiêu đề đơn tầng (Single-row Header), đảm bảo an toàn 100% dữ liệu."""

    def setUp(self):
        self.temp_dir = Path("tests/temp_test_headers")
        self.temp_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        import shutil
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_single_row_header_detection(self):
        """Đảm bảo tiêu đề 1 dòng hoạt động chuẩn xác và dữ liệu bắt đầu ngay dòng tiếp theo."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "TestSheet"
        ws.append(["STT", "Họ và tên", "Ngày sinh", "Giới tính", "Nơi sinh"])
        ws.append([1, "Nguyễn Văn An", "15/05/2008", "Nam", "Hà Nội"])
        ws.append([2, "Trần Thị Bích", "20/10/2008", "Nữ", "Hải Phòng"])

        info = detect_header_and_columns(ws)
        self.assertEqual(info["header_start_row"], 1)
        self.assertEqual(info["header_end_row"], 1)
        self.assertEqual(info["header_row_idx"], 1)
        self.assertEqual(info["raw_headers"], ["STT", "Họ và tên", "Ngày sinh", "Giới tính", "Nơi sinh"])
        self.assertEqual(info["stt_col_idx"], 0)
        self.assertEqual(info["name_col_idx"], 1)

    def test_single_row_header_with_banner_above(self):
        """Kiểm tra file có banner ở các dòng 1-3 và dòng tiêu đề thực tế ở dòng 4."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DanhSach"

        # Dòng 1-3: Banner / Tiêu đề danh sách
        ws.append(["UBND THÀNH PHỐ HÀ NỘI"])
        ws.append(["TRƯỜNG THCS NGUYỄN DU"])
        ws.append(["DANH SÁCH HỌC SINH TỐT NGHIỆP"])

        # Dòng 4: Tiêu đề chuẩn 1 dòng
        ws.append(["STT", "Họ và tên", "Ngày sinh", "Giới tính", "Toán", "Văn", "Ngoại ngữ"])
        # Dòng 5 trở đi là dữ liệu học sinh
        ws.append([1, "Lê Hoàng Long", "15/06/2008", "Nam", 8.5, 9.0, 8.0])
        ws.append([2, "Phạm Ngọc Ánh", "02/11/2008", "Nữ", 9.0, 9.5, 9.0])

        info = detect_header_and_columns(ws)
        self.assertEqual(info["header_start_row"], 4)
        self.assertEqual(info["header_end_row"], 4)
        self.assertEqual(info["header_row_idx"], 4)

        expected_headers = ["STT", "Họ và tên", "Ngày sinh", "Giới tính", "Toán", "Văn", "Ngoại ngữ"]
        self.assertEqual(info["raw_headers"], expected_headers)
        self.assertEqual(info["stt_col_idx"], 0)
        self.assertEqual(info["name_col_idx"], 1)

    def test_single_row_header_with_banner_and_empty_rows(self):
        """Kiểm tra nhận diện chính xác khi có nhiều dòng chữ và dòng trống xen kẽ ở phía trên."""
        wb = openpyxl.Workbook()
        ws = wb.active

        # Dòng 1-4: Các dòng tiêu đề cơ quan / văn bản
        ws.append(["UBND TỈNH VĨNH PHÚC"])
        ws.append(["PHÒNG GIÁO DỤC VÀ ĐÀO TẠO LẬP THẠCH"])
        ws.append(["TRƯỜNG THCS XUÂN HÒA"])
        ws.append(["DANH SÁCH HỌC SINH TỐT NGHIỆP THCS"])

        # Dòng 5-7: Các dòng trống xen kẽ
        ws.append([])
        ws.append([None, None, None])
        ws.append([])

        # Dòng 8: Tiêu đề cột thực tế
        ws.append(["STT", "Họ và tên", "Ngày sinh", "Giới tính", "Nơi sinh", "Xếp loại"])

        # Dòng 9-10: Dữ liệu học sinh
        ws.append([1, "Nguyễn Văn An", "15/05/2008", "Nam", "Vĩnh Phúc", "Giỏi"])
        ws.append([2, "Trần Thị Bình", "20/10/2008", "Nữ", "Vĩnh Phúc", "Khá"])

        info = detect_header_and_columns(ws)
        self.assertEqual(info["header_row_idx"], 8)
        self.assertEqual(info["raw_headers"], ["STT", "Họ và tên", "Ngày sinh", "Giới tính", "Nơi sinh", "Xếp loại"])
        self.assertEqual(info["stt_col_idx"], 0)
        self.assertEqual(info["name_col_idx"], 1)

    def test_auto_suggest_mapping_for_split_columns(self):
        """Kiểm tra auto_suggest_mapping tự động gán Ngày+Tháng+Năm và Họ+Tên từ các cột đơn lẻ."""
        headers = [
            "STT",
            "Họ đệm",
            "Tên",
            "Ngày",
            "Tháng",
            "Năm",
            "Giới tính",
            "Nơi sinh"
        ]
        mapping = auto_suggest_mapping(headers)

        # Cột STT (Target 1) -> 0
        self.assertEqual(mapping[1], 0)
        # Cột Họ tên (Target 5) -> gộp [1, 2]
        self.assertEqual(mapping[5], [1, 2])
        # Cột Ngày sinh (Target 9) -> gộp [3, 4, 5]
        self.assertEqual(mapping[9], [3, 4, 5])
        # Cột Giới tính (Target 10) -> 6
        self.assertEqual(mapping[10], 6)
        # Cột Nơi sinh (Target 13) -> 7
        self.assertEqual(mapping[13], 7)

    def test_smart_merger_with_single_row_header_file(self):
        """Kiểm tra SmartMerger đọc và chuẩn hóa thành công tệp có tiêu đề 1 dòng."""
        src_file = self.temp_dir / "test_single_row_headers.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "K70"

        ws.append(["BẢNG GHI TÊN GHI ĐIỂM"])
        ws.append(["STT", "Họ và tên", "Ngày sinh", "Giới tính", "Nơi sinh"])
        ws.append([1, "Vũ Đức Đam", "03/02/1963", "Nam", "Hải Dương"])
        ws.append([2, "Nguyễn Thị Kim Ngân", "12/04/1954", "Nữ", "Bến Tre"])
        wb.save(src_file)

        out_file = self.temp_dir / "output_standardized.xlsx"
        from src.smart_merger import MappingProfileManager
        prof_mgr = MappingProfileManager(self.temp_dir / "test_profiles.json")
        merger = SmartMerger(profile_mgr=prof_mgr)
        res = merger.process_merge_and_standardize(
            source_files=[src_file],
            output_file_path=out_file,
            mapping_callback=lambda info: (info["suggested_mapping"], "TestProfile", {}, False)
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["total_rows"], 2)

        # Kiểm tra nội dung file đầu ra
        wb_out = openpyxl.load_workbook(out_file)
        ws_out = wb_out.active
        self.assertEqual(ws_out.cell(row=3, column=1).value, 1)
        self.assertEqual(ws_out.cell(row=3, column=5).value, "Vũ Đức Đam")
        self.assertEqual(ws_out.cell(row=3, column=9).value, "03/02/1963")
        self.assertEqual(ws_out.cell(row=3, column=10).value, "Nam")

        self.assertEqual(ws_out.cell(row=4, column=1).value, 2)
        self.assertEqual(ws_out.cell(row=4, column=5).value, "Nguyễn Thị Kim Ngân")
        self.assertEqual(ws_out.cell(row=4, column=9).value, "12/04/1954")
        self.assertEqual(ws_out.cell(row=4, column=10).value, "Nữ")
        wb_out.close()


if __name__ == "__main__":
    unittest.main()
