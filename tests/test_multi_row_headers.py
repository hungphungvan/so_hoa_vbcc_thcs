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
    """Bộ kiểm thử cho tính năng nhận diện dòng tiêu đề đa tầng và các ô merge (Merged Cells)."""

    def setUp(self):
        self.temp_dir = Path("tests/temp_test_headers")
        self.temp_dir.mkdir(parents=True, exist_ok=True)

    def tearDown(self):
        import shutil
        if self.temp_dir.exists():
            shutil.rmtree(self.temp_dir, ignore_errors=True)

    def test_single_row_header_backward_compatibility(self):
        """Đảm bảo tiêu đề 1 dòng truyền thống vẫn hoạt động chính xác 100%."""
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

    def test_two_row_merged_header_detection(self):
        """Kiểm tra file có tiêu đề 2 dòng merge (A4:A5, B4:B5, C4:E4 với Ngày/Tháng/Năm, F4:F5)."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "DanhSach"

        # Dòng 1-3: Banner / Tiêu đề danh sách
        ws.append(["UBND THÀNH PHỐ HÀ NỘI"])
        ws.append(["TRƯỜNG THCS NGUYỄN DU"])
        ws.append(["DANH SÁCH HỌC SINH TỐT NGHIỆP"])

        # Dòng 4: Tiêu đề tầng 1
        ws.append(["STT", "Họ và tên", "Ngày tháng năm sinh", None, None, "Giới tính", "Điểm thi", None, None])
        # Dòng 5: Tiêu đề tầng 2 (sub-headers)
        ws.append([None, None, "Ngày", "Tháng", "Năm", None, "Toán", "Văn", "Ngoại ngữ"])

        # Thiết lập merged_cells trong openpyxl
        ws.merge_cells("A4:A5")  # STT
        ws.merge_cells("B4:B5")  # Họ và tên
        ws.merge_cells("C4:E4")  # Ngày tháng năm sinh (ngang 3 cột)
        ws.merge_cells("F4:F5")  # Giới tính
        ws.merge_cells("G4:I4")  # Điểm thi (ngang 3 cột)

        # Dữ liệu học sinh bắt đầu từ dòng 6
        ws.append([1, "Lê Hoàng Long", 15, 6, 2008, "Nam", 8.5, 9.0, 8.0])
        ws.append([2, "Phạm Ngọc Ánh", 2, 11, 2008, "Nữ", 9.0, 9.5, 9.0])

        info = detect_header_and_columns(ws)
        self.assertEqual(info["header_start_row"], 4)
        self.assertEqual(info["header_end_row"], 5)
        self.assertEqual(info["header_row_idx"], 5)

        expected_headers = [
            "STT",
            "Họ và tên",
            "Ngày tháng năm sinh - Ngày",
            "Ngày tháng năm sinh - Tháng",
            "Ngày tháng năm sinh - Năm",
            "Giới tính",
            "Điểm thi - Toán",
            "Điểm thi - Văn",
            "Điểm thi - Ngoại ngữ"
        ]
        self.assertEqual(info["raw_headers"], expected_headers)
        self.assertEqual(info["stt_col_idx"], 0)
        self.assertEqual(info["name_col_idx"], 1)

    def test_three_row_header_with_index_row(self):
        """Kiểm tra tiêu đề có thêm hàng đánh số thứ tự cột (1), (2), (3)..."""
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(["STT", "Họ và tên", "Ngày sinh", "Địa chỉ"])
        ws.append([None, "Họ và chữ đệm", "Tên", None, None])  # sub-header
        ws.append(["(1)", "(2)", "(3)", "(4)", "(5)"])  # Hàng đánh số cột
        ws.append([1, "Nguyễn Văn", "Bình", "01/01/2008", "Hà Nội"])

        info = detect_header_and_columns(ws)
        self.assertEqual(info["header_start_row"], 1)
        self.assertEqual(info["header_end_row"], 3)
        self.assertEqual(info["header_row_idx"], 3)

        # Hàng (1), (2)... không được thêm vào chuỗi tên cột
        for h in info["raw_headers"]:
            self.assertFalse("(1)" in h or "(2)" in h or "(3)" in h)

    def test_auto_suggest_mapping_for_split_columns(self):
        """Kiểm tra auto_suggest_mapping tự động gán Ngày+Tháng+Năm và Họ+Tên."""
        headers = [
            "STT",
            "Họ và tên - Họ và chữ đệm",
            "Họ và tên - Tên",
            "Ngày tháng năm sinh - Ngày",
            "Ngày tháng năm sinh - Tháng",
            "Ngày tháng năm sinh - Năm",
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

    def test_smart_merger_with_two_row_merged_file(self):
        """Kiểm tra SmartMerger đọc và chuẩn hóa thành công tệp có tiêu đề 2 dòng merge."""
        src_file = self.temp_dir / "test_merged_headers.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "K70"

        ws.append(["BẢNG GHI TÊN GHI ĐIỂM"])
        ws.append(["STT", "Họ và tên", "Ngày tháng năm sinh", None, None, "Giới tính", "Nơi sinh"])
        ws.append([None, None, "Ngày", "Tháng", "Năm", None, None])

        ws.merge_cells("A2:A3")
        ws.merge_cells("B2:B3")
        ws.merge_cells("C2:E2")
        ws.merge_cells("F2:F3")
        ws.merge_cells("G2:G3")

        ws.append([1, "Vũ Đức Đam", 3, 2, 1963, "Nam", "Hải Dương"])
        ws.append([2, "Nguyễn Thị Kim Ngân", 12, 4, 1954, "Nữ", "Bến Tre"])
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
        # Dòng 1: Tiêu đề banner, Dòng 2: Tiêu đề 28 cột chuẩn, Dòng 3-4: Dữ liệu
        self.assertEqual(ws_out.cell(row=3, column=1).value, 1)
        self.assertEqual(ws_out.cell(row=3, column=5).value, "Vũ Đức Đam")
        self.assertEqual(ws_out.cell(row=3, column=9).value, "03/02/1963")
        self.assertEqual(ws_out.cell(row=3, column=10).value, "Nam")

        self.assertEqual(ws_out.cell(row=4, column=1).value, 2)
        self.assertEqual(ws_out.cell(row=4, column=5).value, "Nguyễn Thị Kim Ngân")
        self.assertEqual(ws_out.cell(row=4, column=9).value, "12/04/1954")
        self.assertEqual(ws_out.cell(row=4, column=10).value, "Nữ")
        wb_out.close()

    def test_excel_merger_same_structure_multi_row_header(self):
        """Kiểm tra ExcelMerger (Gộp cùng cấu trúc) bảo toàn khối tiêu đề 2 dòng và gộp dữ liệu."""
        f1 = self.temp_dir / "file1.xlsx"
        f2 = self.temp_dir / "file2.xlsx"

        for f_path, start_stt, name in [(f1, 1, "Trần Văn A"), (f2, 2, "Lê Thị B")]:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Data"
            # Header 2 dòng
            ws.append(["STT", "Họ và tên", "Ngày sinh", "Điểm thi", None])
            ws.append([None, None, None, "Toán", "Văn"])
            ws.merge_cells("A1:A2")
            ws.merge_cells("B1:B2")
            ws.merge_cells("C1:C2")
            ws.merge_cells("D1:E1")
            ws.append([start_stt, name, "10/10/2008", 8.0, 9.0])
            wb.save(f_path)

        out_merged = self.temp_dir / "gop_cung_cau_truc.xlsx"
        merger = ExcelMerger()
        res = merger.merge_files(
            file_paths=[f1, f2],
            output_file_path=out_merged,
            renumber_stt=True
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["total_rows"], 2)

        wb_chk = openpyxl.load_workbook(out_merged)
        ws_chk = wb_chk.active
        # Khối header phải có dòng 1 và dòng 2 nguyên vẹn
        self.assertEqual(ws_chk.cell(row=1, column=1).value, "STT")
        self.assertEqual(ws_chk.cell(row=1, column=4).value, "Điểm thi")
        self.assertEqual(ws_chk.cell(row=2, column=4).value, "Toán")
        self.assertEqual(ws_chk.cell(row=2, column=5).value, "Văn")

        # Dữ liệu học sinh bắt đầu từ dòng 3 (vì header gồm dòng 1 và 2)
        self.assertEqual(ws_chk.cell(row=3, column=1).value, 1)
        self.assertEqual(ws_chk.cell(row=3, column=2).value, "Trần Văn A")
        self.assertEqual(ws_chk.cell(row=4, column=1).value, 2)
        self.assertEqual(ws_chk.cell(row=4, column=2).value, "Lê Thị B")
        wb_chk.close()


if __name__ == "__main__":
    unittest.main()
