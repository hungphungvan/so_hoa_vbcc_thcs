import sys
import unittest
import tempfile
from pathlib import Path
import openpyxl

if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

from src.lookup_manager import LookupManager, find_data_file
from src.excel_processor import ExcelProcessor

TEMP_DIR = Path(tempfile.gettempdir())

class TestGlobalPipeline(unittest.TestCase):
    def setUp(self):
        self.lookup_mgr = LookupManager()
        self.processor = ExcelProcessor(self.lookup_mgr)

    def test_pipeline_stages_order(self):
        # Tạo file Excel giả lập gồm 2 sheets với lỗi Ngày sinh, Tên trường chưa chuẩn, và Niên khóa
        test_file = TEMP_DIR / "test_pipeline_multi_sheets.xlsx"
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

    def test_cancellation_stops_immediately(self):
        """Kiểm tra khi check_cancel=True thì chương trình dừng ngay, không tiếp tục các giai đoạn sau."""
        test_file = TEMP_DIR / "test_cancel.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Sheet1"
        header_row = ["STT", "Số hiệu", "Số vào sổ", "Tên văn bằng", "Họ và tên", "Mã", "CCCD", "Hộ chiếu",
                      "Ngày sinh", "Giới tính", "Dân tộc", "Quốc tịch", "Nơi sinh", "Tên trường", "Niên khóa",
                      "Xếp loại", "Hình thức", "Năm TN", "Ngày cấp", "Số QĐ", "Ngày QĐ", "Người ký", "CQ cấp",
                      "Ghi chú", "Trạng thái", "Lý do", "Mã ĐV", "Tên ĐV"]
        ws.append(["Tiêu đề"])
        ws.append(header_row)
        ws.append([1, "", "", "", "Lê Văn Hủy", "", "", "", "31/02/2005"] + [""] * 19)
        wb.save(test_file)
        wb.close()

        is_cancelled = False
        dialog_called = []
        def mock_dialog(d_type, *args):
            nonlocal is_cancelled
            dialog_called.append(d_type)
            is_cancelled = True # Người dùng bấm dừng tại dialog đầu tiên
            return "15/05/2005", False

        unit_info = {"code": "TEST", "name": "Đơn vị Test", "schools": ["Trường A"]}
        all_st, stats = self.processor.process_all_files(
            files=[test_file],
            unit_info=unit_info,
            dialog_callback=mock_dialog,
            check_cancel=lambda: is_cancelled
        )
        # Vì bị hủy ngay sau Giai đoạn 1, không được chạy tới Giai đoạn 5 xuất file
        self.assertEqual(len(dialog_called), 1)
        self.assertNotIn("school_block", dialog_called)
        self.assertNotIn("cohort_block", dialog_called)
        print("--> Test Cancellation PASSED successfully!")

    def test_empty_school_and_cohort_caching(self):
        """Kiểm tra tính năng tối ưu: nhớ nhóm trường trống cùng tệp và nhớ niên khóa trống an toàn."""
        test_file = TEMP_DIR / "test_empty_cache.xlsx"
        wb = openpyxl.Workbook()
        # Sheet 1: trường trống, niên khóa trống, năm sinh 2005
        ws1 = wb.active
        ws1.title = "Sheet 1"
        ws1.append(["Tiêu đề"])
        header_row = ["STT", "Số hiệu", "Số vào sổ", "Tên văn bằng", "Họ và tên", "Mã", "CCCD", "Hộ chiếu",
                      "Ngày sinh", "Giới tính", "Dân tộc", "Quốc tịch", "Nơi sinh", "Tên trường", "Niên khóa",
                      "Xếp loại", "Hình thức", "Năm TN", "Ngày cấp", "Số QĐ", "Ngày QĐ", "Người ký", "CQ cấp",
                      "Ghi chú", "Trạng thái", "Lý do", "Mã ĐV", "Tên ĐV"]
        ws1.append(header_row)
        ws1.append([1, "", "", "", "Học Sinh 1", "", "", "", "10/05/2005", "", "", "", "", "", ""] + [""] * 13)

        # Sheet 2: trường trống, niên khóa trống, năm sinh 2005 (cùng file -> phải tự nhớ trường và niên khóa)
        ws2 = wb.create_sheet(title="Sheet 2")
        ws2.append(["Tiêu đề"])
        ws2.append(header_row)
        ws2.append([1, "", "", "", "Học Sinh 2", "", "", "", "12/08/2005", "", "", "", "", "", ""] + [""] * 13)
        wb.save(test_file)
        wb.close()

        dialog_counts = {"school_block": 0, "cohort_block": 0}
        def mock_dialog(d_type, *args):
            dialog_counts[d_type] += 1
            if d_type == "school_block":
                return "Trường Tự Nhớ", True # Chọn ghi nhớ
            elif d_type == "cohort_block":
                return "2016-2020", True # Chọn ghi nhớ
            return None, False

        unit_info = {"code": "TEST", "name": "Đơn vị Test", "schools": ["Trường Tự Nhớ"]}
        all_st, stats = self.processor.process_all_files(
            files=[test_file],
            unit_info=unit_info,
            dialog_callback=mock_dialog
        )

        # Vì có 2 sheets cùng bị trống trường & trống niên khóa nhưng cùng năm sinh 2005:
        # Nhờ tính năng ghi nhớ tối ưu, mỗi dialog chỉ cần hỏi 1 lần duy nhất!
        self.assertEqual(dialog_counts["school_block"], 1, "Chỉ được hỏi trường trống 1 lần cho cùng tệp!")
        self.assertEqual(dialog_counts["cohort_block"], 1, "Chỉ được hỏi niên khóa trống 1 lần khi năm sinh đa số khớp!")
        self.assertEqual(all_st[0]["school"], "Trường Tự Nhớ")
        self.assertEqual(all_st[1]["school"], "Trường Tự Nhớ")
        self.assertEqual(all_st[0]["cohort"], "2016-2020")
        self.assertEqual(all_st[1]["cohort"], "2016-2020")
        print("--> Test Empty Cache Optimization PASSED successfully!")

    def test_export_options_and_audit_log(self):
        """Kiểm tra:
        1. Bỏ hoàn toàn file Du_Lieu_Goc.
        2. Tùy chọn ds_loi: chỉ xuất DS_Loi khi được chọn.
        3. Tùy chọn nhat_ky: xuất Nhat_Ky_Thay_Doi.xlsx ghi nhận đầy đủ thay đổi chuẩn hóa.
        """
        test_file = TEMP_DIR / "test_audit_log.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Lop 9"
        header_row = ["STT", "Số hiệu", "Số vào sổ", "Tên văn bằng", "Họ và tên", "Mã", "CCCD", "Hộ chiếu",
                      "Ngày sinh", "Giới tính", "Dân tộc", "Quốc tịch", "Nơi sinh", "Tên trường", "Niên khóa",
                      "Xếp loại", "Hình thức", "Năm TN", "Ngày cấp", "Số QĐ", "Ngày QĐ", "Người ký", "CQ cấp",
                      "Ghi chú", "Trạng thái", "Lý do", "Mã ĐV", "Tên ĐV"]
        ws.append(["Tiêu đề"])
        ws.append(header_row)
        # Học sinh 1: Ngày sinh 15/05/2005, trường hợp lệ
        ws.append([1, "", "", "", "Nguyễn Văn Một", "", "", "", "15/05/2005", "Nam", "Kinh", "Việt Nam", "Hà Nội", "THCS Test", "2016-2020", "Giỏi", "Chính quy", 2020, "15/06/2020"] + [""] * 9)
        # Học sinh 2: Bị thiếu ngày sinh, thiếu tên trường -> Gây lỗi
        ws.append([2, "", "", "", "Trần Thị Hai", "", "", "", "", "Nữ", "Kinh", "Việt Nam", "Hà Nội", "", "", "Khá", "Chính quy", 2020, ""] + [""] * 9)
        wb.save(test_file)
        wb.close()

        out_dir = TEMP_DIR / "test_out_audit"
        if out_dir.exists():
            for f in out_dir.glob("*.xlsx"):
                f.unlink()

        unit_info = {"code": "TEST", "name": "Phòng GD Test", "schools": ["THCS Test"]}

        # 1. Chạy với ds_loi=True, nhat_ky=True
        all_st, stats = self.processor.process_all_files(
            files=[test_file],
            unit_info=unit_info,
            output_dir=out_dir,
            export_options={"ds_loi": True, "nhat_ky": True}
        )

        # File Du_Lieu_Goc KHÔNG được phép tồn tại
        raw_file = out_dir / "Phòng GD Test_Du_Lieu_Goc.xlsx"
        self.assertFalse(raw_file.exists(), "Không được sinh ra file Du_Lieu_Goc.xlsx!")

        # File DS_Loi ĐƯỢC sinh ra vì có lỗi và ds_loi=True
        err_file = out_dir / "Phòng GD Test_DS_Loi.xlsx"
        self.assertTrue(err_file.exists(), "Phải sinh ra file DS_Loi.xlsx khi có học sinh lỗi!")

        # File Nhat_Ky_Thay_Doi ĐƯỢC sinh ra và kiểm tra nội dung
        log_file = out_dir / "Phòng GD Test_Nhat_Ky_Thay_Doi.xlsx"
        self.assertTrue(log_file.exists(), "Phải sinh ra file Nhat_Ky_Thay_Doi.xlsx!")
        log_wb = openpyxl.load_workbook(log_file)
        self.assertIn("Nhat_Ky_Thay_Doi", log_wb.sheetnames)
        log_ws = log_wb["Nhat_Ky_Thay_Doi"]
        self.assertGreaterEqual(log_ws.max_row, 3, "Nhật ký thay đổi phải có ít nhất 1 dòng ghi nhận!")
        log_wb.close()

        # 2. Xóa các file và chạy lại với ds_loi=False, nhat_ky=False
        for f in out_dir.glob("*.xlsx"):
            f.unlink()

        all_st2, stats2 = self.processor.process_all_files(
            files=[test_file],
            unit_info=unit_info,
            output_dir=out_dir,
            export_options={"ds_loi": False, "nhat_ky": False}
        )
        self.assertFalse((out_dir / "Phòng GD Test_Du_Lieu_Goc.xlsx").exists())
        self.assertFalse((out_dir / "Phòng GD Test_DS_Loi.xlsx").exists(), "Khi ds_loi=False, không được xuất file lỗi!")
        self.assertFalse((out_dir / "Phòng GD Test_Nhat_Ky_Thay_Doi.xlsx").exists(), "Khi nhat_ky=False, không được xuất file nhật ký!")
        print("--> Test Export Options & Audit Log PASSED successfully!")

    def test_tcvn3_auto_conversion(self):
        """Kiểm tra tính năng tự động nhận diện và chuyển đổi bảng mã TCVN3 sang Unicode."""
        from src.normalizer import auto_convert_tcvn3, is_tcvn3, tcvn3_to_unicode, clean_whitespace

        # 1. Kiểm tra nhận diện & chuyển đổi chuỗi đơn lẻ
        tcvn3_samples = {
            "NguyÔn V¨n An": "Nguyễn Văn An",
            "Ph¹m ThÞ Mai": "Phạm Thị Mai",
            "TrÇn V¨n B×nh": "Trần Văn Bình",
            "§Æng ThÞ Hång": "Đặng Thị Hồng",
            "Hoµng V¨n Th¸i": "Hoàng Văn Thái",
            "Hä vµ tªn": "Họ và tên",
            "Ngµy sinh": "Ngày sinh",
            "N¬i sinh": "Nơi sinh",
            "Tªn tr\xadêng": "Tên trường"
        }

        for tcvn_str, expected_uni in tcvn3_samples.items():
            self.assertTrue(is_tcvn3(tcvn_str), f"Chuỗi '{tcvn_str}' phải được nhận diện là TCVN3!")
            self.assertEqual(auto_convert_tcvn3(tcvn_str), expected_uni)

        # Chuỗi Unicode chuẩn không được bị biến dạng
        unicode_samples = [
            "Nguyễn Văn An",
            "Phạm Thị Mai",
            "Trần Văn Bình",
            "Đặng Thị Hồng",
            "Bác Hồ",
            "Chú ý",
            "Hoàng Văn Thái"
        ]
        for uni_str in unicode_samples:
            self.assertFalse(is_tcvn3(uni_str), f"Chuỗi Unicode '{uni_str}' không được nhận diện nhầm thành TCVN3!")
            self.assertEqual(auto_convert_tcvn3(uni_str), uni_str)

        # 2. Kiểm tra đọc từ file Excel chứa dữ liệu TCVN3
        tcvn_file = TEMP_DIR / "test_tcvn3_file.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "LÝp 9A"  # Sheet name TCVN3 cho 'Lớp 9A'
        
        # Header TCVN3
        header_row = ["STT", "Sè hiÖu", "Sè vµo sæ", "Tªn v¨n b»ng", "Hä vµ tªn", "M·", "CCCD", "Hé chiÕu",
                      "Ngµy sinh", "Giíi tÝnh", "D©n téc", "Quèc tÞch", "N¬i sinh", "Tªn tr\xadêng", "Niªn khãa",
                      "XÕp lo¹i", "H×nh thøc", "N¨m TN", "Ngµy cÊp", "Sè Q§", "Ngµy Q§", "Ngêi ký", "CQ cÊp",
                      "Ghi chó", "Tr¹ng th¸i", "Lý do", "M· §V", "Tªn §V"]
        ws.append(["Danh s¸ch cÊp b»ng"])
        ws.append(header_row)

        # 2 học sinh với dữ liệu TCVN3
        row1 = [1, "", "", "", "NguyÔn V¨n An", "", "", "", "15/05/2005", "Nam", "Kinh", "ViÖt Nam", "Hµ Néi", "THCS Lª Lîi", "2016-2020", "", "", "", "15/06/2020"] + [""] * 9
        row2 = [2, "", "", "", "TrÇn ThÞ BÝch", "", "", "", "20/10/2005", "N÷", "Kinh", "ViÖt Nam", "Hµ Néi", "THCS Lª Lîi", "2016-2020", "", "", "", "15/06/2020"] + [""] * 9
        ws.append(row1)
        ws.append(row2)
        wb.save(tcvn_file)
        wb.close()

        students = self.processor.extract_students_from_file(
            file_path=tcvn_file,
            allowed_schools=["THCS Lê Lợi"]
        )

        self.assertEqual(len(students), 2)
        self.assertEqual(students[0]["name"], "Nguyễn Văn An")
        self.assertEqual(students[0]["school"], "THCS Lê Lợi")
        self.assertEqual(students[1]["name"], "Trần Thị Bích")
        self.assertEqual(students[1]["school"], "THCS Lê Lợi")
        print("--> Test TCVN3 Auto Conversion PASSED successfully!")

    def test_vni_auto_conversion(self):
        """Kiểm tra tính năng tự động nhận diện và chuyển đổi bảng mã VNI-Windows sang Unicode."""
        from src.normalizer import auto_convert_legacy_encoding, is_vni, vni_to_unicode

        vni_samples = {
            "Nguyeãn Vaên An": "Nguyễn Văn An",
            "Traàn Thò Bích": "Trần Thị Bích",
            "Phaïm Vaên Ñaït": "Phạm Văn Đạt",
            "Tröôøng THCS Leâ Lôïi": "Trường THCS Lê Lợi",
            "Haø Noäi": "Hà Nội",
            "Ñaëng Thò Hoàng": "Đặng Thị Hồng",
            "Löông Thò Aùnh": "Lương Thị Ánh",
            "NGUYEÃN VAÊN AN": "NGUYỄN VĂN AN",
            "TRAÀN THÒ BÍCH": "TRẦN THỊ BÍCH",
            "Toå quoác": "Tổ quốc"
        }

        for vni_str, expected_uni in vni_samples.items():
            self.assertTrue(is_vni(vni_str), f"Chuỗi '{vni_str}' phải được nhận diện là VNI!")
            self.assertEqual(auto_convert_legacy_encoding(vni_str), expected_uni)

        # Kiểm tra đọc từ file Excel chứa dữ liệu VNI
        vni_file = TEMP_DIR / "test_vni_file.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Lôùp 9A"  # VNI cho 'Lớp 9A'

        header_row = ["STT", "Soá hieäu", "Soá vaøo soå", "Teân vaên baèng", "Hoï vaø teân", "Maõ", "CCCD", "Hoä chieáu",
                      "Ngaøy sinh", "Giôùi tính", "Daân toäc", "Quoác tòch", "Nôi sinh", "Teân tröôøng", "Nieân khoùa",
                      "Xeáp loaïi", "Hình thöùc", "Naêm TN", "Ngaøy caáp", "Soá QÑ", "Ngaøy QÑ", "Người ký", "CQ caáp",
                      "Ghi chuù", "Traïng thaùi", "Lyù do", "Maõ ÑV", "Teân ÑV"]
        ws.append(["Danh saùch caáp baèng"])
        ws.append(header_row)

        row1 = [1, "", "", "", "Nguyeãn Vaên An", "", "", "", "15/05/2005", "Nam", "Kinh", "Vieät Nam", "Haø Noäi", "THCS Leâ Lôïi", "2016-2020", "", "", "", "15/06/2020"] + [""] * 9
        row2 = [2, "", "", "", "Traàn Thò Bích", "", "", "", "20/10/2005", "Nöõ", "Kinh", "Vieät Nam", "Haø Noäi", "THCS Leâ Lôïi", "2016-2020", "", "", "", "15/06/2020"] + [""] * 9
        ws.append(row1)
        ws.append(row2)
        wb.save(vni_file)
        wb.close()

        students = self.processor.extract_students_from_file(
            file_path=vni_file,
            allowed_schools=["THCS Lê Lợi"]
        )

        self.assertEqual(len(students), 2)
        self.assertEqual(students[0]["name"], "Nguyễn Văn An")
        self.assertEqual(students[0]["school"], "THCS Lê Lợi")
        self.assertEqual(students[1]["name"], "Trần Thị Bích")
        self.assertEqual(students[1]["school"], "THCS Lê Lợi")
        print("--> Test VNI Auto Conversion PASSED successfully!")

    def test_quote_and_date_format_normalization(self):
        """Kiểm tra chuẩn hóa ngày sinh khi có dấu nháy đơn '04/07/1994, định dạng Date hoặc số serial."""
        import datetime
        from src.normalizer import format_date_to_text, check_dob_validity, format_grad_date_to_text

        cases = [
            ("04/07/1994", "04/07/1994"),
            ("'04/07/1994", "04/07/1994"),
            ("'04/07/1994'", "04/07/1994"),
            ("’04/07/1994", "04/07/1994"),
            ("'4/7/1994", "04/07/1994"),
            ("'1994-07-04", "04/07/1994"),
            (datetime.datetime(1994, 7, 4), "04/07/1994"),
            (datetime.date(1994, 7, 4), "04/07/1994"),
            (34519, "04/07/1994"),
        ]

        for raw_val, expected in cases:
            res = format_date_to_text(raw_val)
            self.assertEqual(res, expected)
            is_valid, err_code, err_desc = check_dob_validity(res)
            self.assertTrue(is_valid, f"Thất bại với {raw_val}: {err_desc}")

            # Kiểm tra cả hàm format_grad_date_to_text
            grad_res = format_grad_date_to_text(raw_val)
            self.assertEqual(grad_res, expected)

        print("--> Test Quote Prefix & Date Format PASSED successfully!")

if __name__ == "__main__":
    unittest.main()
