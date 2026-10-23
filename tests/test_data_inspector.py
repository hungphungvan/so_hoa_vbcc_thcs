import unittest
from pathlib import Path
import tempfile
import openpyxl

from src.data_inspector import (
    normalize_cohort_str,
    parse_and_validate_date,
    validate_student_record,
    DataInspector
)


class TestDataInspector(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp())

    def test_normalize_cohort_str(self):
        self.assertEqual(normalize_cohort_str("2019 - 2023"), "2019-2023")
        self.assertEqual(normalize_cohort_str("2019–2023"), "2019-2023")  # en-dash
        self.assertEqual(normalize_cohort_str("2019—2023"), "2019-2023")  # em-dash
        self.assertEqual(normalize_cohort_str(None), "")

    def test_parse_and_validate_date(self):
        # Hợp lệ chuẩn
        f_val, st, yr = parse_and_validate_date("15/05/2008")
        self.assertEqual(f_val, "15/05/2008")
        self.assertIsNone(st)
        self.assertEqual(yr, 2008)

        # Hỗ trợ dấu nháy đơn phía trước hoặc ở giữa chuỗi ngày ('19/05/1969 hay 19/'05/1969)
        f_val, st, yr = parse_and_validate_date("'19/05/1969")
        self.assertEqual(f_val, "19/05/1969")
        self.assertIsNone(st)
        self.assertEqual(yr, 1969)

        f_val, st, yr = parse_and_validate_date("19/'05/1969")
        self.assertEqual(f_val, "19/05/1969")
        self.assertIsNone(st)
        self.assertEqual(yr, 1969)

        f_val, st, yr = parse_and_validate_date("19/05/'1969")
        self.assertEqual(f_val, "19/05/1969")
        self.assertIsNone(st)
        self.assertEqual(yr, 1969)

        # Trống
        f_val, st, yr = parse_and_validate_date("")
        self.assertEqual(st, "TRỐNG")

        # Sai định dạng (chỉ có năm hoặc chữ)
        f_val, st, yr = parse_and_validate_date("2008")
        self.assertEqual(st, "LỖI")

        # Sai logic ngày tháng (30/02/2011)
        f_val, st, yr = parse_and_validate_date("30/02/2011")
        self.assertEqual(f_val, "30/02/2011")
        self.assertEqual(st, "CẢNH BÁO")
        self.assertEqual(yr, 2011)

    def test_validate_student_record_rules(self):
        # 1. Học sinh hợp lệ hoàn hảo
        row = [""] * 28
        row[1] = "  B123456  "         # Cột 2: Số hiệu bằng -> làm sạch khoảng trắng
        row[2] = "  045/2023  "        # Cột 3: Số vào sổ -> làm sạch khoảng trắng
        row[3] = ""                    # Cột 4: Tên văn bằng -> Luôn tự động điền 'Bằng tốt nghiệp trung học cơ sở'
        row[4] = "nguyễn văn an"
        row[5] = "  0123456789  "      # Cột 6: Mã người học -> làm sạch khoảng trắng
        row[6] = "  001 208 009 999  " # Cột 7: Số định danh CCCD -> làm sạch khoảng trắng
        row[8] = "19/'05/2008"         # Cột 9: Ngày sinh có dấu nháy ở giữa
        row[14] = "2019-2023"          # Y_ra = 2023 = 2008 + 15
        row[18] = "'15/06/2023"        # Y_cap = 2023 = Y_ra (có dấu nháy ở đầu)

        norm, col_st, summary = validate_student_record(row)
        self.assertEqual(norm[1], "B123456")
        self.assertEqual(norm[2], "045/2023")
        self.assertEqual(norm[3], "Bằng tốt nghiệp trung học cơ sở")
        self.assertEqual(norm[4], "Nguyễn Văn An")
        self.assertEqual(norm[5], "0123456789")
        self.assertEqual(norm[6], "001208009999")
        self.assertEqual(norm[8], "19/05/2008")
        self.assertEqual(norm[18], "15/06/2023")
        self.assertIsNone(col_st["ho_ten"])
        self.assertIsNone(col_st["ngay_sinh"])
        self.assertIsNone(col_st["nien_khoa"])
        self.assertIsNone(col_st["ngay_cap"])
        self.assertEqual(summary, "")

        # 2. Họ và tên bị trống -> Lỗi
        row_err_name = list(row)
        row_err_name[4] = ""
        norm, col_st, summary = validate_student_record(row_err_name)
        self.assertEqual(col_st["ho_ten"], "LỖI")
        self.assertEqual(summary, "LỖI")

        # 3. Ngày sinh trống -> Lỗi
        row_err_dob = list(row)
        row_err_dob[8] = ""
        norm, col_st, summary = validate_student_record(row_err_dob)
        self.assertEqual(col_st["ngay_sinh"], "LỖI")
        self.assertEqual(summary, "LỖI")

        # 4. Ngày sinh không đúng dạng dd/MM/yyyy -> Lỗi
        row_err_dob2 = list(row)
        row_err_dob2[8] = "2008"
        norm, col_st, summary = validate_student_record(row_err_dob2)
        self.assertEqual(col_st["ngay_sinh"], "LỖI")
        self.assertEqual(summary, "LỖI")

        # 5. Ngày sinh sai logic (30/02/2011) -> Cảnh báo
        row_warn_dob = list(row)
        row_warn_dob[8] = "30/02/2008"
        norm, col_st, summary = validate_student_record(row_warn_dob)
        self.assertEqual(col_st["ngay_sinh"], "CẢNH BÁO")
        self.assertEqual(summary, "CẢNH BÁO")

        # 6. Niên khoá trống -> Lỗi
        row_err_cohort = list(row)
        row_err_cohort[14] = ""
        norm, col_st, summary = validate_student_record(row_err_cohort)
        self.assertEqual(col_st["nien_khoa"], "LỖI")
        self.assertEqual(summary, "LỖI")

        # 7. Niên khoá không đúng dạng YYYY-YYYY -> Lỗi
        row_err_cohort2 = list(row)
        row_err_cohort2[14] = "Khóa 2023"
        norm, col_st, summary = validate_student_record(row_err_cohort2)
        self.assertEqual(col_st["nien_khoa"], "LỖI")
        self.assertEqual(summary, "LỖI")

        # 8. Niên khoá Y_ra - Y_vao != 4 -> Lỗi
        row_err_cohort3 = list(row)
        row_err_cohort3[14] = "2020-2023" # chênh 3 năm
        norm, col_st, summary = validate_student_record(row_err_cohort3)
        self.assertEqual(col_st["nien_khoa"], "LỖI")
        self.assertEqual(summary, "LỖI")

        # 9. Niên khoá Y_ra != Năm sinh + 15 -> Cảnh báo
        row_warn_cohort = list(row)
        row_warn_cohort[8] = "15/05/2007" # Sinh 2007, tốt nghiệp 2023 (lệch 16)
        norm, col_st, summary = validate_student_record(row_warn_cohort)
        self.assertEqual(col_st["nien_khoa"], "CẢNH BÁO")
        self.assertEqual(summary, "CẢNH BÁO")

        # 10. Ngày cấp bằng trống -> Cảnh báo
        row_warn_grad = list(row)
        row_warn_grad[18] = ""
        norm, col_st, summary = validate_student_record(row_warn_grad)
        self.assertEqual(col_st["ngay_cap"], "CẢNH BÁO")
        self.assertEqual(summary, "CẢNH BÁO")

        # 11. Ngày cấp bằng không đúng dạng dd/MM/yyyy -> Lỗi
        row_err_grad = list(row)
        row_err_grad[18] = "2023"
        norm, col_st, summary = validate_student_record(row_err_grad)
        self.assertEqual(col_st["ngay_cap"], "LỖI")
        self.assertEqual(summary, "LỖI")

        # 12. Ngày cấp bằng sai logic (30/02) -> Cảnh báo
        row_warn_grad2 = list(row)
        row_warn_grad2[18] = "30/02/2023"
        norm, col_st, summary = validate_student_record(row_warn_grad2)
        self.assertEqual(col_st["ngay_cap"], "CẢNH BÁO")
        self.assertEqual(summary, "CẢNH BÁO")

        # 13. Ngày cấp bằng có năm khác Y_ra -> Cảnh báo
        row_warn_grad3 = list(row)
        row_warn_grad3[18] = "15/06/2024" # Y_ra là 2023
        norm, col_st, summary = validate_student_record(row_warn_grad3)
        self.assertEqual(col_st["ngay_cap"], "CẢNH BÁO")
        self.assertEqual(summary, "CẢNH BÁO")

    def test_data_inspector_process_file(self):
        # Tạo file Excel test 28 cột
        test_file = self.tmp_dir / "Test_DS.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "SoHoaTHCS"

        # Dòng 1: Tiêu đề
        ws.append(["DANH SÁCH CẤP BẰNG TỐT NGHIỆP THCS"])
        # Dòng 2: Header 28 cột chuẩn
        headers = [
            "STT", "Số hiệu bằng", "Số vào sổ gốc cấp văn bằng", "Tên văn bằng", "Họ, chữ đệm và tên",
            "Mã người học", "Số định danh cá nhân", "Hộ chiếu", "Ngày, tháng, năm sinh", "Giới tính",
            "Dân tộc", "Quốc tịch", "Nơi sinh", "Tên trường đã học", "Niên khóa",
            "Điểm các môn thi", "Hội đồng thi", "Địa danh nơi cơ quan cấp bằng đặt trụ sở", "Ngày tháng năm cấp bằng", "Tên cơ quan cấp bằng",
            "Chức danh người ký bằng", "Họ, chữ đệm, tên người ký bằng", "Tình trạng văn bằng", "Ghi chú", "Trạng thái số hóa",
            "Thông tin chưa bảo đảm đầy đủ, chính xác và nguyên nhân", "Mã đơn vị", "Tên đơn vị"
        ]
        ws.append(headers)

        # Dòng 3: Học sinh 1 - Chuẩn
        row1 = [""] * 28
        row1[0] = 1
        row1[4] = "lê văn nam"
        row1[8] = "10/04/2008"
        row1[14] = "2019-2023"
        row1[18] = "20/05/2023"
        ws.append(row1)

        # Dòng 4: Học sinh 2 - Lỗi tên trống, lỗi niên khóa 2020-2023
        row2 = [""] * 28
        row2[0] = 2
        row2[4] = ""
        row2[8] = "12/08/2008"
        row2[14] = "2020-2023"
        row2[18] = "20/05/2023"
        ws.append(row2)

        # Dòng 5: Học sinh 3 - Cảnh báo ngày cấp bằng trống
        row3 = [""] * 28
        row3[0] = 3
        row3[4] = "Trần Thị Mai"
        row3[8] = "05/09/2008"
        row3[14] = "2019-2023"
        row3[18] = ""
        ws.append(row3)

        wb.save(test_file)
        wb.close()

        inspector = DataInspector()
        out_f = self.tmp_dir / "Result_DS.xlsx"
        res = inspector.process_file(test_file, out_f)

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["total_rows"], 3)
        self.assertEqual(res["error_count"], 1)
        self.assertEqual(res["warning_count"], 1)
        self.assertEqual(res["valid_count"], 1)

        # Kiểm tra file kết quả
        wb_res = openpyxl.load_workbook(out_f)
        ws_res = wb_res.active

        # Tiêu đề 5 cột mới ở dòng 2
        self.assertEqual(ws_res.cell(2, 29).value, "[TỔNG HỢP]")
        self.assertEqual(ws_res.cell(2, 30).value, "[KT] Họ tên")
        self.assertEqual(ws_res.cell(2, 31).value, "[KT] Ngày sinh")
        self.assertEqual(ws_res.cell(2, 32).value, "[KT] Niên khoá")
        self.assertEqual(ws_res.cell(2, 33).value, "[KT] Ngày cấp bằng")

        # Dòng 3 (Học sinh 1): Hợp lệ
        self.assertEqual(ws_res.cell(3, 5).value, "Lê Văn Nam")
        self.assertIn(ws_res.cell(3, 29).value, ("", None))

        # Dòng 4 (Học sinh 2): Lỗi
        self.assertEqual(ws_res.cell(4, 29).value, "LỖI")
        self.assertEqual(ws_res.cell(4, 30).value, "LỖI")
        self.assertEqual(ws_res.cell(4, 32).value, "LỖI")

        # Dòng 5 (Học sinh 3): Cảnh báo
        self.assertEqual(ws_res.cell(5, 29).value, "CẢNH BÁO")
        self.assertEqual(ws_res.cell(5, 33).value, "CẢNH BÁO")

        wb_res.close()

    def test_detect_header_on_template_file(self):
        """Kiểm tra nhận diện dòng tiêu đề và 4 cột kiểm tra trực tiếp trên tệp mẫu chuẩn Mau_DS_Cap_Bang_THCS.xlsx."""
        template_p = Path("data/Mau_DS_Cap_Bang_THCS.xlsx")
        if not template_p.exists():
            self.skipTest("Template file data/Mau_DS_Cap_Bang_THCS.xlsx not found")
        wb = openpyxl.load_workbook(template_p)
        ws = wb.active
        inspector = DataInspector()
        hdr_row, col_map = inspector.detect_header_row_and_columns(ws)
        wb.close()

        self.assertEqual(hdr_row, 2)
        self.assertEqual(col_map["ho_ten"], 5)
        self.assertEqual(col_map["ngay_sinh"], 9)
        self.assertEqual(col_map["nien_khoa"], 15)
        self.assertEqual(col_map["ngay_cap"], 19)


if __name__ == "__main__":
    unittest.main()
