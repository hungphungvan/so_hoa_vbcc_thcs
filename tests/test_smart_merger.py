from __future__ import annotations

import unittest
import tempfile
import shutil
from pathlib import Path
import openpyxl

from src.smart_merger import (
    TARGET_COLUMNS,
    SYNONYMS_MAP,
    compute_similarity,
    compute_headers_signature,
    auto_suggest_mapping,
    parse_column_expression,
    MappingProfileManager,
    SmartMerger
)


class TestSmartMerger(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp())
        self.profile_file = self.tmp_dir / "test_profiles.json"
        self.profile_mgr = MappingProfileManager(self.profile_file)

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_target_columns_definition(self):
        """Kiểm tra danh sách 28 cột chuẩn."""
        self.assertEqual(len(TARGET_COLUMNS), 28)
        self.assertEqual(TARGET_COLUMNS[0]["name"], "STT")
        self.assertEqual(TARGET_COLUMNS[4]["name"], "Họ, chữ đệm và tên")
        self.assertEqual(TARGET_COLUMNS[6]["name"], "Số định danh cá nhân")
        self.assertEqual(TARGET_COLUMNS[27]["name"], "Tên đơn vị")

    def test_synonyms_and_fuzzy_similarity(self):
        """Kiểm tra khả năng bắt từ khóa nghiệp vụ & so khớp mờ."""
        # 1. Từ khóa đồng nghĩa nghiệp vụ (Tầng 1)
        score_cccd, m_type = compute_similarity("Số định danh cá nhân", SYNONYMS_MAP[7], "CCCD")
        self.assertGreaterEqual(score_cccd, 0.95)
        self.assertEqual(m_type, "keyword")

        score_quequan, _ = compute_similarity("Nơi sinh", SYNONYMS_MAP[13], "Quê quán")
        self.assertGreaterEqual(score_quequan, 0.95)

        score_phai, _ = compute_similarity("Giới tính", SYNONYMS_MAP[10], "Phái")
        self.assertGreaterEqual(score_phai, 0.95)

        score_so_so, _ = compute_similarity("Số vào sổ gốc cấp văn bằng", SYNONYMS_MAP[3], "Số vào sổ")
        self.assertGreaterEqual(score_so_so, 0.95)

        # 2. So khớp mờ chuỗi (Tầng 2)
        score_so_hieu, _ = compute_similarity("Số hiệu bằng", SYNONYMS_MAP[2], "Số hiệu")
        self.assertGreaterEqual(score_so_hieu, 0.85)

        score_ngay_sinh, _ = compute_similarity("Ngày, tháng, năm sinh", SYNONYMS_MAP[9], "Ngày sinh")
        self.assertGreaterEqual(score_ngay_sinh, 0.85)

        score_truong, _ = compute_similarity("Tên trường đã học", SYNONYMS_MAP[14], "Tên trường")
        self.assertGreaterEqual(score_truong, 0.85)

    def test_auto_suggest_mapping(self):
        """Kiểm tra thuật toán tự động gợi ý ánh xạ."""
        src_headers = ["STT", "Họ và tên", "Ngày sinh", "Quê quán", "CCCD", "Khóa học", "Số hiệu", "Số vào sổ"]
        mapping = auto_suggest_mapping(src_headers)

        self.assertEqual(mapping[1], 0)  # STT
        self.assertEqual(mapping[5], 1)  # Họ, chữ đệm và tên <- Họ và tên
        self.assertEqual(mapping[9], 2)  # Ngày sinh
        self.assertEqual(mapping[13], 3) # Nơi sinh <- Quê quán
        self.assertEqual(mapping[7], 4)  # Số định danh cá nhân <- CCCD
        self.assertEqual(mapping[15], 5) # Niên khóa <- Khóa học
        self.assertEqual(mapping[2], 6)  # Số hiệu bằng <- Số hiệu
        self.assertEqual(mapping[3], 7)  # Số vào sổ gốc <- Số vào sổ

        # Các cột không có trong nguồn phải là None (Để trống)
        self.assertIsNone(mapping[4])   # Tên văn bằng
        self.assertIsNone(mapping[8])   # Hộ chiếu
        self.assertIsNone(mapping[12])  # Quốc tịch
        self.assertIsNone(mapping[16])  # Điểm thi

    def test_profile_manager(self):
        """Kiểm tra lưu, nạp, chữ ký và xóa profile."""
        headers = ["STT", "Họ tên", "Ngày sinh"]
        sig = compute_headers_signature(headers)
        self.assertTrue(len(sig) > 0)

        # Lưu profile mới
        prof = self.profile_mgr.save_profile(
            name="Mẫu Test Trường A",
            signature=sig,
            raw_headers=headers,
            mapping={1: 0, 5: 1, 9: 2}
        )
        self.assertEqual(prof["name"], "Mẫu Test Trường A")
        self.assertEqual(prof["signature"], sig)

        # Tìm theo chữ ký
        found = self.profile_mgr.get_profile_by_signature(sig)
        self.assertIsNotNone(found)
        self.assertEqual(found["mapping"]["5"], 1)

        # Xóa profile
        p_id = found["id"]
        deleted = self.profile_mgr.delete_profile(p_id)
        self.assertTrue(deleted)
        self.assertIsNone(self.profile_mgr.get_profile_by_signature(sig))

    def test_smart_merger_end_to_end(self):
        """Kiểm tra quy trình Ghép & Chuẩn hóa thực tế từ file Excel nguồn."""
        # Tạo file nguồn 1 (8 cột, có dòng footer tổng kết)
        file1 = self.tmp_dir / "src_file_1.xlsx"
        wb1 = openpyxl.Workbook()
        ws1 = wb1.active
        ws1.title = "Lop9A"
        ws1.append(["DANH SÁCH TỐT NGHIỆP TRƯỜNG A"])
        ws1.append(["STT", "Họ và tên", "Ngày sinh", "Giới tính", "CCCD", "Quê quán", "Số hiệu", "Số vào sổ"])
        ws1.append([1, "nguyễn văn an", "15/05/2005", "Nam", "001205001234", "Hà Nội", "THCS001", "01/2020"])
        ws1.append([2, "TRẦN THỊ BÌNH", "20/10/2005", "Nữ", "001205005678", "Hải Phòng", "THCS002", "02/2020"])
        ws1.append(["Tổng số học sinh: 2", None, None, None, None, None, None, None])
        ws1.append(["Người lập biểu", None, None, None, None, None, None, "Hiệu trưởng"])
        wb1.save(file1)
        wb1.close()

        # Tạo file nguồn 2 (Cùng cấu trúc cột với file 1 để kiểm tra tự động nhận diện Profile không hỏi lại)
        file2 = self.tmp_dir / "src_file_2.xlsx"
        wb2 = openpyxl.Workbook()
        ws2 = wb2.active
        ws2.title = "Lop9B"
        ws2.append(["DANH SÁCH TỐT NGHIỆP TRƯỜNG B"])
        ws2.append(["STT", "Họ và tên", "Ngày sinh", "Giới tính", "CCCD", "Quê quán", "Số hiệu", "Số vào sổ"])
        ws2.append([1, "LÊ VĂN CƯỜNG", "01/01/2005", "Nam", "001205009999", "Đà Nẵng", "THCS003", "03/2020"])
        wb2.save(file2)
        wb2.close()

        output_file = self.tmp_dir / "Tong_Hop_Chuan_Hoa_28_Cot.xlsx"

        dialog_calls = []

        def mock_mapping_callback(payload):
            dialog_calls.append(payload["file_name"])
            return payload["suggested_mapping"], "Profile Test Tự Lưu", False

        merger = SmartMerger(profile_mgr=self.profile_mgr)
        res = merger.process_merge_and_standardize(
            source_files=[file1, file2],
            output_file_path=output_file,
            mapping_callback=mock_mapping_callback,
            renumber_stt=True,
            fill_defaults=True,
            merge_all_sheets=True
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["total_rows"], 3)  # 2 từ file 1 + 1 từ file 2 (bỏ dòng tổng kết & footer)

        # QUAN TRỌNG: Callback chỉ được gọi 1 lần cho file 1. File 2 tự động áp dụng profile mà KHÔNG gọi callback!
        self.assertEqual(len(dialog_calls), 1)
        self.assertEqual(dialog_calls[0], file1.name)

        # Kiểm tra nội dung file đích đã ghép chuẩn hóa
        self.assertTrue(output_file.exists())
        wb_res = openpyxl.load_workbook(output_file, data_only=True)
        ws_res = wb_res.active

        self.assertEqual(ws_res.max_row, 5) # Dòng 1 banner, Dòng 2 tiêu đề, Dòng 3,4,5 là 3 học sinh

        # Kiểm tra học sinh 1 (nguyen van an -> Nguyễn Văn An)
        r3 = [ws_res.cell(3, c).value for c in range(1, 29)]
        self.assertEqual(r3[0], 1) # STT đánh lại
        self.assertEqual(r3[1], "THCS001") # Số hiệu
        self.assertEqual(r3[2], "01/2020") # Số vào sổ
        self.assertEqual(r3[3], "Bằng tốt nghiệp trung học cơ sở") # Điền mặc định vì để trống
        self.assertEqual(r3[4], "Nguyễn Văn An") # Title case
        self.assertEqual(r3[6], "001205001234") # CCCD
        self.assertEqual(r3[8], "15/05/2005") # Ngày sinh
        self.assertEqual(r3[9], "Nam") # Giới tính
        self.assertEqual(r3[11], "Việt Nam") # Quốc tịch điền mặc định
        self.assertEqual(r3[12], "Hà Nội") # Nơi sinh

        # Kiểm tra học sinh 3 từ file 2 (LE VAN CUONG -> Lê Văn Cường, STT = 3)
        r5 = [ws_res.cell(5, c).value for c in range(1, 29)]
        self.assertEqual(r5[0], 3) # STT liên tục toàn bảng
        self.assertEqual(r5[4], "Lê Văn Cường") # Title case chuẩn
        self.assertEqual(r5[6], "001205009999") # CCCD bảo toàn số 0 đầu
        self.assertEqual(r5[3], "Bằng tốt nghiệp trung học cơ sở")
        self.assertEqual(r5[11], "Việt Nam")

        wb_res.close()

    def test_smart_merger_without_stt_column(self):
        """Kiểm tra trích xuất học sinh từ sheet hoàn toàn KHÔNG CÓ cột STT (chỉ cần có Họ và tên)."""
        file_no_stt = self.tmp_dir / "src_no_stt.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "KhongCoSTT"
        ws.append(["DANH SÁCH KHÔNG CỘT STT"])
        # Tiêu đề không hề có cột STT
        ws.append(["Họ và tên", "Ngày sinh", "Giới tính", "Nơi sinh"])
        ws.append(["Phạm Nhật Minh", "12/03/2005", "Nam", "Hà Nội"])
        ws.append(["Hoàng Thu Trang", "25/11/2005", "Nữ", "Hải Dương"])
        ws.append(["Vũ Đức Đam", "09/08/2005", "Nam", "Quảng Ninh"])
        ws.append(["Tổng số: 3 học sinh", None, None, None]) # Dòng footer
        wb.save(file_no_stt)
        wb.close()

        output_file = self.tmp_dir / "Out_No_STT.xlsx"

        merger = SmartMerger(profile_mgr=self.profile_mgr)
        res = merger.process_merge_and_standardize(
            source_files=[file_no_stt],
            output_file_path=output_file,
            mapping_callback=None, # Tự động nhận diện
            renumber_stt=True,
            fill_defaults=True,
            merge_all_sheets=True
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["total_rows"], 3) # Phải trích xuất đủ cả 3 học sinh

        wb_res = openpyxl.load_workbook(output_file, data_only=True)
        ws_res = wb_res.active

        # Kiểm tra 3 học sinh đều có STT tự đánh 1, 2, 3
        stt_list = [ws_res.cell(r, 1).value for r in [3, 4, 5]]
        name_list = [ws_res.cell(r, 5).value for r in [3, 4, 5]]
        self.assertEqual(stt_list, [1, 2, 3])
        self.assertEqual(name_list, ["Phạm Nhật Minh", "Hoàng Thu Trang", "Vũ Đức Đam"])

        wb_res.close()

    def test_smart_merger_with_custom_default_values(self):
        """Kiểm tra tính năng nhập giá trị mặc định cho cột để trống (Tên trường, Niên khóa, Cơ quan cấp...)."""
        file1 = self.tmp_dir / "src_custom_defaults_1.xlsx"
        wb1 = openpyxl.Workbook()
        ws1 = wb1.active
        ws1.title = "Lop9"
        ws1.append(["Họ và tên", "Ngày sinh", "Giới tính"])
        ws1.append(["Đặng Thái Sơn", "02/07/2005", "Nam"])
        ws1.append(["Ngô Bảo Châu", "28/06/2005", "Nam"])
        wb1.save(file1)
        wb1.close()

        file2 = self.tmp_dir / "src_custom_defaults_2.xlsx"
        wb2 = openpyxl.Workbook()
        ws2 = wb2.active
        ws2.title = "Lop9"
        ws2.append(["Họ và tên", "Ngày sinh", "Giới tính"])
        ws2.append(["Lưu Lệ Hằng", "04/10/2005", "Nữ"])
        wb2.save(file2)
        wb2.close()

        output_file = self.tmp_dir / "Out_Custom_Defaults.xlsx"

        call_count = 0
        def mock_callback(payload):
            nonlocal call_count
            call_count += 1
            mapping = {
                5: 0, # Họ và tên
                9: 1, # Ngày sinh
                10: 2 # Giới tính
            }
            # Người dùng tự nhập giá trị mặc định cho các cột không có trong file nguồn
            custom_defs = {
                14: "THCS Chu Văn An",
                15: "2019-2023",
                18: "Hà Nội",
                20: "Phòng GD&ĐT Tây Hồ"
            }
            return mapping, "Profile Chu Văn An 2019", custom_defs, False

        merger = SmartMerger(profile_mgr=self.profile_mgr)
        res = merger.process_merge_and_standardize(
            source_files=[file1, file2],
            output_file_path=output_file,
            mapping_callback=mock_callback,
            renumber_stt=True,
            fill_defaults=True,
            merge_all_sheets=True
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["total_rows"], 3)
        self.assertEqual(call_count, 1) # File 2 tự động áp dụng Profile kèm cả default_values

        wb_res = openpyxl.load_workbook(output_file, data_only=True)
        ws_res = wb_res.active

        # Kiểm tra học sinh 1 (Đặng Thái Sơn)
        self.assertEqual(ws_res.cell(3, 1).value, 1) # STT
        self.assertEqual(ws_res.cell(3, 4).value, "Bằng tốt nghiệp trung học cơ sở") # Mặc định hệ thống
        self.assertEqual(ws_res.cell(3, 5).value, "Đặng Thái Sơn") # Họ tên
        self.assertEqual(ws_res.cell(3, 12).value, "Việt Nam") # Mặc định hệ thống
        self.assertEqual(ws_res.cell(3, 14).value, "THCS Chu Văn An") # Mặc định do người dùng nhập
        self.assertEqual(ws_res.cell(3, 15).value, "2019-2023") # Mặc định do người dùng nhập
        self.assertEqual(ws_res.cell(3, 18).value, "Hà Nội") # Mặc định do người dùng nhập
        self.assertEqual(ws_res.cell(3, 20).value, "Phòng GD&ĐT Tây Hồ") # Mặc định do người dùng nhập

        # Kiểm tra học sinh 3 (Lưu Lệ Hằng từ file 2 - tự động thừa hưởng default_values của Profile)
        self.assertEqual(ws_res.cell(5, 1).value, 3) # STT
        self.assertEqual(ws_res.cell(5, 5).value, "Lưu Lệ Hằng")
        self.assertEqual(ws_res.cell(5, 14).value, "THCS Chu Văn An")
        self.assertEqual(ws_res.cell(5, 15).value, "2019-2023")
        self.assertEqual(ws_res.cell(5, 20).value, "Phòng GD&ĐT Tây Hồ")

        wb_res.close()

    def test_parse_column_expression(self):
        """Kiểm tra phân tích các cú pháp biểu thức ghép cột (Phương án 3)."""
        formatted = ["Cột A: STT", "Cột B: Họ và chữ đệm", "Cột C: Tên", "Cột D: Ngày sinh", "Cột E: Quê quán"]
        
        # 1. Để trống
        self.assertIsNone(parse_column_expression("", formatted))
        self.assertIsNone(parse_column_expression("-- Để trống --", formatted))
        self.assertIsNone(parse_column_expression("--", formatted))

        # 2. Cột đơn khớp chính xác
        self.assertEqual(parse_column_expression("Cột B: Họ và chữ đệm", formatted), 1)
        self.assertEqual(parse_column_expression("Cột D: Ngày sinh", formatted), 3)

        # 3. Biểu thức ghép bằng chữ cái cột: B + C, b, c, B & C
        self.assertEqual(parse_column_expression("B + C", formatted), [1, 2])
        self.assertEqual(parse_column_expression("b, c", formatted), [1, 2])
        self.assertEqual(parse_column_expression("B & C", formatted), [1, 2])
        self.assertEqual(parse_column_expression("cột b + cột c", formatted), [1, 2])

        # 4. Biểu thức ghép bằng nhãn đầy đủ
        self.assertEqual(parse_column_expression("Cột B: Họ và chữ đệm + Cột C: Tên", formatted), [1, 2])

        # 5. Biểu thức ghép bằng số thứ tự 1-indexed: 2 + 3
        self.assertEqual(parse_column_expression("2 + 3", formatted), [1, 2])

        # 6. Biểu thức ghép bằng từ khóa tương đối
        self.assertEqual(parse_column_expression("Họ và chữ đệm + Tên", formatted), [1, 2])

    def test_auto_suggest_mapping_split_name(self):
        """Kiểm tra thuật toán tự động nhận diện tệp nguồn có cột Họ đệm và Tên tách rời."""
        # Trường hợp 1: Họ và chữ đệm + Tên
        headers1 = ["STT", "Họ và chữ đệm", "Tên", "Ngày sinh", "Giới tính"]
        m1 = auto_suggest_mapping(headers1)
        self.assertEqual(m1[1], 0)        # STT
        self.assertEqual(m1[5], [1, 2])   # Họ, chữ đệm và tên <- ghép cột 1 và 2
        self.assertEqual(m1[9], 3)        # Ngày sinh
        self.assertEqual(m1[10], 4)       # Giới tính

        # Trường hợp 2: Họ đệm + Tên
        headers2 = ["STT", "Họ đệm", "Tên", "Ngày sinh"]
        m2 = auto_suggest_mapping(headers2)
        self.assertEqual(m2[5], [1, 2])

        # Trường hợp 3: Họ + Tên (không có STT)
        headers3 = ["Họ", "Tên", "Ngày sinh"]
        m3 = auto_suggest_mapping(headers3)
        self.assertEqual(m3[5], [0, 1])

        # Trường hợp 4: Có cột Họ và tên gộp thì không tách
        headers4 = ["STT", "Họ và tên", "Ngày sinh"]
        m4 = auto_suggest_mapping(headers4)
        self.assertEqual(m4[5], 1)

    def test_smart_merger_with_split_name_columns(self):
        """Kiểm tra quy trình Ghép & Chuẩn hóa thực tế khi cột Họ và Tên bị tách thành 2 cột."""
        file_split = self.tmp_dir / "src_split_name.xlsx"
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Lop9_TachCot"
        ws.append(["DANH SÁCH TỐT NGHIỆP CÓ HỌ VÀ TÊN TÁCH RỜI"])
        ws.append(["STT", "Họ và chữ đệm", "Tên", "Ngày sinh", "Nơi sinh"])
        ws.append([1, "nguyễn văn", "an", "15/05/2005", "Hà Nội"])
        ws.append([2, "TRẦN THỊ BÍCH", "NGỌC", "20/10/2005", "Hải Phòng"])
        ws.append([3, "  LÊ   VĂN ", " CƯỜNG ", "01/01/2005", "Đà Nẵng"])
        ws.append(["Tổng số học sinh: 3", None, None, None, None]) # Dòng footer
        wb.save(file_split)
        wb.close()

        output_file = self.tmp_dir / "Out_Split_Name.xlsx"

        merger = SmartMerger(profile_mgr=self.profile_mgr)
        res = merger.process_merge_and_standardize(
            source_files=[file_split],
            output_file_path=output_file,
            mapping_callback=None, # Tự động phát hiện mapping[5] = [1, 2]
            renumber_stt=True,
            fill_defaults=True,
            merge_all_sheets=True
        )

        self.assertEqual(res["status"], "success")
        self.assertEqual(res["total_rows"], 3)

        wb_res = openpyxl.load_workbook(output_file, data_only=True)
        ws_res = wb_res.active

        # Kiểm tra học sinh 1: "nguyễn văn" + "an" -> "Nguyễn Văn An"
        self.assertEqual(ws_res.cell(3, 1).value, 1)
        self.assertEqual(ws_res.cell(3, 5).value, "Nguyễn Văn An")
        self.assertEqual(ws_res.cell(3, 9).value, "15/05/2005")
        self.assertEqual(ws_res.cell(3, 13).value, "Hà Nội")

        # Kiểm tra học sinh 2: "TRẦN THỊ BÍCH" + "NGỌC" -> "Trần Thị Bích Ngọc"
        self.assertEqual(ws_res.cell(4, 1).value, 2)
        self.assertEqual(ws_res.cell(4, 5).value, "Trần Thị Bích Ngọc")

        # Kiểm tra học sinh 3: khoảng trắng thừa được dọn sạch -> "Lê Văn Cường"
        self.assertEqual(ws_res.cell(5, 1).value, 3)
        self.assertEqual(ws_res.cell(5, 5).value, "Lê Văn Cường")

        wb_res.close()

    def test_profile_with_combined_columns(self):
        """Kiểm tra lưu và nạp lại Profile chứa cấu hình ghép nhiều cột."""
        headers = ["STT", "Họ và chữ đệm", "Tên", "Ngày sinh"]
        sig = compute_headers_signature(headers)

        # Lưu profile với mapping Cột 5 = [1, 2]
        prof = self.profile_mgr.save_profile(
            name="Mẫu Tách Cột Test",
            signature=sig,
            raw_headers=headers,
            mapping={1: 0, 5: [1, 2], 9: 3}
        )
        self.assertEqual(prof["mapping"]["5"], [1, 2])

        # Tìm lại theo signature
        found = self.profile_mgr.get_profile_by_signature(sig)
        self.assertIsNotNone(found)
        self.assertEqual(found["mapping"]["5"], [1, 2])


if __name__ == '__main__':
    unittest.main()
