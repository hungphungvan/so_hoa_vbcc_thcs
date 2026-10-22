from __future__ import annotations

import unittest
import tempfile
import shutil
from pathlib import Path
import openpyxl

from src.smart_merger import (
    TARGET_COLUMNS,
    SYNONYMS_MAP,
    remove_accents,
    normalize_text_for_match,
    compute_similarity,
    compute_headers_signature,
    is_standard_28_headers,
    auto_suggest_mapping,
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


if __name__ == '__main__':
    unittest.main()
