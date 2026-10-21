import os
import sys
import tempfile
import unittest
from pathlib import Path

# Thêm thư mục gốc vào sys.path
ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.db_manager import DatabaseManager
from src.lookup_manager import LookupManager

class TestPhanHe2(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_vbcc_ph2.db"
        self.db_mgr = DatabaseManager(self.db_path)

        # Tạo một số học sinh mẫu
        self.sample_students = [
            {
                "final_row": [
                    1, "A001", "01/2005", "Bằng tốt nghiệp trung học cơ sở", "Nguyễn Văn Tuấn",
                    None, "025090001111", None, "15/08/1990", "Nam",
                    "Kinh", "Việt Nam", "Hòa Bình", "THCS Lạc Thịnh", "2004-2005",
                    "8.5", "HĐ Lạc Thịnh", "Yên Thủy-Hòa Bình", "25/06/2005", "Phòng GD&ĐT Yên Thủy",
                    "Trưởng phòng", "Bùi Văn Nam", "Cấp mới", None, "Đã số hóa đầy đủ, chính xác",
                    "", "DV01", "UBND Xã Lạc Thịnh"
                ],
                "school_code": "SCH_LT",
                "commune_code": "XA_LT",
                "commune": "Lạc Thịnh",
                "district_code": "H_YT",
                "district_name": "Yên Thủy",
                "province_name": "Hòa Bình",
                "source_file": "2004/THCS_LacThinh.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 4
            },
            {
                "final_row": [
                    2, "A002", "02/2005", "Bằng tốt nghiệp trung học cơ sở", "Trần Thị Hoa",
                    None, "025090002222", None, "1990", "Nữ",
                    "Mường", "Việt Nam", "Hòa Bình", "THCS Lạc Thịnh", "2004-2005",
                    "7.0", "HĐ Lạc Thịnh", "Yên Thủy-Hòa Bình", "25/06/2005", "Phòng GD&ĐT Yên Thủy",
                    "Trưởng phòng", "Bùi Văn Nam", "Cấp mới", None, "Số hoá chưa đầy đủ, chính xác",
                    "Chỉ có năm sinh", "DV01", "UBND Xã Lạc Thịnh"
                ],
                "school_code": "SCH_LT",
                "commune_code": "XA_LT",
                "commune": "Lạc Thịnh",
                "district_code": "H_YT",
                "district_name": "Yên Thủy",
                "province_name": "Hòa Bình",
                "source_file": "2004/THCS_LacThinh.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 5
            },
            {
                "final_row": [
                    3, "A003", "03/2005", "Bằng tốt nghiệp trung học cơ sở", "Lê Văn Hùng",
                    None, None, None, "", "Nam",
                    "Kinh", "Việt Nam", "Hòa Bình", "THCS Bình Hẻm", "2003-2004",
                    "6.5", "HĐ Lạc Thủy", "Lạc Thủy-Hòa Bình", "20/06/2004", "Phòng GD&ĐT Lạc Thủy",
                    "Trưởng phòng", "Lê Văn Thắng", "Cấp mới", None, "Chưa số hoá",
                    "Thiếu ngày sinh", "DV02", "UBND Xã Bình Hẻm"
                ],
                "school_code": "SCH_BH",
                "commune_code": "XA_BH",
                "commune": "Bình Hẻm",
                "district_code": "H_LT",
                "district_name": "Lạc Thủy",
                "province_name": "Hòa Bình",
                "source_file": "2004/THCS_BinhHem.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 6
            },
            {
                "final_row": [
                    4, "A004", "04/2005", "Bằng tốt nghiệp trung học cơ sở", "Phạm Thị Lan",
                    None, None, None, "10/02/1988", "Nữ",
                    "Kinh", "Việt Nam", "Phú Thọ", "THCS Phong Châu", "2002-2003",
                    "9.0", "HĐ Phù Ninh", "Phù Ninh-Phú Thọ", "15/06/2003", "Phòng GD&ĐT Phù Ninh",
                    "Trưởng phòng", "Nguyễn Văn Đức", "Cấp mới", None, "Đã số hóa đầy đủ, chính xác",
                    "", "DV03", "UBND Xã Phong Châu"
                ],
                "school_code": "SCH_PC",
                "commune_code": "XA_PC",
                "commune": "Phong Châu",
                "district_code": "H_PN",
                "district_name": "Phù Ninh",
                "province_name": "Phú Thọ",
                "source_file": "2003/THCS_PhongChau.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 7
            }
        ]
        self.db_mgr.insert_students_batch(self.sample_students)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_all_11_indexes_exist(self):
        """Kiểm tra toàn bộ 11 chỉ mục mới và cũ đã được tạo thành công trong SQLite."""
        with self.db_mgr.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='hoc_sinh';")
            idx_names = {row["name"] for row in cursor.fetchall()}

            expected_indexes = [
                "idx_ho_ten",
                "idx_ngay_sinh",
                "idx_ten_truong",
                "idx_nien_khoa",
                "idx_ten_don_vi",
                "idx_ho_ten_ngay_sinh",
                "idx_truong_nien_khoa",
                "idx_dia_ban",
                "idx_xa_don_vi",
                "idx_trang_thai_loi",
                "idx_source_file_sheet"
            ]
            for idx in expected_indexes:
                self.assertIn(idx, idx_names, f"Chỉ mục {idx} chưa được khởi tạo!")

    def test_cascade_filter_options(self):
        """Kiểm tra hàm get_cascade_filter_options theo 2 nhánh."""
        # 1. Toàn bộ
        opts = self.db_mgr.get_cascade_filter_options()
        self.assertIn("Hòa Bình", opts["tinh_cu"])
        self.assertIn("Phú Thọ", opts["tinh_cu"])
        self.assertEqual(len(opts["tinh_cu"]), 2)

        # 2. Lọc Huyện cũ theo Tỉnh cũ Hòa Bình
        opts_hb = self.db_mgr.get_cascade_filter_options(tinh_cu="Hòa Bình")
        self.assertIn("Yên Thủy", opts_hb["huyen_cu"])
        self.assertIn("Lạc Thủy", opts_hb["huyen_cu"])
        self.assertNotIn("Phù Ninh", opts_hb["huyen_cu"])

        # 3. Lọc Trường theo Huyện cũ Yên Thủy
        opts_yt = self.db_mgr.get_cascade_filter_options(huyen_cu="Yên Thủy")
        self.assertIn("THCS Lạc Thịnh", opts_yt["ten_truong"])
        self.assertNotIn("THCS Bình Hẻm", opts_yt["ten_truong"])

        # 4. Nhánh 2: Lọc Đơn vị theo Xã/phường
        opts_xa = self.db_mgr.get_cascade_filter_options(xa_phuong="Lạc Thịnh")
        self.assertIn("UBND Xã Lạc Thịnh", opts_xa["ten_don_vi"])
        self.assertNotIn("UBND Xã Bình Hẻm", opts_xa["ten_don_vi"])

    def test_quick_search_two_fields(self):
        """Kiểm tra Thanh tìm kiếm tức thì theo Họ tên và Ngày sinh (Cách 2)."""
        # 1. Tìm chỉ theo Họ tên
        rows, count = self.db_mgr.query_students_paginated(search_name="Tuấn")
        self.assertEqual(count, 1)
        self.assertEqual(rows[0]["ho_ten"], "Nguyễn Văn Tuấn")

        # 2. Tìm chỉ theo Ngày sinh / Năm sinh
        rows, count = self.db_mgr.query_students_paginated(search_dob="1990")
        self.assertEqual(count, 2)  # Tuấn (15/08/1990) và Hoa (1990)

        # 3. Kết hợp cả Họ tên và Ngày sinh
        rows, count = self.db_mgr.query_students_paginated(search_name="Tuấn", search_dob="15/08/1990")
        self.assertEqual(count, 1)
        self.assertEqual(rows[0]["ho_ten"], "Nguyễn Văn Tuấn")

        # 4. Tìm không tồn tại
        rows, count = self.db_mgr.query_students_paginated(search_name="Không tồn tại")
        self.assertEqual(count, 0)

    def test_multi_filter_branch_1(self):
        """Kiểm tra Nhánh 1: Tỉnh cũ -> Huyện cũ -> Tên trường -> Niên khóa."""
        # Lọc theo Tỉnh cũ Hòa Bình + Niên khóa 2004-2005
        filters = {
            "tinh_cu": "Hòa Bình",
            "nien_khoa": "2004-2005"
        }
        rows, count = self.db_mgr.query_students_paginated(filters=filters)
        self.assertEqual(count, 2)
        names = {r["ho_ten"] for r in rows}
        self.assertEqual(names, {"Nguyễn Văn Tuấn", "Trần Thị Hoa"})

        # Lọc chính xác 1 trường
        filters_sch = {"ten_truong": "THCS Bình Hẻm"}
        rows, count = self.db_mgr.query_students_paginated(filters=filters_sch)
        self.assertEqual(count, 1)
        self.assertEqual(rows[0]["ho_ten"], "Lê Văn Hùng")

    def test_multi_filter_branch_2(self):
        """Kiểm tra Nhánh 2: Xã/phường mới -> Tên Đơn vị."""
        filters = {
            "xa_phuong": "Phong Châu",
            "ten_don_vi": "UBND Xã Phong Châu"
        }
        rows, count = self.db_mgr.query_students_paginated(filters=filters)
        self.assertEqual(count, 1)
        self.assertEqual(rows[0]["ho_ten"], "Phạm Thị Lan")

    def test_filter_errors_only(self):
        """Kiểm tra lọc chỉ xem hồ sơ có lỗi/cảnh báo."""
        filters = {"chi_xem_loi": True}
        rows, count = self.db_mgr.query_students_paginated(filters=filters)
        # Hoa có lỗi "Chỉ có năm sinh", Hùng có lỗi "Thiếu ngày sinh"
        self.assertEqual(count, 2)
        names = {r["ho_ten"] for r in rows}
        self.assertEqual(names, {"Trần Thị Hoa", "Lê Văn Hùng"})

    def test_pagination_and_sorting(self):
        """Kiểm tra phân trang LIMIT / OFFSET và sắp xếp."""
        # Trang 1: page_size = 2
        rows_p1, count = self.db_mgr.query_students_paginated(page=1, page_size=2, sort_by="id", sort_order="ASC")
        self.assertEqual(count, 4)
        self.assertEqual(len(rows_p1), 2)
        self.assertEqual(rows_p1[0]["id"], 1)
        self.assertEqual(rows_p1[1]["id"], 2)

        # Trang 2: page_size = 2
        rows_p2, count = self.db_mgr.query_students_paginated(page=2, page_size=2, sort_by="id", sort_order="ASC")
        self.assertEqual(len(rows_p2), 2)
        self.assertEqual(rows_p2[0]["id"], 3)
        self.assertEqual(rows_p2[1]["id"], 4)

        # Sắp xếp DESC theo Họ tên
        rows_sort, count = self.db_mgr.query_students_paginated(page=1, page_size=10, sort_by="ho_ten", sort_order="DESC")
        names_ordered = [r["ho_ten"] for r in rows_sort]
        self.assertEqual(names_ordered, sorted(names_ordered, reverse=True))

    def test_update_student_recomputes_status(self):
        """Kiểm tra sửa học sinh đơn lẻ tự động recompute trạng thái & lý do lỗi."""
        # Học sinh 2 (Trần Thị Hoa) ban đầu chỉ có năm sinh 1990 (Số hoá chưa đầy đủ)
        st2 = self.db_mgr.get_student_by_id(2)
        self.assertEqual(st2["trang_thai"], "Số hoá chưa đầy đủ, chính xác")

        # Cập nhật ngày sinh đầy đủ và hợp lệ: 20/10/1990
        ok = self.db_mgr.update_student(2, {"ngay_sinh": "20/10/1990"})
        self.assertTrue(ok)

        # Kiểm tra lại: Trạng thái phải tự động thành Đã số hóa đầy đủ và lý do lỗi rỗng!
        st2_updated = self.db_mgr.get_student_by_id(2)
        self.assertEqual(st2_updated["ngay_sinh"], "20/10/1990")
        self.assertEqual(st2_updated["trang_thai"], "Đã số hóa đầy đủ, chính xác")
        self.assertEqual(st2_updated["ly_do_loi"], "")

        # Sửa thành ngày sinh lỗi: 31/02/1990 -> Tự động thành Số hoá chưa đầy đủ
        self.db_mgr.update_student(2, {"ngay_sinh": "31/02/1990"})
        st2_err = self.db_mgr.get_student_by_id(2)
        self.assertEqual(st2_err["trang_thai"], "Số hoá chưa đầy đủ, chính xác")
        self.assertIn("Tháng 02", st2_err["ly_do_loi"])

    def test_batch_update_students(self):
        """Kiểm tra sửa hàng loạt nhiều học sinh cùng lúc."""
        ids_to_update = [1, 2]
        updated_count = self.db_mgr.batch_update_students(
            student_ids=ids_to_update,
            updates={"nien_khoa": "2005-2006", "ngay_cap_bang": "30/06/2006"}
        )
        self.assertEqual(updated_count, 2)

        st1 = self.db_mgr.get_student_by_id(1)
        st2 = self.db_mgr.get_student_by_id(2)
        self.assertEqual(st1["nien_khoa"], "2005-2006")
        self.assertEqual(st1["ngay_cap_bang"], "30/06/2006")
        self.assertEqual(st2["nien_khoa"], "2005-2006")
        self.assertEqual(st2["ngay_cap_bang"], "30/06/2006")

    def test_delete_students(self):
        """Kiểm tra xóa danh sách học sinh."""
        del_count = self.db_mgr.delete_students([3, 4])
        self.assertEqual(del_count, 2)

        st3 = self.db_mgr.get_student_by_id(3)
        self.assertIsNone(st3)
        total_remaining = self.db_mgr.get_statistics()["total"]
        self.assertEqual(total_remaining, 2)

if __name__ == "__main__":
    unittest.main()
