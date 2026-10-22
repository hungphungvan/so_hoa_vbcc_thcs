import sys
import tempfile
import unittest
from pathlib import Path

# Thêm thư mục gốc vào sys.path
ROOT_DIR = Path(__file__).resolve().parent.parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from src.db_manager import DatabaseManager

class TestDatabaseManager(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.db_path = Path(self.temp_dir.name) / "test_vbcc.db"
        self.db_mgr = DatabaseManager(self.db_path)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_table_and_indexes_created(self):
        """Kiểm tra bảng hoc_sinh và 4 chỉ mục được tạo đầy đủ."""
        with self.db_mgr.get_connection() as conn:
            cursor = conn.cursor()
            
            # Kiểm tra bảng hoc_sinh
            cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='hoc_sinh';")
            self.assertIsNotNone(cursor.fetchone())

            # Kiểm tra các cột trong bảng
            cursor.execute("PRAGMA table_info(hoc_sinh);")
            columns = {row["name"] for row in cursor.fetchall()}
            
            # 28 cột chuẩn
            expected_28 = [
                "stt", "so_hieu_bang", "so_vao_so", "ten_van_bang", "ho_ten",
                "ma_nguoi_hoc", "cccd", "ho_chieu", "ngay_sinh", "gioi_tinh",
                "dan_toc", "quoc_tich", "noi_sinh", "ten_truong", "nien_khoa",
                "diem_thi", "hoi_dong_thi", "dia_danh", "ngay_cap_bang", "co_quan_cap_bang",
                "chuc_danh_nguoi_ky", "nguoi_ky", "tinh_trang_van_bang", "ghi_chu", "trang_thai",
                "ly_do_loi", "ma_don_vi", "ten_don_vi"
            ]
            for col in expected_28:
                self.assertIn(col, columns, f"Cột {col} thiếu trong CSDL")

            # 6 cột hành chính + 3 cột nguồn gốc
            extra_cols = [
                "ma_truong", "ma_xa_phuong", "xa_phuong", "ma_huyen_cu", "huyen_cu", "tinh_cu",
                "source_file", "source_sheet", "source_row"
            ]
            for col in extra_cols:
                self.assertIn(col, columns, f"Cột mở rộng {col} thiếu trong CSDL")

            # Kiểm tra 4 chỉ mục
            cursor.execute("SELECT name FROM sqlite_master WHERE type='index' AND tbl_name='hoc_sinh';")
            indexes = {row["name"] for row in cursor.fetchall()}
            self.assertIn("idx_ho_ten_ngay_sinh", indexes)
            self.assertIn("idx_truong_nien_khoa", indexes)
            self.assertIn("idx_dia_ban", indexes)
            self.assertIn("idx_source_file_sheet", indexes)

    def test_insert_and_query_students(self):
        """Kiểm tra chèn danh sách học sinh và kiểm đếm thống kê."""
        students = [
            {
                "final_row": [
                    1, "B12345", "01/2005", "Bằng tốt nghiệp trung học cơ sở", "Nguyễn Văn An",
                    None, None, None, "15/05/1990", "Nam",
                    "Kinh", "Việt Nam", "Hòa Bình", "THCS Bình Thanh", "2004-2005",
                    None, None, "Kỳ Sơn-Hòa Bình", "25/06/2005", "Phòng GD&ĐT Kỳ Sơn",
                    "Trưởng phòng", "Bùi Văn Hùng", "Cấp mới", None, "Đã số hóa đầy đủ, chính xác",
                    None, "DV01", "Phòng GD&ĐT Kỳ Sơn"
                ],
                "school_code": "SCH01",
                "commune_code": "XA01",
                "commune": "Bình Thanh",
                "district_code": "H01",
                "district_name": "Kỳ Sơn",
                "province_name": "Hòa Bình",
                "source_file": "2004/THCS_BinhThanh.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 4
            },
            {
                "final_row": [
                    2, "B12346", "02/2005", "Bằng tốt nghiệp trung học cơ sở", "Trần Thị Bình",
                    None, None, None, "20/10/1990", "Nữ",
                    "Mường", "Việt Nam", "Hòa Bình", "THCS Bình Thanh", "2004-2005",
                    None, None, "Kỳ Sơn-Hòa Bình", "25/06/2005", "Phòng GD&ĐT Kỳ Sơn",
                    "Trưởng phòng", "Bùi Văn Hùng", "Cấp mới", None, "Đã số hóa đầy đủ, chính xác",
                    None, "DV01", "Phòng GD&ĐT Kỳ Sơn"
                ],
                "school_code": "SCH01",
                "commune_code": "XA01",
                "commune": "Bình Thanh",
                "district_code": "H01",
                "district_name": "Kỳ Sơn",
                "province_name": "Hòa Bình",
                "source_file": "2004/THCS_BinhThanh.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 5
            }
        ]

        res = self.db_mgr.insert_students_batch(students)
        self.assertEqual(res["inserted"], 2)

        stats = self.db_mgr.get_statistics()
        self.assertEqual(stats["total"], 2)
        self.assertEqual(stats["da_so_hoa"], 2)
        self.assertEqual(stats["total_schools"], 1)
        self.assertEqual(stats["total_districts"], 1)

    def test_check_existing_sheets_and_overwrite(self):
        """Kiểm tra cơ chế phát hiện sheet trùng và ghi đè cấp sheet."""
        # Nạp lần 1: 2 học sinh ở Sheet1
        students_batch1 = [
            {
                "final_row": [1, None, None, "Bằng tốt nghiệp THCS", "Học sinh 1"] + [None]*23,
                "source_file": "THCS_A.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 3
            },
            {
                "final_row": [2, None, None, "Bằng tốt nghiệp THCS", "Học sinh 2"] + [None]*23,
                "source_file": "THCS_A.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 4
            }
        ]
        self.db_mgr.insert_students_batch(students_batch1)

        # Kiểm tra tồn tại
        existing = self.db_mgr.check_existing_sheets(["THCS_A.xlsx"])
        self.assertIn(("THCS_A.xlsx", "Sheet1"), existing)
        self.assertEqual(existing[("THCS_A.xlsx", "Sheet1")], 2)

        # Nạp lần 2 (Ghi đè với 3 học sinh mới)
        students_batch2 = [
            {
                "final_row": [1, None, None, "Bằng tốt nghiệp THCS", "Học sinh Mới 1"] + [None]*23,
                "source_file": "THCS_A.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 3
            },
            {
                "final_row": [2, None, None, "Bằng tốt nghiệp THCS", "Học sinh Mới 2"] + [None]*23,
                "source_file": "THCS_A.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 4
            },
            {
                "final_row": [3, None, None, "Bằng tốt nghiệp THCS", "Học sinh Mới 3"] + [None]*23,
                "source_file": "THCS_A.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 5
            }
        ]
        overwrite_set = {("THCS_A.xlsx", "Sheet1")}
        res2 = self.db_mgr.insert_students_batch(students_batch2, overwrite_sheets=overwrite_set)

        self.assertEqual(res2["inserted"], 3)
        self.assertEqual(res2["deleted_old"], 2)

        stats = self.db_mgr.get_statistics()
        # Tổng số bản ghi phải là 3 (không phải 5!)
        self.assertEqual(stats["total"], 3)

    def test_backup_and_restore(self):
        """Kiểm tra sao lưu trực tuyến và khôi phục CSDL."""
        students = [
            {
                "final_row": [1, None, None, "Bằng tốt nghiệp THCS", "Học sinh Test Backup"] + [None]*23,
                "source_file": "test.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 3
            }
        ]
        self.db_mgr.insert_students_batch(students)

        backup_dir = Path(self.temp_dir.name) / "backups"
        b_file = self.db_mgr.backup_database(backup_dir=backup_dir, max_keep=3)
        self.assertTrue(b_file.is_file())

        # Thêm 1 bản ghi mới sau khi backup
        self.db_mgr.insert_students_batch([
            {
                "final_row": [2, None, None, "Bằng tốt nghiệp THCS", "Học sinh Sau Backup"] + [None]*23,
                "source_file": "test.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": 4
            }
        ])
        self.assertEqual(self.db_mgr.get_statistics()["total"], 2)

        # Khôi phục từ bản backup -> số lượng phải trở về 1
        restored = self.db_mgr.restore_database(b_file)
        self.assertTrue(restored)
        self.assertEqual(self.db_mgr.get_statistics()["total"], 1)

    def test_administrative_and_fallback_data(self):
        """Kiểm tra lưu trữ đầy đủ 6 trường hành chính và quy tắc fallback Cột 18, 20."""
        student = {
            "final_row": [
                1, "SERI001", "001", "Bằng tốt nghiệp trung học cơ sở", "Bùi Thị Lan",
                None, "017199000123", None, "12/03/1990", "Nữ",
                "Mường", "Việt Nam", "Hòa Bình", "THCS Lạc Lương", "2004-2005",
                None, None, "Yên Thủy-Hòa Bình", "15/06/2005", "Phòng GD&ĐT Yên Thủy",
                "Trưởng phòng", "Nguyễn Văn C", "Cấp mới", None, "Đã số hóa đầy đủ, chính xác",
                None, "DV_YT", "Phòng GD&ĐT Yên Thủy"
            ],
            "school_code": "SCH_LL",
            "commune_code": "XA_LL",
            "commune": "Lạc Lương",
            "district_code": "H_YT",
            "district_name": "Yên Thủy",
            "province_name": "Hòa Bình",
            "source_file": "2005/THCS_LacLuong.xlsx",
            "source_sheet": "Lop9A",
            "row_idx": 4
        }
        self.db_mgr.insert_students_batch([student])

        with self.db_mgr.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM hoc_sinh WHERE ho_ten = 'Bùi Thị Lan';")
            row = cursor.fetchone()
            self.assertIsNotNone(row)
            self.assertEqual(row["dia_danh"], "Yên Thủy-Hòa Bình")
            self.assertEqual(row["co_quan_cap_bang"], "Phòng GD&ĐT Yên Thủy")
            self.assertEqual(row["ma_huyen_cu"], "H_YT")
            self.assertEqual(row["huyen_cu"], "Yên Thủy")
            self.assertEqual(row["tinh_cu"], "Hòa Bình")
            self.assertEqual(row["ma_xa_phuong"], "XA_LL")
            self.assertEqual(row["xa_phuong"], "Lạc Lương")
            self.assertEqual(row["source_file"], "2005/THCS_LacLuong.xlsx")
            self.assertEqual(row["source_sheet"], "Lop9A")
            self.assertEqual(row["source_row"], 4)

        districts = self.db_mgr.get_all_districts()
        self.assertIn("Yên Thủy", districts)

        schools = self.db_mgr.get_all_schools_in_db()
        self.assertIn("THCS Lạc Lương", schools)

    def test_performance_batch_insert(self):
        """Kiểm tra hiệu năng chèn 2.000 học sinh."""
        students = []
        for i in range(1, 2001):
            students.append({
                "final_row": [i, f"SERI{i:05d}", f"{i}/2005", "Bằng tốt nghiệp THCS", f"Học Sinh Thứ {i}"] + [None]*23,
                "school_code": f"SCH_{i%20}",
                "district_name": f"Huyện {i%5}",
                "province_name": "Hòa Bình",
                "source_file": f"nam_2005/file_{i%10}.xlsx",
                "source_sheet": "Sheet1",
                "row_idx": i
            })
        
        import time
        t0 = time.time()
        res = self.db_mgr.insert_students_batch(students)
        elapsed = time.time() - t0
        
        self.assertEqual(res["inserted"], 2000)
        self.assertLess(elapsed, 1.0, f"Thời gian chèn 2000 dòng mất {elapsed:.2f}s, quá chậm so với kỳ vọng < 1.0s")

if __name__ == "__main__":
    unittest.main()
