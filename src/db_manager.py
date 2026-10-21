from __future__ import annotations

import os
import sys
import sqlite3
import datetime
import shutil
import re
from pathlib import Path
from typing import Callable, Any

try:
    from src.normalizer import check_dob_validity, determine_status_and_reason
    from src.lookup_manager import get_persistent_data_dir
except ImportError:
    from normalizer import check_dob_validity, determine_status_and_reason
    from lookup_manager import get_persistent_data_dir

def get_default_db_path() -> Path:
    """Trả về đường dẫn tệp CSDL mặc định vbcc_database.db trong thư mục dữ liệu bền vững."""
    base_dir = get_persistent_data_dir()
    return base_dir / "vbcc_database.db"

class DatabaseManager:
    """
    Quản lý tập trung CSDL SQLite cho hệ thống Số hóa Văn bằng Chứng chỉ THCS.
    - Áp dụng cấu hình WAL mode, cache tối ưu 64MB RAM.
    - Đảm bảo toàn vẹn dữ liệu (ACID) qua Transaction.
    - Hỗ trợ Online Backup & Restore.
    - Quản lý 28 cột chuẩn biểu mẫu + 6 cột hành chính + 3 cột nguồn gốc.
    """
    def __init__(self, db_path: Path | str | None = None):
        if db_path is None:
            self.db_path = get_default_db_path()
        else:
            self.db_path = Path(db_path)
            self.db_path.parent.mkdir(parents=True, exist_ok=True)
            
        self.initialize_db()

    def get_connection(self) -> sqlite3.Connection:
        """Tạo kết nối SQLite với cấu hình hiệu năng cao."""
        conn = sqlite3.connect(str(self.db_path), timeout=30.0)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        cursor.execute("PRAGMA journal_mode = WAL;")
        cursor.execute("PRAGMA synchronous = NORMAL;")
        cursor.execute("PRAGMA cache_size = -64000;")  # 64MB cache
        cursor.execute("PRAGMA foreign_keys = ON;")
        cursor.close()
        return conn

    def initialize_db(self):
        """Khởi tạo cấu trúc bảng và 4 chỉ mục chuẩn nếu chưa tồn tại."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            
            # 1. Bảng hoc_sinh (28 cột chuẩn + 6 cột hành chính + 3 cột vết nguồn gốc)
            cursor.execute("""
            CREATE TABLE IF NOT EXISTS hoc_sinh (
                id INTEGER PRIMARY KEY AUTOINCREMENT,

                -- 28 CỘT BIỂU MẪU CHUẨN (CỘT 1 ĐẾN 28)
                stt INTEGER,                                -- Cột 1 (A)
                so_hieu_bang TEXT,                          -- Cột 2 (B)
                so_vao_so TEXT,                             -- Cột 3 (C)
                ten_van_bang TEXT,                          -- Cột 4 (D)
                ho_ten TEXT NOT NULL,                       -- Cột 5 (E)
                ma_nguoi_hoc TEXT,                          -- Cột 6 (F)
                cccd TEXT,                                  -- Cột 7 (G)
                ho_chieu TEXT,                              -- Cột 8 (H)
                ngay_sinh TEXT,                             -- Cột 9 (I)
                gioi_tinh TEXT,                             -- Cột 10 (J)
                dan_toc TEXT,                               -- Cột 11 (K)
                quoc_tich TEXT,                             -- Cột 12 (L)
                noi_sinh TEXT,                              -- Cột 13 (M)
                ten_truong TEXT,                            -- Cột 14 (N)
                nien_khoa TEXT,                             -- Cột 15 (O)
                diem_thi TEXT,                              -- Cột 16 (P)
                hoi_dong_thi TEXT,                          -- Cột 17 (Q)
                dia_danh TEXT,                              -- Cột 18 (R)
                ngay_cap_bang TEXT,                         -- Cột 19 (S)
                co_quan_cap_bang TEXT,                      -- Cột 20 (T)
                chuc_danh_nguoi_ky TEXT,                    -- Cột 21 (U)
                nguoi_ky TEXT,                              -- Cột 22 (V)
                tinh_trang_van_bang TEXT,                   -- Cột 23 (W)
                ghi_chu TEXT,                               -- Cột 24 (X)
                trang_thai TEXT,                            -- Cột 25 (Y)
                ly_do_loi TEXT,                             -- Cột 26 (Z)
                ma_don_vi TEXT,                             -- Cột 27 (AA)
                ten_don_vi TEXT,                            -- Cột 28 (AB)

                -- THÔNG TIN ĐỊA DANH HÀNH CHÍNH (TỪ DS_Don_Vi)
                ma_truong TEXT,                             -- Mã trường
                ma_xa_phuong TEXT,                          -- Mã xã phường
                xa_phuong TEXT,                             -- Tên Xã/Phường
                ma_huyen_cu TEXT,                           -- Mã huyện cũ
                huyen_cu TEXT,                              -- Tên Huyện cũ
                tinh_cu TEXT,                               -- Tên Tỉnh cũ

                -- TRUY VẾT NGUỒN GỐC & QUẢN TRỊ
                source_file TEXT,                           -- Tệp nguồn gốc (đường dẫn tương đối)
                source_sheet TEXT,                          -- Tên Sheet nguồn
                source_row INTEGER,                         -- Dòng Excel gốc
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            );
            """)

            # 2. 11 CHỈ MỤC TỐI ƯU TOÀN DIỆN CHO PHÂN HỆ 2
            # Nhóm chỉ mục đơn tường minh (Tối ưu tuyệt đối khi tìm/lọc từng ô riêng lẻ)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ho_ten ON hoc_sinh(ho_ten);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ngay_sinh ON hoc_sinh(ngay_sinh);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ten_truong ON hoc_sinh(ten_truong);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_nien_khoa ON hoc_sinh(nien_khoa);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ten_don_vi ON hoc_sinh(ten_don_vi);")

            # Nhóm chỉ mục kết hợp (Tối ưu siêu tốc khi tìm/lọc phối hợp nhiều ô)
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_ho_ten_ngay_sinh ON hoc_sinh(ho_ten, ngay_sinh);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_truong_nien_khoa ON hoc_sinh(ten_truong, nien_khoa);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_dia_ban ON hoc_sinh(tinh_cu, huyen_cu);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_xa_don_vi ON hoc_sinh(xa_phuong, ten_don_vi);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_trang_thai_loi ON hoc_sinh(trang_thai, ly_do_loi);")
            cursor.execute("CREATE INDEX IF NOT EXISTS idx_source_file_sheet ON hoc_sinh(source_file, source_sheet);")

            conn.commit()

    def backup_database(self, backup_dir: Path | str | None = None, max_keep: int = 5) -> Path:
        """
        Thực hiện sao lưu trực tuyến (Online Backup) an toàn trước khi nạp dữ liệu lớn.
        Tự động xoay vòng giữ lại tối đa `max_keep` bản sao lưu gần nhất.
        """
        if backup_dir is None:
            backup_path_dir = self.db_path.parent / "backups"
        else:
            backup_path_dir = Path(backup_dir)
            
        backup_path_dir.mkdir(parents=True, exist_ok=True)
        now_str = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_file = backup_path_dir / f"vbcc_backup_{now_str}.db"

        # Sử dụng SQLite Online Backup API
        src_conn = self.get_connection()
        try:
            dest_conn = sqlite3.connect(str(backup_file))
            try:
                src_conn.backup(dest_conn)
            finally:
                dest_conn.close()
        finally:
            src_conn.close()

        # Dọn dẹp các bản sao lưu cũ quá hạn
        try:
            all_backups = sorted(
                backup_path_dir.glob("vbcc_backup_*.db"),
                key=lambda p: p.stat().st_mtime,
                reverse=True
            )
            if len(all_backups) > max_keep:
                for old_f in all_backups[max_keep:]:
                    try:
                        old_f.unlink()
                    except OSError:
                        pass
        except Exception:
            pass

        return backup_file

    def restore_database(self, backup_path: Path | str) -> bool:
        """Khôi phục CSDL từ một bản sao lưu cụ thể."""
        b_path = Path(backup_path)
        if not b_path.is_file():
            return False

        # Thực hiện sao lưu bản sao lưu vào tệp hiện tại bằng backup API
        src_conn = sqlite3.connect(str(b_path))
        try:
            dest_conn = self.get_connection()
            try:
                src_conn.backup(dest_conn)
            finally:
                dest_conn.close()
        finally:
            src_conn.close()

        return True

    def check_existing_sheets(self, source_files: list[str]) -> dict[tuple[str, str], int]:
        """
        Kiểm tra nhanh các sheet đã tồn tại trong CSDL cho danh sách các file nguồn.
        Trả về: {(source_file, source_sheet): số_lượng_học_sinh}
        """
        if not source_files:
            return {}

        res = {}
        with self.get_connection() as conn:
            cursor = conn.cursor()
            # Chia nhỏ theo batch nếu số lượng file quá lớn
            batch_size = 500
            for i in range(0, len(source_files), batch_size):
                sub_files = source_files[i:i + batch_size]
                placeholders = ",".join("?" for _ in sub_files)
                sql = f"""
                    SELECT source_file, source_sheet, COUNT(*) as cnt
                    FROM hoc_sinh
                    WHERE source_file IN ({placeholders})
                    GROUP BY source_file, source_sheet
                """
                cursor.execute(sql, sub_files)
                for row in cursor.fetchall():
                    res[(row["source_file"], row["source_sheet"])] = row["cnt"]
        return res

    def delete_sheet_records(self, source_file: str, source_sheet: str, cursor: sqlite3.Cursor | None = None) -> int:
        """Xóa toàn bộ bản ghi của một sheet nguồn cụ thể."""
        sql = "DELETE FROM hoc_sinh WHERE source_file = ? AND source_sheet = ?"
        if cursor is not None:
            cursor.execute(sql, (source_file, source_sheet))
            return cursor.rowcount
        else:
            with self.get_connection() as conn:
                cur = conn.cursor()
                cur.execute(sql, (source_file, source_sheet))
                conn.commit()
                return cur.rowcount

    def insert_students_batch(
        self,
        students: list[dict],
        overwrite_sheets: set[tuple[str, str]] | None = None,
        progress_callback: Callable[[int, int, str], None] | None = None
    ) -> dict[str, int]:
        """
        Nạp toàn bộ danh sách học sinh đã chuẩn hóa vào SQLite trong Transaction duy nhất.
        - Nếu có `overwrite_sheets`: Xóa toàn bộ bản ghi cũ của các sheet này trước khi chèn mới.
        - Tự động map 28 cột chuẩn + 6 cột hành chính + 3 cột vết nguồn gốc.
        """
        if not students:
            return {"inserted": 0, "deleted_old": 0}

        overwrite_sheets = overwrite_sheets or set()
        deleted_count = 0
        total_students = len(students)

        with self.get_connection() as conn:
            cursor = conn.cursor()

            # 1. Thực hiện xóa các sheet cần ghi đè
            for src_f, src_s in overwrite_sheets:
                cursor.execute("DELETE FROM hoc_sinh WHERE source_file = ? AND source_sheet = ?", (src_f, src_s))
                deleted_count += cursor.rowcount

            # 2. Chuẩn bị dữ liệu chèn hàng loạt (executemany)
            insert_sql = """
            INSERT INTO hoc_sinh (
                stt, so_hieu_bang, so_vao_so, ten_van_bang, ho_ten,
                ma_nguoi_hoc, cccd, ho_chieu, ngay_sinh, gioi_tinh,
                dan_toc, quoc_tich, noi_sinh, ten_truong, nien_khoa,
                diem_thi, hoi_dong_thi, dia_danh, ngay_cap_bang, co_quan_cap_bang,
                chuc_danh_nguoi_ky, nguoi_ky, tinh_trang_van_bang, ghi_chu, trang_thai,
                ly_do_loi, ma_don_vi, ten_don_vi,
                ma_truong, ma_xa_phuong, xa_phuong, ma_huyen_cu, huyen_cu, tinh_cu,
                source_file, source_sheet, source_row
            ) VALUES (
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?, ?, ?,
                ?, ?, ?,
                ?, ?, ?, ?, ?, ?,
                ?, ?, ?
            );
            """

            batch_data = []
            for idx, st in enumerate(students, start=1):
                # final_row đảm bảo có 28 phần tử từ Cột 1 đến Cột 28
                final_row = st.get("final_row") or st.get("row_values") or []
                padded = list(final_row)
                while len(padded) < 28:
                    padded.append(None)

                # Ép STT liên tục
                padded[0] = idx

                # 6 Cột hành chính
                ma_truong = st.get("school_code") or ""
                ma_xa_phuong = st.get("commune_code") or ""
                xa_phuong = st.get("commune") or ""
                ma_huyen_cu = st.get("district_code") or ""
                huyen_cu = st.get("district_name") or ""
                tinh_cu = st.get("province_name") or ""

                # 3 Cột nguồn gốc
                source_file = st.get("source_file") or st.get("file_name") or ""
                source_sheet = st.get("source_sheet") or st.get("sheet_name") or ""
                source_row = st.get("row_idx", 0)

                record = (
                    padded[0], padded[1], padded[2], padded[3], padded[4],
                    padded[5], padded[6], padded[7], padded[8], padded[9],
                    padded[10], padded[11], padded[12], padded[13], padded[14],
                    padded[15], padded[16], padded[17], padded[18], padded[19],
                    padded[20], padded[21], padded[22], padded[23], padded[24],
                    padded[25], padded[26], padded[27],
                    ma_truong, ma_xa_phuong, xa_phuong, ma_huyen_cu, huyen_cu, tinh_cu,
                    source_file, source_sheet, source_row
                )
                batch_data.append(record)

            # Thực thi chèn siêu tốc theo chunk (mỗi chunk 5000 dòng để giải phóng RAM)
            chunk_size = 5000
            for start_idx in range(0, total_students, chunk_size):
                chunk = batch_data[start_idx:start_idx + chunk_size]
                cursor.executemany(insert_sql, chunk)
                if progress_callback:
                    cur_done = min(start_idx + chunk_size, total_students)
                    progress_callback(cur_done, total_students, f"Đang lưu vào CSDL SQLite: {cur_done}/{total_students} học sinh...")

            conn.commit()

        return {
            "inserted": total_students,
            "deleted_old": deleted_count
        }

    def get_statistics(self) -> dict[str, Any]:
        """Lấy số liệu thống kê tổng hợp từ CSDL SQLite."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM hoc_sinh;")
            total = cursor.fetchone()[0]

            cursor.execute("""
                SELECT trang_thai, COUNT(*) as cnt
                FROM hoc_sinh
                GROUP BY trang_thai;
            """)
            status_map = {row["trang_thai"]: row["cnt"] for row in cursor.fetchall()}

            cursor.execute("SELECT COUNT(DISTINCT ten_truong) FROM hoc_sinh WHERE ten_truong IS NOT NULL AND ten_truong != '';")
            total_schools = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(DISTINCT huyen_cu) FROM hoc_sinh WHERE huyen_cu IS NOT NULL AND huyen_cu != '';")
            total_districts = cursor.fetchone()[0]

            cursor.execute("SELECT COUNT(DISTINCT source_file) FROM hoc_sinh;")
            total_files = cursor.fetchone()[0]

            return {
                "total": total,
                "da_so_hoa": status_map.get("Đã số hóa đầy đủ, chính xác", 0),
                "chua_day_du": status_map.get("Số hoá chưa đầy đủ, chính xác", 0),
                "chua_so_hoa": status_map.get("Chưa số hoá", 0) + status_map.get("Chưa số hóa", 0),
                "total_schools": total_schools,
                "total_districts": total_districts,
                "total_files": total_files
            }

    def get_all_districts(self) -> list[str]:
        """Lấy danh sách các Huyện cũ có trong CSDL."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT huyen_cu FROM hoc_sinh WHERE huyen_cu IS NOT NULL AND huyen_cu != '' ORDER BY huyen_cu;")
            return [row[0] for row in cursor.fetchall()]

    def get_all_schools_in_db(self) -> list[str]:
        """Lấy danh sách các Trường đã có trong CSDL."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT DISTINCT ten_truong FROM hoc_sinh WHERE ten_truong IS NOT NULL AND ten_truong != '' ORDER BY ten_truong;")
            return [row[0] for row in cursor.fetchall()]

    def get_cascade_filter_options(
        self,
        tinh_cu: str | None = None,
        huyen_cu: str | None = None,
        xa_phuong: str | None = None
    ) -> dict[str, list[str]]:
        """
        Lấy các danh mục động từ CSDL SQLite phục vụ Bộ lọc đa tiêu chí 2 nhánh:
        - Nhánh 1: Tỉnh cũ -> Huyện cũ -> Tên trường, Niên khóa
        - Nhánh 2: Xã/phường mới -> Đơn vị quản lý
        - Chung: Trạng thái số hóa
        """
        with self.get_connection() as conn:
            cursor = conn.cursor()

            # 1. Tỉnh cũ
            cursor.execute("SELECT DISTINCT tinh_cu FROM hoc_sinh WHERE tinh_cu IS NOT NULL AND tinh_cu != '' ORDER BY tinh_cu;")
            tinh_cu_list = [r[0] for r in cursor.fetchall()]

            # 2. Huyện cũ (lọc theo Tỉnh cũ nếu có)
            if tinh_cu and tinh_cu != "Tất cả":
                cursor.execute("SELECT DISTINCT huyen_cu FROM hoc_sinh WHERE tinh_cu = ? AND huyen_cu IS NOT NULL AND huyen_cu != '' ORDER BY huyen_cu;", (tinh_cu,))
            else:
                cursor.execute("SELECT DISTINCT huyen_cu FROM hoc_sinh WHERE huyen_cu IS NOT NULL AND huyen_cu != '' ORDER BY huyen_cu;")
            huyen_cu_list = [r[0] for r in cursor.fetchall()]

            # 3. Tên trường (lọc theo Huyện cũ hoặc Tỉnh cũ nếu có)
            if huyen_cu and huyen_cu != "Tất cả":
                cursor.execute("SELECT DISTINCT ten_truong FROM hoc_sinh WHERE huyen_cu = ? AND ten_truong IS NOT NULL AND ten_truong != '' ORDER BY ten_truong;", (huyen_cu,))
            elif tinh_cu and tinh_cu != "Tất cả":
                cursor.execute("SELECT DISTINCT ten_truong FROM hoc_sinh WHERE tinh_cu = ? AND ten_truong IS NOT NULL AND ten_truong != '' ORDER BY ten_truong;", (tinh_cu,))
            else:
                cursor.execute("SELECT DISTINCT ten_truong FROM hoc_sinh WHERE ten_truong IS NOT NULL AND ten_truong != '' ORDER BY ten_truong;")
            schools_list = [r[0] for r in cursor.fetchall()]

            # 4. Niên khóa
            cursor.execute("SELECT DISTINCT nien_khoa FROM hoc_sinh WHERE nien_khoa IS NOT NULL AND nien_khoa != '' ORDER BY nien_khoa DESC;")
            nien_khoa_list = [r[0] for r in cursor.fetchall()]

            # 5. Xã/phường mới
            cursor.execute("SELECT DISTINCT xa_phuong FROM hoc_sinh WHERE xa_phuong IS NOT NULL AND xa_phuong != '' ORDER BY xa_phuong;")
            xa_phuong_list = [r[0] for r in cursor.fetchall()]

            # 6. Đơn vị quản lý (lọc theo Xã/phường mới nếu có)
            if xa_phuong and xa_phuong != "Tất cả":
                cursor.execute("SELECT DISTINCT ten_don_vi FROM hoc_sinh WHERE xa_phuong = ? AND ten_don_vi IS NOT NULL AND ten_don_vi != '' ORDER BY ten_don_vi;", (xa_phuong,))
            else:
                cursor.execute("SELECT DISTINCT ten_don_vi FROM hoc_sinh WHERE ten_don_vi IS NOT NULL AND ten_don_vi != '' ORDER BY ten_don_vi;")
            don_vi_list = [r[0] for r in cursor.fetchall()]

            # 7. Trạng thái số hóa
            cursor.execute("SELECT DISTINCT trang_thai FROM hoc_sinh WHERE trang_thai IS NOT NULL AND trang_thai != '' ORDER BY trang_thai;")
            trang_thai_list = [r[0] for r in cursor.fetchall()]

            return {
                "tinh_cu": tinh_cu_list,
                "huyen_cu": huyen_cu_list,
                "ten_truong": schools_list,
                "nien_khoa": nien_khoa_list,
                "xa_phuong": xa_phuong_list,
                "ten_don_vi": don_vi_list,
                "trang_thai": trang_thai_list
            }

    def query_students_paginated(
        self,
        search_name: str = "",
        search_dob: str = "",
        filters: dict[str, Any] | None = None,
        page: int = 1,
        page_size: int = 50,
        sort_by: str = "id",
        sort_order: str = "ASC"
    ) -> tuple[list[dict], int]:
        """
        Truy vấn danh sách học sinh theo Thanh tìm kiếm tức thì (Họ tên, Ngày sinh)
        và Bộ lọc đa tiêu chí 2 nhánh, hỗ trợ phân trang SQL (LIMIT/OFFSET) siêu tốc.
        Trả về: (danh_sách_học_sinh, tổng_số_bản_ghi_khớp)
        """
        filters = filters or {}
        conditions = ["1=1"]
        params = []

        # 1. Thanh tìm kiếm tức thì (Cách 2: 2 ô Họ và tên & Ngày sinh)
        s_name = (search_name or "").strip()
        if s_name:
            conditions.append("ho_ten LIKE ?")
            params.append(f"%{s_name}%")

        s_dob = (search_dob or "").strip()
        if s_dob:
            conditions.append("ngay_sinh LIKE ?")
            params.append(f"%{s_dob}%")

        # 2. Bộ lọc Nhánh 1 (Lịch sử: Tỉnh cũ -> Huyện cũ -> Tên trường -> Niên khóa)
        tinh_cu = filters.get("tinh_cu")
        if tinh_cu and tinh_cu != "Tất cả":
            conditions.append("tinh_cu = ?")
            params.append(tinh_cu)

        huyen_cu = filters.get("huyen_cu")
        if huyen_cu and huyen_cu != "Tất cả":
            conditions.append("huyen_cu = ?")
            params.append(huyen_cu)

        ten_truong = filters.get("ten_truong")
        if ten_truong and ten_truong != "Tất cả":
            conditions.append("ten_truong = ?")
            params.append(ten_truong)

        nien_khoa = filters.get("nien_khoa")
        if nien_khoa and nien_khoa != "Tất cả":
            conditions.append("nien_khoa = ?")
            params.append(nien_khoa)

        # 3. Bộ lọc Nhánh 2 (Địa bàn mới: Xã/phường mới -> Tên Đơn vị)
        xa_phuong = filters.get("xa_phuong")
        if xa_phuong and xa_phuong != "Tất cả":
            conditions.append("xa_phuong = ?")
            params.append(xa_phuong)

        ten_don_vi = filters.get("ten_don_vi")
        if ten_don_vi and ten_don_vi != "Tất cả":
            conditions.append("ten_don_vi = ?")
            params.append(ten_don_vi)

        # 4. Trạng thái & Lỗi
        trang_thai = filters.get("trang_thai")
        if trang_thai and trang_thai != "Tất cả":
            conditions.append("trang_thai = ?")
            params.append(trang_thai)

        chi_xem_loi = filters.get("chi_xem_loi", False)
        if chi_xem_loi:
            conditions.append("(ly_do_loi IS NOT NULL AND ly_do_loi != '')")

        where_clause = " AND ".join(conditions)

        # Bảo vệ SQL Injection cho ORDER BY
        valid_cols = {
            "id", "stt", "so_hieu_bang", "so_vao_so", "ho_ten", "ngay_sinh",
            "gioi_tinh", "ten_truong", "nien_khoa", "trang_thai", "ly_do_loi",
            "tinh_cu", "huyen_cu", "xa_phuong", "ten_don_vi"
        }
        order_col = sort_by if sort_by in valid_cols else "id"
        order_dir = "DESC" if str(sort_order).upper() == "DESC" else "ASC"

        page = max(1, page)
        page_size = max(1, min(page_size, 500))
        offset = (page - 1) * page_size

        with self.get_connection() as conn:
            cursor = conn.cursor()

            # Đếm tổng số bản ghi khớp
            count_sql = f"SELECT COUNT(*) FROM hoc_sinh WHERE {where_clause};"
            cursor.execute(count_sql, params)
            total_count = cursor.fetchone()[0]

            # Lấy dữ liệu trang
            data_sql = f"""
                SELECT * FROM hoc_sinh
                WHERE {where_clause}
                ORDER BY {order_col} {order_dir}
                LIMIT ? OFFSET ?;
            """
            cursor.execute(data_sql, params + [page_size, offset])
            rows = [dict(row) for row in cursor.fetchall()]

            return rows, total_count

    def get_student_by_id(self, student_id: int) -> dict | None:
        """Lấy toàn bộ thông tin chi tiết của 1 học sinh theo ID."""
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute("SELECT * FROM hoc_sinh WHERE id = ?;", (student_id,))
            row = cursor.fetchone()
            return dict(row) if row else None

    def update_student(
        self,
        student_id: int,
        fields_dict: dict[str, Any],
        lookup_mgr: Any = None
    ) -> bool:
        """
        Cập nhật thông tin của 1 học sinh.
        - Tự động re-map địa bàn hành chính & đơn vị nếu tên trường thay đổi.
        - Tự động re-validate ngày sinh và recompute `trang_thai`, `ly_do_loi`.
        - Cập nhật `updated_at`.
        """
        current = self.get_student_by_id(student_id)
        if not current:
            return False

        updated_data = dict(fields_dict)

        # 1. Nếu thay đổi Tên trường và có lookup_mgr
        new_school = updated_data.get("ten_truong")
        if new_school and lookup_mgr:
            details = lookup_mgr.get_school_details(new_school)
            if details:
                updated_data.setdefault("ma_truong", details.get("school_code") or "")
                updated_data.setdefault("ma_xa_phuong", details.get("commune_code") or "")
                updated_data.setdefault("xa_phuong", details.get("commune") or "")
                updated_data.setdefault("ma_huyen_cu", details.get("district_code") or "")
                updated_data.setdefault("huyen_cu", details.get("district_name") or "")
                updated_data.setdefault("tinh_cu", details.get("province_name") or "")
                updated_data.setdefault("ma_don_vi", details.get("unit_code") or "")
                updated_data.setdefault("ten_don_vi", details.get("unit_name") or "")

                # Fallback địa danh Cột 18 và cơ quan cấp bằng Cột 20 nếu rỗng
                h_cu = updated_data.get("huyen_cu") or details.get("district_name") or ""
                t_cu = updated_data.get("tinh_cu") or details.get("province_name") or ""
                if not updated_data.get("dia_danh") and not current.get("dia_danh") and h_cu and t_cu:
                    updated_data["dia_danh"] = f"{h_cu}-{t_cu}"
                if not updated_data.get("co_quan_cap_bang") and not current.get("co_quan_cap_bang") and h_cu:
                    updated_data["co_quan_cap_bang"] = f"Phòng GD&ĐT {h_cu}"

        # 2. Re-validate Ngày sinh & Recompute Trạng thái / Lý do lỗi
        effective_dob = updated_data.get("ngay_sinh", current.get("ngay_sinh", ""))
        is_valid, err_type, err_desc = check_dob_validity(effective_dob)
        status, reason = determine_status_and_reason(effective_dob, effective_dob, is_valid, err_desc)
        updated_data["trang_thai"] = status
        updated_data["ly_do_loi"] = reason

        # 3. Mốc thời gian
        updated_data["updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # 4. Thực thi UPDATE
        set_clauses = [f"{k} = ?" for k in updated_data.keys()]
        values = list(updated_data.values()) + [student_id]
        update_sql = f"UPDATE hoc_sinh SET {', '.join(set_clauses)} WHERE id = ?;"

        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(update_sql, values)
            conn.commit()
            return cursor.rowcount > 0

    def batch_update_students(
        self,
        student_ids: list[int],
        updates: dict[str, Any],
        lookup_mgr: Any = None
    ) -> int:
        """
        Cập nhật hàng loạt danh sách học sinh theo các trường được chỉ định.
        - Chạy trong 1 Transaction duy nhất (ACID).
        - Tự động đồng bộ các trường hành chính nếu đổi Tên trường.
        - Tự động recompute trạng thái/lỗi nếu cập nhật Ngày sinh.
        """
        if not student_ids or not updates:
            return 0

        batch_updates = dict(updates)

        # Nếu cập nhật Tên trường
        new_school = batch_updates.get("ten_truong")
        if new_school and lookup_mgr:
            details = lookup_mgr.get_school_details(new_school)
            if details:
                batch_updates["ma_truong"] = details.get("school_code") or ""
                batch_updates["ma_xa_phuong"] = details.get("commune_code") or ""
                batch_updates["xa_phuong"] = details.get("commune") or ""
                batch_updates["ma_huyen_cu"] = details.get("district_code") or ""
                batch_updates["huyen_cu"] = details.get("district_name") or ""
                batch_updates["tinh_cu"] = details.get("province_name") or ""
                batch_updates["ma_don_vi"] = details.get("unit_code") or ""
                batch_updates["ten_don_vi"] = details.get("unit_name") or ""

        # Nếu cập nhật Ngày sinh đồng loạt
        if "ngay_sinh" in batch_updates:
            new_dob = batch_updates["ngay_sinh"]
            is_valid, err_type, err_desc = check_dob_validity(new_dob)
            status, reason = determine_status_and_reason(new_dob, new_dob, is_valid, err_desc)
            batch_updates["trang_thai"] = status
            batch_updates["ly_do_loi"] = reason

        batch_updates["updated_at"] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        set_clauses = [f"{k} = ?" for k in batch_updates.keys()]
        set_str = ", ".join(set_clauses)
        params_base = list(batch_updates.values())

        updated_count = 0
        with self.get_connection() as conn:
            cursor = conn.cursor()
            # Chia nhỏ theo batch 500 ID
            chunk_size = 500
            for i in range(0, len(student_ids), chunk_size):
                sub_ids = student_ids[i:i + chunk_size]
                placeholders = ",".join("?" for _ in sub_ids)
                sql = f"UPDATE hoc_sinh SET {set_str} WHERE id IN ({placeholders});"
                cursor.execute(sql, params_base + sub_ids)
                updated_count += cursor.rowcount
            conn.commit()

        return updated_count

    def delete_students(self, student_ids: list[int]) -> int:
        """Xóa danh sách học sinh theo ID trong 1 Transaction."""
        if not student_ids:
            return 0
        deleted_count = 0
        with self.get_connection() as conn:
            cursor = conn.cursor()
            chunk_size = 500
            for i in range(0, len(student_ids), chunk_size):
                sub_ids = student_ids[i:i + chunk_size]
                placeholders = ",".join("?" for _ in sub_ids)
                cursor.execute(f"DELETE FROM hoc_sinh WHERE id IN ({placeholders});", sub_ids)
                deleted_count += cursor.rowcount
            conn.commit()
        return deleted_count

    def get_distinct_cohorts_for_filter(
        self,
        tinh_cu: str | None = None,
        huyen_cu: str | None = None,
        ten_truong: str | None = None,
        xa_phuong: str | None = None,
        ten_don_vi: str | None = None
    ) -> list[str]:
        """
        Lấy danh sách các Niên khóa duy nhất theo bộ lọc địa bàn đã chọn.
        Sắp xếp theo thứ tự năm tốt nghiệp tăng dần.
        """
        conditions = ["nien_khoa IS NOT NULL AND nien_khoa != ''"]
        params = []

        if tinh_cu and tinh_cu != "Tất cả":
            conditions.append("tinh_cu = ?")
            params.append(tinh_cu)
        if huyen_cu and huyen_cu != "Tất cả":
            conditions.append("huyen_cu = ?")
            params.append(huyen_cu)
        if ten_truong and ten_truong != "Tất cả":
            conditions.append("ten_truong = ?")
            params.append(ten_truong)
        if xa_phuong and xa_phuong != "Tất cả":
            conditions.append("xa_phuong = ?")
            params.append(xa_phuong)
        if ten_don_vi and ten_don_vi != "Tất cả":
            conditions.append("ten_don_vi = ?")
            params.append(ten_don_vi)

        where_clause = " AND ".join(conditions)
        sql = f"SELECT DISTINCT nien_khoa FROM hoc_sinh WHERE {where_clause};"
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(sql, params)
            rows = [r[0] for r in cursor.fetchall()]

            def get_end_year(nk):
                match = re.findall(r'\b(19\d{2}|20\d{2})\b', str(nk))
                if match:
                    return int(match[-1])
                return 0

            return sorted(rows, key=get_end_year)

    def query_students_for_export(
        self,
        tinh_cu: str | None = None,
        huyen_cu: str | None = None,
        ten_truong: str | None = None,
        xa_phuong: str | None = None,
        ten_don_vi: str | None = None,
        cohorts: list[str] | None = None,
        only_errors: bool = False,
        sort_mode: str = "stt_school_cohort"
    ) -> list[dict]:
        """
        Truy vấn danh sách học sinh phục vụ xuất Excel theo yêu cầu:
        - Lọc theo Huyện cũ / Trường hoặc Xã / Đơn vị.
        - Lọc theo danh sách Niên khóa (1 hoặc nhiều).
        - Lọc riêng hồ sơ lỗi (nếu only_errors=True).
        - Sắp xếp mặc định theo STT -> Trường -> Niên khóa.
        """
        conditions = ["1=1"]
        params = []

        if tinh_cu and tinh_cu != "Tất cả":
            conditions.append("tinh_cu = ?")
            params.append(tinh_cu)
        if huyen_cu and huyen_cu != "Tất cả":
            conditions.append("huyen_cu = ?")
            params.append(huyen_cu)
        if ten_truong and ten_truong != "Tất cả":
            conditions.append("ten_truong = ?")
            params.append(ten_truong)
        if xa_phuong and xa_phuong != "Tất cả":
            conditions.append("xa_phuong = ?")
            params.append(xa_phuong)
        if ten_don_vi and ten_don_vi != "Tất cả":
            conditions.append("ten_don_vi = ?")
            params.append(ten_don_vi)

        if cohorts:
            placeholders = ",".join("?" for _ in cohorts)
            conditions.append(f"nien_khoa IN ({placeholders})")
            params.extend(cohorts)

        if only_errors:
            conditions.append("(trang_thai != 'Đã số hóa đầy đủ, chính xác' OR (ly_do_loi IS NOT NULL AND ly_do_loi != ''))")

        where_clause = " AND ".join(conditions)

        if sort_mode == "stt_school_cohort":
            order_clause = "ORDER BY stt ASC, ten_truong ASC, nien_khoa ASC, id ASC"
        else:
            order_clause = "ORDER BY id ASC"

        sql = f"SELECT * FROM hoc_sinh WHERE {where_clause} {order_clause};"
        with self.get_connection() as conn:
            cursor = conn.cursor()
            cursor.execute(sql, params)
            return [dict(r) for r in cursor.fetchall()]
