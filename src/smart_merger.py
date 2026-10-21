from __future__ import annotations

import os
import re
import json
import hashlib
import unicodedata
import difflib
import functools
import io
from pathlib import Path
from typing import Any, Callable
from copy import copy
from datetime import datetime

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

from src.lookup_manager import find_data_file, get_data_dir, get_persistent_data_dir
from src.normalizer import (
    clean_whitespace,
    is_cell_empty,
    vietnamese_title_case,
    format_date_to_text,
    auto_convert_legacy_encoding
)
from src.custom_excel_manager import (
    is_valid_stt,
    is_valid_name,
    is_header_duplicate_row,
    detect_header_and_columns,
    find_best_data_sheet
)
from src.excel_processor import is_summary_or_footer_row
from src.exporter import sanitize_filename

# ==============================================================================
# DANH SÁCH 28 CỘT CHUẨN CỦA MẪU CẤP BẰNG THCS (Mau_DS_Cap_Bang_THCS.xlsx)
# ==============================================================================

TARGET_COLUMNS = [
    {"idx": 1, "field": "stt", "name": "STT", "required": True, "default": "Tự động đánh số 1..N", "type": "int"},
    {"idx": 2, "field": "so_hieu_bang", "name": "Số hiệu bằng", "required": False, "default": "", "type": "text"},
    {"idx": 3, "field": "so_vao_so", "name": "Số vào sổ gốc cấp văn bằng", "required": False, "default": "", "type": "text"},
    {"idx": 4, "field": "ten_van_bang", "name": "Tên văn bằng", "required": False, "default": "Bằng tốt nghiệp trung học cơ sở", "type": "text"},
    {"idx": 5, "field": "ho_ten", "name": "Họ, chữ đệm và tên", "required": True, "default": "", "type": "text"},
    {"idx": 6, "field": "ma_nguoi_hoc", "name": "Mã người học", "required": False, "default": "", "type": "text"},
    {"idx": 7, "field": "cccd", "name": "Số định danh cá nhân", "required": False, "default": "", "type": "text"},
    {"idx": 8, "field": "ho_chieu", "name": "Hộ chiếu", "required": False, "default": "", "type": "text"},
    {"idx": 9, "field": "ngay_sinh", "name": "Ngày, tháng, năm sinh", "required": False, "default": "", "type": "date"},
    {"idx": 10, "field": "gioi_tinh", "name": "Giới tính", "required": False, "default": "", "type": "text"},
    {"idx": 11, "field": "dan_toc", "name": "Dân tộc", "required": False, "default": "", "type": "text"},
    {"idx": 12, "field": "quoc_tich", "name": "Quốc tịch", "required": False, "default": "Việt Nam", "type": "text"},
    {"idx": 13, "field": "noi_sinh", "name": "Nơi sinh", "required": False, "default": "", "type": "text"},
    {"idx": 14, "field": "ten_truong", "name": "Tên trường đã học", "required": False, "default": "", "type": "text"},
    {"idx": 15, "field": "nien_khoa", "name": "Niên khóa", "required": False, "default": "", "type": "text"},
    {"idx": 16, "field": "diem_thi", "name": "Điểm các môn thi", "required": False, "default": "", "type": "text"},
    {"idx": 17, "field": "hoi_dong_thi", "name": "Hội đồng thi", "required": False, "default": "", "type": "text"},
    {"idx": 18, "field": "dia_danh", "name": "Địa danh nơi cơ quan cấp bằng đặt trụ sở", "required": False, "default": "", "type": "text"},
    {"idx": 19, "field": "ngay_cap_bang", "name": "Ngày tháng năm cấp bằng", "required": False, "default": "", "type": "date"},
    {"idx": 20, "field": "co_quan_cap_bang", "name": "Tên cơ quan cấp bằng", "required": False, "default": "", "type": "text"},
    {"idx": 21, "field": "chuc_danh_nguoi_ky", "name": "Chức danh người ký bằng", "required": False, "default": "", "type": "text"},
    {"idx": 22, "field": "nguoi_ky", "name": "Họ, chữ đệm, tên người ký bằng", "required": False, "default": "", "type": "text"},
    {"idx": 23, "field": "tinh_trang_van_bang", "name": "Tình trạng văn bằng", "required": False, "default": "", "type": "text"},
    {"idx": 24, "field": "ghi_chu", "name": "Ghi chú", "required": False, "default": "", "type": "text"},
    {"idx": 25, "field": "trang_thai", "name": "Trạng thái số hóa", "required": False, "default": "", "type": "text"},
    {"idx": 26, "field": "ly_do_loi", "name": "Thông tin chưa bảo đảm đầy đủ, chính xác và nguyên nhân", "required": False, "default": "", "type": "text"},
    {"idx": 27, "field": "ma_don_vi", "name": "Mã đơn vị", "required": False, "default": "", "type": "text"},
    {"idx": 28, "field": "ten_don_vi", "name": "Tên đơn vị", "required": False, "default": "", "type": "text"}
]

# Các cột cần định dạng Text @ trong openpyxl
TEXT_COLS_28 = {2, 3, 6, 7, 8, 9, 19, 24, 27}

# ==============================================================================
# TỪ ĐIỂN ĐỒNG NGHĨA & TỪ KHÓA NGHIỆP VỤ CHO 28 CỘT (TẦNG 1)
# ==============================================================================

SYNONYMS_MAP: dict[int, list[str]] = {
    1: ["stt", "số tt", "số thứ tự", "thứ tự", "tt", "stt."],
    2: ["số hiệu bằng", "số hiệu", "so hieu bang", "số bằng", "so hieu", "so bang"],
    3: ["số vào sổ gốc cấp văn bằng", "số vào sổ gốc", "số vào sổ", "số sổ gốc", "vào sổ số", "số sổ", "so vao so", "so so goc", "số sổ cấp bằng", "vào sổ"],
    4: ["tên văn bằng", "loại văn bằng", "loại bằng", "văn bằng", "loại tốt nghiệp", "xếp loại tốt nghiệp", "bằng tốt nghiệp"],
    5: ["họ, chữ đệm và tên", "họ và tên", "họ tên", "họ và chữ đệm", "tên học sinh", "họ tên học sinh", "người học", "thí sinh", "họ, tên", "họ và tên học sinh", "tên thí sinh"],
    6: ["mã người học", "mã học sinh", "mã định danh học sinh", "mã hs", "sbd", "số báo danh"],
    7: ["số định danh cá nhân", "số định danh", "định danh cá nhân", "cccd", "căn cước", "căn cước công dân", "số cccd", "cmnd", "số cmnd"],
    8: ["hộ chiếu", "passport", "số hộ chiếu"],
    9: ["ngày, tháng, năm sinh", "ngày tháng năm sinh", "ngày sinh", "ngay sinh", "năm sinh", "ngay thang nam sinh", "ns", "ngaysinh"],
    10: ["giới tính", "phái", "nam/nữ", "nam, nữ", "nam nữ", "gioi tinh", "nam", "nữ"],
    11: ["dân tộc", "dan toc", "dt"],
    12: ["quốc tịch", "quoc tich"],
    13: ["nơi sinh", "noi sinh", "quê quán", "nguyên quán", "que quan", "nơi cư trú", "hộ khẩu", "hộ khẩu thường trú"],
    14: ["tên trường đã học", "tên trường", "trường học", "trường thcs", "trường", "ten truong", "tên trường tốt nghiệp", "đơn vị trường"],
    15: ["niên khóa", "khóa học", "năm học", "nien khoa", "khoa hoc", "năm tốt nghiệp", "khóa tốt nghiệp", "niên khoá"],
    16: ["điểm các môn thi", "điểm thi", "kết quả thi", "điểm", "diem thi", "điểm trung bình", "điểm tn", "xếp loại học lực", "học lực"],
    17: ["hội đồng thi", "hội đồng coi thi", "hội đồng chấm thi", "hđ thi"],
    18: ["địa danh nơi cơ quan cấp bằng đặt trụ sở", "địa danh nơi cơ quan", "địa danh", "nơi cấp", "địa phương"],
    19: ["ngày tháng năm cấp bằng", "ngày cấp bằng", "ngày ký bằng", "ngày cấp", "ngay cap", "ngày ra quyết định", "ngày tốt nghiệp"],
    20: ["tên cơ quan cấp bằng", "cơ quan cấp bằng", "đơn vị cấp bằng", "phòng gd&đt", "sở gd&đt", "phòng giáo dục"],
    21: ["chức danh người ký bằng", "chức danh người ký", "chức danh", "chức vụ", "chức vụ người ký"],
    22: ["họ, chữ đệm, tên người ký bằng", "họ, chữ đệm, tên người ký", "người ký bằng", "người ký", "họ tên người ký", "hiệu trưởng", "thủ trưởng", "người ký quyết định"],
    23: ["tình trạng văn bằng", "tình trạng bằng", "tình trạng"],
    24: ["ghi chú", "chú thích", "ghi chu", "note"],
    25: ["trạng thái số hóa", "trạng thái", "tình trạng số hóa"],
    26: ["thông tin chưa bảo đảm đầy đủ, chính xác và nguyên nhân", "thông tin chưa bảo đảm", "nguyên nhân", "lý do chưa bảo đảm", "lý do lỗi"],
    27: ["mã đơn vị", "ma don vi", "mã phòng", "mã trường", "mã sở"],
    28: ["tên đơn vị", "ten don vi", "đơn vị", "tên phòng giáo dục", "phòng gd"]
}

# ==============================================================================
# HÀM SO KHỚP CHUỖI & TỰ ĐỘNG GỢI Ý ÁNH XẠ CỘT
# ==============================================================================

RE_NON_ALPHANUM = re.compile(r'[^a-z0-9\s]')
RE_SPACES = re.compile(r'\s+')

@functools.lru_cache(maxsize=4096)
def remove_accents(input_str: str) -> str:
    """Loại bỏ dấu tiếng Việt để so khớp tương đối."""
    if not input_str:
        return ""
    nfkd_form = unicodedata.normalize('NFKD', str(input_str))
    return "".join([c for c in nfkd_form if not unicodedata.combining(c)])

@functools.lru_cache(maxsize=4096)
def normalize_text_for_match(s: str) -> str:
    """Chuẩn hóa chuỗi text cho việc so khớp mờ."""
    if s is None:
        return ""
    s_conv = auto_convert_legacy_encoding(str(s))
    s_no_acc = remove_accents(s_conv).lower()
    s_clean = RE_NON_ALPHANUM.sub(' ', s_no_acc)
    return RE_SPACES.sub(' ', s_clean).strip()

def compute_similarity(target_name: str, synonyms: list[str], source_header: str) -> tuple[float, str]:
    """
    Tính điểm tương đồng giữa một cột nguồn và một cột đích:
    - Trả về: (score: 0.0 -> 1.0, match_type: 'exact' | 'keyword' | 'fuzzy' | 'none')
    """
    src_norm = normalize_text_for_match(source_header)
    if not src_norm:
        return 0.0, "none"

    tgt_norm = normalize_text_for_match(target_name)

    # 1. Khớp chính xác tuyệt đối với tên cột đích
    if src_norm == tgt_norm:
        return 1.0, "exact"

    # 2. Khớp tuyệt đối với một trong các từ khóa đồng nghĩa (Tầng 1)
    syn_norms = [normalize_text_for_match(syn) for syn in synonyms]
    for syn_n in syn_norms:
        if src_norm == syn_n:
            return 0.98, "keyword"

    # 3. Khớp chuỗi con chứa trọn vẹn từ khóa đồng nghĩa
    for syn_n in syn_norms:
        if len(syn_n) >= 3 and (syn_n in src_norm or src_norm in syn_n):
            ratio_len = min(len(syn_n), len(src_norm)) / max(len(syn_n), len(src_norm))
            if ratio_len >= 0.5:
                return 0.85 + 0.12 * ratio_len, "keyword"

    # 4. So khớp mờ qua Token Set và SequenceMatcher (Tầng 2)
    src_tokens = set(src_norm.split())
    best_fuzzy = 0.0

    all_candidates = [tgt_norm] + syn_norms
    for cand in all_candidates:
        cand_tokens = set(cand.split())
        if src_tokens and cand_tokens:
            inter = src_tokens.intersection(cand_tokens)
            if inter:
                jaccard = len(inter) / len(src_tokens.union(cand_tokens))
                best_fuzzy = max(best_fuzzy, jaccard * 0.9)

        seq_score = difflib.SequenceMatcher(None, src_norm, cand).ratio()
        best_fuzzy = max(best_fuzzy, seq_score)

    if best_fuzzy >= 0.60:
        return round(best_fuzzy, 3), "fuzzy"

    return round(best_fuzzy, 3), "none"

def compute_headers_signature(headers: list[str]) -> str:
    """Tạo chữ ký MD5 hash duy nhất từ danh sách tiêu đề cột nguồn để lưu nhận diện Profile."""
    clean_headers = [normalize_text_for_match(h) for h in headers if str(h or '').strip()]
    raw_sig = "|".join(clean_headers)
    return hashlib.sha256(raw_sig.encode('utf-8')).hexdigest()[:16]

def is_standard_28_headers(headers: list[str]) -> bool:
    """Kiểm tra xem danh sách header có phải chính là mẫu chuẩn 28 cột hay không."""
    if len(headers) < 28:
        return False
    matches = 0
    for idx, col_def in enumerate(TARGET_COLUMNS):
        if idx < len(headers):
            score, _ = compute_similarity(col_def["name"], SYNONYMS_MAP.get(col_def["idx"], []), headers[idx])
            if score >= 0.85:
                matches += 1
    return matches >= 20

def auto_suggest_mapping(source_headers: list[str]) -> dict[int, int | None]:
    """
    Gợi ý ánh xạ thông minh từ 28 cột đích sang chỉ số cột nguồn (0-based) hoặc None.
    Thuật toán đa tầng kết hợp từ điển đồng nghĩa và so khớp mờ, chống gán trùng lặp.
    """
    mapping: dict[int, int | None] = {col["idx"]: None for col in TARGET_COLUMNS}
    if not source_headers:
        return mapping

    # Nếu file nguồn đã là chuẩn 28 cột, ánh xạ 1:1 trực tiếp
    if is_standard_28_headers(source_headers):
        for idx, col in enumerate(TARGET_COLUMNS):
            if idx < len(source_headers):
                mapping[col["idx"]] = idx
        return mapping

    # Bảng điểm cho từng cặp (target_idx, source_idx)
    candidates: list[tuple[float, int, int, str]] = [] # (score, target_idx, source_col_idx, match_type)

    for col in TARGET_COLUMNS:
        t_idx = col["idx"]
        t_name = col["name"]
        syns = SYNONYMS_MAP.get(t_idx, [])

        for s_idx, s_header in enumerate(source_headers):
            score, match_type = compute_similarity(t_name, syns, s_header)
            if score >= 0.62:
                candidates.append((score, t_idx, s_idx, match_type))

    # Sắp xếp theo điểm số giảm dần để ưu tiên gán các cặp có độ tin cậy cao nhất trước
    candidates.sort(key=lambda x: x[0], reverse=True)

    used_sources = set()
    used_targets = set()

    for score, t_idx, s_idx, match_type in candidates:
        if t_idx not in used_targets and s_idx not in used_sources:
            mapping[t_idx] = s_idx
            used_targets.add(t_idx)
            used_sources.add(s_idx)

    return mapping


# ==============================================================================
# QUẢN LÝ PROFILE ÁNH XẠ CỘT (MappingProfileManager)
# ==============================================================================

class MappingProfileManager:
    """Quản lý các cấu hình Profile ánh xạ cột đã lưu trong data/mapping_profiles.json."""

    def __init__(self, storage_path: Path | str | None = None):
        if storage_path:
            self.storage_path = Path(storage_path)
        else:
            self.storage_path = get_persistent_data_dir() / "mapping_profiles.json"
        self._ensure_storage()

    def _ensure_storage(self):
        if not self.storage_path.exists():
            self.storage_path.parent.mkdir(parents=True, exist_ok=True)
            try:
                bundled = find_data_file("mapping_profiles.json")
                if bundled.is_file() and bundled.resolve() != self.storage_path.resolve():
                    import shutil
                    shutil.copy2(bundled, self.storage_path)
                    return
            except Exception:
                pass
            self._save_raw({"profiles": []})

    def _load_raw(self) -> dict:
        try:
            if not self.storage_path.exists():
                return {"profiles": []}
            with open(self.storage_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if not isinstance(data, dict) or "profiles" not in data:
                    return {"profiles": []}
                return data
        except Exception:
            return {"profiles": []}

    def _save_raw(self, data: dict):
        with open(self.storage_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def list_profiles(self) -> list[dict]:
        """Lấy toàn bộ danh sách các Profile đang lưu."""
        return self._load_raw().get("profiles", [])

    def get_profile_by_signature(self, signature: str) -> dict | None:
        """Tìm Profile tương ứng với chữ ký header."""
        if not signature:
            return None
        for p in self.list_profiles():
            if p.get("signature") == signature:
                return p
        return None

    def get_profile_by_id(self, profile_id: str) -> dict | None:
        for p in self.list_profiles():
            if p.get("id") == profile_id:
                return p
        return None

    def save_profile(
        self,
        name: str,
        signature: str,
        raw_headers: list[str],
        mapping: dict[int, int | None]
    ) -> dict:
        """
        Lưu hoặc cập nhật một Profile.
        mapping: {1: 0, 5: 1, 9: None, ...}
        """
        data = self._load_raw()
        profiles = data.get("profiles", [])

        # Chuyển key mapping thành string để lưu JSON chuẩn
        str_mapping = {str(k): v for k, v in mapping.items()}

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Kiểm tra xem signature đã tồn tại chưa
        existing = None
        for p in profiles:
            if p.get("signature") == signature:
                existing = p
                break

        if existing:
            existing["name"] = name or existing.get("name", "Profile Không tên")
            existing["raw_headers"] = raw_headers
            existing["mapping"] = str_mapping
            existing["updated_at"] = now_str
            ret_profile = existing
        else:
            profile_id = f"prof_{signature[:8]}_{int(datetime.now().timestamp())}"
            new_profile = {
                "id": profile_id,
                "name": name or f"Mẫu cột {len(raw_headers)} cột ({now_str[:10]})",
                "signature": signature,
                "raw_headers": raw_headers,
                "mapping": str_mapping,
                "created_at": now_str,
                "updated_at": now_str
            }
            profiles.append(new_profile)
            ret_profile = new_profile

        data["profiles"] = profiles
        self._save_raw(data)
        return ret_profile

    def delete_profile(self, profile_id: str) -> bool:
        data = self._load_raw()
        profiles = data.get("profiles", [])
        init_len = len(profiles)
        profiles = [p for p in profiles if p.get("id") != profile_id]
        if len(profiles) < init_len:
            data["profiles"] = profiles
            self._save_raw(data)
            return True
        return False


# ==============================================================================
# SMART MERGER ENGINE: BỘ XỬ LÝ GHÉP & CHUẨN HÓA DỮ LIỆU ĐA TỆP
# ==============================================================================

class SmartMerger:
    """
    Bộ xử lý Ghép và Chuẩn hóa thông minh:
    - Tự động nhận diện cấu trúc dòng tiêu đề của từng tệp/sheet nguồn.
    - So khớp với mẫu chuẩn 28 cột (Mau_DS_Cap_Bang_THCS.xlsx).
    - Tự động áp dụng Profile nếu đã lưu trước đó.
    - Gọi giao diện tương tác nếu gặp cấu trúc cột mới.
    - Tái sử dụng 100% quy tắc lọc dữ liệu học sinh hợp lệ và chuẩn hóa encoding/date/title_case.
    """

    def __init__(self, profile_mgr: MappingProfileManager | None = None):
        self.profile_mgr = profile_mgr or MappingProfileManager()
        self.template_path = find_data_file("Mau_DS_Cap_Bang_THCS.xlsx")
        self.template_bytes: bytes | None = None
        self.cached_fonts: list[Font | None] = []
        self._load_template()

    def _load_template(self):
        if self.template_path.exists():
            with open(self.template_path, "rb") as f:
                self.template_bytes = f.read()
            wb = openpyxl.load_workbook(io.BytesIO(self.template_bytes))
            ws = wb.active
            self.cached_fonts = []
            for col in range(1, 29):
                cell = ws.cell(row=2, column=col)
                if cell.has_style and cell.font:
                    self.cached_fonts.append(Font(name=cell.font.name, size=cell.font.size))
                else:
                    self.cached_fonts.append(None)
            wb.close()

    def create_destination_workbook(self) -> tuple[openpyxl.Workbook, Any]:
        """Tạo workbook đích từ tệp mẫu Mau_DS_Cap_Bang_THCS.xlsx."""
        if self.template_bytes:
            wb = openpyxl.load_workbook(io.BytesIO(self.template_bytes))
            ws = wb.active
            ws.title = "Data"
            if ws.max_row > 2:
                ws.delete_rows(3, ws.max_row - 2)
            return wb, ws
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Data"
            ws.append(["DANH SÁCH BẰNG TỐT NGHIỆP THPT, THCS, TIỂU HỌC"])
            ws.append([col["name"] for col in TARGET_COLUMNS])
            return wb, ws

    def process_merge_and_standardize(
        self,
        source_files: list[Path | str],
        output_file_path: Path | str,
        mapping_callback: Callable[[dict], tuple[dict[int, int | None] | None, str | None, bool]],
        renumber_stt: bool = True,
        fill_defaults: bool = True,
        merge_all_sheets: bool = True,
        progress_callback: Callable[[int, int, str], None] | None = None,
        log_callback: Callable[[str, str], None] | None = None,
        check_cancel: Callable[[], bool] | None = None
    ) -> dict[str, Any]:
        """
        Thực hiện toàn bộ quy trình Ghép và Chuẩn hóa:
        - mapping_callback(info_dict): Được gọi khi gặp sheet mới chưa có profile.
          Trả về: (confirmed_mapping, profile_name_to_save, should_skip_sheet)
        """
        def log(msg, level="info"):
            if log_callback:
                log_callback(msg, level)
            else:
                print(msg)

        if not source_files:
            return {"status": "error", "message": "Chưa chọn tệp nguồn nào để ghép."}

        out_path = Path(output_file_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        wb_dst, ws_dst = self.create_destination_workbook()

        total_extracted_rows = 0
        total_files = len(source_files)
        sheets_processed = 0
        sheets_skipped = 0

        log(f"🚀 BẮT ĐẦU GHÉP & CHUẨN HÓA {total_files} TỆP NGUỒN VÀO MẪU CHUẨN 28 CỘT...", "info")

        for f_idx, f_p in enumerate(source_files, start=1):
            if check_cancel and check_cancel():
                log("⏹ Đã dừng xử lý theo yêu cầu người dùng.", "warning")
                break

            p = Path(f_p)
            if not p.exists():
                log(f"⚠️ Không tìm thấy tệp: {p.name}, bỏ qua.", "warning")
                continue

            try:
                wb_src = openpyxl.load_workbook(p, data_only=True)
            except Exception as e:
                log(f"❌ Lỗi đọc tệp {p.name}: {e}", "error")
                continue

            # Lấy danh sách sheet cần duyệt
            if merge_all_sheets:
                candidate_sheets = [ws for ws in wb_src.worksheets if ws.sheet_state == 'visible']
            else:
                candidate_sheets = [find_best_data_sheet(wb_src)]

            log(f"📂 [{f_idx}/{total_files}] Đang quét tệp: {p.name} ({len(candidate_sheets)} sheet)...", "info")

            for ws_curr in candidate_sheets:
                if check_cancel and check_cancel():
                    break

                # 1. Quét tìm dòng tiêu đề và danh sách nhãn cột của sheet
                header_info = detect_header_and_columns(ws_curr)
                h_idx = header_info.get("header_row_idx", 1)
                raw_headers = header_info.get("raw_headers", [])

                if not raw_headers:
                    continue

                # 2. Tạo chữ ký cấu trúc cột
                sig = compute_headers_signature(raw_headers)
                mapping: dict[int, int | None] | None = None
                profile_used_name: str = ""

                # Kiểm tra xem có khớp chuẩn 28 cột không
                if is_standard_28_headers(raw_headers):
                    mapping = {col["idx"]: idx for idx, col in enumerate(TARGET_COLUMNS) if idx < len(raw_headers)}
                    profile_used_name = "Chuẩn 28 cột gốc"
                    log(f"  ⚡ Sheet '{ws_curr.title}': Cấu trúc chuẩn 28 cột $\\rightarrow$ Ánh xạ tự động 1:1.", "info")
                else:
                    # Kiểm tra trong kho Profile đã lưu
                    saved_prof = self.profile_mgr.get_profile_by_signature(sig)
                    if saved_prof:
                        # Áp dụng Profile đã lưu tự động mà KHÔNG cần hỏi lại
                        str_map = saved_prof.get("mapping", {})
                        mapping = {int(k): v for k, v in str_map.items()}
                        profile_used_name = saved_prof.get("name", "Profile đã lưu")
                        log(f"  ⚡ Sheet '{ws_curr.title}': Áp dụng Profile đã nhớ: '{profile_used_name}'.", "info")
                    else:
                        # Chưa có Profile -> Gợi ý tự động & Kích hoạt giao diện hỏi người dùng
                        suggested_map = auto_suggest_mapping(raw_headers)
                        formatted_headers = header_info.get("formatted_headers", [])

                        callback_payload = {
                            "file_name": p.name,
                            "file_path": str(p),
                            "sheet_name": ws_curr.title,
                            "header_row_idx": h_idx,
                            "raw_headers": raw_headers,
                            "formatted_headers": formatted_headers,
                            "signature": sig,
                            "suggested_mapping": suggested_map
                        }

                        if mapping_callback:
                            confirmed_map, prof_name, should_skip = mapping_callback(callback_payload)
                            if should_skip or confirmed_map is None:
                                log(f"  ⏭️ Đã bỏ qua Sheet '{ws_curr.title}' theo lựa chọn của bạn.", "warning")
                                sheets_skipped += 1
                                continue

                            mapping = confirmed_map
                            profile_used_name = prof_name or f"Profile_{p.stem}_{ws_curr.title}"

                            # Lưu Profile này vào bộ nhớ cho những lần sau
                            self.profile_mgr.save_profile(
                                name=profile_used_name,
                                signature=sig,
                                raw_headers=raw_headers,
                                mapping=mapping
                            )
                            log(f"  💾 Đã lưu cấu hình ánh xạ: '{profile_used_name}' vào kho Profiles.", "success")
                        else:
                            # Không có UI callback -> sử dụng luôn gợi ý tự động
                            mapping = suggested_map
                            profile_used_name = "Tự động suy luận"

                if not mapping:
                    continue

                # 3. Trích xuất và chuẩn hóa dòng dữ liệu học sinh từ Sheet
                sheet_rows_extracted = 0
                empty_stt_consecutive = 0

                s_src_idx = mapping.get(1) # Cột STT nguồn
                n_src_idx = mapping.get(5) # Cột Họ tên nguồn

                # Đọc dòng tiêu đề chuẩn của sheet để kiểm tra lặp lại
                header_row_vals = [ws_curr.cell(row=h_idx, column=c).value for c in range(1, len(raw_headers) + 1)]

                for row in ws_curr.iter_rows(min_row=h_idx + 1, values_only=True):
                    if check_cancel and check_cancel():
                        break

                    # Kiểm tra ô STT nguồn
                    cell_stt = None
                    if s_src_idx is not None and 0 <= s_src_idx < len(row):
                        cell_stt = row[s_src_idx]

                    # Điều kiện ngắt: 15 dòng liên tiếp trống STT
                    if cell_stt is None or is_cell_empty(cell_stt):
                        empty_stt_consecutive += 1
                        if empty_stt_consecutive >= 15:
                            break
                    else:
                        empty_stt_consecutive = 0

                    # Kiểm tra dòng tiêu đề lặp lại
                    if is_header_duplicate_row(row, header_row_vals, s_src_idx if s_src_idx is not None else -1, n_src_idx if n_src_idx is not None else -1):
                        continue

                    # Kiểm tra dòng tổng kết / chân trang ký tên
                    if is_summary_or_footer_row(row):
                        continue

                    # Kiểm tra tính hợp lệ của dòng học sinh
                    cell_name = row[n_src_idx] if (n_src_idx is not None and 0 <= n_src_idx < len(row)) else None
                    stt_ok = is_valid_stt(cell_stt) if s_src_idx is not None else True
                    name_ok = is_valid_name(cell_name) if n_src_idx is not None else True

                    # Dòng học sinh thật sự phải có Họ tên hoặc STT hợp lệ (nếu cả 2 được map thì cả 2 phải thỏa mãn)
                    if s_src_idx is not None and n_src_idx is not None:
                        if not (stt_ok and name_ok):
                            continue
                    elif n_src_idx is not None:
                        if not name_ok:
                            continue
                    elif s_src_idx is not None:
                        if not stt_ok:
                            continue
                    else:
                        # Nếu không map cả STT lẫn Họ tên -> bỏ qua dòng trống
                        if not any(row):
                            continue

                    # Dòng hợp lệ -> Tiến hành ánh xạ sang 28 cột chuẩn
                    sheet_rows_extracted += 1
                    total_extracted_rows += 1

                    row_28: list[Any] = [None] * 28

                    for target_col in TARGET_COLUMNS:
                        t_idx = target_col["idx"]
                        t_col_0 = t_idx - 1
                        s_col_idx = mapping.get(t_idx)

                        raw_val = None
                        if s_col_idx is not None and 0 <= s_col_idx < len(row):
                            raw_val = row[s_col_idx]

                        # Chuẩn hóa giá trị ô
                        if raw_val is not None:
                            # 1. Chuyển đổi bảng mã cổ TCVN3 / VNI sang Unicode
                            val_str = auto_convert_legacy_encoding(str(raw_val))
                            val_str = clean_whitespace(val_str)

                            # 2. Chuẩn hóa đặc thù theo từng cột
                            if t_idx == 5: # Họ, chữ đệm và tên -> Title Case
                                row_28[t_col_0] = vietnamese_title_case(val_str)
                            elif t_idx in (9, 19): # Ngày sinh, Ngày cấp bằng -> format_date_to_text
                                row_28[t_col_0] = format_date_to_text(raw_val)
                            else:
                                row_28[t_col_0] = val_str
                        else:
                            row_28[t_col_0] = None

                        # Xử lý các giá trị mặc định nếu để trống
                        if fill_defaults and (row_28[t_col_0] is None or row_28[t_col_0] == ""):
                            if t_idx == 4: # Tên văn bằng
                                row_28[t_col_0] = "Bằng tốt nghiệp trung học cơ sở"
                            elif t_idx == 12: # Quốc tịch
                                row_28[t_col_0] = "Việt Nam"

                    # Đánh lại STT tăng dần 1..N
                    if renumber_stt:
                        row_28[0] = total_extracted_rows

                    # Ghi vào worksheet đích
                    ws_dst.append(row_28)
                    c_row = ws_dst.max_row
                    dst_cells = ws_dst[c_row]

                    for c_i in range(1, 29):
                        cell = dst_cells[c_i - 1]
                        if self.cached_fonts and c_i - 1 < len(self.cached_fonts) and self.cached_fonts[c_i - 1]:
                            cell.font = self.cached_fonts[c_i - 1]
                        if c_i in TEXT_COLS_28:
                            cell.number_format = '@'
                            if cell.value is not None:
                                cell.data_type = 's'
                                cell.value = str(cell.value)

                sheets_processed += 1
                log(f"    -> Sheet '{ws_curr.title}': Đã ghép {sheet_rows_extracted} học sinh hợp lệ.", "info")

            try:
                wb_src.close()
            except Exception:
                pass

            if progress_callback:
                progress_callback(f_idx, total_files, f"Đã ghép xong tệp {f_idx}/{total_files}: {p.name}")

        # Hoàn tất: Freeze panes & Auto-filter
        ws_dst.freeze_panes = "A3"
        ws_dst.auto_filter.ref = f"A2:AB{max(2, ws_dst.max_row)}"

        wb_dst.save(out_path)
        wb_dst.close()

        log(f"🎉 HOÀN TẤT GHÉP & CHUẨN HÓA! Đã thu nạp {total_extracted_rows} dòng dữ liệu vào:", "success")
        log(f"   📄 {out_path.name} (Đường dẫn: {out_path})", "success")

        return {
            "status": "success",
            "output_file": str(out_path),
            "total_rows": total_extracted_rows,
            "sheets_processed": sheets_processed,
            "sheets_skipped": sheets_skipped
        }
