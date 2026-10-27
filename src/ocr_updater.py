from __future__ import annotations

import json
import os
import re
import time
import unicodedata
from pathlib import Path
from typing import Any, Callable

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.custom_excel_manager import detect_header_and_columns
from src.normalizer import clean_whitespace, format_date_to_text, vietnamese_title_case
from src.ocr_processor import (
    DEFAULT_BASE_URL,
    DEFAULT_MODEL_NAME,
    OCRConfigManager,
    OCRProcessor,
    extract_table_from_image,
    parse_ocr_json_response,
)
from src.lookup_manager import get_persistent_data_dir
from src.smart_merger import TARGET_COLUMNS, normalize_text_for_match

# Định nghĩa các tiêu chí so khớp dòng học sinh
AVAILABLE_MATCH_FIELDS = [
    {"key": "stt", "label": "STT (Cột 1)", "col_idx": 1},
    {"key": "so_hieu_bang", "label": "Số hiệu bằng (Cột 2)", "col_idx": 2},
    {"key": "so_vao_so", "label": "Số vào sổ (Cột 3)", "col_idx": 3},
    {"key": "ho_ten", "label": "Họ và tên (Cột 5)", "col_idx": 5},
    {"key": "ngay_sinh", "label": "Ngày sinh (Cột 9)", "col_idx": 9},
    {"key": "ten_truong", "label": "Trường đang học (Cột 14)", "col_idx": 14},
]

# Ánh xạ từ tên trường dữ liệu sang chỉ số cột chuẩn 28 (1-based)
FIELD_TO_COL_IDX: dict[str, int] = {
    "stt": 1,
    "so_hieu_bang": 2,
    "so_vao_so": 3,
    "ten_van_bang": 4,
    "ho_ten": 5,
    "ma_nguoi_hoc": 6,
    "cccd": 7,
    "ho_chieu": 8,
    "ngay_sinh": 9,
    "gioi_tinh": 10,
    "dan_toc": 11,
    "quoc_tich": 12,
    "noi_sinh": 13,
    "ten_truong": 14,
    "nien_khoa": 15,
    "diem_thi": 16,
    "xep_loai": 16,
    "hoi_dong_thi": 17,
    "dia_danh": 18,
    "ngay_cap_bang": 19,
    "co_quan_cap_bang": 20,
    "chuc_danh_nguoi_ky": 21,
    "nguoi_ky": 22,
    "tinh_trang_van_bang": 23,
    "ghi_chu": 24,
    "trang_thai": 25,
    "ly_do_loi": 26,
    "ma_don_vi": 27,
    "ten_don_vi": 28,
}

# Từ đồng nghĩa và tên tiếng Việt của các cột
FIELD_SYNONYMS: dict[str, int] = {
    "so thu tu": 1,
    "stt": 1,
    "so hieu": 2,
    "so hieu bang": 2,
    "so vao so": 3,
    "so vao so goc": 3,
    "vao so": 3,
    "ten van bang": 4,
    "ho ten": 5,
    "ho va ten": 5,
    "ten": 5,
    "ho, chu dem va ten": 5,
    "ma nguoi hoc": 6,
    "cccd": 7,
    "dinh danh": 7,
    "so dinh danh": 7,
    "ho chieu": 8,
    "ngay sinh": 9,
    "nam sinh": 9,
    "ngay thang nam sinh": 9,
    "gioi tinh": 10,
    "gt": 10,
    "dan toc": 11,
    "quoc tich": 12,
    "noi sinh": 13,
    "ten truong": 14,
    "truong": 14,
    "truong dang hoc": 14,
    "nien khoa": 15,
    "diem thi": 16,
    "xep loai": 16,
    "xep loai tn": 16,
    "hoi dong thi": 17,
    "dia danh": 18,
    "ngay cap bang": 19,
    "co quan cap bang": 20,
    "chuc danh nguoi ky": 21,
    "nguoi ky": 22,
    "tinh trang": 23,
    "ghi chu": 24,
    "trang thai": 25,
}

DEFAULT_PATCH_PROMPT = """Bạn là chuyên gia số hóa văn bằng chứng chỉ THCS, có nhiệm vụ trích xuất dữ liệu từ các trang tài liệu/sổ bộ CÓ CHỨA ĐÁNH DẤU SỬA ĐỔI HOẶC BỔ SUNG THÊM HỌC SINH.

ĐẶC ĐIỂM TÀI LIỆU CẦN ĐỌC:
1. SỬA ĐỔI (Sửa trên học sinh đã in sẵn):
   - Có chữ in máy bị gạch ngang/gạch xóa và được viết đè chữ viết tay hoặc bút đỏ (sửa Họ tên, Ngày sinh, Giới tính, Nơi sinh, Xếp loại, Số hiệu bằng, Số vào sổ, Ghi chú...).
   - BẮT BUỘC lấy giá trị viết tay/bút đỏ sửa đổi làm giá trị mới ("gia_tri_moi").
2. BỔ SUNG (Học sinh mới được chèn thêm):
   - Là những dòng học sinh được viết tay toàn bộ hoặc ghi thêm bằng bút mực/bút đỏ ở dòng trống, cuối trang, hoặc chèn giữa danh sách.

QUY TẮC TRÍCH XUẤT:
- BỎ QUA những dòng in máy bình thường không hề có bất kỳ dấu hiệu sửa đổi, gạch xóa hay ghi chú viết tay nào.
- CHỈ trích xuất những dòng có SỬA ĐỔI hoặc BỔ SUNG.

ĐẦU RA BẮT BUỘC:
Trả về DUY NHẤT một mảng JSON (chuỗi JSON chuẩn), không bọc trong ```json, không kèm giải thích ngoài.
Mỗi phần tử trong mảng là một object đại diện cho một học sinh với cấu trúc:
[
  {
    "loai": "SUA_DOI" hoặc "BO_SUNG",
    "stt": <số thứ tự trên trang, nếu có, ví dụ 15>,
    "so_hieu_bang": "<số hiệu bằng sau khi sửa nếu có>",
    "so_vao_so": "<số vào sổ gốc sau khi sửa nếu có>",
    "ho_ten": "<họ và tên đầy đủ sau khi sửa>",
    "ngay_sinh": "<ngày sinh dạng dd/mm/yyyy sau khi sửa>",
    "gioi_tinh": "<Nam hoặc Nữ>",
    "dan_toc": "<dân tộc nếu có, vd: Kinh>",
    "noi_sinh": "<nơi sinh tỉnh/huyện sau khi sửa nếu có>",
    "ten_truong": "<tên trường hoặc hội đồng thi>",
    "xep_loai": "<xếp loại tốt nghiệp sau khi sửa nếu có: Giỏi/Khá/TB>",
    "ghi_chu": "<ghi chú sau khi sửa/bổ sung>",
    "cac_truong_sua": [
      {
        "ten_truong": "<tên trường dữ liệu bị sửa, ví dụ: ho_ten, ngay_sinh, gioi_tinh, noi_sinh, so_hieu_bang, so_vao_so, xep_loai, ghi_chu>",
        "gia_tri_cu": "<giá trị in máy cũ bị gạch xóa>",
        "gia_tri_moi": "<giá trị sửa đổi mới viết tay/mực đỏ>"
      }
    ],
    "can_kiem_tra": false,
    "ly_do_nghi_ngo": null
  }
]

Nếu trang không có dòng nào bị sửa đổi hay bổ sung, trả về mảng rỗng: []"""


def normalize_match_val(val: Any, field_key: str) -> str:
    """Chuẩn hóa giá trị để so khớp chính xác giữa OCR và file Excel."""
    if val is None:
        return ""
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ("none", "nan", "null"):
        return ""

    if field_key == "stt":
        val_clean = re.sub(r"\.0+$", "", val_str)
        if val_clean.isdigit():
            return str(int(val_clean))
        return val_clean.lower()

    if field_key == "ngay_sinh":
        norm_date = format_date_to_text(val_str)
        return norm_date or val_str

    if field_key == "ho_ten":
        val_norm = unicodedata.normalize("NFC", val_str)
        return re.sub(r"\s+", " ", val_norm).strip().lower()

    if field_key in ("so_hieu_bang", "so_vao_so"):
        return re.sub(r"\s+", "", val_str).lower()

    if field_key == "ten_truong":
        val_norm = unicodedata.normalize("NFC", val_str)
        return re.sub(r"\s+", " ", val_norm).strip().lower()

    return re.sub(r"\s+", " ", val_str).strip().lower()


class OCRColumnMappingManager:
    """
    Quản lý cấu hình ánh xạ cột từ tên cột trên ảnh sổ sang 28 cột chuẩn.
    Lưu trữ bền vững trong file JSON (ocr_column_mappings.json).
    Cho phép người dùng tạo Profile riêng cho từng loại sổ / trường học và tái sử dụng.
    """
    def __init__(self, config_path: Path | str | None = None):
        if config_path:
            self.config_path = Path(config_path)
        else:
            self.config_path = get_persistent_data_dir() / "ocr_column_mappings.json"
        self._cache: dict[str, Any] | None = None

    def get_default_mapping(self) -> dict[str, int]:
        return {
            "STT": 1,
            "Số thứ tự": 1,
            "Số hiệu bằng": 2,
            "Số hiệu": 2,
            "Số bằng": 2,
            "Số vào sổ gốc": 3,
            "Số vào sổ": 3,
            "Vào sổ": 3,
            "Tên văn bằng": 4,
            "Họ, chữ đệm và tên": 5,
            "Họ và tên": 5,
            "Họ tên": 5,
            "Tên học sinh": 5,
            "Mã người học": 6,
            "Số định danh cá nhân": 7,
            "CCCD": 7,
            "Hộ chiếu": 8,
            "Ngày, tháng, năm sinh": 9,
            "Ngày sinh": 9,
            "Năm sinh": 9,
            "Giới tính": 10,
            "GT": 10,
            "Dân tộc": 11,
            "Quốc tịch": 12,
            "Nơi sinh": 13,
            "Quê quán": 13,
            "Tên trường đã học": 14,
            "Tên trường": 14,
            "Trường": 14,
            "Niên khóa": 15,
            "Điểm các môn thi": 16,
            "Điểm thi": 16,
            "Xếp loại": 16,
            "Hội đồng thi": 17,
            "HĐT": 17,
            "Địa danh": 18,
            "Ngày tháng năm cấp bằng": 19,
            "Ngày cấp bằng": 19,
            "Tên cơ quan cấp bằng": 20,
            "Chức danh người ký": 21,
            "Người ký bằng": 22,
            "Tình trạng văn bằng": 23,
            "Ghi chú": 24,
            "Trạng thái số hóa": 25,
            "Mã đơn vị": 27,
            "Tên đơn vị": 28,
        }

    def load_data(self) -> dict[str, Any]:
        if self._cache is not None:
            return self._cache
        if self.config_path.is_file():
            try:
                data = json.loads(self.config_path.read_text(encoding="utf-8"))
                if isinstance(data, dict) and "profiles" in data:
                    self._cache = data
                    return data
            except Exception:
                pass

        default_map = self.get_default_mapping()
        data = {
            "active_profile": "Mặc định (Chuẩn 28 cột)",
            "profiles": {
                "Mặc định (Chuẩn 28 cột)": {
                    "name": "Mặc định (Chuẩn 28 cột)",
                    "mapping": default_map
                }
            }
        }
        self._cache = data
        self.save_data(data)
        return data

    def save_data(self, data: dict[str, Any]) -> None:
        self._cache = data
        try:
            self.config_path.parent.mkdir(parents=True, exist_ok=True)
            self.config_path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except Exception:
            pass

    def list_profiles(self) -> list[str]:
        data = self.load_data()
        return list(data.get("profiles", {}).keys())

    def get_active_profile_name(self) -> str:
        data = self.load_data()
        return data.get("active_profile", "Mặc định (Chuẩn 28 cột)")

    def set_active_profile_name(self, name: str) -> None:
        data = self.load_data()
        if name in data.get("profiles", {}):
            data["active_profile"] = name
            self.save_data(data)

    def get_mapping(self, profile_name: str | None = None) -> dict[str, int]:
        data = self.load_data()
        p_name = profile_name or data.get("active_profile") or "Mặc định (Chuẩn 28 cột)"
        prof = data.get("profiles", {}).get(p_name)
        if prof and isinstance(prof.get("mapping"), dict):
            return prof["mapping"]
        return self.get_default_mapping()

    def save_profile(self, profile_name: str, mapping: dict[str, int]) -> None:
        data = self.load_data()
        if "profiles" not in data:
            data["profiles"] = {}
        data["profiles"][profile_name] = {
            "name": profile_name,
            "mapping": mapping
        }
        data["active_profile"] = profile_name
        self.save_data(data)

    def delete_profile(self, profile_name: str) -> bool:
        data = self.load_data()
        profiles = data.get("profiles", {})
        if profile_name in profiles and len(profiles) > 1:
            del profiles[profile_name]
            if data.get("active_profile") == profile_name:
                data["active_profile"] = list(profiles.keys())[0]
            self.save_data(data)
            return True
        return False


def resolve_field_to_col_idx(field_name: str, custom_mapping: dict[str, int] | None = None) -> int | None:
    """
    Nhận diện tên trường dữ liệu từ ảnh sang số cột chuẩn 1..28.
    Ưu tiên 100% theo bảng cấu hình ánh xạ do người dùng chỉ định.
    """
    if not field_name:
        return None
    raw_key = str(field_name).strip()
    norm = normalize_text_for_match(raw_key)

    # 1. Tra cứu trực tiếp trong cấu hình người dùng (nếu có)
    if custom_mapping:
        for k, col in custom_mapping.items():
            if str(k).strip().lower() == raw_key.lower():
                return int(col)
            if normalize_text_for_match(str(k)) == norm:
                return int(col)

    # 2. Tra cứu trong bảng chuẩn mặc định
    lower_k = raw_key.lower()
    if lower_k in FIELD_TO_COL_IDX:
        return FIELD_TO_COL_IDX[lower_k]

    if norm in FIELD_SYNONYMS:
        return FIELD_SYNONYMS[norm]

    for syn_text, c_idx in FIELD_SYNONYMS.items():
        if syn_text in norm or norm in syn_text:
            return c_idx

    return None


def match_student_row(
    candidate: dict[str, Any],
    sheet_rows: list[dict[str, Any]],
    selected_keys: list[str],
) -> tuple[int | None, list[int], str]:
    """
    Tìm dòng học sinh trong danh sách Excel khớp với thông tin trích xuất từ OCR.
    
    Hỗ trợ đối chiếu thông minh:
    - Nếu trường đang so sánh (ví dụ Họ tên) bị sửa đổi, cho phép khớp với cả giá trị cũ
      hoặc giá trị mới.
    - Trả về (matched_row_idx, all_candidate_indices, note)
    """
    if not selected_keys:
        return None, [], "Chưa chọn tiêu chí so khớp nào"

    # Trích xuất danh sách giá trị cũ nếu có trong cac_truong_sua
    old_vals: dict[str, str] = {}
    for mod in candidate.get("cac_truong_sua", []):
        t_key = mod.get("ten_truong", "")
        old_v = mod.get("gia_tri_cu")
        if t_key and old_v:
            old_vals[t_key] = str(old_v).strip()

    matches: list[int] = []

    for r_idx, row_data in enumerate(sheet_rows):
        is_all_match = True

        for k in selected_keys:
            cand_val = candidate.get(k)
            excel_val = row_data.get(k)

            norm_excel = normalize_match_val(excel_val, k)
            norm_cand_new = normalize_match_val(cand_val, k)
            norm_cand_old = normalize_match_val(old_vals.get(k), k) if k in old_vals else ""

            # Nếu phía OCR hoàn toàn không có dữ liệu cho tiêu chí này -> coi như không khớp
            if not norm_cand_new and not norm_cand_old:
                is_all_match = False
                break

            # Khớp nếu giá trị trong Excel trùng với giá trị mới HOẶC giá trị cũ của OCR
            matched_k = False
            if norm_cand_new and norm_excel == norm_cand_new:
                matched_k = True
            elif norm_cand_old and norm_excel == norm_cand_old:
                matched_k = True

            if not matched_k:
                is_all_match = False
                break

        if is_all_match:
            matches.append(r_idx)

    if len(matches) == 1:
        return matches[0], matches, "Khớp duy nhất"
    elif len(matches) > 1:
        return None, matches, f"Trùng {len(matches)} dòng (Cần kiểm tra lại)"
    else:
        return None, [], "Không tìm thấy trong danh sách"


class OCRUpdaterProcessor:
    """Bộ xử lý cập nhật danh sách 28 cột từ ảnh chụp sửa đổi/bổ sung."""

    def __init__(self, config_mgr: OCRConfigManager | None = None):
        self.config_mgr = config_mgr or OCRConfigManager()

    def update_excel_from_images(
        self,
        excel_path: Path | str,
        image_paths: list[Path | str],
        selected_match_keys: list[str] | None = None,
        column_mapping: dict[str, int] | None = None,
        auto_append_new: bool = True,
        output_path: Path | str | None = None,
        api_key: str | None = None,
        base_url: str | None = None,
        model_name: str | None = None,
        prompt_text: str | None = None,
        progress_callback: Callable[[int, int, str], None] | None = None,
        log_callback: Callable[[str, str], None] | None = None,
        check_cancel: Callable[[], bool] | None = None,
        manual_review_callback: Callable[[dict[str, Any]], dict[str, Any]] | None = None,
    ) -> dict[str, Any]:
        """
        Thực thi OCR từng ảnh sửa đổi/bổ sung, định vị học sinh trong file Excel 28 cột
        và áp dụng các thay đổi kèm tô màu highlight và tạo sheet lịch sử.
        Hỗ trợ cấu hình ánh xạ cột tùy biến 100% chuẩn xác và hộp thoại duyệt thủ công.
        """
        def log(msg: str, tag: str = "info"):
            if log_callback:
                log_callback(msg, tag)

        excel_p = Path(excel_path)
        if not excel_p.is_file():
            raise FileNotFoundError(f"Không tìm thấy tệp Excel danh sách: {excel_p}")

        imgs = [Path(p) for p in image_paths if Path(p).is_file()]
        if not imgs:
            return {"status": "error", "message": "Không có tệp ảnh hợp lệ để xử lý."}

        keys_to_match = selected_match_keys or ["ho_ten", "ngay_sinh"]

        # Cấu hình Ánh xạ cột (100% chuẩn xác theo người dùng thiết lập)
        active_mapping = column_mapping
        if active_mapping is None:
            active_mapping = OCRColumnMappingManager().get_mapping()
        log(f"  🎯 Áp dụng bảng ánh xạ cột ({len(active_mapping)} quy tắc xác định 100%).", "info")

        # Cấu hình API OCR
        cfg = self.config_mgr.load_config()
        active_api_key = (api_key or cfg.get("api_key") or "").strip()
        active_base_url = (base_url or cfg.get("base_url") or DEFAULT_BASE_URL).strip()
        active_model_name = (model_name or cfg.get("model_name") or DEFAULT_MODEL_NAME).strip()
        active_prompt = (prompt_text or DEFAULT_PATCH_PROMPT).strip()

        if not active_api_key:
            return {
                "status": "error",
                "message": "Chưa có API Key. Vui lòng nhập API Key trên giao diện hoặc tệp cấu hình.",
            }

        log(f"📖 Đang nạp tệp Excel: {excel_p.name}...", "info")
        wb = openpyxl.load_workbook(excel_p, data_only=False)

        # Chọn Sheet dữ liệu chuẩn
        ws = None
        for s in wb.worksheets:
            if s.sheet_state == "visible" and s.title != "Lich_Su_Thay_Doi":
                ws = s
                break
        if ws is None:
            ws = wb.active

        header_info = detect_header_and_columns(ws)
        header_row_idx = header_info.get("header_row_idx", 1)
        raw_headers = header_info.get("raw_headers", [])
        log(f"  🔍 Dòng tiêu đề tại dòng {header_row_idx} với {len(raw_headers)} cột. Sheet: '{ws.title}'.", "info")

        # Nạp dữ liệu các dòng hiện có trong Excel vào bộ nhớ
        sheet_rows: list[dict[str, Any]] = []
        excel_row_indices: list[int] = []

        max_existing_stt = 0
        last_non_empty_row = header_row_idx

        for r_num in range(header_row_idx + 1, ws.max_row + 1):
            row_dict: dict[str, Any] = {}
            has_data = False

            # Đọc các trường nhận diện chính
            stt_val = ws.cell(row=r_num, column=1).value
            so_hieu_val = ws.cell(row=r_num, column=2).value
            so_vao_so_val = ws.cell(row=r_num, column=3).value
            ho_ten_val = ws.cell(row=r_num, column=5).value
            ngay_sinh_val = ws.cell(row=r_num, column=9).value
            ten_truong_val = ws.cell(row=r_num, column=14).value

            row_dict["stt"] = stt_val
            row_dict["so_hieu_bang"] = so_hieu_val
            row_dict["so_vao_so"] = so_vao_so_val
            row_dict["ho_ten"] = ho_ten_val
            row_dict["ngay_sinh"] = ngay_sinh_val
            row_dict["ten_truong"] = ten_truong_val

            # Kiểm tra dòng có dữ liệu không
            if ho_ten_val or stt_val or so_hieu_val:
                has_data = True
                last_non_empty_row = r_num

            if stt_val is not None:
                try:
                    s_int = int(float(str(stt_val).strip()))
                    if s_int > max_existing_stt:
                        max_existing_stt = s_int
                except Exception:
                    pass

            sheet_rows.append(row_dict)
            excel_row_indices.append(r_num)

        log(f"  📊 Đã lập chỉ mục {len(sheet_rows):,} học sinh trong danh sách gốc (STT cao nhất: {max_existing_stt}).", "info")

        # Định dạng màu cho các ô thay đổi / bổ sung
        FILL_MODIFIED = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")  # Vàng nhạt
        FILL_ADDED = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")     # Xanh lá nhạt
        FONT_MODIFIED = Font(name="Times New Roman", size=11, bold=True, color="B25900")
        FONT_ADDED = Font(name="Times New Roman", size=11, bold=True, color="276A3C")
        BORDER_THIN = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        audit_entries: list[dict[str, Any]] = []
        total_images = len(imgs)
        success_images = 0
        modified_students_count = 0
        added_students_count = 0
        unmatched_count = 0

        # Lấy một dòng mẫu để kế thừa thông tin chung (Hội đồng thi, Ngày cấp, Tên trường...) khi bổ sung học sinh mới
        sample_context_row: dict[int, Any] = {}
        if last_non_empty_row > header_row_idx:
            for c in range(1, 29):
                sample_context_row[c] = ws.cell(row=last_non_empty_row, column=c).value

        log(f"🚀 BẮT ĐẦU CHẠY OCR & CẬP NHẬT TRÊN {total_images} ẢNH CHỤP...", "info")

        # Biến trạng thái duyệt thủ công
        review_active = bool(manual_review_callback)

        for idx, img_p in enumerate(imgs, start=1):
            if check_cancel and check_cancel():
                log(f"🛑 Đã dừng theo yêu cầu tại ảnh {img_p.name}.", "warning")
                break

            prog_pct = int(((idx - 1) / total_images) * 100)
            if progress_callback:
                progress_callback(prog_pct, 100, f"[{idx}/{total_images}] OCR: {img_p.name}")

            log(f"📷 [{idx}/{total_images}] Đang đọc ảnh: {img_p.name}...", "info")

            try:
                raw_response = extract_table_from_image(
                    image_path=img_p,
                    prompt_text=active_prompt,
                    api_key=active_api_key,
                    base_url=active_base_url,
                    model_name=active_model_name,
                    log_callback=log_callback,
                )
            except Exception as e:
                log(f"  ❌ Lỗi khi gửi ảnh {img_p.name}: {e}", "error")
                continue

            records, err_msg = parse_ocr_json_response(raw_response)
            if records is None:
                log(f"  ⚠️ {img_p.name}: Không trích xuất được JSON ({err_msg}).", "warning")
                continue

            success_images += 1
            if not records:
                log(f"  ℹ️ {img_p.name}: Không phát hiện dòng nào có sửa đổi hoặc bổ sung.", "info")
                continue

            log(f"  ✨ {img_p.name}: Phát hiện {len(records)} học sinh có đánh dấu sửa đổi/bổ sung.", "success")

            stop_process = False

            for rec in records:
                if stop_process or (check_cancel and check_cancel()):
                    break

                rec_type = str(rec.get("loai") or "").strip().upper()
                ho_ten = str(rec.get("ho_ten") or "").strip()
                stt_val = rec.get("stt")

                # 1. Thử so khớp học sinh với các dòng hiện có
                matched_idx, all_cand, match_note = match_student_row(rec, sheet_rows, keys_to_match)

                if matched_idx is not None:
                    # KHỚP ĐƯỢC 1 DÒNG DUY NHẤT -> TIẾN HÀNH CẬP NHẬT
                    excel_r = excel_row_indices[matched_idx]
                    mod_fields = rec.get("cac_truong_sua", [])

                    # Nếu AI không liệt kê cac_truong_sua cụ thể, so sánh trực tiếp các trường cơ bản
                    if not mod_fields:
                        for field_k in ("ho_ten", "ngay_sinh", "gioi_tinh", "noi_sinh", "so_hieu_bang", "so_vao_so", "xep_loai", "ghi_chu"):
                            new_v = rec.get(field_k)
                            if new_v is not None and str(new_v).strip():
                                c_idx = resolve_field_to_col_idx(field_k, custom_mapping=active_mapping)
                                if c_idx:
                                    old_cell_v = ws.cell(row=excel_r, column=c_idx).value
                                    if normalize_match_val(old_cell_v, field_k) != normalize_match_val(new_v, field_k):
                                        mod_fields.append({
                                            "ten_truong": field_k,
                                            "gia_tri_cu": str(old_cell_v or ""),
                                            "gia_tri_moi": str(new_v).strip()
                                        })

                    # Chuẩn bị danh sách trường thay đổi
                    candidate_changes: list[dict[str, Any]] = []
                    for m_item in mod_fields:
                        f_name = m_item.get("ten_truong", "")
                        col_idx = resolve_field_to_col_idx(f_name, custom_mapping=active_mapping)
                        if not col_idx:
                            continue

                        val_new = m_item.get("gia_tri_moi", "")
                        val_old = m_item.get("gia_tri_cu", "")
                        curr_excel_val = ws.cell(row=excel_r, column=col_idx).value

                        col_label = TARGET_COLUMNS[col_idx - 1]["name"] if (1 <= col_idx <= len(TARGET_COLUMNS)) else f_name
                        candidate_changes.append({
                            "field_name": f_name,
                            "col_idx": col_idx,
                            "col_name": col_label,
                            "old_val": str(curr_excel_val if curr_excel_val is not None else (val_old or "")),
                            "new_val": str(val_new or ""),
                        })

                    if not candidate_changes:
                        continue

                    # NẾU BẬT DUYỆT THỦ CÔNG: GỌI CALLBACK HIỆN HỘP THOẠI
                    if review_active and manual_review_callback:
                        review_req = {
                            "type": "SUA_DOI",
                            "student_name": ho_ten or sheet_rows[matched_idx].get("ho_ten", ""),
                            "row_num": excel_r,
                            "image_name": img_p.name,
                            "reason": rec.get("ly_do_nghi_ngo"),
                            "changes": candidate_changes,
                        }
                        user_decision = manual_review_callback(review_req)
                        act = user_decision.get("action", "apply")
                        if act == "skip":
                            log(f"    ⏭️ Người dùng bỏ qua sửa học sinh '{ho_ten}' (Dòng {excel_r}).", "warning")
                            audit_entries.append({
                                "row_num": excel_r,
                                "type": "Bỏ qua",
                                "student_name": ho_ten,
                                "field_name": "---",
                                "old_val": "---",
                                "new_val": "---",
                                "image_name": img_p.name,
                                "status": "Người dùng bỏ qua",
                                "note": "Bỏ qua khi duyệt thủ công"
                            })
                            continue
                        elif act == "apply_all":
                            review_active = False
                            log("    ⚡ Chế độ: Tự động áp dụng tất cả các học sinh còn lại.", "info")
                        elif act == "stop":
                            log("🛑 Người dùng yêu cầu dừng tiến trình duyệt.", "warning")
                            stop_process = True
                            break

                        # Nhận các giá trị người dùng đã sửa đổi thủ công
                        candidate_changes = user_decision.get("changes", candidate_changes)

                    modified_students_count += 1
                    updated_cols_log = []
                    for c_info in candidate_changes:
                        col_idx = c_info["col_idx"]
                        f_name = c_info["field_name"]
                        val_new = c_info["new_val"]
                        val_old = c_info["old_val"]

                        # Chuẩn hóa giá trị theo nghiệp vụ cột
                        if col_idx == 5:
                            val_new = vietnamese_title_case(str(val_new))
                        elif col_idx == 9:
                            val_new = format_date_to_text(val_new)
                        elif col_idx in (2, 3):
                            val_new = re.sub(r"\s+", "", str(val_new))

                        cell = ws.cell(row=excel_r, column=col_idx)
                        curr_excel_val = cell.value
                        cell.value = val_new
                        cell.fill = FILL_MODIFIED
                        cell.font = FONT_MODIFIED

                        # Cập nhật lại chỉ mục trong bộ nhớ
                        for k, v_c in FIELD_TO_COL_IDX.items():
                            if v_c == col_idx:
                                sheet_rows[matched_idx][k] = val_new

                        updated_cols_log.append(f_name)
                        audit_entries.append({
                            "row_num": excel_r,
                            "type": "Sửa đổi",
                            "student_name": ho_ten or sheet_rows[matched_idx].get("ho_ten", ""),
                            "field_name": c_info.get("col_name", f_name),
                            "old_val": val_old or str(curr_excel_val or ""),
                            "new_val": val_new,
                            "image_name": img_p.name,
                            "status": "Thành công",
                            "note": rec.get("ly_do_nghi_ngo") or "Đã duyệt/cập nhật"
                        })

                    log(f"    ✏️ Dòng {excel_r} ({ho_ten}): Đã sửa {len(candidate_changes)} ô ({', '.join(updated_cols_log)}).", "info")

                elif rec_type == "BO_SUNG" or auto_append_new:
                    # BỔ SUNG HỌC SINH MỚI VÀO CUỐI DANH SÁCH
                    max_existing_stt += 1
                    new_r = ws.max_row + 1

                    # Tạo dữ liệu 28 cột cho học sinh mới
                    new_row_vals: list[Any] = [None] * 28

                    # 1: STT
                    new_row_vals[0] = max_existing_stt
                    # 2: Số hiệu bằng
                    new_row_vals[1] = re.sub(r"\s+", "", str(rec.get("so_hieu_bang") or ""))
                    # 3: Số vào sổ
                    new_row_vals[2] = re.sub(r"\s+", "", str(rec.get("so_vao_so") or ""))
                    # 4: Tên văn bằng
                    new_row_vals[3] = sample_context_row.get(4) or "Bằng tốt nghiệp trung học cơ sở"
                    # 5: Họ và tên
                    new_row_vals[4] = vietnamese_title_case(str(rec.get("ho_ten") or ""))
                    # 6: Mã người học
                    new_row_vals[5] = ""
                    # 7: CCCD
                    new_row_vals[6] = ""
                    # 8: Hộ chiếu
                    new_row_vals[7] = ""
                    # 9: Ngày sinh
                    new_row_vals[8] = format_date_to_text(rec.get("ngay_sinh") or "")
                    # 10: Giới tính
                    new_row_vals[9] = clean_whitespace(str(rec.get("gioi_tinh") or ""))
                    # 11: Dân tộc
                    new_row_vals[10] = clean_whitespace(str(rec.get("dan_toc") or sample_context_row.get(11) or "Kinh"))
                    # 12: Quốc tịch
                    new_row_vals[11] = sample_context_row.get(12) or "Việt Nam"
                    # 13: Nơi sinh
                    new_row_vals[12] = clean_whitespace(str(rec.get("noi_sinh") or ""))
                    # 14: Tên trường
                    new_row_vals[13] = clean_whitespace(str(rec.get("ten_truong") or sample_context_row.get(14) or ""))
                    # 15: Niên khóa
                    new_row_vals[14] = sample_context_row.get(15) or ""
                    # 16: Điểm/Xếp loại
                    new_row_vals[15] = clean_whitespace(str(rec.get("xep_loai") or ""))
                    # 17: Hội đồng thi
                    new_row_vals[16] = sample_context_row.get(17) or ""
                    # 18: Địa danh
                    new_row_vals[17] = sample_context_row.get(18) or ""
                    # 19: Ngày cấp bằng
                    new_row_vals[18] = sample_context_row.get(19) or ""
                    # 20: Tên cơ quan cấp bằng
                    new_row_vals[19] = sample_context_row.get(20) or ""
                    # 21: Chức danh người ký
                    new_row_vals[20] = sample_context_row.get(21) or ""
                    # 22: Họ tên người ký
                    new_row_vals[21] = sample_context_row.get(22) or ""
                    # 23: Tình trạng văn bằng
                    new_row_vals[22] = sample_context_row.get(23) or ""
                    # 24: Ghi chú
                    new_row_vals[23] = clean_whitespace(str(rec.get("ghi_chu") or "Bổ sung từ ảnh sửa đổi"))
                    # 25: Trạng thái số hóa
                    new_row_vals[24] = "Đã bổ sung từ ảnh sửa đổi"
                    # 26: Lý do lỗi
                    new_row_vals[25] = ""
                    # 27: Mã đơn vị
                    new_row_vals[26] = sample_context_row.get(27) or ""
                    # 28: Tên đơn vị
                    new_row_vals[27] = sample_context_row.get(28) or ""

                    # Áp dụng các trường tùy biến khác trên ảnh theo bảng ánh xạ người dùng
                    for k_rec, v_rec in rec.items():
                        if k_rec in ("loai", "stt", "can_kiem_tra", "ly_do_nghi_ngo", "cac_truong_sua"):
                            continue
                        if v_rec is None or not str(v_rec).strip():
                            continue
                        c_i = resolve_field_to_col_idx(k_rec, custom_mapping=active_mapping)
                        if c_i and 1 <= c_i <= 28:
                            val_str = str(v_rec).strip()
                            if c_i == 5:
                                val_str = vietnamese_title_case(val_str)
                            elif c_i == 9:
                                val_str = format_date_to_text(val_str)
                            elif c_i in (2, 3):
                                val_str = re.sub(r"\s+", "", val_str)
                            new_row_vals[c_i - 1] = val_str

                    # NẾU BẬT DUYỆT THỦ CÔNG KHI THÊM MỚI
                    if review_active and manual_review_callback:
                        new_student_fields = [
                            {"field_name": "ho_ten", "col_idx": 5, "col_name": "Họ và tên", "val": str(new_row_vals[4] or "")},
                            {"field_name": "ngay_sinh", "col_idx": 9, "col_name": "Ngày, tháng, năm sinh", "val": str(new_row_vals[8] or "")},
                            {"field_name": "gioi_tinh", "col_idx": 10, "col_name": "Giới tính", "val": str(new_row_vals[9] or "")},
                            {"field_name": "noi_sinh", "col_idx": 13, "col_name": "Nơi sinh", "val": str(new_row_vals[12] or "")},
                            {"field_name": "so_hieu_bang", "col_idx": 2, "col_name": "Số hiệu bằng", "val": str(new_row_vals[1] or "")},
                            {"field_name": "so_vao_so", "col_idx": 3, "col_name": "Số vào sổ", "val": str(new_row_vals[2] or "")},
                            {"field_name": "xep_loai", "col_idx": 16, "col_name": "Điểm / Xếp loại", "val": str(new_row_vals[15] or "")},
                            {"field_name": "ghi_chu", "col_idx": 24, "col_name": "Ghi chú", "val": str(new_row_vals[23] or "")},
                        ]
                        review_req = {
                            "type": "BO_SUNG",
                            "student_name": ho_ten,
                            "row_num": new_r,
                            "image_name": img_p.name,
                            "reason": rec.get("ly_do_nghi_ngo"),
                            "fields": new_student_fields,
                        }
                        user_decision = manual_review_callback(review_req)
                        act = user_decision.get("action", "apply")
                        if act == "skip":
                            log(f"    ⏭️ Người dùng bỏ qua thêm học sinh mới '{ho_ten}'.", "warning")
                            max_existing_stt -= 1
                            continue
                        elif act == "apply_all":
                            review_active = False
                            log("    ⚡ Chế độ: Tự động áp dụng tất cả các học sinh còn lại.", "info")
                        elif act == "stop":
                            log("🛑 Người dùng yêu cầu dừng tiến trình duyệt.", "warning")
                            max_existing_stt -= 1
                            stop_process = True
                            break

                        for ef in user_decision.get("fields", new_student_fields):
                            c_i = ef.get("col_idx")
                            v_ed = ef.get("val", "")
                            if c_i and 1 <= c_i <= 28:
                                new_row_vals[c_i - 1] = v_ed

                    added_students_count += 1

                    # Ghi vào sheet
                    for c_i, v_val in enumerate(new_row_vals, start=1):
                        c_cell = ws.cell(row=new_r, column=c_i, value=v_val)
                        c_cell.fill = FILL_ADDED
                        c_cell.font = FONT_ADDED
                        c_cell.border = BORDER_THIN

                    # Cập nhật chỉ mục bộ nhớ
                    sheet_rows.append({
                        "stt": max_existing_stt,
                        "so_hieu_bang": new_row_vals[1],
                        "so_vao_so": new_row_vals[2],
                        "ho_ten": new_row_vals[4],
                        "ngay_sinh": new_row_vals[8],
                        "ten_truong": new_row_vals[13],
                    })
                    excel_row_indices.append(new_r)

                    audit_entries.append({
                        "row_num": new_r,
                        "type": "Bổ sung mới",
                        "student_name": ho_ten,
                        "field_name": "Toàn bộ dòng",
                        "old_val": "---",
                        "new_val": f"STT {max_existing_stt} - {ho_ten}",
                        "image_name": img_p.name,
                        "status": "Đã thêm mới",
                        "note": rec.get("ly_do_nghi_ngo") or "Học sinh mới ghi thêm trên sổ"
                    })
                    log(f"    ➕ Thêm mới dòng {new_r} (STT {max_existing_stt}): {ho_ten}.", "success")

                else:
                    # KHÔNG TÌM THẤY VÀ KHÔNG BẬT TỰ ĐỘNG THÊM MỚI
                    unmatched_count += 1
                    audit_entries.append({
                        "row_num": "---",
                        "type": "Không khớp",
                        "student_name": ho_ten,
                        "field_name": "---",
                        "old_val": "---",
                        "new_val": "---",
                        "image_name": img_p.name,
                        "status": "Không tìm thấy",
                        "note": match_note
                    })
                    log(f"    ⚠️ Không tìm thấy vị trí của học sinh '{ho_ten}' ({match_note}).", "warning")

            if stop_process:
                break

        # 4. TẠO HOẶC CẬP NHẬT SHEET 'Lich_Su_Thay_Doi'
        if "Lich_Su_Thay_Doi" in wb.sheetnames:
            del wb["Lich_Su_Thay_Doi"]

        ws_audit = wb.create_sheet(title="Lich_Su_Thay_Doi")
        audit_headers = [
            ("stt", "STT", 8),
            ("row_num", "Dòng Excel", 12),
            ("type", "Loại thao tác", 16),
            ("student_name", "Họ và tên", 26),
            ("field_name", "Trường thay đổi", 18),
            ("old_val", "Giá trị cũ trong file", 24),
            ("new_val", "Giá trị mới (từ ảnh)", 26),
            ("image_name", "Tệp ảnh nguồn", 22),
            ("status", "Trạng thái", 18),
            ("note", "Ghi chú đối soát", 32),
        ]

        # Tiêu đề bảng lịch sử
        HDR_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
        HDR_FONT = Font(name="Times New Roman", size=11, bold=True, color="FFFFFF")

        for col_i, (h_key, h_label, h_w) in enumerate(audit_headers, start=1):
            c = ws_audit.cell(row=1, column=col_i, value=h_label)
            c.fill = HDR_FILL
            c.font = HDR_FONT
            c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            c.border = BORDER_THIN
            col_letter = get_column_letter(col_i)
            ws_audit.column_dimensions[col_letter].width = h_w

        # Ghi các dòng lịch sử
        for a_idx, entry in enumerate(audit_entries, start=1):
            r_a = a_idx + 1
            ws_audit.cell(row=r_a, column=1, value=a_idx).alignment = Alignment(horizontal="center")
            ws_audit.cell(row=r_a, column=2, value=entry["row_num"]).alignment = Alignment(horizontal="center")
            ws_audit.cell(row=r_a, column=3, value=entry["type"]).alignment = Alignment(horizontal="center")
            ws_audit.cell(row=r_a, column=4, value=entry["student_name"])
            ws_audit.cell(row=r_a, column=5, value=entry["field_name"])
            ws_audit.cell(row=r_a, column=6, value=entry["old_val"])
            ws_audit.cell(row=r_a, column=7, value=entry["new_val"])
            ws_audit.cell(row=r_a, column=8, value=entry["image_name"])
            ws_audit.cell(row=r_a, column=9, value=entry["status"]).alignment = Alignment(horizontal="center")
            ws_audit.cell(row=r_a, column=10, value=entry["note"])

            for c_i in range(1, 11):
                cell_item = ws_audit.cell(row=r_a, column=c_i)
                cell_item.border = BORDER_THIN
                cell_item.font = Font(name="Times New Roman", size=11)
                if entry["type"] == "Sửa đổi":
                    cell_item.fill = PatternFill(start_color="FFF9E6", end_color="FFF9E6", fill_type="solid")
                elif entry["type"] == "Bổ sung mới":
                    cell_item.fill = PatternFill(start_color="F2F9F2", end_color="F2F9F2", fill_type="solid")
                elif entry["type"] == "Không khớp":
                    cell_item.fill = PatternFill(start_color="FDE8E8", end_color="FDE8E8", fill_type="solid")

        # Lưu tệp kết quả
        if output_path:
            final_out_path = Path(output_path)
        else:
            final_out_path = excel_p.parent / f"{excel_p.stem}_DaCapNhat_{time.strftime('%Y%m%d_%H%M%S')}.xlsx"

        final_out_path.parent.mkdir(parents=True, exist_ok=True)
        wb.save(final_out_path)

        if progress_callback:
            progress_callback(100, 100, "Hoàn tất cập nhật.")

        log(
            f"🎉 HOÀN TẤT CẬP NHẬT: Đã sửa {modified_students_count} học sinh, thêm mới {added_students_count} học sinh "
            f"từ {success_images}/{total_images} ảnh. ➔ Đã lưu: {final_out_path.name}",
            "success"
        )

        return {
            "status": "success",
            "total_images": total_images,
            "success_images": success_images,
            "modified_students": modified_students_count,
            "added_students": added_students_count,
            "unmatched_records": unmatched_count,
            "total_changes": len(audit_entries),
            "output_file": final_out_path,
        }
