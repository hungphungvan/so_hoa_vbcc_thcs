import os
import sys
import io
import re
import copy
import unicodedata
from pathlib import Path
import openpyxl
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

from src.lookup_manager import LookupManager, get_data_dir
from src.normalizer import (
    clean_whitespace,
    is_cell_empty,
    vietnamese_title_case,
    normalize_birthplace,
    normalize_ethnic,
    normalize_nationality,
    format_date_to_text,
    check_dob_validity,
    parse_nien_khoa,
    infer_cohort_from_birth_year,
    find_mode_birth_year,
    fuzzy_match_school,
    auto_fill_graduation_dates_by_school_and_cohort,
    determine_status_and_reason
)
from src.exporter import DataExporter

SUMMARY_KEYWORDS = [
    "danh sách này có", "danh sách có", "danh sách gồm",
    "tổng số", "tổng cộng", "số học sinh",
    "công nhận tốt nghiệp", "người lập", "người ký", "hiệu trưởng", "thủ trưởng"
]

def is_valid_stt(val):
    """Kiểm tra ô có phải là Số thứ tự hợp lệ của học sinh (1, 2, 3...)."""
    if val is None or isinstance(val, bool):
        return False
    if isinstance(val, (int, float)):
        import math
        if math.isnan(val):
            return False
        return val > 0 and float(val).is_integer()
    s = str(val).strip()
    s_clean = re.sub(r'\.0+$', '', s)
    return s_clean.isdigit() and int(s_clean) > 0

def is_summary_or_footer_row(row_cells):
    """Kiểm tra xem dòng có phải là dòng tổng kết, chân trang ký tên hay không."""
    for cell in row_cells[:15]:
        val = str(cell.value or '').strip().lower()
        if any(kw in val for kw in SUMMARY_KEYWORDS):
            return True
    return False

def get_cached_row_fonts(ws, reference_row_idx, max_col=28):
    """Trích xuất và lưu sẵn định dạng Font từ dòng mẫu."""
    cached = []
    for col in range(1, max_col + 1):
        cell = ws.cell(row=reference_row_idx, column=col)
        if cell.has_style and cell.font:
            cached.append(Font(name=cell.font.name, size=cell.font.size))
        else:
            cached.append(None)
    return cached

def apply_cached_fonts(ws, target_row_idx, cached_fonts):
    """Gán font từ bộ nhớ cache vào dòng mục tiêu."""
    for col_idx, font in enumerate(cached_fonts, start=1):
        if font:
            ws.cell(row=target_row_idx, column=col_idx).font = font

def get_file_format_info(ws):
    """
    Nhận diện định dạng sheet:
    - Quét tối đa 30 dòng đầu tìm dòng tiêu đề chứa 'STT' hoặc 'Họ và tên'.
    - Phân loại: 'AE' (31 cột -> loại bỏ Q, R, S) hoặc 'AB' (28 cột chuẩn).
    - Trả về: (format_type, headers, header_row_idx)
    """
    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=30, values_only=True), start=1):
        if not row:
            continue
        row_str = [clean_whitespace(str(cell)) if cell is not None else "" for cell in row]
        row_lower = [s.lower() for s in row_str]

        has_stt = any(s in ["stt", "số tt", "số thứ tự", "stt.", "tt", "tt."] for s in row_lower[:5])
        has_name = any(("họ" in s and "tên" in s) for s in row_lower)

        if has_stt or has_name:
            headers = row_str
            col_idx = -1
            for idx, h in enumerate(row_lower):
                if "tên đơn vị" in h or h == "đơn vị":
                    col_idx = idx
                    break

            if col_idx >= 29:
                return "AE", headers, r_idx
            elif 25 <= col_idx <= 28:
                return "AB", headers, r_idx
            elif col_idx == -1 and has_name:
                non_empty_cols = sum(1 for c in headers if c)
                if non_empty_cols >= 30:
                    return "AE", headers, r_idx
                elif non_empty_cols >= 26:
                    return "AB", headers, r_idx
                return "INVALID", headers, r_idx
            else:
                return "INVALID", headers, r_idx

    return "INVALID", [], None

def extract_cohort_from_text(text: str) -> str:
    """Trích xuất chuỗi niên khóa dạng 2018-2022, 2018_2022, 18-22, 2020 từ chuỗi text (file name / sheet name)."""
    if not text:
        return ""
    # Mẫu 4 số - 4 số: 2018-2022, 2018_2022, 2018 - 2022
    m = re.search(r'\b(19\d{2}|20\d{2})\s*[-_–]\s*(19\d{2}|20\d{2})\b', text)
    if m:
        return f"{m.group(1)}-{m.group(2)}"
    # Mẫu 2 số - 2 số: 18-22
    m2 = re.search(r'\b(\d{2})\s*[-_–]\s*(\d{2})\b', text)
    if m2:
        y1 = int(m2.group(1))
        y2 = int(m2.group(2))
        full_y1 = (1900 + y1) if y1 >= 80 else (2000 + y1)
        full_y2 = (1900 + y2) if y2 >= 80 else (2000 + y2)
        if full_y2 > full_y1 and (full_y2 - full_y1) in (3, 4, 5):
            return f"{full_y1}-{full_y2}"
    # Mẫu 1 năm duy nhất: vd 2022
    m3 = re.search(r'\b(19\d{2}|20\d{2})\b', text)
    if m3:
        y_val = int(m3.group(1))
        return f"{y_val - 4}-{y_val}"
    return ""

def match_school_from_text(text: str, allowed_schools: list[str]) -> str:
    """Tìm trường khớp trong chuỗi văn bản (tên file, tên thư mục, tên sheet). Hỗ trợ chuẩn xác tiếng Việt có dấu, ký tự & và Unicode NFD/NFC."""
    if not text or not allowed_schools:
        return ""
    text_nfc = unicodedata.normalize('NFC', str(text))
    text_clean = re.sub(r'[_\-\.]', ' ', text_nfc).lower()
    for sch in allowed_schools:
        sch_nfc = unicodedata.normalize('NFC', str(sch)).lower()
        if sch_nfc in text_clean:
            return sch
    # Fuzzy match nếu gần khớp
    close = fuzzy_match_school(text_clean, allowed_schools, n=1, cutoff=0.75)
    if close:
        return close[0]
    return ""


class ExcelProcessor:
    """Bộ xử lý đọc, chuẩn hóa và xuất file Excel theo đúng quy chuẩn 28 cột của Bộ GD&ĐT."""
    def __init__(self, lookup_mgr: LookupManager, template_path: Path | str | None = None):
        self.lookup_mgr = lookup_mgr
        if template_path is None:
            self.template_path = get_data_dir() / "Mau_DS_Cap_Bang_THCS.xlsx"
        else:
            self.template_path = Path(template_path)

        if not self.template_path.exists():
            raise FileNotFoundError(f"Không tìm thấy file mẫu template: {self.template_path}")

        # Nạp template bytes và font cache
        with open(self.template_path, "rb") as f:
            self.template_bytes = f.read()

        wb_tpl = openpyxl.load_workbook(io.BytesIO(self.template_bytes))
        ws_tpl = wb_tpl.active
        self.cached_fonts = get_cached_row_fonts(ws_tpl, ws_tpl.max_row, max_col=28)
        wb_tpl.close()

        self.yellow_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
        self.exporter = DataExporter(self.template_bytes, self.cached_fonts, self.yellow_fill)

    def process_file(
        self,
        file_path: Path,
        unit_code: str,
        unit_name: str,
        allowed_schools: list[str],
        output_dir: Path | None = None,
        dialog_callback=None,
        log_callback=None,
        cache_state=None,
        save_individual_file: bool = False
    ) -> tuple[list[dict], dict]:
        """
        Trích xuất và chuẩn hóa học sinh từ 1 tệp Excel:
        - Tự động nhận diện Tên trường và Niên khóa từ Context (Thư mục / Tên file / Sheet).
        - Xử lý tương tác theo khối đối với Trường, Niên khóa, Ngày sinh.
        - Đính kèm tọa độ nguồn gốc vào Cột X (Ghi chú): [Nguồn: File | Sheet | Dòng].
        - Trả về: (file_students, stats)
        """
        def log(msg, level="info"):
            if log_callback:
                log_callback(msg, level)
            else:
                print(msg)

        if cache_state is None:
            cache_state = {
                "school_cache": {},
                "cohort_cache": {},
                "skip_all_dob": False
            }

        school_cache = cache_state.setdefault("school_cache", {})
        cohort_cache = cache_state.setdefault("cohort_cache", {})

        stats = {
            "file_name": file_path.name,
            "total_students": 0,
            "da_so_hoa": 0,
            "chua_day_du": 0,
            "chua_so_hoa": 0,
            "output_file": None,
            "error_file": None
        }

        try:
            wb_in = openpyxl.load_workbook(file_path, data_only=True)
        except Exception as e:
            log(f"❌ Không thể mở file {file_path.name}: {e}", "error")
            return [], stats

        all_file_students = []

        # Heuristic 1 & 2: Nhận diện trường và niên khóa từ đường dẫn và tên tệp
        # Trường từ thư mục cha (Cách 2) hoặc từ tên file (Cách 1)
        fallback_sch_parent = match_school_from_text(file_path.parent.name, allowed_schools)
        fallback_sch_file = match_school_from_text(file_path.stem, allowed_schools)
        file_level_school = fallback_sch_parent or fallback_sch_file

        # Niên khóa từ tên tệp (Cách 2 hoặc 3)
        file_level_cohort = extract_cohort_from_text(file_path.stem)

        try:
            for ws_in in wb_in.worksheets:
                if ws_in.sheet_state != 'visible':
                    continue

                fmt_type, headers, header_r_idx = get_file_format_info(ws_in)
                if fmt_type == "INVALID":
                    log(f"  -> Bỏ qua Sheet '{ws_in.title}' (Cấu trúc không khớp 28 hoặc 31 cột)", "warning")
                    continue

                remove_qrs = (fmt_type == "AE")
                log(f"  -> Quét Sheet '{ws_in.title}': {'Dạng AE 31 cột (cắt Q, R, S)' if remove_qrs else 'Dạng AB 28 cột chuẩn'}", "info")

                # Heuristic 3: Nhận diện trường hoặc niên khóa từ tên Sheet
                sheet_level_school = match_school_from_text(ws_in.title, allowed_schools)
                sheet_level_cohort = extract_cohort_from_text(ws_in.title)

                effective_context_school = sheet_level_school or file_level_school
                effective_context_cohort = sheet_level_cohort or file_level_cohort

                start_extracting = False
                empty_consecutive = 0
                sheet_students = []

                for row in ws_in.iter_rows():
                    r_idx = row[0].row if hasattr(row[0], 'row') else 0
                    if header_r_idx is not None and r_idx <= header_r_idx:
                        continue

                    cell_a_val = row[0].value
                    cell_a_str = str(cell_a_val).strip() if cell_a_val is not None else ""

                    name_cell_val = row[4].value if len(row) > 4 else None
                    name_cell_str = clean_whitespace(str(name_cell_val)) if name_cell_val is not None else ""
                    has_letters = bool(re.search(r'[a-zA-ZÀ-ỹ]', name_cell_str))
                    is_row_empty = not any(c.value is not None and str(c.value).strip() != "" for c in row[:15])

                    if not start_extracting:
                        is_stt_one = bool(re.match(r'^0*1(\.0*)?$', cell_a_str))
                        if (is_stt_one and has_letters) or (has_letters and name_cell_str.lower() not in ["họ và tên", "họ tên", "họ, chữ đệm và tên"]):
                            start_extracting = True

                    if start_extracting:
                        if is_row_empty:
                            empty_consecutive += 1
                            if empty_consecutive >= 20:
                                break
                            continue
                        else:
                            empty_consecutive = 0

                        if is_summary_or_footer_row(row):
                            continue

                        # Bắt buộc: STT hợp lệ và Họ tên có chữ cái
                        if not is_valid_stt(cell_a_val) or is_cell_empty(name_cell_val) or not has_letters:
                            continue

                        row_vals = [cell.value for cell in row]
                        req_len = 31 if remove_qrs else 28
                        while len(row_vals) < req_len:
                            row_vals.append(None)

                        if remove_qrs:
                            # Cắt bỏ 3 cột Q, R, S (index 16, 17, 18)
                            new_row = (row_vals[:16] + row_vals[19:])[:28]
                        else:
                            new_row = row_vals[:28]

                        while len(new_row) < 28:
                            new_row.append(None)

                        clean_stt = int(cell_a_val) if isinstance(cell_a_val, (int, float)) and float(cell_a_val).is_integer() else str(cell_a_val or '')

                        raw_dob = new_row[8]
                        raw_school = new_row[13]
                        raw_cohort = new_row[14]
                        raw_grad = new_row[18]

                        # Áp dụng heuristic fallback nếu ô trống
                        init_school = clean_whitespace(str(raw_school)) if not is_cell_empty(raw_school) else (effective_context_school or "")
                        init_cohort = clean_whitespace(str(raw_cohort)) if not is_cell_empty(raw_cohort) else (effective_context_cohort or "")

                        student_dict = {
                            "file_name": file_path.name,
                            "sheet_name": ws_in.title,
                            "row_idx": r_idx,
                            "clean_stt": clean_stt,
                            "raw_name": name_cell_str,
                            "name": vietnamese_title_case(name_cell_str),
                            "raw_dob": raw_dob,
                            "dob": format_date_to_text(raw_dob),
                            "raw_school": raw_school,
                            "school": init_school,
                            "raw_cohort": raw_cohort,
                            "cohort": init_cohort,
                            "raw_grad": raw_grad,
                            "grad_date": clean_whitespace(str(raw_grad)) if raw_grad is not None else "",
                            "row_values": new_row
                        }
                        sheet_students.append(student_dict)

                total_sheet = len(sheet_students)
                if total_sheet == 0:
                    continue

                log(f"  -> Trích xuất được {total_sheet} học sinh từ Sheet '{ws_in.title}'.", "info")

                # ==============================================================
                # GIAI ĐOẠN A: XỬ LÝ KHỐI THIẾU/SAI TÊN TRƯỜNG (CỘT N)
                # ==============================================================
                i = 0
                while i < total_sheet:
                    st_item = sheet_students[i]
                    curr_sch = st_item["school"]
                    is_valid_sch = any(curr_sch.lower() == s.lower() for s in allowed_schools) if allowed_schools else bool(curr_sch)

                    if not is_valid_sch:
                        start_k = i
                        raw_val_key = clean_whitespace(str(st_item["raw_school"]))

                        while i < total_sheet:
                            nxt = sheet_students[i]
                            nxt_sch = nxt["school"]
                            nxt_is_valid = any(nxt_sch.lower() == s.lower() for s in allowed_schools) if allowed_schools else bool(nxt_sch)
                            nxt_raw_key = clean_whitespace(str(nxt["raw_school"]))
                            if (not nxt_is_valid) and (nxt_raw_key == raw_val_key or (is_cell_empty(raw_val_key) and is_cell_empty(nxt_raw_key))):
                                i += 1
                            else:
                                break

                        end_k = i - 1
                        block_len = end_k - start_k + 1

                        if raw_val_key and raw_val_key in school_cache:
                            applied_sch = school_cache[raw_val_key]
                            for idx_in_blk in range(start_k, end_k + 1):
                                sheet_students[idx_in_blk]["school"] = applied_sch
                            log(f"   ⚡ Tự động áp dụng Tên trường đã nhớ cho {block_len} học sinh (STT {sheet_students[start_k]['clean_stt']}->{sheet_students[end_k]['clean_stt']}): '{applied_sch}'", "info")
                            continue

                        # Ngữ cảnh +-5 học sinh
                        ctx_start = max(0, start_k - 5)
                        ctx_end = min(total_sheet, end_k + 6)
                        context_rows = []
                        for c_idx in range(ctx_start, ctx_end):
                            item_c = sheet_students[c_idx]
                            context_rows.append({
                                "stt": item_c["clean_stt"],
                                "name": item_c["name"],
                                "school": item_c["school"],
                                "cohort": item_c["cohort"],
                                "grad": item_c["grad_date"],
                                "is_in_block": (start_k <= c_idx <= end_k)
                            })

                        # Gợi ý thông minh: Lấy trường từ Context fallback -> liền trước -> fuzzy match
                        sugg_school = effective_context_school
                        if not sugg_school and start_k > 0:
                            prev_sch = sheet_students[start_k - 1]["school"]
                            if prev_sch and (not allowed_schools or any(prev_sch.lower() == s.lower() for s in allowed_schools)):
                                sugg_school = prev_sch
                        if not sugg_school and allowed_schools:
                            if raw_val_key:
                                close_matches = fuzzy_match_school(raw_val_key, allowed_schools, n=1)
                                if close_matches:
                                    sugg_school = close_matches[0]
                            if not sugg_school:
                                sugg_school = allowed_schools[0]

                        chosen_school = sugg_school
                        if dialog_callback:
                            block_info = {
                                "file_name": file_path.name,
                                "sheet_name": ws_in.title,
                                "start_stt": sheet_students[start_k]["clean_stt"],
                                "end_stt": sheet_students[end_k]["clean_stt"],
                                "count": block_len,
                                "raw_val": raw_val_key if raw_val_key else "(Trống)",
                                "context_rows": context_rows
                            }
                            res_sch, rem = dialog_callback("school_block", block_info, allowed_schools, sugg_school)
                            if res_sch:
                                chosen_school = res_sch
                            if rem and raw_val_key:
                                school_cache[raw_val_key] = chosen_school
                        else:
                            if sugg_school and raw_val_key:
                                school_cache[raw_val_key] = sugg_school

                        for idx_in_blk in range(start_k, end_k + 1):
                            sheet_students[idx_in_blk]["school"] = chosen_school
                    else:
                        for s_std in allowed_schools:
                            if s_std.lower() == curr_sch.lower():
                                st_item["school"] = s_std
                                break
                        i += 1

                # ==============================================================
                # GIAI ĐOẠN B: XỬ LÝ KHỐI THIẾU NIÊN KHÓA (CỘT O)
                # ==============================================================
                i = 0
                while i < total_sheet:
                    st_item = sheet_students[i]
                    is_valid_nk, y1, y2, fmt_nk = parse_nien_khoa(st_item["cohort"])

                    if not is_valid_nk:
                        start_k = i
                        raw_nk_key = clean_whitespace(str(st_item["raw_cohort"]))

                        while i < total_sheet:
                            nxt = sheet_students[i]
                            nxt_valid, _, _, _ = parse_nien_khoa(nxt["cohort"])
                            nxt_nk_key = clean_whitespace(str(nxt["raw_cohort"]))
                            if (not nxt_valid) and (nxt_nk_key == raw_nk_key or (is_cell_empty(raw_nk_key) and is_cell_empty(nxt_nk_key))):
                                i += 1
                            else:
                                break

                        end_k = i - 1
                        block_len = end_k - start_k + 1

                        if raw_nk_key and raw_nk_key in cohort_cache:
                            applied_nk = cohort_cache[raw_nk_key]
                            for idx_in_blk in range(start_k, end_k + 1):
                                sheet_students[idx_in_blk]["cohort"] = applied_nk
                            log(f"   ⚡ Tự động áp dụng Niên khóa đã nhớ cho {block_len} học sinh: '{applied_nk}'", "info")
                            continue

                        # Ngữ cảnh +-5 học sinh
                        ctx_start = max(0, start_k - 5)
                        ctx_end = min(total_sheet, end_k + 6)
                        context_rows = []
                        for c_idx in range(ctx_start, ctx_end):
                            item_c = sheet_students[c_idx]
                            context_rows.append({
                                "stt": item_c["clean_stt"],
                                "name": item_c["name"],
                                "dob": item_c["dob"],
                                "cohort": item_c["cohort"],
                                "grad": item_c["grad_date"],
                                "is_in_block": (start_k <= c_idx <= end_k)
                            })

                        # Gợi ý thông minh: Tính từ năm sinh phổ biến của khối hoặc context
                        block_slice = sheet_students[start_k:end_k + 1]
                        mode_yr = find_mode_birth_year(block_slice)
                        if not mode_yr and start_k > 0:
                            mode_yr = find_mode_birth_year(sheet_students[max(0, start_k - 5):start_k])

                        sugg_nk = effective_context_cohort
                        if not sugg_nk and mode_yr:
                            _, _, sugg_nk = infer_cohort_from_birth_year(mode_yr)
                        elif not sugg_nk and start_k > 0:
                            prev_valid, _, _, prev_fmt = parse_nien_khoa(sheet_students[start_k - 1]["cohort"])
                            if prev_valid:
                                sugg_nk = prev_fmt

                        chosen_nk = sugg_nk if sugg_nk else fmt_nk
                        if dialog_callback:
                            block_info = {
                                "file_name": file_path.name,
                                "sheet_name": ws_in.title,
                                "start_stt": sheet_students[start_k]["clean_stt"],
                                "end_stt": sheet_students[end_k]["clean_stt"],
                                "count": block_len,
                                "raw_val": raw_nk_key if raw_nk_key else "(Trống)",
                                "context_rows": context_rows
                            }
                            res_nk, rem = dialog_callback("cohort_block", block_info, sugg_nk, mode_yr)
                            if res_nk:
                                chosen_nk = res_nk
                            if rem and raw_nk_key:
                                cohort_cache[raw_nk_key] = chosen_nk
                        else:
                            if sugg_nk and raw_nk_key:
                                cohort_cache[raw_nk_key] = sugg_nk

                        for idx_in_blk in range(start_k, end_k + 1):
                            sheet_students[idx_in_blk]["cohort"] = chosen_nk
                    else:
                        st_item["cohort"] = fmt_nk
                        i += 1

                # ==============================================================
                # GIAI ĐOẠN C: XỬ LÝ NGÀY SINH LỖI (CỘT I)
                # ==============================================================
                for idx_st, st_item in enumerate(sheet_students):
                    norm_dob = st_item["dob"]
                    is_v_dob, err_t, err_d = check_dob_validity(norm_dob)
                    st_item["is_valid_dob"] = is_v_dob
                    st_item["dob_err_desc"] = err_d

                    if not is_v_dob and not cache_state.get("skip_all_dob", False):
                        if not is_cell_empty(norm_dob):
                            ctx_start = max(0, idx_st - 5)
                            ctx_end = min(total_sheet, idx_st + 6)
                            context_rows = []
                            for c_idx in range(ctx_start, ctx_end):
                                item_c = sheet_students[c_idx]
                                context_rows.append({
                                    "stt": item_c["clean_stt"],
                                    "name": item_c["name"],
                                    "dob": item_c["dob"],
                                    "cohort": item_c["cohort"],
                                    "grad": item_c["grad_date"],
                                    "is_current": (c_idx == idx_st)
                                })

                            sugg_dob = ""
                            m_nk = re.search(r'(\d{4})$', str(st_item["cohort"]))
                            if m_nk:
                                y_expected = int(m_nk.group(1)) - 15
                                m_dm = re.search(r'\b(\d{1,2})[/-](\d{1,2})\b', str(st_item["raw_dob"] or norm_dob))
                                if m_dm:
                                    d_v, m_v = int(m_dm.group(1)), int(m_dm.group(2))
                                    if 1 <= d_v <= 31 and 1 <= m_v <= 12:
                                        cand = f"{d_v:02d}/{m_v:02d}/{y_expected}"
                                        c_v, _, _ = check_dob_validity(cand)
                                        if c_v:
                                            sugg_dob = cand

                            if dialog_callback:
                                student_info = {
                                    "file_name": file_path.name,
                                    "sheet_name": ws_in.title,
                                    "stt": st_item["clean_stt"],
                                    "name": st_item["name"],
                                    "raw_val": st_item["raw_dob"],
                                    "curr_norm": norm_dob,
                                    "err_desc": err_d,
                                    "context_rows": context_rows
                                }
                                res_dob, skip_all_flag = dialog_callback("date_fix", student_info, "Ngày sinh", sugg_dob)
                                if skip_all_flag:
                                    cache_state["skip_all_dob"] = True
                                if res_dob:
                                    st_item["dob"] = res_dob
                                    is_v_after, _, err_d_after = check_dob_validity(res_dob)
                                    st_item["is_valid_dob"] = is_v_after
                                    st_item["dob_err_desc"] = err_d_after

                # ==============================================================
                # GIAI ĐOẠN D: ĐIỀN NGÀY CẤP BẰNG THEO (TRƯỜNG, NIÊN KHÓA)
                # ==============================================================
                auto_fill_graduation_dates_by_school_and_cohort(sheet_students)

                # ==============================================================
                # GIAI ĐOẠN E: ÁP DỤNG DANH MỤC & ĐÍNH KÈM TỌA ĐỘ NGUỒN CỘT X
                # ==============================================================
                for st_item in sheet_students:
                    row_vals = list(st_item["row_values"])

                    # Cột D (Index 3): Tên văn bằng
                    row_vals[3] = "Bằng tốt nghiệp trung học cơ sở"
                    # Cột E (Index 4): Họ và tên Title Case
                    row_vals[4] = st_item["name"]
                    # Cột I (Index 8): Ngày sinh chuẩn dd/MM/yyyy
                    row_vals[8] = st_item["dob"]
                    # Cột J (Index 9): Giới tính
                    row_vals[9] = self.lookup_mgr.map_gender(row_vals[9])
                    # Cột K (Index 10): Dân tộc (Title Case)
                    row_vals[10] = normalize_ethnic(row_vals[10], self.lookup_mgr)
                    # Cột L (Index 11): Quốc tịch (Title Case)
                    row_vals[11] = normalize_nationality(row_vals[11], self.lookup_mgr)
                    # Cột M (Index 12): Nơi sinh (Title Case + Thành phố/Thị xã/Thị trấn)
                    row_vals[12] = normalize_birthplace(row_vals[12], self.lookup_mgr)
                    # Cột N (Index 13): Tên trường đã học đã chuẩn hóa
                    row_vals[13] = st_item["school"]
                    # Cột O (Index 14): Niên khóa đã chuẩn hóa
                    row_vals[14] = st_item["cohort"]
                    # Cột S (Index 18): Ngày tháng năm cấp bằng
                    row_vals[18] = st_item["grad_date"]
                    # Cột V (Index 21): Người ký bằng Title Case
                    if row_vals[21]:
                        row_vals[21] = vietnamese_title_case(str(row_vals[21]))

                    # Cột X (Index 23): Ghi chú -> ĐÍNH KÈM TỌA ĐỘ NGUỒN GỐC
                    source_tag = f"[Nguồn: {file_path.name} | Sheet: {ws_in.title} | Dòng {st_item['row_idx']}]"
                    orig_note = str(row_vals[23]).strip() if row_vals[23] is not None else ""
                    if orig_note:
                        row_vals[23] = f"{orig_note} | {source_tag}"
                    else:
                        row_vals[23] = source_tag

                    # Cột Y (Index 24) & Cột Z (Index 25): Trạng thái số hóa & Nguyên nhân
                    status_y, reason_z = determine_status_and_reason(
                        st_item["raw_dob"],
                        st_item["dob"],
                        st_item["is_valid_dob"],
                        st_item["dob_err_desc"]
                    )
                    row_vals[24] = status_y
                    row_vals[25] = reason_z if reason_z else None

                    # Cột AA (Index 26): Mã đơn vị
                    row_vals[26] = unit_code
                    # Cột AB (Index 27): Tên đơn vị
                    row_vals[27] = unit_name

                    st_item["final_row"] = row_vals
                    all_file_students.append(st_item)

                    if status_y == "Đã số hóa đầy đủ, chính xác":
                        stats["da_so_hoa"] += 1
                    elif status_y == "Số hoá chưa đầy đủ, chính xác":
                        stats["chua_day_du"] += 1
                    else:
                        stats["chua_so_hoa"] += 1

            wb_in.close()
        except Exception as e:
            log(f"❌ Lỗi khi xử lý nội dung file {file_path.name}: {e}", "error")
            return [], stats

        stats["total_students"] = len(all_file_students)

        # Lưu file riêng lẻ nếu được yêu cầu (ví dụ khi người dùng chọn xử lý 1 file duy nhất)
        if save_individual_file and output_dir and all_file_students:
            output_dir.mkdir(parents=True, exist_ok=True)
            out_wb = openpyxl.load_workbook(io.BytesIO(self.template_bytes))
            out_ws = out_wb.active
            out_ws.title = "Data"
            if out_ws.max_row > 2:
                out_ws.delete_rows(3, out_ws.max_row - 2)

            for r_idx, st_item in enumerate(all_file_students, start=1):
                row_vals = list(st_item["final_row"])
                row_vals[0] = r_idx
                out_ws.append(row_vals)
                curr_row = out_ws.max_row
                apply_cached_fonts(out_ws, curr_row, self.cached_fonts)
                for col in range(1, 28):
                    if is_cell_empty(row_vals[col]):
                        out_ws.cell(row=curr_row, column=col + 1).fill = self.yellow_fill

            out_file_path = output_dir / file_path.name
            try:
                out_wb.save(out_file_path)
                stats["output_file"] = out_file_path
                log(f"  ✅ Đã lưu file chuẩn hóa: Da_Chuan_Hoa/{out_file_path.name}", "success")
            except PermissionError:
                log(f"❌ Không thể ghi file '{out_file_path.name}' do đang mở trong Excel.", "error")
            finally:
                out_wb.close()

        return all_file_students, stats

    def process_all_files(
        self,
        files: list[Path],
        unit_info: dict | str,
        unit_name: str = "",
        allowed_schools: list[str] | None = None,
        output_dir: Path | None = None,
        export_options: dict | None = None,
        dialog_callback=None,
        log_callback=None,
        progress_callback=None,
        check_cancel=None
    ) -> tuple[list[dict], dict]:
        """
        Xử lý toàn diện danh sách tệp Excel cho một đơn vị:
        - Quét và trích xuất học sinh từ tất cả các file.
        - Gom nhóm và tối ưu hóa ngày cấp bằng toàn đơn vị.
        - Xuất các file gộp, tách, file DS_Loi duy nhất, file Du_Lieu_Goc và Danh sách tài liệu số hoá (Mô-đun 2).
        """
        def log(msg, level="info"):
            if log_callback:
                log_callback(msg, level)
            else:
                print(msg)

        if isinstance(unit_info, dict):
            eff_unit_info = unit_info
            u_code = unit_info.get("code", "")
            u_name = unit_info.get("name", "")
            u_schools = unit_info.get("schools", [])
        else:
            u_code = str(unit_info)
            u_name = unit_name
            u_schools = allowed_schools or []
            eff_unit_info = self.lookup_mgr.get_unit_info(u_name) or {
                "code": u_code,
                "name": u_name,
                "schools": u_schools,
                "school_map": {}
            }

        all_unit_students = []
        overall_stats = {
            "total_files": len(files),
            "total_students": 0,
            "da_so_hoa": 0,
            "chua_day_du": 0,
            "chua_so_hoa": 0,
            "export_results": {}
        }

        cache_state = {
            "school_cache": {},
            "cohort_cache": {},
            "skip_all_dob": False
        }

        is_single_file = (len(files) == 1)

        for idx, f_path in enumerate(files, start=1):
            if check_cancel and check_cancel():
                log("⏹ Quá trình xử lý đã bị hủy theo yêu cầu của người dùng.", "warning")
                break

            log(f"\n[{idx}/{len(files)}] ĐANG XỬ LÝ: {f_path.name}...", "info")

            file_students, file_stats = self.process_file(
                file_path=f_path,
                unit_code=u_code,
                unit_name=u_name,
                allowed_schools=u_schools,
                output_dir=output_dir,
                dialog_callback=dialog_callback,
                log_callback=log_callback,
                cache_state=cache_state,
                save_individual_file=is_single_file
            )

            all_unit_students.extend(file_students)
            overall_stats["total_students"] += file_stats["total_students"]
            overall_stats["da_so_hoa"] += file_stats["da_so_hoa"]
            overall_stats["chua_day_du"] += file_stats["chua_day_du"]
            overall_stats["chua_so_hoa"] += file_stats["chua_so_hoa"]

            if progress_callback:
                progress_callback(idx, len(files), overall_stats)

        if not all_unit_students:
            log("⚠️ Không có học sinh nào được trích xuất từ các file nguồn.", "warning")
            return [], overall_stats

        # Tối ưu hóa điền ngày cấp bằng theo trường & niên khóa trên toàn bộ dữ liệu đơn vị
        if len(files) > 1:
            log("\n⚡ Đang tối ưu hóa đối khớp Ngày cấp bằng toàn đơn vị...", "info")
            auto_fill_graduation_dates_by_school_and_cohort(all_unit_students)
            for st in all_unit_students:
                st["final_row"][18] = str(st["grad_date"]) if st["grad_date"] else ""

        # Xuất các file gộp, tách, file DS_Loi duy nhất, file Dữ liệu gốc và Tài liệu số hóa
        log("\n📦 BẮT ĐẦU XUẤT CÁC TỆP DỮ LIỆU THEO TÙY CHỌN...", "info")
        export_res = self.exporter.export_all(
            all_students=all_unit_students,
            unit_info=eff_unit_info,
            output_dir=output_dir,
            options=export_options or {},
            log_callback=log_callback
        )
        overall_stats["export_results"] = export_res

        return all_unit_students, overall_stats
