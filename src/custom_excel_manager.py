from __future__ import annotations

import sys
import re
import math
from pathlib import Path
from typing import Any
import openpyxl
from openpyxl.utils import get_column_letter

from src.normalizer import (
    auto_convert_legacy_encoding,
    clean_whitespace,
    is_tcvn3,
    tcvn3_to_unicode
)
from src.exporter import sanitize_filename

def safe_print(msg: Any) -> None:
    try:
        print(msg)
    except Exception:
        try:
            if hasattr(sys.stdout, 'buffer') and sys.stdout.buffer:
                sys.stdout.buffer.write((str(msg) + "\n").encode('utf-8', errors='replace'))
            else:
                print(str(msg).encode('ascii', errors='replace').decode('ascii'))
        except Exception:
            pass


# Từ khóa nhận dạng dòng tiêu đề STT
STT_KEYWORDS = ["stt", "số tt", "số thứ tự", "stt.", "tt", "tt."]

# Từ khóa nhận dạng dòng tiêu đề Họ và tên
NAME_KEYWORDS = [
    "họ và tên", "họ tên", "họ, chữ đệm và tên", "họ và chữ đệm", "tên học sinh", "họ tên học sinh",
    "họ đệm", "họ, đệm", "họ và", "họ chữ đệm", "họ lót", "họ", "hodem", "ho dem",
    "tên", "ten", "tên thí sinh", "tên người học"
]

# Các từ khóa đặc trưng nhận diện mẫu chuẩn 28 cột (Mau_DS_Cap_Bang_THCS.xlsx)
STANDARD_28_SIGNATURES = [
    "stt", "số hiệu bằng", "số vào sổ gốc cấp văn bằng", "số vào sổ", "tên văn bằng",
    "họ, chữ đệm và tên", "mã người học", "số định danh cá nhân", "hộ chiếu",
    "ngày, tháng, năm sinh", "giới tính", "dân tộc", "quốc tịch", "nơi sinh",
    "tên trường đã học", "niên khóa", "niên khoá", "điểm các môn thi", "hội đồng thi",
    "địa danh nơi cơ quan cấp bằng đặt trụ sở", "ngày tháng năm cấp bằng", "tên cơ quan cấp bằng",
    "chức danh người ký bằng", "tình trạng văn bằng", "ghi chú", "trạng thái số hóa",
    "trạng thái số hoá", "thông tin chưa bảo đảm đầy đủ, chính xác và nguyên nhân",
    "mã đơn vị", "tên đơn vị"
]

# Các họ phổ biến của Việt Nam phục vụ Heuristic dò tìm
VIETNAMESE_SURNAMES = [
    "nguyễn", "nguyen", "trần", "tran", "lê", "le", "phạm", "pham", 
    "hoàng", "hoang", "huỳnh", "huynh", "phan", "vũ", "vu", "võ", "vo", 
    "đặng", "dang", "bùi", "bui", "đỗ", "do", "hồ", "ho", "ngô", "ngo", 
    "dương", "duong", "lý", "ly", "đinh", "dinh", "đoàn", "doan", "lâm", "lam"
]

def is_valid_stt(val) -> bool:
    """Kiểm tra ô có phải là Số thứ tự hợp lệ (1, 2, 3...)."""
    if val is None or isinstance(val, bool):
        return False
    if isinstance(val, (int, float)):
        if math.isnan(val):
            return False
        return val > 0 and float(val).is_integer()
    s = str(val).strip()
    s_lower = s.lower()
    if any(kw in s_lower for kw in ["stt", "số tt", "số thứ tự", "thứ tự", "tt."]):
        return False
    s_clean = re.sub(r'\.0+$', '', s)
    return s_clean.isdigit() and int(s_clean) > 0

RE_HAS_LETTER = re.compile(r'[a-zA-ZÀ-ỹ]')
NAME_EXCLUDE_KEYWORDS = (
    "họ và tên", "họ tên", "họ, chữ đệm", "họ và chữ đệm", "tên học sinh", 
    "họ tên học sinh", "người học", "thí sinh", "tiêu đề", "họ, tên", "tên người"
)

def is_valid_name(val) -> bool:
    """Kiểm tra ô có phải là Họ và tên người hợp lệ (chứa chữ cái, không rỗng, không phải header)."""
    if val is None or isinstance(val, (int, float, bool)):
        return False
    s = str(val).strip()
    if len(s) < 2 or s.isdigit():
        return False
    s_lower = s.lower()
    # Không được là tiêu đề lặp lại hoặc chứa các từ khóa cột tiêu đề
    if any(kw in s_lower for kw in NAME_EXCLUDE_KEYWORDS):
        return False
    # Bắt buộc phải chứa ký tự chữ cái tiếng Việt hoặc Latin
    return bool(RE_HAS_LETTER.search(s))

# Từ khóa nhận dạng sub-header trong tiêu đề đa tầng
SUB_HEADER_KEYWORDS = [
    "ngày", "tháng", "năm", "toán", "văn", "ngữ văn", "ngoại ngữ", "tiếng anh",
    "vật lý", "hóa học", "sinh học", "lịch sử", "địa lý", "gdcd", "nam", "nữ",
    "học lực", "hạnh kiểm", "tb", "tổng", "xếp loại", "chữ đệm", "tên", "họ", "đạt", "đỗ", "hỏng"
]

def _is_index_row(row_cells: list | tuple) -> bool:
    """Kiểm tra xem dòng có phải là dòng đánh số thứ tự cột (1), (2), (3)... hoặc 1, 2, 3... hay không."""
    non_empty = [c for c in row_cells if c is not None and str(c).strip() != ""]
    if not non_empty or len(non_empty) < 2:
        return False
    # Nếu dòng chứa tên người hợp lệ -> chắc chắn là dữ liệu, không phải dòng đánh số cột
    if any(is_valid_name(c) for c in non_empty):
        return False
    num_count = sum(1 for c in non_empty if re.match(r'^\(?\d{1,3}\)?$', str(c).strip()))
    return (num_count / len(non_empty)) >= 0.7

def is_header_duplicate_row(row: list | tuple, header_row_data: list | tuple | None = None, s_idx: int = -1, n_idx: int = -1) -> bool:
    """
    Kiểm tra xem một dòng dữ liệu có phải là dòng tiêu đề bị lặp lại hay không.
    Bảo vệ triệt để: loại bỏ 100% các dòng trùng nội dung tiêu đề khỏi danh sách dữ liệu con.
    """
    if not row:
        return True
    
    # 1. Kiểm tra ô STT xem có chứa từ khóa STT không (ví dụ ngắt trang lặp lại 'STT')
    if 0 <= s_idx < len(row) and row[s_idx] is not None:
        s_str = clean_whitespace(str(row[s_idx])).lower()
        if any(kw in s_str for kw in ["stt", "số tt", "số thứ tự", "thứ tự", "tt."]):
            return True

    # 2. Kiểm tra ô Họ tên xem có chứa từ khóa tiêu đề không
    if 0 <= n_idx < len(row) and row[n_idx] is not None:
        n_str = clean_whitespace(str(row[n_idx])).lower()
        if any(kw in n_str for kw in [
            "họ và tên", "họ tên", "họ, chữ đệm", "họ và chữ đệm", "tên học sinh",
            "họ tên học sinh", "người học", "thí sinh", "họ, tên", "tên người"
        ]):
            return True

    # 3. Nếu ô Họ tên chứa tên người thật hợp lệ -> chắc chắn là dòng học sinh, KHÔNG phải tiêu đề!
    if 0 <= n_idx < len(row) and row[n_idx] is not None:
        if is_valid_name(row[n_idx]):
            return False

    # 4. Kiểm tra xem dòng có phải là dòng đánh số thứ tự cột (1), (2), (3)... lặp lại
    if _is_index_row(row):
        return True

    # 5. So khớp trực tiếp với dòng tiêu đề chuẩn (header_row_data)
    if header_row_data:
        header_norm = [clean_whitespace(str(c or '')).lower() for c in header_row_data]
        row_norm = [clean_whitespace(str(c or '')).lower() for c in row]
        
        matches = 0
        total_filled = 0
        for c_idx in range(min(len(header_norm), len(row_norm))):
            h_val = header_norm[c_idx]
            r_val = row_norm[c_idx]
            if r_val:
                total_filled += 1
                if h_val and (r_val == h_val or (len(h_val) >= 3 and h_val in r_val)):
                    matches += 1
        
        # Nếu trùng >= 35% số ô hoặc có >= 3 ô tiêu đề trùng khớp -> đây là dòng tiêu đề lặp lại!
        if total_filled > 0 and ((matches / total_filled >= 0.35) or matches >= 3):
            return True

    return False

def is_cell_empty(val) -> bool:
    """Kiểm tra ô có bị rỗng hoàn toàn hay không."""
    if val is None:
        return True
    s = str(val).strip()
    return s == "" or s.lower() == "none"

def detect_header_and_columns(ws, max_scan_rows: int = 25) -> dict:
    """
    Quét tìm dòng tiêu đề và xác định tọa độ cột:
    - Hỗ trợ tiêu đề đơn dòng VÀ tiêu đề đa dòng (2 hoặc nhiều dòng merge với nhau).
    - Tự động quét ws.merged_cells và kết hợp tên cha - con (vd: 'Ngày tháng năm sinh - Ngày').
    - Dòng tiêu đề chính (header_row_idx: dòng cuối cùng của khối tiêu đề để dữ liệu bắt đầu sau đó)
    - Danh sách nhãn cột (raw_headers, formatted_headers)
    - Tọa độ cột STT (stt_col_idx, 0-based)
    - Tọa độ cột Họ và tên (name_col_idx, 0-based)
    """
    best_header_idx = None
    best_headers: list[str] = []
    stt_idx = 0
    name_idx = -1

    for r_idx, row in enumerate(ws.iter_rows(min_row=1, max_row=max_scan_rows, values_only=True), start=1):
        if not row:
            continue
        row_raw = [str(c or '').strip() for c in row]
        row_is_tcvn3 = any(is_tcvn3(s) for s in row_raw)
        row_str = []
        for c in row:
            s_raw = str(c or '').strip()
            if row_is_tcvn3 and any(ord(ch) > 127 for ch in s_raw):
                row_str.append(clean_whitespace(tcvn3_to_unicode(s_raw)))
            else:
                row_str.append(auto_convert_legacy_encoding(clean_whitespace(str(c or ''))))
        row_lower = [s.lower() for s in row_str]

        # 1. Nhận diện ưu tiên tuyệt đối: Cấu trúc mẫu chuẩn 28 cột (Mau_DS_Cap_Bang_THCS.xlsx)
        std_28_score = sum(1 for kw in STANDARD_28_SIGNATURES if any(kw in s for s in row_lower))
        if std_28_score >= 4:
            best_header_idx = r_idx
            best_headers = row_str
            for c_i, s in enumerate(row_lower):
                if any(kw == s or kw in s for kw in STT_KEYWORDS):
                    stt_idx = c_i
                    break
            for c_i, s in enumerate(row_lower):
                s_clean = s.strip()
                if any(kw in s_clean for kw in ["họ, chữ đệm", "họ và tên", "họ tên"]) and "người ký" not in s_clean:
                    name_idx = c_i
                    break
            break

        # 2. Kiểm tra tổng quát: xem dòng có chứa STT hoặc Họ và tên không
        has_stt = any(any(kw == s or kw in s for kw in STT_KEYWORDS) for s in row_lower[:6])
        has_name = any(any(kw in s for kw in NAME_KEYWORDS) for s in row_lower)
        non_empty = sum(1 for s in row_str if s)

        # Dòng tiêu đề thật sự phải có ít nhất 2 cột có nội dung và có STT hoặc Họ tên
        if (has_stt and has_name and non_empty >= 2) or ((has_stt or has_name) and non_empty >= 3):
            best_header_idx = r_idx
            best_headers = row_str

            # Tìm cột STT sơ bộ
            for c_i, s in enumerate(row_lower):
                if any(kw == s or kw in s for kw in STT_KEYWORDS):
                    stt_idx = c_i
                    break

            # Tìm cột Họ và tên sơ bộ (ưu tiên cột Họ đệm/Họ và tên trước, sau đó tới Tên)
            for c_i, s in enumerate(row_lower):
                s_clean = s.strip()
                if any(kw in s_clean for kw in ["họ và tên", "họ tên", "họ, chữ đệm", "họ và chữ đệm", "họ đệm", "họ, đệm", "họ và", "hodem", "ho dem", "họ lót"]) and "người ký" not in s_clean:
                    name_idx = c_i
                    break
            if name_idx == -1:
                for c_i, s in enumerate(row_lower):
                    s_clean = s.strip()
                    if s_clean in ["tên", "ten", "tên học sinh", "tên thí sinh", "tên người học"] or any(kw in s_clean for kw in ["tên học sinh", "tên thí sinh"]):
                        name_idx = c_i
                        break
            break

    # Nếu không tìm thấy theo từ khóa, mặc định lấy dòng 1 hoặc 2
    if best_header_idx is None:
        best_header_idx = 1
        first_row = next(ws.iter_rows(min_row=1, max_row=1, values_only=True), [])
        first_is_tcvn3 = any(is_tcvn3(str(c or '').strip()) for c in first_row)
        best_headers = []
        for c in first_row:
            s_raw = str(c or '').strip()
            if first_is_tcvn3 and any(ord(ch) > 127 for ch in s_raw):
                best_headers.append(clean_whitespace(tcvn3_to_unicode(s_raw)))
            else:
                best_headers.append(auto_convert_legacy_encoding(clean_whitespace(str(c or ''))))

    header_start_row = best_header_idx
    header_end_row = best_header_idx

    # Cắt gọn đuôi rỗng của header
    while best_headers and not best_headers[-1]:
        best_headers.pop()

    # Chuẩn hóa tên cột hiển thị
    formatted_headers = []
    for c_i, h in enumerate(best_headers):
        col_letter = get_column_letter(c_i + 1)
        name = h if h else f"Cột {col_letter}"
        formatted_headers.append(f"[{col_letter}] {name}")

    # Tìm lại cột STT và Họ và tên từ danh sách header hoàn chỉnh
    row_lower = [s.lower() for s in best_headers]
    for c_i, s in enumerate(row_lower):
        if any(kw == s or kw in s for kw in STT_KEYWORDS):
            stt_idx = c_i
            break

    name_idx = -1
    for c_i, s in enumerate(row_lower):
        if any(kw in s for kw in NAME_KEYWORDS):
            name_idx = c_i
            break

    # Fallback Heuristic tìm cột Họ và tên nếu chưa thấy ở header
    if name_idx == -1:
        name_idx = _heuristic_find_name_column(ws, header_end_row, stt_idx)
        if name_idx == -1 and len(best_headers) > 4:
            name_idx = 4  # Mặc định cột E nếu là chuẩn 28 cột
        elif name_idx == -1 and len(best_headers) > 1:
            name_idx = 1  # Mặc định cột B

    return {
        "header_row_idx": header_end_row,
        "header_start_row": header_start_row,
        "header_end_row": header_end_row,
        "raw_headers": best_headers,
        "formatted_headers": formatted_headers,
        "stt_col_idx": stt_idx,
        "name_col_idx": name_idx,
        "total_cols": len(best_headers)
    }

def _heuristic_find_name_column(ws, header_row_idx: int, stt_col_idx: int) -> int:
    """Dò tìm cột họ tên dựa trên mẫu tên người Việt ở các dòng dữ liệu đầu tiên."""
    col_scores: dict[int, int] = {}
    rows_checked = 0
    
    for row in ws.iter_rows(min_row=header_row_idx + 1, max_row=header_row_idx + 30, values_only=True):
        if not row:
            continue
        cell_stt = row[stt_col_idx] if stt_col_idx < len(row) else None
        if not is_valid_stt(cell_stt):
            continue
            
        rows_checked += 1
        for c_i, val in enumerate(row):
            if c_i == stt_col_idx or val is None:
                continue
            s = clean_whitespace(str(val))
            if not s or not re.search(r'^[a-zA-ZÀ-ỹ\s]{3,40}$', s):
                continue
            words = s.split()
            if 2 <= len(words) <= 5:
                first_word = words[0].lower()
                if first_word in VIETNAMESE_SURNAMES:
                    col_scores[c_i] = col_scores.get(c_i, 0) + 3
                else:
                    col_scores[c_i] = col_scores.get(c_i, 0) + 1

        if rows_checked >= 10:
            break

    if col_scores:
        return max(col_scores, key=col_scores.get)
    return -1


class ExcelSplitter:
    """Bộ xử lý Tách file Excel thành nhiều file con theo cột tùy chọn."""

    @staticmethod
    def preview_split(
        file_path: Path | str,
        sheet_name: str | None = None,
        split_col_idx: int | None = None,
        header_row_idx: int | None = None,
        stt_col_idx: int | None = None,
        name_col_idx: int | None = None
    ) -> dict:
        """
        Quét xem trước file nguồn:
        - Đếm số dòng hợp lệ (thỏa mãn STT + Họ tên).
        - Ngắt khi gặp 15 dòng liên tiếp có STT trống.
        - Thống kê các nhóm theo cột tách.
        """
        p = Path(file_path)
        if not p.exists():
            raise FileNotFoundError(f"Không tìm thấy file: {p}")

        wb = openpyxl.load_workbook(p, data_only=True)
        try:
            ws = wb[sheet_name] if (sheet_name and sheet_name in wb.sheetnames) else wb.active
            info = detect_header_and_columns(ws)

            h_idx = header_row_idx or info["header_row_idx"]
            s_idx = stt_col_idx if stt_col_idx is not None else info["stt_col_idx"]
            n_idx = name_col_idx if name_col_idx is not None else info["name_col_idx"]
            sp_idx = split_col_idx if split_col_idx is not None else s_idx

            groups: dict[str, int] = {}
            total_valid_rows = 0
            empty_stt_consecutive = 0

            best_header_row = [ws.cell(row=h_idx, column=c).value for c in range(1, ws.max_column + 1)] if ws.max_column else []

            for r_idx, row in enumerate(ws.iter_rows(min_row=h_idx + 1, values_only=True), start=h_idx + 1):
                cell_stt = row[s_idx] if s_idx < len(row) else None
                cell_name = row[n_idx] if n_idx < len(row) else None

                # Điều kiện dừng sớm: 15 dòng liên tiếp STT trống
                if is_cell_empty(cell_stt):
                    empty_stt_consecutive += 1
                    if empty_stt_consecutive >= 15:
                        break
                    continue
                else:
                    empty_stt_consecutive = 0

                # Bỏ qua dòng trùng lặp tiêu đề
                if is_header_duplicate_row(row, best_header_row, s_idx, n_idx):
                    continue

                # Điều kiện dòng hợp lệ: có STT hợp lệ VÀ có Họ tên hợp lệ
                if is_valid_stt(cell_stt) and is_valid_name(cell_name):
                    total_valid_rows += 1
                    raw_val = row[sp_idx] if sp_idx < len(row) else None
                    val_str = clean_whitespace(str(raw_val)) if not is_cell_empty(raw_val) else "(Trống)"
                    groups[val_str] = groups.get(val_str, 0) + 1

            return {
                "file_name": p.name,
                "sheet_name": ws.title,
                "all_sheets": wb.sheetnames,
                "header_info": info,
                "total_valid_rows": total_valid_rows,
                "groups": groups,
                "group_count": len(groups)
            }
        finally:
            wb.close()

    @staticmethod
    def split_file(
        file_path: Path | str,
        split_col_idx: int,
        output_dir: Path | str,
        sheet_name: str | None = None,
        header_row_idx: int | None = None,
        stt_col_idx: int | None = None,
        name_col_idx: int | None = None,
        renumber_stt: bool = True,
        progress_callback=None,
        log_callback=None,
        check_cancel=None
    ) -> dict:
        """
        Thực hiện tách file:
        - Giữ nguyên khối header (1..h_idx).
        - Đọc dòng hợp lệ và phân loại vào các nhóm.
        - Ghi từng nhóm ra file riêng, giữ nguyên style và định dạng.
        """
        def log(msg, level="info"):
            if log_callback:
                log_callback(msg, level)
            else:
                safe_print(msg)

        p = Path(file_path)
        out_dir = Path(output_dir)
        out_dir.mkdir(parents=True, exist_ok=True)

        wb_in = openpyxl.load_workbook(p, data_only=False)
        ws_in = wb_in[sheet_name] if (sheet_name and sheet_name in wb_in.sheetnames) else wb_in.active

        info = detect_header_and_columns(ws_in)
        h_idx = header_row_idx or info["header_row_idx"]
        s_idx = stt_col_idx if stt_col_idx is not None else info["stt_col_idx"]
        n_idx = name_col_idx if name_col_idx is not None else info["name_col_idx"]

        log(f"🔍 Bắt đầu phân tích tệp: {p.name} (Sheet: '{ws_in.title}')", "info")
        log(f"   Dòng tiêu đề: {h_idx} | Cột STT: {s_idx + 1} | Cột Họ tên: {n_idx + 1} | Cột tách: {split_col_idx + 1}", "info")

        # 1. Trích xuất toàn bộ các dòng dữ liệu hợp lệ
        best_header_row = [ws_in.cell(row=h_idx, column=c).value for c in range(1, ws_in.max_column + 1)] if ws_in.max_column else []
        grouped_rows: dict[str, list[list]] = {}
        total_extracted = 0
        empty_stt_consecutive = 0

        for row in ws_in.iter_rows(min_row=h_idx + 1, values_only=True):
            if check_cancel and check_cancel():
                log("⏹ Quá trình tách tệp đã dừng theo yêu cầu.", "warning")
                wb_in.close()
                return {"status": "cancelled", "files": []}

            cell_stt = row[s_idx] if s_idx < len(row) else None
            cell_name = row[n_idx] if n_idx < len(row) else None

            # Quy tắc dừng sớm: 15 dòng liên tiếp có STT trống
            if is_cell_empty(cell_stt):
                empty_stt_consecutive += 1
                if empty_stt_consecutive >= 15:
                    break
                continue
            else:
                empty_stt_consecutive = 0

            # Bỏ qua dòng trùng lặp tiêu đề
            if is_header_duplicate_row(row, best_header_row, s_idx, n_idx):
                continue

            # Quy tắc dòng hợp lệ: Có STT hợp lệ VÀ Có Họ tên hợp lệ
            if is_valid_stt(cell_stt) and is_valid_name(cell_name):
                total_extracted += 1
                raw_val = row[split_col_idx] if split_col_idx < len(row) else None
                val_str = clean_whitespace(str(raw_val)) if not is_cell_empty(raw_val) else "(Gia_Tri_Trong)"
                grouped_rows.setdefault(val_str, []).append(list(row))
                if progress_callback and (total_extracted % 2000 == 0):
                    progress_callback(total_extracted, 0, f"Đang phân tích dòng {total_extracted:,}...")

        log(f"✅ Đã thu nạp {total_extracted} dòng dữ liệu hợp lệ thuộc {len(grouped_rows)} nhóm.", "success")

        if total_extracted == 0:
            wb_in.close()
            return {"status": "empty", "files": [], "total_rows": 0, "total_groups": 0}

        # 2. Thu thập khối Header (Dòng 1 đến h_idx) kèm định dạng
        header_rows_data = []
        for r in range(1, h_idx + 1):
            row_data = [ws_in.cell(row=r, column=c).value for c in range(1, ws_in.max_column + 1)]
            header_rows_data.append(row_data)

        # Lưu lại độ rộng cột từ sheet gốc
        col_widths = {}
        for col_letter, col_dim in ws_in.column_dimensions.items():
            if col_dim.width is not None:
                col_widths[col_letter] = col_dim.width

        wb_in.close()

        # 3. Xuất từng nhóm ra file riêng
        created_files = []
        total_groups = len(grouped_rows)

        for g_idx, (grp_name, rows_list) in enumerate(grouped_rows.items(), start=1):
            if check_cancel and check_cancel():
                log("⏹ Đã dừng xuất các file còn lại theo yêu cầu.", "warning")
                break

            safe_name = sanitize_filename(grp_name) if grp_name != "(Gia_Tri_Trong)" else f"{p.stem}_Gia_Tri_Trong"
            if not safe_name:
                safe_name = f"Nhom_{g_idx}"

            out_file = out_dir / f"{safe_name}.xlsx"

            # Tạo file mới và sao chép khối header
            wb_out = openpyxl.Workbook()
            ws_out = wb_out.active
            ws_out.title = "Data"

            # Ghi khối header
            for h_row in header_rows_data:
                ws_out.append(h_row)

            # Ghi dữ liệu học sinh
            for idx, r_vals in enumerate(rows_list, start=1):
                row_copy = list(r_vals)
                if renumber_stt and s_idx < len(row_copy):
                    row_copy[s_idx] = idx
                ws_out.append(row_copy)

            # Áp dụng độ rộng cột
            for col_letter, w in col_widths.items():
                ws_out.column_dimensions[col_letter].width = w

            wb_out.save(out_file)
            wb_out.close()
            created_files.append(out_file)

            if progress_callback:
                progress_callback(g_idx, total_groups, f"Đang tách ({g_idx}/{total_groups}): {out_file.name}")

            log(f"  📄 Đã xuất: {out_file.name} ({len(rows_list)} dòng)", "info")

        log(f"🎉 Hoàn tất! Đã tách thành {len(created_files)} tệp Excel tại: {out_dir.name}", "success")
        return {
            "status": "success",
            "files": created_files,
            "total_rows": total_extracted,
            "total_groups": len(created_files)
        }


def find_best_data_sheet(wb, target_sheet_name: str | None = None):
    """
    Tìm sheet dữ liệu học sinh tốt nhất trong workbook:
    - Nếu có target_sheet_name trong wb -> lấy sheet đó.
    - Duyệt qua các sheet visible, tìm sheet có đầy đủ cột STT và Họ và tên.
    - Fallback: wb.active
    """
    if target_sheet_name and target_sheet_name in wb.sheetnames:
        return wb[target_sheet_name]

    for ws in wb.worksheets:
        if ws.sheet_state != 'visible':
            continue
        info = detect_header_and_columns(ws)
        if info.get("stt_col_idx", -1) >= 0 and info.get("name_col_idx", -1) >= 0:
            return ws

    return wb.active


class ExcelMerger:
    """Bộ xử lý Nhập / Gộp nhiều file Excel có cùng cấu trúc tiêu đề."""

    @staticmethod
    def validate_headers(file_paths: list[Path | str], sheet_name: str | None = None) -> dict:
        """
        Kiểm tra độ tương thích giữa các file cần gộp:
        - Tự động tìm sheet dữ liệu chuẩn trong từng file (tránh lấy nhầm sheet bìa/thống kê).
        - So sánh số lượng và nhãn cột của các file tiếp theo với Master Header.
        """
        if not file_paths:
            return {"compatible": False, "error": "Chưa chọn file nào để gộp."}

        reports = []
        master_headers = None
        master_file = Path(file_paths[0])

        wb_m = openpyxl.load_workbook(master_file, data_only=True)
        try:
            ws_m = find_best_data_sheet(wb_m, sheet_name)
            m_info = detect_header_and_columns(ws_m)
            master_headers = m_info["raw_headers"]
            reports.append({
                "file": master_file.name,
                "sheet": ws_m.title,
                "status": "master",
                "col_count": len(master_headers),
                "header_row": m_info["header_row_idx"]
            })
        finally:
            wb_m.close()

        all_match = True
        for f_p in file_paths[1:]:
            p = Path(f_p)
            wb = openpyxl.load_workbook(p, data_only=True)
            try:
                ws = find_best_data_sheet(wb, sheet_name)
                info = detect_header_and_columns(ws)
                headers = info["raw_headers"]

                # So sánh độ dài cột
                len_match = (len(headers) == len(master_headers))
                reports.append({
                    "file": p.name,
                    "sheet": ws.title,
                    "status": "match" if len_match else "mismatch",
                    "col_count": len(headers),
                    "header_row": info["header_row_idx"]
                })
                if not len_match:
                    all_match = False
            finally:
                wb.close()

        return {
            "compatible": all_match,
            "master_headers": master_headers,
            "master_col_count": len(master_headers),
            "reports": reports
        }

    @staticmethod
    def merge_files(
        file_paths: list[Path | str],
        output_file_path: Path | str,
        sheet_name: str | None = None,
        renumber_stt: bool = True,
        add_source_col: bool = False,
        merge_all_sheets: bool = True,
        progress_callback=None,
        log_callback=None,
        check_cancel=None
    ) -> dict:
        """
        Gộp nhiều file Excel thông minh:
        - Tự động nhận diện và quét tất cả các Sheet có dữ liệu học sinh (hoặc sheet cụ thể).
        - Bỏ qua các sheet bìa/thống kê không chứa cột STT & Họ tên.
        - File đầu tiên: giữ nguyên khối header.
        - Các file và sheet tiếp theo: bỏ header, chỉ lấy dòng hợp lệ (STT + Họ tên).
        - Ngắt khi gặp 15 dòng liên tiếp có STT trống trong mỗi sheet.
        - Đánh lại STT liên tục toàn bảng từ 1..N.
        """
        def log(msg, level="info"):
            if log_callback:
                log_callback(msg, level)
            else:
                safe_print(msg)

        if not file_paths:
            return {"status": "error", "message": "Không có tệp để gộp"}

        out_path = Path(output_file_path)
        out_path.parent.mkdir(parents=True, exist_ok=True)

        master_file = Path(file_paths[0])
        log(f"🚀 Bắt đầu gộp {len(file_paths)} tệp. Tệp mẫu chuẩn: {master_file.name}", "info")

        # 1. Đọc tệp mẫu (Master) để lấy cấu trúc Header và định dạng từ sheet chuẩn nhất
        wb_master = openpyxl.load_workbook(master_file, data_only=False)
        ws_master = find_best_data_sheet(wb_master, sheet_name)
        m_info = detect_header_and_columns(ws_master)
        h_idx = m_info["header_row_idx"]
        s_idx = m_info["stt_col_idx"]
        n_idx = m_info["name_col_idx"]

        log(f"📋 Sử dụng cấu trúc tiêu đề từ Sheet '{ws_master.title}' (Tệp: {master_file.name})", "info")

        header_rows_data = []
        for r in range(1, h_idx + 1):
            row_data = [ws_master.cell(row=r, column=c).value for c in range(1, ws_master.max_column + 1)]
            if add_source_col and r == h_idx:
                row_data.append("Tệp nguồn gốc")
            header_rows_data.append(row_data)

        col_widths = {c: ws_master.column_dimensions[c].width for c in ws_master.column_dimensions if ws_master.column_dimensions[c].width}
        wb_master.close()

        # 2. Tạo workbook gộp
        wb_merged = openpyxl.Workbook()
        ws_merged = wb_merged.active
        ws_merged.title = "Data_Tong_Hop"

        for h_row in header_rows_data:
            ws_merged.append(h_row)

        # 3. Quét từng tệp và trích xuất dữ liệu
        total_merged_rows = 0
        total_files = len(file_paths)

        for f_idx, f_p in enumerate(file_paths, start=1):
            if check_cancel and check_cancel():
                log("⏹ Đã dừng gộp tệp theo yêu cầu.", "warning")
                break

            p = Path(f_p)
            try:
                wb_curr = openpyxl.load_workbook(p, data_only=True)
            except Exception as e:
                log(f"❌ Không thể đọc tệp {p.name}: {e}", "error")
                continue

            # Xác định danh sách sheet cần gộp trong file này
            if merge_all_sheets:
                candidate_sheets = [ws for ws in wb_curr.worksheets if ws.sheet_state == 'visible']
            else:
                candidate_sheets = [find_best_data_sheet(wb_curr, sheet_name)]

            file_row_count = 0

            for ws_curr in candidate_sheets:
                curr_info = detect_header_and_columns(ws_curr)
                curr_h_idx = curr_info.get("header_row_idx", 1)
                curr_s_idx = curr_info.get("stt_col_idx", -1)
                curr_n_idx = curr_info.get("name_col_idx", -1)

                # Bỏ qua sheet không có STT hoặc Họ và tên (ví dụ: sheet bìa, sheet tổng hợp số lượng)
                if curr_s_idx < 0 or curr_n_idx < 0:
                    if len(candidate_sheets) > 1:
                        log(f"  ℹ️ Bỏ qua Sheet '{ws_curr.title}' trong {p.name} (Không chứa bảng danh sách học sinh).", "info")
                    continue

                sheet_row_count = 0
                empty_stt_consecutive = 0

                for row in ws_curr.iter_rows(min_row=curr_h_idx + 1, values_only=True):
                    # Đảm bảo index nằm trong giới hạn row
                    if curr_s_idx >= len(row) or curr_n_idx >= len(row):
                        empty_stt_consecutive += 1
                        if empty_stt_consecutive >= 15:
                            break
                        continue

                    cell_stt = row[curr_s_idx]
                    cell_name = row[curr_n_idx]

                    # Điều kiện dừng: 15 dòng liên tiếp có STT trống trong sheet này
                    if is_cell_empty(cell_stt):
                        empty_stt_consecutive += 1
                        if empty_stt_consecutive >= 15:
                            break
                        continue
                    else:
                        empty_stt_consecutive = 0

                    # Bỏ qua dòng trùng lặp tiêu đề
                    if is_header_duplicate_row(row, header_rows_data[-1] if header_rows_data else None, curr_s_idx, curr_n_idx):
                        continue

                    # Điều kiện dòng hợp lệ: Có STT hợp lệ VÀ Có Họ tên hợp lệ
                    if is_valid_stt(cell_stt) and is_valid_name(cell_name):
                        total_merged_rows += 1
                        sheet_row_count += 1
                        file_row_count += 1
                        row_data = list(row)

                        # Đánh lại STT nếu được yêu cầu
                        if renumber_stt and s_idx >= 0 and s_idx < len(row_data):
                            row_data[s_idx] = total_merged_rows

                        # Thêm cột nguồn (tên file + tên sheet nếu file có nhiều sheet)
                        if add_source_col:
                            src_desc = f"{p.name} ({ws_curr.title})" if len(candidate_sheets) > 1 else p.name
                            row_data.append(src_desc)

                        ws_merged.append(row_data)

                        if progress_callback and (total_merged_rows % 2000 == 0):
                            progress_callback(f_idx, total_files, f"Đang gộp tệp {f_idx}/{total_files} ({p.name}) - Đã gộp {total_merged_rows:,} dòng...")

                if sheet_row_count > 0 and len(candidate_sheets) > 1:
                    log(f"    -> Sheet '{ws_curr.title}': Đã lấy {sheet_row_count} học sinh.", "info")

            wb_curr.close()
            log(f"  📥 ({f_idx}/{total_files}) Đã gộp tổng cộng {file_row_count} dòng từ: {p.name}", "info")
            if progress_callback:
                progress_callback(f_idx, total_files, f"Đang gộp ({f_idx}/{total_files}): {p.name}")

        # Gán lại độ rộng cột từ tệp Master
        for col_letter, w in col_widths.items():
            ws_merged.column_dimensions[col_letter].width = w

        wb_merged.save(out_path)
        wb_merged.close()

        log(f"🎉 Gộp tệp thành công! Đã gom {total_merged_rows} dòng từ {total_files} file vào: {out_path.name}", "success")
        return {
            "status": "success",
            "output_file": out_path,
            "total_rows": total_merged_rows,
            "total_files": total_files
        }
