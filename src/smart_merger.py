from __future__ import annotations

import sys
import re
import json
import hashlib
import unicodedata
import difflib
import functools
import io
from pathlib import Path
from typing import Any, Callable
from datetime import datetime

import openpyxl
from openpyxl.styles import Font
from openpyxl.utils import column_index_from_string

from src.lookup_manager import find_data_file, get_persistent_data_dir
from src.normalizer import (
    clean_whitespace,
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
    {"idx": 25, "field": "trang_thai", "name": "Trạng thái số hóa", "required": False, "default": "Đã số hoá đầy đủ, chính xác", "type": "text"},
    {"idx": 26, "field": "ly_do_loi", "name": "Thông tin chưa bảo đảm đầy đủ, chính xác và nguyên nhân", "required": False, "default": "", "type": "text"},
    {"idx": 27, "field": "ma_don_vi", "name": "Mã đơn vị", "required": False, "default": "", "type": "text"},
    {"idx": 28, "field": "ten_don_vi", "name": "Tên đơn vị", "required": False, "default": "", "type": "text"}
]

# Các cột cần định dạng Text @ trong openpyxl
TEXT_COLS_28 = {2, 3, 6, 7, 8, 9, 19, 24, 27}

# ==============================================================================
# TỪ ĐIỂN ĐỒNG NGHĨA & TỪ KHÓA NGHIỆP VỤ CHO 28 CỘT (TẦNG 1)
# ==============================================================================

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
    res = "".join([c for c in nfkd_form if not unicodedata.combining(c)])
    return res.replace('đ', 'd').replace('Đ', 'D')

def clean_header_base(s: str) -> str:
    """Loại bỏ phần chú thích trong ngoặc đơn, ngoặc vuông để lấy tiêu đề cốt lõi."""
    s_clean = re.sub(r'[\(\[\{].*?[\)\]\}]', '', str(s or ''))
    return s_clean.strip()

@functools.lru_cache(maxsize=4096)
def normalize_text_for_match(s: str) -> str:
    """Chuẩn hóa chuỗi text cho việc so khớp mờ."""
    if s is None:
        return ""
    base = clean_header_base(s)
    s_conv = auto_convert_legacy_encoding(str(base))
    s_no_acc = remove_accents(s_conv).lower()
    s_clean = RE_NON_ALPHANUM.sub(' ', s_no_acc)
    return RE_SPACES.sub(' ', s_clean).strip()

def compute_fuzzy_similarity(name1: str, name2: str) -> float:
    """
    Tính độ tương đồng mờ trực tiếp giữa hai tiêu đề cột (không dùng từ điển).
    Hỗ trợ:
    - Bỏ dấu tiếng Việt, viết thường, loại bỏ ký tự đặc biệt.
    - So khớp chuỗi cốt lõi bỏ phần trong ngoặc.
    - So khớp dạng compact (viết liền không dấu cách, vd: dantoc <-> dan toc).
    - So khớp tập từ (tokens Jaccard / Subset) và SequenceMatcher.
    """
    if not name1 or not name2:
        return 0.0

    n1 = normalize_text_for_match(name1)
    n2 = normalize_text_for_match(name2)
    if not n1 or not n2:
        return 0.0

    if n1 == n2:
        return 1.0

    # 1. So khớp phần cốt lõi sau khi bỏ ngoặc đơn/vuông
    b1 = normalize_text_for_match(clean_header_base(name1))
    b2 = normalize_text_for_match(clean_header_base(name2))
    if b1 and b2 and b1 == b2:
        return 0.98

    # 2. So khớp dạng compact (viết liền không dấu cách)
    c1 = n1.replace(" ", "")
    c2 = n2.replace(" ", "")
    if c1 == c2:
        return 0.98

    cb1 = b1.replace(" ", "") if b1 else ""
    cb2 = b2.replace(" ", "") if b2 else ""
    if cb1 and cb2 and cb1 == cb2:
        return 0.96

    compact_score = 0.0
    if c1 in c2 or c2 in c1:
        ratio = min(len(c1), len(c2)) / max(len(c1), len(c2))
        compact_score = max(compact_score, 0.88 * ratio)
    if cb1 and cb2 and (cb1 in cb2 or cb2 in cb1):
        r_b = min(len(cb1), len(cb2)) / max(len(cb1), len(cb2))
        compact_score = max(compact_score, 0.88 * r_b)

    # 3. So khớp tập từ (Token Jaccard & Subset)
    toks1 = set(n1.split())
    toks2 = set(n2.split())
    inter = toks1.intersection(toks2)
    jaccard = len(inter) / len(toks1.union(toks2)) if (toks1 and toks2) else 0.0

    subset_score = 0.0
    if toks1 and toks2:
        if toks1.issubset(toks2):
            subset_score = 0.70 + 0.25 * (len(toks1) / len(toks2))
        elif toks2.issubset(toks1):
            subset_score = 0.70 + 0.25 * (len(toks2) / len(toks1))

    # 4. So khớp tuần tự SequenceMatcher
    seq = difflib.SequenceMatcher(None, n1, n2).ratio()
    seq_c = difflib.SequenceMatcher(None, c1, c2).ratio()
    seq_b = difflib.SequenceMatcher(None, b1, b2).ratio() if (b1 and b2) else 0.0

    score = max(compact_score, jaccard * 0.95, subset_score, seq, seq_c * 0.9, seq_b)
    if len(toks1) > 1 and len(toks2) > 1 and not inter:
        score *= 0.70

    return round(score, 3)

def compute_similarity(target_name: str, synonyms: list[str] | None, source_header: str) -> tuple[float, str]:
    """
    Tính điểm tương đồng giữa một cột nguồn và một cột đích (hỗ trợ tương thích ngược).
    """
    tgt_score = compute_fuzzy_similarity(target_name, source_header)
    if tgt_score >= 0.95:
        return tgt_score, "exact"

    best_syn_score = 0.0
    if synonyms:
        for syn in synonyms:
            s = compute_fuzzy_similarity(syn, source_header)
            if s > best_syn_score:
                best_syn_score = s

    if best_syn_score >= 0.90:
        return best_syn_score, "keyword"

    final_score = max(tgt_score, best_syn_score)
    if final_score >= 0.60:
        return final_score, "fuzzy"
    return final_score, "none"

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
            score = compute_fuzzy_similarity(col_def["name"], headers[idx])
            if score >= 0.85:
                matches += 1
    return matches >= 20

def auto_suggest_mapping_with_details(
    source_headers: list[str],
    profile_mgr: Any = None
) -> tuple[dict[int, int | list[int] | None], dict[int, str], str | None]:
    """
    Gợi ý ánh xạ thông minh:
    1. Kiểm tra tiêu đề: nếu là null/rỗng thì không khớp mờ, mặc định để None (-- Để trống --).
    2. Nếu nguồn là chuẩn 28 cột -> map 1-1.
    3. Từ lần thứ 2 trở đi (đã có Profile lưu trước đó):
       - So sánh xem có Profile khớp 100% tên cột không -> Nếu có: áp dụng luôn.
       - Nếu không khớp 100%: Khớp mờ tên cột nguồn với tên cột trong Profile có độ tương đồng cao nhất.
    4. Lần đầu tiên (hoặc các cột còn thiếu):
       - Khớp mờ trực tiếp với tên 28 cột chuẩn của file đích (TARGET_COLUMNS), không dùng từ điển.
    """
    mapping: dict[int, int | list[int] | None] = {col["idx"]: None for col in TARGET_COLUMNS}
    suggested_defaults: dict[int, str] = {}
    matched_profile_name: str | None = None

    if not source_headers:
        return mapping, suggested_defaults, matched_profile_name

    # Lọc các cột nguồn có nội dung hợp lệ (bỏ qua tuyệt đối các cột null / rỗng)
    valid_source_indices = {
        i for i, h in enumerate(source_headers)
        if h is not None and str(h).strip() != "" and normalize_text_for_match(h) != ""
    }

    if not valid_source_indices:
        return mapping, suggested_defaults, matched_profile_name

    # Nếu nguồn là chuẩn 28 cột
    if is_standard_28_headers(source_headers):
        for idx, col in enumerate(TARGET_COLUMNS):
            if idx < len(source_headers) and idx in valid_source_indices:
                mapping[col["idx"]] = idx
        return mapping, suggested_defaults, "Chuẩn 28 cột gốc"

    used_sources = set()
    used_targets = set()

    # =========================================================================
    # TẦNG 1: TỪ LẦN THỨ 2 TRỞ ĐI - SO SÁNH VỚI PROFILE ĐÃ LƯU
    # =========================================================================
    if profile_mgr and hasattr(profile_mgr, "list_profiles"):
        # 1.1. Khớp 100% Profile đã lưu -> Áp dụng luôn
        if hasattr(profile_mgr, "find_exact_matching_profile"):
            exact_prof, resolved_map = profile_mgr.find_exact_matching_profile(source_headers)
            if exact_prof and resolved_map:
                for k, v in exact_prof.get("default_values", {}).items():
                    if str(v).strip():
                        suggested_defaults[int(k)] = str(v).strip()
                return resolved_map, suggested_defaults, f"100% khớp: {exact_prof.get('name')}"

    # =========================================================================
    # Heuristic cho cột Họ đệm + Tên tách biệt (nếu Cột 5 chưa được map)
    # =========================================================================
    if 5 not in used_targets:
        has_full_name_col = False
        ho_col_idx = None
        ten_col_idx = None

        HO_KEYWORDS = ["ho", "ho dem", "ho va chu dem", "ho va ten dem", "ho lot", "hodem", "ho va", "ho dum", "hodum"]
        TEN_KEYWORDS = ["ten", "ten hs", "ten hoc sinh", "ten thi sinh"]

        for s_idx in valid_source_indices:
            if s_idx in used_sources:
                continue
            s_h = source_headers[s_idx]
            parts = [p.strip() for p in str(s_h).split(" - ")]
            full_norm = normalize_text_for_match(s_h)
            sub_norms = [normalize_text_for_match(p) for p in parts]

            if full_norm in ["ho va ten", "ho ten", "ho chu dem va ten", "ho va chu dem va ten"]:
                has_full_name_col = True
                break

            if ho_col_idx is None and (full_norm in HO_KEYWORDS or any(sn in HO_KEYWORDS for sn in sub_norms)):
                ho_col_idx = s_idx
                continue
            if ten_col_idx is None and (full_norm in TEN_KEYWORDS or any(sn in TEN_KEYWORDS for sn in sub_norms)):
                ten_col_idx = s_idx
                continue

        if not has_full_name_col and ho_col_idx is not None and ten_col_idx is not None and ho_col_idx != ten_col_idx:
            mapping[5] = [ho_col_idx, ten_col_idx]
            used_targets.add(5)
            used_sources.add(ho_col_idx)
            used_sources.add(ten_col_idx)

    # =========================================================================
    # Heuristic cho cột Ngày + Tháng + Năm sinh tách biệt (nếu Cột 9 chưa được map)
    # =========================================================================
    if 9 not in used_targets:
        has_full_dob_col = False
        ngay_col_idx = None
        thang_col_idx = None
        nam_col_idx = None

        NGAY_KEYWORDS = ["ngay", "ngay sinh"]
        THANG_KEYWORDS = ["thang", "thang sinh"]
        NAM_KEYWORDS = ["nam", "nam sinh"]

        for s_idx in valid_source_indices:
            if s_idx in used_sources:
                continue
            s_h = source_headers[s_idx]
            parts = [p.strip() for p in str(s_h).split(" - ")]
            full_norm = normalize_text_for_match(s_h)
            sub_norms = [normalize_text_for_match(p) for p in parts]

            if full_norm in ["ngay thang nam sinh", "ngay sinh", "ns"] and len(parts) == 1:
                has_full_dob_col = True
                break

            last_sub = sub_norms[-1] if sub_norms else ""
            if len(parts) > 1:
                if ngay_col_idx is None and last_sub in NGAY_KEYWORDS:
                    ngay_col_idx = s_idx
                    continue
                if thang_col_idx is None and last_sub in THANG_KEYWORDS:
                    thang_col_idx = s_idx
                    continue
                if nam_col_idx is None and last_sub in NAM_KEYWORDS:
                    nam_col_idx = s_idx
                    continue
            else:
                if ngay_col_idx is None and full_norm in NGAY_KEYWORDS:
                    ngay_col_idx = s_idx
                    continue
                if thang_col_idx is None and full_norm in THANG_KEYWORDS:
                    thang_col_idx = s_idx
                    continue
                if nam_col_idx is None and full_norm in NAM_KEYWORDS:
                    nam_col_idx = s_idx
                    continue

        if not has_full_dob_col and ngay_col_idx is not None and thang_col_idx is not None and nam_col_idx is not None:
            if len({ngay_col_idx, thang_col_idx, nam_col_idx}) == 3:
                mapping[9] = [ngay_col_idx, thang_col_idx, nam_col_idx]
                used_targets.add(9)
                used_sources.add(ngay_col_idx)
                used_sources.add(thang_col_idx)
                used_sources.add(nam_col_idx)

    # =========================================================================
    # TẦNG 1 (tiếp tục): Không khớp 100% -> Khớp mờ với Kho tri thức tập trung
    # =========================================================================
    if profile_mgr and hasattr(profile_mgr, "get_columns_knowledge"):
        knowledge = profile_mgr.get_columns_knowledge()
        p_candidates: list[tuple[float, int, int | list[int]]] = []

        for t_str, terms in knowledge.items():
            t_idx = int(t_str)
            if t_idx in used_targets:
                continue

            for term in terms:
                if isinstance(term, list):
                    matched_indices = []
                    sub_scores = []
                    for name in term:
                        best_idx = None
                        best_score = 0.0
                        for s_idx in valid_source_indices:
                            if s_idx in used_sources:
                                continue
                            sc = compute_fuzzy_similarity(name, source_headers[s_idx])
                            if sc > best_score:
                                best_score = sc
                                best_idx = s_idx
                        if best_idx is not None and best_score >= 0.60:
                            matched_indices.append(best_idx)
                            sub_scores.append(best_score)
                    if len(matched_indices) == len(term) and len(set(matched_indices)) == len(term):
                        avg_sub = sum(sub_scores) / len(sub_scores)
                        p_candidates.append((avg_sub, t_idx, matched_indices))
                else:
                    best_idx = None
                    best_score = 0.0
                    for s_idx in valid_source_indices:
                        if s_idx in used_sources:
                            continue
                        sc = compute_fuzzy_similarity(term, source_headers[s_idx])
                        if sc > best_score:
                            best_score = sc
                            best_idx = s_idx
                    if best_idx is not None and best_score >= 0.60:
                        p_candidates.append((best_score, t_idx, best_idx))

        p_candidates.sort(key=lambda x: x[0], reverse=True)
        for sc, t_idx, s_val in p_candidates:
            if t_idx in used_targets:
                continue
            if isinstance(s_val, list):
                if not any(i in used_sources for i in s_val):
                    mapping[t_idx] = s_val
                    used_targets.add(t_idx)
                    for i in s_val:
                        used_sources.add(i)
            else:
                if s_val not in used_sources:
                    mapping[t_idx] = s_val
                    used_targets.add(t_idx)
                    used_sources.add(s_val)

    # =========================================================================
    # TẦNG 2: LẦN ĐẦU TIÊN (HOẶC CỘT CÒN THIẾU): KHỚP MỜ VỚI TIÊU ĐỀ FILE 28 CỘT
    # =========================================================================
    candidates: list[tuple[float, int, int]] = []
    for col in TARGET_COLUMNS:
        t_idx = col["idx"]
        if t_idx in used_targets:
            continue
        t_name = col["name"]

        for s_idx in valid_source_indices:
            if s_idx in used_sources:
                continue
            s_header = source_headers[s_idx]

            # Cột 4 (Tên văn bằng) và Cột 28 (Tên đơn vị) không tự khớp với tiêu đề chỉ có chữ 'tên' đơn thuần
            if t_idx in (4, 28):
                s_norm = normalize_text_for_match(s_header)
                if t_idx == 4 and "bang" not in s_norm and "van bang" not in s_norm:
                    continue
                if t_idx == 28 and "don vi" not in s_norm and "phong" not in s_norm and "truong" not in s_norm:
                    continue

            score = compute_fuzzy_similarity(t_name, s_header)
            if score >= 0.58:
                candidates.append((score, t_idx, s_idx))

    candidates.sort(key=lambda x: x[0], reverse=True)

    for score, t_idx, s_idx in candidates:
        if t_idx not in used_targets and s_idx not in used_sources:
            mapping[t_idx] = s_idx
            used_targets.add(t_idx)
            used_sources.add(s_idx)

    return mapping, suggested_defaults, matched_profile_name


def auto_suggest_mapping(
    source_headers: list[str],
    profile_mgr: Any = None
) -> dict[int, int | list[int] | None]:
    """
    Gợi ý ánh xạ từ 28 cột đích sang chỉ số cột nguồn (0-based) hoặc None.
    Đảm bảo tương thích 100% với các mã nguồn cũ và các bài kiểm tra.
    """
    mapping, _, _ = auto_suggest_mapping_with_details(source_headers, profile_mgr)
    return mapping


def parse_column_expression(expr: str, formatted_headers: list[str]) -> int | list[int] | None:
    """
    Phân tích biểu thức ánh xạ cột từ người dùng:
    - Rỗng hoặc '-- Để trống --' -> None
    - Cột đơn chính xác (vd: 'Cột B: Họ và tên') -> int (chỉ số 0-based)
    - Biểu thức gộp nhiều cột qua dấu '+', ',', '&':
        + Theo chữ cái: 'B + C', 'Cột B + Cột C', 'B, C', 'B & C'
        + Theo nhãn đầy đủ: 'Cột B: Họ và chữ đệm + Cột C: Tên'
        + Theo số thứ tự 1-indexed: '2 + 3'
        + Khớp theo từ khóa tiêu đề: 'Họ và chữ đệm + Tên'
    - Trả về: int nếu là 1 cột, list[int] nếu là nhiều cột, None nếu không hợp lệ hoặc để trống.
    """
    if not expr:
        return None
    s = str(expr).strip()
    if s in ("", "-- Để trống --", "--", "None", "none"):
        return None

    # Khớp chính xác với 1 nhãn hoàn chỉnh trong formatted_headers
    if s in formatted_headers:
        return formatted_headers.index(s)

    # Tách biểu thức theo các toán tử ghép cột: '+', ',', '&'
    tokens = [t.strip() for t in re.split(r'[,+&]', s) if t.strip()]
    if not tokens:
        return None

    indices: list[int] = []
    for tok in tokens:
        # 1. Trực tiếp có trong formatted_headers
        if tok in formatted_headers:
            idx = formatted_headers.index(tok)
            if idx not in indices:
                indices.append(idx)
            continue

        # 2. Định dạng chữ cái cột Excel: "Cột B", "B", "Cột AA", "AA"
        m_col = re.match(r'^(?:cột\s+)?([A-Za-z]+)(?::.*)?$', tok, re.IGNORECASE)
        if m_col:
            letter = m_col.group(1).upper()
            try:
                c_idx = column_index_from_string(letter) - 1
                if 0 <= c_idx < len(formatted_headers) and c_idx not in indices:
                    indices.append(c_idx)
                    continue
            except Exception:
                pass

        # 3. Định dạng số 1-indexed: "1", "2"
        if tok.isdigit():
            c_idx = int(tok) - 1
            if 0 <= c_idx < len(formatted_headers) and c_idx not in indices:
                indices.append(c_idx)
                continue

        # 4. Tìm kiếm tương đối / từ khóa trong formatted_headers
        tok_norm = normalize_text_for_match(tok)
        matched_idx = None
        for i, fh in enumerate(formatted_headers):
            fh_norm = normalize_text_for_match(fh)
            if tok_norm and (tok_norm == fh_norm or tok_norm in fh_norm or fh_norm in tok_norm):
                matched_idx = i
                break
        if matched_idx is not None and matched_idx not in indices:
            indices.append(matched_idx)

    if not indices:
        return None
    if len(indices) == 1:
        return indices[0]
    return indices


# ==============================================================================
# QUẢN LÝ PROFILE ÁNH XẠ CỘT (MappingProfileManager)
# ==============================================================================

class MappingProfileManager:
    """Quản lý Kho tri thức ánh xạ tập trung theo 28 cột trong data/mapping_profiles.json."""

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
            self._save_raw({"version": "2.0", "columns": {}, "updated_at": ""})

    def _load_raw(self) -> dict:
        try:
            if not self.storage_path.exists():
                return {"version": "2.0", "columns": {}}
            with open(self.storage_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if not isinstance(data, dict):
                    return {"version": "2.0", "columns": {}}
                if "columns" not in data and "profiles" in data:
                    migrated = {"version": "2.0", "columns": {}}
                    for p in data["profiles"]:
                        named_m = p.get("named_mapping", {})
                        for t_str, val in named_m.items():
                            if val:
                                migrated["columns"].setdefault(t_str, []).append(val)
                    return migrated
                return data
        except Exception:
            return {"version": "2.0", "columns": {}}

    def _save_raw(self, data: dict):
        with open(self.storage_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    def get_columns_knowledge(self) -> dict[str, list[str | list[str]]]:
        """Lấy toàn bộ từ điển tri thức của 28 cột chuẩn."""
        data = self._load_raw()
        return data.get("columns", {})

    def get_learned_terms(self, t_idx: int) -> list[str | list[str]]:
        """Lấy danh sách các biến thể tiêu đề đã học cho cột đích t_idx."""
        cols = self.get_columns_knowledge()
        return cols.get(str(t_idx), [])

    def learn_mapping(
        self,
        mapping: dict[int, int | list[int] | None],
        raw_headers: list[str]
    ) -> None:
        """Học và bổ sung các biến thể tên cột mới vào kho tri thức tập trung."""
        if not mapping or not raw_headers:
            return
        data = self._load_raw()
        cols = data.setdefault("columns", {})

        for t_idx, s_val in mapping.items():
            if s_val is None:
                continue
            str_t = str(t_idx)
            existing_terms = cols.setdefault(str_t, [])

            if isinstance(s_val, list):
                combo = [
                    clean_whitespace(auto_convert_legacy_encoding(str(raw_headers[i])))
                    for i in s_val
                    if isinstance(i, int) and 0 <= i < len(raw_headers) and str(raw_headers[i] or "").strip()
                ]
                if combo:
                    combo_norms = [normalize_text_for_match(c) for c in combo]
                    already_exists = False
                    for ex in existing_terms:
                        if isinstance(ex, list) and len(ex) == len(combo):
                            if [normalize_text_for_match(c) for c in ex] == combo_norms:
                                already_exists = True
                                break
                    if not already_exists:
                        existing_terms.append(combo)

            elif isinstance(s_val, int) and 0 <= s_val < len(raw_headers):
                raw_text = clean_whitespace(auto_convert_legacy_encoding(str(raw_headers[s_val] or "")))
                if raw_text:
                    norm_text = normalize_text_for_match(raw_text)
                    already_exists = False
                    for ex in existing_terms:
                        if isinstance(ex, str) and normalize_text_for_match(ex) == norm_text:
                            already_exists = True
                            break
                    if not already_exists:
                        existing_terms.append(raw_text)

        data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._save_raw(data)

    def delete_learned_term(self, t_idx: int, term_idx: int) -> bool:
        """Xóa một từ khóa đã học của cột đích t_idx."""
        data = self._load_raw()
        cols = data.get("columns", {})
        str_t = str(t_idx)
        terms = cols.get(str_t, [])
        if 0 <= term_idx < len(terms):
            terms.pop(term_idx)
            data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._save_raw(data)
            return True
        return False

    def find_exact_matching_profile(
        self,
        current_raw_headers: list[str]
    ) -> tuple[dict | None, dict[int, int | list[int] | None] | None]:
        """
        Kiểm tra xem 100% các cột có dữ liệu trong current_raw_headers có khớp chính xác
        với kho tri thức đã học hay không.
        Nếu 100% khớp: trả về (profile_info, resolved_mapping).
        Nếu có cột chưa biết: trả về (None, None).
        """
        valid_indices = [i for i, h in enumerate(current_raw_headers) if str(h or "").strip()]
        if not valid_indices:
            return None, None

        knowledge = self.get_columns_knowledge()
        if not knowledge:
            return None, None

        resolved_map: dict[int, int | list[int] | None] = {col["idx"]: None for col in TARGET_COLUMNS}
        used_sources: set[int] = set()
        used_targets: set[int] = set()

        # 1. So khớp các cụm cột gộp trước (e.g. ['Họ đệm', 'Tên'], ['Ngày', 'Tháng', 'Năm'])
        for t_str, terms in knowledge.items():
            t_idx = int(t_str)
            if t_idx in used_targets:
                continue
            for term in terms:
                if isinstance(term, list):
                    matched = []
                    for sub in term:
                        norm_sub = normalize_text_for_match(sub)
                        found = None
                        for si in valid_indices:
                            if si not in used_sources and normalize_text_for_match(current_raw_headers[si]) == norm_sub:
                                found = si
                                break
                        if found is not None:
                            matched.append(found)
                    if len(matched) == len(term) and len(set(matched)) == len(term):
                        resolved_map[t_idx] = matched
                        used_targets.add(t_idx)
                        used_sources.update(matched)
                        break

        # 2. So khớp các cột đơn
        for si in valid_indices:
            if si in used_sources:
                continue
            norm_s = normalize_text_for_match(current_raw_headers[si])
            for t_str, terms in knowledge.items():
                t_idx = int(t_str)
                if t_idx in used_targets:
                    continue
                matched_term = False
                for term in terms:
                    if isinstance(term, str) and normalize_text_for_match(term) == norm_s:
                        matched_term = True
                        break
                if matched_term:
                    resolved_map[t_idx] = si
                    used_targets.add(t_idx)
                    used_sources.add(si)
                    break

        # Điều kiện tự động 100%: TẤT CẢ các cột hợp lệ trong file đều phải được nhận diện chính xác
        if len(used_sources) == len(valid_indices):
            prof_info = {
                "id": "knowledge_pool",
                "name": "Tri thức đã học",
                "total_matched": len(used_sources)
            }
            return prof_info, resolved_map

        return None, None

    def save_profile(
        self,
        name: str,
        raw_headers: list[str],
        mapping: dict[int, int | list[int] | None] | None = None,
        default_values: dict[int | str, str] | None = None,
        named_mapping: dict[str, str | list[str] | None] | None = None,
        signature: str | None = None
    ) -> dict:
        """Tương thích ngược: ghi nhận ánh xạ vào kho tri thức tập trung."""
        if mapping:
            self.learn_mapping(mapping, raw_headers)
        elif named_mapping:
            # Chuyển đổi named_mapping sang mapping nếu có
            norm_map = {}
            for t_str, val in named_mapping.items():
                if val is not None:
                    t_idx = int(t_str)
                    if isinstance(val, list):
                        indices = []
                        for sub in val:
                            n_sub = normalize_text_for_match(sub)
                            for i, h in enumerate(raw_headers):
                                if normalize_text_for_match(h) == n_sub and i not in indices:
                                    indices.append(i)
                                    break
                        norm_map[t_idx] = indices
                    else:
                        n_val = normalize_text_for_match(val)
                        for i, h in enumerate(raw_headers):
                            if normalize_text_for_match(h) == n_val:
                                norm_map[t_idx] = i
                                break
            self.learn_mapping(norm_map, raw_headers)
        return {"id": "knowledge_pool", "name": name, "raw_headers": raw_headers}

    def list_profiles(self) -> list[dict]:
        """Tương thích ngược cho UI: trả về danh sách các cột đã học tri thức."""
        cols = self.get_columns_knowledge()
        res = []
        for col in TARGET_COLUMNS:
            t_idx = col["idx"]
            terms = cols.get(str(t_idx), [])
            if terms:
                res.append({
                    "id": f"col_{t_idx}",
                    "name": f"Cột {t_idx}: {col['name']}",
                    "raw_headers": terms,
                    "named_mapping": {str(t_idx): terms[-1]},
                    "updated_at": ""
                })
        return res

    def get_named_mapping(self, profile: dict) -> dict[str, str | list[str] | None]:
        return profile.get("named_mapping", {})

    def delete_profile(self, profile_id: str) -> bool:
        if profile_id.startswith("col_"):
            try:
                t_idx = int(profile_id.replace("col_", ""))
                data = self._load_raw()
                if str(t_idx) in data.get("columns", {}):
                    del data["columns"][str(t_idx)]
                    self._save_raw(data)
                    return True
            except Exception:
                pass
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
                current_defaults: dict[int, str] = {}
                profile_used_name: str = ""

                # 1. Kiểm tra xem có khớp chuẩn 28 cột không
                if is_standard_28_headers(raw_headers):
                    mapping = {col["idx"]: idx for idx, col in enumerate(TARGET_COLUMNS) if idx < len(raw_headers)}
                    current_defaults = {}
                    profile_used_name = "Chuẩn 28 cột gốc"
                    log(f"  ⚡ Sheet '{ws_curr.title}': Cấu trúc chuẩn 28 cột $\\rightarrow$ Ánh xạ tự động 1:1.", "info")
                else:
                    # 2. Tìm Profile khớp 100% theo TÊN CỘT (không phụ thuộc thứ tự index)
                    exact_prof, resolved_map = self.profile_mgr.find_exact_matching_profile(raw_headers)
                    if exact_prof and resolved_map:
                        mapping = resolved_map
                        str_defs = exact_prof.get("default_values", {})
                        current_defaults = {int(k): str(v).strip() for k, v in str_defs.items() if str(v).strip()}
                        profile_used_name = exact_prof.get("name", "Profile đã lưu")
                        log(f"  ⚡ Sheet '{ws_curr.title}': Tự động áp dụng Profile '{profile_used_name}' (Khớp 100% tên cột).", "info")
                    else:
                        # 3. Chưa có Profile khớp 100% tên cột -> Gợi ý thông minh 2 tầng:
                        #    Tầng 1: Khớp mờ theo Profile gần nhất
                        #    Tầng 2: Điền tiếp theo từ điển đồng nghĩa 28 cột
                        suggested_map, suggested_defs, matched_pname = auto_suggest_mapping_with_details(
                            raw_headers, profile_mgr=self.profile_mgr
                        )
                        formatted_headers = header_info.get("formatted_headers", [])

                        callback_payload = {
                            "file_name": p.name,
                            "file_path": str(p),
                            "sheet_name": ws_curr.title,
                            "header_row_idx": h_idx,
                            "raw_headers": raw_headers,
                            "formatted_headers": formatted_headers,
                            "signature": sig,
                            "suggested_mapping": suggested_map,
                            "suggested_defaults": suggested_defs,
                            "matched_profile_name": matched_pname
                        }

                        if mapping_callback:
                            cb_res = mapping_callback(callback_payload)
                            if len(cb_res) == 4:
                                confirmed_map, prof_name, custom_defs, should_skip = cb_res
                            elif len(cb_res) == 3:
                                confirmed_map, prof_name, should_skip = cb_res
                                custom_defs = {}
                            else:
                                confirmed_map, prof_name, custom_defs, should_skip = None, None, {}, True

                            if should_skip or confirmed_map is None:
                                log(f"  ⏭️ Đã bỏ qua Sheet '{ws_curr.title}' theo lựa chọn của bạn.", "warning")
                                sheets_skipped += 1
                                continue

                            mapping = confirmed_map
                            current_defaults = {int(k): str(v).strip() for k, v in (custom_defs or {}).items() if str(v).strip()}
                            profile_used_name = prof_name or f"Profile_{p.stem}_{ws_curr.title}"

                            # Lưu Profile này vào bộ nhớ cho những lần sau (bao gồm cả Tên cột và Index)
                            self.profile_mgr.save_profile(
                                name=profile_used_name,
                                raw_headers=raw_headers,
                                mapping=mapping,
                                default_values=current_defaults
                            )
                            log(f"  💾 Đã lưu cấu hình ánh xạ: '{profile_used_name}' vào kho Profiles.", "success")
                        else:
                            # Không có UI callback -> sử dụng luôn gợi ý tự động
                            mapping = suggested_map
                            current_defaults = suggested_defs
                            profile_used_name = matched_pname or "Tự động suy luận"

                if not mapping:
                    continue

                # 3. Trích xuất và chuẩn hóa dòng dữ liệu học sinh từ Sheet
                sheet_rows_extracted = 0
                empty_name_consecutive = 0

                s_src_idx = mapping.get(1) # Cột STT nguồn
                n_src_idx = mapping.get(5) # Cột Họ tên nguồn

                first_s_idx = s_src_idx[0] if isinstance(s_src_idx, list) else (s_src_idx if s_src_idx is not None else -1)
                first_n_idx = n_src_idx[0] if isinstance(n_src_idx, list) else (n_src_idx if n_src_idx is not None else -1)

                # Đọc dòng tiêu đề chuẩn của sheet để kiểm tra lặp lại
                header_row_vals = [ws_curr.cell(row=h_idx, column=c).value for c in range(1, len(raw_headers) + 1)]

                for row in ws_curr.iter_rows(min_row=h_idx + 1, values_only=True):
                    if check_cancel and check_cancel():
                        break

                    # Kiểm tra dòng tiêu đề lặp lại (ví dụ ngắt trang lặp lại dòng tiêu đề)
                    if is_header_duplicate_row(row, header_row_vals, first_s_idx, first_n_idx):
                        continue

                    # Kiểm tra dòng tổng kết / chân trang ký tên
                    if is_summary_or_footer_row(row):
                        continue

                    # Kiểm tra tính hợp lệ của dòng học sinh: CHỈ CẦN CÓ HỌ VÀ TÊN HỢP LỆ
                    if n_src_idx is not None:
                        if isinstance(n_src_idx, list):
                            parts = []
                            for ni in n_src_idx:
                                if ni is not None and 0 <= ni < len(row) and row[ni] is not None:
                                    p_text = clean_whitespace(auto_convert_legacy_encoding(str(row[ni])))
                                    if p_text:
                                        parts.append(p_text)
                            cell_name = " ".join(parts) if parts else None
                        else:
                            cell_name = row[n_src_idx] if (0 <= n_src_idx < len(row)) else None
                        name_ok = is_valid_name(cell_name)

                        # Điều kiện ngắt: Nếu 20 dòng liên tiếp không có Họ tên hợp lệ -> kết thúc bảng dữ liệu
                        if not name_ok:
                            empty_name_consecutive += 1
                            if empty_name_consecutive >= 20:
                                break
                            continue
                        else:
                            empty_name_consecutive = 0
                    else:
                        # Dự phòng trường hợp hy hữu không map cột Họ tên: kiểm tra STT hoặc dòng không rỗng
                        cell_stt = row[first_s_idx] if (first_s_idx >= 0 and first_s_idx < len(row)) else None
                        if first_s_idx >= 0:
                            if not is_valid_stt(cell_stt):
                                empty_name_consecutive += 1
                                if empty_name_consecutive >= 20:
                                    break
                                continue
                            else:
                                empty_name_consecutive = 0
                        else:
                            if not any(row):
                                empty_name_consecutive += 1
                                if empty_name_consecutive >= 20:
                                    break
                                continue
                            else:
                                empty_name_consecutive = 0

                    # Dòng hợp lệ -> Tiến hành ánh xạ sang 28 cột chuẩn
                    sheet_rows_extracted += 1
                    total_extracted_rows += 1

                    row_28: list[Any] = [None] * 28

                    for target_col in TARGET_COLUMNS:
                        t_idx = target_col["idx"]
                        t_col_0 = t_idx - 1
                        s_col_idx = mapping.get(t_idx)

                        raw_val = None
                        if isinstance(s_col_idx, list):
                            parts = []
                            for sc in s_col_idx:
                                if sc is not None and 0 <= sc < len(row) and row[sc] is not None:
                                    p_str = auto_convert_legacy_encoding(str(row[sc]))
                                    p_str = clean_whitespace(p_str)
                                    if p_str:
                                        parts.append(p_str)
                            raw_val = " ".join(parts) if parts else None
                        elif isinstance(s_col_idx, int) and 0 <= s_col_idx < len(row):
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
                        if row_28[t_col_0] is None or row_28[t_col_0] == "":
                            # 1. Ưu tiên giá trị mặc định do người dùng tự nhập trong Profile cho cột này
                            cust_def = current_defaults.get(t_idx)
                            if cust_def is not None and str(cust_def).strip() != "":
                                c_str = str(cust_def).strip()
                                if t_idx in (9, 19):
                                    row_28[t_col_0] = format_date_to_text(c_str)
                                elif t_idx == 5:
                                    row_28[t_col_0] = vietnamese_title_case(c_str)
                                else:
                                    row_28[t_col_0] = clean_whitespace(c_str)
                            elif fill_defaults:
                                # 2. Giá trị mặc định chuẩn của hệ thống (nếu bật fill_defaults)
                                if t_idx == 4: # Tên văn bằng
                                    row_28[t_col_0] = "Bằng tốt nghiệp trung học cơ sở"
                                elif t_idx == 12: # Quốc tịch
                                    row_28[t_col_0] = "Việt Nam"

                    # Đánh lại STT tăng dần 1..N (hoặc tự điền nếu sheet nguồn không có STT)
                    if renumber_stt or row_28[0] is None or row_28[0] == "":
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
