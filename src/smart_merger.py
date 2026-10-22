from __future__ import annotations

import sys
import re
import json
import hashlib
import unicodedata
import difflib
import functools
import io
import uuid
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
                mapping[5] = s_idx
                used_targets.add(5)
                used_sources.add(s_idx)
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
                mapping[9] = s_idx
                used_targets.add(9)
                used_sources.add(s_idx)
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

    # =========================================================================
    # TẦNG 3: ÁP DỤNG QUY TẮC DỰ PHÒNG (FALLBACK RULES)
    # Ví dụ: Cột 17 (Hội đồng thi) dự phòng lấy từ Cột 14 (Tên trường đã học)
    # khi trong file không có cột Hội đồng thi riêng biệt.
    # =========================================================================
    fallback_rules = {}
    if profile_mgr and hasattr(profile_mgr, "get_fallback_columns"):
        fallback_rules = profile_mgr.get_fallback_columns()
    else:
        fallback_rules = {17: 14}

    for t_idx, fb_target in fallback_rules.items():
        if t_idx not in used_targets and fb_target in used_targets:
            fb_val = mapping.get(fb_target)
            if fb_val is not None:
                mapping[t_idx] = fb_val
                used_targets.add(t_idx)

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
    """
    Quản lý các Profile ánh xạ cột theo triết lý đảo ngược (Inverted Profile):
    - Profile lưu theo danh sách tiêu đề file Excel nguồn (source_headers) -> ánh xạ sang 28 cột mẫu.
    - Hỗ trợ so khớp tập hợp: Trùng khớp 100% (Exact), Tập con (Subset), Tập mẹ (Superset).
    - Hỗ trợ xóa đơn lẻ, xóa hàng loạt (nhiều profile cùng lúc) và xóa tất cả profile.
    """

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
            self._save_raw({
                "version": "3.0",
                "profiles": [],
                "columns": {},
                "fallback_columns": {"17": 14},
                "updated_at": ""
            })

    def _load_raw(self) -> dict:
        try:
            if not self.storage_path.exists():
                return {"version": "3.0", "profiles": [], "columns": {}, "fallback_columns": {"17": 14}}
            with open(self.storage_path, "r", encoding="utf-8") as f:
                data = json.load(f)
                if not isinstance(data, dict):
                    return {"version": "3.0", "profiles": [], "columns": {}, "fallback_columns": {"17": 14}}
                if "profiles" not in data:
                    data["profiles"] = []
                if "version" not in data:
                    data["version"] = "3.0"
                return data
        except Exception:
            return {"version": "3.0", "profiles": [], "columns": {}, "fallback_columns": {"17": 14}}

    def _save_raw(self, data: dict):
        with open(self.storage_path, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)

    DEFAULT_FALLBACK_COLUMNS: dict[str, int] = {
        "17": 14  # Cột 17 (Hội đồng thi) dự phòng lấy từ Cột 14 (Tên trường đã học) khi không có cột riêng
    }

    def get_fallback_columns(self) -> dict[int, int]:
        """Lấy các quy tắc dự phòng cột đích (ví dụ: Cột 17 lấy theo Cột 14 khi không có cột riêng)."""
        data = self._load_raw()
        fbs = data.get("fallback_columns")
        if fbs is None:
            fbs = self.DEFAULT_FALLBACK_COLUMNS
        return {int(k): int(v) for k, v in fbs.items()}

    def get_fallback_for_column(self, t_idx: int) -> int | None:
        """Lấy chỉ số cột dự phòng cho cột t_idx nếu có cấu hình."""
        return self.get_fallback_columns().get(t_idx)

    def set_fallback_columns(self, fallbacks: dict[int | str, int]):
        """Lưu cấu hình quy tắc dự phòng cột đích."""
        data = self._load_raw()
        data["fallback_columns"] = {str(k): int(v) for k, v in fallbacks.items()}
        data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._save_raw(data)

    def add_fallback_column(self, target_idx: int, fallback_idx: int) -> None:
        """Thêm hoặc cập nhật một quy tắc dự phòng cột đích."""
        fbs = self.get_fallback_columns()
        fbs[int(target_idx)] = int(fallback_idx)
        self.set_fallback_columns(fbs)

    def delete_fallback_column(self, target_idx: int) -> bool:
        """Xóa quy tắc dự phòng của một cột đích."""
        fbs = self.get_fallback_columns()
        if int(target_idx) in fbs:
            del fbs[int(target_idx)]
            self.set_fallback_columns(fbs)
            return True
        return False

    # -------------------------------------------------------------------------
    # QUẢN LÝ DANH SÁCH PROFILES ĐẢO NGƯỢC (Inverted Profiles)
    # -------------------------------------------------------------------------

    def get_all_profiles(self) -> list[dict]:
        """Lấy danh sách tất cả các Profile đã lưu."""
        data = self._load_raw()
        return data.get("profiles", [])

    def list_profiles(self) -> list[dict]:
        """Bí danh cho get_all_profiles(), tương thích UI."""
        return self.get_all_profiles()

    def get_profile_by_id(self, profile_id: str) -> dict | None:
        """Tìm Profile theo ID."""
        for p in self.get_all_profiles():
            if p.get("id") == profile_id:
                return p
        return None

    def save_profile(
        self,
        name: str,
        source_headers: list[str],
        mapping: dict[str, int | list[int] | None] | dict[int, int | list[int] | None],
        profile_id: str | None = None,
        default_values: dict[int | str, str] | None = None,
        sync_knowledge: bool = True
    ) -> dict:
        """
        Lưu hoặc cập nhật một Profile đảo ngược.
        Mapping được lưu dưới dạng: {"tên cột nguồn chuẩn hóa": chỉ_số_cột_mẫu | [cột, thứ_tự] | None}
        """
        data = self._load_raw()
        profiles = data.setdefault("profiles", [])

        # Chuẩn hóa danh sách tiêu đề nguồn
        clean_headers: list[str] = []
        norm_headers: list[str] = []
        for h in source_headers:
            if h is not None and str(h).strip():
                clean_str = clean_whitespace(auto_convert_legacy_encoding(str(h)))
                norm_str = normalize_text_for_match(clean_str)
                if clean_str and norm_str:
                    clean_headers.append(clean_str)
                    norm_headers.append(norm_str)

        # Chuẩn hóa mapping
        norm_mapping: dict[str, Any] = {}
        is_target_keyed = False
        if mapping:
            # Kiểm tra xem mapping truyền vào là {norm_header: target} hay {target_col_idx: source_idx}
            is_target_keyed = any(isinstance(k, int) and 1 <= k <= 28 for k in mapping.keys())
            if is_target_keyed:
                # Chuyển đổi từ {target_idx: source_idx_or_list} sang {norm_header: target}
                for t_idx, s_val in mapping.items():
                    if s_val is None:
                        continue
                    if isinstance(s_val, list):
                        for ord_idx, si in enumerate(s_val):
                            if 0 <= si < len(source_headers):
                                h_norm = normalize_text_for_match(source_headers[si])
                                if h_norm:
                                    norm_mapping[h_norm] = [int(t_idx), ord_idx]
                    elif isinstance(s_val, int) and 0 <= s_val < len(source_headers):
                        h_norm = normalize_text_for_match(source_headers[s_val])
                        if h_norm:
                            norm_mapping[h_norm] = int(t_idx)
                # Các cột nguồn không được ánh xạ -> map to None
                for nh in norm_headers:
                    if nh not in norm_mapping:
                        norm_mapping[nh] = None
            else:
                for k, v in mapping.items():
                    nh = normalize_text_for_match(str(k))
                    if nh:
                        norm_mapping[nh] = v

        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        now_ts = int(datetime.now().timestamp())

        # Tìm profile trùng lặp (theo ID hoặc tập hợp header giống hệt 100%)
        existing_idx = None
        s_set = set(norm_headers)
        for idx, p in enumerate(profiles):
            if profile_id and p.get("id") == profile_id:
                existing_idx = idx
                break
            if set(p.get("source_headers_norm", [])) == s_set:
                existing_idx = idx
                break

        prof_obj = {
            "id": profile_id or (profiles[existing_idx].get("id") if existing_idx is not None else f"prof_{now_ts}_{uuid.uuid4().hex[:6]}"),
            "name": name.strip() or f"Mẫu {len(clean_headers)} cột ({now_str})",
            "source_headers": clean_headers,
            "source_headers_norm": norm_headers,
            "mapping": norm_mapping,
            "timestamp": now_ts,
            "updated_at": now_str
        }

        if existing_idx is not None:
            profiles[existing_idx] = prof_obj
        else:
            profiles.append(prof_obj)

        data["updated_at"] = now_str
        self._save_raw(data)

        # Cập nhật song song từ điển biến thể tương thích
        if sync_knowledge and mapping:
            self.learn_mapping(
                mapping if is_target_keyed else self.resolve_execution_mapping(source_headers, norm_mapping),
                source_headers,
                save_prof=False
            )

        return prof_obj

    def delete_profile(self, profile_id: str) -> bool:
        """Xóa một Profile theo ID."""
        data = self._load_raw()
        profiles = data.get("profiles", [])
        init_len = len(profiles)
        data["profiles"] = [p for p in profiles if p.get("id") != profile_id]
        if len(data["profiles"]) < init_len:
            data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._save_raw(data)
            return True
        return False

    def delete_profiles(self, profile_ids: list[str]) -> int:
        """Xóa nhiều Profile cùng lúc theo danh sách ID."""
        if not profile_ids:
            return 0
        data = self._load_raw()
        profiles = data.get("profiles", [])
        id_set = set(profile_ids)
        new_profiles = [p for p in profiles if p.get("id") not in id_set]
        deleted_count = len(profiles) - len(new_profiles)
        if deleted_count > 0:
            data["profiles"] = new_profiles
            data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            self._save_raw(data)
        return deleted_count

    def delete_all_profiles(self) -> int:
        """Xóa toàn bộ Profile đã lưu."""
        data = self._load_raw()
        profiles = data.get("profiles", [])
        total = len(profiles)
        data["profiles"] = []
        data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._save_raw(data)
        return total

    # -------------------------------------------------------------------------
    # CHUYỂN ĐỔI MAPPING ĐẢO NGƯỢC -> MAPPING THỰC THI (Execution Mapping)
    # -------------------------------------------------------------------------

    def resolve_execution_mapping(
        self,
        raw_headers: list[str],
        inverted_mapping: dict[str, Any]
    ) -> dict[int, int | list[int] | None]:
        """
        Chuyển đổi từ Inverted Mapping (tên cột nguồn -> chỉ số cột mẫu)
        sang Execution Mapping (chỉ số cột mẫu 1..28 -> chỉ số cột nguồn 0..N-1).
        Đồng thời áp dụng các quy tắc cột dự phòng (fallback).
        """
        resolved: dict[int, int | list[int] | None] = {c["idx"]: None for c in TARGET_COLUMNS}
        composites: dict[int, list[tuple[int, int]]] = {}

        for si, h in enumerate(raw_headers):
            if h is None or not str(h).strip():
                continue
            norm_h = normalize_text_for_match(str(h))
            if not norm_h:
                continue

            target_spec = inverted_mapping.get(norm_h)
            if target_spec is None:
                continue

            if isinstance(target_spec, list) and len(target_spec) >= 2:
                # Dạng composite: [target_idx, order]
                t_idx, order = int(target_spec[0]), int(target_spec[1])
                composites.setdefault(t_idx, []).append((order, si))
            elif isinstance(target_spec, int) and 1 <= target_spec <= 28:
                resolved[target_spec] = si

        # Sắp xếp và đưa các cột composite vào resolved
        for t_idx, parts in composites.items():
            sorted_parts = sorted(parts, key=lambda x: x[0])
            indices = [si for _, si in sorted_parts]
            resolved[t_idx] = indices[0] if len(indices) == 1 else indices

        # Áp dụng quy tắc dự phòng cột đích (ví dụ Cột 17 lấy theo Cột 14)
        fallback_rules = self.get_fallback_columns()
        for t_idx, fb_target in fallback_rules.items():
            if resolved.get(t_idx) is None and resolved.get(fb_target) is not None:
                resolved[t_idx] = resolved[fb_target]

        return resolved

    # -------------------------------------------------------------------------
    # BỘ SO KHỚP TẬP HỢP 4 CẤP ĐỘ (Exact, Subset, Superset, Fallback)
    # -------------------------------------------------------------------------

    def find_matching_profile(
        self,
        raw_headers: list[str],
        sample_rows: list[list[Any]] | None = None
    ) -> tuple[str | None, dict | None, dict[int, int | list[int] | None] | None, list[str]]:
        """
        So khớp tập hợp tiêu đề của file Excel nguồn với kho Profiles:
        - Cấp 1: Khớp 100% (S == P)
        - Cấp 2: Khớp Tập con (S c P): chứa cột Họ tên, chọn Profile cha gần nhất.
        - Cấp 3: Khớp Tập mẹ (S > P):
            + Nếu các cột thừa không có dữ liệu -> Tự động bỏ qua và áp dụng.
            + Nếu các cột thừa có dữ liệu -> Báo về để mở dialog hỏi người dùng.
        - Cấp 4: Không khớp -> Trả về (None, None, None, [])

        Trả về: (match_type, matched_profile, resolved_mapping, extra_headers_with_data)
        """
        norm_source: list[str] = []
        for h in raw_headers:
            if h is not None and str(h).strip():
                nh = normalize_text_for_match(str(h))
                if nh:
                    norm_source.append(nh)

        if not norm_source:
            return None, None, None, []

        S = set(norm_source)
        profiles = self.get_all_profiles()
        if not profiles:
            return None, None, None, []

        # CẤP 1: TRÙNG KHỚP 100% (Exact Match: S == P)
        for p in profiles:
            P = set(p.get("source_headers_norm", []))
            if S == P:
                res_map = self.resolve_execution_mapping(raw_headers, p.get("mapping", {}))
                return "exact", p, res_map, []

        # CẤP 2: TẬP CON (Subset Match: S c P)
        subset_candidates: list[dict] = []
        for p in profiles:
            P = set(p.get("source_headers_norm", []))
            if S.issubset(P) and len(S) < len(P):
                # Kiểm tra điều kiện tiên quyết: Các cột còn lại trong S phải chứa Cột 5 (Họ tên)
                p_map = p.get("mapping", {})
                has_name = False
                for sh in S:
                    spec = p_map.get(sh)
                    if spec == 5 or (isinstance(spec, list) and len(spec) > 0 and spec[0] == 5):
                        has_name = True
                        break
                if has_name:
                    subset_candidates.append(p)

        if subset_candidates:
            # Chọn Profile cha có số lượng cột nhỏ nhất (gần S nhất), nếu bằng nhau chọn bản mới nhất
            subset_candidates.sort(key=lambda x: (len(x.get("source_headers_norm", [])), -x.get("timestamp", 0)))
            best_p = subset_candidates[0]
            res_map = self.resolve_execution_mapping(raw_headers, best_p.get("mapping", {}))
            return "subset", best_p, res_map, []

        # CẤP 3: TẬP MẸ (Superset Match: S > P)
        superset_candidates: list[dict] = []
        for p in profiles:
            P = set(p.get("source_headers_norm", []))
            if P.issubset(S) and len(P) < len(S):
                superset_candidates.append(p)

        if superset_candidates:
            # Chọn Profile con bao quát được nhiều cột của S nhất
            superset_candidates.sort(key=lambda x: (-len(x.get("source_headers_norm", [])), -x.get("timestamp", 0)))
            best_p = superset_candidates[0]
            P = set(best_p.get("source_headers_norm", []))
            extra_headers = S - P

            # Xác định những cột thừa nào thực sự có dữ liệu trong các dòng mẫu
            extra_indices = [
                i for i, h in enumerate(raw_headers)
                if h is not None and normalize_text_for_match(str(h)) in extra_headers
            ]

            extra_with_data_names: list[str] = []
            if sample_rows:
                for idx in extra_indices:
                    has_val = any(
                        idx < len(row) and row[idx] is not None and str(row[idx]).strip() != ""
                        for row in sample_rows
                    )
                    if has_val:
                        extra_with_data_names.append(str(raw_headers[idx]))
            else:
                extra_with_data_names = [str(raw_headers[i]) for i in extra_indices]

            if not extra_with_data_names:
                # Toàn bộ cột thừa KHÔNG có dữ liệu -> Tự động bỏ qua và áp dụng
                augmented_mapping = dict(best_p.get("mapping", {}))
                for eh in extra_headers:
                    augmented_mapping[eh] = None
                res_map = self.resolve_execution_mapping(raw_headers, augmented_mapping)
                return "superset_auto", best_p, res_map, []
            else:
                # Có cột thừa CHỨA dữ liệu -> Cần hiển thị để người dùng quyết định
                return "superset_ask", best_p, None, extra_with_data_names

        return None, None, None, []

    # -------------------------------------------------------------------------
    # TỰ ĐỘNG GỢI Ý MAPPING ĐẢO NGƯỢC (Heuristics & Fuzzy Match)
    # -------------------------------------------------------------------------

    def auto_suggest_inverted_mapping(
        self,
        source_headers: list[str]
    ) -> dict[str, Any]:
        """
        Tự động tạo gợi ý ánh xạ đảo ngược: {norm_header: target_spec}
        - Nhận diện cột Họ đệm + Tên: [5, 0], [5, 1]
        - Nhận diện cột Ngày + Tháng + Năm: [9, 0], [9, 1], [9, 2]
        - Khớp mờ các cột còn lại với 28 cột chuẩn. Nếu độ khớp thấp thì để None (Bỏ qua).
        """
        inverted_map: dict[str, Any] = {}
        if not source_headers:
            return inverted_map

        # Danh sách chuẩn hóa
        norm_list = [normalize_text_for_match(str(h or "")) for h in source_headers]

        # 1. Phát hiện Họ đệm + Tên tách rời
        ho_idx = None
        ten_idx = None
        HO_KW = {"ho", "ho dem", "ho va chu dem", "ho va ten dem", "ho lot", "hodem", "ho va"}
        TEN_KW = {"ten", "ten hs", "ten hoc sinh", "ten thi sinh"}

        has_full_name = any(n in {"ho va ten", "ho ten", "ho chu dem va ten", "ho va chu dem va ten"} for n in norm_list)
        if not has_full_name:
            for i, n in enumerate(norm_list):
                if not n:
                    continue
                if ho_idx is None and n in HO_KW:
                    ho_idx = i
                elif ten_idx is None and n in TEN_KW:
                    ten_idx = i

            if ho_idx is not None and ten_idx is not None and ho_idx != ten_idx:
                inverted_map[norm_list[ho_idx]] = [5, 0]
                inverted_map[norm_list[ten_idx]] = [5, 1]

        # 2. Phát hiện Ngày + Tháng + Năm tách rời
        ngay_idx, thang_idx, nam_idx = None, None, None
        has_full_dob = any(n in {"ngay sinh", "ngay thang nam sinh", "ns"} for n in norm_list)
        if not has_full_dob:
            for i, n in enumerate(norm_list):
                if not n:
                    continue
                if ngay_idx is None and n in {"ngay", "ngay sinh"}:
                    ngay_idx = i
                elif thang_idx is None and n in {"thang", "thang sinh"}:
                    thang_idx = i
                elif nam_idx is None and n in {"nam", "nam sinh"}:
                    nam_idx = i

            if ngay_idx is not None and thang_idx is not None and nam_idx is not None:
                inverted_map[norm_list[ngay_idx]] = [9, 0]
                inverted_map[norm_list[thang_idx]] = [9, 1]
                inverted_map[norm_list[nam_idx]] = [9, 2]

        # 3. So khớp mờ các cột còn lại với 28 cột chuẩn
        used_targets = {v if isinstance(v, int) else v[0] for v in inverted_map.values() if v is not None}
        for i, n in enumerate(norm_list):
            if not n or n in inverted_map:
                continue

            best_target = None
            best_score = 0.0

            for col in TARGET_COLUMNS:
                t_idx = col["idx"]
                if t_idx in used_targets:
                    continue
                score = compute_fuzzy_similarity(col["name"], n)
                if score > best_score:
                    best_score = score
                    best_target = t_idx

            if best_target is not None and best_score >= 0.55:
                inverted_map[n] = best_target
                used_targets.add(best_target)
            else:
                inverted_map[n] = None

        return inverted_map

    # -------------------------------------------------------------------------
    # CÁC HÀM TƯƠNG THÍCH NGƯỢC (Backward Compatibility)
    # -------------------------------------------------------------------------

    def get_columns_knowledge(self) -> dict[str, list[str | list[str]]]:
        """Tương thích ngược: lấy từ điển tri thức cột."""
        data = self._load_raw()
        return data.get("columns", {})

    def get_learned_terms(self, t_idx: int) -> list[str | list[str]]:
        cols = self.get_columns_knowledge()
        return cols.get(str(t_idx), [])

    def learn_mapping(
        self,
        mapping: dict[int, int | list[int] | None],
        raw_headers: list[str],
        save_prof: bool = True
    ) -> None:
        """Tương thích ngược: bổ sung các từ khóa vào data['columns'] và profiles."""
        if not mapping or not raw_headers:
            return
        data = self._load_raw()
        cols = data.setdefault("columns", {})
        fallback_rules = self.get_fallback_columns()

        for t_idx, s_val in mapping.items():
            if s_val is None:
                continue
            fb_target = fallback_rules.get(t_idx)
            if fb_target and mapping.get(fb_target) == s_val:
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
                    already_exists = any(
                        isinstance(ex, list) and len(ex) == len(combo) and
                        [normalize_text_for_match(c) for c in ex] == combo_norms
                        for ex in existing_terms
                    )
                    if not already_exists:
                        existing_terms.append(combo)

            elif isinstance(s_val, int) and 0 <= s_val < len(raw_headers):
                raw_text = clean_whitespace(auto_convert_legacy_encoding(str(raw_headers[s_val] or "")))
                if raw_text:
                    norm_text = normalize_text_for_match(raw_text)
                    already_exists = any(
                        isinstance(ex, str) and normalize_text_for_match(ex) == norm_text
                        for ex in existing_terms
                    )
                    if not already_exists:
                        existing_terms.append(raw_text)

        data["updated_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self._save_raw(data)

        if save_prof:
            self.save_profile(
                name="Tri thức đã học",
                source_headers=raw_headers,
                mapping=mapping,
                sync_knowledge=False
            )

    def delete_learned_term(self, t_idx: int, term_idx: int) -> bool:
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
        """Tương thích ngược: tìm profile khớp 100%."""
        match_type, prof, res_map, _ = self.find_matching_profile(current_raw_headers)
        if match_type == "exact":
            return prof, res_map
        return None, None

    def get_named_mapping(self, profile: dict) -> dict[str, str | list[str] | None]:
        return profile.get("mapping", {})


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

                # 2. Đọc các dòng mẫu dữ liệu (sample_rows) để kiểm tra cột có dữ liệu hay rỗng
                sample_rows: list[list[Any]] = []
                for row_vals in ws_curr.iter_rows(min_row=h_idx + 1, max_row=min(h_idx + 25, ws_curr.max_row), values_only=True):
                    sample_rows.append(list(row_vals))

                sample_values: dict[int, str] = {}
                for c_idx in range(len(raw_headers)):
                    for r in sample_rows:
                        if c_idx < len(r) and r[c_idx] is not None and str(r[c_idx]).strip():
                            sample_values[c_idx] = str(r[c_idx]).strip()
                            break

                mapping: dict[int, int | list[int] | None] | None = None
                current_defaults: dict[int, str] = {}
                profile_used_name: str = ""

                # 3. Kiểm tra xem có khớp chuẩn 28 cột không
                if is_standard_28_headers(raw_headers):
                    mapping = {col["idx"]: idx for idx, col in enumerate(TARGET_COLUMNS) if idx < len(raw_headers)}
                    current_defaults = {}
                    profile_used_name = "Chuẩn 28 cột gốc"
                    log(f"  ⚡ Sheet '{ws_curr.title}': Cấu trúc chuẩn 28 cột ➔ Ánh xạ tự động 1:1.", "info")
                else:
                    # 4. So khớp theo lý thuyết tập hợp (Trùng khớp 100%, Tập con, Tập mẹ)
                    match_type, matched_prof, resolved_map, extra_cols = self.profile_mgr.find_matching_profile(
                        raw_headers, sample_rows=sample_rows
                    )

                    if match_type in ("exact", "subset", "superset_auto") and resolved_map:
                        mapping = resolved_map
                        profile_used_name = matched_prof.get("name", "Profile")
                        if match_type == "exact":
                            log(f"  ⚡ Sheet '{ws_curr.title}': Tự động áp dụng Profile '{profile_used_name}' (Khớp 100% tiêu đề).", "info")
                        elif match_type == "subset":
                            log(f"  ⚡ Sheet '{ws_curr.title}': Tự động áp dụng Profile '{profile_used_name}' (Khớp tập con, kế thừa quy tắc).", "info")
                        elif match_type == "superset_auto":
                            log(f"  ⚡ Sheet '{ws_curr.title}': Tự động áp dụng Profile '{profile_used_name}' (Khớp tập mẹ, cột thừa không có dữ liệu ➔ Bỏ qua).", "info")
                    else:
                        # Cần sự can thiệp của người dùng:
                        # Tạo gợi ý đảo ngược
                        suggested_inverted = self.profile_mgr.auto_suggest_inverted_mapping(raw_headers)
                        if match_type == "superset_ask" and matched_prof:
                            # Kế thừa ánh xạ từ profile cha cho các cột đã biết
                            for k, v in matched_prof.get("mapping", {}).items():
                                suggested_inverted[k] = v
                            # Tự động bỏ qua các cột thừa KHÔNG có dữ liệu
                            extra_data_norm_set = {normalize_text_for_match(c) for c in extra_cols}
                            for h in raw_headers:
                                nh = normalize_text_for_match(str(h or ''))
                                if nh and nh not in matched_prof.get("mapping", {}) and nh not in extra_data_norm_set:
                                    suggested_inverted[nh] = None
                            log(f"  🔍 Sheet '{ws_curr.title}': Khớp tập mẹ với '{matched_prof.get('name')}', các cột phát hiện thêm có dữ liệu: {extra_cols} ➔ Hiển thị để xác nhận.", "info")

                        suggested_map = self.profile_mgr.resolve_execution_mapping(raw_headers, suggested_inverted)
                        formatted_headers = header_info.get("formatted_headers", [])

                        callback_payload = {
                            "file_name": p.name,
                            "file_path": str(p),
                            "sheet_name": ws_curr.title,
                            "header_row_idx": h_idx,
                            "raw_headers": raw_headers,
                            "formatted_headers": formatted_headers,
                            "suggested_mapping": suggested_map,
                            "suggested_inverted_mapping": suggested_inverted,
                            "sample_rows": sample_rows,
                            "sample_values": sample_values,
                            "extra_cols_with_data": extra_cols,
                            "matched_profile_name": matched_prof.get("name") if matched_prof else None
                        }

                        if mapping_callback:
                            cb_res = mapping_callback(callback_payload)
                            action = "skip"
                            confirmed_map = None
                            prof_name = None
                            custom_inverted = None

                            if isinstance(cb_res, tuple):
                                if cb_res[0] in ("cancel", "skip", "apply"):
                                    action = cb_res[0]
                                    confirmed_map = cb_res[1]
                                    prof_name = cb_res[2] if len(cb_res) > 2 else None
                                    custom_inverted = cb_res[3] if len(cb_res) > 3 else None
                                elif len(cb_res) == 4:
                                    c_map, p_name, _, is_skip = cb_res
                                    if is_skip is True:
                                        action = "skip"
                                    elif is_skip is False and c_map is None:
                                        action = "cancel"
                                    else:
                                        action = "apply"
                                        confirmed_map = c_map
                                        prof_name = p_name
                                elif len(cb_res) == 3:
                                    c_map, p_name, is_skip = cb_res
                                    if is_skip is True:
                                        action = "skip"
                                    elif is_skip is False and c_map is None:
                                        action = "cancel"
                                    else:
                                        action = "apply"
                                        confirmed_map = c_map
                                        prof_name = p_name

                            if action == "cancel":
                                log("⏹ Đã dừng toàn bộ quá trình ghép theo yêu cầu của người dùng.", "warning")
                                return {
                                    "status": "cancelled",
                                    "total_rows": total_extracted_rows,
                                    "sheets_processed": sheets_processed,
                                    "sheets_skipped": sheets_skipped
                                }
                            elif action == "skip" or confirmed_map is None:
                                log(f"  ⏭️ Đã bỏ qua Sheet '{ws_curr.title}' theo lựa chọn của bạn.", "warning")
                                sheets_skipped += 1
                                continue

                            mapping = confirmed_map
                            profile_used_name = prof_name or f"Mẫu {p.stem}_{ws_curr.title}"

                            # Lưu Profile đảo ngược vào kho Profiles
                            self.profile_mgr.save_profile(
                                name=profile_used_name,
                                source_headers=raw_headers,
                                mapping=custom_inverted or mapping
                            )
                            log(f"  💾 Đã lưu cấu hình ánh xạ: '{profile_used_name}' vào kho Profiles.", "success")
                        else:
                            mapping = suggested_map
                            profile_used_name = "Tự động suy luận"

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
