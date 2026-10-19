import re
import math
import datetime
import unicodedata
from collections import Counter
from openpyxl.utils.datetime import from_excel
import difflib

def clean_whitespace(value):
    """Làm sạch khoảng trắng thừa trong chuỗi, xử lý nan/none."""
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
    if isinstance(value, str):
        cleaned = unicodedata.normalize('NFC', value)
        cleaned = re.sub(r'\s+', ' ', cleaned).strip()
        if cleaned.lower() in ("none", "nan", "null", "<na>"):
            return ""
        return cleaned
    return value

def is_cell_empty(val):
    """Kiểm tra ô có bị trống dữ liệu không."""
    if val is None:
        return True
    if isinstance(val, float):
        if math.isnan(val):
            return True
    s = str(val).strip()
    return s == "" or s.lower() in ["none", "nan", "null", "<na>"]

def vietnamese_title_case(text: str) -> str:
    """
    Chuyển đổi chuỗi tiếng Việt thành dạng Title Case (viết hoa chữ cái đầu mỗi từ).
    Hỗ trợ chuẩn xác các ký tự tiếng Việt có dấu.
    """
    if not text:
        return ""
    s = clean_whitespace(str(text))
    if not s:
        return ""
    
    # Chuẩn hóa Unicode NFC
    s = unicodedata.normalize('NFC', s)
    
    words = s.split(' ')
    capitalized_words = []
    for w in words:
        if not w:
            continue
        # Nếu từ chứa dấu gạch nối (ví dụ Ê-đê, Ba-na, Bru-Vân Kiều)
        if '-' in w:
            sub_parts = w.split('-')
            cap_sub = [p.capitalize() if p else '' for p in sub_parts]
            capitalized_words.append('-'.join(cap_sub))
        else:
            capitalized_words.append(w.capitalize())
            
    return ' '.join(capitalized_words)

def normalize_birthplace(val, lookup_mgr=None) -> str:
    """
    Chuẩn hóa nơi sinh:
    - Nếu là mã số: tra cứu từ sheet Noi_Sinh.
    - Viết hoa chữ cái đầu mỗi từ.
    - ĐẶC BIỆT: Cụm từ phân loại hành chính 'Thành phố', 'Thị xã', 'Thị trấn' chỉ viết hoa chữ cái đầu tiên!
    """
    if is_cell_empty(val):
        return ""
    
    raw_str = str(val).strip()
    if lookup_mgr:
        mapped = lookup_mgr.map_birthplace(raw_str)
    else:
        mapped = raw_str

    if not mapped:
        return ""

    title_cased = vietnamese_title_case(mapped)
    
    # Hiệu chỉnh các cụm từ hành chính chỉ viết hoa chữ cái đầu tiên
    # "Thành Phố " -> "Thành phố "
    res = re.sub(r'\bThành\s+Phố\b', 'Thành phố', title_cased, flags=re.IGNORECASE)
    # "Thị Xã " -> "Thị xã "
    res = re.sub(r'\bThị\s+Xã\b', 'Thị xã', res, flags=re.IGNORECASE)
    # "Thị Trấn " -> "Thị trấn "
    res = re.sub(r'\bThị\s+Trấn\b', 'Thị trấn', res, flags=re.IGNORECASE)
    
    return res

def normalize_ethnic(val, lookup_mgr=None) -> str:
    """Chuẩn hóa dân tộc: tra cứu mã và viết hoa chữ cái đầu mỗi từ."""
    if is_cell_empty(val):
        return ""
    raw_str = str(val).strip()
    if lookup_mgr:
        mapped = lookup_mgr.map_ethnic(raw_str)
    else:
        mapped = raw_str
    return vietnamese_title_case(mapped)

def normalize_nationality(val, lookup_mgr=None) -> str:
    """Chuẩn hóa quốc tịch: tra cứu mã và viết hoa chữ cái đầu mỗi từ."""
    if is_cell_empty(val):
        return ""
    raw_str = str(val).strip()
    if lookup_mgr:
        mapped = lookup_mgr.map_nationality(raw_str)
    else:
        mapped = raw_str
    return vietnamese_title_case(mapped)

# ==============================================================================
# XỬ LÝ NGÀY THÁNG: NGÀY SINH & NGÀY CẤP BẰNG
# ==============================================================================

def check_dob_validity(val):
    """
    Kiểm tra tính hợp lệ của ngày sinh:
    Trả về: (is_valid, err_type, err_desc)
    - Hợp lệ: (True, 0, "")
    - Kiểu 1 (Sai định dạng dd/MM/yyyy): (False, 1, "Sai định dạng ngày sinh")
    - Kiểu 2 (Sai logic thực tế): (False, 2, "Ngày sinh không hợp lệ: <giải thích>")
    """
    if is_cell_empty(val):
        return False, 1, "Thiếu ngày sinh"
    
    s = str(val).strip()
    m = re.match(r'^(\d{2})/(\d{2})/(\d{4})$', s)
    if not m:
        return False, 1, "Sai định dạng ngày sinh"

    d, m_val, y = map(int, m.groups())

    if m_val < 1 or m_val > 12:
        return False, 2, f"Ngày sinh không hợp lệ: Tháng {m_val:02d} không tồn tại"

    if d == 0:
        return False, 2, "Ngày sinh không hợp lệ: Ngày 00 không tồn tại"

    if not (1900 <= y <= 2100):
        return False, 2, f"Ngày sinh không hợp lệ: Năm sinh {y} không hợp lý (1900-2100)"

    if m_val == 2:
        is_leap = (y % 4 == 0 and y % 100 != 0) or (y % 400 == 0)
        max_d = 29 if is_leap else 28
        if d > max_d:
            if not is_leap and d == 29:
                return False, 2, f"Ngày sinh không hợp lệ: Năm {y} không phải năm nhuận nên tháng 02 không thể có 29 ngày"
            else:
                return False, 2, f"Ngày sinh không hợp lệ: Tháng 02 năm {y} không thể có {d:02d} ngày"
    elif m_val in (4, 6, 9, 11):
        if d > 30:
            return False, 2, f"Ngày sinh không hợp lệ: Tháng {m_val:02d} không thể có {d:02d} ngày"
    else:
        if d > 31:
            return False, 2, f"Ngày sinh không hợp lệ: Tháng {m_val:02d} không thể có {d:02d} ngày"

    try:
        datetime.date(y, m_val, d)
        return True, 0, ""
    except ValueError as e:
        return False, 2, f"Ngày sinh không hợp lệ: {e}"

def format_date_to_text(value):
    """
    Chuyển đổi mọi loại dữ liệu ngày tháng thành văn bản thuần túy dạng dd/MM/yyyy.
    Hỗ trợ datetime, số serial Excel, số liền 7-8 chữ số, lỗi lặp dấu (//, .., --),
    chuỗi ISO, VÀ TỰ ĐỘNG ĐẢO NGƯỢC THÁNG TRƯỚC NGÀY SAU (MM/DD/YYYY -> DD/MM/YYYY) KHI A <= 12 VÀ B > 12.
    """
    if is_cell_empty(value):
        return ""

    # 1. Hỗ trợ datetime.datetime và datetime.date
    if isinstance(value, (datetime.datetime, datetime.date)):
        return value.strftime("%d/%m/%Y")

    # 2. Xử lý số serial ngày tháng của Excel (10000 đến 80000, năm 1927 đến 2119)
    if isinstance(value, (int, float)) and 10000 <= value <= 80000:
        try:
            dt = from_excel(value)
            return dt.strftime("%d/%m/%Y")
        except:
            pass

    val_str = str(value).strip()
    if not val_str or val_str.lower() in ("none", "nan", "null"):
        return ""

    val_clean = re.sub(r'\.0+$', '', val_str)

    # 3. Xử lý dạng số liền 7-8 chữ số (ví dụ: 19061998 hoặc 1061998)
    if val_clean.isdigit() and len(val_clean) in (7, 8):
        padded = val_clean.zfill(8)
        # Thử kiểu DDMMYYYY
        d, m, y = int(padded[:2]), int(padded[2:4]), int(padded[4:])
        if 1900 <= y <= 2100:
            if 1 <= m <= 12 and 1 <= d <= 31:
                try:
                    dt = datetime.date(y, m, d)
                    return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
                except ValueError:
                    pass
            # Thử nếu bị viết ngược MMDDYYYY (m <= 12, d > 12)
            if 1 <= d <= 12 and 12 < m <= 31:
                try:
                    dt = datetime.date(y, d, m)
                    return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
                except ValueError:
                    pass
        # Thử kiểu YYYYMMDD
        y, m, d = int(padded[:4]), int(padded[4:6]), int(padded[6:])
        if 1900 <= y <= 2100:
            try:
                dt = datetime.date(y, m, d)
                return f"{dt.day:02d}/{dt.month:02d}/{dt.year}"
            except ValueError:
                pass

    # 4. Gom các dấu phân cách lặp lại (//, ..., --) và khoảng trắng
    val_norm = re.sub(r'[/.\-\s]+', '/', val_clean).strip('/')
    if not val_norm:
        return ""

    # 5. Chuỗi dạng ISO YYYY/MM/DD -> đảo về DD/MM/YYYY
    m_iso = re.match(r'^(\d{4})/(\d{1,2})/(\d{1,2})$', val_norm)
    if m_iso:
        y, m, d = m_iso.groups()
        return f"{int(d):02d}/{int(m):02d}/{y}"

    # 6. Chuỗi dạng A/B/YYYY
    m_dmy = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', val_norm)
    if m_dmy:
        a_str, b_str, y_str = m_dmy.groups()
        a, b, y = int(a_str), int(b_str), int(y_str)
        # ĐẶC BIỆT: Tự động phát hiện viết ngược MM/DD/YYYY (Tháng <= 12, Ngày > 12)
        if a <= 12 and b > 12:
            # a là tháng, b là ngày -> đảo lại thành DD/MM/YYYY
            return f"{b:02d}/{a:02d}/{y}"
        else:
            # a là ngày, b là tháng
            return f"{a:02d}/{b:02d}/{y}"

    # 6b. Chuỗi dạng A/B/YY (năm 2 chữ số)
    m_dmy2 = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{2})$', val_norm)
    if m_dmy2:
        a_str, b_str, y2_str = m_dmy2.groups()
        a, b, y_int = int(a_str), int(b_str), int(y2_str)
        full_y = 2000 + y_int if y_int <= 35 else 1900 + y_int
        if a <= 12 and b > 12:
            return f"{b:02d}/{a:02d}/{full_y}"
        return f"{a:02d}/{b:02d}/{full_y}"

    # 7. Nếu chỉ có năm (ví dụ "1990")
    m_year = re.match(r'^(\d{4})$', val_norm)
    if m_year:
        return m_year.group(1)

    return val_norm

def format_grad_date_to_text(value):
    """
    Chuẩn hóa Ngày, tháng, năm cấp bằng (Cột S):
    Cho phép: đầy đủ dd/MM/yyyy, tháng/năm MM/yyyy, chỉ có năm yyyy, hoặc để trống "".
    Tự động đảo ngược MM/DD/YYYY -> DD/MM/YYYY nếu A <= 12 và B > 12.
    """
    if is_cell_empty(value):
        return ""
    val_str = str(value).strip()
    if val_str.lower() in ("trong", "trống", "xoa", "xóa"):
        return ""

    val_clean = re.sub(r'\.0+$', '', val_str)

    # Nếu là số serial Excel
    if isinstance(value, (int, float)) and 10000 <= value <= 80000:
        try:
            dt = from_excel(value)
            return dt.strftime("%d/%m/%Y")
        except:
            pass

    # Dạng số liền 5-6 chữ số (062020 -> MM/yyyy)
    if val_clean.isdigit() and len(val_clean) in (5, 6):
        padded = val_clean.zfill(6)
        m, y = int(padded[:2]), int(padded[2:])
        if 1 <= m <= 12 and 1900 <= y <= 2100:
            return f"{m:02d}/{y}"

    # Làm sạch dấu phân cách
    val_norm = re.sub(r'[/.\-\s]+', '/', val_clean).strip('/')
    if not val_norm:
        return ""

    # Dạng A/B/YYYY
    m_dmy = re.match(r'^(\d{1,2})/(\d{1,2})/(\d{4})$', val_norm)
    if m_dmy:
        a_str, b_str, y_str = m_dmy.groups()
        a, b, y = int(a_str), int(b_str), int(y_str)
        if a <= 12 and b > 12:
            return f"{b:02d}/{a:02d}/{y}"
        return f"{a:02d}/{b:02d}/{y}"

    # Dạng MM/yyyy hoặc M/yyyy -> chuẩn hóa MM/yyyy
    m_my = re.match(r'^(\d{1,2})/(\d{4})$', val_norm)
    if m_my:
        m, y = m_my.groups()
        if 1 <= int(m) <= 12 and 1900 <= int(y) <= 2100:
            return f"{int(m):02d}/{y}"

    # Dạng chỉ có 4 chữ số năm yyyy
    m_year = re.match(r'^(\d{4})$', val_norm)
    if m_year:
        y_int = int(m_year.group(1))
        if 1900 <= y_int <= 2100:
            return m_year.group(1)

    return format_date_to_text(value)

def is_valid_grad_date(val):
    """Kiểm tra ngày cấp bằng có hợp lệ không (hợp lệ: dd/MM/yyyy, MM/yyyy, yyyy, hoặc rỗng)."""
    if is_cell_empty(val):
        return True
    s = str(val).strip()
    is_v, _, _ = check_dob_validity(s)
    if is_v:
        return True
    if re.match(r'^(0[1-9]|1[0-2])/\d{4}$', s):
        return True
    if re.match(r'^(19\d{2}|20\d{2})$', s):
        return True
    return False

# ==============================================================================
# NIÊN KHÓA & SUY LUẬN NIÊN KHÓA TỪ NĂM SINH
# ==============================================================================

def parse_nien_khoa(val_nien_khoa):
    """
    Kiểm tra và bóc tách Niên khóa:
    - Chuẩn hóa về dạng YYYY-YYYY.
    - Trả về: (is_valid_diff_4, y_vao, y_ra, formatted_str)
    """
    if is_cell_empty(val_nien_khoa):
        return False, None, None, ""
    s_clean = clean_whitespace(str(val_nien_khoa))
    s_clean = re.sub(r'\.0+$', '', s_clean)

    m = re.search(r'(\d{4})\s*[-/–—._~]\s*(\d{2,4})', s_clean)
    if m:
        y1_str, y2_str = m.groups()
        y1 = int(y1_str)
        if len(y2_str) == 2:
            y2_val = int(y2_str)
            century = (y1 // 100) * 100
            if y2_val < (y1 % 100):
                y2 = century + 100 + y2_val
            else:
                y2 = century + y2_val
        else:
            y2 = int(y2_str)

        return (y2 - y1 == 4), y1, y2, f"{y1}-{y2}"

    # Nếu chỉ có 1 năm lẻ
    m_single = re.search(r'\b(19\d{2}|20\d{2})\b', s_clean)
    if m_single:
        y_ra = int(m_single.group(1))
        return False, y_ra - 4, y_ra, f"{y_ra - 4}-{y_ra}"

    return False, None, None, s_clean

def infer_cohort_from_birth_year(birth_year: int) -> tuple[int, int, str]:
    """
    Suy luận Niên khóa THCS từ Năm sinh của học sinh:
    Học sinh tốt nghiệp lớp 9 ở độ tuổi 15 tuổi:
    - Y_ra = birth_year + 15
    - Y_vao = Y_ra - 4 = birth_year + 11
    - Niên khóa chuẩn: f"{Y_vao}-{Y_ra}"
    """
    y_ra = birth_year + 15
    y_vao = y_ra - 4
    return y_vao, y_ra, f"{y_vao}-{y_ra}"

def extract_birth_year_from_dob(dob_val) -> int | None:
    """Trích xuất năm sinh từ ngày sinh đã chuẩn hóa hoặc chuỗi gốc."""
    if is_cell_empty(dob_val):
        return None
    s = str(dob_val).strip()
    m = re.search(r'\b(19\d{2}|20\d{2})\b', s)
    if m:
        return int(m.group(1))
    return None

def find_mode_birth_year(student_list) -> int | None:
    """Tìm năm sinh phổ biến nhất từ danh sách học sinh."""
    years = []
    for s in student_list:
        dob = s.get("dob", "") if isinstance(s, dict) else (s[8] if len(s) > 8 else "")
        yr = extract_birth_year_from_dob(dob)
        if yr and 1950 <= yr <= 2030:
            years.append(yr)
    if years:
        return Counter(years).most_common(1)[0][0]
    return None

# ==============================================================================
# FUZZY MATCHING TÊN TRƯỜNG
# ==============================================================================

def fuzzy_match_school(raw_school: str, allowed_schools: list[str], n: int = 5, cutoff: float = 0.3) -> list[str]:
    """Tìm các tên trường trong danh mục trực thuộc gần khớp nhất với raw_school."""
    if not raw_school or not allowed_schools:
        return allowed_schools[:n]
    
    clean_raw = clean_whitespace(raw_school).lower()
    
    # 1. Thử khớp gần đúng bằng difflib
    matches = difflib.get_close_matches(clean_raw, [s.lower() for s in allowed_schools], n=n, cutoff=cutoff)
    res = []
    for m in matches:
        for orig in allowed_schools:
            if orig.lower() == m and orig not in res:
                res.append(orig)
                break
                
    # 2. Bổ sung các trường có chứa từ khóa
    words = [w for w in re.split(r'\W+', clean_raw) if len(w) > 2]
    for orig in allowed_schools:
        if orig not in res:
            orig_lower = orig.lower()
            if any(w in orig_lower for w in words):
                res.append(orig)
                if len(res) >= n:
                    break
                    
    # Nếu vẫn chưa đủ, lấy thêm từ danh sách trường
    for orig in allowed_schools:
        if orig not in res:
            res.append(orig)
            if len(res) >= n:
                break
                
    return res

# ==============================================================================
# THUẬT TOÁN TỰ ĐỘNG ĐIỀN NGÀY CẤP BẰNG THEO NHÓM (TRƯỜNG, NIÊN KHÓA)
# ==============================================================================

def auto_fill_graduation_dates_by_school_and_cohort(students: list[dict]):
    """
    Gom nhóm học sinh theo (Tên trường đã học, Niên khóa) và tự động điền Ngày cấp bằng (Cột S):
    - Đã có dd/MM/yyyy: giữ nguyên.
    - Dạng MM/yyyy: tìm ngày cùng tháng/năm trong nhóm -> nếu có lấy ngày nhiều nhất, nếu không lấy Mode_group.
    - Dạng yyyy: tìm ngày cùng năm trong nhóm -> nếu có lấy ngày nhiều nhất, nếu không lấy Mode_group.
    - Để trống: lấy Mode_group.
    - Nếu cả nhóm không có ngày nào: để trống.
    """
    # 1. Gom nhóm
    groups: dict[tuple[str, str], list[dict]] = {}
    for st in students:
        sch = st.get("school", "") or "Chua_Xac_Dinh_Truong"
        nk = st.get("cohort", "") or "Chua_Xac_Dinh_Nien_Khoa"
        key = (sch.strip().lower(), nk.strip().lower())
        if key not in groups:
            groups[key] = []
        groups[key].append(st)

    for key, group_students in groups.items():
        # Thu thập tất cả ngày dd/MM/yyyy hợp lệ thực tế trong nhóm
        valid_dates = []
        for st in group_students:
            g_val = st.get("grad_date", "")
            if g_val:
                formatted = format_grad_date_to_text(g_val)
                is_v, _, _ = check_dob_validity(formatted)
                if is_v:
                    valid_dates.append(formatted)

        mode_group = Counter(valid_dates).most_common(1)[0][0] if valid_dates else None

        for st in group_students:
            curr_grad = st.get("grad_date", "")
            formatted = format_grad_date_to_text(curr_grad)
            
            # 1. Đã là dd/MM/yyyy hợp lệ -> Giữ nguyên
            is_v, _, _ = check_dob_validity(formatted)
            if is_v:
                st["grad_date"] = formatted
                continue

            # 2. Dạng MM/yyyy
            m_my = re.match(r'^(\d{2})/(\d{4})$', formatted)
            if m_my:
                target_m, target_y = m_my.groups()
                same_my = [d for d in valid_dates if d.endswith(f"/{target_m}/{target_y}")]
                if same_my:
                    st["grad_date"] = Counter(same_my).most_common(1)[0][0]
                elif mode_group:
                    st["grad_date"] = mode_group
                else:
                    st["grad_date"] = formatted
                continue

            # 3. Dạng yyyy
            m_y = re.match(r'^(\d{4})$', formatted)
            if m_y:
                target_y = m_y.group(1)
                same_y = [d for d in valid_dates if d.endswith(f"/{target_y}")]
                if same_y:
                    st["grad_date"] = Counter(same_y).most_common(1)[0][0]
                elif mode_group:
                    st["grad_date"] = mode_group
                else:
                    st["grad_date"] = formatted
                continue

            # 4. Trường hợp để trống
            if is_cell_empty(formatted):
                if mode_group:
                    st["grad_date"] = mode_group
                else:
                    st["grad_date"] = ""
            else:
                if mode_group:
                    st["grad_date"] = mode_group
                else:
                    st["grad_date"] = formatted

# ==============================================================================
# XÁC ĐỊNH TRẠNG THÁI SỐ HÓA (CỘT Y) VÀ NGUYÊN NHÂN (CỘT Z)
# ==============================================================================

def determine_status_and_reason(dob_raw, dob_formatted, is_valid_dob, dob_err_desc):
    """
    Quy tắc ấn định Cột Y và Cột Z:
    - Chỉ 'Chưa số hoá' khi thiếu hẳn ngày sinh -> Cột Z: 'Thiếu ngày sinh'.
    - Nếu có ngày sinh nhưng không chính xác (sai format hoặc sai logic) -> 'Số hoá chưa đầy đủ, chính xác', Cột Z ghi rõ nguyên nhân.
    - Còn lại (ngày sinh đầy đủ và chính xác) -> 'Đã số hóa đầy đủ, chính xác', Cột Z: ''.
    """
    if is_cell_empty(dob_raw) and is_cell_empty(dob_formatted):
        return "Chưa số hoá", "Thiếu ngày sinh"
    
    if not is_valid_dob:
        reason = dob_err_desc if dob_err_desc else "Sai định dạng ngày sinh"
        return "Số hoá chưa đầy đủ, chính xác", reason
        
    return "Đã số hóa đầy đủ, chính xác", ""
