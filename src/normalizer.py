from __future__ import annotations

import re
import math
import datetime
import unicodedata
from collections import Counter
from openpyxl.utils.datetime import from_excel
import difflib

# Bảng mã TCVN3 (ABC / .VnTime) và bảng mã Unicode tương ứng
TCVN3_CHARS = "µ¸¶·¹¨»¾¼½Æ©ÇÊÈÉË®ÌÐÎÏÑªÒÕÓÔÖ×ÝØÜÞßãáâä«åèæçé¬êíëìîïóñòô\u00adõøö÷ùúýûüþ¡¢§£¤¥¦"
UNICODE_CHARS = "àáảãạăằắẳẵặâầấẩẫậđèéẻẽẹêềếểễệìíỉĩịòóỏõọôồốổỗộơờớởỡợùúủũụưừứửữựỳýỷỹỵĂÂĐÊÔƠƯ"

TCVN3_TO_UNI_MAP = dict(zip(TCVN3_CHARS, UNICODE_CHARS))
TCVN3_REGEX = re.compile("|".join(re.escape(c) for c in TCVN3_CHARS))

# Ký tự đặc thù chỉ có ở TCVN3, không bao giờ xuất hiện trong từ tiếng Việt Unicode chuẩn
TCVN3_EXCLUSIVE = set(TCVN3_CHARS) - set(UNICODE_CHARS)
# Ký tự độc quyền của bảng mã Unicode tiếng Việt, không bao giờ xuất hiện trong TCVN3
UNI_EXCLUSIVE = set(UNICODE_CHARS) - set(TCVN3_CHARS)

# Bảng mã VNI-Windows (VNI-Times / VNI-Helve...) sang Unicode
VNI_PAIRS = [
    # 4-char / 3-char combinations (ưu tiên thay thế trước)
    ("öôù", "ướ"), ("öôø", "ườ"), ("öôû", "ưở"), ("öôõ", "ưỡng"), ("öôï", "ược"),
    ("ÖÔÙ", "ƯỚ"), ("ÖÔØ", "ƯỜ"), ("ÖÔÛ", "ƯỞ"), ("ÖÔÕ", "ƯỠ"), ("ÖÔÏ", "ƯỢ"),
    ("Öôù", "Ướ"), ("Öôø", "Ườ"), ("Öôû", "Ưở"), ("Öôõ", "Ưỡ"), ("Öôï", "Ượ"),
    ("aâù", "ấ"), ("aâø", "ầ"), ("aâuû", "ẩ"), ("aâõ", "ẫ"), ("aâï", "ậ"),
    ("AÂÙ", "Ấ"), ("AÂØ", "Ầ"), ("AÂUÛ", "Ẩ"), ("AÂÕ", "Ẫ"), ("AÂÏ", "Ậ"),
    ("Aâù", "Ấ"), ("Aâø", "Ầ"), ("Aâuû", "Ẩ"), ("Aâõ", "Ẫ"), ("Aâï", "Ậ"),
    ("eâù", "ế"), ("eâø", "ề"), ("eâuû", "ể"), ("eâõ", "ễ"), ("eâï", "ệ"),
    ("EÂÙ", "Ế"), ("EÂØ", "Ề"), ("EÂUÛ", "Ể"), ("EÂÕ", "Ễ"), ("EÂÏ", "Ệ"),
    ("Eâù", "Ế"), ("Eâø", "Ề"), ("Eâuû", "Ể"), ("Eâõ", "Ễ"), ("Eâï", "Ệ"),
    ("oâù", "ố"), ("oâø", "ồ"), ("oâuû", "ổ"), ("oâõ", "ỗ"), ("oâï", "ộ"),
    ("OÂÙ", "Ố"), ("OÂØ", "Ồ"), ("OÂUÛ", "Ổ"), ("OÂÕ", "Ỗ"), ("OÂÏ", "Ộ"),
    ("Oâù", "Ố"), ("Oâø", "Ồ"), ("Oâuû", "Ổ"), ("Oâõ", "Ỗ"), ("Oâï", "Ộ"),
    ("aêù", "ắ"), ("aêø", "ằ"), ("aêuû", "ẳ"), ("aêõ", "ẵ"), ("aêï", "ặ"),
    ("AÊÙ", "Ắ"), ("AÊØ", "Ằ"), ("AÊUÛ", "Ẳ"), ("AÊÕ", "Ẵ"), ("AÊÏ", "Ặ"),
    ("Aêù", "Ắ"), ("Aêø", "Ằ"), ("Aêuû", "Ẳ"), ("Aêõ", "Ẵ"), ("Aêï", "Ặ"),

    # 2-char combinations
    ("öô", "ươ"), ("ÖÔ", "ƯƠ"), ("Öô", "Ươ"),
    ("öù", "ứ"), ("öø", "ừ"), ("öû", "ử"), ("öõ", "ữ"), ("öï", "ự"),
    ("ÖÙ", "Ứ"), ("ÖØ", "Ừ"), ("ÖÛ", "Ử"), ("ÖÕ", "Ữ"), ("ÖÏ", "Ự"),
    ("Öù", "Ứ"), ("Öø", "Ừ"), ("Öû", "Ử"), ("Öõ", "Ữ"), ("Öï", "Ự"),
    ("ôù", "ớ"), ("ôø", "ờ"), ("ôû", "ở"), ("ôõ", "ỡ"), ("ôï", "ợ"),
    ("ÔÙ", "Ớ"), ("ÔØ", "Ờ"), ("ÔÛ", "Ở"), ("ÔÕ", "Ỡ"), ("ÔÏ", "Ợ"),
    ("Ôù", "Ớ"), ("Ôø", "Ờ"), ("Ôû", "Ở"), ("Ôõ", "Ỡ"), ("Ôï", "Ợ"),
    ("aù", "á"), ("aø", "à"), ("aû", "ả"), ("aõ", "ã"), ("aï", "ạ"),
    ("AÙ", "Á"), ("AØ", "À"), ("AÛ", "Ả"), ("AÕ", "Ã"), ("AÏ", "Ạ"),
    ("Aù", "Á"), ("Aø", "À"), ("Aû", "Ả"), ("Aõ", "Ã"), ("Aï", "Ạ"),
    ("aê", "ă"), ("aé", "ắ"), ("aè", "ằ"), ("aú", "ẳ"), ("aü", "ẵ"), ("aë", "ặ"),
    ("AÊ", "Ă"), ("AÉ", "Ắ"), ("AÈ", "Ằ"), ("AÚ", "Ẳ"), ("AÜ", "Ẵ"), ("AË", "Ặ"),
    ("Aê", "Ă"), ("Aé", "Ắ"), ("Aè", "Ằ"), ("Aú", "Ẳ"), ("Aü", "Ẵ"), ("Aë", "Ặ"),
    ("aâ", "â"), ("aá", "ấ"), ("aà", "ầ"), ("aå", "ẩ"), ("aã", "ẫ"), ("aä", "ậ"),
    ("AÂ", "Â"), ("AÁ", "Ấ"), ("AÀ", "Ầ"), ("AÅ", "Ẩ"), ("AÃ", "Ẫ"), ("AÄ", "Ậ"),
    ("Aâ", "Â"), ("Aá", "Ấ"), ("Aà", "Ầ"), ("Aå", "Ẩ"), ("Aã", "Ẫ"), ("Aä", "Ậ"),
    ("eù", "é"), ("eø", "è"), ("eû", "ẻ"), ("eõ", "ẽ"), ("eï", "ẹ"),
    ("EÙ", "É"), ("EØ", "È"), ("EÛ", "Ẻ"), ("EÕ", "Ẽ"), ("EÏ", "Ẹ"),
    ("Eù", "É"), ("Eø", "È"), ("Eû", "Ẻ"), ("Eõ", "Ẽ"), ("Eï", "Ẹ"),
    ("eâ", "ê"), ("eá", "ế"), ("eà", "ề"), ("eå", "ể"), ("eã", "ễ"), ("eä", "ệ"),
    ("EÂ", "Ê"), ("EÁ", "Ế"), ("EÀ", "Ề"), ("EÅ", "Ể"), ("EÃ", "Ễ"), ("EÄ", "Ệ"),
    ("Eâ", "Ê"), ("Eá", "Ế"), ("Eà", "Ề"), ("Eå", "Ể"), ("Eã", "Ễ"), ("Eä", "Ệ"),
    ("où", "ó"), ("oø", "ò"), ("oû", "ỏ"), ("oõ", "õ"), ("oï", "ọ"),
    ("OÙ", "Ó"), ("OØ", "Ò"), ("OÛ", "Ỏ"), ("OÕ", "Õ"), ("OÏ", "Ọ"),
    ("Où", "Ó"), ("Oø", "Ò"), ("Oû", "Ỏ"), ("Oõ", "Õ"), ("Oï", "Ọ"),
    ("oâ", "ô"), ("oá", "ố"), ("oà", "ồ"), ("oå", "ổ"), ("oã", "ỗ"), ("oä", "ộ"),
    ("OÂ", "Ô"), ("OÁ", "Ố"), ("OÀ", "Ồ"), ("OÅ", "Ổ"), ("OÃ", "Ỗ"), ("OÄ", "Ộ"),
    ("Oâ", "Ô"), ("Oá", "Ố"), ("Oà", "Ồ"), ("Oå", "Ổ"), ("Oã", "Ỗ"), ("Oä", "Ộ"),
    ("uù", "ú"), ("uø", "ù"), ("uû", "ủ"), ("uõ", "ũ"), ("uï", "ụ"),
    ("UÙ", "Ú"), ("UØ", "Ù"), ("UÛ", "Ủ"), ("UÕ", "Ũ"), ("UÏ", "Ụ"),
    ("Uù", "Ú"), ("Uø", "Ù"), ("Uû", "Ủ"), ("Uõ", "Ũ"), ("Uï", "Ụ"),
    ("yù", "ý"), ("yø", "ỳ"), ("yû", "ỷ"), ("yõ", "ỹ"), ("yï", "ỵ"),
    ("YÙ", "Ý"), ("YØ", "Ỳ"), ("YÛ", "Ỷ"), ("YÕ", "Ỹ"), ("YÏ", "Ỵ"),
    ("Yù", "Ý"), ("Yø", "Ỳ"), ("Yû", "Ỷ"), ("Yõ", "Ỹ"), ("Yï", "Ỵ"),
    ("iù", "í"), ("iø", "ì"), ("iû", "ỉ"), ("iõ", "ĩ"), ("iï", "ị"),
    ("IÙ", "Í"), ("IØ", "Ì"), ("IÛ", "Ỉ"), ("IÕ", "Ĩ"), ("IÏ", "Ị"),
    ("Iù", "Í"), ("Iø", "Ì"), ("Iû", "Ỉ"), ("Iõ", "Ĩ"), ("Iï", "Ị"),

    # 1-char combinations
    ("ñ", "đ"), ("Ñ", "Đ"),
    ("ô", "ơ"), ("Ô", "Ơ"),
    ("ö", "ư"), ("Ö", "Ư"),
    ("ò", "ị"), ("Ò", "Ị"),
    ("æ", "ỉ"), ("Æ", "Ỉ"),
    ("î", "ỵ"), ("Î", "Ỵ")
]

# Khử trùng lặp và sắp xếp theo độ dài giảm dần
_seen_vni = set()
_unique_vni_pairs = []
for _k, _v in VNI_PAIRS:
    if _k not in _seen_vni:
        _seen_vni.add(_k)
        _unique_vni_pairs.append((_k, _v))
_unique_vni_pairs.sort(key=lambda x: len(x[0]), reverse=True)

VNI_TO_UNI_MAP = dict(_unique_vni_pairs)
VNI_REGEX = re.compile("|".join(re.escape(k) for k in VNI_TO_UNI_MAP.keys()))

# Biểu thức chính quy phát hiện dấu hiệu đặc thù của bảng mã VNI-Windows
VNI_EXCLUSIVE_PATTERN = re.compile(
    r"(?i:[aeouyöô][ùøûï]|a[êeéèúüëâáàåãä]|e[âáàåãä]|o[âåãä]|öô|ñ|ö)"
)

def is_vni(text: str) -> bool:
    """Kiểm tra xem chuỗi có dấu hiệu mã hóa VNI-Windows (VNI-Times) hay không."""
    if not text or not isinstance(text, str) or text.isascii():
        return False
    # Nếu chuỗi chứa các ký tự Unicode tiếng Việt cấp cao (ord >= 256) -> chắc chắn là Unicode
    if any(ord(c) >= 256 for c in text):
        return False
    return bool(VNI_EXCLUSIVE_PATTERN.search(text))

def vni_to_unicode(text: str) -> str:
    """Chuyển đổi chuỗi từ bảng mã VNI-Windows sang Unicode dựng sẵn chuẩn."""
    if not text or not isinstance(text, str):
        return "" if text is None else str(text)
    return VNI_REGEX.sub(lambda m: VNI_TO_UNI_MAP.get(m.group(0), m.group(0)), text)

def is_tcvn3(text: str) -> bool:
    """Kiểm tra xem chuỗi có dấu hiệu mã hóa TCVN3 (ABC/.VnTime) hay không."""
    if not text or not isinstance(text, str) or text.isascii():
        return False
    # Nếu chuỗi chứa các ký tự đặc trưng riêng của Unicode tiếng Việt -> chắc chắn là Unicode
    if any(c in UNI_EXCLUSIVE for c in text):
        return False
    # Ưu tiên kiểm tra VNI trước nếu có dấu hiệu ghép đôi
    if is_vni(text):
        return False
    # Nếu chuỗi chứa bất kỳ ký tự đặc thù độc quyền nào của TCVN3 -> là TCVN3
    return any(c in TCVN3_EXCLUSIVE for c in text)

def tcvn3_to_unicode(text: str) -> str:
    """Chuyển đổi chuỗi từ bảng mã TCVN3 (ABC/.VnTime) sang Unicode dựng sẵn."""
    if not text or not isinstance(text, str):
        return "" if text is None else str(text)
    return TCVN3_REGEX.sub(lambda m: TCVN3_TO_UNI_MAP.get(m.group(0), m.group(0)), text)

def auto_convert_legacy_encoding(text: str) -> str:
    """Tự động phát hiện và chuyển đổi chuỗi từ TCVN3 hoặc VNI-Windows sang Unicode (nếu đã là Unicode thì giữ nguyên)."""
    if not isinstance(text, str):
        return text
    if is_vni(text):
        return vni_to_unicode(text)
    if is_tcvn3(text):
        return tcvn3_to_unicode(text)
    return text

# Giữ nguyên alias auto_convert_tcvn3 để tương thích ngược
auto_convert_tcvn3 = auto_convert_legacy_encoding

def clean_whitespace(value):
    """Làm sạch khoảng trắng thừa trong chuỗi, tự động chuyển đổi TCVN3/VNI sang Unicode nếu phát hiện, xử lý nan/none."""
    if value is None:
        return ""
    if isinstance(value, float):
        if math.isnan(value):
            return ""
    if isinstance(value, str):
        if is_vni(value):
            value = vni_to_unicode(value)
        elif is_tcvn3(value):
            value = tcvn3_to_unicode(value)
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
    # Cắt bỏ phần giờ nếu có (ví dụ ISO datetime: "2022-06-25 00:00:00" -> "2022-06-25")
    val_clean = re.sub(r'\s+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?$', '', val_clean)

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
    # Cắt bỏ phần giờ nếu có (ví dụ ISO datetime: "2022-06-25 00:00:00" -> "2022-06-25")
    val_clean = re.sub(r'\s+\d{1,2}:\d{2}(:\d{2})?(\.\d+)?$', '', val_clean)

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

    # Dạng 2 số - 2 số (ví dụ: 18-22 hoặc 18_22)
    m2 = re.search(r'\b(\d{2})\s*[-/–—._~]\s*(\d{2})\b', s_clean)
    if m2:
        y1_2d, y2_2d = int(m2.group(1)), int(m2.group(2))
        full_y1 = (1900 + y1_2d) if y1_2d >= 80 else (2000 + y1_2d)
        full_y2 = (1900 + y2_2d) if y2_2d >= 80 else (2000 + y2_2d)
        if full_y2 > full_y1:
            return (full_y2 - full_y1 == 4), full_y1, full_y2, f"{full_y1}-{full_y2}"

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
    Chuẩn hóa và điền ngày tháng năm cấp bằng (Cột S) theo quy trình 2 bước chuẩn:
    + Bước 1: Quét và tìm những ngày cấp bằng có dạng chuẩn dd/MM/yyyy cùng với Tên trường và Niên khóa tương ứng.
      Hình thành bảng dữ liệu: Tên trường | Niên khóa | Ngày cấp bằng, sắp xếp theo thứ tự từ ngày có nhiều HS nhất đến ít nhất.
      Những HS có ngày cấp bằng hợp lệ dd/MM/yyyy thì giữ nguyên.
    + Bước 2: Tìm những ngày cấp bằng không đúng:
      - Dạng MM/yyyy: tìm trong danh sách ngày hợp lệ của đúng (Trường, Niên khóa) ngày có cùng tháng và năm (ưu tiên ngày nhiều HS nhất).
        Nếu không thấy -> bỏ trống "".
      - Dạng yyyy: tìm trong danh sách ngày hợp lệ của đúng (Trường, Niên khóa) ngày có cùng năm, lấy ngày có nhiều HS hơn.
        Nếu không thấy -> bỏ trống "".
      - Dạng trống hoặc trường hợp khác: bỏ trống "".
    """
    # 1. Thu thập danh sách ngày chuẩn dd/MM/yyyy theo (Tên trường, Niên khóa)
    valid_map: dict[tuple[str, str], list[tuple[str, int]]] = {}
    school_cohort_dates: dict[tuple[str, str], list[str]] = {}

    for st in students:
        sch = clean_whitespace(str(st.get("school") or "")).lower()
        nk = clean_whitespace(str(st.get("cohort") or "")).lower()
        key = (sch, nk)

        g_val = st.get("grad_date", "")
        if not is_cell_empty(g_val):
            formatted = format_grad_date_to_text(g_val)
            is_v, _, _ = check_dob_validity(formatted)
            if is_v:
                school_cohort_dates.setdefault(key, []).append(formatted)

    # Đếm tần suất và sắp xếp giảm dần theo số lượng HS
    for key, date_list in school_cohort_dates.items():
        valid_map[key] = Counter(date_list).most_common()

    # Bước 2: Xử lý các ngày chưa đúng chuẩn
    for st in students:
        sch = clean_whitespace(str(st.get("school") or "")).lower()
        nk = clean_whitespace(str(st.get("cohort") or "")).lower()
        key = (sch, nk)
        valid_dates_for_group = valid_map.get(key, [])

        curr_grad = st.get("grad_date", "")
        formatted = format_grad_date_to_text(curr_grad)

        # 1. Đã là dd/MM/yyyy hợp lệ -> Giữ nguyên
        is_v, _, _ = check_dob_validity(formatted)
        if is_v:
            st["grad_date"] = formatted
            continue

        # 2. Dạng MM/yyyy -> Tìm ngày chuẩn cùng tháng & năm
        m_my = re.match(r'^(\d{2})/(\d{4})$', formatted)
        if m_my:
            target_m, target_y = m_my.groups()
            matched_date = None
            for d_str, _ in valid_dates_for_group:
                if d_str.endswith(f"/{target_m}/{target_y}"):
                    matched_date = d_str
                    break
            st["grad_date"] = matched_date if matched_date else ""
            continue

        # 3. Dạng yyyy -> Tìm ngày chuẩn cùng năm có nhiều HS nhất
        m_y = re.match(r'^(\d{4})$', formatted)
        if m_y:
            target_y = m_y.group(1)
            matched_date = None
            for d_str, _ in valid_dates_for_group:
                if d_str.endswith(f"/{target_y}"):
                    matched_date = d_str
                    break
            st["grad_date"] = matched_date if matched_date else ""
            continue

        # 4. Trường hợp trống hoặc khác -> Bỏ trống
        st["grad_date"] = ""

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
