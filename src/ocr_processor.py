from __future__ import annotations

import base64
import io
import json
import os
import re
import ssl
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
from pathlib import Path
from typing import Any, Callable

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from src.lookup_manager import find_data_file, get_persistent_data_dir

# Thử nạp PIL và pillow_heif nếu môi trường có cài đặt
try:
    from PIL import Image
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    from pillow_heif import register_heif_opener
    register_heif_opener()
    HAS_HEIF = True
except ImportError:
    HAS_HEIF = False


VALID_IMAGE_EXTENSIONS = {".jpg", ".jpeg", ".png", ".heic", ".heif"}

DEFAULT_BASE_URL = "https://api.xah.io/v1"
DEFAULT_MODEL_NAME = "dungcsnd113/gemini-flash-3.8"
AVAILABLE_OCR_MODELS = [
    "dungcsnd113/gemini-flash-3.8",
    "dungcsnd113/gemini-pro-3.1",
]

DEFAULT_PROMPT_TEXT = """Bạn là hệ thống OCR tài liệu độ chính xác cao.
Hãy trích xuất toàn bộ dữ liệu từ bảng "DANH SÁCH HỌC SINH ĐƯỢC CÔNG NHẬN TỐT NGHIỆP" trong ảnh.

QUY TẮC XỬ LÝ DỮ LIỆU SỬA ĐỔI (CHỮ VIẾT TAY / MỰC ĐỎ):
1. Ưu tiên dữ liệu sửa đổi: Nếu có dữ liệu in máy bị gạch bỏ và được sửa lại bằng chữ viết tay (thường là mực đỏ), BẮT BUỘC phải lấy phần chữ viết tay làm giá trị cuối cùng. Bỏ qua phần chữ in máy đã bị gạch.
2. Dữ liệu viết tay bổ sung: Nếu có chữ viết tay (mực đỏ) ghi chú thêm bên cạnh chữ in máy (ví dụ các con số, mũi tên ở cột Ghi chú), hãy ghi nhận đầy đủ cả phần in máy và phần viết tay.

YÊU CẦU ĐẦU RA:
- Trả về KẾT QUẢ DUY NHẤT là một chuỗi JSON chuẩn (mảng các object).
- Không kèm theo markdown block (như ```json) và không có bất kỳ văn bản giải thích nào khác.
- Cấu trúc key của từng dòng:
  + stt: Số thứ tự
  + sbd: Số báo danh
  + ho_ten: Họ và tên thí sinh
  + gioi_tinh: "Nam" hoặc "Nữ" (dựa theo cột GT. Chú ý chữ viết tay sửa đổi).
  + ngay_sinh: Ngày sinh (Chú ý các năm sinh bị gạch xóa và sửa bằng mực đỏ).
  + noi_sinh: Nơi sinh
  + hoi_dong_thi: Hội đồng thi (hoặc trường)
  + xep_loai: Xếp loại TN
  + ghi_chu: Ghi chú (Bao gồm cả các số liệu in máy và các số hiệu ghi tay/mực đỏ bổ sung).
  + can_kiem_tra: Điền true nếu bạn cảm thấy phân vân, không chắc chắn về bất kỳ dữ liệu nào trên hàng này (do chữ mờ, lem mực, viết tay khó đọc). Nếu chắc chắn đọc đúng 100% thì để trống.
  + ly_do_nghi_ngo: Nếu can_kiem_tra là true, hãy ghi ngắn gọn lý do (VD: "Ngày sinh viết tay bị nhòe", "Tên đệm mờ"). Nếu false, để null.
- Nếu ô dữ liệu trống hoặc không rõ, điền null."""

ORIGINAL_OCR_COLUMNS = [
    ("stt", "STT", 8),
    ("sbd", "SBD", 12),
    ("ho_ten", "Họ và tên", 26),
    ("gioi_tinh", "Giới tính", 11),
    ("ngay_sinh", "Ngày sinh", 14),
    ("noi_sinh", "Nơi sinh", 22),
    ("hoi_dong_thi", "Hội đồng thi", 26),
    ("xep_loai", "Xếp loại", 12),
    ("ghi_chu", "Ghi chú", 22),
    ("can_kiem_tra", "Cần kiểm tra", 14),
    ("ly_do_nghi_ngo", "Lý do nghi ngờ", 30),
]


def natural_sort_key(file_path: str | Path) -> tuple[list[int | str], str]:
    """
    Tách tên file thành các phần chữ và số để sắp xếp tự nhiên (Natural Sort).
    Đảm bảo thứ tự 1, 2, ..., 9, 10 cũng như 01, 02, ..., 10.
    """
    name = Path(file_path).name
    tokens: list[int | str] = [
        int(text) if text.isdigit() else text.lower()
        for text in re.split(r"(\d+)", name)
    ]
    return (tokens, name)


def encode_image(
    image_path: str | Path,
    max_dimension: int = 1800,
    jpeg_quality: int = 85,
) -> str:
    """
    Đọc ảnh (tự động convert HEIC/HEIF sang JPEG, nén và giới hạn độ phân giải tối ưu)
    rồi chuyển sang chuỗi base64 để tăng tốc độ truyền tải qua API OCR gấp 3-5 lần.
    Hỗ trợ:
    1. PIL + pillow_heif (nếu đã cài đặt).
    2. Công cụ native `/usr/bin/sips` trên macOS cho ảnh HEIC/HEIF/JPG/PNG dung lượng lớn.
    3. Đọc trực tiếp bytes đối với ảnh nhỏ hoặc khi không có công cụ nén.
    """
    path = Path(image_path)
    if not path.exists() or not path.is_file():
        raise FileNotFoundError(f"Không tìm thấy tệp ảnh: {path}")

    suffix = path.suffix.lower()
    file_size = path.stat().st_size

    # 1. Nếu có PIL, dùng PIL để resize & nén chuẩn xác
    if HAS_PIL and (suffix not in (".heic", ".heif") or HAS_HEIF):
        try:
            img = Image.open(path)
            if img.mode != "RGB":
                img = img.convert("RGB")
            w, h = img.size
            if max_dimension > 0 and max(w, h) > max_dimension:
                ratio = max_dimension / float(max(w, h))
                img = img.resize((int(w * ratio), int(h * ratio)), Image.Resampling.LANCZOS)
            buffer = io.BytesIO()
            img.save(buffer, format="JPEG", quality=jpeg_quality, optimize=True)
            return base64.b64encode(buffer.getvalue()).decode("utf-8")
        except Exception:
            pass

    # 2. Trên macOS: dùng /usr/bin/sips cho ảnh HEIC/HEIF hoặc ảnh nặng > 800KB
    if sys.platform == "darwin" and Path("/usr/bin/sips").exists():
        if suffix in (".heic", ".heif") or (max_dimension > 0 and file_size > 800 * 1024):
            with tempfile.NamedTemporaryFile(suffix=".jpg", delete=False) as tmp_f:
                tmp_jpg = Path(tmp_f.name)
            try:
                cmd = ["/usr/bin/sips", "-s", "format", "jpeg"]
                if jpeg_quality > 0:
                    cmd.extend(["-s", "formatOptions", str(jpeg_quality)])
                if max_dimension > 0:
                    cmd.extend(["-Z", str(max_dimension)])
                cmd.extend([str(path.resolve()), "--out", str(tmp_jpg)])
                res = subprocess.run(
                    cmd,
                    capture_output=True,
                    timeout=30,
                    check=False,
                )
                if res.returncode == 0 and tmp_jpg.exists() and tmp_jpg.stat().st_size > 0:
                    data = tmp_jpg.read_bytes()
                    return base64.b64encode(data).decode("utf-8")
            finally:
                if tmp_jpg.exists():
                    try:
                        tmp_jpg.unlink()
                    except Exception:
                        pass

    with open(path, "rb") as image_file:
        return base64.b64encode(image_file.read()).decode("utf-8")


def _parse_dotenv_file(env_path: Path) -> dict[str, str]:
    """Đọc các cặp KEY=VALUE từ tệp .env mà không cần phụ thuộc bắt buộc vào python-dotenv."""
    result: dict[str, str] = {}
    if not env_path.is_file():
        return result
    try:
        content = env_path.read_text(encoding="utf-8", errors="ignore")
        for line in content.splitlines():
            line_s = line.strip()
            if not line_s or line_s.startswith("#") or "=" not in line_s:
                continue
            k, v = line_s.split("=", 1)
            k = k.strip()
            v = v.strip().strip("'\"")
            if k:
                result[k] = v
    except Exception:
        pass
    return result


class OCRConfigManager:
    """
    Quản lý cấu hình OCR (API Key, Base URL, Model, Default Prompt):
    - Tự động đọc từ biến môi trường, tệp `.env` (trong thư mục dự án hoặc dự án `so-hoa-vbcc` lân cận)
      và tệp lưu trữ bền vững `ocr_config.json`.
    - Hỗ trợ tìm kiếm Prompt 3 cấp độ:
      1. `local_prompt`: `<thư_mục_ảnh>/prompt.txt` (Prompt riêng của trường)
      2. `year_prompt`: `<thư_mục_cha>/prompt.txt` (Prompt chung của năm học)
      3. `default_prompt`: `data/default_prompt.txt` (Prompt mặc định của hệ thống)
    """

    def __init__(self, config_path: Path | str | None = None):
        if config_path:
            self.config_path = Path(config_path)
        else:
            self.config_path = get_persistent_data_dir() / "ocr_config.json"

    def load_config(self) -> dict[str, str]:
        """Nạp cấu hình kết hợp từ .env, biến môi trường và ocr_config.json."""
        env_vars: dict[str, str] = {}

        # Quét các vị trí có thể chứa tệp .env
        candidate_envs = [
            Path.cwd() / ".env",
            Path(__file__).resolve().parent.parent / ".env",
            Path.cwd().parent / "so-hoa-vbcc" / ".env",
            Path.home() / "Desktop" / "so-hoa-vbcc" / ".env",
        ]
        for env_p in candidate_envs:
            if env_p.is_file():
                parsed = _parse_dotenv_file(env_p)
                for k, v in parsed.items():
                    if k not in env_vars and v:
                        env_vars[k] = v

        # Ưu tiên biến môi trường hệ thống nếu có
        for key in ("CKEY_API_KEY", "CKEY_BASE_URL", "GEMINI_MODEL"):
            os_val = os.environ.get(key)
            if os_val:
                env_vars[key] = os_val

        saved_cfg: dict[str, str] = {}
        if self.config_path.is_file():
            try:
                saved_cfg = json.loads(self.config_path.read_text(encoding="utf-8"))
                if not isinstance(saved_cfg, dict):
                    saved_cfg = {}
            except Exception:
                saved_cfg = {}

        api_key = saved_cfg.get("api_key") or env_vars.get("CKEY_API_KEY", "")
        base_url = saved_cfg.get("base_url") or env_vars.get("CKEY_BASE_URL", DEFAULT_BASE_URL)
        model_name = saved_cfg.get("model_name") or env_vars.get("GEMINI_MODEL", DEFAULT_MODEL_NAME)
        custom_prompt = saved_cfg.get("custom_prompt", "")
        custom_patch_prompt = saved_cfg.get("custom_patch_prompt", "")

        return {
            "api_key": str(api_key).strip(),
            "base_url": str(base_url).strip() or DEFAULT_BASE_URL,
            "model_name": str(model_name).strip() or DEFAULT_MODEL_NAME,
            "custom_prompt": str(custom_prompt).strip(),
            "custom_patch_prompt": str(custom_patch_prompt).strip(),
        }

    def save_config(
        self,
        api_key: str,
        base_url: str = DEFAULT_BASE_URL,
        model_name: str = DEFAULT_MODEL_NAME,
        custom_prompt: str | None = None,
        custom_patch_prompt: str | None = None,
    ) -> None:
        """Lưu cấu hình OCR vào tệp JSON bền vững và đồng bộ tệp .env cục bộ."""
        self.config_path.parent.mkdir(parents=True, exist_ok=True)
        existing_custom_prompt = ""
        existing_custom_patch_prompt = ""
        if (custom_prompt is None or custom_patch_prompt is None) and self.config_path.is_file():
            try:
                old_data = json.loads(self.config_path.read_text(encoding="utf-8"))
                if isinstance(old_data, dict):
                    existing_custom_prompt = str(old_data.get("custom_prompt", "")).strip()
                    existing_custom_patch_prompt = str(old_data.get("custom_patch_prompt", "")).strip()
            except Exception:
                pass
        payload = {
            "api_key": api_key.strip(),
            "base_url": base_url.strip() or DEFAULT_BASE_URL,
            "model_name": model_name.strip() or DEFAULT_MODEL_NAME,
            "custom_prompt": existing_custom_prompt if custom_prompt is None else custom_prompt.strip(),
            "custom_patch_prompt": existing_custom_patch_prompt if custom_patch_prompt is None else custom_patch_prompt.strip(),
        }
        self.config_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        # Cập nhật tệp .env trong thư mục làm việc nếu có thể ghi
        try:
            local_env = Path.cwd() / ".env"
            env_lines = [
                f"CKEY_API_KEY={payload['api_key']}",
                f"CKEY_BASE_URL={payload['base_url']}",
                f"GEMINI_MODEL={payload['model_name']}",
            ]
            local_env.write_text("\n".join(env_lines) + "\n", encoding="utf-8")
        except Exception:
            pass

    def get_default_prompt(self) -> str:
        """Đọc prompt mặc định từ data/default_prompt.txt hoặc trả về hằng số mặc định."""
        try:
            p = find_data_file("default_prompt.txt")
            if p.is_file():
                text = p.read_text(encoding="utf-8").strip()
                if text:
                    return text
        except Exception:
            pass
        return DEFAULT_PROMPT_TEXT

    def resolve_prompt_for_folder(self, folder_path: Path | str | None = None) -> tuple[str, str]:
        """
        Tìm kiếm prompt theo 3 cấp độ giống hệt dự án `so-hoa-vbcc`:
        1. `local_prompt`: `<folder>/prompt.txt` (Prompt ưu tiên của TRƯỜNG)
        2. `year_prompt`: `<folder.parent>/prompt.txt` (Prompt chung của NĂM HỌC)
        3. `default_prompt`: `data/default_prompt.txt` (Prompt MẶC ĐỊNH của hệ thống)
        Trả về: (prompt_text, source_description)
        """
        if folder_path:
            data_dir = Path(folder_path)
            if data_dir.is_file():
                data_dir = data_dir.parent

            local_prompt = data_dir / "prompt.txt"
            year_prompt = data_dir.parent / "prompt.txt"

            if local_prompt.is_file():
                try:
                    txt = local_prompt.read_text(encoding="utf-8").strip()
                    if txt:
                        return txt, f"Prompt riêng của TRƯỜNG ({local_prompt})"
                except Exception:
                    pass

            if year_prompt.is_file():
                try:
                    txt = year_prompt.read_text(encoding="utf-8").strip()
                    if txt:
                        return txt, f"Prompt chung của NĂM HỌC ({year_prompt})"
                except Exception:
                    pass

        cfg = self.load_config()
        if cfg.get("custom_prompt"):
            return cfg["custom_prompt"], "Prompt tùy chỉnh đã lưu trên giao diện"

        default_txt = self.get_default_prompt()
        return default_txt, "Prompt MẶC ĐỊNH của hệ thống (data/default_prompt.txt)"


def _open_url_with_ssl(req: urllib.request.Request, timeout: int = 180):
    """Mở kết nối HTTPS sử dụng CA bundle từ certifi kèm cơ chế fallback SSL trên macOS."""
    ssl_context = None
    try:
        import certifi
        ssl_context = ssl.create_default_context(cafile=certifi.where())
    except Exception:
        try:
            ssl_context = ssl.create_default_context()
        except Exception:
            pass

    try:
        if ssl_context:
            return urllib.request.urlopen(req, timeout=timeout, context=ssl_context)
        return urllib.request.urlopen(req, timeout=timeout)
    except Exception as exc:
        err_str = str(exc)
        if "CERTIFICATE_VERIFY_FAILED" in err_str or "certificate verify failed" in err_str:
            unverified_ctx = getattr(ssl, "_create_unverified_context", None)
            if unverified_ctx:
                return urllib.request.urlopen(req, timeout=timeout, context=unverified_ctx())
        raise


def extract_table_from_image(
    image_path: str | Path,
    prompt_text: str,
    api_key: str,
    base_url: str = DEFAULT_BASE_URL,
    model_name: str = DEFAULT_MODEL_NAME,
    max_retries: int = 3,
    timeout: int = 180,
    retry_delay: float = 5.0,
    log_callback: Callable[[str, str], None] | None = None,
) -> str:
    """
    Gửi ảnh qua OpenAI-compatible Chat Completions API kèm cơ chế tự động thử lại khi lỗi mạng.
    Không phụ thuộc bắt buộc vào thư viện `openai` bên ngoài (sử dụng chuẩn REST API `/chat/completions`).
    """
    if not api_key or not api_key.strip():
        raise ValueError("Chưa thiết lập API Key (CKEY_API_KEY). Vui lòng nhập API Key trên giao diện.")

    base64_image = encode_image(image_path)
    endpoint = f"{base_url.rstrip('/')}/chat/completions"

    payload = {
        "model": model_name.strip() or DEFAULT_MODEL_NAME,
        "messages": [
            {
                "role": "user",
                "content": [
                    {"type": "text", "text": prompt_text},
                    {
                        "type": "image_url",
                        "image_url": {"url": f"data:image/jpeg;base64,{base64_image}"},
                    },
                ],
            }
        ],
        "temperature": 0.1,
    }
    body_bytes = json.dumps(payload).encode("utf-8")

    for attempt in range(max_retries):
        try:
            req = urllib.request.Request(
                endpoint,
                data=body_bytes,
                headers={
                    "Content-Type": "application/json",
                    "Authorization": f"Bearer {api_key.strip()}",
                    "User-Agent": "SohoaTHCS-OCR/3.0",
                },
                method="POST",
            )
            with _open_url_with_ssl(req, timeout=timeout) as resp:
                resp_text = resp.read().decode("utf-8", errors="replace")
                resp_data = json.loads(resp_text)

            choices = resp_data.get("choices", [])
            if choices and isinstance(choices, list):
                msg = choices[0].get("message", {})
                content = msg.get("content")
                return content if content is not None else ""
            return ""

        except urllib.error.HTTPError as he:
            err_detail = ""
            try:
                err_detail = he.read().decode("utf-8", errors="replace")[:300]
            except Exception:
                pass
            err_msg = f"HTTP {he.code}: {he.reason} {err_detail}".strip()
            if log_callback:
                log_callback(f"  ⚠️ Lỗi API (Lần thử {attempt + 1}/{max_retries}): {err_msg}", "warning")
            if attempt == max_retries - 1:
                return ""
            if retry_delay > 0:
                time.sleep(retry_delay)

        except Exception as e:
            if log_callback:
                log_callback(f"  ⚠️ Lỗi kết nối API (Lần thử {attempt + 1}/{max_retries}): {e}", "warning")
            if attempt == max_retries - 1:
                return ""
            if retry_delay > 0:
                time.sleep(retry_delay)

    return ""


def parse_ocr_json_response(result_text: str) -> tuple[list[dict[str, Any]] | None, str | None]:
    """
    Trích xuất và phân tích mảng JSON từ phản hồi của AI.
    Trả về: (list_of_records, error_message)
    """
    if not result_text or not str(result_text).strip():
        return None, "BỎ QUA DO LỖI MẠNG HOẶC PHẢN HỒI RỖNG"

    text = str(result_text).strip()
    match = re.search(r"\[.*\]", text, re.DOTALL)
    if not match:
        return None, "AI không trả về mảng JSON hợp lệ"

    json_str = match.group(0)
    try:
        page_data = json.loads(json_str)
        if isinstance(page_data, list):
            valid_items = [item for item in page_data if isinstance(item, dict)]
            return valid_items, None
        return None, "Dữ liệu trả về không phải là danh sách JSON"
    except json.JSONDecodeError as e:
        return None, f"Lỗi cú pháp JSON ({e})"


class OCRProcessor:
    """
    Bộ xử lý OCR ảnh sổ văn bằng sang Excel:
    - Quét và sắp xếp ảnh tự nhiên (1, 2, ..., 9, 10).
    - Xuất ra file Excel thông thường theo cấu trúc cột OCR (`stt`, `sbd`, `ho_ten`, ..., `can_kiem_tra`, `ly_do_nghi_ngo`).
    """

    FILL_HEADER = PatternFill(start_color="E0EBF5", end_color="E0EBF5", fill_type="solid")
    FONT_HEADER = Font(name="Times New Roman", size=11, bold=True, color="003366")
    FONT_DATA = Font(name="Times New Roman", size=11)
    FILL_WARN = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    FONT_WARN = Font(name="Times New Roman", size=11, bold=True, color="9C6500")
    BORDER_THIN = Border(
        left=Side(style="thin", color="D9D9D9"),
        right=Side(style="thin", color="D9D9D9"),
        top=Side(style="thin", color="D9D9D9"),
        bottom=Side(style="thin", color="D9D9D9"),
    )

    def __init__(self, config_mgr: OCRConfigManager | None = None):
        self.config_mgr = config_mgr or OCRConfigManager()

    @staticmethod
    def scan_images(paths: list[Path | str]) -> list[dict[str, Any]]:
        """
        Quét danh sách các tệp ảnh hoặc thư mục chứa ảnh (.jpg, .jpeg, .png, .heic, .heif)
        và sắp xếp theo thứ tự tự nhiên (natural_sort_key).
        """
        found_files: list[Path] = []
        for p_item in paths:
            p = Path(p_item)
            if not p.exists():
                continue
            if p.is_file():
                if p.suffix.lower() in VALID_IMAGE_EXTENSIONS and not p.name.startswith("._"):
                    found_files.append(p)
            elif p.is_dir():
                dir_imgs = [
                    f for f in p.iterdir()
                    if f.is_file() and f.suffix.lower() in VALID_IMAGE_EXTENSIONS and not f.name.startswith("._")
                ]
                dir_imgs.sort(key=natural_sort_key)
                found_files.extend(dir_imgs)

        # Khử trùng lặp giữ nguyên thứ tự
        unique_files: list[Path] = []
        seen: set[str] = set()
        for f in found_files:
            res_str = str(f.resolve())
            if res_str not in seen:
                seen.add(res_str)
                unique_files.append(f)

        unique_files.sort(key=natural_sort_key)

        results: list[dict[str, Any]] = []
        for f in unique_files:
            try:
                sz = f.stat().st_size
                sz_str = f"{sz / 1024:.1f} KB" if sz < 1024 * 1024 else f"{sz / (1024 * 1024):.1f} MB"
            except Exception:
                sz = 0
                sz_str = "---"
            results.append({
                "path": f,
                "name": f.name,
                "ext": f.suffix.lower(),
                "size": sz,
                "size_str": sz_str,
                "status": "Chờ OCR",
                "students_count": 0,
            })
        return results

    def export_to_original_excel(
        self,
        records: list[dict[str, Any]],
        output_path: Path | str,
    ) -> Path:
        """
        Xuất danh sách kết quả OCR ra file Excel với các cột động 100% dựa theo các key JSON
        mà AI trả về từ Prompt (tương đương `pd.DataFrame(records).to_excel`),
        đồng thời tô màu vàng các dòng có `can_kiem_tra = true` (nếu có trường này).
        """
        out_p = Path(output_path)
        out_p.parent.mkdir(parents=True, exist_ok=True)

        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "KetQua_OCR"

        # Thu thập danh sách cột động theo đúng thứ tự xuất hiện trong dữ liệu trả về từ Prompt
        dynamic_keys: list[str] = []
        for rec in records:
            if isinstance(rec, dict):
                for k in rec.keys():
                    if k not in dynamic_keys:
                        dynamic_keys.append(k)

        # Fallback nếu records rỗng
        if not dynamic_keys:
            dynamic_keys = [col[0] for col in ORIGINAL_OCR_COLUMNS]

        known_widths = {col[0]: col[2] for col in ORIGINAL_OCR_COLUMNS}

        # Dòng 1: Header
        for c_idx, key in enumerate(dynamic_keys, start=1):
            cell = ws.cell(row=1, column=c_idx, value=key)
            cell.fill = self.FILL_HEADER
            cell.font = self.FONT_HEADER
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
            cell.border = self.BORDER_THIN
            width = known_widths.get(key, max(14, min(36, len(str(key)) + 6)))
            ws.column_dimensions[get_column_letter(c_idx)].width = width

        # Dữ liệu từng học sinh
        for r_idx, rec in enumerate(records, start=2):
            need_check = bool(rec.get("can_kiem_tra")) and str(rec.get("can_kiem_tra")).strip().lower() not in ("false", "0", "none", "null", "")
            for c_idx, key in enumerate(dynamic_keys, start=1):
                val = rec.get(key)
                if val is None:
                    val = ""
                cell = ws.cell(row=r_idx, column=c_idx, value=val)
                cell.border = self.BORDER_THIN
                if need_check:
                    cell.fill = self.FILL_WARN
                    cell.font = self.FONT_WARN if key in ("can_kiem_tra", "ly_do_nghi_ngo", "ho_ten") else self.FONT_DATA
                else:
                    cell.font = self.FONT_DATA

        wb.save(out_p)
        wb.close()
        return out_p

    def process_ocr_batch(
        self,
        image_paths: list[Path | str],
        output_excel_path: Path | str,
        prompt_text: str,
        api_key: str,
        base_url: str = DEFAULT_BASE_URL,
        model_name: str = DEFAULT_MODEL_NAME,
        max_retries: int = 3,
        retry_delay: float = 5.0,
        progress_callback: Callable[[int, int, str], None] | None = None,
        file_status_callback: Callable[[Path, str, int], None] | None = None,
        log_callback: Callable[[str, str], None] | None = None,
        cancel_token: Callable[[], bool] | None = None,
    ) -> dict[str, Any]:
        """
        Thực thi OCR hàng loạt trên danh sách ảnh và lưu ra tệp Excel.
        """
        def log(msg: str, level: str = "info"):
            if log_callback:
                log_callback(msg, level)

        sorted_paths = sorted([Path(p) for p in image_paths], key=natural_sort_key)
        total_images = len(sorted_paths)
        if total_images == 0:
            return {"status": "error", "message": "Không có tệp ảnh hợp lệ để xử lý OCR."}

        all_students_data: list[dict[str, Any]] = []
        success_images = 0
        failed_images = 0
        suspect_rows_count = 0

        log(f"🚀 BẮT ĐẦU CHẠY OCR TRÊN {total_images} ẢNH (Model: {model_name})...", "info")

        for idx, img_p in enumerate(sorted_paths, start=1):
            if cancel_token and cancel_token():
                log(f"🛑 Đã dừng tiến trình OCR theo yêu cầu tại ảnh {img_p.name}.", "warning")
                break

            if progress_callback:
                prog_pct = int(((idx - 1) / total_images) * 100)
                progress_callback(prog_pct, 100, f"[{idx}/{total_images}] Đang đọc OCR: {img_p.name}")

            if file_status_callback:
                file_status_callback(img_p, "⏳ Đang nhận diện...", 0)

            log(f"📸 [{idx}/{total_images}] Đang xử lý ảnh: {img_p.name}...", "info")

            try:
                raw_response = extract_table_from_image(
                    image_path=img_p,
                    prompt_text=prompt_text,
                    api_key=api_key,
                    base_url=base_url,
                    model_name=model_name,
                    max_retries=max_retries,
                    retry_delay=retry_delay,
                    log_callback=log_callback,
                )

                page_records, err_msg = parse_ocr_json_response(raw_response)
                if page_records is not None:
                    count = len(page_records)
                    all_students_data.extend(page_records)
                    success_images += 1
                    page_suspects = sum(
                        1 for r in page_records
                        if bool(r.get("can_kiem_tra")) and str(r.get("can_kiem_tra")).strip().lower() not in ("false", "0", "none", "null", "")
                    )
                    suspect_rows_count += page_suspects
                    status_str = f"✅ OK ({count} HS)" + (f" [⚠️ {page_suspects} cần KT]" if page_suspects else "")
                    if file_status_callback:
                        file_status_callback(img_p, status_str, count)
                    log(f"  ✅ {img_p.name}: Trích xuất thành công {count} học sinh" + (f" ({page_suspects} dòng cần kiểm tra lại)." if page_suspects else "."), "success")
                else:
                    failed_images += 1
                    if file_status_callback:
                        file_status_callback(img_p, f"❌ {err_msg}", 0)
                    log(f"  ❌ {img_p.name}: {err_msg}", "error")
                    if raw_response:
                        log(f"     [Phản hồi thô]: {raw_response[:200]}...", "warning")

            except Exception as e:
                failed_images += 1
                if file_status_callback:
                    file_status_callback(img_p, f"❌ Lỗi: {e}", 0)
                log(f"  ❌ Lỗi hệ thống khi xử lý {img_p.name}: {e}", "error")

            if progress_callback:
                prog_pct = int((idx / total_images) * 100)
                progress_callback(prog_pct, 100, f"Đã xử lý {idx}/{total_images} ảnh ({len(all_students_data)} học sinh)")

        if not all_students_data:
            return {
                "status": "empty",
                "message": "Không trích xuất được dữ liệu học sinh hợp lệ nào từ các ảnh đã chọn.",
                "total_images": total_images,
                "success_images": success_images,
                "failed_images": failed_images,
                "total_students": 0,
            }

        excel_out = Path(output_excel_path)
        excel_out.parent.mkdir(parents=True, exist_ok=True)

        # Xuất file Excel kết quả OCR
        self.export_to_original_excel(all_students_data, excel_out)

        log(
            f"🎉 HOÀN TẤT OCR: Đã trích xuất {len(all_students_data):,} học sinh từ {success_images}/{total_images} ảnh.",
            "success",
        )
        log(f"  📄 Tệp Excel: {excel_out}", "info")

        return {
            "status": "success",
            "output_file": excel_out,
            "total_images": total_images,
            "success_images": success_images,
            "failed_images": failed_images,
            "total_students": len(all_students_data),
            "suspect_rows_count": suspect_rows_count,
            "records": all_students_data,
        }
