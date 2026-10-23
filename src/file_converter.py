from __future__ import annotations

import os
import sys
import subprocess
import shutil
import io
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
from typing import Callable, Any
import base64

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False

try:
    import win32com.client
    HAS_WIN32COM = True
except ImportError:
    HAS_WIN32COM = False

from src.normalizer import (
    auto_convert_legacy_encoding,
    is_tcvn3,
    is_vni
)


class FileFormatConverter:
    """
    Bộ xử lý chuyển đổi định dạng tệp cũ sang định dạng mới:
    - .doc  -> .docx (Word 97-2003 -> Word 2007+)
    - .xls  -> .xlsx (Excel 97-2003 -> Excel OpenXML)
    Hỗ trợ đa nền tảng (macOS và Windows).
    """
    _excel_app = None
    _word_app = None

    @classmethod
    def get_excel_app(cls):
        """Lấy hoặc khởi tạo phiên Microsoft Excel qua win32com (tái sử dụng instance)."""
        if cls._excel_app is None and HAS_WIN32COM:
            try:
                cls._excel_app = win32com.client.Dispatch("Excel.Application")
                cls._excel_app.Visible = False
                cls._excel_app.DisplayAlerts = False
                cls._excel_app.ScreenUpdating = False
                cls._excel_app.EnableEvents = False
            except Exception:
                cls._excel_app = None
        return cls._excel_app

    @classmethod
    def get_word_app(cls):
        """Lấy hoặc khởi tạo phiên Microsoft Word qua win32com (tái sử dụng instance)."""
        if cls._word_app is None and HAS_WIN32COM:
            try:
                cls._word_app = win32com.client.Dispatch("Word.Application")
                cls._word_app.Visible = False
                cls._word_app.DisplayAlerts = 0
                cls._word_app.ScreenUpdating = False
            except Exception:
                cls._word_app = None
        return cls._word_app

    @staticmethod
    def _run_powershell(script: str, timeout: int = 120) -> tuple[int, str, str]:
        """
        Thực thi script PowerShell ẩn dưới nền, hỗ trợ 100% đường dẫn Unicode tiếng Việt
        thông qua chuẩn mã hóa UTF-16LE Base64 (-EncodedCommand).
        """
        full_script = f'''
        $OutputEncoding = [System.Text.Encoding]::UTF8
        [Console]::OutputEncoding = [System.Text.Encoding]::UTF8
        {script}
        '''
        enc = base64.b64encode(full_script.encode('utf-16le')).decode('ascii')
        cmd = ["powershell", "-NoProfile", "-NonInteractive", "-WindowStyle", "Hidden", "-EncodedCommand", enc]
        res = subprocess.run(
            cmd,
            capture_output=True,
            timeout=timeout,
            creationflags=subprocess.CREATE_NO_WINDOW if sys.platform == "win32" else 0
        )
        stdout = res.stdout.decode('utf-8', errors='replace').strip()
        stderr = res.stderr.decode('utf-8', errors='replace').strip()
        return res.returncode, stdout, stderr

    @staticmethod
    def scan_files(paths: list[Path | str], recursive: bool = True) -> list[dict]:
        """
        Quét danh sách các tệp hoặc thư mục để tìm các tệp .doc và .xls:
        - Bỏ qua các tệp ẩn, tệp tạm (~$..., ._...)
        - Trả về danh sách thông tin chi tiết từng tệp.
        """
        found_files: list[Path] = []
        for p_item in paths:
            p = Path(p_item)
            if not p.exists():
                continue
            if p.is_file():
                ext = p.suffix.lower()
                if ext in (".doc", ".xls") and not p.name.startswith("~$") and not p.name.startswith("._"):
                    found_files.append(p)
            elif p.is_dir():
                pattern = "**/*" if recursive else "*"
                for f in p.glob(pattern):
                    if f.is_file():
                        ext = f.suffix.lower()
                        if ext in (".doc", ".xls") and not f.name.startswith("~$") and not f.name.startswith("._"):
                            found_files.append(f)

        # Loại bỏ trùng lặp
        unique_files: list[Path] = []
        seen = set()
        for f in found_files:
            resolved = str(f.resolve())
            if resolved not in seen:
                seen.add(resolved)
                unique_files.append(f)

        results = []
        for f in unique_files:
            try:
                sz = f.stat().st_size
                sz_str = f"{sz / 1024:.1f} KB" if sz < 1024 * 1024 else f"{sz / (1024 * 1024):.1f} MB"
            except Exception:
                sz = 0
                sz_str = "---"

            ext = f.suffix.lower()
            target_ext = ".docx" if ext == ".doc" else ".xlsx"
            type_label = "Word (.doc)" if ext == ".doc" else "Excel (.xls)"

            results.append({
                "path": f,
                "name": f.name,
                "ext": ext,
                "target_ext": target_ext,
                "type_label": type_label,
                "size": sz,
                "size_str": sz_str,
                "status": "Chờ chuyển đổi"
            })

        return results

    @staticmethod
    def detect_file_type(file_path: Path | str) -> str:
        """
        Nhận diện định dạng thực tế của tệp thông qua Magic Bytes / Header:
        - 'doc_binary' / 'xls_binary': OLE2 Compound File (Word/Excel 97-2003: D0 CF 11 E0 A1 B1 1A E1)
        - 'docx_zip': Gói nén PK ZIP chứa word/
        - 'xlsx_zip': Gói nén PK ZIP chứa xl/
        - 'rtf': Rich Text Format ({\\rtf)
        - 'html': HTML Table/Document (<html, <!DOCTYPE, <table)
        - 'unknown': Không xác định
        """
        p = Path(file_path)
        if not p.exists() or not p.is_file():
            return "unknown"
        try:
            with open(p, "rb") as f:
                header = f.read(512)
            if header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
                return "xls_binary" if p.suffix.lower() == ".xls" else "doc_binary"
            if header.startswith(b"PK\x03\x04"):
                try:
                    with zipfile.ZipFile(p, "r") as z:
                        namelist = z.namelist()
                        if any(n.startswith("xl/") for n in namelist):
                            return "xlsx_zip"
                        if any(n.startswith("word/") for n in namelist):
                            return "docx_zip"
                except Exception:
                    pass
                return "xlsx_zip" if p.suffix.lower() == ".xls" else "docx_zip"
            if header.strip().startswith(b"{\\rtf"):
                return "rtf"
            header_str = header.decode("latin1", errors="ignore").lower().strip()
            if "<html" in header_str or "<!doctype html" in header_str or "<table" in header_str:
                return "html"
            return "unknown"
        except Exception:
            return "unknown"

    @staticmethod
    def find_libreoffice() -> str | None:
        """Tìm đường dẫn thực thi của LibreOffice / soffice trên cả macOS và Windows."""
        candidates = [
            shutil.which("soffice"),
            shutil.which("libreoffice"),
        ]
        if sys.platform == "darwin":
            candidates.extend([
                "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                "/Applications/OpenOffice.app/Contents/MacOS/soffice",
                os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice"),
                "/opt/homebrew/bin/soffice",
                "/usr/local/bin/soffice",
            ])
        elif sys.platform == "win32":
            candidates.extend([
                os.path.expandvars(r"%ProgramFiles%\LibreOffice\program\soffice.exe"),
                os.path.expandvars(r"%ProgramFiles(x86)%\LibreOffice\program\soffice.exe"),
                os.path.expandvars(r"%LOCALAPPDATA%\Programs\LibreOffice\program\soffice.exe"),
            ])
        for c in candidates:
            if c and os.path.isfile(c) and os.access(c, os.X_OK):
                return c
        return None

    @staticmethod
    def validate_docx_integrity(docx_path: Path | str) -> tuple[bool, str]:
        """
        Kiểm tra tính toàn vẹn của tệp .docx sau khi chuyển đổi:
        1. Tệp tồn tại và có dung lượng hợp lệ (> 100 bytes).
        2. Cấu trúc nén ZIP hợp lệ.
        3. Chứa các thành phần cốt lõi: [Content_Types].xml và word/document.xml.
        4. Cú pháp XML trong word/document.xml hợp lệ, không bị lỗi cú pháp.
        """
        p = Path(docx_path)
        if not p.exists():
            return False, "Tệp đích không tồn tại."
        if p.stat().st_size < 100:
            return False, f"Tệp đích rỗng hoặc quá nhỏ ({p.stat().st_size} bytes)."
        if not zipfile.is_zipfile(p):
            return False, "Tệp đích không phải là tệp nén OpenXML (ZIP) hợp lệ."

        try:
            with zipfile.ZipFile(p, "r") as z:
                namelist = z.namelist()
                if "[Content_Types].xml" not in namelist:
                    return False, "Thiếu [Content_Types].xml trong gói OOXML."
                if "word/document.xml" not in namelist:
                    return False, "Thiếu word/document.xml trong gói OOXML."

                # Kiểm tra parse XML document.xml
                doc_xml = z.read("word/document.xml")
                ET.fromstring(doc_xml)
            return True, "Hợp lệ"
        except Exception as e:
            return False, f"Lỗi cấu trúc OpenXML: {e}"

    @staticmethod
    def validate_xlsx_integrity(xlsx_path: Path | str) -> tuple[bool, str]:
        """
        Kiểm tra tính toàn vẹn của tệp .xlsx sau khi chuyển đổi:
        1. Tệp tồn tại và có dung lượng hợp lệ (> 100 bytes).
        2. Cấu trúc nén ZIP hợp lệ.
        3. Chứa các thành phần cốt lõi: [Content_Types].xml và xl/workbook.xml.
        4. Cú pháp XML trong xl/workbook.xml hợp lệ.
        """
        p = Path(xlsx_path)
        if not p.exists():
            return False, "Tệp đích không tồn tại."
        if p.stat().st_size < 100:
            return False, f"Tệp đích rỗng hoặc quá nhỏ ({p.stat().st_size} bytes)."
        if not zipfile.is_zipfile(p):
            return False, "Tệp đích không phải là tệp nén OpenXML (ZIP) hợp lệ."

        try:
            with zipfile.ZipFile(p, "r") as z:
                namelist = z.namelist()
                if "[Content_Types].xml" not in namelist:
                    return False, "Thiếu [Content_Types].xml trong gói OOXML."
                if "xl/workbook.xml" not in namelist:
                    return False, "Thiếu xl/workbook.xml trong gói OOXML."

                wb_xml = z.read("xl/workbook.xml")
                ET.fromstring(wb_xml)
            return True, "Hợp lệ"
        except Exception as e:
            return False, f"Lỗi cấu trúc OpenXML Excel: {e}"

    @staticmethod
    def sanitize_and_repair_docx(docx_path: Path | str) -> bool:
        """
        Sửa chữa và chuẩn hóa các lỗi OpenXML thường gặp từ các công cụ cũ/textutil:
        - Thay thế thẻ không hợp lệ/lạc hậu <w:sz-cs> thành <w:szCs> (nguyên nhân gây unreadable content trong MS Word)
        - Loại bỏ các file rác phi tiêu chuẩn (docProps/meta.xml của Cocoa)
        - Đảm bảo gói ZIP hoàn chỉnh và có thể mở được trên mọi bản MS Word.
        """
        p = Path(docx_path)
        if not p.exists() or not zipfile.is_zipfile(p):
            return False

        try:
            repaired_buffer = io.BytesIO()
            with zipfile.ZipFile(p, "r") as zin:
                with zipfile.ZipFile(repaired_buffer, "w", zipfile.ZIP_DEFLATED) as zout:
                    for item in zin.infolist():
                        # Bỏ qua file rác phi tiêu chuẩn gây lỗi xác thực trong Word
                        if item.filename == "docProps/meta.xml":
                            continue
                        data = zin.read(item.filename)
                        if item.filename == "word/document.xml":
                            text = data.decode("utf-8", errors="ignore")
                            # Sửa thẻ lỗi w:sz-cs -> w:szCs
                            text = re.sub(r"<w:sz-cs(\b|/)", r"<w:szCs\1", text)
                            text = re.sub(r"</w:sz-cs>", r"</w:szCs>", text)
                            data = text.encode("utf-8")
                        zout.writestr(item, data)

            # Ghi đè lại file gốc
            repaired_buffer.seek(0)
            with open(p, "wb") as f:
                f.write(repaired_buffer.read())
            return True
        except Exception:
            return False

    @classmethod
    def close_office_apps(cls, apps: tuple[str, ...] = ("word", "excel")) -> None:
        """
        Đóng an toàn và dọn sạch các tệp/tiến trình Microsoft Word và Microsoft Excel.
        - Trên macOS: Sử dụng AppleScript đóng các file đã mở và thoát ứng dụng.
        - Trên Windows: Thu dọn COM (win32com & PowerShell) và thu hồi rác bộ nhớ.
        """
        if sys.platform == "darwin":
            script_parts = []
            if "word" in apps:
                script_parts.append('''
                if application "Microsoft Word" is running then
                    tell application "Microsoft Word"
                        set origAlerts to display alerts
                        set display alerts to none
                        close every document saving no
                        set display alerts to origAlerts
                        quit saving no
                    end tell
                end if
                ''')
            if "excel" in apps:
                script_parts.append('''
                if application "Microsoft Excel" is running then
                    tell application "Microsoft Excel"
                        set origAlerts to display alerts
                        set display alerts to false
                        close every workbook saving no
                        set display alerts to origAlerts
                        quit
                    end tell
                end if
                ''')
            if script_parts:
                full_script = "\n".join(script_parts)
                try:
                    subprocess.run(["osascript", "-e", full_script], capture_output=True, timeout=10)
                except Exception:
                    pass
        elif sys.platform == "win32":
            if "excel" in apps and cls._excel_app is not None:
                try:
                    cls._excel_app.Quit()
                except Exception:
                    pass
                cls._excel_app = None
            if "word" in apps and cls._word_app is not None:
                try:
                    cls._word_app.Quit()
                except Exception:
                    pass
                cls._word_app = None
            import gc
            gc.collect()

    @staticmethod
    def _convert_doc_via_word_macos(src_p: Path, out_p: Path, auto_quit: bool = True) -> tuple[bool, str]:
        """
        Chuyển đổi .doc sang .docx thông qua Microsoft Word trên macOS bằng AppleScript:
        - Tự động mở file .doc trong Microsoft Word
        - Lưu sang .docx chuẩn native bằng engine Word (bảo toàn 100% định dạng, font chữ, bảng biểu, layout)
        - Đóng file .doc ngay sau khi lưu để không bị kẹt cửa sổ
        - Chờ người dùng cấp phép truy cập thư mục nếu macOS hiển thị hộp thoại
        """
        if not Path("/Applications/Microsoft Word.app").exists():
            return False, "Microsoft Word không được cài đặt trong /Applications."

        src_abs = str(src_p.resolve())
        out_abs = str(out_p.resolve())

        # Nếu tệp đích đã tồn tại, xóa trước để Word không hiển thị hộp thoại hỏi ghi đè
        if out_p.exists():
            try:
                out_p.unlink()
            except Exception:
                pass

        quit_clause = "if (count of documents) is 0 then quit saving no" if auto_quit else ""

        as_script = f'''
        tell application "Microsoft Word"
            set origAlerts to display alerts
            set display alerts to none
            try
                open file name "{src_abs}"
                set theDoc to active document
                set hfsOut to (POSIX file "{out_abs}") as text
                save as theDoc file name hfsOut file format format document default
                close active document saving no
                try
                    set display alerts to origAlerts
                end try
                {quit_clause}
                return "OK"
            on error errMsg
                try
                    close active document saving no
                end try
                try
                    set display alerts to origAlerts
                end try
                {quit_clause}
                return "ERR:" & errMsg
            end try
        end tell
        '''
        try:
            res = subprocess.run(
                ["osascript", "-e", as_script],
                capture_output=True,
                text=True,
                timeout=120
            )
            out_str = (res.stdout or "").strip()
            if "OK" in out_str and out_p.exists() and out_p.stat().st_size > 0:
                return True, "Thành công qua Microsoft Word (macOS)"
            return False, res.stderr.strip() or out_str
        except subprocess.TimeoutExpired:
            return False, "Quá thời gian chờ (120s) khi chuyển đổi bằng Microsoft Word."
        except Exception as e:
            return False, str(e)

    @classmethod
    def _convert_doc_via_word_windows(cls, src_p: Path, out_p: Path, auto_quit: bool = True) -> tuple[bool, str]:
        """
        Chuyển đổi .doc sang .docx thông qua Microsoft Word trên Windows:
        - Ưu tiên win32com native (in-process, tốc độ cực nhanh ~0.2s - 0.5s/file).
        - Dự phòng qua PowerShell COM nếu win32com không khả dụng.
        """
        src_abs = str(src_p.resolve())
        out_abs = str(out_p.resolve())

        # Xóa tệp đích trước để không bị hỏi ghi đè
        if out_p.exists():
            try:
                out_p.unlink()
            except Exception:
                pass

        # 1. Ưu tiên win32com native (In-process, không mở process phụ)
        if HAS_WIN32COM:
            try:
                word = cls.get_word_app()
                if word is None:
                    word = win32com.client.Dispatch("Word.Application")
                    word.Visible = False
                    word.DisplayAlerts = 0
                    word.ScreenUpdating = False

                # ConfirmConversions=False, ReadOnly=True, AddToRecentFiles=False
                doc = word.Documents.Open(src_abs, False, True, False)
                # wdFormatXMLDocument = 16
                doc.SaveAs2(out_abs, 16)
                doc.Close(0)

                if auto_quit:
                    cls.close_office_apps(apps=("word",))

                if out_p.exists() and out_p.stat().st_size > 0:
                    return True, "Thành công qua Microsoft Word (Native COM)"
            except Exception:
                if auto_quit:
                    cls.close_office_apps(apps=("word",))

        # 2. Phương án dự phòng qua PowerShell COM
        src_abs_ps = src_abs.replace("'", "''")
        out_abs_ps = out_abs.replace("'", "''")
        quit_clause = """
            if ($word -ne $null) {
                try { $word.Quit() } catch {}
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null } catch {}
            }
        """ if auto_quit else ""

        ps_script = f'''
        $ErrorActionPreference = 'Stop'
        $ProgressPreference = 'SilentlyContinue'
        $word = $null
        $doc = $null
        try {{
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            $word.ScreenUpdating = $false
            try {{ $word.AutomationSecurity = 3 }} catch {{}}

            $doc = $word.Documents.Open('{src_abs_ps}', $false, $true, $false)
            $doc.SaveAs2('{out_abs_ps}', 16)
            $doc.Close(0)
            $doc = $null
            Write-Output "OK"
        }} catch {{
            Write-Error $_.Exception.Message
        }} finally {{
            if ($doc -ne $null) {{
                try {{ $doc.Close(0) }} catch {{}}
            }}
            {quit_clause}
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }}
        '''
        try:
            code, stdout, stderr = FileFormatConverter._run_powershell(ps_script, timeout=120)
            if "OK" in stdout and out_p.exists() and out_p.stat().st_size > 0:
                return True, "Thành công qua Microsoft Word (Windows)"
            return False, stderr or stdout or "Lỗi PowerShell COM Word"
        except subprocess.TimeoutExpired:
            return False, "Quá thời gian chờ (120s) khi gọi Word COM trên Windows."
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _convert_doc_via_libreoffice(soffice_path: str, src_p: Path, out_p: Path) -> tuple[bool, str]:
        """Chuyển đổi thông qua LibreOffice headless CLI."""
        try:
            extra_kwargs: dict[str, Any] = {}
            if sys.platform == "win32":
                extra_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

            cmd = [
                soffice_path,
                "--headless",
                "--convert-to", "docx:MS Word 2007 XML",
                "--outdir", str(out_p.parent),
                str(src_p.resolve())
            ]
            subprocess.run(cmd, capture_output=True, text=True, timeout=60, **extra_kwargs)
            gen_file = out_p.parent / f"{src_p.stem}.docx"
            if gen_file.exists():
                if gen_file.resolve() != out_p.resolve():
                    if out_p.exists():
                        out_p.unlink()
                    gen_file.replace(out_p)
                return True, "Thành công qua LibreOffice"

            # Thử lại nếu không dùng bộ lọc tên
            cmd_simple = [
                soffice_path,
                "--headless",
                "--convert-to", "docx",
                "--outdir", str(out_p.parent),
                str(src_p.resolve())
            ]
            subprocess.run(cmd_simple, capture_output=True, text=True, timeout=60, **extra_kwargs)
            if gen_file.exists():
                if gen_file.resolve() != out_p.resolve():
                    if out_p.exists():
                        out_p.unlink()
                    gen_file.replace(out_p)
                return True, "Thành công qua LibreOffice"

            return False, "LibreOffice không sinh tệp .docx đích."
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _convert_doc_via_textutil_macos(src_p: Path, out_p: Path) -> tuple[bool, str]:
        """Chuyển đổi bằng textutil trên macOS và tự động sửa các lỗi XML OpenXML."""
        try:
            cmd = ["textutil", "-convert", "docx", "-output", str(out_p.resolve()), str(src_p.resolve())]
            res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
            if res.returncode == 0 and out_p.exists() and out_p.stat().st_size > 0:
                FileFormatConverter.sanitize_and_repair_docx(out_p)
                return True, "Thành công qua textutil (đã tối ưu hóa OpenXML)"
            return False, res.stderr.strip() or "textutil không thể chuyển đổi tệp."
        except Exception as e:
            return False, str(e)

    @staticmethod
    def convert_doc_to_docx(
        doc_path: Path | str,
        output_path: Path | str | None = None,
        delete_source: bool = False,
        auto_quit: bool = True
    ) -> tuple[bool, Path | None, str]:
        """
        Chuyển đổi một tệp .doc (Word 97-2003) sang .docx trực tiếp bằng Microsoft Word:
        - Mở trực tiếp bằng Microsoft Word (AppleScript trên macOS, COM trên Windows).
        - Lưu sang .docx chuẩn Native để giữ nguyên 100% định dạng, font chữ, bảng biểu, layout.
        - Tự động đóng tệp .doc cũ ngay sau khi lưu để không lưu lại bất kỳ cửa sổ nào.
        - Chờ người dùng cấp phép truy cập thư mục nếu cần (hỗ trợ timeout lên đến 120s).
        - Thực hiện nghiêm ngặt qua Microsoft Word, không sử dụng các phương án dự phòng làm sai lệch định dạng.
        """
        p = Path(doc_path)
        if not p.exists() or p.suffix.lower() != ".doc":
            return False, None, f"Tệp không hợp lệ: {p.name}"

        out_p = Path(output_path) if output_path else p.with_suffix(".docx")
        out_p.parent.mkdir(parents=True, exist_ok=True)

        # 0. Nhận diện định dạng thực tế (trường hợp tệp thực chất là .docx bị đổi đuôi nhầm thành .doc)
        real_type = FileFormatConverter.detect_file_type(p)
        if real_type == "docx_zip":
            try:
                shutil.copy2(p, out_p)
                valid, msg = FileFormatConverter.validate_docx_integrity(out_p)
                if valid:
                    if delete_source:
                        try:
                            p.unlink()
                        except Exception:
                            pass
                    return True, out_p, "Thành công (Tệp gốc là DOCX đổi tên)"
                else:
                    if out_p.exists():
                        out_p.unlink()
            except Exception:
                pass

        # Chuyển đổi trực tiếp bằng Microsoft Word Native theo yêu cầu (không dùng phương án dự phòng khác)
        if sys.platform == "darwin":
            ok, msg = FileFormatConverter._convert_doc_via_word_macos(p, out_p, auto_quit=auto_quit)
            method_used = "Microsoft Word (macOS)"
        elif sys.platform == "win32":
            ok, msg = FileFormatConverter._convert_doc_via_word_windows(p, out_p, auto_quit=auto_quit)
            method_used = "Microsoft Word (Windows)"
        else:
            return False, None, "Hệ điều hành hiện tại không hỗ trợ Microsoft Word."

        if not ok:
            if out_p.exists():
                try:
                    out_p.unlink()
                except Exception:
                    pass
            return False, None, f"Lỗi chuyển đổi qua Microsoft Word: {msg}"

        # Kiểm tra tính toàn vẹn của tệp .docx sinh ra
        if out_p.exists():
            final_valid, final_val_msg = FileFormatConverter.validate_docx_integrity(out_p)
            if not final_valid:
                try:
                    out_p.unlink()
                except Exception:
                    pass
                return False, None, f"Tệp .docx sinh ra không hợp lệ: {final_val_msg}"

            if delete_source:
                try:
                    p.unlink()
                except Exception:
                    pass
            return True, out_p, f"Thành công ({method_used})"

        return False, None, "Microsoft Word không tạo được tệp .docx đích."

    @staticmethod
    def _convert_xls_via_excel_macos(src_p: Path, out_p: Path, auto_quit: bool = True) -> tuple[bool, str]:
        """
        Chuyển đổi .xls sang .xlsx thông qua Microsoft Excel trên macOS bằng AppleScript:
        - Tự động mở file .xls trong Microsoft Excel
        - Lưu sang .xlsx chuẩn native bằng engine Excel (bảo toàn 100% định dạng, viền, ô gộp, công thức, font)
        - Đóng file .xls ngay sau khi lưu để không bị kẹt cửa sổ
        - Chờ người dùng cấp phép truy cập thư mục nếu macOS hiển thị hộp thoại
        """
        if not Path("/Applications/Microsoft Excel.app").exists():
            return False, "Microsoft Excel không được cài đặt trong /Applications."

        src_abs = str(src_p.resolve())
        out_abs = str(out_p.resolve())

        # Nếu tệp đích đã tồn tại, xóa trước để Excel không hiển thị hộp thoại hỏi ghi đè
        if out_p.exists():
            try:
                out_p.unlink()
            except Exception:
                pass

        quit_clause = "if (count of workbooks) is 0 then quit" if auto_quit else ""

        as_script = f'''
        tell application "Microsoft Excel"
            set origAlerts to display alerts
            set display alerts to false
            try
                open (POSIX file "{src_abs}")
                set theWb to active workbook
                set hfsOut to (POSIX file "{out_abs}") as text
                save workbook as theWb filename hfsOut file format workbook normal file format
                close active workbook saving no
                try
                    set display alerts to origAlerts
                end try
                {quit_clause}
                return "OK"
            on error errMsg
                try
                    close active workbook saving no
                end try
                try
                    set display alerts to origAlerts
                end try
                {quit_clause}
                return "ERR:" & errMsg
            end try
        end tell
        '''
        try:
            res = subprocess.run(
                ["osascript", "-e", as_script],
                capture_output=True,
                text=True,
                timeout=120
            )
            out_str = (res.stdout or "").strip()
            if "OK" in out_str and out_p.exists() and out_p.stat().st_size > 0:
                return True, "Thành công qua Microsoft Excel (macOS)"
            return False, res.stderr.strip() or out_str
        except subprocess.TimeoutExpired:
            return False, "Quá thời gian chờ (120s) khi chuyển đổi bằng Microsoft Excel."
        except Exception as e:
            return False, str(e)

    @classmethod
    def _convert_xls_via_excel_windows(cls, src_p: Path, out_p: Path, auto_quit: bool = True) -> tuple[bool, str]:
        """
        Chuyển đổi .xls sang .xlsx thông qua Microsoft Excel trên Windows:
        - Ưu tiên win32com native (in-process, tốc độ cực nhanh ~0.2s - 0.5s/file).
        - Tự động tái sử dụng phiên Excel khi chuyển đổi nhiều tệp liên tiếp.
        - Dự phòng qua PowerShell COM (-EncodedCommand UTF-16LE) nếu win32com không khả dụng.
        """
        src_abs = str(src_p.resolve())
        out_abs = str(out_p.resolve())

        # Xóa tệp đích trước để không bị hỏi ghi đè
        if out_p.exists():
            try:
                out_p.unlink()
            except Exception:
                pass

        # 1. Ưu tiên win32com native (In-process, siêu tốc)
        if HAS_WIN32COM:
            try:
                excel = cls.get_excel_app()
                if excel is None:
                    excel = win32com.client.Dispatch("Excel.Application")
                    excel.Visible = False
                    excel.DisplayAlerts = False
                    excel.ScreenUpdating = False
                    excel.EnableEvents = False

                # UpdateLinks=0, ReadOnly=True
                wb = excel.Workbooks.Open(src_abs, 0, True)
                # xlOpenXMLWorkbook = 51
                wb.SaveAs(out_abs, 51)
                wb.Close(False)

                if auto_quit:
                    cls.close_office_apps(apps=("excel",))

                if out_p.exists() and out_p.stat().st_size > 0:
                    return True, "Thành công qua Microsoft Excel (Native COM)"
            except Exception:
                if auto_quit:
                    cls.close_office_apps(apps=("excel",))

        # 2. Phương án dự phòng qua PowerShell COM
        src_abs_ps = src_abs.replace("'", "''")
        out_abs_ps = out_abs.replace("'", "''")
        quit_clause = """
            if ($excel -ne $null) {
                try { $excel.Quit() } catch {}
                try { [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null } catch {}
            }
        """ if auto_quit else ""

        ps_script = f'''
        $ErrorActionPreference = 'Stop'
        $ProgressPreference = 'SilentlyContinue'
        $excel = $null
        $wb = $null
        try {{
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false
            $excel.AskToUpdateLinks = $false
            $excel.ScreenUpdating = $false
            $excel.EnableEvents = $false
            try {{ $excel.AutomationSecurity = 3 }} catch {{}}

            $wb = $excel.Workbooks.Open('{src_abs_ps}', 0, $true)
            $wb.SaveAs('{out_abs_ps}', 51)
            $wb.Close($false)
            $wb = $null
            Write-Output "OK"
        }} catch {{
            Write-Error $_.Exception.Message
        }} finally {{
            if ($wb -ne $null) {{
                try {{ $wb.Close($false) }} catch {{}}
            }}
            {quit_clause}
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }}
        '''
        try:
            code, stdout, stderr = FileFormatConverter._run_powershell(ps_script, timeout=120)
            if "OK" in stdout and out_p.exists() and out_p.stat().st_size > 0:
                return True, "Thành công qua Microsoft Excel (Windows)"
            return False, stderr or stdout or "Lỗi PowerShell COM Excel"
        except subprocess.TimeoutExpired:
            return False, "Quá thời gian chờ (120s) khi gọi Excel COM trên Windows."
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _convert_xls_via_libreoffice(soffice_path: str, src_p: Path, out_p: Path) -> tuple[bool, str]:
        """Chuyển đổi .xls sang .xlsx thông qua LibreOffice Calc headless CLI."""
        try:
            extra_kwargs: dict[str, Any] = {}
            if sys.platform == "win32":
                extra_kwargs["creationflags"] = subprocess.CREATE_NO_WINDOW

            cmd = [
                soffice_path,
                "--headless",
                "--convert-to", "xlsx:Calc MS Excel 2007 XML",
                "--outdir", str(out_p.parent),
                str(src_p.resolve())
            ]
            subprocess.run(cmd, capture_output=True, text=True, timeout=60, **extra_kwargs)
            gen_file = out_p.parent / f"{src_p.stem}.xlsx"
            if gen_file.exists():
                if gen_file.resolve() != out_p.resolve():
                    if out_p.exists():
                        out_p.unlink()
                    gen_file.replace(out_p)
                return True, "Thành công qua LibreOffice Calc"

            cmd_simple = [
                soffice_path,
                "--headless",
                "--convert-to", "xlsx",
                "--outdir", str(out_p.parent),
                str(src_p.resolve())
            ]
            subprocess.run(cmd_simple, capture_output=True, text=True, timeout=60, **extra_kwargs)
            if gen_file.exists():
                if gen_file.resolve() != out_p.resolve():
                    if out_p.exists():
                        out_p.unlink()
                    gen_file.replace(out_p)
                return True, "Thành công qua LibreOffice Calc"

            return False, "LibreOffice không sinh tệp .xlsx đích."
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _convert_xls_via_openpyxl_enhanced(src_p: Path, out_p: Path) -> tuple[bool, str]:
        """
        Chuyển đổi .xls sang .xlsx bằng bộ thư viện thuần Python xlrd + openpyxl nâng cao:
        - Khôi phục 100% các ô gộp (Merged Cells) từ sh.merged_cells
        - Tự động tính toán độ rộng cột (Auto-fit Widths) chống lỗi ### và tràn text
        - Đóng khung viền bảng mỏng (Thin Grid Borders) và bật hiển thị lưới (Gridlines)
        - Tạo kiểu nổi bật cho hàng tiêu đề (In đậm, Căn giữa, Màu nền)
        - Bảo toàn kiểu dữ liệu (Date, Số, Text tiếng Việt).
        """
        if not HAS_XLRD:
            return False, "Thiếu thư viện xlrd."

        try:
            try:
                wb_rd = xlrd.open_workbook(str(src_p), formatting_info=True)
            except Exception:
                wb_rd = xlrd.open_workbook(str(src_p), formatting_info=False)

            wb_out = openpyxl.Workbook()
            wb_out.remove(wb_out.active)

            thin_side = Side(border_style="thin", color="D3D3D3")
            data_border = Border(left=thin_side, right=thin_side, top=thin_side, bottom=thin_side)
            header_fill = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
            header_font = Font(name="Times New Roman", size=11, bold=True)
            data_font = Font(name="Times New Roman", size=11)

            for s_idx in range(wb_rd.nsheets):
                sh = wb_rd.sheet_by_index(s_idx)
                sh_title = sh.name[:31] if sh.name else f"Sheet{s_idx + 1}"
                ws_out = wb_out.create_sheet(title=sh_title)
                try:
                    ws_out.views.sheetView[0].showGridLines = True
                except Exception:
                    pass

                col_widths: dict[int, int] = {}

                for r in range(sh.nrows):
                    row_vals = []
                    for c in range(sh.ncols):
                        cell = sh.cell(r, c)
                        val = None
                        if cell.ctype == xlrd.XL_CELL_DATE:
                            try:
                                dt = datetime(*xlrd.xldate_as_tuple(cell.value, wb_rd.datemode))
                                val = dt.strftime("%d/%m/%Y")
                            except Exception:
                                val = str(cell.value)
                        elif cell.ctype == xlrd.XL_CELL_NUMBER:
                            if cell.value.is_integer():
                                val = int(cell.value)
                            else:
                                val = cell.value
                        elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                            val = bool(cell.value)
                        elif cell.ctype == xlrd.XL_CELL_EMPTY:
                            val = None
                        else:
                            val = cell.value

                        row_vals.append(val)
                        val_str = str(val) if val is not None else ""
                        col_widths[c] = max(col_widths.get(c, 8), len(val_str) + 3)

                    ws_out.append(row_vals)

                    row_idx = r + 1
                    for col_idx in range(1, sh.ncols + 1):
                        cell_out = ws_out.cell(row=row_idx, column=col_idx)
                        cell_out.border = data_border
                        if row_idx == 1:
                            cell_out.font = header_font
                            cell_out.fill = header_fill
                            cell_out.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                        else:
                            cell_out.font = data_font
                            cell_out.alignment = Alignment(vertical="center")

                # Khôi phục các ô gộp (Merged Cells)
                if hasattr(sh, "merged_cells") and sh.merged_cells:
                    for (rlo, rhi, clo, chi) in sh.merged_cells:
                        try:
                            ws_out.merge_cells(
                                start_row=rlo + 1,
                                end_row=rhi,
                                start_column=clo + 1,
                                end_column=chi
                            )
                        except Exception:
                            pass

                # Thiết lập độ rộng cột (Auto-fit Widths)
                for c_idx, width in col_widths.items():
                    col_letter = get_column_letter(c_idx + 1)
                    ws_out.column_dimensions[col_letter].width = min(max(width, 10), 60)

            wb_out.save(out_p)
            wb_out.close()

            if out_p.exists() and out_p.stat().st_size > 0:
                return True, "Thành công qua xlrd + openpyxl (Enhanced)"
            return False, "Không thể tạo file .xlsx qua openpyxl"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def convert_xls_to_xlsx(
        xls_path: Path | str,
        output_path: Path | str | None = None,
        delete_source: bool = False,
        auto_quit: bool = True
    ) -> tuple[bool, Path | None, str]:
        """
        Chuyển đổi một tệp .xls (Excel 97-2003) sang .xlsx trực tiếp bằng Microsoft Excel:
        - Mở trực tiếp bằng Microsoft Excel (AppleScript trên macOS, COM trên Windows).
        - Lưu sang .xlsx chuẩn Native để giữ nguyên 100% định dạng, kẻ khung viền, ô gộp, công thức, độ rộng cột, font chữ.
        - Tự động đóng tệp .xls cũ ngay sau khi lưu để không lưu lại bất kỳ cửa sổ nào.
        - Chờ người dùng cấp phép truy cập thư mục nếu cần (hỗ trợ timeout lên đến 120s).
        - Thực hiện nghiêm ngặt qua Microsoft Excel, không sử dụng các phương án dự phòng làm sai lệch định dạng.
        """
        p = Path(xls_path)
        if not p.exists() or p.suffix.lower() != ".xls":
            return False, None, f"Tệp không hợp lệ: {p.name}"

        out_p = Path(output_path) if output_path else p.with_suffix(".xlsx")
        out_p.parent.mkdir(parents=True, exist_ok=True)

        # 0. Nhận diện định dạng thực tế (trường hợp tệp thực chất là .xlsx bị đổi đuôi nhầm thành .xls)
        real_type = FileFormatConverter.detect_file_type(p)
        if real_type == "xlsx_zip":
            try:
                shutil.copy2(p, out_p)
                valid, msg = FileFormatConverter.validate_xlsx_integrity(out_p)
                if valid:
                    if delete_source:
                        try:
                            p.unlink()
                        except Exception:
                            pass
                    return True, out_p, "Thành công (Tệp gốc là XLSX đổi tên)"
                else:
                    if out_p.exists():
                        out_p.unlink()
            except Exception:
                pass

        # Chuyển đổi trực tiếp bằng Microsoft Excel Native theo yêu cầu (không dùng phương án dự phòng khác)
        if sys.platform == "darwin":
            ok, msg = FileFormatConverter._convert_xls_via_excel_macos(p, out_p, auto_quit=auto_quit)
            method_used = "Microsoft Excel (macOS)"
        elif sys.platform == "win32":
            ok, msg = FileFormatConverter._convert_xls_via_excel_windows(p, out_p, auto_quit=auto_quit)
            method_used = "Microsoft Excel (Windows)"
        else:
            return False, None, "Hệ điều hành hiện tại không hỗ trợ Microsoft Excel."

        if not ok:
            if out_p.exists():
                try:
                    out_p.unlink()
                except Exception:
                    pass
            return False, None, f"Lỗi chuyển đổi qua Microsoft Excel: {msg}"

        # Kiểm tra tính toàn vẹn của tệp .xlsx sinh ra
        if out_p.exists():
            final_valid, final_val_msg = FileFormatConverter.validate_xlsx_integrity(out_p)
            if not final_valid:
                try:
                    out_p.unlink()
                except Exception:
                    pass
                return False, None, f"Tệp .xlsx sinh ra không hợp lệ: {final_val_msg}"

            if delete_source:
                try:
                    p.unlink()
                except Exception:
                    pass
            return True, out_p, f"Thành công ({method_used})"

        return False, None, "Microsoft Excel không tạo được tệp .xlsx đích."

    @staticmethod
    def batch_convert(
        files: list[Path],
        output_dir: Path | None = None,
        delete_source: bool = False,
        skip_existing: bool = True,
        progress_callback: Callable[[int, int, str], None] | None = None,
        log_callback: Callable[[str, str], None] | None = None,
        check_cancel: Callable[[], bool] | None = None,
        file_status_callback: Callable[[Path, str], None] | None = None
    ) -> dict[str, Any]:
        """
        Chuyển đổi hàng loạt danh sách các file .doc và .xls:
        - Báo cáo tiến trình thời gian thực
        - Hỗ trợ dừng giữa chừng qua check_cancel
        - Cập nhật trạng thái từng file qua file_status_callback
        """
        def log(msg, level="info"):
            if log_callback:
                log_callback(msg, level)
            else:
                print(msg)

        total = len(files)
        success_count = 0
        skipped_count = 0
        failed_count = 0
        results = []

        log(f"🚀 BẮT ĐẦU CHUYỂN ĐỔI HÀNG LOẠT {total} TỆP ĐỊNH DẠNG CŨ...", "info")

        # Đảm bảo đóng sạch các tiến trình cũ trước khi bắt đầu
        try:
            FileFormatConverter.close_office_apps()
        except Exception:
            pass

        try:
            for idx, f in enumerate(files, start=1):
                if check_cancel and check_cancel():
                    log("⏹ Đã dừng tiến trình chuyển đổi theo yêu cầu người dùng.", "warning")
                    break

                # Cập nhật tiến trình ngay khi bắt đầu xử lý file này
                if progress_callback:
                    progress_callback(idx - 1, total, f"Đang chuyển [{idx}/{total}]: {f.name}")
                if file_status_callback:
                    file_status_callback(f, "Đang chuyển...")

                ext = f.suffix.lower()
                target_ext = ".docx" if ext == ".doc" else ".xlsx"

                if output_dir:
                    dest_file = output_dir / f"{f.stem}{target_ext}"
                else:
                    dest_file = f.with_suffix(target_ext)

                if skip_existing and dest_file.exists() and dest_file.stat().st_size > 0:
                    log(f"  ⏭️ [{idx}/{total}] Đã tồn tại: {dest_file.name}, bỏ qua.", "info")
                    skipped_count += 1
                    results.append({"path": f, "out": dest_file, "status": "Bỏ qua (đã có)"})
                    if file_status_callback:
                        file_status_callback(f, "Bỏ qua (đã có)")
                    if progress_callback:
                        progress_callback(idx, total, f"Bỏ qua: {f.name}")
                    continue

                if ext == ".doc":
                    ok, out_p, msg = FileFormatConverter.convert_doc_to_docx(
                        doc_path=f,
                        output_path=dest_file,
                        delete_source=delete_source,
                        auto_quit=False
                    )
                elif ext == ".xls":
                    ok, out_p, msg = FileFormatConverter.convert_xls_to_xlsx(
                        xls_path=f,
                        output_path=dest_file,
                        delete_source=delete_source,
                        auto_quit=False
                    )
                else:
                    ok, out_p, msg = False, None, "Định dạng không được hỗ trợ"

                if ok:
                    success_count += 1
                    log(f"  ✅ [{idx}/{total}] Chuyển thành công: {f.name} $\\rightarrow$ {dest_file.name}", "success")
                    results.append({"path": f, "out": dest_file, "status": "Thành công"})
                    if file_status_callback:
                        file_status_callback(f, "Thành công")
                else:
                    failed_count += 1
                    log(f"  ❌ [{idx}/{total}] Lỗi chuyển đổi {f.name}: {msg}", "error")
                    results.append({"path": f, "out": None, "status": f"Lỗi: {msg}"})
                    if file_status_callback:
                        file_status_callback(f, f"Lỗi: {msg}")

                if progress_callback:
                    progress_callback(idx, total, f"Đã xong: {f.name}")
        finally:
            # Tự động đóng Word & Excel khi kết thúc hoặc huỷ
            FileFormatConverter.close_office_apps()

        log(f"🎉 HOÀN TẤT CHUYỂN ĐỔI: {success_count} thành công, {skipped_count} bỏ qua, {failed_count} lỗi.", "success" if failed_count == 0 else "warning")
        log("🧹 Đã tự động đóng và giải phóng tiến trình Microsoft Word & Excel.", "info")

        return {
            "status": "success",
            "total": total,
            "success": success_count,
            "skipped": skipped_count,
            "failed": failed_count,
            "results": results
        }

    @staticmethod
    def convert_excel_encoding_to_unicode(
        xlsx_path: Path | str,
        output_path: Path | str | None = None
    ) -> tuple[bool, Path | None, str, int]:
        """
        Quét toàn bộ ô trong tệp .xlsx và tự động chuyển đổi các ô dính bảng mã TCVN3/VNI sang Unicode.
        Nếu font chữ là .VnTime hoặc VNI-Times, tự động chuyển sang Times New Roman.
        Trả về (success, output_path, message, converted_cells_count).
        """
        p = Path(xlsx_path)
        if not p.exists() or p.suffix.lower() != ".xlsx":
            return False, None, f"Tệp không hợp lệ: {p.name}", 0

        out_p = Path(output_path) if output_path else p.parent / f"{p.stem}_Unicode.xlsx"

        try:
            wb = openpyxl.load_workbook(p, data_only=False)
        except Exception as e:
            return False, None, f"Không thể đọc tệp Excel: {e}", 0

        converted_count = 0
        try:
            for ws in wb.worksheets:
                for row in ws.iter_rows():
                    for cell in row:
                        val = cell.value
                        if isinstance(val, str) and val.strip():
                            if is_tcvn3(val) or is_vni(val):
                                converted_val = auto_convert_legacy_encoding(val)
                                if converted_val != val:
                                    cell.value = converted_val
                                    converted_count += 1

                        # Sửa font nếu là font cổ TCVN3 / VNI
                        if cell.font and cell.font.name:
                            f_name_lower = cell.font.name.lower()
                            if f_name_lower.startswith(".vn") or f_name_lower.startswith("vni-"):
                                new_font = Font(
                                    name="Times New Roman",
                                    size=cell.font.size,
                                    bold=cell.font.bold,
                                    italic=cell.font.italic,
                                    color=cell.font.color
                                )
                                cell.font = new_font

            wb.save(out_p)
            wb.close()
            return True, out_p, f"Đã chuyển đổi thành công {converted_count} ô sang Unicode chuẩn.", converted_count
        except Exception as e:
            try:
                wb.close()
            except Exception:
                pass
            return False, None, f"Lỗi chuyển đổi mã Unicode: {e}", 0
