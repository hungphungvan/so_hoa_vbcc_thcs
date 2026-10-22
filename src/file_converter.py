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

import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

try:
    import xlrd
    HAS_XLRD = True
except ImportError:
    HAS_XLRD = False


class FileFormatConverter:
    """
    Bộ xử lý chuyển đổi định dạng tệp cũ sang định dạng mới:
    - .doc  -> .docx (Word 97-2003 -> Word 2007+)
    - .xls  -> .xlsx (Excel 97-2003 -> Excel OpenXML)
    Hỗ trợ đa nền tảng (macOS và Windows).
    """

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
        - 'doc_binary': OLE2 Compound File (Word 97-2003 chuẩn: D0 CF 11 E0 A1 B1 1A E1)
        - 'docx_zip': Gói nén PK ZIP (thực chất là .docx bị đổi tên thành .doc)
        - 'rtf': Rich Text Format ({\\rtf)
        - 'html': HTML Table/Document (<html, <!DOCTYPE)
        - 'unknown': Không xác định
        """
        p = Path(file_path)
        if not p.exists() or not p.is_file():
            return "unknown"
        try:
            with open(p, "rb") as f:
                header = f.read(512)
            if header.startswith(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"):
                return "doc_binary"
            if header.startswith(b"PK\x03\x04"):
                return "docx_zip"
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

    @staticmethod
    def _convert_doc_via_word_macos(src_p: Path, out_p: Path) -> tuple[bool, str]:
        """Chuyển đổi .doc sang .docx thông qua Microsoft Word trên macOS bằng AppleScript."""
        if not Path("/Applications/Microsoft Word.app").exists():
            return False, "Microsoft Word không được cài đặt trong /Applications."

        src_abs = str(src_p.resolve())
        out_abs = str(out_p.resolve())

        as_script = f'''
        tell application "Microsoft Word"
            set origAlerts to display alerts
            set display alerts to none
            try
                set theDoc to open (POSIX file "{src_abs}")
                save as theDoc file name (POSIX file "{out_abs}" as text) file format format document default
                close theDoc saving no
                set display alerts to origAlerts
                return "OK"
            on error errMsg
                try
                    set display alerts to origAlerts
                end try
                return "ERR:" & errMsg
            end try
        end tell
        '''
        try:
            res = subprocess.run(
                ["osascript", "-e", as_script],
                capture_output=True,
                text=True,
                timeout=45
            )
            out_str = (res.stdout or "").strip()
            if "OK" in out_str and out_p.exists() and out_p.stat().st_size > 0:
                return True, "Thành công qua Microsoft Word (macOS)"
            return False, res.stderr.strip() or out_str
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _convert_doc_via_word_windows(src_p: Path, out_p: Path) -> tuple[bool, str]:
        """Chuyển đổi .doc sang .docx thông qua PowerShell COM Automation trên Windows."""
        src_abs = str(src_p.resolve()).replace("'", "''")
        out_abs = str(out_p.resolve()).replace("'", "''")

        ps_script = f'''
        $ErrorActionPreference = 'Stop'
        $word = $null
        $doc = $null
        try {{
            $word = New-Object -ComObject Word.Application
            $word.Visible = $false
            $word.DisplayAlerts = 0
            try {{ $word.AutomationSecurity = 3 }} catch {{}}

            $doc = $word.Documents.Open('{src_abs}', $false, $true, $false)
            $doc.SaveAs2('{out_abs}', 16)
            $doc.Close(0)
            $doc = $null
            Write-Output "OK"
        }} catch {{
            Write-Error $_.Exception.Message
        }} finally {{
            if ($doc -ne $null) {{
                try {{ $doc.Close(0) }} catch {{}}
            }}
            if ($word -ne $null) {{
                try {{ $word.Quit() }} catch {{}}
                try {{ [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null }} catch {{}}
            }}
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()
        }}
        '''
        try:
            res = subprocess.run(
                ["powershell", "-NoProfile", "-NonInteractive", "-Command", ps_script],
                capture_output=True,
                text=True,
                timeout=45
            )
            if "OK" in (res.stdout or "") and out_p.exists() and out_p.stat().st_size > 0:
                return True, "Thành công qua Microsoft Word (Windows)"
            return False, res.stderr.strip() or "Lỗi PowerShell COM Word"
        except Exception as e:
            return False, str(e)

    @staticmethod
    def _convert_doc_via_libreoffice(soffice_path: str, src_p: Path, out_p: Path) -> tuple[bool, str]:
        """Chuyển đổi thông qua LibreOffice headless CLI."""
        try:
            cmd = [
                soffice_path,
                "--headless",
                "--convert-to", "docx:MS Word 2007 XML",
                "--outdir", str(out_p.parent),
                str(src_p.resolve())
            ]
            subprocess.run(cmd, capture_output=True, text=True, timeout=60)
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
            subprocess.run(cmd_simple, capture_output=True, text=True, timeout=60)
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
        delete_source: bool = False
    ) -> tuple[bool, Path | None, str]:
        """
        Chuyển đổi một tệp .doc sang .docx theo chiến lược đa tầng:
        1. Kiểm tra định dạng thực tế (xử lý tệp .docx bị đổi tên nhầm thành .doc).
        2. Tầng 1: Sử dụng Microsoft Word Native (macOS AppleScript hoặc Windows COM) - bảo toàn 100% bảng biểu.
        3. Tầng 2: Sử dụng LibreOffice / soffice CLI (chuẩn mực nguồn mở).
        4. Tầng 3: Sử dụng textutil trên macOS (tự động khử lỗi XML w:sz-cs và file rác).
        5. Tầng 4: Kiểm tra tính toàn vẹn (Validation) của tệp .docx sinh ra. Tự động xóa file hỏng nếu phát hiện lỗi.
        """
        p = Path(doc_path)
        if not p.exists() or p.suffix.lower() != ".doc":
            return False, None, f"Tệp không hợp lệ: {p.name}"

        out_p = Path(output_path) if output_path else p.with_suffix(".docx")
        out_p.parent.mkdir(parents=True, exist_ok=True)

        # 0. Nhận diện định dạng thực tế
        real_type = FileFormatConverter.detect_file_type(p)

        # Trường hợp tệp thực chất là .docx bị đổi đuôi nhầm thành .doc
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

        success = False
        method_used = ""
        attempt_errors: list[str] = []

        # TẦNG 1: Microsoft Word Native (chính xác tuyệt đối)
        if sys.platform == "darwin":
            ok, msg = FileFormatConverter._convert_doc_via_word_macos(p, out_p)
            if ok:
                valid, val_err = FileFormatConverter.validate_docx_integrity(out_p)
                if valid:
                    success = True
                    method_used = "Microsoft Word (macOS)"
                else:
                    attempt_errors.append(f"Word Mac: {val_err}")
            else:
                attempt_errors.append(f"Word Mac: {msg}")

        elif sys.platform == "win32":
            ok, msg = FileFormatConverter._convert_doc_via_word_windows(p, out_p)
            if ok:
                valid, val_err = FileFormatConverter.validate_docx_integrity(out_p)
                if valid:
                    success = True
                    method_used = "Microsoft Word (Windows)"
                else:
                    attempt_errors.append(f"Word Win: {val_err}")
            else:
                attempt_errors.append(f"Word Win: {msg}")

        # TẦNG 2: LibreOffice CLI (chất lượng mã nguồn mở cao cấp)
        if not success:
            soffice_bin = FileFormatConverter.find_libreoffice()
            if soffice_bin:
                ok, msg = FileFormatConverter._convert_doc_via_libreoffice(soffice_bin, p, out_p)
                if ok:
                    valid, val_err = FileFormatConverter.validate_docx_integrity(out_p)
                    if valid:
                        success = True
                        method_used = "LibreOffice"
                    else:
                        attempt_errors.append(f"LibreOffice: {val_err}")
                else:
                    attempt_errors.append(f"LibreOffice: {msg}")

        # TẦNG 3: macOS textutil (chỉ chạy trên macOS khi không có Word/LibreOffice)
        if not success and sys.platform == "darwin":
            ok, msg = FileFormatConverter._convert_doc_via_textutil_macos(p, out_p)
            if ok:
                valid, val_err = FileFormatConverter.validate_docx_integrity(out_p)
                if valid:
                    success = True
                    method_used = "textutil (macOS)"
                else:
                    attempt_errors.append(f"textutil: {val_err}")
            else:
                attempt_errors.append(f"textutil: {msg}")

        # TẦNG 4: Thẩm định và Xử lý kết quả cuối cùng
        if success and out_p.exists():
            final_valid, final_val_msg = FileFormatConverter.validate_docx_integrity(out_p)
            if not final_valid:
                # XÓA BỎ NGAY FILE HỎNG để tránh người dùng mở ra gặp lỗi
                try:
                    out_p.unlink()
                except Exception:
                    pass
                success = False
                err_detail = f"Tệp .docx sinh ra không hợp lệ: {final_val_msg}"
            else:
                if delete_source:
                    try:
                        p.unlink()
                    except Exception:
                        pass
                return True, out_p, f"Thành công ({method_used})"

        # Nếu thất bại hoàn toàn
        if out_p.exists():
            try:
                out_p.unlink()
            except Exception:
                pass

        err_summary = " | ".join(attempt_errors) if attempt_errors else "Không tìm thấy công cụ chuyển đổi tương thích."
        advice = " Khuyến nghị: Cài đặt LibreOffice (miễn phí) hoặc Microsoft Word để chuyển đổi tệp .doc phức tạp."
        return False, None, f"Lỗi chuyển đổi .doc: {err_summary}.{advice}"

    @staticmethod
    def convert_xls_to_xlsx(
        xls_path: Path | str,
        output_path: Path | str | None = None,
        delete_source: bool = False
    ) -> tuple[bool, Path | None, str]:
        """
        Chuyển đổi một tệp .xls (Excel 97-2003) sang .xlsx:
        - Sử dụng thư viện Python thuần xlrd + openpyxl (hoạt động độc lập 100% không cần Office).
        - Đọc đầy đủ các sheet, bảo toàn kiểu dữ liệu (Date, Number, Text, Boolean).
        - Fallback: LibreOffice nếu có.
        """
        p = Path(xls_path)
        if not p.exists() or p.suffix.lower() != ".xls":
            return False, None, f"Tệp không hợp lệ: {p.name}"

        out_p = Path(output_path) if output_path else p.with_suffix(".xlsx")
        out_p.parent.mkdir(parents=True, exist_ok=True)

        # 1. Thử bằng xlrd + openpyxl (nhanh, nhẹ, độc lập nền tảng)
        if HAS_XLRD:
            try:
                wb_rd = xlrd.open_workbook(str(p), formatting_info=False)
                wb_out = openpyxl.Workbook()
                wb_out.remove(wb_out.active) # Xóa sheet mặc định

                for s_idx in range(wb_rd.nsheets):
                    sh = wb_rd.sheet_by_index(s_idx)
                    # Giới hạn tên sheet tối đa 31 ký tự theo chuẩn Excel
                    sh_title = sh.name[:31] if sh.name else f"Sheet{s_idx + 1}"
                    ws_out = wb_out.create_sheet(title=sh_title)

                    for r in range(sh.nrows):
                        row_vals = []
                        for c in range(sh.ncols):
                            cell = sh.cell(r, c)
                            if cell.ctype == xlrd.XL_CELL_DATE:
                                try:
                                    dt = datetime(*xlrd.xldate_as_tuple(cell.value, wb_rd.datemode))
                                    # Định dạng ngày chuẩn dd/mm/yyyy
                                    row_vals.append(dt.strftime("%d/%m/%Y"))
                                except Exception:
                                    row_vals.append(str(cell.value))
                            elif cell.ctype == xlrd.XL_CELL_NUMBER:
                                if cell.value.is_integer():
                                    row_vals.append(int(cell.value))
                                else:
                                    row_vals.append(cell.value)
                            elif cell.ctype == xlrd.XL_CELL_BOOLEAN:
                                row_vals.append(bool(cell.value))
                            elif cell.ctype == xlrd.XL_CELL_EMPTY:
                                row_vals.append(None)
                            else:
                                row_vals.append(cell.value)
                        ws_out.append(row_vals)

                wb_out.save(out_p)
                wb_out.close()

                if out_p.exists() and out_p.stat().st_size > 0:
                    if delete_source:
                        try:
                            p.unlink()
                        except Exception:
                            pass
                    return True, out_p, "Thành công"
            except Exception as e:
                pass

        # 2. Fallback dùng LibreOffice (nếu có trên máy)
        soffice_cmd = FileFormatConverter.find_libreoffice()
        if soffice_cmd:
            try:
                cmd = [soffice_cmd, "--headless", "--convert-to", "xlsx", "--outdir", str(out_p.parent), str(p)]
                subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                gen_file = out_p.parent / f"{p.stem}.xlsx"
                if gen_file.exists():
                    if gen_file != out_p:
                        gen_file.replace(out_p)
                    if delete_source:
                        try:
                            p.unlink()
                        except Exception:
                            pass
                    return True, out_p, "Thành công (qua LibreOffice)"
            except Exception as e:
                return False, None, f"Lỗi chuyển đổi LibreOffice: {e}"

        return False, None, "Không thể chuyển đổi file .xls (Yêu cầu thư viện xlrd hoặc LibreOffice)."

    @staticmethod
    def batch_convert(
        files: list[Path],
        output_dir: Path | None = None,
        delete_source: bool = False,
        skip_existing: bool = True,
        progress_callback: Callable[[int, int, str], None] | None = None,
        log_callback: Callable[[str, str], None] | None = None,
        check_cancel: Callable[[], bool] | None = None
    ) -> dict[str, Any]:
        """
        Chuyển đổi hàng loạt danh sách các file .doc và .xls:
        - Báo cáo tiến trình thời gian thực
        - Hỗ trợ dừng giữa chừng qua check_cancel
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

        for idx, f in enumerate(files, start=1):
            if check_cancel and check_cancel():
                log("⏹ Đã dừng tiến trình chuyển đổi theo yêu cầu người dùng.", "warning")
                break

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
                if progress_callback:
                    progress_callback(idx, total, f"Bỏ qua: {f.name}")
                continue

            if ext == ".doc":
                ok, out_p, msg = FileFormatConverter.convert_doc_to_docx(
                    doc_path=f,
                    output_path=dest_file,
                    delete_source=delete_source
                )
            elif ext == ".xls":
                ok, out_p, msg = FileFormatConverter.convert_xls_to_xlsx(
                    xls_path=f,
                    output_path=dest_file,
                    delete_source=delete_source
                )
            else:
                ok, out_p, msg = False, None, "Định dạng không được hỗ trợ"

            if ok:
                success_count += 1
                log(f"  ✅ [{idx}/{total}] Chuyển thành công: {f.name} $\\rightarrow$ {dest_file.name}", "success")
                results.append({"path": f, "out": dest_file, "status": "Thành công"})
            else:
                failed_count += 1
                log(f"  ❌ [{idx}/{total}] Lỗi chuyển đổi {f.name}: {msg}", "error")
                results.append({"path": f, "out": None, "status": f"Lỗi: {msg}"})

            if progress_callback:
                progress_callback(idx, total, f"Đang chuyển: {f.name}")

        log(f"🎉 HOÀN TẤT CHUYỂN ĐỔI: {success_count} thành công, {skipped_count} bỏ qua, {failed_count} lỗi.", "success" if failed_count == 0 else "warning")

        return {
            "status": "success",
            "total": total,
            "success": success_count,
            "skipped": skipped_count,
            "failed": failed_count,
            "results": results
        }
