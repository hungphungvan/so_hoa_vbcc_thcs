from __future__ import annotations

import os
import sys
import subprocess
import shutil
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
    def convert_doc_to_docx(
        doc_path: Path | str,
        output_path: Path | str | None = None,
        delete_source: bool = False
    ) -> tuple[bool, Path | None, str]:
        """
        Chuyển đổi một tệp .doc sang .docx:
        - Trên macOS: Sử dụng tiện ích textutil có sẵn của hệ thống.
        - Trên Windows: Sử dụng PowerShell gọi Word.Application COM.
        - Fallback: LibreOffice nếu có.
        """
        p = Path(doc_path)
        if not p.exists() or p.suffix.lower() != ".doc":
            return False, None, f"Tệp không hợp lệ: {p.name}"

        out_p = Path(output_path) if output_path else p.with_suffix(".docx")
        out_p.parent.mkdir(parents=True, exist_ok=True)

        success = False
        error_msg = ""

        # 1. Thử trên macOS bằng textutil (nhanh nhất, tích hợp sẵn, không cần mở Word)
        if sys.platform == "darwin":
            try:
                cmd = ["textutil", "-convert", "docx", "-output", str(out_p.resolve()), str(p.resolve())]
                res = subprocess.run(cmd, capture_output=True, text=True, timeout=30)
                if res.returncode == 0 and out_p.exists() and out_p.stat().st_size > 0:
                    success = True
                else:
                    error_msg = res.stderr or "textutil không thể chuyển đổi tệp."
            except Exception as e:
                error_msg = str(e)

        # 2. Thử trên Windows bằng PowerShell Word.Application
        elif sys.platform == "win32":
            try:
                ps_script = (
                    f'$word = New-Object -ComObject Word.Application; '
                    f'$word.Visible = $false; '
                    f'$doc = $word.Documents.Open("{p.resolve()}"); '
                    f'$doc.SaveAs2("{out_p.resolve()}", 16); '
                    f'$doc.Close(); '
                    f'$word.Quit()'
                )
                res = subprocess.run(["powershell", "-NoProfile", "-Command", ps_script], capture_output=True, text=True, timeout=45)
                if out_p.exists() and out_p.stat().st_size > 0:
                    success = True
                else:
                    error_msg = res.stderr or "Không thể gọi Microsoft Word trên Windows."
            except Exception as e:
                error_msg = str(e)

        # 3. Fallback dùng LibreOffice (nếu có trên máy)
        if not success:
            soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
            if soffice_cmd:
                try:
                    cmd = [soffice_cmd, "--headless", "--convert-to", "docx", "--outdir", str(out_p.parent), str(p)]
                    subprocess.run(cmd, capture_output=True, text=True, timeout=60)
                    gen_file = out_p.parent / f"{p.stem}.docx"
                    if gen_file.exists():
                        if gen_file != out_p:
                            gen_file.replace(out_p)
                        success = True
                except Exception as e:
                    error_msg = f"{error_msg} | LibreOffice: {e}"

        if success:
            if delete_source:
                try:
                    p.unlink()
                except Exception:
                    pass
            return True, out_p, "Thành công"
        else:
            return False, None, error_msg or "Không tìm thấy công cụ chuyển đổi phù hợp cho định dạng .doc."

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
        soffice_cmd = shutil.which("soffice") or shutil.which("libreoffice")
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
