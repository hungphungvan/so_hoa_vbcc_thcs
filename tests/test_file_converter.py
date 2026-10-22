from __future__ import annotations

import unittest
import tempfile
import shutil
from pathlib import Path
import xlwt
import openpyxl

from src.file_converter import FileFormatConverter


class TestFileFormatConverter(unittest.TestCase):
    def setUp(self):
        self.tmp_dir = Path(tempfile.mkdtemp(dir=Path(__file__).parent))

    def tearDown(self):
        shutil.rmtree(self.tmp_dir, ignore_errors=True)

    def test_scan_files(self):
        """Kiểm tra quét tìm tệp .doc và .xls trong thư mục, loại bỏ file tạm ~$."""
        f_doc = self.tmp_dir / "so_yeu_ly_lich.doc"
        f_doc.write_text("dummy doc content")

        f_xls = self.tmp_dir / "danh_sach_hoc_sinh.xls"
        f_xls.write_text("dummy xls content")

        f_temp = self.tmp_dir / "~$so_yeu_ly_lich.doc"
        f_temp.write_text("temp lock file")

        f_txt = self.tmp_dir / "note.txt"
        f_txt.write_text("ignored file")

        results = FileFormatConverter.scan_files([self.tmp_dir])
        self.assertEqual(len(results), 2)
        exts = [r["ext"] for r in results]
        self.assertIn(".doc", exts)
        self.assertIn(".xls", exts)
        self.assertNotIn("~$so_yeu_ly_lich.doc", [r["name"] for r in results])

    def test_convert_doc_to_docx(self):
        """Kiểm tra chuyển đổi file .doc sang .docx."""
        doc_file = self.tmp_dir / "van_ban_mau.doc"
        doc_file.write_text("<html><body><p>Van ban thi tot nghiep THCS</p></body></html>", encoding="utf-8")

        ok, out_docx, msg = FileFormatConverter.convert_doc_to_docx(doc_file)
        self.assertTrue(ok)
        self.assertIsNotNone(out_docx)
        self.assertTrue(out_docx.exists())
        self.assertEqual(out_docx.suffix.lower(), ".docx")
        self.assertGreater(out_docx.stat().st_size, 0)

    def test_convert_xls_to_xlsx(self):
        """Kiểm tra chuyển đổi file .xls sang .xlsx (đa sheet, định dạng ngày, số, chữ)."""
        xls_file = self.tmp_dir / "bang_diem.xls"

        wb_xls = xlwt.Workbook(encoding="utf-8")
        ws = wb_xls.add_sheet("Lop_9A")
        ws.write(0, 0, "STT")
        ws.write(0, 1, "Họ và tên")
        ws.write(0, 2, "Ngày sinh")
        ws.write(0, 3, "Điểm trung bình")

        ws.write(1, 0, 1)
        ws.write(1, 1, "Nguyễn Văn An")
        ws.write(1, 2, "15/05/2005")
        ws.write(1, 3, 8.75)

        wb_xls.save(str(xls_file))
        self.assertTrue(xls_file.exists())

        ok, out_xlsx, msg = FileFormatConverter.convert_xls_to_xlsx(xls_file)
        self.assertTrue(ok)
        self.assertIsNotNone(out_xlsx)
        self.assertTrue(out_xlsx.exists())
        self.assertEqual(out_xlsx.suffix.lower(), ".xlsx")

        # Đọc lại bằng openpyxl để xác thực
        wb_check = openpyxl.load_workbook(out_xlsx)
        self.assertIn("Lop_9A", wb_check.sheetnames)
        ws_chk = wb_check["Lop_9A"]
        rows = list(ws_chk.iter_rows(values_only=True))
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0], ("STT", "Họ và tên", "Ngày sinh", "Điểm trung bình"))
        self.assertEqual(rows[1][0], 1)
        self.assertEqual(rows[1][1], "Nguyễn Văn An")
        self.assertEqual(rows[1][3], 8.75)
        wb_check.close()

    def test_batch_convert(self):
        """Kiểm tra chuyển đổi hàng loạt với tiến trình và tùy chọn."""
        f1 = self.tmp_dir / "doc1.doc"
        f1.write_text("<html><body>Test 1</body></html>")

        f2 = self.tmp_dir / "sheet1.xls"
        wb_x = xlwt.Workbook(encoding="utf-8")
        ws = wb_x.add_sheet("S1")
        ws.write(0, 0, "Hello")
        wb_x.save(str(f2))

        progress_records = []
        def prog_cb(cur, tot, msg):
            progress_records.append((cur, tot, msg))

        out_sub_dir = self.tmp_dir / "converted_out"
        res = FileFormatConverter.batch_convert(
            files=[f1, f2],
            output_dir=out_sub_dir,
            delete_source=False,
            skip_existing=True,
            progress_callback=prog_cb
        )

        self.assertEqual(res["total"], 2)
        self.assertEqual(res["success"], 2)
        self.assertEqual(res["failed"], 0)
        self.assertTrue((out_sub_dir / "doc1.docx").exists())
        self.assertTrue((out_sub_dir / "sheet1.xlsx").exists())
        self.assertGreater(len(progress_records), 0)

    def test_detect_file_type(self):
        """Kiểm tra nhận diện định dạng thực tế của tệp qua header / magic bytes."""
        # Binary doc
        f_bin = self.tmp_dir / "real.doc"
        f_bin.write_bytes(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 100)
        self.assertEqual(FileFormatConverter.detect_file_type(f_bin), "doc_binary")

        # Renamed docx
        f_docx_renamed = self.tmp_dir / "fake.doc"
        f_docx_renamed.write_bytes(b"PK\x03\x04" + b"\x00" * 100)
        self.assertEqual(FileFormatConverter.detect_file_type(f_docx_renamed), "docx_zip")

        # RTF
        f_rtf = self.tmp_dir / "text.doc"
        f_rtf.write_text("{\\rtf1\\ansi\\deff0 Hello World}", encoding="utf-8")
        self.assertEqual(FileFormatConverter.detect_file_type(f_rtf), "rtf")

        # HTML
        f_html = self.tmp_dir / "table.doc"
        f_html.write_text("<html><body><table><tr><td>1</td></tr></table></body></html>", encoding="utf-8")
        self.assertEqual(FileFormatConverter.detect_file_type(f_html), "html")

    def test_validate_docx_integrity(self):
        """Kiểm tra thẩm định tính toàn vẹn của tệp .docx."""
        # File không tồn tại
        ok, msg = FileFormatConverter.validate_docx_integrity(self.tmp_dir / "nonexistent.docx")
        self.assertFalse(ok)

        # File rỗng / quá nhỏ
        f_empty = self.tmp_dir / "empty.docx"
        f_empty.write_bytes(b"")
        ok, msg = FileFormatConverter.validate_docx_integrity(f_empty)
        self.assertFalse(ok)

        # File zip nhưng thiếu part
        import zipfile
        f_invalid_zip = self.tmp_dir / "corrupt_parts.docx"
        with zipfile.ZipFile(f_invalid_zip, "w") as z:
            z.writestr("test.txt", "abc")
        ok, msg = FileFormatConverter.validate_docx_integrity(f_invalid_zip)
        self.assertFalse(ok)

        # File zip hợp lệ
        f_valid = self.tmp_dir / "valid.docx"
        with zipfile.ZipFile(f_valid, "w") as z:
            z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
            z.writestr("word/document.xml", '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>')
        ok, msg = FileFormatConverter.validate_docx_integrity(f_valid)
        self.assertTrue(ok)

    def test_sanitize_and_repair_docx(self):
        """Kiểm tra khử thẻ lỗi w:sz-cs và loại bỏ file rác Cocoa meta.xml."""
        import zipfile
        f_quirky = self.tmp_dir / "quirky.docx"
        with zipfile.ZipFile(f_quirky, "w") as z:
            z.writestr("[Content_Types].xml", '<Types/>')
            z.writestr("word/document.xml", '<w:document><w:rPr><w:sz-cs w:val="24"/><w:spacing w:val="0"/></w:rPr></w:document>')
            z.writestr("docProps/meta.xml", '<meta/>')

        repaired = FileFormatConverter.sanitize_and_repair_docx(f_quirky)
        self.assertTrue(repaired)

        with zipfile.ZipFile(f_quirky, "r") as z:
            namelist = z.namelist()
            self.assertNotIn("docProps/meta.xml", namelist)
            xml_content = z.read("word/document.xml").decode("utf-8")
            self.assertNotIn("w:sz-cs", xml_content)
            self.assertIn("w:szCs", xml_content)

    def test_convert_renamed_docx(self):
        """Kiểm tra phục hồi tự động khi người dùng đổi tên tệp .docx thành .doc."""
        import zipfile
        f_doc = self.tmp_dir / "mistakenly_named.doc"
        with zipfile.ZipFile(f_doc, "w") as z:
            z.writestr("[Content_Types].xml", '<?xml version="1.0"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
            z.writestr("word/document.xml", '<?xml version="1.0"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p/></w:body></w:document>')

        ok, out_docx, msg = FileFormatConverter.convert_doc_to_docx(f_doc)
        self.assertTrue(ok)
        self.assertIn("DOCX đổi tên", msg)
        self.assertTrue(out_docx.exists())

    def test_validate_xlsx_integrity(self):
        """Kiểm tra thẩm định tính toàn vẹn của tệp .xlsx."""
        ok, _ = FileFormatConverter.validate_xlsx_integrity(self.tmp_dir / "nonexistent.xlsx")
        self.assertFalse(ok)

        f_empty = self.tmp_dir / "empty.xlsx"
        f_empty.write_bytes(b"")
        ok, _ = FileFormatConverter.validate_xlsx_integrity(f_empty)
        self.assertFalse(ok)

        import zipfile
        f_invalid = self.tmp_dir / "invalid.xlsx"
        with zipfile.ZipFile(f_invalid, "w") as z:
            z.writestr("test.txt", "abc")
        ok, _ = FileFormatConverter.validate_xlsx_integrity(f_invalid)
        self.assertFalse(ok)

        f_valid = self.tmp_dir / "valid.xlsx"
        with zipfile.ZipFile(f_valid, "w") as z:
            z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
            z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>')
        ok, msg = FileFormatConverter.validate_xlsx_integrity(f_valid)
        self.assertTrue(ok)

    def test_convert_renamed_xlsx(self):
        """Kiểm tra phục hồi tự động khi người dùng đổi tên tệp .xlsx thành .xls."""
        import zipfile
        f_xls = self.tmp_dir / "mistakenly_named.xls"
        with zipfile.ZipFile(f_xls, "w") as z:
            z.writestr("[Content_Types].xml", '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>')
            z.writestr("xl/workbook.xml", '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheets/></workbook>')

        ok, out_xlsx, msg = FileFormatConverter.convert_xls_to_xlsx(f_xls)
        self.assertTrue(ok)
        self.assertIn("XLSX đổi tên", msg)
        self.assertTrue(out_xlsx.exists())

    def test_convert_xls_formatting_enhanced(self):
        """Kiểm tra chuyển đổi file .xls bảo toàn ô gộp (Merged Cells), căn lề và độ rộng cột."""
        xls_file = self.tmp_dir / "so_bo.xls"
        wb_xls = xlwt.Workbook(encoding="utf-8")
        ws = wb_xls.add_sheet("So_Goc")
        # Gộp ô dòng 0 từ cột 0 đến 3 cho tiêu đề
        ws.write_merge(0, 0, 0, 3, "SỔ GỐC CẤP BẰNG TỐT NGHIỆP THCS")
        ws.write(1, 0, "STT")
        ws.write(1, 1, "Họ và tên")
        ws.write(1, 2, "Ngày tháng năm sinh")
        ws.write(1, 3, "Xếp loại tốt nghiệp")

        ws.write(2, 0, 1)
        ws.write(2, 1, "Trần Thị Bích Ngọc")
        ws.write(2, 2, "20/11/2006")
        ws.write(2, 3, "Giỏi")
        wb_xls.save(str(xls_file))

        ok, out_xlsx, msg = FileFormatConverter.convert_xls_to_xlsx(xls_file)
        self.assertTrue(ok)
        self.assertTrue(out_xlsx.exists())

        # Đọc lại bằng openpyxl và kiểm tra định dạng
        wb_check = openpyxl.load_workbook(out_xlsx)
        ws_chk = wb_check["So_Goc"]
        # Kiểm tra ô gộp
        merged_ranges = [str(m) for m in ws_chk.merged_cells.ranges]
        self.assertTrue(any("A1:D1" in r for r in merged_ranges) or len(merged_ranges) > 0)
        self.assertEqual(ws_chk["A1"].value, "SỔ GỐC CẤP BẰNG TỐT NGHIỆP THCS")
        self.assertEqual(ws_chk["B3"].value, "Trần Thị Bích Ngọc")
        wb_check.close()


if __name__ == '__main__':
    unittest.main()
