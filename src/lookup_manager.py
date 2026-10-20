from __future__ import annotations

import os
import sys
import re
from pathlib import Path
import openpyxl

def get_data_dir() -> Path:
    """Xác định đường dẫn thư mục data chứa các file mẫu và danh mục tra cứu."""
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        base = Path(sys._MEIPASS)
        if (base / "data").exists():
            return base / "data"
        if (base / "Danh_Muc_Tra_Cuu.xlsx").exists():
            return base
    # Chạy bình thường từ mã nguồn
    cur = Path(__file__).resolve().parent.parent / "data"
    if cur.exists():
        return cur
    cwd_data = Path.cwd() / "data"
    if cwd_data.exists():
        return cwd_data
    return Path(__file__).resolve().parent.parent

class LookupManager:
    """Quản lý nạp và tra cứu từ tệp Danh_Muc_Tra_Cuu.xlsx."""
    def __init__(self, excel_path: Path | str | None = None):
        if excel_path is None:
            self.excel_path = get_data_dir() / "Danh_Muc_Tra_Cuu.xlsx"
        else:
            self.excel_path = Path(excel_path)
            
        self.units: dict[str, dict] = {}  # {unit_key: {"key": ..., "code": ..., "name": ..., "commune": ..., "schools": [...], "school_map": {...}}}
        self.unit_names_list: list[str] = []
        self.gender_map: dict[str, str] = {}
        self.ethnic_map: dict[str, str] = {}
        self.nationality_map: dict[str, str] = {}
        self.birthplace_map: dict[str, str] = {}
        
        self.load_all()

    def load_all(self):
        """Nạp toàn bộ 5 sheet từ tệp Danh_Muc_Tra_Cuu.xlsx vào bộ nhớ."""
        if not self.excel_path.exists():
            raise FileNotFoundError(f"Không tìm thấy tệp danh mục tra cứu: {self.excel_path}")

        wb = openpyxl.load_workbook(self.excel_path, data_only=True)
        
        # 1. Sheet DS_Don_Vi
        if "DS_Don_Vi" in wb.sheetnames:
            ws = wb["DS_Don_Vi"]
            # Cột: 1: Mã trường, 2: Tên trường đã học, 3: Mã đơn vị, 4: Tên đơn vị, 5: Xã/phường
            for r in range(2, ws.max_row + 1):
                sch_code = str(ws.cell(r, 1).value or '').strip()
                sch_name = str(ws.cell(r, 2).value or '').strip()
                u_code = str(ws.cell(r, 3).value or '').strip()
                u_name = str(ws.cell(r, 4).value or '').strip()
                commune = str(ws.cell(r, 5).value or '').strip()
                
                if u_name:
                    # Tạo khóa định danh duy nhất hiển thị cả Tên đơn vị, Xã/phường và Mã đơn vị
                    if commune:
                        unit_key = f"{u_name} - {commune} [Mã: {u_code}]"
                    else:
                        unit_key = f"{u_name} [Mã: {u_code}]"

                    if unit_key not in self.units:
                        self.units[unit_key] = {
                            "key": unit_key,
                            "code": u_code,
                            "name": u_name,
                            "commune": commune,
                            "schools": [],
                            "school_map": {}
                        }

                    u_dict = self.units[unit_key]
                    if sch_name and sch_name not in u_dict["schools"]:
                        u_dict["schools"].append(sch_name)
                    if sch_name and sch_code:
                        u_dict["school_map"][sch_name.lower()] = sch_code
                        
            self.unit_names_list = sorted(list(self.units.keys()))

        # 2. Sheet Gioi_Tinh
        if "Gioi_Tinh" in wb.sheetnames:
            ws = wb["Gioi_Tinh"]
            for r in range(2, ws.max_row + 1):
                code = str(ws.cell(r, 1).value or '').strip()
                name = str(ws.cell(r, 2).value or '').strip()
                if code and name:
                    code_clean = re.sub(r'\.0+$', '', code)
                    self.gender_map[code_clean] = name

        # 3. Sheet Dan_Toc
        if "Dan_Toc" in wb.sheetnames:
            ws = wb["Dan_Toc"]
            for r in range(2, ws.max_row + 1):
                code = str(ws.cell(r, 1).value or '').strip()
                name = str(ws.cell(r, 2).value or '').strip()
                if code and name:
                    code_clean = re.sub(r'\.0+$', '', code)
                    self.ethnic_map[code_clean] = name

        # 4. Sheet Quoc_Tich
        if "Quoc_Tich" in wb.sheetnames:
            ws = wb["Quoc_Tich"]
            for r in range(2, ws.max_row + 1):
                code = str(ws.cell(r, 1).value or '').strip()
                name = str(ws.cell(r, 2).value or '').strip()
                if code and name:
                    code_clean = re.sub(r'\.0+$', '', code)
                    self.nationality_map[code_clean] = name

        # 5. Sheet Noi_Sinh
        if "Noi_Sinh" in wb.sheetnames:
            ws = wb["Noi_Sinh"]
            for r in range(2, ws.max_row + 1):
                code = str(ws.cell(r, 1).value or '').strip()
                name = str(ws.cell(r, 2).value or '').strip()
                if code and name:
                    code_clean = re.sub(r'\.0+$', '', code)
                    self.birthplace_map[code_clean] = name

        wb.close()

    def get_unit_info(self, unit_identifier: str) -> dict | None:
        """Lấy thông tin đơn vị từ unit_key, tên đơn vị hoặc mã đơn vị."""
        if not unit_identifier:
            return None
        u_clean = unit_identifier.strip()
        # 1. Khớp trực tiếp unit_key
        if u_clean in self.units:
            return self.units[u_clean]

        # 2. Tìm theo tên đơn vị hoặc mã đơn vị
        for key, info in self.units.items():
            if key.lower() == u_clean.lower() or info["name"].lower() == u_clean.lower():
                return info
            if info["code"] == u_clean:
                return info

        return None

    def get_school_code(self, unit_info: dict, school_name: str) -> str:
        """Tra cứu mã trường trực thuộc đơn vị từ tên trường đã học."""
        if not unit_info or not school_name:
            return ""
        s_map = unit_info.get("school_map", {})
        s_clean = school_name.strip().lower()
        if s_clean in s_map:
            return s_map[s_clean]

        # Khớp mềm (bỏ khoảng trắng và ký tự đặc biệt)
        s_alphanumeric = re.sub(r'\W+', '', s_clean)
        for k, code in s_map.items():
            k_alphanumeric = re.sub(r'\W+', '', k)
            if s_alphanumeric == k_alphanumeric:
                return code

        # Tìm kiếm chuỗi con
        for k, code in s_map.items():
            if k in s_clean or s_clean in k:
                return code

        return ""

    def map_gender(self, val) -> str:
        """Ánh xạ giới tính từ mã hoặc làm sạch chuỗi."""
        if val is None:
            return ""
        s = str(val).strip()
        if not s or s.lower() in ("none", "nan", "null"):
            return ""
        s_clean = re.sub(r'\.0+$', '', s)
        if s_clean in self.gender_map:
            return self.gender_map[s_clean]
        s_lower = s_clean.lower()
        if s_lower == "nam":
            return "Nam"
        elif s_lower in ("nữ", "nu"):
            return "Nữ"
        elif s_lower in ("khác", "khac"):
            return "Khác"
        return s_clean

    def map_ethnic(self, val) -> str:
        """Ánh xạ dân tộc từ mã hoặc trả về chuỗi gốc."""
        if val is None:
            return ""
        s = str(val).strip()
        if not s or s.lower() in ("none", "nan", "null"):
            return ""
        s_clean = re.sub(r'\.0+$', '', s)
        if s_clean in self.ethnic_map:
            return self.ethnic_map[s_clean]
        return s_clean

    def map_nationality(self, val) -> str:
        """Ánh xạ quốc tịch từ mã hoặc trả về chuỗi gốc."""
        if val is None:
            return ""
        s = str(val).strip()
        if not s or s.lower() in ("none", "nan", "null"):
            return ""
        s_clean = re.sub(r'\.0+$', '', s)
        if s_clean in self.nationality_map:
            return self.nationality_map[s_clean]
        if s_clean.lower() in ("việt nam", "viet nam", "vn"):
            return "Việt Nam"
        return s_clean

    def map_birthplace(self, val) -> str:
        """Ánh xạ nơi sinh từ mã hoặc trả về chuỗi gốc."""
        if val is None:
            return ""
        s = str(val).strip()
        if not s or s.lower() in ("none", "nan", "null"):
            return ""
        s_clean = re.sub(r'\.0+$', '', s)
        if s_clean in self.birthplace_map:
            return self.birthplace_map[s_clean]
        return s_clean
