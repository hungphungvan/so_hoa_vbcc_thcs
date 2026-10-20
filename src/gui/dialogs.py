from __future__ import annotations

import re
import tkinter as tk
from tkinter import ttk, messagebox
from tkinter.scrolledtext import ScrolledText

def center_window(window, parent=None, width=None, height=None):
    """Căn giữa cửa sổ dialog trên cửa sổ cha hoặc màn hình."""
    window.update_idletasks()
    w = width or window.winfo_width()
    h = height or window.winfo_height()
    if parent:
        x = parent.winfo_x() + (parent.winfo_width() // 2) - (w // 2)
        y = parent.winfo_y() + (parent.winfo_height() // 2) - (h // 2)
    else:
        x = (window.winfo_screenwidth() // 2) - (w // 2)
        y = (window.winfo_screenheight() // 2) - (h // 2)
    window.geometry(f"{w}x{h}+{max(0, x)}+{max(0, y)}")

class SchoolBlockDialog(tk.Toplevel):
    """
    Hộp thoại tương tác khi phát hiện Khối học sinh thiếu hoặc sai Tên trường (Cột N).
    Hiển thị bảng đối chiếu 5 bạn trước và 5 bạn sau, cho phép chọn/áp dụng cho cả khối.
    """
    def __init__(self, parent, block_info: dict, allowed_schools: list[str], default_suggestion: str = ""):
        super().__init__(parent)
        self.title("Xử lý Tên trường học (Cột N)")
        self.transient(parent)
        self.grab_set()
        
        self.block_info = block_info
        self.allowed_schools = allowed_schools
        self.default_suggestion = default_suggestion
        self.result = None
        self.apply_to_all_similar = False
        
        self._build_ui()
        center_window(self, parent, 850, 520)

    def _build_ui(self):
        hdr_frame = ttk.Frame(self, padding=(12, 10))
        hdr_frame.pack(fill=tk.X)
        
        start_stt = self.block_info.get("start_stt", 1)
        end_stt = self.block_info.get("end_stt", 1)
        count = self.block_info.get("count", 1)
        raw_val = self.block_info.get("raw_val", "(Trống)")
        
        title_lbl = ttk.Label(
            hdr_frame, 
            text=f"⚠️ PHÁT HIỆN KHỐI {count} HỌC SINH CẦN XÁC ĐỊNH TÊN TRƯỜNG (STT {start_stt} -> {end_stt})",
            font=("Segoe UI", 11, "bold"),
            foreground="#B71C1C"
        )
        title_lbl.pack(anchor=tk.W)
        
        sub_lbl = ttk.Label(
            hdr_frame,
            text=f"Tệp: {self.block_info.get('file_name', '')} | Sheet: {self.block_info.get('sheet_name', '')} | Giá trị gốc: '{raw_val}'",
            font=("Segoe UI", 9)
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        tbl_frame = ttk.LabelFrame(self, text="📋 Bảng đối chiếu học sinh tham chiếu (5 trước và 5 sau)", padding=6)
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        
        columns = ("pos", "stt", "name", "school", "cohort", "grad")
        self.tree = ttk.Treeview(tbl_frame, columns=columns, show="headings", height=8)
        
        self.tree.heading("pos", text="")
        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Họ và tên")
        self.tree.heading("school", text="Tên trường đã học")
        self.tree.heading("cohort", text="Niên khóa")
        self.tree.heading("grad", text="Ngày cấp bằng")
        
        self.tree.column("pos", width=40, anchor=tk.CENTER)
        self.tree.column("stt", width=60, anchor=tk.CENTER)
        self.tree.column("name", width=180, anchor=tk.W)
        self.tree.column("school", width=220, anchor=tk.W)
        self.tree.column("cohort", width=100, anchor=tk.CENTER)
        self.tree.column("grad", width=110, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        context_rows = self.block_info.get("context_rows", [])
        for r in context_rows:
            is_curr = r.get("is_in_block", False)
            pos_icon = "👉" if is_curr else ""
            item_id = self.tree.insert("", tk.END, values=(
                pos_icon,
                r.get("stt", ""),
                r.get("name", ""),
                r.get("school", "") or "(Chưa có)",
                r.get("cohort", ""),
                r.get("grad", "")
            ))
            if is_curr:
                self.tree.tag_configure("highlight", background="#FFF3E0", foreground="#BF360C")
                self.tree.item(item_id, tags=("highlight",))

        act_frame = ttk.LabelFrame(self, text="💡 Lựa chọn Tên trường áp dụng cho khối học sinh này", padding=10)
        act_frame.pack(fill=tk.X, padx=12, pady=8)
        
        if self.default_suggestion:
            sugg_box = ttk.Frame(act_frame)
            sugg_box.pack(fill=tk.X, pady=(0, 6))
            ttk.Label(sugg_box, text="Gợi ý tự động từ bạn liền kề: ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
            ttk.Label(sugg_box, text=f"'{self.default_suggestion}'", font=("Segoe UI", 9, "bold"), foreground="#1B5E20").pack(side=tk.LEFT)
            
            sugg_btn = ttk.Button(sugg_box, text="👉 Áp dụng gợi ý", command=self._apply_suggestion)
            sugg_btn.pack(side=tk.RIGHT)

        select_box = ttk.Frame(act_frame)
        select_box.pack(fill=tk.X, pady=4)
        ttk.Label(select_box, text="Hoặc chọn trường trực thuộc: ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        
        self.school_var = tk.StringVar(value=self.default_suggestion or (self.allowed_schools[0] if self.allowed_schools else ""))
        self.school_combo = ttk.Combobox(select_box, textvariable=self.school_var, values=self.allowed_schools, width=45)
        self.school_combo.pack(side=tk.LEFT, padx=6)
        
        self.remember_var = tk.BooleanVar(value=True)
        rem_chk = ttk.Checkbutton(act_frame, text="Tự động ghi nhớ cho các khối tương tự tiếp theo", variable=self.remember_var)
        rem_chk.pack(anchor=tk.W, pady=(4, 0))

        btn_frame = ttk.Frame(self, padding=(12, 8))
        btn_frame.pack(fill=tk.X)
        
        count = self.block_info.get("count", 1)
        ok_btn = ttk.Button(btn_frame, text=f"✅ Áp dụng cho cả {count} học sinh", command=self._confirm)
        ok_btn.pack(side=tk.RIGHT, padx=4)
        
        skip_btn = ttk.Button(btn_frame, text="⏩ Bỏ qua (Giữ nguyên trống)", command=self._skip)
        skip_btn.pack(side=tk.RIGHT, padx=4)

    def _apply_suggestion(self):
        self.school_var.set(self.default_suggestion)
        self._confirm()

    def _confirm(self):
        val = self.school_var.get().strip()
        if not val:
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn hoặc nhập tên trường!", parent=self)
            return
        self.result = val
        self.apply_to_all_similar = self.remember_var.get()
        self.destroy()

    def _skip(self):
        self.result = ""
        self.apply_to_all_similar = False
        self.destroy()

class CohortBlockDialog(tk.Toplevel):
    """
    Hộp thoại tương tác khi phát hiện Khối học sinh thiếu Niên khóa (Cột O).
    Gợi ý thông minh tính từ Năm sinh của học sinh (Y_ra = Y_birth + 15, Y_vao = Y_ra - 4).
    """
    def __init__(self, parent, block_info: dict, suggested_cohort: str = "", inferred_from_year: int | None = None):
        super().__init__(parent)
        self.title("Xử lý Niên khóa (Cột O)")
        self.transient(parent)
        self.grab_set()
        
        self.block_info = block_info
        self.suggested_cohort = suggested_cohort
        self.inferred_from_year = inferred_from_year
        self.result = None
        self.apply_to_all_similar = False
        
        self._build_ui()
        center_window(self, parent, 850, 520)

    def _build_ui(self):
        hdr_frame = ttk.Frame(self, padding=(12, 10))
        hdr_frame.pack(fill=tk.X)
        
        start_stt = self.block_info.get("start_stt", 1)
        end_stt = self.block_info.get("end_stt", 1)
        count = self.block_info.get("count", 1)
        raw_val = self.block_info.get("raw_val", "(Trống)")
        
        title_lbl = ttk.Label(
            hdr_frame, 
            text=f"⚠️ PHÁT HIỆN KHỐI {count} HỌC SINH CẦN XÁC ĐỊNH NIÊN KHÓA (STT {start_stt} -> {end_stt})",
            font=("Segoe UI", 11, "bold"),
            foreground="#E65100"
        )
        title_lbl.pack(anchor=tk.W)
        
        sub_lbl = ttk.Label(
            hdr_frame,
            text=f"Tệp: {self.block_info.get('file_name', '')} | Sheet: {self.block_info.get('sheet_name', '')} | Giá trị gốc: '{raw_val}'",
            font=("Segoe UI", 9)
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        tbl_frame = ttk.LabelFrame(self, text="📋 Bảng đối chiếu học sinh tham chiếu (5 trước và 5 sau)", padding=6)
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        
        columns = ("pos", "stt", "name", "dob", "cohort", "grad")
        self.tree = ttk.Treeview(tbl_frame, columns=columns, show="headings", height=8)
        
        self.tree.heading("pos", text="")
        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Họ và tên")
        self.tree.heading("dob", text="Ngày sinh")
        self.tree.heading("cohort", text="Niên khóa")
        self.tree.heading("grad", text="Ngày cấp bằng")
        
        self.tree.column("pos", width=40, anchor=tk.CENTER)
        self.tree.column("stt", width=60, anchor=tk.CENTER)
        self.tree.column("name", width=200, anchor=tk.W)
        self.tree.column("dob", width=120, anchor=tk.CENTER)
        self.tree.column("cohort", width=140, anchor=tk.CENTER)
        self.tree.column("grad", width=130, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        for r in self.block_info.get("context_rows", []):
            is_curr = r.get("is_in_block", False)
            pos_icon = "👉" if is_curr else ""
            item_id = self.tree.insert("", tk.END, values=(
                pos_icon,
                r.get("stt", ""),
                r.get("name", ""),
                r.get("dob", "") or "(Chưa có)",
                r.get("cohort", "") or "(Trống)",
                r.get("grad", "") or "(Trống)"
            ))
            if is_curr:
                self.tree.tag_configure("highlight", background="#FFF3E0", foreground="#BF360C")
                self.tree.item(item_id, tags=("highlight",))

        act_frame = ttk.LabelFrame(self, text="💡 Xác định Niên khóa (Dạng YYYY-YYYY, hiệu 2 năm = 4)", padding=10)
        act_frame.pack(fill=tk.X, padx=12, pady=8)
        
        if self.suggested_cohort:
            sugg_box = ttk.Frame(act_frame)
            sugg_box.pack(fill=tk.X, pady=(0, 6))
            source_txt = f"từ năm sinh phổ biến {self.inferred_from_year}" if self.inferred_from_year else "từ học sinh liền kề"
            ttk.Label(sugg_box, text=f"Gợi ý tự động ({source_txt}): ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
            ttk.Label(sugg_box, text=f"'{self.suggested_cohort}'", font=("Segoe UI", 9, "bold"), foreground="#1B5E20").pack(side=tk.LEFT)
            
            sugg_btn = ttk.Button(sugg_box, text="👉 Áp dụng gợi ý", command=self._apply_suggestion)
            sugg_btn.pack(side=tk.RIGHT)

        input_box = ttk.Frame(act_frame)
        input_box.pack(fill=tk.X, pady=4)
        ttk.Label(input_box, text="Nhập Niên khóa (hoặc chỉ gõ Năm ra trường, ví dụ 2020): ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        
        self.cohort_var = tk.StringVar(value=self.suggested_cohort)
        self.cohort_entry = ttk.Entry(input_box, textvariable=self.cohort_var, width=20)
        self.cohort_entry.pack(side=tk.LEFT, padx=6)
        
        self.remember_var = tk.BooleanVar(value=True)
        rem_chk = ttk.Checkbutton(act_frame, text="Tự động ghi nhớ cho các khối tương tự tiếp theo", variable=self.remember_var)
        rem_chk.pack(anchor=tk.W, pady=(4, 0))

        btn_frame = ttk.Frame(self, padding=(12, 8))
        btn_frame.pack(fill=tk.X)
        
        count = self.block_info.get("count", 1)
        ok_btn = ttk.Button(btn_frame, text=f"✅ Áp dụng cho cả {count} học sinh", command=self._confirm)
        ok_btn.pack(side=tk.RIGHT, padx=4)
        
        skip_btn = ttk.Button(btn_frame, text="⏩ Bỏ qua (Giữ nguyên)", command=self._skip)
        skip_btn.pack(side=tk.RIGHT, padx=4)

    def _apply_suggestion(self):
        self.cohort_var.set(self.suggested_cohort)
        self._confirm()

    def _confirm(self):
        val = self.cohort_var.get().strip()
        if not val:
            messagebox.showwarning("Cảnh báo", "Vui lòng nhập Niên khóa hoặc Năm ra trường!", parent=self)
            return
        
        m4 = re.match(r'^(19\d{2}|20\d{2})$', val)
        if m4:
            y_ra = int(m4.group(1))
            val = f"{y_ra - 4}-{y_ra}"
        m2 = re.match(r'^(\d{2})$', val)
        if m2:
            v2 = int(m2.group(1))
            y_ra = 2000 + v2 if v2 <= 50 else 1900 + v2
            val = f"{y_ra - 4}-{y_ra}"

        from src.normalizer import parse_nien_khoa
        is_valid, y1, y2, fmt = parse_nien_khoa(val)
        if is_valid:
            self.result = fmt
            self.apply_to_all_similar = self.remember_var.get()
            self.destroy()
        else:
            if y1 and y2:
                ans = messagebox.askyesno(
                    "Xác nhận",
                    f"Niên khóa '{val}' có hiệu số {y2} - {y1} = {y2 - y1} (khác 4 theo chuẩn THCS).\nBạn có chắc chắn muốn lưu giá trị này không?",
                    parent=self
                )
                if ans:
                    self.result = val
                    self.apply_to_all_similar = self.remember_var.get()
                    self.destroy()
            else:
                messagebox.showerror("Lỗi", f"Niên khóa '{val}' không đúng định dạng YYYY-YYYY!", parent=self)

    def _skip(self):
        self.result = self.block_info.get("raw_val", "")
        self.apply_to_all_similar = False
        self.destroy()

class DateFixDialog(tk.Toplevel):
    """
    Hộp thoại sửa Ngày sinh bị sai định dạng/lôgic.
    Nút 'Đảo Ngày <-> Tháng' chỉ kích hoạt khi cả hai số ngày và tháng đều <= 12 (trường hợp nhập nhằng).
    """
    def __init__(self, parent, student_info: dict, field_name: str = "Ngày sinh", smart_suggestion: str = ""):
        super().__init__(parent)
        self.title(f"Sửa {field_name}")
        self.transient(parent)
        self.grab_set()
        
        self.student_info = student_info
        self.field_name = field_name
        self.smart_suggestion = smart_suggestion
        self.result = None
        self.skip_all = False
        
        self._build_ui()
        center_window(self, parent, 850, 520)

    def _build_ui(self):
        hdr_frame = ttk.Frame(self, padding=(12, 10))
        hdr_frame.pack(fill=tk.X)
        
        err_desc = self.student_info.get("err_desc", "Sai định dạng ngày tháng")
        title_lbl = ttk.Label(
            hdr_frame, 
            text=f"⚠️ PHÁT HIỆN {self.field_name.upper()} CHƯA HỢP LỆ: {err_desc}",
            font=("Segoe UI", 11, "bold"),
            foreground="#C62828"
        )
        title_lbl.pack(anchor=tk.W)
        
        stt = self.student_info.get("stt", "")
        name = self.student_info.get("name", "")
        raw_val = self.student_info.get("raw_val", "(Trống)")
        curr_norm = self.student_info.get("curr_norm", "")
        
        sub_lbl = ttk.Label(
            hdr_frame,
            text=f"Học sinh: {name} (STT: {stt}) | File: {self.student_info.get('file_name', '')} (Sheet: {self.student_info.get('sheet_name', '')})\nGiá trị gốc: '{raw_val}' | Kết quả hiện tại: '{curr_norm}'",
            font=("Segoe UI", 9)
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        tbl_frame = ttk.LabelFrame(self, text="📋 Bảng đối chiếu học sinh tham chiếu (5 trước và 5 sau)", padding=6)
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)
        
        columns = ("pos", "stt", "name", "cohort", "dob", "grad")
        self.tree = ttk.Treeview(tbl_frame, columns=columns, show="headings", height=8)
        
        self.tree.heading("pos", text="")
        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Họ và tên")
        self.tree.heading("cohort", text="Niên khóa")
        self.tree.heading("dob", text="Ngày sinh")
        self.tree.heading("grad", text="Ngày cấp bằng")
        
        self.tree.column("pos", width=40, anchor=tk.CENTER)
        self.tree.column("stt", width=60, anchor=tk.CENTER)
        self.tree.column("name", width=200, anchor=tk.W)
        self.tree.column("cohort", width=120, anchor=tk.CENTER)
        self.tree.column("dob", width=130, anchor=tk.CENTER)
        self.tree.column("grad", width=130, anchor=tk.CENTER)
        
        scrollbar = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        for r in self.student_info.get("context_rows", []):
            is_curr = r.get("is_current", False)
            pos_icon = "👉" if is_curr else ""
            item_id = self.tree.insert("", tk.END, values=(
                pos_icon,
                r.get("stt", ""),
                r.get("name", ""),
                r.get("cohort", "") or "(Trống)",
                r.get("dob", "") or "(Trống)",
                r.get("grad", "") or "(Trống)"
            ))
            if is_curr:
                self.tree.tag_configure("highlight", background="#FFEBEE", foreground="#C62828")
                self.tree.item(item_id, tags=("highlight",))

        act_frame = ttk.LabelFrame(self, text=f"💡 Điều chỉnh {self.field_name}", padding=10)
        act_frame.pack(fill=tk.X, padx=12, pady=6)
        
        input_box = ttk.Frame(act_frame)
        input_box.pack(fill=tk.X, pady=4)
        
        ttk.Label(input_box, text=f"Nhập {self.field_name} chuẩn (dd/MM/yyyy): ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        
        init_val = self.smart_suggestion or curr_norm or ""
        self.date_var = tk.StringVar(value=init_val)
        self.date_entry = ttk.Entry(input_box, textvariable=self.date_var, width=18)
        self.date_entry.pack(side=tk.LEFT, padx=6)
        
        # Nút Đảo Ngày <-> Tháng: Kiểm tra xem có thể đảo được không
        self.swap_btn = ttk.Button(input_box, text="🔄 Đảo Ngày ↔ Tháng", command=self._swap_day_month)
        self.swap_btn.pack(side=tk.LEFT, padx=6)
        self._check_swap_enable()

        self.date_var.trace_add("write", lambda *args: self._check_swap_enable())

        if self.smart_suggestion:
            sugg_btn = ttk.Button(input_box, text=f"👉 Lấy gợi ý ({self.smart_suggestion})", command=self._apply_suggestion)
            sugg_btn.pack(side=tk.LEFT, padx=6)

        btn_frame = ttk.Frame(self, padding=(12, 8))
        btn_frame.pack(fill=tk.X)
        
        ok_btn = ttk.Button(btn_frame, text="✅ Cập nhật", command=self._confirm)
        ok_btn.pack(side=tk.RIGHT, padx=4)
        
        keep_btn = ttk.Button(btn_frame, text="⏩ Giữ nguyên (Ghi nhận Chưa đầy đủ)", command=self._keep)
        keep_btn.pack(side=tk.RIGHT, padx=4)
        
        skip_all_btn = ttk.Button(btn_frame, text="⏭️ Bỏ qua tất cả các lỗi sau", command=self._skip_all)
        skip_all_btn.pack(side=tk.RIGHT, padx=4)

    def _check_swap_enable(self):
        val = self.date_var.get().strip()
        m = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$', val)
        if m:
            a, b, _ = map(int, m.groups())
            # Nút sáng khi có đủ ngày tháng
            self.swap_btn.config(state=tk.NORMAL)
        else:
            self.swap_btn.config(state=tk.DISABLED)

    def _swap_day_month(self):
        val = self.date_var.get().strip()
        m = re.match(r'^(\d{1,2})[/.-](\d{1,2})[/.-](\d{4})$', val)
        if m:
            d, m_val, y = m.groups()
            self.date_var.set(f"{int(m_val):02d}/{int(d):02d}/{y}")

    def _apply_suggestion(self):
        self.date_var.set(self.smart_suggestion)

    def _confirm(self):
        val = self.date_var.get().strip()
        from src.normalizer import format_date_to_text, check_dob_validity
        formatted = format_date_to_text(val)
        is_v, err_t, err_d = check_dob_validity(formatted)
        if is_v:
            self.result = formatted
            self.destroy()
        else:
            ans = messagebox.askyesno(
                "Xác nhận",
                f"Giá trị '{formatted}' vẫn chưa hợp lệ: {err_d}.\nBạn có muốn giữ nguyên giá trị này để ghi nhận trạng thái 'Số hoá chưa đầy đủ, chính xác' không?",
                parent=self
            )
            if ans:
                self.result = formatted
                self.destroy()

    def _keep(self):
        self.result = self.student_info.get("curr_norm", "")
        self.destroy()

    def _skip_all(self):
        self.result = self.student_info.get("curr_norm", "")
        self.skip_all = True
        self.destroy()

class GradDateBlockDialog(tk.Toplevel):
    """
    Hộp thoại tương tác khi phát hiện cả nhóm (Trường, Niên khóa) chưa có Ngày cấp bằng chuẩn.
    Cho phép nhập 1 ngày và áp dụng cho toàn bộ học sinh cùng nhóm (Trường, Niên khóa).
    """
    def __init__(self, parent, group_info: dict, default_date: str = ""):
        super().__init__(parent)
        self.title("Xử lý Ngày cấp bằng (Cột S)")
        self.transient(parent)
        self.grab_set()

        self.group_info = group_info
        self.default_date = default_date
        self.result = None
        self.apply_to_group = True

        self._build_ui()
        center_window(self, parent, 820, 480)

    def _build_ui(self):
        hdr_frame = ttk.Frame(self, padding=(12, 10))
        hdr_frame.pack(fill=tk.X)

        sch = self.group_info.get("school", "Chưa rõ trường")
        nk = self.group_info.get("cohort", "Chưa rõ niên khóa")
        count = self.group_info.get("count", 1)

        title_lbl = ttk.Label(
            hdr_frame,
            text=f"⚠️ XÁC ĐỊNH NGÀY CẤP BẰNG CHO NHÓM: {sch} (Niên khóa: {nk})",
            font=("Segoe UI", 11, "bold"),
            foreground="#1565C0"
        )
        title_lbl.pack(anchor=tk.W)

        sub_lbl = ttk.Label(
            hdr_frame,
            text=f"Nhóm gồm {count} học sinh cùng trường và cùng niên khóa hiện chưa có ngày cấp bằng chuẩn dd/MM/yyyy.",
            font=("Segoe UI", 9)
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        tbl_frame = ttk.LabelFrame(self, text="📋 Bảng đối chiếu học sinh tham chiếu", padding=6)
        tbl_frame.pack(fill=tk.BOTH, expand=True, padx=12, pady=4)

        columns = ("stt", "name", "school", "cohort", "dob", "grad")
        self.tree = ttk.Treeview(tbl_frame, columns=columns, show="headings", height=6)

        self.tree.heading("stt", text="STT")
        self.tree.heading("name", text="Họ và tên")
        self.tree.heading("school", text="Tên trường")
        self.tree.heading("cohort", text="Niên khóa")
        self.tree.heading("dob", text="Ngày sinh")
        self.tree.heading("grad", text="Ngày cấp gốc")

        self.tree.column("stt", width=50, anchor=tk.CENTER)
        self.tree.column("name", width=180, anchor=tk.W)
        self.tree.column("school", width=180, anchor=tk.W)
        self.tree.column("cohort", width=100, anchor=tk.CENTER)
        self.tree.column("dob", width=110, anchor=tk.CENTER)
        self.tree.column("grad", width=110, anchor=tk.CENTER)

        scrollbar = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        for r in self.group_info.get("sample_students", []):
            self.tree.insert("", tk.END, values=(
                r.get("clean_stt", ""),
                r.get("name", ""),
                r.get("school", ""),
                r.get("cohort", ""),
                r.get("dob", ""),
                r.get("grad_date", "") or "(Trống)"
            ))

        act_frame = ttk.LabelFrame(self, text="💡 Nhập Ngày tháng năm cấp bằng chuẩn", padding=10)
        act_frame.pack(fill=tk.X, padx=12, pady=6)

        in_box = ttk.Frame(act_frame)
        in_box.pack(fill=tk.X)

        ttk.Label(in_box, text="Ngày cấp bằng chuẩn (dd/MM/yyyy): ", font=("Segoe UI", 9)).pack(side=tk.LEFT)
        self.grad_var = tk.StringVar(value=self.default_date)
        self.grad_entry = ttk.Entry(in_box, textvariable=self.grad_var, width=18)
        self.grad_entry.pack(side=tk.LEFT, padx=6)

        self.apply_chk_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(act_frame, text=f"Tự động áp dụng ngày này cho TẤT CẢ {count} học sinh cùng trường & khóa này", variable=self.apply_chk_var).pack(anchor=tk.W, pady=(6, 0))

        btn_frame = ttk.Frame(self, padding=(12, 8))
        btn_frame.pack(fill=tk.X)

        ok_btn = ttk.Button(btn_frame, text="✅ Cập nhật", command=self._confirm)
        ok_btn.pack(side=tk.RIGHT, padx=4)

        skip_btn = ttk.Button(btn_frame, text="⏩ Bỏ qua (Để trống)", command=self._skip)
        skip_btn.pack(side=tk.RIGHT, padx=4)

    def _confirm(self):
        val = self.grad_var.get().strip()
        from src.normalizer import format_grad_date_to_text, is_valid_grad_date
        formatted = format_grad_date_to_text(val)
        if is_valid_grad_date(formatted) and formatted:
            self.result = formatted
            self.apply_to_group = self.apply_chk_var.get()
            self.destroy()
        else:
            ans = messagebox.askyesno("Cảnh báo", f"Giá trị '{val}' chưa đúng định dạng chuẩn ngày cấp bằng dd/MM/yyyy.\nBạn có muốn giữ nguyên không?", parent=self)
            if ans:
                self.result = val
                self.apply_to_group = self.apply_chk_var.get()
                self.destroy()

    def _skip(self):
        self.result = ""
        self.destroy()

class DataInspectorDialog(tk.Toplevel):
    """
    Màn hình Tra cứu & Đối soát Dữ liệu gốc (In-App Data Lineage & Inspector).
    Cho phép tìm kiếm nhanh học sinh và đối chiếu song song: [Dữ liệu gốc] vs [Dữ liệu chuẩn hóa].
    """
    def __init__(self, parent, all_students: list[dict]):
        super().__init__(parent)
        self.title("🔍 Tra cứu & Đối soát Dữ liệu gốc (Data Inspector)")
        self.geometry("950, 650")
        self.transient(parent)

        self.all_students = all_students
        self.filtered_students = list(all_students)

        self._build_ui()
        center_window(self, parent, 950, 650)

    def _build_ui(self):
        # 1. Top Search Bar
        top_frame = ttk.Frame(self, padding=(12, 10))
        top_frame.pack(fill=tk.X)

        ttk.Label(top_frame, text="🔍 Tìm kiếm nhanh: ", font=("Segoe UI", 9, "bold")).pack(side=tk.LEFT)
        self.search_var = tk.StringVar()
        self.search_var.trace_add("write", self._on_search)
        search_entry = ttk.Entry(top_frame, textvariable=self.search_var, width=35)
        search_entry.pack(side=tk.LEFT, padx=6)

        # Filter radio buttons
        self.filter_var = tk.StringVar(value="all")
        ttk.Radiobutton(top_frame, text="Tất cả", variable=self.filter_var, value="all", command=self._on_search).pack(side=tk.LEFT, padx=4)
        ttk.Radiobutton(top_frame, text="Chỉ xem học sinh lỗi (Cột Y != Đã số hóa)", variable=self.filter_var, value="errors", command=self._on_search).pack(side=tk.LEFT, padx=4)

        self.count_lbl = ttk.Label(top_frame, text=f"Tổng: {len(self.all_students)} học sinh", font=("Segoe UI", 9, "bold"), foreground="#0D47A1")
        self.count_lbl.pack(side=tk.RIGHT)

        # Split Pane
        paned = ttk.PanedWindow(self, orient=tk.VERTICAL)
        paned.pack(fill=tk.BOTH, expand=True, padx=12, pady=6)

        # Upper Table: Student List
        tbl_frame = ttk.Frame(paned)
        paned.add(tbl_frame, weight=3)

        columns = ("idx", "name", "dob", "school", "cohort", "status", "file", "sheet", "row")
        self.tree = ttk.Treeview(tbl_frame, columns=columns, show="headings", height=8)

        self.tree.heading("idx", text="STT")
        self.tree.heading("name", text="Họ và tên")
        self.tree.heading("dob", text="Ngày sinh")
        self.tree.heading("school", text="Tên trường")
        self.tree.heading("cohort", text="Niên khóa")
        self.tree.heading("status", text="Trạng thái số hóa")
        self.tree.heading("file", text="Tệp gốc")
        self.tree.heading("sheet", text="Sheet")
        self.tree.heading("row", text="Dòng")

        self.tree.column("idx", width=45, anchor=tk.CENTER)
        self.tree.column("name", width=160, anchor=tk.W)
        self.tree.column("dob", width=95, anchor=tk.CENTER)
        self.tree.column("school", width=180, anchor=tk.W)
        self.tree.column("cohort", width=90, anchor=tk.CENTER)
        self.tree.column("status", width=160, anchor=tk.W)
        self.tree.column("file", width=120, anchor=tk.W)
        self.tree.column("sheet", width=80, anchor=tk.W)
        self.tree.column("row", width=50, anchor=tk.CENTER)

        s_bar = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=s_bar.set)
        s_bar.pack(side=tk.RIGHT, fill=tk.Y)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.tree.bind("<<TreeviewSelect>>", self._on_select_student)

        # Lower Detail: Side-by-side Inspector
        detail_frame = ttk.LabelFrame(paned, text="📄 Chi tiết đối soát: [Dữ liệu gốc ban đầu] vs [Dữ liệu sau chuẩn hóa]", padding=10)
        paned.add(detail_frame, weight=4)

        diff_cols = ("field", "raw_val", "norm_val")
        self.diff_tree = ttk.Treeview(detail_frame, columns=diff_cols, show="headings", height=7)
        self.diff_tree.heading("field", text="Trường thông tin")
        self.diff_tree.heading("raw_val", text="📄 Dữ liệu gốc trong File nguồn")
        self.diff_tree.heading("norm_val", text="✨ Kết quả sau chuẩn hóa")

        self.diff_tree.column("field", width=220, anchor=tk.W)
        self.diff_tree.column("raw_val", width=340, anchor=tk.W)
        self.diff_tree.column("norm_val", width=340, anchor=tk.W)

        diff_scrollbar = ttk.Scrollbar(detail_frame, orient=tk.VERTICAL, command=self.diff_tree.yview)
        self.diff_tree.configure(yscrollcommand=diff_scrollbar.set)
        diff_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.diff_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Bottom Bar
        b_frame = ttk.Frame(self, padding=(12, 8))
        b_frame.pack(fill=tk.X)
        ttk.Button(b_frame, text="Đóng", command=self.destroy).pack(side=tk.RIGHT)

        self._populate_tree()

    def _populate_tree(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, st in enumerate(self.filtered_students, start=1):
            stt = st.get("clean_stt", idx)
            name = st.get("name", "")
            dob = st.get("dob", "")
            sch = st.get("school", "")
            nk = st.get("cohort", "")
            final_r = st.get("final_row", [])
            status = final_r[24] if len(final_r) > 24 else ""
            f_name = st.get("file_name", "")
            s_name = st.get("sheet_name", "")
            r_idx = st.get("row_idx", "")

            item_id = self.tree.insert("", tk.END, values=(
                stt, name, dob, sch, nk, status, f_name, s_name, r_idx
            ))

            if status == "Chưa số hoá":
                self.tree.tag_configure("red", foreground="#C62828", font=("Segoe UI", 9, "bold"))
                self.tree.item(item_id, tags=("red",))
            elif status == "Số hoá chưa đầy đủ, chính xác":
                self.tree.tag_configure("orange", foreground="#E65100")
                self.tree.item(item_id, tags=("orange",))

        self.count_lbl.config(text=f"Hiển thị: {len(self.filtered_students)}/{len(self.all_students)} học sinh")

    def _on_search(self, *args):
        kw = self.search_var.get().strip().lower()
        mode = self.filter_var.get()

        res = []
        for st in self.all_students:
            final_r = st.get("final_row", [])
            status = final_r[24] if len(final_r) > 24 else ""
            if mode == "errors" and status == "Đã số hóa đầy đủ, chính xác":
                continue

            if not kw:
                res.append(st)
            else:
                combined = f"{st.get('name', '')} {st.get('school', '')} {st.get('cohort', '')} {st.get('clean_stt', '')} {st.get('file_name', '')}".lower()
                if kw in combined:
                    res.append(st)

        self.filtered_students = res
        self._populate_tree()

    def _on_select_student(self, event=None):
        sel = self.tree.selection()
        if not sel:
            return
        idx = self.tree.index(sel[0])
        if idx >= len(self.filtered_students):
            return
        st = self.filtered_students[idx]

        for item in self.diff_tree.get_children():
            self.diff_tree.delete(item)

        raw_r = st.get("row_values", [])
        final_r = st.get("final_row", [])

        fields = [
            ("Họ, chữ đệm và tên (Cột E)", st.get("raw_name", ""), st.get("name", "")),
            ("Ngày, tháng, năm sinh (Cột I)", st.get("raw_dob", ""), st.get("dob", "")),
            ("Giới tính (Cột J)", raw_r[9] if len(raw_r) > 9 else "", final_r[9] if len(final_r) > 9 else ""),
            ("Dân tộc (Cột K)", raw_r[10] if len(raw_r) > 10 else "", final_r[10] if len(final_r) > 10 else ""),
            ("Quốc tịch (Cột L)", raw_r[11] if len(raw_r) > 11 else "", final_r[11] if len(final_r) > 11 else ""),
            ("Nơi sinh (Cột M)", raw_r[12] if len(raw_r) > 12 else "", final_r[12] if len(final_r) > 12 else ""),
            ("Tên trường đã học (Cột N)", st.get("raw_school", ""), st.get("school", "")),
            ("Niên khóa (Cột O)", st.get("raw_cohort", ""), st.get("cohort", "")),
            ("Ngày cấp bằng (Cột S)", st.get("raw_grad", ""), st.get("grad_date", "")),
            ("Trạng thái số hóa (Cột Y)", "(Chưa xác định)", final_r[24] if len(final_r) > 24 else ""),
            ("Nguyên nhân lỗi (Cột Z)", "(Không có)", final_r[25] if len(final_r) > 25 and final_r[25] else "(Trống)"),
            ("Vị trí nguồn gốc", f"File: {st.get('file_name', '')} | Sheet: {st.get('sheet_name', '')} | Dòng: {st.get('row_idx', '')}", "Đã gắn vào Cột X (Ghi chú)")
        ]

        for f_name, r_val, n_val in fields:
            self.diff_tree.insert("", tk.END, values=(
                f_name,
                r_val if r_val is not None and str(r_val).strip() != "" else "(Trống)",
                n_val if n_val is not None and str(n_val).strip() != "" else "(Trống)"
            ))
