from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Any, Callable

from src.gui.theme import (
    FONT_LABEL,
    FONT_LABEL_BOLD,
    FONT_BUTTON,
    FONT_TITLE,
    CLR_PRIMARY,
    CLR_SUCCESS,
    CLR_WARNING,
    CLR_DANGER,
    CLR_MUTED
)

class StudentEditDialog(tk.Toplevel):
    """
    Hộp thoại chỉnh sửa thông tin chi tiết của 1 học sinh.
    Tự động suy thông tin hành chính khi đổi Tên trường và tự động recompute Trạng thái / Lỗi.
    """
    def __init__(self, parent, student_data: dict, db_manager: Any, lookup_mgr: Any = None, on_saved: Callable[[], None] | None = None):
        super().__init__(parent)
        self.title(f"Chỉnh sửa thông tin học sinh: {student_data.get('ho_ten', '')} (ID: {student_data.get('id', '')})")
        self.geometry("900x680")
        self.minsize(800, 600)
        self.transient(parent)
        self.grab_set()

        self.student_data = dict(student_data)
        self.student_id = student_data["id"]
        self.db_manager = db_manager
        self.lookup_mgr = lookup_mgr
        self.on_saved = on_saved

        self.entries: dict[str, tk.StringVar] = {}
        self._build_ui()
        self._center_window(parent)

    def _center_window(self, parent):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = self.winfo_width()
        h = self.winfo_height()
        x = max(0, px + (pw - w) // 2)
        y = max(0, py + (ph - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # Header banner
        hdr = ttk.Frame(self, padding=(16, 12), style="Card.TFrame")
        hdr.pack(fill=tk.X)

        title_lbl = ttk.Label(
            hdr,
            text=f"✏️ CHỈNH SỬA HỒ SƠ HỌC SINH (ID: {self.student_id})",
            font=FONT_TITLE,
            foreground=CLR_PRIMARY
        )
        title_lbl.pack(side=tk.LEFT)

        sub_lbl = ttk.Label(
            hdr,
            text="Hệ thống tự động cập nhật Địa bàn khi đổi Trường và tính lại Trạng thái số hóa",
            font=FONT_LABEL,
            foreground=CLR_MUTED
        )
        sub_lbl.pack(side=tk.RIGHT)

        # Main scrollable canvas container
        container = ttk.Frame(self, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        canvas = tk.Canvas(container, highlightthickness=0)
        scrollbar = ttk.Scrollbar(container, orient=tk.VERTICAL, command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.bind('<Configure>', lambda event: canvas.itemconfig(canvas_window, width=event.width))

        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # -------------------------------------------------------------
        # KHUNG 1: THÔNG TIN CÁ NHÂN NGƯỜI HỌC
        # -------------------------------------------------------------
        grp1 = ttk.LabelFrame(scrollable_frame, text="1. Thông tin cá nhân người học", padding=10)
        grp1.pack(fill=tk.X, pady=(0, 10))

        r1 = ttk.Frame(grp1)
        r1.pack(fill=tk.X, pady=3)
        self._add_field(r1, "ho_ten", "Họ và tên (*):", 0, 0, width=28, bold=True)
        self._add_field(r1, "ngay_sinh", "Ngày sinh (dd/MM/yyyy):", 0, 2, width=16)
        self._add_combobox(r1, "gioi_tinh", "Giới tính:", ["Nam", "Nữ"], 0, 4, width=10)

        r2 = ttk.Frame(grp1)
        r2.pack(fill=tk.X, pady=3)
        self._add_field(r2, "dan_toc", "Dân tộc:", 0, 0, width=16)
        self._add_field(r2, "quoc_tich", "Quốc tịch:", 0, 2, width=16)
        self._add_field(r2, "noi_sinh", "Nơi sinh:", 0, 4, width=22)

        r3 = ttk.Frame(grp1)
        r3.pack(fill=tk.X, pady=3)
        self._add_field(r3, "cccd", "Số CCCD (12 số):", 0, 0, width=22)
        self._add_field(r3, "ma_nguoi_hoc", "Mã người học:", 0, 2, width=18)
        self._add_field(r3, "ho_chieu", "Hộ chiếu:", 0, 4, width=18)

        # -------------------------------------------------------------
        # KHUNG 2: THÔNG TIN TRƯỜNG HỌC & TỐT NGHIỆP
        # -------------------------------------------------------------
        grp2 = ttk.LabelFrame(scrollable_frame, text="2. Thông tin trường học & Tốt nghiệp", padding=10)
        grp2.pack(fill=tk.X, pady=(0, 10))

        r4 = ttk.Frame(grp2)
        r4.pack(fill=tk.X, pady=3)
        
        # Danh sách gợi ý trường học từ LookupManager
        school_options = []
        if self.lookup_mgr:
            school_options = self.lookup_mgr.get_all_schools()
        
        self._add_combobox_or_entry(r4, "ten_truong", "Tên trường (*):", school_options, 0, 0, width=32, bold=True)
        self._add_field(r4, "nien_khoa", "Niên khóa (YYYY-YYYY):", 0, 2, width=16)
        self._add_field(r4, "diem_thi", "Điểm thi:", 0, 4, width=10)

        r5 = ttk.Frame(grp2)
        r5.pack(fill=tk.X, pady=3)
        self._add_field(r5, "so_hieu_bang", "Số hiệu bằng:", 0, 0, width=20)
        self._add_field(r5, "so_vao_so", "Số vào sổ gốc:", 0, 2, width=20)
        self._add_field(r5, "ngay_cap_bang", "Ngày cấp bằng:", 0, 4, width=16)

        r6 = ttk.Frame(grp2)
        r6.pack(fill=tk.X, pady=3)
        self._add_field(r6, "hoi_dong_thi", "Hội đồng thi:", 0, 0, width=28)
        self._add_field(r6, "tinh_trang_van_bang", "Tình trạng văn bằng:", 0, 2, width=18)
        self._add_field(r6, "ghi_chu", "Ghi chú:", 0, 4, width=22)

        # -------------------------------------------------------------
        # KHUNG 3: ĐƠN VỊ & CƠ QUAN CẤP BẰNG (TỰ ĐỘNG THEO TRƯỜNG)
        # -------------------------------------------------------------
        grp3 = ttk.LabelFrame(scrollable_frame, text="3. Cơ quan cấp bằng & Địa bàn hành chính", padding=10)
        grp3.pack(fill=tk.X, pady=(0, 10))

        r7 = ttk.Frame(grp3)
        r7.pack(fill=tk.X, pady=3)
        self._add_field(r7, "dia_danh", "Địa danh trụ sở:", 0, 0, width=25)
        self._add_field(r7, "co_quan_cap_bang", "Cơ quan cấp bằng:", 0, 2, width=28)

        r8 = ttk.Frame(grp3)
        r8.pack(fill=tk.X, pady=3)
        self._add_field(r8, "chuc_danh_nguoi_ky", "Chức danh người ký:", 0, 0, width=25)
        self._add_field(r8, "nguoi_ky", "Họ tên người ký:", 0, 2, width=28)

        r9 = ttk.Frame(grp3)
        r9.pack(fill=tk.X, pady=3)
        self._add_field(r9, "tinh_cu", "Tỉnh cũ:", 0, 0, width=16)
        self._add_field(r9, "huyen_cu", "Huyện cũ:", 0, 2, width=16)
        self._add_field(r9, "xa_phuong", "Xã/phường mới:", 0, 4, width=18)

        r10 = ttk.Frame(grp3)
        r10.pack(fill=tk.X, pady=3)
        self._add_field(r10, "ten_don_vi", "Tên đơn vị quản lý:", 0, 0, width=32)
        self._add_field(r10, "ma_don_vi", "Mã đơn vị:", 0, 2, width=16)

        # -------------------------------------------------------------
        # KHUNG 4: TRẠNG THÁI & NGUỒN GỐC (HỆ THỐNG KIỂM SOÁT)
        # -------------------------------------------------------------
        grp4 = ttk.LabelFrame(scrollable_frame, text="4. Trạng thái số hóa & Vết dữ liệu gốc", padding=10)
        grp4.pack(fill=tk.X, pady=(0, 10))

        r11 = ttk.Frame(grp4)
        r11.pack(fill=tk.X, pady=3)

        status_val = self.student_data.get("trang_thai", "")
        stat_color = CLR_SUCCESS if "Đã số hóa" in status_val else (CLR_WARNING if "chưa đầy đủ" in status_val else CLR_DANGER)
        
        ttk.Label(r11, text="Trạng thái hiện tại: ", font=FONT_LABEL_BOLD).pack(side=tk.LEFT)
        self.status_preview_lbl = ttk.Label(r11, text=status_val or "Chưa xác định", font=FONT_LABEL_BOLD, foreground=stat_color)
        self.status_preview_lbl.pack(side=tk.LEFT, padx=(0, 20))

        ttk.Label(r11, text="Lý do lỗi / Cảnh báo: ", font=FONT_LABEL_BOLD).pack(side=tk.LEFT)
        self.reason_preview_lbl = ttk.Label(
            r11,
            text=self.student_data.get("ly_do_loi") or "(Không có lỗi)",
            font=FONT_LABEL,
            foreground=CLR_DANGER if self.student_data.get("ly_do_loi") else CLR_MUTED
        )
        self.reason_preview_lbl.pack(side=tk.LEFT)

        r12 = ttk.Frame(grp4)
        r12.pack(fill=tk.X, pady=(4, 0))
        src_info = f"Tệp: {self.student_data.get('source_file', '')} | Sheet: {self.student_data.get('source_sheet', '')} | Dòng gốc: {self.student_data.get('source_row', '')}"
        ttk.Label(r12, text=f"📍 Nguồn gốc: {src_info}", font=FONT_LABEL, foreground=CLR_MUTED).pack(side=tk.LEFT)

        # -------------------------------------------------------------
        # BOTTOM ACTION BUTTONS
        # -------------------------------------------------------------
        bot_bar = ttk.Frame(self, padding=(16, 10))
        bot_bar.pack(side=tk.BOTTOM, fill=tk.X)

        save_btn = ttk.Button(bot_bar, text="💾 Lưu thay đổi", style="Action.TButton", command=self._on_save)
        save_btn.pack(side=tk.RIGHT, padx=6)

        cancel_btn = ttk.Button(bot_bar, text="Hủy bỏ", command=self.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=6)

    def _add_field(self, parent, key: str, label: str, row: int, col: int, width: int = 20, bold: bool = False):
        frame = ttk.Frame(parent)
        frame.pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)

        lbl = ttk.Label(frame, text=label, font=FONT_LABEL_BOLD if bold else FONT_LABEL)
        lbl.pack(anchor=tk.W)

        var = tk.StringVar(value=str(self.student_data.get(key) or ''))
        self.entries[key] = var
        entry = ttk.Entry(frame, textvariable=var, width=width)
        entry.pack(fill=tk.X, expand=True, pady=(2, 0))

    def _add_combobox(self, parent, key: str, label: str, values: list[str], row: int, col: int, width: int = 15):
        frame = ttk.Frame(parent)
        frame.pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)

        lbl = ttk.Label(frame, text=label, font=FONT_LABEL)
        lbl.pack(anchor=tk.W)

        var = tk.StringVar(value=str(self.student_data.get(key) or ''))
        self.entries[key] = var
        cb = ttk.Combobox(frame, textvariable=var, values=values, width=width, state="readonly")
        cb.pack(fill=tk.X, expand=True, pady=(2, 0))

    def _add_combobox_or_entry(self, parent, key: str, label: str, values: list[str], row: int, col: int, width: int = 25, bold: bool = False):
        frame = ttk.Frame(parent)
        frame.pack(side=tk.LEFT, padx=6, fill=tk.X, expand=True)

        lbl = ttk.Label(frame, text=label, font=FONT_LABEL_BOLD if bold else FONT_LABEL)
        lbl.pack(anchor=tk.W)

        var = tk.StringVar(value=str(self.student_data.get(key) or ''))
        self.entries[key] = var
        if values:
            cb = ttk.Combobox(frame, textvariable=var, values=values, width=width)
            cb.pack(fill=tk.X, expand=True, pady=(2, 0))
            cb.bind("<<ComboboxSelected>>", self._on_school_changed)
        else:
            entry = ttk.Entry(frame, textvariable=var, width=width)
            entry.pack(fill=tk.X, expand=True, pady=(2, 0))

    def _on_school_changed(self, event=None):
        """Khi người dùng chọn trường khác từ Combobox, tự động điền các thông tin hành chính."""
        sch_name = self.entries.get("ten_truong", tk.StringVar()).get().strip()
        if sch_name and self.lookup_mgr:
            details = self.lookup_mgr.get_school_details(sch_name)
            if details:
                if "tinh_cu" in self.entries and details.get("province_name"):
                    self.entries["tinh_cu"].set(details["province_name"])
                if "huyen_cu" in self.entries and details.get("district_name"):
                    self.entries["huyen_cu"].set(details["district_name"])
                if "xa_phuong" in self.entries and details.get("commune"):
                    self.entries["xa_phuong"].set(details["commune"])
                if "ten_don_vi" in self.entries and details.get("unit_name"):
                    self.entries["ten_don_vi"].set(details["unit_name"])
                if "ma_don_vi" in self.entries and details.get("unit_code"):
                    self.entries["ma_don_vi"].set(details["unit_code"])

    def _on_save(self):
        ho_ten = self.entries.get("ho_ten", tk.StringVar()).get().strip()
        if not ho_ten:
            messagebox.showwarning("Thiếu thông tin", "Họ và tên học sinh không được để trống!", parent=self)
            return

        fields_dict = {k: v.get().strip() for k, v in self.entries.items()}

        try:
            ok = self.db_manager.update_student(self.student_id, fields_dict, self.lookup_mgr)
            if ok:
                if self.on_saved:
                    self.on_saved()
                messagebox.showinfo("Thành công", f"Đã cập nhật thông tin học sinh: {ho_ten}", parent=self)
                self.destroy()
            else:
                messagebox.showerror("Lỗi", "Không thể cập nhật thông tin học sinh vào CSDL.", parent=self)
        except Exception as e:
            messagebox.showerror("Lỗi CSDL", f"Đã xảy ra lỗi khi lưu: {e}", parent=self)


class BatchEditDialog(tk.Toplevel):
    """
    Hộp thoại chỉnh sửa thông tin hàng loạt cho nhiều học sinh được chọn.
    """
    def __init__(self, parent, student_ids: list[int], db_manager: Any, lookup_mgr: Any = None, on_saved: Callable[[], None] | None = None):
        super().__init__(parent)
        self.title(f"Sửa đổi hàng loạt: {len(student_ids)} học sinh được chọn")
        self.geometry("640x520")
        self.minsize(560, 460)
        self.transient(parent)
        self.grab_set()

        self.student_ids = student_ids
        self.db_manager = db_manager
        self.lookup_mgr = lookup_mgr
        self.on_saved = on_saved

        self.checks: dict[str, tk.BooleanVar] = {}
        self.vars: dict[str, tk.StringVar] = {}

        self._build_ui()
        self._center_window(parent)

    def _center_window(self, parent):
        self.update_idletasks()
        pw = parent.winfo_width()
        ph = parent.winfo_height()
        px = parent.winfo_rootx()
        py = parent.winfo_rooty()
        w = self.winfo_width()
        h = self.winfo_height()
        x = max(0, px + (pw - w) // 2)
        y = max(0, py + (ph - h) // 2)
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        # Header
        hdr = ttk.Frame(self, padding=(16, 12), style="Card.TFrame")
        hdr.pack(fill=tk.X)

        title_lbl = ttk.Label(
            hdr,
            text=f"⚡ SỬA ĐỔI HÀNG LOẠT ({len(self.student_ids)} HỌC SINH)",
            font=FONT_TITLE,
            foreground=CLR_PRIMARY
        )
        title_lbl.pack(anchor=tk.W)

        sub_lbl = ttk.Label(
            hdr,
            text="Tích chọn các trường cần gán đồng loạt. Các trường không tích chọn sẽ giữ nguyên giá trị gốc.",
            font=FONT_LABEL,
            foreground=CLR_MUTED
        )
        sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # Main content
        body = ttk.Frame(self, padding=16)
        body.pack(fill=tk.BOTH, expand=True)

        school_list = self.lookup_mgr.get_all_schools() if self.lookup_mgr else []

        # 1. Tên trường
        self._add_batch_field(body, "ten_truong", "Tên trường đã học:", is_combo=True, combo_values=school_list, width=32)

        # 2. Niên khóa
        self._add_batch_field(body, "nien_khoa", "Niên khóa (YYYY-YYYY):", width=20)

        # 3. Ngày cấp bằng
        self._add_batch_field(body, "ngay_cap_bang", "Ngày cấp bằng (dd/MM/yyyy):", width=20)

        # 4. Hội đồng thi
        self._add_batch_field(body, "hoi_dong_thi", "Hội đồng thi:", width=28)

        # 5. Cơ quan cấp bằng
        self._add_batch_field(body, "co_quan_cap_bang", "Tên cơ quan cấp bằng:", width=28)

        # 6. Người ký bằng
        self._add_batch_field(body, "nguoi_ky", "Họ tên người ký bằng:", width=28)

        # 7. Chức danh người ký
        self._add_batch_field(body, "chuc_danh_nguoi_ky", "Chức danh người ký:", width=24)

        # Bottom buttons
        bot = ttk.Frame(self, padding=(16, 12))
        bot.pack(side=tk.BOTTOM, fill=tk.X)

        apply_btn = ttk.Button(bot, text="⚡ Áp dụng cho các học sinh đã chọn", style="Action.TButton", command=self._on_apply)
        apply_btn.pack(side=tk.RIGHT, padx=6)

        cancel_btn = ttk.Button(bot, text="Hủy bỏ", command=self.destroy)
        cancel_btn.pack(side=tk.RIGHT, padx=6)

    def _add_batch_field(self, parent, key: str, label_text: str, is_combo: bool = False, combo_values: list[str] = None, width: int = 20):
        row = ttk.Frame(parent)
        row.pack(fill=tk.X, pady=6)

        chk_var = tk.BooleanVar(value=False)
        self.checks[key] = chk_var

        chk = ttk.Checkbutton(row, text=label_text, variable=chk_var, width=28)
        chk.pack(side=tk.LEFT)

        val_var = tk.StringVar()
        self.vars[key] = val_var

        if is_combo and combo_values:
            entry = ttk.Combobox(row, textvariable=val_var, values=combo_values, width=width)
        else:
            entry = ttk.Entry(row, textvariable=val_var, width=width)
        entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        # Tự động bật check khi người dùng gõ vào ô
        val_var.trace_add("write", lambda *args: chk_var.set(True) if val_var.get().strip() else None)

    def _on_apply(self):
        updates = {}
        for key, chk_var in self.checks.items():
            if chk_var.get():
                val = self.vars[key].get().strip()
                updates[key] = val

        if not updates:
            messagebox.showwarning("Chưa chọn trường sửa", "Vui lòng tích chọn ít nhất 1 trường thông tin để cập nhật hàng loạt!", parent=self)
            return

        summary = "\n".join([f"• {k}: {v}" for k, v in updates.items()])
        confirm = messagebox.askyesno(
            "Xác nhận cập nhật hàng loạt",
            f"Bạn có chắc chắn muốn cập nhật đồng loạt cho {len(self.student_ids)} học sinh với các giá trị sau không?\n\n{summary}\n\nThao tác này sẽ ghi trực tiếp vào CSDL SQLite.",
            parent=self
        )
        if not confirm:
            return

        try:
            count = self.db_manager.batch_update_students(self.student_ids, updates, self.lookup_mgr)
            if self.on_saved:
                self.on_saved()
            messagebox.showinfo("Thành công", f"Đã cập nhật thành công cho {count} học sinh!", parent=self)
            self.destroy()
        except Exception as e:
            messagebox.showerror("Lỗi CSDL", f"Đã xảy ra lỗi khi cập nhật hàng loạt: {e}", parent=self)
