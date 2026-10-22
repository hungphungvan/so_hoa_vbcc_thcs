from __future__ import annotations

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Any

from src.gui.student_edit_dialog import StudentEditDialog, BatchEditDialog
from src.gui.export_dialog import ExportDialog
from src.gui.theme import (
    FONT_LABEL,
    FONT_LABEL_BOLD,
    FONT_DATA_BOLD,
    FONT_TREE_BOLD,
    CLR_PRIMARY,
    CLR_WARNING,
    CLR_DANGER,
    CLR_MUTED,
    CLR_TEXT
)

class DataManagementView(ttk.Frame):
    """
    Giao diện Phân hệ 2: Lọc, Tìm kiếm & Sửa đổi dữ liệu CSDL SQLite.
    Bao gồm:
    1. Thanh tìm kiếm tức thì theo Họ tên & Ngày sinh (Cách 2).
    2. Bộ lọc đa tiêu chí 2 nhánh (Lịch sử: Tỉnh cũ -> Huyện cũ -> Trường -> Khóa; Địa bàn mới: Xã -> Đơn vị).
    3. Bảng dữ liệu phân trang hiển thị đầy đủ 28 cột chuẩn của Bộ/Sở.
    4. Thanh điều hướng phân trang (50/100/200 dòng) và công cụ sửa đơn / sửa hàng loạt.
    """
    def __init__(self, parent, db_manager: Any, lookup_mgr: Any = None):
        super().__init__(parent, style="Window.TFrame")
        self.db_manager = db_manager
        self.lookup_mgr = lookup_mgr

        # Trạng thái phân trang & sắp xếp
        self.current_page = 1
        self.page_size = 50
        self.total_records = 0
        self.total_pages = 1
        self.sort_by = "id"
        self.sort_order = "ASC"

        # Biến tìm kiếm & lọc
        self.search_name_var = tk.StringVar()
        self.search_dob_var = tk.StringVar()
        self.debounce_timer = None

        self.tinh_cu_var = tk.StringVar(value="Tất cả")
        self.huyen_cu_var = tk.StringVar(value="Tất cả")
        self.ten_truong_var = tk.StringVar(value="Tất cả")
        self.nien_khoa_var = tk.StringVar(value="Tất cả")

        self.xa_phuong_var = tk.StringVar(value="Tất cả")
        self.ten_don_vi_var = tk.StringVar(value="Tất cả")

        self.trang_thai_var = tk.StringVar(value="Tất cả")
        self.chi_xem_loi_var = tk.BooleanVar(value=False)

        self._build_ui()
        self.refresh_filter_options()
        self.refresh_data(reset_page=True)

    def _build_ui(self):
        # ==============================================================
        # 1. KHU VỰC TÌM KIẾM TỨC THÌ (CÁCH 2: 2 Ô NHẬP HỌ TÊN & NGÀY SINH)
        # ==============================================================
        search_card = ttk.LabelFrame(
            self,
            text="1. Tìm kiếm tức thì theo Học sinh (Bắn trúng đích)",
            style="Card.TLabelframe",
            padding=(10, 6)
        )
        search_card.pack(fill=tk.X, padx=10, pady=(6, 4))

        s_row = ttk.Frame(search_card)
        s_row.pack(fill=tk.X)

        ttk.Label(s_row, text="👤 Họ và tên:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 4))
        self.entry_name = ttk.Entry(s_row, textvariable=self.search_name_var, width=28)
        self.entry_name.pack(side=tk.LEFT, padx=(0, 16))
        self.entry_name.bind("<Return>", lambda e: self._on_quick_search())
        self.search_name_var.trace_add("write", self._on_search_typing)

        ttk.Label(s_row, text="📅 Ngày sinh:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 4))
        self.entry_dob = ttk.Entry(s_row, textvariable=self.search_dob_var, width=18)
        self.entry_dob.pack(side=tk.LEFT, padx=(0, 12))
        self.entry_dob.bind("<Return>", lambda e: self._on_quick_search())
        self.search_dob_var.trace_add("write", self._on_search_typing)

        btn_search = ttk.Button(s_row, text="🔍 Tìm ngay", command=self._on_quick_search)
        btn_search.pack(side=tk.LEFT, padx=4)

        btn_clear_search = ttk.Button(s_row, text="❌ Xóa tìm kiếm", command=self._clear_search)
        btn_clear_search.pack(side=tk.LEFT, padx=4)

        self.search_tip_lbl = ttk.Label(
            s_row,
            text="Gợi ý: Nhập tên (vd: Tuấn) hoặc ngày sinh (vd: 15/08/1990 hoặc 1990)",
            font=FONT_LABEL,
            foreground=CLR_MUTED
        )
        self.search_tip_lbl.pack(side=tk.RIGHT)

        # ==============================================================
        # 2. KHU VỰC BỘ LỌC ĐA TIÊU CHÍ (2 NHÁNH)
        # ==============================================================
        filter_card = ttk.LabelFrame(
            self,
            text="2. Bộ lọc đa tiêu chí (Gom nhóm & Quản trị hồ sơ)",
            style="Card.TLabelframe",
            padding=(10, 6)
        )
        filter_card.pack(fill=tk.X, padx=10, pady=(0, 4))

        # Nhánh 1: Lịch sử cấp bằng (Tỉnh cũ -> Huyện cũ -> Trường -> Niên khóa)
        f_row1 = ttk.Frame(filter_card)
        f_row1.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(f_row1, text="🏛️ Tỉnh cũ:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_tinh_cu = ttk.Combobox(f_row1, textvariable=self.tinh_cu_var, width=14, state="readonly")
        self.cb_tinh_cu.pack(side=tk.LEFT, padx=(0, 10))
        self.cb_tinh_cu.bind("<<ComboboxSelected>>", self._on_tinh_cu_changed)

        ttk.Label(f_row1, text="➔ Huyện cũ:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_huyen_cu = ttk.Combobox(f_row1, textvariable=self.huyen_cu_var, width=16, state="readonly")
        self.cb_huyen_cu.pack(side=tk.LEFT, padx=(0, 10))
        self.cb_huyen_cu.bind("<<ComboboxSelected>>", self._on_huyen_cu_changed)

        ttk.Label(f_row1, text="➔ Tên trường:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_ten_truong = ttk.Combobox(f_row1, textvariable=self.ten_truong_var, width=24)
        self.cb_ten_truong.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Label(f_row1, text="Niên khóa:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_nien_khoa = ttk.Combobox(f_row1, textvariable=self.nien_khoa_var, width=12, state="readonly")
        self.cb_nien_khoa.pack(side=tk.LEFT)

        # Nhánh 2: Địa bàn mới & Đơn vị + Trạng thái
        f_row2 = ttk.Frame(filter_card)
        f_row2.pack(fill=tk.X, pady=(2, 0))

        ttk.Label(f_row2, text="📍 Xã/phường mới:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_xa_phuong = ttk.Combobox(f_row2, textvariable=self.xa_phuong_var, width=16, state="readonly")
        self.cb_xa_phuong.pack(side=tk.LEFT, padx=(0, 10))
        self.cb_xa_phuong.bind("<<ComboboxSelected>>", self._on_xa_phuong_changed)

        ttk.Label(f_row2, text="➔ Đơn vị quản lý:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_ten_don_vi = ttk.Combobox(f_row2, textvariable=self.ten_don_vi_var, width=24)
        self.cb_ten_don_vi.pack(side=tk.LEFT, padx=(0, 14))

        ttk.Label(f_row2, text="Trạng thái:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_trang_thai = ttk.Combobox(
            f_row2,
            textvariable=self.trang_thai_var,
            values=["Tất cả", "Đã số hóa đầy đủ, chính xác", "Số hoá chưa đầy đủ, chính xác", "Chưa số hoá"],
            width=20,
            state="readonly"
        )
        self.cb_trang_thai.pack(side=tk.LEFT, padx=(0, 10))

        chk_err = ttk.Checkbutton(f_row2, text="⚠️ Chỉ xem hồ sơ có lỗi/cảnh báo", variable=self.chi_xem_loi_var, command=self._on_apply_filter)
        chk_err.pack(side=tk.LEFT, padx=(0, 14))

        btn_apply = ttk.Button(f_row2, text="🔍 Áp dụng lọc", style="Action.TButton", command=self._on_apply_filter)
        btn_apply.pack(side=tk.LEFT, padx=4)

        btn_reset = ttk.Button(f_row2, text="🔄 Đặt lại", command=self._on_reset_filter)
        btn_reset.pack(side=tk.LEFT, padx=4)

        self.filter_stat_lbl = ttk.Label(
            f_row2,
            text="Tìm thấy: 0 học sinh",
            font=FONT_DATA_BOLD,
            foreground=CLR_PRIMARY
        )
        self.filter_stat_lbl.pack(side=tk.RIGHT)

        # ==============================================================
        # 3. BẢNG DỮ LIỆU GRID PHÂN TRANG (28 CỘT CHUẨN)
        # ==============================================================
        grid_card = ttk.Frame(self, style="Card.TFrame", padding=(10, 4))
        grid_card.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 4))

        # Cấu hình danh sách cột Treeview
        self.columns = [
            ("id", "ID", 50, tk.CENTER),
            ("stt", "STT", 45, tk.CENTER),
            ("ho_ten", "Họ và tên", 160, tk.W),
            ("ngay_sinh", "Ngày sinh", 95, tk.CENTER),
            ("gioi_tinh", "Giới tính", 70, tk.CENTER),
            ("ten_truong", "Tên trường đã học", 180, tk.W),
            ("nien_khoa", "Niên khóa", 90, tk.CENTER),
            ("so_hieu_bang", "Số hiệu bằng", 100, tk.CENTER),
            ("so_vao_so", "Số vào sổ", 85, tk.CENTER),
            ("trang_thai", "Trạng thái số hóa", 160, tk.W),
            ("ly_do_loi", "Lý do lỗi / Cảnh báo", 180, tk.W),
            ("dan_toc", "Dân tộc", 75, tk.CENTER),
            ("noi_sinh", "Nơi sinh", 110, tk.W),
            ("cccd", "Số CCCD", 115, tk.CENTER),
            ("diem_thi", "Điểm thi", 70, tk.CENTER),
            ("hoi_dong_thi", "Hội đồng thi", 140, tk.W),
            ("dia_danh", "Địa danh trụ sở", 130, tk.W),
            ("ngay_cap_bang", "Ngày cấp bằng", 95, tk.CENTER),
            ("co_quan_cap_bang", "Cơ quan cấp bằng", 160, tk.W),
            ("chuc_danh_nguoi_ky", "Chức danh ký", 110, tk.W),
            ("nguoi_ky", "Người ký bằng", 130, tk.W),
            ("tinh_trang_van_bang", "Tình trạng", 90, tk.CENTER),
            ("ghi_chu", "Ghi chú", 110, tk.W),
            ("tinh_cu", "Tỉnh cũ", 100, tk.W),
            ("huyen_cu", "Huyện cũ", 110, tk.W),
            ("xa_phuong", "Xã/phường", 110, tk.W),
            ("ten_don_vi", "Tên đơn vị", 160, tk.W),
            ("ma_don_vi", "Mã đơn vị", 80, tk.CENTER),
            ("source_file", "Tệp nguồn gốc", 130, tk.W),
            ("source_sheet", "Sheet nguồn", 90, tk.W)
        ]

        col_ids = [c[0] for c in self.columns]
        self.tree = ttk.Treeview(grid_card, columns=col_ids, show="headings", selectmode="extended")

        for col_id, col_text, col_width, col_anchor in self.columns:
            self.tree.heading(col_id, text=col_text, command=lambda cid=col_id: self._on_sort_column(cid))
            self.tree.column(col_id, width=col_width, minwidth=40, anchor=col_anchor)

        # Thanh cuộn dọc & ngang
        scroll_y = ttk.Scrollbar(grid_card, orient=tk.VERTICAL, command=self.tree.yview)
        scroll_x = ttk.Scrollbar(grid_card, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)

        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Sự kiện Treeview
        self.tree.bind("<Double-1>", lambda e: self._on_edit_selected())
        self.tree.bind("<F2>", lambda e: self._on_edit_selected())
        self.tree.bind("<<TreeviewSelect>>", self._on_row_selected)

        # Tags tô màu trạng thái
        self.tree.tag_configure("da_so_hoa", foreground=CLR_TEXT)
        self.tree.tag_configure("chua_day_du", foreground=CLR_WARNING, font=FONT_TREE_BOLD)
        self.tree.tag_configure("chua_so_hoa", foreground=CLR_DANGER, font=FONT_TREE_BOLD)

        # ==============================================================
        # 4. THANH PHÂN TRANG & THANH TÁC VỤ (PAGINATION & ACTION BAR)
        # ==============================================================
        action_bar = ttk.Frame(self, style="Window.TFrame", padding=(10, 4))
        action_bar.pack(side=tk.BOTTOM, fill=tk.X)

        # Bên trái: Điều hướng phân trang
        nav_frame = ttk.Frame(action_bar)
        nav_frame.pack(side=tk.LEFT)

        self.btn_first = ttk.Button(nav_frame, text="|<<", width=4, command=lambda: self._go_to_page(1))
        self.btn_first.pack(side=tk.LEFT, padx=1)

        self.btn_prev = ttk.Button(nav_frame, text="<", width=3, command=lambda: self._go_to_page(self.current_page - 1))
        self.btn_prev.pack(side=tk.LEFT, padx=1)

        self.page_info_lbl = ttk.Label(nav_frame, text="Trang 1 / 1", font=FONT_LABEL_BOLD)
        self.page_info_lbl.pack(side=tk.LEFT, padx=8)

        self.btn_next = ttk.Button(nav_frame, text=">", width=3, command=lambda: self._go_to_page(self.current_page + 1))
        self.btn_next.pack(side=tk.LEFT, padx=1)

        self.btn_last = ttk.Button(nav_frame, text=">>|", width=4, command=lambda: self._go_to_page(self.total_pages))
        self.btn_last.pack(side=tk.LEFT, padx=1)

        ttk.Label(nav_frame, text="Hiển thị:", font=FONT_LABEL).pack(side=tk.LEFT, padx=(12, 4))
        self.cb_page_size = ttk.Combobox(nav_frame, values=["50", "100", "200"], width=5, state="readonly")
        self.cb_page_size.set("50")
        self.cb_page_size.pack(side=tk.LEFT)
        self.cb_page_size.bind("<<ComboboxSelected>>", self._on_page_size_changed)
        ttk.Label(nav_frame, text="dòng/trang", font=FONT_LABEL).pack(side=tk.LEFT, padx=(4, 0))

        # Bên phải: Nút tác vụ
        act_frame = ttk.Frame(action_bar)
        act_frame.pack(side=tk.RIGHT)

        self.btn_edit_single = ttk.Button(
            act_frame,
            text="✏️ Sửa dòng (F2)",
            style="Action.TButton",
            state=tk.DISABLED,
            command=self._on_edit_selected
        )
        self.btn_edit_single.pack(side=tk.LEFT, padx=3)

        self.btn_edit_batch = ttk.Button(
            act_frame,
            text="⚡ Sửa hàng loạt (0)",
            state=tk.DISABLED,
            command=self._on_edit_batch
        )
        self.btn_edit_batch.pack(side=tk.LEFT, padx=3)

        self.btn_delete = ttk.Button(
            act_frame,
            text="🗑️ Xóa đã chọn",
            state=tk.DISABLED,
            command=self._on_delete_selected
        )
        self.btn_delete.pack(side=tk.LEFT, padx=3)

        self.btn_export = ttk.Button(
            act_frame,
            text="📤 Xuất Excel...",
            style="Action.TButton",
            command=self._on_open_export_dialog
        )
        self.btn_export.pack(side=tk.LEFT, padx=3)

        btn_refresh = ttk.Button(act_frame, text="🔄 Tải lại", command=lambda: self.refresh_data(reset_page=False))
        btn_refresh.pack(side=tk.LEFT, padx=3)

    # ------------------------------------------------------------------
    # QUẢN LÝ TƯƠNG TÁC LỌC ĐA CẤP (CASCADING FILTERS)
    # ------------------------------------------------------------------
    def refresh_filter_options(self):
        """Nạp danh sách các tùy chọn bộ lọc thực tế từ CSDL SQLite."""
        try:
            tinh = self.tinh_cu_var.get()
            huyen = self.huyen_cu_var.get()
            xa = self.xa_phuong_var.get()

            opts = self.db_manager.get_cascade_filter_options(
                tinh_cu=tinh if tinh != "Tất cả" else None,
                huyen_cu=huyen if huyen != "Tất cả" else None,
                xa_phuong=xa if xa != "Tất cả" else None
            )

            # 1. Tỉnh cũ
            self.cb_tinh_cu["values"] = ["Tất cả"] + opts.get("tinh_cu", [])
            # 2. Huyện cũ
            self.cb_huyen_cu["values"] = ["Tất cả"] + opts.get("huyen_cu", [])
            # 3. Tên trường
            self.cb_ten_truong["values"] = ["Tất cả"] + opts.get("ten_truong", [])
            # 4. Niên khóa
            self.cb_nien_khoa["values"] = ["Tất cả"] + opts.get("nien_khoa", [])
            # 5. Xã/phường
            self.cb_xa_phuong["values"] = ["Tất cả"] + opts.get("xa_phuong", [])
            # 6. Đơn vị
            self.cb_ten_don_vi["values"] = ["Tất cả"] + opts.get("ten_don_vi", [])
        except Exception:
            pass

    def _on_tinh_cu_changed(self, event=None):
        self.huyen_cu_var.set("Tất cả")
        self.ten_truong_var.set("Tất cả")
        self.refresh_filter_options()
        self._on_apply_filter()

    def _on_huyen_cu_changed(self, event=None):
        self.ten_truong_var.set("Tất cả")
        self.refresh_filter_options()
        self._on_apply_filter()

    def _on_xa_phuong_changed(self, event=None):
        self.ten_don_vi_var.set("Tất cả")
        self.refresh_filter_options()
        self._on_apply_filter()

    def _on_search_typing(self, *args):
        """Cơ chế Debounce 300ms khi người dùng gõ vào ô tìm kiếm Họ tên hoặc Ngày sinh."""
        if self.debounce_timer:
            self.after_cancel(self.debounce_timer)
        self.debounce_timer = self.after(300, self._on_quick_search)

    def _on_quick_search(self):
        self.refresh_data(reset_page=True)

    def _clear_search(self):
        self.search_name_var.set("")
        self.search_dob_var.set("")
        self.refresh_data(reset_page=True)

    def _on_apply_filter(self):
        self.refresh_data(reset_page=True)

    def _on_reset_filter(self):
        self.search_name_var.set("")
        self.search_dob_var.set("")
        self.tinh_cu_var.set("Tất cả")
        self.huyen_cu_var.set("Tất cả")
        self.ten_truong_var.set("Tất cả")
        self.nien_khoa_var.set("Tất cả")
        self.xa_phuong_var.set("Tất cả")
        self.ten_don_vi_var.set("Tất cả")
        self.trang_thai_var.set("Tất cả")
        self.chi_xem_loi_var.set(False)
        self.refresh_filter_options()
        self.refresh_data(reset_page=True)

    # ------------------------------------------------------------------
    # PHÂN TRANG & HIỂN THỊ DỮ LIỆU
    # ------------------------------------------------------------------
    def refresh_data(self, reset_page: bool = False):
        if reset_page:
            self.current_page = 1

        filters = {
            "tinh_cu": self.tinh_cu_var.get(),
            "huyen_cu": self.huyen_cu_var.get(),
            "ten_truong": self.ten_truong_var.get(),
            "nien_khoa": self.nien_khoa_var.get(),
            "xa_phuong": self.xa_phuong_var.get(),
            "ten_don_vi": self.ten_don_vi_var.get(),
            "trang_thai": self.trang_thai_var.get(),
            "chi_xem_loi": self.chi_xem_loi_var.get()
        }

        try:
            rows, total = self.db_manager.query_students_paginated(
                search_name=self.search_name_var.get(),
                search_dob=self.search_dob_var.get(),
                filters=filters,
                page=self.current_page,
                page_size=self.page_size,
                sort_by=self.sort_by,
                sort_order=self.sort_order
            )
        except Exception as e:
            messagebox.showerror("Lỗi truy vấn", f"Không thể lấy dữ liệu: {e}", parent=self)
            return

        self.total_records = total
        self.total_pages = max(1, (total + self.page_size - 1) // self.page_size)
        if self.current_page > self.total_pages:
            self.current_page = self.total_pages

        # Xóa dữ liệu cũ trên Treeview
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Đổ dữ liệu mới vào bảng
        col_keys = [c[0] for c in self.columns]
        for row in rows:
            values = [row.get(k, '') for k in col_keys]
            
            # Phân loại tag màu
            st_val = str(row.get("trang_thai") or "")
            if "chưa đầy đủ" in st_val.lower():
                tag = "chua_day_du"
            elif "chưa số" in st_val.lower():
                tag = "chua_so_hoa"
            else:
                tag = "da_so_hoa"

            self.tree.insert("", tk.END, iid=str(row["id"]), values=values, tags=(tag,))

        # Cập nhật UI phân trang
        self.page_info_lbl.config(text=f"Trang {self.current_page} / {self.total_pages}")
        self.filter_stat_lbl.config(text=f"Tìm thấy: {self.total_records:,} học sinh")

        self.btn_first.config(state=tk.NORMAL if self.current_page > 1 else tk.DISABLED)
        self.btn_prev.config(state=tk.NORMAL if self.current_page > 1 else tk.DISABLED)
        self.btn_next.config(state=tk.NORMAL if self.current_page < self.total_pages else tk.DISABLED)
        self.btn_last.config(state=tk.NORMAL if self.current_page < self.total_pages else tk.DISABLED)

        self._update_action_buttons()

    def _go_to_page(self, page_num: int):
        target = max(1, min(page_num, self.total_pages))
        if target != self.current_page:
            self.current_page = target
            self.refresh_data(reset_page=False)

    def _on_page_size_changed(self, event=None):
        try:
            self.page_size = int(self.cb_page_size.get())
        except ValueError:
            self.page_size = 50
        self.refresh_data(reset_page=True)

    def _on_sort_column(self, col_id: str):
        if self.sort_by == col_id:
            self.sort_order = "DESC" if self.sort_order == "ASC" else "ASC"
        else:
            self.sort_by = col_id
            self.sort_order = "ASC"

        # Đặt lại text heading để hiển thị mũi tên
        for c_id, c_text, _, _ in self.columns:
            if c_id == self.sort_by:
                arrow = " ▲" if self.sort_order == "ASC" else " ▼"
                self.tree.heading(c_id, text=f"{c_text}{arrow}")
            else:
                self.tree.heading(c_id, text=c_text)

        self.refresh_data(reset_page=True)

    def _on_row_selected(self, event=None):
        self._update_action_buttons()

    def _update_action_buttons(self):
        selected = self.tree.selection()
        count = len(selected)

        self.btn_edit_single.config(state=tk.NORMAL if count == 1 else tk.DISABLED)
        self.btn_edit_batch.config(
            state=tk.NORMAL if count > 1 else tk.DISABLED,
            text=f"⚡ Sửa hàng loạt ({count})"
        )
        self.btn_delete.config(state=tk.NORMAL if count > 0 else tk.DISABLED)

    # ------------------------------------------------------------------
    # THAO TÁC SỬA / XÓA DỮ LIỆU
    # ------------------------------------------------------------------
    def _on_edit_selected(self):
        selected = self.tree.selection()
        if not selected:
            return

        student_id = int(selected[0])
        student_data = self.db_manager.get_student_by_id(student_id)
        if not student_data:
            messagebox.showwarning("Không tìm thấy", "Không tìm thấy hồ sơ học sinh trong CSDL.", parent=self)
            return

        dialog = StudentEditDialog(
            parent=self.winfo_toplevel(),
            student_data=student_data,
            db_manager=self.db_manager,
            lookup_mgr=self.lookup_mgr,
            on_saved=lambda: self.refresh_data(reset_page=False)
        )

    def _on_edit_batch(self):
        selected = self.tree.selection()
        if len(selected) < 2:
            messagebox.showinfo("Chọn học sinh", "Vui lòng chọn từ 2 học sinh trở lên để sửa hàng loạt (giữ phím Shift hoặc Ctrl/Cmd khi chọn).", parent=self)
            return

        student_ids = [int(s_id) for s_id in selected]
        dialog = BatchEditDialog(
            parent=self.winfo_toplevel(),
            student_ids=student_ids,
            db_manager=self.db_manager,
            lookup_mgr=self.lookup_mgr,
            on_saved=lambda: self.refresh_data(reset_page=False)
        )

    def _on_delete_selected(self):
        selected = self.tree.selection()
        if not selected:
            return

        count = len(selected)
        confirm = messagebox.askyesno(
            "Xác nhận xóa dữ liệu",
            f"Bạn có chắc chắn muốn xóa vĩnh viễn {count} bản ghi học sinh đã chọn khỏi CSDL SQLite không?\n\nThao tác này KHÔNG thể hoàn tác!",
            parent=self
        )
        if not confirm:
            return

        student_ids = [int(s_id) for s_id in selected]
        try:
            deleted_count = self.db_manager.delete_students(student_ids)
            messagebox.showinfo("Đã xóa", f"Đã xóa thành công {deleted_count} học sinh khỏi CSDL.", parent=self)
            self.refresh_data(reset_page=False)
        except Exception as e:
            messagebox.showerror("Lỗi xóa", f"Không thể xóa dữ liệu: {e}", parent=self)

    def _on_open_export_dialog(self):
        """Mở Hộp thoại Xuất Báo Cáo & Dữ Liệu Excel (Phân hệ 3)."""
        current_filters = {
            "tinh_cu": self.tinh_cu_var.get(),
            "huyen_cu": self.huyen_cu_var.get(),
            "ten_truong": self.ten_truong_var.get(),
            "nien_khoa": self.nien_khoa_var.get(),
            "xa_phuong": self.xa_phuong_var.get(),
            "ten_don_vi": self.ten_don_vi_var.get(),
            "trang_thai": self.trang_thai_var.get(),
            "chi_xem_loi": self.chi_xem_loi_var.get()
        }
        dialog = ExportDialog(
            parent=self.winfo_toplevel(),
            db_manager=self.db_manager,
            lookup_mgr=self.lookup_mgr,
            initial_filters=current_filters
        )
