from __future__ import annotations

import os
import sys
import threading
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Any

from src.export_manager import ExportManager, analyze_cohorts
from src.gui.theme import (
    FONT_LABEL,
    FONT_LABEL_BOLD,
    FONT_TITLE,
    FONT_SUBTITLE,
    CLR_PRIMARY,
    CLR_SUCCESS,
    CLR_WARNING,
    CLR_DANGER,
    CLR_MUTED,
    CLR_TEXT
)

def open_directory(path: Path | str):
    """Mở thư mục trên hệ điều hành (macOS Finder, Windows Explorer, Linux xdg-open)."""
    p = Path(path)
    if not p.exists():
        return
    try:
        if sys.platform == "darwin":
            subprocess.Popen(["open", str(p)])
        elif sys.platform == "win32":
            os.startfile(str(p))
        else:
            subprocess.Popen(["xdg-open", str(p)])
    except Exception as e:
        print(f"Không thể mở thư mục: {e}")

class ExportDialog(tk.Toplevel):
    """
    Hộp thoại Xuất Báo Cáo & Dữ Liệu Excel VBCC THCS (Phân hệ 3).
    Hỗ trợ:
    1. Bộ lọc địa bàn 2 nhánh (Huyện cũ -> Trường; Xã/phường -> Đơn vị) với tự động nhận diện cấp độ.
    2. Bộ chọn Niên khóa tương tác hỗ trợ Click (1 năm), Shift+Click (liên tiếp), Ctrl+Click (rời rạc).
    3. Tự động áp dụng quy tắc tạo file dữ liệu và file Danh sách tài liệu số hóa theo đúng 3 tình huống.
    4. Tùy chọn xuất file lỗi kèm 3 cột vết nguồn gốc.
    5. Luồng xử lý nền (Background Thread) với thanh Progressbar và nút Mở thư mục kết quả.
    """
    def __init__(
        self,
        parent: Any,
        db_manager: Any,
        lookup_mgr: Any = None,
        initial_filters: dict[str, Any] | None = None,
        default_output_dir: Path | str | None = None
    ):
        super().__init__(parent)
        self.title("Xuất Báo Cáo & Dữ Liệu Excel VBCC THCS")
        self.geometry("920x760")
        self.minsize(860, 680)
        self.transient(parent)
        self.grab_set()

        self.db_manager = db_manager
        self.lookup_mgr = lookup_mgr
        self.export_mgr = ExportManager(lookup_mgr=lookup_mgr)
        self.initial_filters = initial_filters or {}
        self.is_exporting = False
        self.last_exported_dir: Path | None = None

        # Thư mục lưu mặc định
        if default_output_dir:
            self.output_dir_var = tk.StringVar(value=str(Path(default_output_dir).resolve()))
        else:
            default_dir = Path(__file__).resolve().parent.parent.parent / "Xuat_Bao_Cao"
            self.output_dir_var = tk.StringVar(value=str(default_dir))

        # Biến trạng thái bộ lọc
        self.branch_var = tk.StringVar(value="branch_history")
        self.huyen_cu_var = tk.StringVar(value="Tất cả")
        self.ten_truong_var = tk.StringVar(value="Tất cả")
        self.xa_phuong_var = tk.StringVar(value="Tất cả")
        self.ten_don_vi_var = tk.StringVar(value="Tất cả")

        # Tùy chọn xuất
        self.export_errors_var = tk.BooleanVar(value=True)
        self.renumber_stt_var = tk.BooleanVar(value=True)
        self.format_text_var = tk.BooleanVar(value=True)

        self._build_ui()
        self._apply_initial_filters()
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
        # 1. Header Banner
        hdr = ttk.Frame(self, padding=(16, 10), style="Card.TFrame")
        hdr.pack(fill=tk.X)

        ttk.Label(
            hdr,
            text="📤 XUẤT BÁO CÁO & DỮ LIỆU EXCEL VBCC THCS",
            font=FONT_TITLE,
            foreground=CLR_PRIMARY
        ).pack(anchor=tk.W)

        ttk.Label(
            hdr,
            text="Xuất biểu mẫu chuẩn 28 cột của Bộ/Sở, tệp Danh sách tài liệu số hóa kèm và Danh sách lỗi",
            font=FONT_SUBTITLE,
            foreground=CLR_MUTED
        ).pack(anchor=tk.W, pady=(2, 0))

        # Scrollable / Main Container
        main_frame = ttk.Frame(self, padding=(14, 8), style="Window.TFrame")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # ==============================================================
        # CARD 1: BỘ LỌC ĐỊA BÀN VÀ ĐƠN VỊ (2 NHÁNH)
        # ==============================================================
        card1 = ttk.LabelFrame(
            main_frame,
            text="1. Bộ lọc địa bàn & đơn vị (Tự động nhận diện cấp xuất)",
            style="Card.TLabelframe",
            padding=(12, 8)
        )
        card1.pack(fill=tk.X, pady=(0, 6))

        # Nhánh 1: Lịch sử (Huyện cũ -> Trường)
        r1_frame = ttk.Frame(card1)
        r1_frame.pack(fill=tk.X, pady=(0, 4))

        rb1 = ttk.Radiobutton(
            r1_frame,
            text="Theo Lịch sử Cấp bằng:",
            variable=self.branch_var,
            value="branch_history",
            command=self._on_branch_changed
        )
        rb1.pack(side=tk.LEFT, padx=(0, 8))

        ttk.Label(r1_frame, text="🏛️ Huyện cũ:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_huyen_cu = ttk.Combobox(r1_frame, textvariable=self.huyen_cu_var, width=18, state="readonly")
        self.cb_huyen_cu.pack(side=tk.LEFT, padx=(0, 12))
        self.cb_huyen_cu.bind("<<ComboboxSelected>>", self._on_huyen_cu_changed)

        ttk.Label(r1_frame, text="🏫 Trường:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_ten_truong = ttk.Combobox(r1_frame, textvariable=self.ten_truong_var, width=28, state="readonly")
        self.cb_ten_truong.pack(side=tk.LEFT)
        self.cb_ten_truong.bind("<<ComboboxSelected>>", self._on_ten_truong_changed)

        # Nhánh 2: Địa bàn mới (Xã/phường -> Đơn vị)
        r2_frame = ttk.Frame(card1)
        r2_frame.pack(fill=tk.X, pady=(2, 4))

        rb2 = ttk.Radiobutton(
            r2_frame,
            text="Theo Địa bàn mới:",
            variable=self.branch_var,
            value="branch_new_admin",
            command=self._on_branch_changed
        )
        rb2.pack(side=tk.LEFT, padx=(0, 28))

        ttk.Label(r2_frame, text="📍 Xã/phường:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_xa_phuong = ttk.Combobox(r2_frame, textvariable=self.xa_phuong_var, width=18, state="readonly")
        self.cb_xa_phuong.pack(side=tk.LEFT, padx=(0, 12))
        self.cb_xa_phuong.bind("<<ComboboxSelected>>", self._on_xa_phuong_changed)

        ttk.Label(r2_frame, text="🏢 Đơn vị:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 2))
        self.cb_ten_don_vi = ttk.Combobox(r2_frame, textvariable=self.ten_don_vi_var, width=28, state="readonly")
        self.cb_ten_don_vi.pack(side=tk.LEFT)
        self.cb_ten_don_vi.bind("<<ComboboxSelected>>", self._on_ten_don_vi_changed)

        # Badge hiển thị cấp xuất đã nhận diện
        badge_frame = ttk.Frame(card1)
        badge_frame.pack(fill=tk.X, pady=(4, 0))

        self.level_badge_lbl = ttk.Label(
            badge_frame,
            text="🏷️ Đang xác định cấp xuất...",
            font=FONT_LABEL_BOLD,
            foreground=CLR_PRIMARY
        )
        self.level_badge_lbl.pack(side=tk.LEFT)

        # ==============================================================
        # CARD 2: CHỌN NIÊN KHÓA (CLICK / SHIFT / CTRL)
        # ==============================================================
        card2 = ttk.LabelFrame(
            main_frame,
            text="2. Chọn Niên khóa (Click: 1 năm | Shift+Click: Liên tiếp | Ctrl+Click: Rời rạc)",
            style="Card.TLabelframe",
            padding=(12, 8)
        )
        card2.pack(fill=tk.X, pady=(0, 6))

        c2_top = ttk.Frame(card2)
        c2_top.pack(fill=tk.X, pady=(0, 4))

        ttk.Label(c2_top, text="📅 Danh sách niên khóa khả dụng:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT)

        btn_select_all = ttk.Button(c2_top, text="Chọn tất cả", width=12, command=self._select_all_cohorts)
        btn_select_all.pack(side=tk.RIGHT, padx=2)

        btn_clear_cohorts = ttk.Button(c2_top, text="Bỏ chọn", width=10, command=self._clear_cohorts)
        btn_clear_cohorts.pack(side=tk.RIGHT, padx=2)

        # Listbox chọn nhiều niên khóa
        lb_frame = ttk.Frame(card2)
        lb_frame.pack(fill=tk.X, pady=(0, 4))

        self.cohort_listbox = tk.Listbox(
            lb_frame,
            selectmode=tk.EXTENDED,
            height=5,
            font=FONT_LABEL,
            exportselection=False
        )
        lb_scroll = ttk.Scrollbar(lb_frame, orient=tk.VERTICAL, command=self.cohort_listbox.yview)
        self.cohort_listbox.configure(yscrollcommand=lb_scroll.set)

        self.cohort_listbox.pack(side=tk.LEFT, fill=tk.X, expand=True)
        lb_scroll.pack(side=tk.RIGHT, fill=tk.Y)

        self.cohort_listbox.bind("<<ListboxSelect>>", self._on_cohort_selection_changed)

        # Nhãn phản hồi trực quan theo 3 tình huống
        self.cohort_feedback_lbl = ttk.Label(
            card2,
            text="⚠️ Chưa chọn niên khóa nào (Vui lòng chọn ít nhất 1 năm)",
            font=FONT_LABEL_BOLD,
            foreground=CLR_WARNING
        )
        self.cohort_feedback_lbl.pack(anchor=tk.W, pady=(2, 0))

        # ==============================================================
        # CARD 3: TÙY CHỌN TỆP XUẤT
        # ==============================================================
        card3 = ttk.LabelFrame(
            main_frame,
            text="3. Tùy chọn biểu mẫu & xử lý dữ liệu",
            style="Card.TLabelframe",
            padding=(12, 6)
        )
        card3.pack(fill=tk.X, pady=(0, 6))

        c3_row = ttk.Frame(card3)
        c3_row.pack(fill=tk.X)

        chk_err = ttk.Checkbutton(
            c3_row,
            text="⚠️ Xuất kèm tệp Danh sách hồ sơ bị Lỗi / Cảnh báo (nếu có lỗi)",
            variable=self.export_errors_var
        )
        chk_err.pack(side=tk.LEFT, padx=(0, 16))

        chk_renumber = ttk.Checkbutton(
            c3_row,
            text="Đánh lại STT liên tục (1 .. N) theo thứ tự: STT ➔ Trường ➔ Niên khóa",
            variable=self.renumber_stt_var
        )
        chk_renumber.pack(side=tk.LEFT, padx=(0, 16))

        chk_format = ttk.Checkbutton(
            c3_row,
            text="Ép kiểu Text (@) bảo toàn số 0 ở đầu",
            variable=self.format_text_var
        )
        chk_format.pack(side=tk.LEFT)

        # ==============================================================
        # CARD 4: THƯ MỤC LƯU & TIẾN TRÌNH THỰC HIỆN
        # ==============================================================
        card4 = ttk.LabelFrame(
            main_frame,
            text="4. Thư mục lưu trữ & Tiến trình xuất",
            style="Card.TLabelframe",
            padding=(12, 8)
        )
        card4.pack(fill=tk.X, pady=(0, 6))

        path_row = ttk.Frame(card4)
        path_row.pack(fill=tk.X, pady=(0, 6))

        ttk.Label(path_row, text="📁 Thư mục lưu:", font=FONT_LABEL_BOLD).pack(side=tk.LEFT, padx=(0, 4))
        self.path_entry = ttk.Entry(path_row, textvariable=self.output_dir_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))

        btn_browse = ttk.Button(path_row, text="Chọn thư mục...", command=self._browse_output_dir)
        btn_browse.pack(side=tk.RIGHT)

        # Progress bar & Status
        self.progress_bar = ttk.Progressbar(card4, orient=tk.HORIZONTAL, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(0, 4))

        self.status_lbl = ttk.Label(card4, text="Sẵn sàng xuất dữ liệu.", font=FONT_LABEL, foreground=CLR_TEXT)
        self.status_lbl.pack(anchor=tk.W)

        # ==============================================================
        # THANH NÚT BẤM TÁC VỤ (ACTION BAR DƯỚI ĐÁY)
        # ==============================================================
        btn_bar = ttk.Frame(self, padding=(16, 10), style="Window.TFrame")
        btn_bar.pack(side=tk.BOTTOM, fill=tk.X)

        self.btn_count = ttk.Button(
            btn_bar,
            text="📊 Đếm số lượng học sinh",
            command=self._count_matching_students
        )
        self.btn_count.pack(side=tk.LEFT, padx=4)

        self.btn_close = ttk.Button(btn_bar, text="Đóng", command=self.destroy)
        self.btn_close.pack(side=tk.RIGHT, padx=4)

        self.btn_open_folder = ttk.Button(
            btn_bar,
            text="📂 Mở thư mục kết quả",
            state=tk.DISABLED,
            command=self._open_output_folder
        )
        self.btn_open_folder.pack(side=tk.RIGHT, padx=4)

        self.btn_start_export = ttk.Button(
            btn_bar,
            text="🚀 Bắt đầu xuất",
            style="Action.TButton",
            command=self._start_export_thread
        )
        self.btn_start_export.pack(side=tk.RIGHT, padx=4)

    # ------------------------------------------------------------------
    # NẠP DỮ LIỆU BỘ LỌC VÀ ĐIỀU KHIỂN
    # ------------------------------------------------------------------
    def _apply_initial_filters(self):
        """Khởi tạo các dropdown theo bộ lọc đang mở bên Phân hệ 2."""
        opts = self.db_manager.get_cascade_filter_options()

        self.cb_huyen_cu["values"] = ["Tất cả"] + opts.get("huyen_cu", [])
        self.cb_xa_phuong["values"] = ["Tất cả"] + opts.get("xa_phuong", [])

        # Áp dụng giá trị ban đầu nếu có
        ini_huyen = self.initial_filters.get("huyen_cu")
        if ini_huyen and ini_huyen in self.cb_huyen_cu["values"]:
            self.huyen_cu_var.set(ini_huyen)
        else:
            first_h = opts.get("huyen_cu", [])
            self.huyen_cu_var.set(first_h[0] if first_h else "Tất cả")

        ini_truong = self.initial_filters.get("ten_truong")
        self._refresh_schools_dropdown(keep_val=ini_truong)

        ini_xa = self.initial_filters.get("xa_phuong")
        if ini_xa and ini_xa in self.cb_xa_phuong["values"]:
            self.xa_phuong_var.set(ini_xa)
        else:
            self.xa_phuong_var.set("Tất cả")
        self._refresh_units_dropdown()

        self._refresh_cohorts_listbox()
        self._update_level_badge()

    def _on_branch_changed(self):
        self._update_level_badge()
        self._refresh_cohorts_listbox()

    def _on_huyen_cu_changed(self, event=None):
        self._refresh_schools_dropdown()
        self._update_level_badge()
        self._refresh_cohorts_listbox()

    def _on_ten_truong_changed(self, event=None):
        self._update_level_badge()
        self._refresh_cohorts_listbox()

    def _on_xa_phuong_changed(self, event=None):
        self._refresh_units_dropdown()
        self._update_level_badge()
        self._refresh_cohorts_listbox()

    def _on_ten_don_vi_changed(self, event=None):
        self._update_level_badge()
        self._refresh_cohorts_listbox()

    def _refresh_schools_dropdown(self, keep_val: str | None = None):
        huyen = self.huyen_cu_var.get()
        opts = self.db_manager.get_cascade_filter_options(huyen_cu=huyen if huyen != "Tất cả" else None)
        sch_vals = ["Tất cả (Xuất cấp Huyện)"] + opts.get("ten_truong", [])
        self.cb_ten_truong["values"] = sch_vals

        if keep_val and keep_val in sch_vals:
            self.ten_truong_var.set(keep_val)
        else:
            self.ten_truong_var.set(sch_vals[0])

    def _refresh_units_dropdown(self, keep_val: str | None = None):
        xa = self.xa_phuong_var.get()
        opts = self.db_manager.get_cascade_filter_options(xa_phuong=xa if xa != "Tất cả" else None)
        unit_vals = ["Tất cả (Xuất cấp Xã)"] + opts.get("ten_don_vi", [])
        self.cb_ten_don_vi["values"] = unit_vals

        if keep_val and keep_val in unit_vals:
            self.ten_don_vi_var.set(keep_val)
        else:
            self.ten_don_vi_var.set(unit_vals[0])

    def _update_level_badge(self):
        """Cập nhật nhãn nhận diện cấp xuất và thông báo Cột 27, 28 tức thì."""
        branch = self.branch_var.get()
        huyen = self.huyen_cu_var.get()
        truong = self.ten_truong_var.get()
        xa = self.xa_phuong_var.get()
        don_vi = self.ten_don_vi_var.get()

        clean_truong = None if "Tất cả" in truong else truong
        clean_don_vi = None if "Tất cả" in don_vi else don_vi

        lvl = self.export_mgr.determine_export_level(
            branch=branch,
            huyen_cu=huyen if huyen != "Tất cả" else None,
            ten_truong=clean_truong,
            xa_phuong=xa if xa != "Tất cả" else None,
            ten_don_vi=clean_don_vi
        )

        level_name = lvl["level"]
        if level_name == "HUYEN_CU":
            txt = f"🏷️ CẤP HUYỆN CŨ [{lvl['base_name']}]: Cột 27 = Mã huyện ({lvl['override_unit_code'] or 'Mã huyện'}), Cột 28 = {lvl['override_unit_name']}"
            color = CLR_PRIMARY
        elif level_name == "TRUONG":
            txt = f"🏷️ CẤP TRƯỜNG HỌC [{lvl['base_name']}]: Giữ nguyên Đơn vị quản lý ở Cột 27 và 28"
            color = CLR_SUCCESS
        elif level_name == "XA_PHUONG":
            txt = f"🏷️ CẤP XÃ/PHƯỜNG MỚI [{lvl['base_name']}]: Cột 27 = Mã xã ({lvl['override_unit_code'] or 'Mã xã'}), Cột 28 = {lvl['override_unit_name']}"
            color = "#7C3AED"  # Purple
        else:
            txt = f"🏷️ CẤP ĐƠN VỊ QUẢN LÝ [{lvl['base_name']}]: Giữ nguyên Đơn vị quản lý ở Cột 27 và 28"
            color = CLR_SUCCESS

        self.level_badge_lbl.config(text=txt, foreground=color)

    def _refresh_cohorts_listbox(self):
        """Nạp danh sách niên khóa thực tế có trong CSDL theo bộ lọc hiện tại."""
        branch = self.branch_var.get()
        huyen = self.huyen_cu_var.get() if branch == "branch_history" else None
        truong_raw = self.ten_truong_var.get() if branch == "branch_history" else None
        truong = None if not truong_raw or "Tất cả" in truong_raw else truong_raw

        xa = self.xa_phuong_var.get() if branch == "branch_new_admin" else None
        don_vi_raw = self.ten_don_vi_var.get() if branch == "branch_new_admin" else None
        don_vi = None if not don_vi_raw or "Tất cả" in don_vi_raw else don_vi_raw

        cohorts = self.db_manager.get_distinct_cohorts_for_filter(
            huyen_cu=huyen if huyen != "Tất cả" else None,
            ten_truong=truong,
            xa_phuong=xa if xa != "Tất cả" else None,
            ten_don_vi=don_vi
        )

        self.cohort_listbox.delete(0, tk.END)
        for c in cohorts:
            self.cohort_listbox.insert(tk.END, c)

        # Mặc định chọn tất cả
        if cohorts:
            self.cohort_listbox.selection_set(0, tk.END)
        self._on_cohort_selection_changed()

    def _select_all_cohorts(self):
        self.cohort_listbox.selection_set(0, tk.END)
        self._on_cohort_selection_changed()

    def _clear_cohorts(self):
        self.cohort_listbox.selection_clear(0, tk.END)
        self._on_cohort_selection_changed()

    def _get_selected_cohorts(self) -> list[str]:
        indices = self.cohort_listbox.curselection()
        return [self.cohort_listbox.get(i) for i in indices]

    def _on_cohort_selection_changed(self, event=None):
        """Cập nhật nhãn phản hồi theo 3 tình huống Niên khóa."""
        selected = self._get_selected_cohorts()
        if not selected:
            self.cohort_feedback_lbl.config(
                text="⚠️ Chưa chọn niên khóa nào (Vui lòng chọn ít nhất 1 năm)",
                foreground=CLR_WARNING
            )
            return

        analysis = analyze_cohorts(selected)
        sorted_years = analysis["sorted_years"]
        num_years = len(sorted_years)

        if analysis["is_single"]:
            yr = sorted_years[0]
            txt = f"🟢 Đã chọn 1 NĂM DUY NHẤT: {yr} ➔ Xuất 1 file dữ liệu [*{yr}.xlsx] + 1 file DS tài liệu số hóa (GHI 1 DÒNG)"
            color = CLR_SUCCESS
        elif analysis["is_consecutive"]:
            y_min = analysis["min_year"]
            y_max = analysis["max_year"]
            txt = f"🔵 Đã chọn {num_years} NĂM LIÊN TIẾP ({y_min} - {y_max}) ➔ Xuất 1 FILE GỘP DUY NHẤT [*{y_min}_{y_max}.xlsx] + 1 file DS tài liệu số hóa (GHI 1 DÒNG)"
            color = CLR_PRIMARY
        else:
            yr_str = ", ".join(str(y) for y in sorted_years)
            txt = f"🟠 Đã chọn {num_years} NĂM RỜI RẠC ({yr_str}) ➔ Xuất {num_years} FILE DỮ LIỆU RIÊNG + 1 file DS tài liệu số hóa chung (GHI {num_years} DÒNG)"
            color = "#D97706"  # Amber

        self.cohort_feedback_lbl.config(text=txt, foreground=color)

    def _browse_output_dir(self):
        d = filedialog.askdirectory(
            title="Chọn thư mục lưu trữ báo cáo Excel",
            initialdir=self.output_dir_var.get(),
            parent=self
        )
        if d:
            self.output_dir_var.set(d)

    def _open_output_folder(self):
        target = self.last_exported_dir or Path(self.output_dir_var.get())
        open_directory(target)

    def _count_matching_students(self):
        """Đếm nhanh số lượng học sinh khớp bộ lọc."""
        branch = self.branch_var.get()
        huyen = self.huyen_cu_var.get() if branch == "branch_history" else None
        truong_raw = self.ten_truong_var.get() if branch == "branch_history" else None
        truong = None if not truong_raw or "Tất cả" in truong_raw else truong_raw

        xa = self.xa_phuong_var.get() if branch == "branch_new_admin" else None
        don_vi_raw = self.ten_don_vi_var.get() if branch == "branch_new_admin" else None
        don_vi = None if not don_vi_raw or "Tất cả" in don_vi_raw else don_vi_raw

        selected_cohorts = self._get_selected_cohorts()
        if not selected_cohorts:
            messagebox.showwarning("Chưa chọn niên khóa", "Vui lòng chọn ít nhất 1 niên khóa.", parent=self)
            return

        students = self.db_manager.query_students_for_export(
            huyen_cu=huyen if huyen != "Tất cả" else None,
            ten_truong=truong,
            xa_phuong=xa if xa != "Tất cả" else None,
            ten_don_vi=don_vi,
            cohorts=selected_cohorts
        )

        messagebox.showinfo(
            "Kết quả kiểm tra",
            f"Tìm thấy: {len(students):,} học sinh khớp toàn bộ tiêu chí lọc hiện tại.",
            parent=self
        )

    # ------------------------------------------------------------------
    # TIẾN TRÌNH XUẤT NỀN (BACKGROUND THREAD)
    # ------------------------------------------------------------------
    def _start_export_thread(self):
        if self.is_exporting:
            return

        selected_cohorts = self._get_selected_cohorts()
        if not selected_cohorts:
            messagebox.showwarning("Chưa chọn niên khóa", "Vui lòng chọn ít nhất 1 niên khóa trước khi xuất.", parent=self)
            return

        out_path = Path(self.output_dir_var.get().strip())
        if not out_path:
            messagebox.showwarning("Chưa chọn thư mục", "Vui lòng chọn thư mục lưu trữ kết quả.", parent=self)
            return

        self.is_exporting = True
        self.btn_start_export.config(state=tk.DISABLED)
        self.btn_open_folder.config(state=tk.DISABLED)
        self.progress_bar["value"] = 0
        self.status_lbl.config(text="Đang chuẩn bị dữ liệu...", foreground=CLR_TEXT)

        thread = threading.Thread(target=self._run_export, daemon=True)
        thread.start()

    def _run_export(self):
        try:
            branch = self.branch_var.get()
            huyen = self.huyen_cu_var.get() if branch == "branch_history" else None
            truong_raw = self.ten_truong_var.get() if branch == "branch_history" else None
            truong = None if not truong_raw or "Tất cả" in truong_raw else truong_raw

            xa = self.xa_phuong_var.get() if branch == "branch_new_admin" else None
            don_vi_raw = self.ten_don_vi_var.get() if branch == "branch_new_admin" else None
            don_vi = None if not don_vi_raw or "Tất cả" in don_vi_raw else don_vi_raw

            selected_cohorts = self._get_selected_cohorts()
            out_dir = Path(self.output_dir_var.get().strip())
            export_err = self.export_errors_var.get()

            def progress_cb(cur, total, msg):
                self.after(0, lambda: self._update_progress_ui(cur, total, msg))

            results = self.export_mgr.execute_export(
                db_manager=self.db_manager,
                output_dir=out_dir,
                branch=branch,
                huyen_cu=huyen if huyen != "Tất cả" else None,
                ten_truong=truong,
                xa_phuong=xa if xa != "Tất cả" else None,
                ten_don_vi=don_vi,
                selected_cohorts=selected_cohorts,
                export_error_files=export_err,
                progress_callback=progress_cb
            )

            self.last_exported_dir = out_dir
            self.after(0, lambda: self._on_export_success(results))

        except Exception as e:
            self.after(0, lambda err=e: self._on_export_error(err))

    def _update_progress_ui(self, current: int, total: int, message: str):
        self.progress_bar["maximum"] = total
        self.progress_bar["value"] = current
        self.status_lbl.config(text=message, foreground=CLR_TEXT)

    def _on_export_success(self, results: dict):
        self.is_exporting = False
        self.btn_start_export.config(state=tk.NORMAL)
        self.btn_open_folder.config(state=tk.NORMAL)
        self.progress_bar["value"] = 100

        total_st = results.get("total_students", 0)
        num_st_files = len(results.get("student_files", []))
        num_doc_files = len(results.get("doc_files", []))
        num_err_files = len(results.get("error_files", []))

        msg_lines = [
            f"🎉 Xuất báo cáo Excel thành công!",
            f"• Tổng số học sinh: {total_st:,}",
            f"• Số tệp dữ liệu học sinh (28 cột): {num_st_files}",
            f"• Số tệp Danh sách tài liệu số hóa (11 cột): {num_doc_files}",
        ]
        if num_err_files > 0:
            msg_lines.append(f"• Số tệp Danh sách lỗi kèm theo: {num_err_files}")

        msg_lines.append(f"\nĐã lưu an toàn tại:\n{self.last_exported_dir}")
        self.status_lbl.config(text="✅ Xuất hoàn tất 100%!", foreground=CLR_SUCCESS)

        messagebox.showinfo("Xuất báo cáo thành công", "\n".join(msg_lines), parent=self)

    def _on_export_error(self, err: Exception):
        self.is_exporting = False
        self.btn_start_export.config(state=tk.NORMAL)
        self.status_lbl.config(text=f"❌ Lỗi: {err}", foreground=CLR_DANGER)
        messagebox.showerror("Lỗi xuất dữ liệu", f"Đã xảy ra lỗi trong quá trình xuất Excel:\n{err}", parent=self)
