from __future__ import annotations

import os
import sys
import threading
import queue
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from src.lookup_manager import LookupManager
from src.excel_processor import ExcelProcessor
from src.db_manager import DatabaseManager
from src.version import APP_VERSION, APP_NAME
from src.updater import check_for_updates_async, check_for_updates, open_release_page
from src.gui.dialogs import (
    SchoolBlockDialog,
    CohortBlockDialog,
    DateFixDialog,
    GradDateBlockDialog,
    DataInspectorDialog
)
from src.gui.custom_tools_dialog import CustomToolsView
from src.gui.data_management_view import DataManagementView
from src.gui.theme import (
    IS_DARK,
    FONT_TITLE,
    FONT_SUBTITLE,
    FONT_LABEL,
    FONT_LABEL_BOLD,
    FONT_DATA_BOLD,
    FONT_MUTED,
    FONT_LOG,
    get_theme_palette,
    configure_ttk_theme,
)

class MainWindow:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} - v{APP_VERSION}")
        
        # Đặt kích thước và vị trí cửa sổ tối ưu đa nền tảng (rộng rãi, chữ to rõ, không bị tràn màn hình)
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        win_w = min(1280, max(1040, sw - 40))
        win_h = min(820, max(680, sh - 80))
        x = max(0, (sw - win_w) // 2)
        y = max(0, (sh - win_h - 40) // 2)
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(1000, 620)

        # Quản lý giao diện Sáng / Tối (Dark Mode)
        self.is_dark = IS_DARK
        self.palette = get_theme_palette(self.is_dark)
        self.root.configure(bg=self.palette.WINDOW_BG)

        # Cấu hình style
        self.style = ttk.Style(self.root)
        self._setup_styles()

        # Biến trạng thái
        self.lookup_mgr = None
        self.processor = None
        self.db_manager = None
        self.data_mgmt_view = None
        self.dialog_queue = queue.Queue()
        self.log_queue = queue.Queue()
        self.is_processing = False
        self.cancel_requested = False
        self.all_unit_students = []
        self.last_output_dir = None
        self.latest_update_info = None

        self._build_ui()
        self._init_data()

        # Bắt đầu vòng lặp kiểm tra queue cho thread-safe GUI
        self.root.after(50, self._process_queues)

        # Cập nhật layout và kích hoạt khắc phục lỗi Cocoa redraw lag trên macOS
        self.root.update_idletasks()
        self.root.lift()
        if sys.platform == "darwin":
            self.root.after(80, self._force_macos_redraw)

    def _force_macos_redraw(self):
        """Khắc phục triệt để lỗi macOS Tkinter không vẽ nội dung (trắng/trống cửa sổ) cho đến khi người dùng resize."""
        try:
            w = self.root.winfo_width()
            h = self.root.winfo_height()
            if w > 10 and h > 10:
                self.root.geometry(f"{w+1}x{h+1}")
                self.root.update_idletasks()
                self.root.geometry(f"{w}x{h}")
                self.root.update_idletasks()
                self.root.lift()
        except Exception:
            pass

    def _setup_styles(self):
        configure_ttk_theme(self.style, self.palette)

    def _toggle_theme(self):
        """Chuyển đổi tức thì giữa Chế độ Sáng (Light) và Chế độ Tối (Dark)."""
        self.is_dark = not self.is_dark
        self.palette = get_theme_palette(self.is_dark)
        self._apply_theme()

    def _apply_theme(self):
        """Áp dụng bảng màu mới cho toàn bộ widget của ứng dụng."""
        p = self.palette
        self.root.configure(bg=p.WINDOW_BG)
        configure_ttk_theme(self.style, p)

        # Header banner
        if hasattr(self, 'hdr_frame'):
            self.hdr_frame.configure(bg=p.HEADER_BANNER_BG)
            self.hdr_title_lbl.configure(bg=p.HEADER_BANNER_BG, fg=p.HEADER_BANNER_FG)
            self.hdr_sub_lbl.configure(bg=p.HEADER_BANNER_BG, fg=p.HEADER_SUB_FG)
            if hasattr(self, 'hdr_left'):
                self.hdr_left.configure(bg=p.HEADER_BANNER_BG)

        # Labels với màu chuyên biệt
        if hasattr(self, 'db_path_lbl'):
            self.db_path_lbl.configure(foreground=p.PRIMARY)
        if hasattr(self, 'db_schools_lbl'):
            self.db_schools_lbl.configure(foreground=p.SUCCESS)
        if hasattr(self, 'db_districts_lbl'):
            self.db_districts_lbl.configure(foreground=p.ACCENT)
        if hasattr(self, 'unit_code_lbl'):
            self.unit_code_lbl.configure(foreground=p.PRIMARY)
            self.unit_name_lbl.configure(foreground=p.PRIMARY)
            self.unit_commune_lbl.configure(foreground=p.ACCENT)
            self.school_count_lbl.configure(foreground=p.SUCCESS)
            self.schools_display_lbl.configure(foreground=p.MUTED)
            self.files_count_lbl.configure(foreground=p.MUTED)
            self.status_lbl.configure(foreground=p.PRIMARY)
            if hasattr(self, 'stat_total_lbl'):
                self.stat_total_lbl.configure(foreground=p.TEXT)
            self.stat_da_so_hoa_lbl.configure(foreground=p.SUCCESS)
            self.stat_chua_day_du_lbl.configure(foreground=p.WARNING)
            self.stat_chua_so_hoa_lbl.configure(foreground=p.DANGER)

        # Log console
        if hasattr(self, 'log_text'):
            self.log_text.configure(
                bg=p.LOG_BG,
                fg=p.LOG_FG,
                insertbackground=p.LOG_FG
            )
            self.log_text.tag_config("info", foreground=p.LOG_FG)
            self.log_text.tag_config("success", foreground=p.SUCCESS, font=FONT_LOG)
            self.log_text.tag_config("warning", foreground=p.WARNING)
            self.log_text.tag_config("error", foreground=p.DANGER, font=FONT_LOG)

        # Nút theme
        if hasattr(self, 'theme_btn'):
            self.theme_btn.configure(text="☀️ Chế độ Sáng (Đổi 🌙)" if not self.is_dark else "🌙 Chế độ Tối (Đổi ☀️)")

    def _build_ui(self):
        p = self.palette

        # 1. Header Banner (TOP)
        self.hdr_frame = tk.Frame(self.root, bg=p.HEADER_BANNER_BG, padx=16, pady=10)
        self.hdr_frame.pack(side=tk.TOP, fill=tk.X)

        self.hdr_left = tk.Frame(self.hdr_frame, bg=p.HEADER_BANNER_BG)
        self.hdr_left.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.hdr_title_lbl = tk.Label(
            self.hdr_left,
            text=f"{APP_NAME} - SỐ HÓA VĂN BẰNG CHỨNG CHỈ THCS",
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_BANNER_FG,
            font=FONT_TITLE
        )
        self.hdr_title_lbl.pack(anchor=tk.W)

        self.hdr_sub_lbl = tk.Label(
            self.hdr_left,
            text="Chuẩn hóa dữ liệu cấp bằng và lập hồ sơ số hóa",
            bg=p.HEADER_BANNER_BG,
            fg=p.HEADER_SUB_FG,
            font=FONT_SUBTITLE
        )
        self.hdr_sub_lbl.pack(anchor=tk.W, pady=(2, 0))

        # 2. Bottom Bar (GHIM CỐ ĐỊNH Ở ĐÁY CỬA SỔ - LUÔN LUÔN NHÌN THẤY 100%)
        bottom_frame = ttk.Frame(self.root, style="Window.TFrame", padding=(12, 6))
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.open_output_btn = ttk.Button(bottom_frame, text="📂 Mở thư mục CSDL", command=self._open_data_dir)
        self.open_output_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.export_btn = ttk.Button(
            bottom_frame,
            text="📤 Xuất Báo Cáo Excel...",
            command=self._open_export_dialog
        )
        self.export_btn.pack(side=tk.LEFT, padx=6)

        self.theme_btn = ttk.Button(
            bottom_frame,
            text="☀️ Chế độ Sáng (Đổi 🌙)" if not self.is_dark else "🌙 Chế độ Tối (Đổi ☀️)",
            command=self._toggle_theme
        )
        self.theme_btn.pack(side=tk.LEFT, padx=6)

        self.update_btn = ttk.Button(
            bottom_frame,
            text=f"🔄 Bản {APP_VERSION}",
            command=self._check_update_manually
        )
        self.update_btn.pack(side=tk.LEFT, padx=6)

        exit_btn = ttk.Button(bottom_frame, text="Thoát", command=self._on_exit)
        exit_btn.pack(side=tk.RIGHT)

        # 3. Main Container (Sử dụng ttk.Notebook để điều hướng giữa Phân hệ 1 và Phân hệ 2)
        main_container = ttk.Frame(self.root, style="Window.TFrame", padding=(10, 4))
        main_container.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        self.notebook = ttk.Notebook(main_container)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        self.notebook.bind("<<NotebookTabChanged>>", self._on_tab_changed)

        # Tab 1: 📥 Nạp & Chuẩn hóa Dữ liệu (Phân hệ 1)
        tab1_frame = ttk.Frame(self.notebook, style="Window.TFrame", padding=(4, 4))
        self.notebook.add(tab1_frame, text="  📥 1. Nạp & Chuẩn hóa Dữ liệu  ")

        # Tab 2: 📊 Quản trị, Lọc & Sửa đổi CSDL (Phân hệ 2)
        self.tab2_frame = ttk.Frame(self.notebook, style="Window.TFrame", padding=(2, 2))
        self.notebook.add(self.tab2_frame, text="  📊 2. Quản trị, Lọc & Sửa đổi CSDL  ")

        # Tab 3: 🛠️ Công cụ & Tiện ích mở rộng
        self.tab3_frame = ttk.Frame(self.notebook, style="Window.TFrame", padding=(2, 2))
        self.notebook.add(self.tab3_frame, text="  🛠️ 3. Công cụ & Tiện ích mở rộng  ")

        # ==============================================================
        # Card 1: Nguồn dữ liệu Excel đầu vào (Chọn 1 File hoặc Thư mục)
        # ==============================================================
        file_card = ttk.LabelFrame(tab1_frame, text="1. Nguồn dữ liệu Excel (Chọn 1 File hoặc Thư mục)", style="Card.TLabelframe", padding=10)
        file_card.pack(fill=tk.X, pady=(0, 6))

        f_row1 = ttk.Frame(file_card)
        f_row1.pack(fill=tk.X)

        self.path_var = tk.StringVar()
        self.path_entry = ttk.Entry(f_row1, textvariable=self.path_var)
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))

        browse_file_btn = ttk.Button(f_row1, text="📂 Chọn 1 file...", command=self._browse_file)
        browse_file_btn.pack(side=tk.LEFT, padx=3)

        browse_dir_btn = ttk.Button(f_row1, text="📁 Chọn thư mục...", command=self._browse_dir)
        browse_dir_btn.pack(side=tk.LEFT, padx=3)


        self.files_count_lbl = ttk.Label(
            file_card,
            text="Hệ thống tự động quét nhận diện các trường, niên khóa; Đơn vị quản lý được tự động suy ra sau khi chuẩn hóa Trường học.",
            font=FONT_MUTED,
            foreground=p.MUTED
        )
        self.files_count_lbl.pack(anchor=tk.W, pady=(4, 0))

        # ==============================================================
        # Card 2: Quản trị CSDL SQLite tập trung
        # ==============================================================
        db_card = ttk.LabelFrame(tab1_frame, text="2. Cơ sở dữ liệu SQLite tập trung", style="Card.TLabelframe", padding=10)
        db_card.pack(fill=tk.X, pady=(0, 6))

        db_row1 = ttk.Frame(db_card)
        db_row1.pack(fill=tk.X, pady=2)

        ttk.Label(db_row1, text="Tệp CSDL: ", font=FONT_LABEL_BOLD).pack(side=tk.LEFT)
        self.db_path_lbl = ttk.Label(db_row1, text="data/vbcc_database.db", font=FONT_DATA_BOLD, foreground=p.PRIMARY)
        self.db_path_lbl.pack(side=tk.LEFT, padx=(0, 16))

        self.db_total_lbl = ttk.Label(db_row1, text="Tổng học sinh trong CSDL: 0", font=FONT_LABEL_BOLD)
        self.db_total_lbl.pack(side=tk.LEFT, padx=(0, 16))

        self.db_schools_lbl = ttk.Label(db_row1, text="Số trường: 0", font=FONT_LABEL, foreground=p.SUCCESS)
        self.db_schools_lbl.pack(side=tk.LEFT, padx=(0, 16))

        self.db_districts_lbl = ttk.Label(db_row1, text="Số huyện: 0", font=FONT_LABEL, foreground=p.ACCENT)
        self.db_districts_lbl.pack(side=tk.LEFT, padx=(0, 16))

        backup_btn = ttk.Button(db_row1, text="🛡️ Sao lưu CSDL ngay", command=self._manual_backup)
        backup_btn.pack(side=tk.RIGHT, padx=4)

        open_db_dir_btn = ttk.Button(db_row1, text="📂 Mở thư mục CSDL", command=self._open_data_dir)
        open_db_dir_btn.pack(side=tk.RIGHT, padx=4)

        # ==============================================================
        # Card 3: Tiến trình chuẩn hóa & Điều khiển
        # ==============================================================
        prog_card = ttk.LabelFrame(tab1_frame, text="3. Tiến trình chuẩn hóa & Nạp CSDL", style="Card.TLabelframe", padding=8)
        prog_card.pack(fill=tk.X, pady=(0, 6))

        self.progress_bar = ttk.Progressbar(prog_card, orient=tk.HORIZONTAL, mode='determinate')
        self.progress_bar.pack(fill=tk.X, pady=(0, 6))

        ctrl_row = ttk.Frame(prog_card)
        ctrl_row.pack(fill=tk.X)

        self.status_lbl = ttk.Label(ctrl_row, text="Sẵn sàng thực hiện.", font=FONT_LABEL_BOLD, foreground=p.PRIMARY)
        self.status_lbl.pack(side=tk.LEFT)

        self.stop_btn = ttk.Button(ctrl_row, text="⏹ Dừng", state=tk.DISABLED, command=self._stop_processing)
        self.stop_btn.pack(side=tk.RIGHT, padx=4)

        self.start_btn = ttk.Button(ctrl_row, text="▶️ Bắt đầu chuẩn hóa & Nạp CSDL", style="Action.TButton", command=self._start_processing)
        self.start_btn.pack(side=tk.RIGHT, padx=4)

        # ==============================================================
        # Card 4: Thống kê phiên nạp & Nhật ký hoạt động
        # ==============================================================
        log_card = ttk.LabelFrame(tab1_frame, text="4. Thống kê phiên nạp & Nhật ký hoạt động", style="Card.TLabelframe", padding=8)
        log_card.pack(fill=tk.BOTH, expand=True, pady=(0, 6))

        stat_row = ttk.Frame(log_card)
        stat_row.pack(fill=tk.X, pady=(0, 4))

        self.stat_total_lbl = ttk.Label(stat_row, text="Tổng học sinh: 0", font=FONT_LABEL_BOLD)
        self.stat_total_lbl.pack(side=tk.LEFT, padx=(0, 20))

        self.stat_da_so_hoa_lbl = ttk.Label(stat_row, text="✅ Đã số hóa đầy đủ: 0", font=FONT_LABEL_BOLD, foreground=p.SUCCESS)
        self.stat_da_so_hoa_lbl.pack(side=tk.LEFT, padx=(0, 20))

        self.stat_chua_day_du_lbl = ttk.Label(stat_row, text="⚠️ Số hoá chưa đầy đủ: 0", font=FONT_LABEL_BOLD, foreground=p.WARNING)
        self.stat_chua_day_du_lbl.pack(side=tk.LEFT, padx=(0, 20))

        self.stat_chua_so_hoa_lbl = ttk.Label(stat_row, text="❌ Chưa số hoá: 0", font=FONT_LABEL_BOLD, foreground=p.DANGER)
        self.stat_chua_so_hoa_lbl.pack(side=tk.LEFT)

        # Log text console
        self.log_text = ScrolledText(
            log_card,
            height=5,
            font=FONT_LOG,
            wrap=tk.WORD,
            bg=p.LOG_BG,
            fg=p.LOG_FG,
            insertbackground=p.LOG_FG
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        self.log_text.tag_config("info", foreground=p.LOG_FG)
        self.log_text.tag_config("success", foreground=p.SUCCESS, font=FONT_LOG)
        self.log_text.tag_config("warning", foreground=p.WARNING)
        self.log_text.tag_config("error", foreground=p.DANGER, font=FONT_LOG)

    def _refresh_db_stats_ui(self):
        """Cập nhật các nhãn số liệu thống kê CSDL SQLite lên giao diện."""
        if not self.db_manager:
            return
        try:
            stats = self.db_manager.get_statistics()
            self.db_path_lbl.config(text=str(self.db_manager.db_path.name))
            self.db_total_lbl.config(text=f"Tổng học sinh trong CSDL: {stats['total']:,}")
            self.db_schools_lbl.config(text=f"Số trường: {stats['total_schools']:,}")
            self.db_districts_lbl.config(text=f"Số huyện: {stats['total_districts']:,}")
        except Exception as e:
            self._log(f"⚠️ Không thể cập nhật số liệu CSDL: {e}", "warning")

    def _manual_backup(self):
        """Sao lưu nhanh CSDL SQLite."""
        if not self.db_manager:
            return
        try:
            b_path = self.db_manager.backup_database()
            messagebox.showinfo("Sao lưu thành công", f"Đã sao lưu CSDL an toàn tại:\n{b_path.name}", parent=self.root)
            self._log(f"🛡️ Đã tạo bản sao lưu CSDL an toàn: {b_path.name}", "success")
        except Exception as e:
            messagebox.showerror("Lỗi sao lưu", f"Không thể sao lưu CSDL: {e}", parent=self.root)

    def _open_data_dir(self):
        """Mở thư mục data chứa CSDL SQLite và các bản sao lưu."""
        if not self.db_manager:
            return
        data_dir = self.db_manager.db_path.parent
        try:
            if sys.platform == "win32":
                os.startfile(str(data_dir))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(data_dir)])
            else:
                subprocess.Popen(["xdg-open", str(data_dir)])
        except Exception as e:
            messagebox.showerror("Lỗi mở thư mục", f"Không thể mở thư mục CSDL: {e}", parent=self.root)

    def _init_data(self):
        """Khởi tạo LookupManager, DatabaseManager, nạp danh mục và khởi tạo Giao diện Phân hệ 2."""
        try:
            self.lookup_mgr = LookupManager()
            self.db_manager = DatabaseManager()
            self.processor = ExcelProcessor(self.lookup_mgr)
            self._refresh_db_stats_ui()
            
            # Khởi tạo Giao diện Phân hệ 2 tại Tab 2
            if hasattr(self, 'tab2_frame'):
                self.data_mgmt_view = DataManagementView(
                    self.tab2_frame,
                    db_manager=self.db_manager,
                    lookup_mgr=self.lookup_mgr
                )
                self.data_mgmt_view.pack(fill=tk.BOTH, expand=True)

            # Khởi tạo Giao diện Tab 3: CÔNG CỤ (Tách & Gộp Excel - Nhóm 1)
            if hasattr(self, 'tab3_frame'):
                self.custom_tools_view = CustomToolsView(
                    self.tab3_frame,
                    is_dark=self.is_dark
                )
                self.custom_tools_view.pack(fill=tk.BOTH, expand=True)

            db_stats = self.db_manager.get_statistics()
            self._log(f"🗄️ CSDL SQLite: {self.db_manager.db_path.name} (Hiện có {db_stats['total']:,} học sinh, {db_stats['total_schools']} trường, {db_stats['total_districts']} huyện).", "info")
            self._log(f"📖 Nguồn danh mục: {self.lookup_mgr.excel_path.name} ({len(self.lookup_mgr.all_schools_list)} trường học thuộc 3 tỉnh).", "success")
        except Exception as e:
            self._log(f"❌ Lỗi khởi tạo hệ thống: {e}", "error")
            messagebox.showerror("Lỗi khởi tạo", f"Không thể khởi tạo CSDL hoặc đọc file danh mục:\n{e}", parent=self.root)

        # Kiểm tra phiên bản mới từ GitHub Releases ngầm
        check_for_updates_async(self._on_update_checked, current_version=APP_VERSION)

    def _switch_to_tab(self, tab_idx: int):
        """Chuyển đổi giữa các tab trên giao diện chính."""
        if hasattr(self, 'notebook'):
            self.notebook.select(tab_idx)

    def _open_export_dialog(self):
        """Mở Hộp thoại Xuất Báo Cáo & Dữ Liệu Excel (Phân hệ 3)."""
        if not self.db_manager:
            return
        initial_filters = {}
        if hasattr(self, 'data_mgmt_view') and self.data_mgmt_view:
            initial_filters = {
                "tinh_cu": self.data_mgmt_view.tinh_cu_var.get(),
                "huyen_cu": self.data_mgmt_view.huyen_cu_var.get(),
                "ten_truong": self.data_mgmt_view.ten_truong_var.get(),
                "nien_khoa": self.data_mgmt_view.nien_khoa_var.get(),
                "xa_phuong": self.data_mgmt_view.xa_phuong_var.get(),
                "ten_don_vi": self.data_mgmt_view.ten_don_vi_var.get(),
            }
        from src.gui.export_dialog import ExportDialog
        ExportDialog(
            parent=self.root,
            db_manager=self.db_manager,
            lookup_mgr=self.lookup_mgr,
            initial_filters=initial_filters
        )

    def _on_tab_changed(self, event=None):
        """Khi người dùng chuyển tab, tự động làm mới dữ liệu phù hợp."""
        if hasattr(self, 'notebook') and hasattr(self, 'data_mgmt_view') and self.data_mgmt_view:
            try:
                idx = self.notebook.index(self.notebook.select())
                if idx == 1:
                    self.data_mgmt_view.refresh_filter_options()
                    self.data_mgmt_view.refresh_data(reset_page=False)
            except Exception:
                pass

    def _on_update_checked(self, res: dict):
        """Callback khi kiểm tra phiên bản trên GitHub hoàn thành."""
        self.latest_update_info = res
        if res.get("has_update"):
            latest_v = res.get("latest_version", "")
            self.root.after(0, lambda: self.update_btn.config(text=f"🔥 Có bản mới {latest_v}!"))
            self.root.after(0, lambda: self._log(f"\n🔔 PHÁT HIỆN BẢN MỚI: {latest_v} trên GitHub (Bản hiện tại: v{APP_VERSION}). Bấm 'Có bản mới' ở góc dưới để cập nhật!", "warning"))

    def _check_update_manually(self):
        """Người dùng bấm nút kiểm tra cập nhật thủ công."""
        info = self.latest_update_info
        if not info or not info.get("success"):
            info = check_for_updates(current_version=APP_VERSION)
            self.latest_update_info = info

        if info.get("has_update"):
            latest_v = info.get("latest_version", "")
            notes = info.get("release_notes", "").strip()
            msg = f"Đã có phiên bản mới: {latest_v} trên GitHub!\n(Phiên bản bạn đang dùng: v{APP_VERSION})\n\n"
            if notes:
                msg += f"Thông tin cập nhật:\n{notes[:350]}\n\n"
            msg += "Bạn có muốn mở trang tải bản cập nhật mới nhất không?"
            ans = messagebox.askyesno("Cập nhật phần mềm", msg, parent=self.root)
            if ans:
                open_release_page(info.get("release_url", ""))
        elif info.get("success"):
            messagebox.showinfo("Cập nhật phần mềm", f"Bạn đang sử dụng phiên bản mới nhất (v{APP_VERSION})!", parent=self.root)
        else:
            err = info.get("error", "Không thể kết nối Internet")
            ans = messagebox.askyesno("Kiểm tra cập nhật", f"Không thể kiểm tra cập nhật tự động:\n{err}\n\nBạn có muốn mở trang GitHub Releases để kiểm tra thủ công không?", parent=self.root)
            if ans:
                open_release_page()

    def _update_unit_combobox(self, unit_list: list[str]):
        self.unit_combo["values"] = unit_list
        if unit_list:
            self.unit_combo.current(0)
            self._on_unit_selected()

    def _on_search_unit(self, *args):
        """Lọc danh sách đơn vị khi người dùng gõ vào ô tìm kiếm (hỗ trợ tìm tên, xã, hoặc mã)."""
        if not self.lookup_mgr:
            return
        kw = self.unit_search_var.get().strip().lower()
        if not kw:
            filtered = self.lookup_mgr.unit_names_list
        else:
            filtered = [u for u in self.lookup_mgr.unit_names_list if kw in u.lower()]
        self.unit_combo["values"] = filtered
        if filtered:
            self.unit_combo.current(0)
            self._on_unit_selected()

    def _on_unit_selected(self, event=None):
        """Khi người dùng chọn một đơn vị từ combobox."""
        selected_unit = self.unit_combo_var.get()
        if not selected_unit or not self.lookup_mgr:
            return
        info = self.lookup_mgr.get_unit_info(selected_unit)
        if info:
            self.unit_code_lbl.config(text=info.get("code", "---"))
            self.unit_name_lbl.config(text=info.get("name", "---"))
            self.unit_commune_lbl.config(text=info.get("commune", "(Trống)") or "(Trống)")
            sch_list = info.get("schools", [])
            self.school_count_lbl.config(text=f"{len(sch_list)} trường")
            if sch_list:
                disp = "; ".join(sch_list)
                if len(disp) > 130:
                    disp = disp[:127] + "..."
                self.schools_display_lbl.config(text=disp)
            else:
                self.schools_display_lbl.config(text="(Không có trường trực thuộc trong danh mục)")

    def _browse_file(self):
        file_selected = filedialog.askopenfilename(
            title="Chọn file Excel nguồn cần chuẩn hóa",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
            parent=self.root
        )
        if file_selected:
            self.path_var.set(file_selected)
            self.files_count_lbl.config(text=f"Đã chọn 1 tệp: {Path(file_selected).name}")

    def _browse_dir(self):
        dir_selected = filedialog.askdirectory(
            title="Chọn thư mục Đơn vị chứa các file Excel nguồn",
            parent=self.root
        )
        if dir_selected:
            self.path_var.set(dir_selected)
            p = Path(dir_selected)
            files = [
                f for f in p.rglob("*.xlsx")
                if not f.name.startswith("~$")
                and not f.name.startswith("._")
                and not any(tag in f.name for tag in ["_Đa_Gop_", "_DS_Loi", "_Du_Lieu_Goc", "_DS_Lỗi", "_TongHop", "_DS_Tai_Lieu_So_Hoa", "_Nhat_Ky_Thay_Doi"])
                and not any(tag in part.lower() for tag in ["da_chuan_hoa", "ket_qua_so_hoa"] for part in f.parts)
            ]
            self.files_count_lbl.config(
                text=f"Tìm thấy {len(files)} tệp Excel trong thư mục {p.name} (bao gồm cả các thư mục trường con)"
            )

    def _log(self, msg: str, level: str = "info"):
        self.log_queue.put((msg, level))

    def _process_queues(self):
        """Xử lý hàng đợi log và hàng đợi dialog tương tác trên Main GUI Thread."""
        while not self.log_queue.empty():
            msg, level = self.log_queue.get_nowait()
            self.log_text.insert(tk.END, msg + "\n", level)
            self.log_text.see(tk.END)

        if not self.dialog_queue.empty():
            req = self.dialog_queue.get_nowait()
            d_type = req["type"]
            args = req["args"]
            res_box = req["res_box"]
            evt = req["event"]

            try:
                if d_type == "school_block":
                    block_info, allowed_schools, sugg_school = args
                    dlg = SchoolBlockDialog(self.root, block_info, allowed_schools, sugg_school)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.apply_to_all_similar)
                elif d_type == "cohort_block":
                    block_info = args[0]
                    default_y_ra = args[1] if len(args) > 1 else ""
                    reason = args[2] if len(args) > 2 else ""
                    dlg = CohortBlockDialog(self.root, block_info, default_y_ra, reason)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.apply_to_all_similar)
                elif d_type == "date_fix":
                    student_info, field_name, sugg_dob = args
                    dlg = DateFixDialog(self.root, student_info, field_name, sugg_dob)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.skip_all)
                elif d_type == "grad_date_block":
                    block_info, sugg_grad = args
                    dlg = GradDateBlockDialog(self.root, block_info, sugg_grad)
                    self.root.wait_window(dlg)
                    res_box["result"] = (dlg.result, dlg.apply_to_group)
            except Exception as e:
                print(f"Lỗi hiển thị dialog: {e}")
                try:
                    if 'dlg' in locals() and dlg and dlg.winfo_exists():
                        dlg.destroy()
                except Exception:
                    pass
                res_box["result"] = (None, False)
            finally:
                evt.set()

        self.root.after(50, self._process_queues)

    def _dialog_callback(self, dialog_type: str, *args):
        """Hàm callback được gọi từ Worker Thread để mở Dialog GUI trên Main Thread."""
        evt = threading.Event()
        res_box = {"result": None}
        self.dialog_queue.put({
            "type": dialog_type,
            "args": args,
            "res_box": res_box,
            "event": evt
        })
        evt.wait()
        return res_box["result"]

    def _start_processing(self):
        if self.is_processing:
            return

        raw_path = self.path_var.get().strip().strip("'\"")
        if not raw_path or not Path(raw_path).exists():
            messagebox.showwarning("Cảnh báo", "Vui lòng chọn tệp hoặc thư mục Excel hợp lệ!", parent=self.root)
            return

        target_path = Path(raw_path)
        if target_path.is_file():
            files_to_process = [target_path]
            base_dir = target_path.parent
        else:
            files_to_process = sorted(
                [
                    f for f in target_path.rglob("*.xlsx")
                    if not f.name.startswith("~$")
                    and not f.name.startswith("._")
                    and not any(tag in f.name for tag in ["_Đa_Gop_", "_DS_Loi", "_Du_Lieu_Goc", "_DS_Lỗi", "_TongHop", "_DS_Tai_Lieu_So_Hoa", "_Nhat_Ky_Thay_Doi"])
                    and not any(tag in part.lower() for tag in ["da_chuan_hoa", "ket_qua_so_hoa"] for part in f.parts)
                ],
                key=lambda x: str(x).lower()
            )
            base_dir = target_path

        if not files_to_process:
            messagebox.showwarning("Cảnh báo", "Không tìm thấy file Excel (*.xlsx) nào để xử lý!", parent=self.root)
            return

        # Tính relative paths để kiểm tra trùng
        rel_files = []
        for f in files_to_process:
            if f.is_relative_to(base_dir):
                rel_files.append(f.relative_to(base_dir).as_posix())
            else:
                rel_files.append(f.name)

        # Kiểm tra các sheet đã tồn tại trong CSDL SQLite
        overwrite_sheets = set()
        if self.db_manager:
            existing_sheets = self.db_manager.check_existing_sheets(rel_files)
            if existing_sheets:
                total_dup_students = sum(existing_sheets.values())
                sample_sheets = [f"• {f} [{s}] ({c} HS)" for (f, s), c in list(existing_sheets.items())[:5]]
                sample_str = "\n".join(sample_sheets)
                if len(existing_sheets) > 5:
                    sample_str += f"\n... và {len(existing_sheets) - 5} sheet khác."

                msg = (
                    f"Phát hiện {len(existing_sheets)} Sheet (khoảng {total_dup_students:,} học sinh) "
                    f"đã tồn tại trong CSDL SQLite:\n\n{sample_str}\n\n"
                    f"Bạn muốn xử lý các sheet này như thế nào?\n"
                    f"• Chọn 'Yes' (Có): GHI ĐÈ (làm mới toàn bộ dữ liệu của các sheet này theo file Excel hiện tại).\n"
                    f"• Chọn 'No' (Không): BỎ QUA các sheet đã có, chỉ nạp tiếp các sheet/file mới.\n"
                    f"• Chọn 'Cancel' (Hủy): Dừng lại để kiểm tra."
                )
                ans = messagebox.askyesnocancel("Phát hiện dữ liệu trùng lặp trong CSDL", msg, parent=self.root)
                if ans is None:
                    return
                elif ans is True:
                    overwrite_sheets = set(existing_sheets.keys())
                    self._log(f"⚠️ Đã chọn GHI ĐÈ cho {len(overwrite_sheets)} sheet đã tồn tại.", "warning")
                else:
                    self._log(f"ℹ️ Đã chọn BỎ QUA các sheet đã tồn tại trong CSDL.", "info")

        self.is_processing = True
        self.cancel_requested = False
        self.start_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.progress_bar["value"] = 0
        self.progress_bar["maximum"] = len(files_to_process)

        # Reset thống kê
        self.stat_total_lbl.config(text="Tổng học sinh: 0")
        self.stat_da_so_hoa_lbl.config(text="✅ Đã số hóa đầy đủ: 0")
        self.stat_chua_day_du_lbl.config(text="⚠️ Số hoá chưa đầy đủ: 0")
        self.stat_chua_so_hoa_lbl.config(text="❌ Chưa số hoá: 0")

        # Khởi chạy worker thread
        t = threading.Thread(
            target=self._worker_thread,
            args=(files_to_process, base_dir, overwrite_sheets),
            daemon=True
        )
        t.start()

    def _stop_processing(self):
        if self.is_processing:
            ans = messagebox.askyesno("Xác nhận", "Bạn có chắc chắn muốn dừng quá trình xử lý không?", parent=self.root)
            if ans:
                self.cancel_requested = True
                self._log("⚠️ Đã gửi yêu cầu dừng xử lý...", "warning")

    def _worker_thread(
        self,
        files: list[Path],
        base_dir: Path,
        overwrite_sheets: set[tuple[str, str]]
    ):
        try:
            self._log("\n" + "=" * 80, "info")
            self._log(f"🚀 BẮT ĐẦU CHUẨN HÓA DỮ LIỆU & NẠP VÀO CSDL SQLITE", "info")
            self._log(f"📁 Tổng số file Excel phát hiện: {len(files)} tệp.", "info")
            if self.db_manager:
                self._log(f"🗄️ CSDL SQLite đích: {self.db_manager.db_path.name}", "info")
            self._log("=" * 80, "info")

            # 1. Tự động sao lưu an toàn CSDL trước khi nạp batch
            if self.db_manager:
                try:
                    backup_f = self.db_manager.backup_database()
                    self._log(f"🛡️ Đã tự động tạo bản sao lưu an toàn CSDL tại: {backup_f.name}", "success")
                except Exception as e:
                    self._log(f"⚠️ Không thể tạo bản sao lưu CSDL: {e}", "warning")

            def progress_cb(curr_idx, total_cnt, current_stats, stage_msg=None):
                self.root.after(0, lambda: self._update_progress_ui(curr_idx, total_cnt, current_stats, stage_msg))

            def check_cancel_cb():
                return self.cancel_requested

            # 2. Quy trình chuẩn hóa 5 giai đoạn toàn cục
            all_students, overall_stats = self.processor.process_all_files(
                files=files,
                unit_info=None,
                dialog_callback=self._dialog_callback,
                log_callback=self._log,
                progress_callback=progress_cb,
                check_cancel=check_cancel_cb,
                base_input_dir=base_dir
            )

            if check_cancel_cb():
                self._log("⏹ Quá trình đã dừng theo yêu cầu. Dữ liệu chưa ghi vào CSDL.", "warning")
                self.root.after(0, self._on_processing_stopped)
                return

            self.all_unit_students = all_students

            # 3. Nạp danh sách học sinh vào CSDL SQLite
            db_res = {"inserted": 0, "deleted_old": 0}
            if self.db_manager:
                self._log("\n" + "=" * 80, "info")
                self._log(f"💾 [GHI CSDL] Đang lưu {len(all_students)} học sinh vào CSDL SQLite...", "info")
                self._log("=" * 80, "info")
                db_res = self.db_manager.insert_students_batch(
                    all_students,
                    overwrite_sheets=overwrite_sheets,
                    progress_callback=lambda cur, tot, msg: self.root.after(0, lambda: self.status_lbl.config(text=msg))
                )

            total_st = overall_stats["total_students"]
            da_so_hoa = overall_stats["da_so_hoa"]
            chua_day_du = overall_stats["chua_day_du"]
            chua_so_hoa = overall_stats["chua_so_hoa"]

            self._log("\n" + "=" * 80, "info")
            self._log(f"🎉 HOÀN THÀNH XỬ LÝ & NẠP VÀO CSDL SQLITE THÀNH CÔNG!", "success")
            self._log(f"✅ Đã nạp mới / cập nhật: {db_res['inserted']} học sinh.", "success")
            if db_res['deleted_old'] > 0:
                self._log(f"🗑️ Đã làm mới (xóa bản ghi cũ): {db_res['deleted_old']} học sinh.", "info")
            self._log(f"   • Đã số hóa đầy đủ, chính xác: {da_so_hoa}", "success")
            self._log(f"   • Số hoá chưa đầy đủ, chính xác: {chua_day_du}", "warning")
            self._log(f"   • Chưa số hoá: {chua_so_hoa}", "error")
            self._log("=" * 80, "info")

            self.root.after(0, self._on_processing_finished, total_st, da_so_hoa, chua_day_du, chua_so_hoa, db_res)
        except Exception as e:
            self._log(f"\n❌ LỖI TRONG QUÁ TRÌNH XỬ LÝ: {e}", "error")
            import traceback
            traceback.print_exc()
            self.root.after(0, self._on_processing_error, str(e))

    def _on_processing_stopped(self):
        self.status_lbl.config(text="Đã dừng xử lý theo yêu cầu.")
        self.is_processing = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

    def _on_processing_error(self, err_msg: str):
        self.status_lbl.config(text=f"Lỗi: {err_msg}")
        self.is_processing = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        messagebox.showerror(
            "Lỗi xử lý",
            f"Đã xảy ra lỗi trong quá trình xử lý dữ liệu:\n\n{err_msg}\n\nVui lòng kiểm tra lại cấu trúc file hoặc nhật ký lỗi.",
            parent=self.root
        )

    def _update_progress_ui(self, curr_idx, total_cnt, stats, stage_msg=None):
        self.progress_bar["maximum"] = total_cnt
        self.progress_bar["value"] = curr_idx
        if stage_msg:
            self.status_lbl.config(text=stage_msg)
        else:
            self.status_lbl.config(text=f"Đang xử lý [{curr_idx}/{total_cnt}] tệp...")
        self.stat_total_lbl.config(text=f"Tổng học sinh: {stats.get('total_students', 0)}")
        self.stat_da_so_hoa_lbl.config(text=f"✅ Đã số hóa đầy đủ: {stats.get('da_so_hoa', 0)}")
        self.stat_chua_day_du_lbl.config(text=f"⚠️ Số hoá chưa đầy đủ: {stats.get('chua_day_du', 0)}")
        self.stat_chua_so_hoa_lbl.config(text=f"❌ Chưa số hoá: {stats.get('chua_so_hoa', 0)}")

    def _on_processing_finished(self, total_st, da_so_hoa, chua_day_du, chua_so_hoa, db_res):
        self.status_lbl.config(text="Hoàn thành nạp dữ liệu vào CSDL SQLite!")
        self.is_processing = False
        self.start_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)

        # Cập nhật số liệu CSDL trên giao diện
        self._refresh_db_stats_ui()

        # Tự động cập nhật dữ liệu mới nạp sang Phân hệ 2
        if hasattr(self, 'data_mgmt_view') and self.data_mgmt_view:
            try:
                self.data_mgmt_view.refresh_filter_options()
                self.data_mgmt_view.refresh_data(reset_page=True)
            except Exception:
                pass

        ins_count = db_res.get("inserted", total_st)
        del_count = db_res.get("deleted_old", 0)
        del_msg = f"\n• Đã làm mới (ghi đè): {del_count:,} bản ghi cũ" if del_count > 0 else ""

        messagebox.showinfo(
            "Hoàn thành nạp CSDL",
            f"Đã hoàn thành chuẩn hóa và nạp dữ liệu vào CSDL SQLite!\n\n"
            f"• Học sinh đã nạp: {ins_count:,}{del_msg}\n"
            f"• Đã số hóa đầy đủ: {da_so_hoa:,}\n"
            f"• Số hoá chưa đầy đủ: {chua_day_du:,}\n"
            f"• Chưa số hoá: {chua_so_hoa:,}\n\n"
            f"Dữ liệu đã được lưu trữ an toàn vào {self.db_manager.db_path.name}.\n"
            f"Bạn có thể chuyển sang Tab '📊 2. Quản trị, Lọc & Sửa đổi CSDL' để tra cứu và chỉnh sửa trực tiếp!",
            parent=self.root
        )

    def _open_data_inspector(self):
        """Mở cửa sổ tra cứu và đối soát dữ liệu gốc song song."""
        if not self.all_unit_students:
            messagebox.showinfo("Thông báo", "Chưa có dữ liệu học sinh nào được xử lý.", parent=self.root)
            return
        DataInspectorDialog(self.root, self.all_unit_students)

    def _open_output_dir(self):
        """Mở thư mục kết quả trên hệ điều hành."""
        out_path = getattr(self, "last_output_dir", None)
        if not out_path or not out_path.exists():
            raw_path = self.path_var.get().strip().strip("'\"")
            if raw_path:
                p = Path(raw_path)
                cand = (p.parent if p.is_file() else p) / "Ket_Qua_So_Hoa"
                if not cand.exists():
                    cand = (p.parent if p.is_file() else p) / "Da_Chuan_Hoa"
                if cand.exists():
                    out_path = cand

        if not out_path or not out_path.exists():
            messagebox.showinfo("Thông báo", "Chưa có thư mục kết quả Ket_Qua_So_Hoa nào được tạo.", parent=self.root)
            return

        try:
            if sys.platform == "win32":
                os.startfile(str(out_path))
            elif sys.platform == "darwin":
                subprocess.Popen(["open", str(out_path)])
            else:
                subprocess.Popen(["xdg-open", str(out_path)])
        except Exception as e:
            messagebox.showerror("Lỗi mở thư mục", f"Không thể mở thư mục: {e}", parent=self.root)


    def _on_exit(self):
        """Xác nhận an toàn trước khi đóng ứng dụng để tránh bấm nhầm."""
        if self.is_processing:
            if messagebox.askyesno("Đang xử lý dữ liệu", "Hệ thống đang trong quá trình xử lý dữ liệu. Bạn có chắc chắn muốn dừng và thoát?", parent=self.root):
                self.root.destroy()
        else:
            if messagebox.askyesno("Xác nhận thoát", "Bạn có chắc chắn muốn đóng chương trình?", parent=self.root):
                self.root.destroy()

def main():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
