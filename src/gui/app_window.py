from __future__ import annotations

import os
import sys
import subprocess
from pathlib import Path
import tkinter as tk
from tkinter import ttk, messagebox

from src.version import APP_VERSION, APP_NAME
from src.updater import check_for_updates_async, check_for_updates, open_release_page
from src.gui.custom_tools_dialog import CustomToolsView, open_path_in_os
from src.gui.theme import (
    IS_DARK,
    FONT_LABEL,
    FONT_LABEL_BOLD,
    get_theme_palette,
    configure_ttk_theme,
)


class MainWindow:
    """
    Cửa sổ chính của ứng dụng Số hóa Văn bằng THCS.
    Tập trung vào 5 nhóm công cụ cốt lõi:
    - Tab 1: Gộp vào mẫu chuẩn 28 cột
    - Tab 2: Kiểm tra & Chuẩn hoá lỗi dữ liệu
    - Tab 3: Tách / Ghép tệp Excel
    - Tab 4: Đổi đuôi (.doc, .xls) & Chuyển mã sang Unicode
    - Tab 5: OCR trích xuất ảnh sổ sang Excel
    """
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title(f"{APP_NAME} - v{APP_VERSION} (Số hóa Văn bằng THCS)")

        # Đặt kích thước và vị trí cửa sổ tối ưu đa nền tảng
        sw = self.root.winfo_screenwidth()
        sh = self.root.winfo_screenheight()
        win_w = min(1280, max(1040, sw - 40))
        win_h = min(880, max(720, sh - 60))
        x = max(0, (sw - win_w) // 2)
        y = max(0, (sh - win_h - 40) // 2)
        self.root.geometry(f"{win_w}x{win_h}+{x}+{y}")
        self.root.minsize(1000, 660)

        # Quản lý giao diện Sáng / Tối (Dark Mode)
        self.is_dark = IS_DARK
        self.palette = get_theme_palette(self.is_dark)
        self.root.configure(bg=self.palette.WINDOW_BG)

        # Cấu hình style ttk
        self.style = ttk.Style(self.root)
        self._setup_styles()

        self.latest_update_info = None

        self._build_ui()

        # Cập nhật layout và kích hoạt khắc phục lỗi Cocoa redraw lag trên macOS
        self.root.update_idletasks()
        self.root.lift()
        if sys.platform == "darwin":
            self.root.after(80, self._force_macos_redraw)

        # Kiểm tra cập nhật ngầm sau 3 giây
        self.root.after(3000, self._check_update_silent)

    def _force_macos_redraw(self):
        """Khắc phục triệt để lỗi macOS Tkinter không vẽ nội dung cho đến khi người dùng resize."""
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
        self.root.configure(bg=self.palette.WINDOW_BG)
        configure_ttk_theme(self.style, self.palette)

        if hasattr(self, 'custom_view'):
            self.custom_view.apply_theme(self.is_dark)

        if hasattr(self, 'theme_btn'):
            self.theme_btn.configure(text="☀️ Chế độ Sáng (Đổi 🌙)" if not self.is_dark else "🌙 Chế độ Tối (Đổi ☀️)")

    def _build_ui(self):
        # 1. Bottom Bar (GHIM CỐ ĐỊNH Ở ĐÁY CỬA SỔ - LUÔN LUÔN NHÌN THẤY)
        bottom_frame = ttk.Frame(self.root, style="Window.TFrame", padding=(12, 6))
        bottom_frame.pack(side=tk.BOTTOM, fill=tk.X)

        self.open_output_btn = ttk.Button(
            bottom_frame,
            text="📂 Mở thư mục chứa file",
            command=self._open_output_dir
        )
        self.open_output_btn.pack(side=tk.LEFT, padx=(0, 6))

        self.open_file_btn = ttk.Button(
            bottom_frame,
            text="📄 Mở file Excel vừa tạo",
            command=self._open_last_file
        )
        self.open_file_btn.pack(side=tk.LEFT, padx=6)

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

        # 2. Main Content View: Khởi tạo Bộ Công Cụ & Tiện Ích Toàn Diện
        self.custom_view = CustomToolsView(self.root, is_dark=self.is_dark)
        self.custom_view.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

    def _open_last_file(self):
        """Mở tệp Excel được sinh ra gần nhất của ứng dụng."""
        last_f = getattr(self.custom_view, "last_created_file", None)
        if not last_f:
            last_f = getattr(self.custom_view, "last_inspect_output_file", None)
        if not last_f:
            last_f = getattr(self.custom_view, "last_smart_output_file", None)
        if not last_f:
            last_f = getattr(self.custom_view, "last_merge_output_file", None)
        if not last_f:
            last_f = getattr(self.custom_view, "last_encoding_output_file", None)
        if not last_f:
            last_f = getattr(self.custom_view, "last_ocr_output_file", None)

        if last_f and Path(last_f).exists():
            open_path_in_os(Path(last_f))
        else:
            messagebox.showinfo("Thông báo", "Chưa có tệp Excel nào được tạo gần đây.\nVui lòng thực hiện một thao tác để tạo tệp kết quả.", parent=self.root)

    def _open_output_dir(self):
        """Mở thư mục chứa tệp kết quả gần nhất trên hệ điều hành."""
        last_f = getattr(self.custom_view, "last_created_file", None)
        if last_f and Path(last_f).exists():
            target_dir = Path(last_f).parent
        else:
            target_dir = Path.cwd() / "data"
            if not target_dir.exists():
                target_dir = Path.cwd()

        open_path_in_os(target_dir)

    def _check_update_silent(self):
        """Kiểm tra bản cập nhật ngầm mà không làm phiền người dùng nếu chưa có bản mới."""
        def on_done(info):
            if info and info.get("has_update"):
                self.latest_update_info = info
                tag = info.get("latest_version", "")
                self.root.after(0, lambda: self.update_btn.config(
                    text=f"⭐ Có bản mới {tag}",
                    style="Action.TButton"
                ))

        check_for_updates_async(on_done)

    def _check_update_manually(self):
        """Kiểm tra bản cập nhật thủ công khi người dùng click nút."""
        self.update_btn.config(text="⏳ Đang kiểm tra...")
        self.root.update_idletasks()

        info = check_for_updates()
        if info and info.get("has_update"):
            self.latest_update_info = info
            tag = info.get("latest_version", "")
            self.update_btn.config(text=f"⭐ Có bản mới {tag}")
            if messagebox.askyesno(
                "Đã có phiên bản mới!",
                f"Đã có phiên bản mới {tag} trên GitHub!\n\n"
                f"Bạn có muốn mở trang tải bản cập nhật mới không?",
                parent=self.root
            ):
                open_release_page()
        else:
            self.update_btn.config(text=f"🔄 Bản {APP_VERSION}")
            messagebox.showinfo(
                "Kiểm tra cập nhật",
                f"Bạn đang sử dụng phiên bản mới nhất ({APP_VERSION})!",
                parent=self.root
            )

    def _on_exit(self):
        """Xác nhận an toàn trước khi đóng ứng dụng."""
        if messagebox.askyesno("Xác nhận thoát", "Bạn có chắc chắn muốn đóng chương trình?", parent=self.root):
            self.root.destroy()


def main():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
