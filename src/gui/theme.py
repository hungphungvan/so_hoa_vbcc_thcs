from __future__ import annotations

import sys
import subprocess

def is_dark_mode() -> bool:
    """Kiểm tra hệ thống có đang bật Dark Mode hay không (hỗ trợ macOS và Windows)."""
    if sys.platform == "darwin":
        try:
            res = subprocess.run(
                ["defaults", "read", "-g", "AppleInterfaceStyle"],
                capture_output=True,
                text=True,
                timeout=1
            )
            return "Dark" in res.stdout
        except Exception:
            return False
    elif sys.platform == "win32":
        try:
            import winreg
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize"
            )
            value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
            return value == 0
        except Exception:
            return False
    return False

IS_DARK = is_dark_mode()

# Phông chữ đa nền tảng
FONT_FAMILY = "Helvetica Neue" if sys.platform == "darwin" else "Segoe UI"
FONT_FAMILY_MONO = "Menlo" if sys.platform == "darwin" else "Consolas"

# Bảng màu tương thích Dark Mode & Light Mode
if IS_DARK:
    CLR_PRIMARY = "#64B5F6"      # Xanh dương sáng (Tiêu đề card, mã đơn vị, tên đơn vị, trạng thái)
    CLR_ACCENT = "#CE93D8"       # Tím nhạt sáng (Xã/phường)
    CLR_SUCCESS = "#81C784"      # Xanh lá sáng (Số trường, Đã số hóa đầy đủ)
    CLR_WARNING = "#FFB74D"      # Cam vàng sáng (Số hoá chưa đầy đủ)
    CLR_DANGER = "#EF5350"       # Đỏ hồng sáng (Chưa số hoá, lỗi)
    CLR_MUTED = "#B0BEC5"        # Xám bạc sáng (Chữ chú thích, hướng dẫn, danh sách trường)
    CLR_TEXT = "#ECEFF1"         # Chữ nội dung sáng
    CLR_WINDOW_BG = "#2D2D2D"    # Nền cửa sổ Dark Mode
    CLR_LOG_BG = "#1E1E1E"       # Nền console nhật ký
    CLR_LOG_FG = "#ECEFF1"       # Chữ console nhật ký
    CLR_HIGHLIGHT_BG = "#3E2723" # Nền dòng nổi bật trong bảng
    CLR_HIGHLIGHT_FG = "#FFE082" # Chữ dòng nổi bật trong bảng
else:
    CLR_PRIMARY = "#0D47A1"
    CLR_ACCENT = "#7B1FA2"
    CLR_SUCCESS = "#2E7D32"
    CLR_WARNING = "#EF6C00"
    CLR_DANGER = "#C62828"
    CLR_MUTED = "#555555"
    CLR_TEXT = "#212121"
    CLR_WINDOW_BG = "#F5F5F7"
    CLR_LOG_BG = "#FFFFFF"
    CLR_LOG_FG = "#000000"
    CLR_HIGHLIGHT_BG = "#FFF3E0"
    CLR_HIGHLIGHT_FG = "#BF360C"
