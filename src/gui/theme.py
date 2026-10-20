from __future__ import annotations

import sys
import subprocess
from tkinter import ttk

def is_system_dark_mode() -> bool:
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

# Trạng thái Dark Mode ban đầu theo hệ thống
IS_DARK = is_system_dark_mode()

# Phông chữ đa nền tảng
FONT_FAMILY = "Helvetica Neue" if sys.platform == "darwin" else "Segoe UI"
FONT_FAMILY_MONO = "Menlo" if sys.platform == "darwin" else "Consolas"

# Thang kích thước phông chữ chuẩn (Đã loại bỏ hoàn toàn cỡ chữ 8pt, tối thiểu 10pt)
FONT_TITLE = (FONT_FAMILY, 15, "bold")
FONT_SUBTITLE = (FONT_FAMILY, 10)
FONT_CARD_TITLE = (FONT_FAMILY, 11, "bold")
FONT_LABEL = (FONT_FAMILY, 10)
FONT_LABEL_BOLD = (FONT_FAMILY, 10, "bold")
FONT_DATA_BOLD = (FONT_FAMILY, 10, "bold")
FONT_MUTED = (FONT_FAMILY, 10)
FONT_BUTTON = (FONT_FAMILY, 10, "bold")
FONT_BUTTON_REG = (FONT_FAMILY, 10)
FONT_LOG = (FONT_FAMILY_MONO, 10)
FONT_TREE = (FONT_FAMILY, 10)
FONT_TREE_BOLD = (FONT_FAMILY, 10, "bold")
FONT_TREE_HEADING = (FONT_FAMILY, 10, "bold")

class ThemePalette:
    """Bảng màu tương thích Light Mode & Dark Mode với độ tương phản cực cao (High Contrast)."""
    def __init__(self, is_dark: bool):
        self.is_dark = is_dark
        if is_dark:
            # DARK MODE (Hiện đại, tương phản cao, dịu mắt)
            self.WINDOW_BG = "#18181B"          # Nền cửa sổ Dark Mode (Zinc-900)
            self.CARD_BG = "#27272A"            # Nền khối Card (Zinc-800)
            self.CARD_BORDER = "#3F3F46"        # Viền Card (Zinc-700)
            self.HEADER_BANNER_BG = "#0F172A"   # Nền banner trên cùng (Slate-900)
            self.HEADER_BANNER_FG = "#F8FAFC"   # Chữ banner trên cùng (Trắng sáng)
            self.HEADER_SUB_FG = "#94A3B8"      # Chữ tiêu đề phụ
            self.PRIMARY = "#60A5FA"            # Xanh dương sáng (Blue-400)
            self.ACCENT = "#C084FC"             # Tím nhạt sáng (Purple-400)
            self.SUCCESS = "#4ADE80"            # Xanh lá sáng (Green-400)
            self.WARNING = "#FB923C"            # Cam sáng (Orange-400)
            self.DANGER = "#F87171"             # Đỏ hồng sáng (Red-400)
            self.MUTED = "#D1D5DB"              # Xám bạc sáng (Gray-300, tương phản rất cao)
            self.TEXT = "#F9FAFB"               # Chữ nội dung sáng (Gray-50)
            self.ENTRY_BG = "#18181B"           # Nền ô nhập liệu
            self.ENTRY_FG = "#F9FAFB"           # Chữ ô nhập liệu
            self.LOG_BG = "#121214"             # Nền console nhật ký
            self.LOG_FG = "#F9FAFB"             # Chữ console nhật ký
            self.HIGHLIGHT_BG = "#451A03"       # Nền dòng nổi bật trong bảng
            self.HIGHLIGHT_FG = "#FDE68A"       # Chữ dòng nổi bật trong bảng
            self.TREE_BG = "#27272A"            # Nền Treeview
            self.TREE_FG = "#F9FAFB"            # Chữ Treeview
            self.TREE_HEADING_BG = "#3F3F46"    # Tiêu đề cột Treeview
            self.TREE_HEADING_FG = "#F9FAFB"    # Chữ tiêu đề cột
            self.TREE_SELECT_BG = "#2563EB"     # Dòng chọn Treeview
            self.TREE_SELECT_FG = "#FFFFFF"     # Chữ dòng chọn Treeview
            self.BUTTON_BG = "#3F3F46"          # Nền nút bấm
            self.BUTTON_FG = "#F9FAFB"          # Chữ nút bấm
        else:
            # LIGHT MODE (Trắng sáng tinh khôi, sạch sẽ, tương phản cao)
            self.WINDOW_BG = "#F0F2F5"          # Nền cửa sổ Light Mode (Xám xanh rất nhạt)
            self.CARD_BG = "#FFFFFF"            # Nền khối Card trắng tinh khôi
            self.CARD_BORDER = "#D1D5DB"        # Viền Card xám sáng
            self.HEADER_BANNER_BG = "#1565C0"   # Nền banner trên cùng (Xanh dương Hoàng gia)
            self.HEADER_BANNER_FG = "#FFFFFF"   # Chữ banner trên cùng
            self.HEADER_SUB_FG = "#E3F2FD"      # Chữ tiêu đề phụ
            self.PRIMARY = "#0D47A1"            # Xanh dương Hoàng gia đậm (tương phản cao)
            self.ACCENT = "#6B21A8"             # Tím đậm
            self.SUCCESS = "#15803D"            # Xanh lục đậm (tương phản cao)
            self.WARNING = "#C2410C"            # Cam cháy đậm (dễ đọc)
            self.DANGER = "#B91C1C"             # Đỏ thẫm (tương phản mạnh)
            self.MUTED = "#374151"              # Xanh xám đậm (tương phản cao, không xám mờ)
            self.TEXT = "#111827"               # Đen than đậm
            self.ENTRY_BG = "#FFFFFF"           # Nền ô nhập liệu
            self.ENTRY_FG = "#111827"           # Chữ ô nhập liệu
            self.LOG_BG = "#FFFFFF"             # Nền console nhật ký
            self.LOG_FG = "#111827"             # Chữ console nhật ký
            self.HIGHLIGHT_BG = "#FEF3C7"       # Nền dòng nổi bật trong bảng
            self.HIGHLIGHT_FG = "#92400E"       # Chữ dòng nổi bật trong bảng
            self.TREE_BG = "#FFFFFF"            # Nền Treeview
            self.TREE_FG = "#111827"            # Chữ Treeview
            self.TREE_HEADING_BG = "#E5E7EB"    # Tiêu đề cột Treeview
            self.TREE_HEADING_FG = "#111827"    # Chữ tiêu đề cột
            self.TREE_SELECT_BG = "#1D4ED8"     # Dòng chọn Treeview
            self.TREE_SELECT_FG = "#FFFFFF"     # Chữ dòng chọn Treeview
            self.BUTTON_BG = "#E5E7EB"          # Nền nút bấm
            self.BUTTON_FG = "#111827"          # Chữ nút bấm

def get_theme_palette(is_dark: bool | None = None) -> ThemePalette:
    """Lấy bảng màu theo trạng thái Sáng / Tối."""
    if is_dark is None:
        is_dark = IS_DARK
    return ThemePalette(is_dark)

# Bảng màu mặc định hiện thời
current_palette = get_theme_palette(IS_DARK)

CLR_PRIMARY = current_palette.PRIMARY
CLR_ACCENT = current_palette.ACCENT
CLR_SUCCESS = current_palette.SUCCESS
CLR_WARNING = current_palette.WARNING
CLR_DANGER = current_palette.DANGER
CLR_MUTED = current_palette.MUTED
CLR_TEXT = current_palette.TEXT
CLR_WINDOW_BG = current_palette.WINDOW_BG
CLR_CARD_BG = current_palette.CARD_BG
CLR_CARD_BORDER = current_palette.CARD_BORDER
CLR_ENTRY_BG = current_palette.ENTRY_BG
CLR_ENTRY_FG = current_palette.ENTRY_FG
CLR_LOG_BG = current_palette.LOG_BG
CLR_LOG_FG = current_palette.LOG_FG
CLR_HIGHLIGHT_BG = current_palette.HIGHLIGHT_BG
CLR_HIGHLIGHT_FG = current_palette.HIGHLIGHT_FG
CLR_HEADER_BANNER_BG = current_palette.HEADER_BANNER_BG
CLR_HEADER_BANNER_FG = current_palette.HEADER_BANNER_FG
CLR_HEADER_SUB_FG = current_palette.HEADER_SUB_FG
CLR_TREE_BG = current_palette.TREE_BG
CLR_TREE_FG = current_palette.TREE_FG
CLR_TREE_HEADING_BG = current_palette.TREE_HEADING_BG
CLR_TREE_HEADING_FG = current_palette.TREE_HEADING_FG
CLR_TREE_SELECT_BG = current_palette.TREE_SELECT_BG
CLR_TREE_SELECT_FG = current_palette.TREE_SELECT_FG

def configure_ttk_theme(style: ttk.Style, palette: ThemePalette):
    """Cấu hình đồng bộ toàn diện cho các widget ttk theo bảng màu palette."""
    p = palette

    # Thiết lập theme cơ sở: Trên Windows/Linux dùng 'clam' có thể tùy biến 100% màu sắc
    if sys.platform != "darwin":
        try:
            style.theme_use('clam')
        except Exception:
            pass

    # 1. Cấu hình toàn cục (.)
    style.configure(".",
        background=p.WINDOW_BG,
        foreground=p.TEXT,
        font=FONT_LABEL
    )

    # 2. Frame
    style.configure("TFrame", background=p.CARD_BG)
    style.configure("Window.TFrame", background=p.WINDOW_BG)
    style.configure("Card.TFrame", background=p.CARD_BG)

    # 3. LabelFrame (Card)
    style.configure("TLabelframe",
        background=p.CARD_BG,
        bordercolor=p.CARD_BORDER,
        darkcolor=p.CARD_BORDER,
        lightcolor=p.CARD_BORDER,
        relief="solid",
        borderwidth=1
    )
    style.configure("TLabelframe.Label",
        background=p.CARD_BG,
        foreground=p.PRIMARY,
        font=FONT_CARD_TITLE
    )
    style.configure("Card.TLabelframe",
        background=p.CARD_BG,
        bordercolor=p.CARD_BORDER,
        darkcolor=p.CARD_BORDER,
        lightcolor=p.CARD_BORDER,
        relief="solid",
        borderwidth=1
    )
    style.configure("Card.TLabelframe.Label",
        background=p.CARD_BG,
        foreground=p.PRIMARY,
        font=FONT_CARD_TITLE
    )

    # 4. Label
    style.configure("TLabel",
        background=p.CARD_BG,
        foreground=p.TEXT,
        font=FONT_LABEL
    )
    style.configure("Window.TLabel",
        background=p.WINDOW_BG,
        foreground=p.TEXT,
        font=FONT_LABEL
    )
    style.configure("Muted.TLabel",
        background=p.CARD_BG,
        foreground=p.MUTED,
        font=FONT_MUTED
    )
    style.configure("Primary.TLabel",
        background=p.CARD_BG,
        foreground=p.PRIMARY,
        font=FONT_LABEL_BOLD
    )
    style.configure("Success.TLabel",
        background=p.CARD_BG,
        foreground=p.SUCCESS,
        font=FONT_LABEL_BOLD
    )
    style.configure("Warning.TLabel",
        background=p.CARD_BG,
        foreground=p.WARNING,
        font=FONT_LABEL_BOLD
    )
    style.configure("Danger.TLabel",
        background=p.CARD_BG,
        foreground=p.DANGER,
        font=FONT_LABEL_BOLD
    )

    # 5. Checkbutton
    style.configure("TCheckbutton",
        background=p.CARD_BG,
        foreground=p.TEXT,
        font=FONT_LABEL
    )
    style.map("TCheckbutton",
        background=[('active', p.CARD_BG)],
        foreground=[('active', p.PRIMARY)]
    )

    # 6. Button
    style.configure("TButton",
        font=FONT_BUTTON_REG,
        padding=5
    )
    style.configure("Action.TButton",
        font=FONT_BUTTON,
        padding=6
    )
    style.configure("Inspect.TButton",
        font=FONT_BUTTON,
        foreground=p.PRIMARY,
        padding=6
    )
    style.configure("Success.TButton",
        font=FONT_BUTTON,
        background="#15803D" if not p.is_dark else "#166534",
        foreground="#FFFFFF",
        padding=6
    )
    style.configure("Stop.TButton",
        font=FONT_BUTTON,
        background="#B91C1C" if not p.is_dark else "#991B1B",
        foreground="#FFFFFF",
        padding=6
    )

    # 7. Treeview
    style.configure("Treeview",
        background=p.TREE_BG,
        foreground=p.TREE_FG,
        fieldbackground=p.TREE_BG,
        font=FONT_TREE,
        rowheight=26
    )
    style.map("Treeview",
        background=[('selected', p.TREE_SELECT_BG)],
        foreground=[('selected', p.TREE_SELECT_FG)]
    )
    style.configure("Treeview.Heading",
        background=p.TREE_HEADING_BG,
        foreground=p.TREE_HEADING_FG,
        font=FONT_TREE_HEADING,
        padding=5
    )

    # 8. Entry & Combobox
    style.configure("TEntry",
        fieldbackground=p.ENTRY_BG,
        foreground=p.ENTRY_FG,
        font=FONT_LABEL
    )
    style.configure("TCombobox",
        fieldbackground=p.ENTRY_BG,
        foreground=p.ENTRY_FG,
        background=p.WINDOW_BG,
        font=FONT_LABEL
    )
