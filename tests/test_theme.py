import unittest
import tkinter as tk
from tkinter import ttk

from src.gui.theme import (
    get_theme_palette,
    configure_ttk_theme,
    FONT_TITLE,
    FONT_SUBTITLE,
    FONT_CARD_TITLE,
    FONT_LABEL,
    FONT_LABEL_BOLD,
    FONT_DATA_BOLD,
    FONT_MUTED,
    FONT_BUTTON,
    FONT_LOG,
    FONT_TREE,
    FONT_TREE_HEADING
)

class TestThemeSystem(unittest.TestCase):
    def test_font_sizes_at_least_10(self):
        """Kiểm tra toàn bộ hệ thống phông chữ đã được nâng chuẩn lên tối thiểu 10pt (không còn chữ 8pt, 9pt)."""
        fonts_to_check = [
            FONT_TITLE, FONT_SUBTITLE, FONT_CARD_TITLE,
            FONT_LABEL, FONT_LABEL_BOLD, FONT_DATA_BOLD,
            FONT_MUTED, FONT_BUTTON, FONT_LOG,
            FONT_TREE, FONT_TREE_HEADING
        ]
        for f in fonts_to_check:
            self.assertIsInstance(f, tuple)
            self.assertGreaterEqual(f[1], 10, f"Phông chữ {f} có kích cỡ < 10pt!")

    def test_palette_light_and_dark(self):
        """Kiểm tra màu sắc tương phản cao của cả 2 bảng màu Light và Dark."""
        light = get_theme_palette(is_dark=False)
        self.assertFalse(light.is_dark)
        self.assertEqual(light.WINDOW_BG, "#F0F2F5")
        self.assertEqual(light.CARD_BG, "#FFFFFF")
        self.assertEqual(light.TEXT, "#111827")
        self.assertEqual(light.MUTED, "#374151")

        dark = get_theme_palette(is_dark=True)
        self.assertTrue(dark.is_dark)
        self.assertEqual(dark.WINDOW_BG, "#18181B")
        self.assertEqual(dark.CARD_BG, "#27272A")
        self.assertEqual(dark.TEXT, "#F9FAFB")
        self.assertEqual(dark.MUTED, "#D1D5DB")

    def test_configure_ttk_theme_execution(self):
        """Kiểm tra hàm configure_ttk_theme thực thi thành công không phát sinh lỗi Tkinter."""
        root = tk.Tk()
        style = ttk.Style(root)
        try:
            # Test áp dụng Light Mode
            light_palette = get_theme_palette(is_dark=False)
            configure_ttk_theme(style, light_palette)

            # Test áp dụng Dark Mode
            dark_palette = get_theme_palette(is_dark=True)
            configure_ttk_theme(style, dark_palette)
        finally:
            root.destroy()

if __name__ == "__main__":
    unittest.main()
