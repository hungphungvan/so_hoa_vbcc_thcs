# /// script
# requires-python = ">=3.9"
# dependencies = [
#     "openpyxl",
#     "xlrd",
#     "python-docx",
# ]
# ///
import os
import sys
from pathlib import Path

# Tắt cảnh báo Tkinter cũ trên macOS
os.environ["TK_SILENCE_DEPRECATION"] = "1"

# Đảm bảo mã hóa UTF-8 cho stdout trên mọi nền tảng
if sys.stdout and hasattr(sys.stdout, 'reconfigure'):
    try:
        sys.stdout.reconfigure(encoding='utf-8')
    except Exception:
        pass

# Thêm thư mục gốc vào sys.path để import src
ROOT_DIR = Path(__file__).resolve().parent
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

# Kích hoạt nhận diện High-DPI trên Windows giúp phông chữ và giao diện sắc nét tuyệt đối
if sys.platform == "win32":
    try:
        import ctypes
        # 1: PROCESS_SYSTEM_DPI_AWARE, 2: PROCESS_PER_MONITOR_DPI_AWARE
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        try:
            ctypes.windll.user32.SetProcessDPIAware()
        except Exception:
            pass

def main():
    """Khởi chạy giao diện đồ họa (GUI) cho Mô-đun CHUẨN HOÁ DỮ LIỆU."""
    import tkinter as tk
    from src.gui.app_window import MainWindow

    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()

if __name__ == "__main__":
    main()
