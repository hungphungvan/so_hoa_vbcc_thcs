#!/bin/bash
cd "$(dirname "$0")"

# Nạp đường dẫn môi trường trên macOS Apple Silicon và Intel
export PATH="/opt/homebrew/bin:/usr/local/bin:$HOME/.local/bin:$PATH"
export TK_SILENCE_DEPRECATION=1

# Cưỡng chế chạy chế độ Aqua (Light Mode) cho Python để khắc phục hoàn toàn lỗi màn hình đen của Tkinter trên macOS Dark Mode
defaults write org.python.python NSRequiresAquaSystemAppearance -bool true 2>/dev/null || true
defaults write com.apple.python3 NSRequiresAquaSystemAppearance -bool true 2>/dev/null || true

echo "========================================================================"
echo "    CÔNG CỤ XỬ LÝ & SỐ HÓA VĂN BẰNG CHỨNG CHỈ THCS    "
echo "========================================================================"
echo ""

# Tự động đồng bộ mã nguồn mới nhất từ GitHub
if [ -d ".git" ] && command -v git >/dev/null 2>&1; then
    echo "--> Đang kiểm tra cập nhật mới nhất từ GitHub..."
    if git pull --quiet 2>/dev/null; then
        echo "✅ Đã đồng bộ mã nguồn mới nhất thành công!"
    else
        echo "ℹ️ Đang chạy phiên bản hiện tại trên máy."
    fi
    echo ""
fi

# Ưu tiên UV để đảm bảo chạy trên Python 3.12 kèm Tkinter 9.0 hiện đại
if command -v uv >/dev/null 2>&1; then
    if [ ! -d ".venv" ] || [ ! -f ".venv/bin/python" ]; then
        echo "--> Đang tự động cấu hình môi trường Python 3.12 (Tkinter 9.0 hiện đại)..."
        uv venv --python 3.12 .venv >/dev/null 2>&1
        uv pip install openpyxl --python .venv/bin/python >/dev/null 2>&1
    fi
    .venv/bin/python main.py
elif [ -d ".venv" ] && [ -f ".venv/bin/python" ]; then
    .venv/bin/python main.py
else
    python3 -c "import openpyxl" 2>/dev/null || python3 -m pip install openpyxl -q 2>/dev/null
    python3 main.py
fi

echo ""
read -p "Nhấn Enter để đóng cửa sổ này..."
