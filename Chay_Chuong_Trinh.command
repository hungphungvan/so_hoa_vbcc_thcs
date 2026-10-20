#!/bin/bash
cd "$(dirname "$0")"

# Nạp đường dẫn môi trường trên macOS Apple Silicon và Intel
export PATH="/opt/homebrew/bin:/usr/local/bin:$HOME/.local/bin:$PATH"
export TK_SILENCE_DEPRECATION=1

echo "========================================================================"
echo "    CÔNG CỤ XỬ LÝ & SỐ HÓA VĂN BẰNG CHỨNG CHỈ THCS    "
echo "========================================================================"
echo ""

if [ -d ".venv" ] && [ -f ".venv/bin/python" ]; then
    .venv/bin/python main.py
elif command -v uv >/dev/null 2>&1; then
    uv run main.py
else
    python3 -c "import openpyxl" 2>/dev/null || python3 -m pip install openpyxl -q 2>/dev/null
    python3 main.py
fi

echo ""
read -p "Nhấn Enter để đóng cửa sổ này..."
