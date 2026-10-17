#!/bin/bash
cd "$(dirname "$0")"

# Nạp đường dẫn môi trường trên macOS Apple Silicon và Intel
export PATH="/opt/homebrew/bin:/usr/local/bin:$HOME/.local/bin:$PATH"

echo "========================================================================"
echo "    CÔNG CỤ XỬ LÝ & SỐ HÓA VĂN BẰNG CHỨNG CHỈ THCS    "
echo "========================================================================"
echo ""

if command -v uv >/dev/null 2>&1; then
    uv run main.py
else
    python3 main.py
fi

echo ""
read -p "Nhấn Enter để đóng cửa sổ này..."
