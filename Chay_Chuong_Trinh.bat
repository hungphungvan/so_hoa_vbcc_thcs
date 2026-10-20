@echo off
chcp 65001 >nul
cd /d "%~dp0"
title CÔNG CỤ XỬ LÝ & SỐ HÓA VĂN BẰNG CHỨNG CHỈ THCS

echo ========================================================================
echo     CÔNG CỤ XỬ LÝ & SỐ HÓA VĂN BẰNG CHỨNG CHỈ THCS    
echo ========================================================================
echo.

if exist .venv\Scripts\python.exe (
    .venv\Scripts\python.exe main.py
    goto end
)

where uv >nul 2>nul
if %errorlevel% equ 0 (
    uv run main.py
    goto end
)

where py >nul 2>nul
if %errorlevel% equ 0 (
    py -m pip install openpyxl -q >nul 2>nul
    py main.py
    goto end
)

where python >nul 2>nul
if %errorlevel% equ 0 (
    python -m pip install openpyxl -q >nul 2>nul
    python main.py
    goto end
)

echo [LỖI] Máy tính chưa cài đặt Python hoặc công cụ UV!
echo Vui lòng cài đặt Python từ https://www.python.org/downloads/ (nhớ tích chọn Add Python to PATH)
echo hoặc cài đặt UV để sử dụng chương trình.
echo.

:end
echo.
pause
