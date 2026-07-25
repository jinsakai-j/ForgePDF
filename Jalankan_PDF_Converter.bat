@echo off
title PDF to Word Converter Launcher
chcp 65001 > nul
cd /d "%~dp0"

echo ============================================================
echo      📄 PDF to Word Converter (Offline Desktop Edition)
echo ============================================================
echo.
echo Sedang memeriksa & menyiapkan environment...

:: Check python installation
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo [ERROR] Python belum terdeteksi di sistem Anda!
    echo Silakan install Python dari https://www.python.org/ atau Microsoft Store.
    echo Pastikan opsi "Add Python to PATH" dicentang saat instalasi.
    pause
    exit /b
)

:: Install required packages automatically if missing
echo Memeriksa modul pendukung (pdf2docx, customtkinter, Pillow)...
python -c "import pdf2docx, customtkinter, PIL, docx" >nul 2>&1
if %errorlevel% neq 0 (
    echo.
    echo Mengunduh & menginstall pustaka yang diperlukan...
    python -m pip install -r requirements.txt --quiet
    if %errorlevel% neq 0 (
        echo [INFO] Memasang dependencies dasar...
        python -m pip install pdf2docx customtkinter Pillow python-docx
    )
)

echo.
echo Menginstansiasi Aplikasi GUI...
start "" pythonw main.py
if %errorlevel% neq 0 (
    python main.py
)

exit
