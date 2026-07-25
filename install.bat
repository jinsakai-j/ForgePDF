@echo off
title Installer ForgePDF - Setup dan Desktop Shortcut
echo ===================================================
echo   Menyiapkan dan Menginstall ForgePDF di PC Anda
echo ===================================================
echo.

:: 1. Install required Python packages
echo [1/2] Menginstall pustaka Python yang dibutuhkan...
python -m pip install --upgrade pip
python -m pip install -r requirements.txt

echo.
:: 2. Create VBS Launcher and Desktop Shortcut automatically
echo [2/2] Membuat Shortcut ForgePDF di Desktop Anda...

set "TARGET_DIR=%~dp0"
set "TARGET_DIR=%TARGET_DIR:~0,-1%"

:: Create ForgePDF.vbs launcher in project folder
(
echo Set WshShell = CreateObject^("WScript.Shell"^)^
echo WshShell.CurrentDirectory = "%TARGET_DIR%"
echo WshShell.Run "pythonw.exe main.py", 0, False
) > "%TARGET_DIR%\ForgePDF.vbs"

:: Create Desktop Shortcut using PowerShell
powershell -Command "$userDesktop = [Environment]::GetFolderPath('Desktop'); $s = (New-Object -ComObject WScript.Shell).CreateShortcut(\"$userDesktop\ForgePDF.lnk\"); $s.TargetPath = 'C:\Windows\System32\wscript.exe'; $s.Arguments = '\"%TARGET_DIR%\ForgePDF.vbs\"'; $s.WorkingDirectory = '%TARGET_DIR%'; $s.Save()"

echo.
echo ===================================================
echo   SUKSES! ForgePDF Berhasil Terinstall
echo   Shortcut 'ForgePDF' telah dibuat di Desktop Anda.
echo ===================================================
echo.
pause
