@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title KEO LEN GITHUB - DAY CODE VA DU LIEU (SCT)
color 0a

:: Chuyen thu muc hien tai den dung noi chua file bat
cd /d "%~dp0"

echo ======================================================================
echo     FPT MVNO - HE THONG KEO LEN GITHUB (DAY CODE VA DU LIEU)
echo ======================================================================
echo.
echo [*] Dang chuan bi dong bo va day du lieu len GitHub (SCT)...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\keo_len.ps1"

echo.
echo Nhan phim bat ky de thoat...
pause >nul
