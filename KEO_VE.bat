@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title KEO VE MAY - DONG BO DU LIEU TU GITHUB (SCT)
color 0b

:: Chuyen thu muc hien tai den dung noi chua file bat
cd /d "%~dp0"

echo ======================================================================
echo     FPT MVNO - HE THONG KEO VE MAY (DONG BO CODE TU GITHUB)
echo ======================================================================
echo.
echo [*] Dang ket noi va tai ban moi nhat tu GitHub ve may nay...
echo.

powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\keo_ve.ps1"

echo.
echo Nhan phim bat ky de thoat...
pause >nul
