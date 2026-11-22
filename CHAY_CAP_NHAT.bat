@echo off
setlocal EnableDelayedExpansion
chcp 65001 >nul
title CAP NHAT KET QUA THI DUA SIM FPT
color 0b

:: Chuyen thu muc hien tai den dung noi chua file bat
cd /d "%~dp0"

echo ======================================================================
echo     FPT MVNO - HE THONG CAP NHAT BANG VINH DANH THI DUA SIM FPT 3 TUOI
echo ======================================================================
echo.
echo [*] Dang quet du lieu tu file Excel trong thu muc input_data...
echo [*] Dang tinh toan xep hang va render Infographic...
echo.

:: Chay Powershell an toan bang cach goi script tu thu muc hien tai (hoat dong 100% tren moi may)
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "scripts\update_infographic.ps1"

if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [THONG BAO] Neu gap loi ve quyen, ban co the bam chuot phai vao file nay chon 'Run as administrator'
)

echo.
echo Nhan phim bat ky de thoat...
pause >nul
