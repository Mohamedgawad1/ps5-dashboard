@echo off
title PS5 Master Tracker Auto Update
cd /d "C:\Users\mylap\OneDrive\Desktop\dashboard"
echo ========================================
echo  PS5 MASTER TRACKER - AUTO UPDATE
echo ========================================
echo.
:retry
python "C:\Users\mylap\OneDrive\Desktop\dashboard\update_master.py"
if %errorlevel% equ 0 (
    echo.
    echo  Done! File updated successfully.
) else (
    echo.
    echo  ERROR: Update failed. The Excel file may be open.
    set /p ans=Close Excel then press R to retry, or any key to exit... 
    if /i "%ans%"=="R" goto retry
)
echo.
pause
