@echo off
chcp 65001 >nul
title PS5 Dashboard - Update & Direct Upload
color 0B
cd /d "%~dp0"
echo ============================================
echo  PS5 DASHBOARD - UPDATE AND UPLOAD DIRECTLY
echo  %date% %time%
echo ============================================
echo.

echo [1/3] Building platform...
python cpp_agi_dashboard.py
if %errorlevel% neq 0 goto :err

echo.
echo [2/3] Committing and uploading directly...
git add -u
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo   No changes to upload.
) else (
    git -c user.name=ps5-bot -c user.email=ps5@local commit -m "platform update %date% %time%"
)
git push origin main
if %errorlevel% neq 0 goto :err

echo.
echo [3/3] Done! Opening platform...
start "" "https://mohamedgawad1.github.io/ps5-dashboard/"
timeout /t 10 >nul
exit /b 0

:err
echo.
echo  ERROR! Check messages above (file open in Excel? internet?).
pause
exit /b 1
