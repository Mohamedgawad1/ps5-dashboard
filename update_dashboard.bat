@echo off
chcp 65001 >nul
title PS5 Dashboard - Cloud Sync & Update
color 0B
cd /d "%~dp0"
echo ============================================
echo  PS5 DASHBOARD - CLOUD SYNC & UPDATE
echo  %date% %time%
echo ============================================
echo.

echo [1/4] Syncing latest data from OneDrive cloud...
python -c "import sys; sys.path.insert(0,'.'); from cpp_agi_dashboard import sync_cloud_files; sync_cloud_files()"
if %errorlevel% neq 0 (
    echo   [WARN] Cloud sync failed, continuing with local files...
)

echo.
echo [2/4] Building dashboard...
python cpp_agi_dashboard.py
if %errorlevel% neq 0 goto :err

echo.
echo [3/4] Committing and pushing to GitHub...
git add -u
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo   No changes to upload.
) else (
    git -c user.name=ps5-bot -c user.email=ps5@local commit -m "cloud update %date% %time%"
    git push origin main
    if %errorlevel% neq 0 goto :err
)

echo.
echo [4/4] Done! Opening platform...
start "" "https://mohamedgawad1.github.io/ps5-dashboard/"
timeout /t 10 >nul
exit /b 0

:err
echo.
echo  ERROR! Check messages above.
pause
exit /b 1
