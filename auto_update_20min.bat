@echo off
chcp 65001 >nul
cd /d "%~dp0"
set LOG=%TEMP%\ps5_auto_update.log
echo [%date% %time%] START >> "%LOG%"

python -c "import sys; sys.path.insert(0,'.'); from cpp_agi_dashboard import sync_cloud_files; sync_cloud_files()" >> "%LOG%" 2>&1

python cpp_agi_dashboard.py >> "%LOG%" 2>&1
if %errorlevel% neq 0 (
    echo [%date% %time%] BUILD FAILED >> "%LOG%"
    exit /b 1
)

git add -u
git diff --cached --quiet
if %errorlevel% equ 0 (
    echo [%date% %time%] No changes >> "%LOG%"
) else (
    git -c user.name=ps5-bot -c user.email=ps5@local commit -m "auto update %date% %time%" >> "%LOG%" 2>&1
    git push origin main >> "%LOG%" 2>&1
    echo [%date% %time%] pushed >> "%LOG%"
)
echo [%date% %time%] END >> "%LOG%"
exit /b 0