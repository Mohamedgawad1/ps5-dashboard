@echo off
chcp 65001 >nul
title PS5 - UPDATE ALL
color 0B
echo ==========================================
echo   PS5  -  UPDATING EVERYTHING ...
echo ==========================================
echo.
cd /d "C:\Users\mylap\OneDrive\Desktop\PS5-COMPLETION-PLATFORM"
echo [1/4] Rebuild data from XLSX files ...
python "C:\Users\mylap\OneDrive\Desktop\PS5-COMPLETION-PLATFORM\rebuild_data.py"
if errorlevel 1 goto err
echo.
echo [2/4] Git commit ...
git add index.html
git commit -m "Update all from XLSX (auto)" --no-verify
echo.
echo [3/4] Git sync with remote (pull --rebase) ...
git pull --rebase --no-edit origin main
if errorlevel 1 goto err
echo.
echo [4/4] Git push to GitHub (live site) ...
git push origin main --no-verify
if errorlevel 1 goto err
echo.
echo ==========================================
echo   DONE!
echo   Site: https://mohamedgawad1.github.io/PS5-COMPLETION-PLATFORM/
echo ==========================================
pause
exit /b 0
:err
echo.
echo [ERROR] Update failed - check message above
pause
