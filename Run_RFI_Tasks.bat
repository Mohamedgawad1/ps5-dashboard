@echo off
REM ==============================================================
REM  SmartCloud - open the task pages for the PDFs in this folder
REM  The browser is started ONCE and stays open, so the session
REM  is never lost. Every run attaches to that same window.
REM ==============================================================
setlocal
cd /d "C:\Users\mylap\OneDrive\Desktop\dashboard"

echo ============================================================
echo   SmartCloud - open RFI task pages
echo   PDF folder: C:\Users\mylap\Downloads\rfi
echo ============================================================
echo.

powershell -NoProfile -Command "try { (Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:9222/json/version' -TimeoutSec 3).StatusCode } catch { exit 1 }" >nul 2>&1
if errorlevel 1 (
  echo [1/2] Chrome is not running - starting it now ...
  call "C:\Users\mylap\OneDrive\Desktop\dashboard\Start_SmartCloud.bat"
) else (
  echo [1/2] Chrome is already running.
)

echo.
echo [2/2] opening the task pages ...
echo        if a login page shows up, sign in once - it stays logged in.
echo.
python "C:\Users\mylap\OneDrive\Desktop\dashboard\open_tasks_batch.py"
echo.
echo ============================================================
echo  Done. The browser stays open.
echo  In every tab: Files tab - click the upload icon - pick the
echo  PDF - then click Save.
echo  This window keeps the session alive. Leave it open.
echo ============================================================
pause
endlocal
