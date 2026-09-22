@echo off
chcp 65001 >nul
title PS5 - Platform to Excel Sync
echo ============================================
echo   Sync platform edits into DPR SUMMERY Excel
echo ============================================
python "%~dp0sync_cloud_to_excel.py" %*
if %ERRORLEVEL% NEQ 0 (
  echo.
  echo [ERROR] sync failed - see messages above
  pause
  exit /b 1
)
echo.
echo Done. Report: _platform_sync_report.txt
if /I NOT "%1"=="--watch" pause
