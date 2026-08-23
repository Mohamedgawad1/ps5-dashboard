@echo off
title PS5 DAILY UPDATE + UPLOAD
cd /d "C:\Users\mylap\OneDrive\Desktop\dashboard"
color 0B
echo ==========================================
echo    PS5   DAILY   UPDATE   +   UPLOAD
echo ==========================================
echo.
echo  TIP: Close "PS5 DPR DASHBOARD.xlsm" in Excel first for a smooth update.
echo.
echo [%time:~0,8%] [1/3]  Building EXCEL dashboard + PLATFORM ...
echo        (takes about one or two minutes - please wait)
echo.
python punch_itr_explorer.py
if errorlevel 1 goto fail
python dpr_dashboard.py
if errorlevel 1 goto fail
if not exist "C:\Users\mylap\Downloads\PS5 - CPP AGI Completion Progress Dashboard_files\PS5 DPR DASHBOARD.xlsm" goto fail
echo.
echo [%time:~0,8%] [2/3]  Uploading platform to GitHub Pages ...
call "_deploy_github.bat"
if errorlevel 1 goto fail
echo.
echo [%time:~0,8%] [3/3]  DONE - everything is up to date:
echo.
echo        SITE    : https://mohamedgawad1.github.io/PS5-COMPLETION-PLATFORM/
echo        EXCEL   : PS5 DPR DASHBOARD.xlsm  (with new MILESTONE ^& DIS. sheet)
echo        PLATFORM: subsystem_explorer.html
echo.
start "" https://mohamedgawad1.github.io/PS5-COMPLETION-PLATFORM/
pause
exit /b 0

:fail
echo.
color 0C
echo        ERROR - close Excel and any open files, then run again.
echo.
pause
exit /b 1
