@echo off
chcp 65001 >nul
title PS5 Dashboard - Daily Update
color 0B
echo ============================================
echo  PS5 CPP AGI Dashboard - Daily Update
echo ============================================
echo.

set "DASH=%~dp0"
set "DOWNLOADS=C:\Users\mylap\Downloads\asset and punch"

echo [1/4] Syncing latest files...
:: Copy newest Excel files from Downloads to dashboard folder
for %%F in ("ovTasks_TestsPlanned_1369.xlsx" "PS-5 EIT PUNCH LIST REGISTER (1).xlsx" "PS-5 EIT PUNCH LIST REGISTER.xlsx" "PS-5 INSPECTION REGISTER.xlsx" "PS5_EIT_Dashboard_CPP_AGI.xlsx") do (
    if exist "%DOWNLOADS%\%%F" (
        copy /Y "%DOWNLOADS%\%%F" "%DASH%%%F" >nul 2>&1
        echo   Synced: %%F
    )
)

echo.
echo [2/4] Running Excel refresh (refresh.py)...
python "%DASH%asset and punch\refresh.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [WARNING] Excel refresh had issues, continuing...
)

echo.
echo [3/4] Building HTML dashboard...
python "%DASH%cpp_agi_dashboard.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Failed to build dashboard!
    pause
    exit /b 1
)

echo.
echo [4/4] Uploading to GitHub...
cd /d "%DASH%"
git add -A
git commit -m "auto update %DATE% %TIME%" 2>nul
git push

echo.
echo ============================================
echo  Dashboard updated & uploaded to GitHub!
echo ============================================
echo.
echo Opening in browser...
start "" "https://mohamedgawad1.github.io/ps5-dashboard/"
echo.
pause
