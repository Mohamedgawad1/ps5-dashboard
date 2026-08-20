@echo off
chcp 65001 >nul
title PS5 Dashboard - Daily Update
color 0B
echo ============================================
echo  PS5 Dashboard - Daily Update
echo ============================================
echo.

set "HERE=%~dp0"
set "LATEST=C:\Users\mylap\Downloads\asset and punch"

echo [1/3] Syncing latest files...
for %%F in (
    "ovTasks_TestsPlanned_1369.xlsx"
    "ovPunchlist_1399.xlsx"
    "PS-5 EIT PUNCH LIST REGISTER (1).xlsx"
    "PS-5 EIT PUNCH LIST REGISTER.xlsx"
    "PS-5 INSPECTION REGISTER.xlsx"
    "PS5 Master tracker EIT Combined.xlsx"
    "PS5 EIT CPP AGI Dashboard.xlsx"
) do (
    if exist "%LATEST%\%%~nxF" (
        copy /Y "%LATEST%\%%~nxF" "%HERE%%%~nxF" >nul 2>&1
        echo   OK: %%~nxF
    )
)

echo.
echo [2/3] Building HTML dashboard...
python "%HERE%cpp_agi_dashboard.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Failed to build dashboard!
    pause
    exit /b 1
)

echo.
echo [3/3] Done!
echo.
echo Opening in browser...
start "" "https://mohamedgawad1.github.io/ps5-dashboard/"
echo.
pause
