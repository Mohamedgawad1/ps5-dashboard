@echo off
chcp 65001 >nul
echo ============================================
echo  PS5 CPP AGI Dashboard Updater
echo ============================================
echo.

echo [1/1] Building & uploading dashboard...
python "%~dp0cpp_agi_dashboard.py"
if %ERRORLEVEL% NEQ 0 (
    echo.
    echo [ERROR] Failed to generate dashboard!
    pause
    exit /b 1
)

echo.
echo ============================================
echo  Dashboard updated & uploaded to GitHub!
echo ============================================
echo.
echo Opening in browser...
start "" "https://mohamedgawad1.github.io/ps5-dashboard/"
echo.
pause
