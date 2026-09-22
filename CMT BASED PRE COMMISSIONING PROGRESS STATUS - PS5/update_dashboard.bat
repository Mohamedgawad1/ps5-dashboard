@echo off
cd /d "%~dp0"
python generate_dashboard.py
if %errorlevel% equ 0 (
    echo Deploying to GitHub...
    powershell -ExecutionPolicy Bypass -File "%~dp0deploy.ps1"
)
pause
