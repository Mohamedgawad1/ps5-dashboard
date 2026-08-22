@echo off
echo ============================================
echo  CPP AGI EIT - Server + Tunnel
echo ============================================
echo.

echo [1/3] Starting server...
start /min python "%~dp0mobile_app\server.py"
timeout /t 4 /nobreak >nul

echo [2/3] Testing server...
curl -s -o nul -w "Server: HTTP %%{http_code}\n" http://localhost:8080/data
if errorlevel 1 (
    echo Server FAILED! Check python error.
    pause
    exit /b 1
)

echo [3/3] Starting tunnel via localhost.run...
echo.
echo ============================================
echo  SHARE THIS URL WITH YOUR PHONE:
echo ============================================
echo.

ssh -o StrictHostKeyChecking=no -R 80:localhost:8080 nokey@localhost.run

echo.
echo Tunnel closed. Server still running at http://localhost:8080
