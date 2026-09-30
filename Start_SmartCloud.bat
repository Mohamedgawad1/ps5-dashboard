@echo off
REM ==============================================================
REM  Starts Chrome ONCE with the SmartCloud profile and a debug
REM  port, then leaves it running. The session stays alive.
REM ==============================================================
setlocal
set PROFILE=%LOCALAPPDATA%\sc_itr_profile
set URL=https://wly04-sc.intergraphsmartcloud.com/ISC/Tools/vDashboardsUsers/Switchboard.htm

set CHROME=
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set CHROME=%ProgramFiles%\Google\Chrome\Application\chrome.exe
if not defined CHROME if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set CHROME=%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe
if not defined CHROME if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" set CHROME=%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe

if "%CHROME%"=="" (
  echo Chrome is not installed.
  pause
  exit /b 1
)

powershell -NoProfile -Command "Start-Process -FilePath '%CHROME%' -ArgumentList '--remote-debugging-port=9222','--user-data-dir=\"%PROFILE%\"','--start-maximized','%URL%'"

echo.
echo Waiting for the debug port ...
for /l %%i in (1,1,30) do (
  powershell -NoProfile -Command "try { (Invoke-WebRequest -UseBasicParsing -Uri 'http://127.0.0.1:9222/json/version' -TimeoutSec 2).StatusCode } catch { exit 1 }" >nul 2>&1
  if not errorlevel 1 (
    echo Chrome is ready on port 9222.
    exit /b 0
  )
  ping -n 2 127.0.0.1 >nul
)
echo Chrome did not answer on 9222 - sign in once, then try again.
pause
exit /b 1
