@echo off
py -3.14 "%~dp0show_url.py"
echo.
echo Current URL:
type "%USERPROFILE%\Desktop\EIT_URL.txt"
echo.
pause
