@echo off
echo ============================================
echo  Daily Data Update - CPP AGI EIT
echo  %date% %time%
echo ============================================
echo.
cd /d "%~dp0"
python daily_update.py
echo.
echo Update complete. Press any key to exit.
pause > nul
