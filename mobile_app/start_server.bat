@echo off
title CPP AGI - EIT Mobile App Server
echo.
echo  ============================================
echo   CPP AGI - EIT Cable Schedule Mobile App
echo  ============================================
echo.
echo  Starting server...
echo.
cd /d "%~dp0"
python server.py
pause
