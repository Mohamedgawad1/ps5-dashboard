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
py -3.14 server.py
pause
