@echo off
rem Starts the ITR live update (one sync only, it keeps its own browser on port 9233).
title ITR Live Update
cd /d "C:\Users\mylap\OneDrive\Desktop\dashboard"
python -u -W ignore sc_pull\itr_online_sync.py --loop
pause
