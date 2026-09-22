@echo off
rem Silent scheduled pipeline: rebuild explorer + dashboard, then deploy site
cd /d "C:\Users\mylap\OneDrive\Desktop\dashboard"
set LOG=C:\Users\mylap\Downloads\PS5 - CPP AGI Completion Progress Dashboard_files\_auto_update_log.txt
echo ====== %date% %time% ====== >> "%LOG%"

python punch_itr_explorer.py >> "%LOG%" 2>&1
if errorlevel 1 (echo [explorer] FAILED >> "%LOG%") else (echo [explorer] OK >> "%LOG%")

python "C:\Users\mylap\OneDrive\Desktop\dashboard\apply_platform_edits_to_summery.py" >> "%LOG%" 2>&1
if errorlevel 1 (echo [summery-sync] FAILED >> "%LOG%") else (echo [summery-sync] OK >> "%LOG%")

python dpr_dashboard.py >> "%LOG%" 2>&1
if errorlevel 1 (echo [dashboard] FAILED >> "%LOG%") else (echo [dashboard] OK >> "%LOG%")

call "C:\Users\mylap\OneDrive\Desktop\dashboard\_deploy_github.bat" >> "%LOG%" 2>&1

echo ------ done %time% ------ >> "%LOG%"
