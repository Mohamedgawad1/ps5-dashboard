@echo off
rem ============================================================
rem Deploy PS5 ITR-live files -> GitHub Pages (direct upload)
rem
rem index.html / app.html are NOT written here on purpose:
rem they are produced by rebuild_data.py (the "update from Excel"
rem pipeline). Copying subsystem_explorer.html over index.html
rem silently reverted every page-7 rebuild, so this script only
rem syncs the ITR live files and mirrors the published index.html.
rem ============================================================
setlocal enabledelayedexpansion
set REPO=PS5-COMPLETION-PLATFORM
set D=%TEMP%\gh_deploy_kentplc
set TOK=
for /f "usebackq delims=" %%t in ("%~dp0github_token.txt") do set TOK=%%t

if "%TOK%"=="" (
    echo [deploy] ERROR : github_token.txt is empty / missing
    exit /b 1
)

if not exist "%D%\.git" (
    git clone --depth 1 "https://github.com/Mohamedgawad1/%REPO%.git" "%D%" >nul 2>&1
)
if not exist "%D%\.git" (
    echo [deploy] clone FAILED
    exit /b 1
)

pushd "%D%"

rem --- keep token OUT of the remote config (security) ---
git remote set-url origin "https://github.com/Mohamedgawad1/%REPO%.git" >nul 2>&1

rem --- refresh local copy to current origin main ---
git fetch --depth 1 origin main >nul 2>&1
if not errorlevel 1 (
    git reset --hard origin/main >nul 2>&1
)
git checkout main >nul 2>&1

rem --- refuse to publish a broken index.html ---
findstr /C:"<<<<<<< HEAD" index.html >nul
if not errorlevel 1 (
    echo [deploy] ERROR ^: index.html has git conflict markers - aborting
    popd
    exit /b 1
)
findstr /C:"const RFCK=" index.html >nul
if errorlevel 1 (
    echo [deploy] ERROR ^: index.html missing RFCK ^(page 7 data^) - aborting
    popd
    exit /b 1
)
findstr /C:"const WIRE=" index.html >nul
if errorlevel 1 (
    echo [deploy] ERROR ^: index.html missing WIRE data - aborting
    popd
    exit /b 1
)

rem --- keep app.html identical to the published index.html ---
copy /y "index.html" "app.html" >nul

rem --- ITR LIVE page (standalone + badge + live state) ---
if exist "%~dp0itr_live.html" (
    copy /y "%~dp0itr_live.html" "%D%\itr_live.html" >nul
    copy /y "%~dp0live_itr.js" "%D%\live_itr.js" >nul
    copy /y "%~dp0itr_live_state.json" "%D%\itr_live_state.json" >nul
)
git add index.html app.html itr_live.html live_itr.js itr_live_state.json
git diff --cached --quiet
if not errorlevel 1 (
    echo [deploy] no changes - already up to date
    popd
    exit /b 0
)
git -c user.name=ps5-bot -c user.email=ps5@local commit -m "PS5 ITR live update %date% %time%" >nul
git push "https://x-access-ps5:%TOK%@github.com/Mohamedgawad1/%REPO%.git" main >nul 2>&1
if errorlevel 1 (
    echo [deploy] push FAILED
    popd
    exit /b 1
) else (
    echo [deploy] pushed OK
)

rem --- refresh the local offline mirror so it cannot rot ---
if exist "C:\Users\mylap\OneDrive\Desktop\dashboard\PS5 - CPP AGI Completion Progress Dashboard_files\" (
    copy /y "index.html" "C:\Users\mylap\OneDrive\Desktop\dashboard\PS5 - CPP AGI Completion Progress Dashboard_files\subsystem_explorer.html" >nul
)

popd
