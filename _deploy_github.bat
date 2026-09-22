@echo off
rem Deploy subsystem_explorer.html -> GitHub Pages
setlocal enabledelayedexpansion
set REPO=PS5-COMPLETION-PLATFORM
set SRC=C:\Users\mylap\Downloads\PS5 - CPP AGI Completion Progress Dashboard_files\subsystem_explorer.html
set D=%TEMP%\gh_deploy_kentplc
set TOK=
for /f "usebackq delims=" %%t in ("%~dp0github_token.txt") do set TOK=%%t

if not exist "%D%\.git" (
    git clone --depth 1 "https://x-access-ps5:%TOK%@github.com/Mohamedgawad1/%REPO%.git" "%D%" >nul 2>&1
)
if not exist "%D%\.git" (
    echo [deploy] clone FAILED
    exit /b 1
)
copy /y "%SRC%" "%D%\index.html" >nul
copy /y "%SRC%" "%D%\app.html" >nul
pushd "%D%"
git add index.html app.html
git diff --cached --quiet
if not errorlevel 1 (
    echo [deploy] no changes
    popd
    exit /b 0
)
git -c user.name=ps5-bot -c user.email=ps5@local commit -m "PS5 platform update %date% %time%" >nul
git push >nul 2>&1
if errorlevel 1 (
    echo [deploy] push FAILED
) else (
    echo [deploy] pushed OK
)
popd
