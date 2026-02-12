@echo off
REM Daily Net Worth Newsletter Runner
REM This script is designed to be run by Windows Task Scheduler
REM Auto-restores config from OneDrive backup before each run

set VAULT=%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\configs
set SKILL_DIR=%~dp0

if exist "%VAULT%\networth_config.py" (
    copy /Y "%VAULT%\networth_config.py" "%SKILL_DIR%config.py" >nul 2>&1
)

cd /d "%SKILL_DIR%"
python networth.py >> networth.log 2>&1
