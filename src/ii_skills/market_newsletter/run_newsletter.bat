@echo off
REM Daily Market Newsletter Runner
REM This script is designed to be run by Windows Task Scheduler
REM Auto-restores config from OneDrive backup before each run

set VAULT=%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\configs
set SKILL_DIR=%~dp0

if exist "%VAULT%\market_config.py" (
    copy /Y "%VAULT%\market_config.py" "%SKILL_DIR%config.py" >nul 2>&1
)

cd /d "%SKILL_DIR%"
python newsletter.py >> newsletter.log 2>&1
