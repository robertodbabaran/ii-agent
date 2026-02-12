@echo off
REM Daily WHOOP Health Dashboard Runner
REM This script is designed to be run by Windows Task Scheduler
REM Auto-restores config from OneDrive backup before each run

set VAULT=%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\configs
set SKILL_DIR=%~dp0

if exist "%VAULT%\whoop_config.py" (
    copy /Y "%VAULT%\whoop_config.py" "%SKILL_DIR%config.py" >nul 2>&1
)

cd /d "%SKILL_DIR%"
python whoop_newsletter.py >> whoop.log 2>&1
