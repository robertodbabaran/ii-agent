@echo off
REM ═══════════════════════════════════════════════════════════
REM  Portfolio Tracker Runner
REM  Restores config from vault, generates workbook
REM ═══════════════════════════════════════════════════════════

set VAULT=%USERPROFILE%\OneDrive\Desktop\RJ\II-Agent Credentials\configs
set SKILL_DIR=%~dp0
set NW_DIR=%SKILL_DIR%..\networth_newsletter

echo [Portfolio Tracker] Restoring config from vault...
if exist "%VAULT%\networth_config.py" (
    copy /Y "%VAULT%\networth_config.py" "%NW_DIR%\config.py" >nul
    echo [Portfolio Tracker] Config restored.
) else (
    echo [Portfolio Tracker] WARNING: Vault config not found at %VAULT%
)

echo [Portfolio Tracker] Generating workbook...
cd /d "%SKILL_DIR%..\..\..\"
python -c "from ii_skills.portfolio_tracker import get_tracker; t = get_tracker(); r = t.execute('generate_workbook'); print(f'Output: {r[\"output_path\"]}')"

echo [Portfolio Tracker] Done.
pause
