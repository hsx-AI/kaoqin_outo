@echo off
cd /d "%~dp0"
echo Creating auto-start task...
echo NOTE: Please Run as Administrator.

set "TASK_NAME=KaoqinAutoReport"
set "SCRIPT_PATH=%~dp0run_service.bat"

schtasks /create /tn "%TASK_NAME%" /tr "\"%SCRIPT_PATH%\"" /sc onlogon /rl highest /f

if %errorlevel% equ 0 (
    echo.
    echo [SUCCESS] Task created!
    echo 1. The service will start automatically after you LOGON.
    echo 2. You can manually run run_service.bat to start it now.
    echo 3. Please ensure Power Settings are set to 'Never Sleep'.
) else (
    echo.
    echo [FAILURE] Failed to create task. Please run as Administrator.
)

pause
