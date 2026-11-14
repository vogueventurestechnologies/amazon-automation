@echo off
REM ============================================================================
REM Amazon.in Order + Payment Reconciliation - 1-Click Runner
REM Usage:
REM   run_reconciliation.bat
REM   run_reconciliation.bat --basis total_net
REM   run_reconciliation.bat --orders file.txt --payments file.csv
REM ============================================================================

setlocal enabledelayedexpansion

set "PROJECT_ROOT=%~dp0"
cd /d "%PROJECT_ROOT%"

set "PYTHON_EXE=%PROJECT_ROOT%.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtualenv python not found at "%PYTHON_EXE%"
    echo Please create the virtual environment first:
    echo   python -m venv .venv
    echo   .\.venv\Scripts\pip install -r requirements.txt
    pause
    exit /b 1
)

if "%~1"=="" (
    echo Running in AUTO mode (scanning /data/inbox)...
    "%PYTHON_EXE%" main.py --auto
) else (
    "%PYTHON_EXE%" main.py %*
)

set "EXIT_CODE=%ERRORLEVEL%"
echo.
if %EXIT_CODE% EQU 0 (
    echo [SUCCESS] Reconciliation finished successfully!
) else (
    echo [FAILED] Reconciliation encountered an error (exit code %EXIT_CODE%). Check logs/reconciliation.log.
)
echo.
pause
exit /b %EXIT_CODE%

