@echo off
REM ============================================================================
REM Amazon.in Reconciliation Web App - 1-Click Browser Launcher
REM Starts local server and opens http://localhost:8000 in your browser
REM ============================================================================

setlocal enabledelayedexpansion

set "PROJECT_ROOT=%~dp0"
cd /d "%PROJECT_ROOT%"

set "PYTHON_EXE=%PROJECT_ROOT%.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtual environment not found at "%PYTHON_EXE%"
    pause
    exit /b 1
)

echo ============================================================================
echo   Starting Amazon.in Reconciliation Web App
echo   Local Address: http://localhost:8000
echo   Vercel Ready:  Deploy anytime with 'vercel' command
echo ============================================================================
echo.
echo Opening browser...
start http://localhost:8000
echo.
echo Server running. Press Ctrl+C in this terminal window to stop.
echo.

"%PYTHON_EXE%" -m uvicorn api.index:app --host 127.0.0.1 --port 8000 --reload

pause

