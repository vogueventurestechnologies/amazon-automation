@echo off
setlocal enabledelayedexpansion
set "PROJECT_ROOT=%~dp0"
cd /d "%PROJECT_ROOT%"
set "PYTHON_EXE=%PROJECT_ROOT%.venv\Scripts\python.exe"

if not exist "%PYTHON_EXE%" (
    echo [ERROR] Virtualenv python not found at "%PYTHON_EXE%"
    pause
    exit /b 1
)

"%PYTHON_EXE%" step2_payment_pivot.py %*
pause

