@echo off
setlocal
chcp 65001 >nul
title MKR FlowSync V12.0 - Python 3.14 Setup
cd /d "%~dp0"

set "PYTHON_CMD=py -3.14"
set "ENV_DIR=.venv314"

%PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 14) and sys.maxsize > 2**32 else 1)" >nul 2>&1
if errorlevel 1 (
  set "PYTHON_CMD=python"
  python -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 14) and sys.maxsize > 2**32 else 1)" >nul 2>&1
  if errorlevel 1 (
    echo [ERROR] Python 3.14 64-bit was not found.
    echo Install Python 3.14 64-bit and select "Add Python to PATH".
    pause
    exit /b 2
  )
)

if not exist "%ENV_DIR%\Scripts\python.exe" (
  echo [1/3] Creating the Python environment...
  %PYTHON_CMD% -m venv "%ENV_DIR%"
  if errorlevel 1 goto :failed
)

echo [2/3] Updating pip...
"%ENV_DIR%\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :failed

echo [3/3] Installing MKR dependencies...
"%ENV_DIR%\Scripts\python.exe" -m pip install -r "requirements.txt"
if errorlevel 1 goto :failed

echo [OK] MKR Python 3.14 environment is ready.
exit /b 0

:failed
echo [ERROR] Setup failed. Check the internet connection and company security policy.
pause
exit /b 3
