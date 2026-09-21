@echo off
setlocal
chcp 65001 >nul
title MKR FlowSync V12.0 - Setup
cd /d "%~dp0"

py -3.12 -c "import sys; raise SystemExit(0 if sys.maxsize > 2**32 else 1)" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python 3.12 64-bit was not found.
  echo Install Python 3.12 64-bit and select "Add Python to PATH".
  pause
  exit /b 2
)

if not exist ".venv\Scripts\python.exe" (
  echo [1/3] Creating the Python environment...
  py -3.12 -m venv ".venv"
  if errorlevel 1 goto :failed
)

echo [2/3] Updating pip...
".venv\Scripts\python.exe" -m pip install --upgrade pip
if errorlevel 1 goto :failed

echo [3/3] Installing MKR dependencies...
".venv\Scripts\python.exe" -m pip install -r "requirements.txt"
if errorlevel 1 goto :failed

echo [OK] MKR Python environment is ready.
exit /b 0

:failed
echo [ERROR] Setup failed. Check the internet connection and company security policy.
pause
exit /b 3
