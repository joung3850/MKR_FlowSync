@echo off
setlocal
chcp 65001 >nul
title MKR SYNC - Build EXE
cd /d "%~dp0"

set "PYTHON_CMD=py -3.14"
%PYTHON_CMD% --version >nul 2>&1
if errorlevel 1 set "PYTHON_CMD=python"

%PYTHON_CMD% -c "import sys; raise SystemExit(0 if sys.version_info[:2] == (3, 14) and sys.maxsize > 2**32 else 1)" >nul 2>&1
if errorlevel 1 (
  echo [ERROR] Python 3.14 64-bit was not found.
  pause
  exit /b 2
)

%PYTHON_CMD% -m pip install --upgrade -r requirements.txt
if errorlevel 1 goto :failed

%PYTHON_CMD% -m PyInstaller --clean --noconfirm --distpath "%~dp0" "MKR_SYNC.spec"
if errorlevel 1 goto :failed

echo.
echo [OK] MKR_SYNC.exe was created in the project folder.
echo Keep MKR_TEMPLATE.xlsx beside the EXE, or select a template on first use.
pause
exit /b 0

:failed
echo [ERROR] EXE build failed.
pause
exit /b 1
