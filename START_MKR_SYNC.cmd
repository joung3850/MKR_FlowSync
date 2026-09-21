@echo off
setlocal
chcp 65001 >nul
title MKR FlowSync V12.0 Python Edition
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  call "SETUP_MKR.cmd"
  if errorlevel 1 exit /b 1
)

if not exist "MKR_TEST.xlsx" (
  echo [ERROR] MKR_TEST.xlsx was not found.
  pause
  exit /b 2
)

echo [1/3] Checking Python syntax...
".venv\Scripts\python.exe" -m compileall -q "mkr_sync" "tests"
if errorlevel 1 goto :failed

echo [2/3] Running quantity and parser self-tests...
".venv\Scripts\python.exe" -m mkr_sync self-test
if errorlevel 1 goto :failed

echo [3/3] Reading Classic Outlook and updating Excel safely...
".venv\Scripts\python.exe" -m mkr_sync sync
if errorlevel 1 goto :failed

start "" "%~dp0MKR_TEST.xlsx"
echo [OK] Synchronization completed.
pause
exit /b 0

:failed
echo [ERROR] Synchronization stopped. The original Excel file was not replaced.
echo Review the latest file in the Logs folder.
pause
exit /b 1
