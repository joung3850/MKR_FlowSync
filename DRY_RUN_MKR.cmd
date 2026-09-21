@echo off
setlocal
chcp 65001 >nul
title MKR FlowSync V12.0 - Dry Run
cd /d "%~dp0"

if not exist ".venv\Scripts\python.exe" (
  call "SETUP_MKR.cmd"
  if errorlevel 1 exit /b 1
)

echo [1/3] Checking Python syntax...
".venv\Scripts\python.exe" -m compileall -q "mkr_sync" "tests"
if errorlevel 1 goto :failed

echo [2/3] Running self-tests...
".venv\Scripts\python.exe" -m mkr_sync self-test
if errorlevel 1 goto :failed

echo [3/3] Creating a preview without changing Excel...
".venv\Scripts\python.exe" -m mkr_sync dry-run
if errorlevel 1 goto :failed

echo [OK] Preview report was created in the Reports folder.
pause
exit /b 0

:failed
echo [ERROR] Dry run stopped. Excel was not changed.
echo Review the latest file in the Logs folder.
pause
exit /b 1
