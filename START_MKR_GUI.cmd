@echo off
setlocal
cd /d "%~dp0"

if not exist ".venv314\Scripts\python.exe" (
  call "SETUP_MKR.cmd"
  if errorlevel 1 exit /b 1
)

".venv314\Scripts\python.exe" -m mkr_sync gui
if errorlevel 1 pause
endlocal
