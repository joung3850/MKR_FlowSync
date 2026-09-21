@echo off
setlocal
chcp 65001 >nul
title MKR Multi Outlook Sync V11.1 - Net Order Qty - Max 500 Mails
cd /d "%~dp0"
set "SYNC_SCRIPT=%~dp0MKR48_Outlook_AllItems_500.ps1"
set "CHECK_SCRIPT=%~dp0CHECK_MKR48_SCRIPT.ps1"

if not exist "MKR_TEST.xlsx" (
  echo [ERROR] MKR_TEST.xlsx not found.
  pause
  exit /b 1
)

if not exist "MKR48_Outlook_AllItems_500.ps1" (
  echo [ERROR] PowerShell script not found.
  pause
  exit /b 1
)

if not exist "CHECK_MKR48_SCRIPT.ps1" (
  echo [ERROR] PowerShell syntax checker not found.
  pause
  exit /b 1
)

echo [0/3] Checking PowerShell 5.1 script syntax...
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%CHECK_SCRIPT%" "%SYNC_SCRIPT%"
set "PREFLIGHT_EXIT=%ERRORLEVEL%"
if not "%PREFLIGHT_EXIT%"=="0" (
  echo The workbook and Outlook were not changed.
  pause
  exit /b %PREFLIGHT_EXIT%
)

echo [1/3] Testing net new order, no double deduction, separate backorder, and MKR56 logic...
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%SYNC_SCRIPT%" -SelfTest
set "SELFTEST_EXIT=%ERRORLEVEL%"
if not "%SELFTEST_EXIT%"=="0" (
  echo The workbook and Outlook were not changed.
  pause
  exit /b %SELFTEST_EXIT%
)

echo [2/3] Close MKR_TEST.xlsx and keep Classic Outlook open.
echo [3/3] Searching no more than 500 Outlook messages...
powershell.exe -NoLogo -NoProfile -STA -ExecutionPolicy Bypass -File "%SYNC_SCRIPT%"
set "SYNC_EXIT=%ERRORLEVEL%"

if not "%SYNC_EXIT%"=="0" (
  echo.
  echo [ERROR] Synchronization failed. The workbook was not replaced.
  echo Review MKR_SYNC_LOG.txt and the Backup folder.
  pause
  exit /b %SYNC_EXIT%
)

echo.
echo [SUCCESS] The workbook was updated and opened.
pause
exit /b 0
