@echo off
setlocal
set "PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
set "SCRIPT=%~dp0WOW-Inventory-Scanner-v2.3.20.ps1"

rem Run in STA (needed for WinForms), no profile, bypass policy.
rem If it fails, write errors to last_error.txt and pause so you can see them.
"%PS%" -NoProfile -ExecutionPolicy Bypass -STA -File "%SCRIPT%" 2> "%~dp0last_error.txt"
if errorlevel 1 (
  echo.
  echo PowerShell returned error level %errorlevel%.
  echo A copy of the error (if any) is in: "%~dp0last_error.txt"
  echo.
  pause
)
