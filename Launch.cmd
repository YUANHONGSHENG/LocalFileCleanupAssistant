@echo off
setlocal
set "APP=%~dp0app\App.ps1"
if not exist "%APP%" (
  echo Missing app\App.ps1
  pause
  exit /b 1
)
powershell.exe -STA -NoProfile -ExecutionPolicy Bypass -File "%APP%"
if errorlevel 1 pause
endlocal
