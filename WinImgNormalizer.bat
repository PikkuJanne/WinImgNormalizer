@echo off
setlocal EnableExtensions

if "%~1"=="" (
  echo Drag and drop a FOLDER onto this .bat file.
  pause
  exit /b 1
)

set "SRC=%~1"
set "PS1=%~dpn0.ps1"

if not exist "%PS1%" (
  echo ERROR: Missing PowerShell script: "%PS1%"
  pause
  exit /b 1
)

rem # Positional only, avoids PS 5.1 parameter set ambiguity
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" "%SRC%"
echo.
echo Done. Press any key to close.
pause >nul

