@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem An inherited variable can shadow CMD's dynamic native exit-code value.
set "ERRORLEVEL="

if "%~1"=="" (
  echo ERROR: Drag and drop exactly one FOLDER onto this .bat file.
  pause
  exit /b 1
)

rem Check stripped content first, then distinguish explicit "" from no argument.
rem Only the empty case reaches the raw check, avoiding nested path quotes.
if not "%~2"=="" goto extra_arguments
if not "%2"=="" goto extra_arguments

set "SRC=%~1"
rem Preserve a trailing root/folder separator across native quoted argv parsing.
if "%SRC:~-1%"=="\" set "SRC=%SRC%."
if "%SRC:~-1%"=="/" set "SRC=%SRC%."
set "PS1=%~dpn0.ps1"

if not exist "%PS1%" (
  echo ERROR: Missing PowerShell script: "%PS1%"
  pause
  exit /b 1
)

rem Retain the Windows PowerShell host and process-only execution policy.
powershell -NoProfile -ExecutionPolicy Bypass -File "%PS1%" "%SRC%"
set "WINIMG_EXIT=%ERRORLEVEL%"
echo.
if "%WINIMG_EXIT%"=="0" goto completed
if "%WINIMG_EXIT%"=="2" goto partial
if "%WINIMG_EXIT%"=="130" goto cancelled
echo ERROR: Processing did not complete reliably (exit code %WINIMG_EXIT%). Press any key to close.
goto finish

:completed
echo Completed successfully. Press any key to close.
goto finish

:partial
echo Completed with warnings or partial failures. Review the report. Press any key to close.
goto finish

:cancelled
echo Cancelled. Completed outputs were retained. Press any key to close.

:finish
pause >nul
exit /b %WINIMG_EXIT%

:extra_arguments
echo ERROR: Drop exactly one FOLDER; extra arguments are not supported.
pause
exit /b 1
