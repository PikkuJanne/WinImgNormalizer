@echo off
setlocal EnableExtensions DisableDelayedExpansion
rem Let Windows PowerShell discover its own modules instead of inherited PS7 modules.
set "PSModulePath="
rem Preserve the native status even if an inherited variable shadows ERRORLEVEL.
set "ERRORLEVEL="
set "WINIMG_PAUSE="
set "WINIMG_PS=%SystemRoot%\System32\WindowsPowerShell\v1.0\powershell.exe"
rem The probe emits nothing. Its inherited handles must remain unredirected.
if "%WINIMG_NO_PAUSE%"=="1" goto arguments
"%WINIMG_PS%" -NoLogo -NoProfile -NonInteractive -Command "try { if ([Environment]::UserInteractive -and -not [Console]::IsInputRedirected -and -not [Console]::IsOutputRedirected -and -not [Console]::IsErrorRedirected) { exit 0 } } catch {}; exit 1"
if not errorlevel 1 set "WINIMG_PAUSE=1"

:arguments
if "%~1"=="" goto missing_argument
if not "%~2"=="" goto extra_arguments
if not "%2"=="" goto extra_arguments
set "SRC=%~1"
if "%SRC:~-1%"=="\" set "SRC=%SRC%."
if "%SRC:~-1%"=="/" set "SRC=%SRC%."
set "PS1=%~dpn0.ps1"
if not exist "%PS1%" goto missing_script
rem Keep Windows PowerShell 5.1 and process-only execution policy.
"%WINIMG_PS%" -NoLogo -NoProfile -ExecutionPolicy Bypass -File "%PS1%" "%SRC%"
set "WINIMG_EXIT=%ERRORLEVEL%"
goto finish

:missing_argument
echo ERROR: Drag and drop exactly one FOLDER onto this .bat file.
set "WINIMG_EXIT=1"
goto finish
:extra_arguments
echo ERROR: Drop exactly one FOLDER; extra arguments are not supported.
set "WINIMG_EXIT=1"
goto finish
:missing_script
echo ERROR: Missing PowerShell script: "%PS1%". Re-extract the official portable ZIP.
set "WINIMG_EXIT=1"
:finish
if not defined WINIMG_PAUSE goto return
echo.
echo Press any key to close.
pause >nul
:return
exit /b %WINIMG_EXIT%
