@echo off
setlocal
set "DOUBLE_CLICKED=0"
echo %cmdcmdline% | find /i "/c" >nul
if %ERRORLEVEL%==0 set "DOUBLE_CLICKED=1"

set "PROJECT_ROOT=%~dp0"
set "SRC_PATH=%PROJECT_ROOT%src"
set "VENV_PY=%PROJECT_ROOT%.venv\Scripts\python.exe"

if "%~1"=="" (
  echo Usage: run_refresh.bat "D:\Downloads\2026-Import tracking list.xlsx" [extra args]
  echo Example: run_refresh.bat "D:\Downloads\2026-Import tracking list.xlsx" --dry-run --limit 3
  call :finish 2
  exit /b 2
)

set "PYTHONPATH=%SRC_PATH%"

if exist "%VENV_PY%" (
  "%VENV_PY%" -m shipment_tracking.refresh_excel %*
  call :finish %ERRORLEVEL%
  exit /b %ERRORLEVEL%
)

where py >nul 2>nul
if %ERRORLEVEL%==0 (
  py -3 -m shipment_tracking.refresh_excel %*
  call :finish %ERRORLEVEL%
  exit /b %ERRORLEVEL%
)

where python >nul 2>nul
if %ERRORLEVEL%==0 (
  python -m shipment_tracking.refresh_excel %*
  call :finish %ERRORLEVEL%
  exit /b %ERRORLEVEL%
)

echo Python was not found. Install Python 3.10+ first.
echo Suggested: winget install Python.Python.3.12
call :finish 1
exit /b 1

:finish
set "EXIT_CODE=%~1"
if "%DOUBLE_CLICKED%"=="1" (
  echo.
  echo Exit code: %EXIT_CODE%
  pause
)
exit /b %EXIT_CODE%
