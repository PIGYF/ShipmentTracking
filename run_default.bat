@echo off
setlocal

if not "%~1"=="" (
  call "%~dp0run_refresh.bat" %*
  exit /b %ERRORLEVEL%
)

set "DEFAULT_EXCEL=C:\Users\320305826\Philips\IWD ISC CT AMI DT Center - api_test_tmp\api_test.xlsx"
if defined SHIPMENT_TRACKING_DEFAULT_EXCEL set "DEFAULT_EXCEL=%SHIPMENT_TRACKING_DEFAULT_EXCEL%"

if not exist "%DEFAULT_EXCEL%" (
  echo Default Excel file was not found:
  echo   %DEFAULT_EXCEL%
  echo.
  echo Run with an explicit workbook path instead:
  echo   run_default.bat "C:\path\to\workbook.xlsx"
  echo.
  echo Or set SHIPMENT_TRACKING_DEFAULT_EXCEL to this computer's workbook path.
  exit /b 1
)

call "%~dp0run_refresh.bat" "%DEFAULT_EXCEL%" %*
