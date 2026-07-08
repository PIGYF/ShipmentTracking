@echo off
setlocal

::set "DEFAULT_EXCEL=C:\Users\320305826\Philips\IWD ISC CT AMI DT Center - 4. Inbound\2026-Import tracking list.xlsx"

set "DEFAULT_EXCEL=C:\Users\320305826\Philips\IWD ISC CT AMI DT Center - api_test_tmp\api_test.xlsx"

call "%~dp0run_refresh.bat" "%DEFAULT_EXCEL%" %*
