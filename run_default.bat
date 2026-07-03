@echo off
setlocal

set "DEFAULT_EXCEL=D:\Downloads\2026-Import tracking list.xlsx"

call "%~dp0run_refresh.bat" "%DEFAULT_EXCEL%" %*
