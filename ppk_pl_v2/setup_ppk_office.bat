@echo off
setlocal
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup_ppk_office.ps1" %*
if not "%ERRORLEVEL%"=="0" pause
endlocal
