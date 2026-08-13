@echo off
setlocal
set "REPO_DIR=%~dp0"
if not defined POKJA_PYTHON set "POKJA_PYTHON=%REPO_DIR%..\Runtime\WPy64-313110\python\python.exe"
set "SCRAPE_TAHUN=2026"
set "SCRAPE_KODE_LPSE="
cd /d "%REPO_DIR%"

set "SCRAPE_KATEGORI=Tender"
"%POKJA_PYTHON%" "%REPO_DIR%scrape_spse.py"
set "RC_TENDER=%ERRORLEVEL%"

set "SCRAPE_KATEGORI=Non Tender"
"%POKJA_PYTHON%" "%REPO_DIR%scrape_spse.py"
set "RC_NON_TENDER=%ERRORLEVEL%"

if not "%RC_TENDER%"=="0" exit /b %RC_TENDER%
exit /b %RC_NON_TENDER%
