@echo off
setlocal
set "REPO_DIR=%~dp0"
if not defined POKJA_PYTHON set "POKJA_PYTHON=%REPO_DIR%..\Runtime\WPy64-313110\python\python.exe"
set "SCRAPE_TAHUN=2026"
set "SCRAPE_KATEGORI=Tender"
set "SCRAPE_KODE_LPSE="
cd /d "%REPO_DIR%"
"%POKJA_PYTHON%" "%REPO_DIR%scrape_spse.py"
exit /b %ERRORLEVEL%
