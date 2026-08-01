@echo off
setlocal EnableExtensions

set "RUNTIME_DIR=%~dp0"
for %%I in ("%RUNTIME_DIR%..") do set "PACKAGE_ROOT=%%~fI"
set "PPK_PACKAGE_ROOT=%PACKAGE_ROOT%"
set "PPK_LOG_DIR=%RUNTIME_DIR%"
set "SOURCE_DIR="

:: Source generator wajib berasal dari clone lokal per-PC.
if defined POKJA_PPK_ROOT if exist "%POKJA_PPK_ROOT%\generate_dokumen_ppk.py" set "SOURCE_DIR=%POKJA_PPK_ROOT%"
if not defined SOURCE_DIR if defined POKJA_V19_ROOT if exist "%POKJA_V19_ROOT%\ppk_pl_v2\generate_dokumen_ppk.py" set "SOURCE_DIR=%POKJA_V19_ROOT%\ppk_pl_v2"
if not defined SOURCE_DIR if defined POKJA_CODE_ROOT if exist "%POKJA_CODE_ROOT%\..\procurement_core\ppk_pl_v2\generate_dokumen_ppk.py" set "SOURCE_DIR=%POKJA_CODE_ROOT%\..\procurement_core\ppk_pl_v2"
:: Fallback hanya untuk paket lama yang belum memakai source Git; generatornya
:: tetap portable dan memakai PPK_PACKAGE_ROOT di atas.
if not defined SOURCE_DIR if exist "%RUNTIME_DIR%generate_dokumen_ppk.py" set "SOURCE_DIR=%RUNTIME_DIR%"

if not defined SOURCE_DIR (
    echo ERROR: Source generator PPK V2 tidak ditemukan. > "%RUNTIME_DIR%generate_log.txt"
    echo Set POKJA_V19_ROOT ke clone procurement_core lokal. >> "%RUNTIME_DIR%generate_log.txt"
    exit /b 1
)
set "SCRIPT_PATH=%SOURCE_DIR%\generate_dokumen_ppk.py"

:: Python portable per-PC. Jangan memakai Python dari Google Drive.
set "PYTHON="
if defined POKJA_PYTHON if exist "%POKJA_PYTHON%" set "PYTHON=%POKJA_PYTHON%"
if not defined PYTHON if defined POKJA_V19_ROOT if exist "%POKJA_V19_ROOT%\python\python.exe" set "PYTHON=%POKJA_V19_ROOT%\python\python.exe"
if not defined PYTHON if defined POKJA_CODE_ROOT if exist "%POKJA_CODE_ROOT%\..\Runtime\WPy64-313110\python\python.exe" set "PYTHON=%POKJA_CODE_ROOT%\..\Runtime\WPy64-313110\python\python.exe"
if not defined PYTHON if exist "C:\WinPython313\python\python.exe" set "PYTHON=C:\WinPython313\python\python.exe"
if not defined PYTHON for /f "delims=" %%P in ('where python 2^>nul') do if not defined PYTHON set "PYTHON=%%P"

if not defined PYTHON (
    echo ERROR: Python portable tidak ditemukan. > "%RUNTIME_DIR%generate_log.txt"
    echo Set POKJA_PYTHON ke python.exe lokal. >> "%RUNTIME_DIR%generate_log.txt"
    exit /b 1
)
if not exist "%PACKAGE_ROOT%\0. Master_Data_PL_PPK.xlsm" (
    echo ERROR: Master_Data_PL_PPK.xlsm tidak ditemukan di %PACKAGE_ROOT%. > "%RUNTIME_DIR%generate_log.txt"
    exit /b 1
)

pushd "%PACKAGE_ROOT%"
"%PYTHON%" "%SCRIPT_PATH%" %* > "%RUNTIME_DIR%generate_log.txt" 2>&1
set "EXIT_CODE=%ERRORLEVEL%"
popd
exit /b %EXIT_CODE%
