@echo off
cd /d "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"

echo ---------------------------------------------------
echo Inject VBA ke semua file .xlsm di @ POKJA 2026
echo ---------------------------------------------------

if exist "python\python.exe" (
    "python\python.exe" inject_all.py
    goto end
)

if exist "WPy64-313110\python\python.exe" (
    "WPy64-313110\python\python.exe" inject_all.py
    goto end
)

echo [ERROR] Python tidak ditemukan. Cek folder WinPython kamu.

:end
pause
