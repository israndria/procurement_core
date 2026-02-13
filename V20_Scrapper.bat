@echo off
cd /d "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"

echo ---------------------------------------------------
echo 🕵️ Membuka V20 Scraper (Supabase)...
echo ---------------------------------------------------

:: Cek posisi Python otomatis
if exist "python\python.exe" (
    "python\python.exe" -m streamlit run V20_Scrapper.py
    goto end
)

if exist "WPy64-313110\python\python.exe" (
    "WPy64-313110\python\python.exe" -m streamlit run V20_Scrapper.py
    goto end
)

echo [ERROR] Python tidak ditemukan. Cek folder WinPython kamu.
pause

:end