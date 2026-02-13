@echo off
cd /d "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110"

echo ---------------------------------------------------
echo 🚀 Mempersiapkan V19 Scheduler...
echo ---------------------------------------------------

:: KEMUNGKINAN 1: File bat ada DI DALAM folder WinPython (sebelahan sama folder 'python')
if exist "python\python.exe" (
    echo [INFO] Menggunakan Python mode dekat.
    "python\python.exe" -m streamlit run V19_Scheduler.py
    goto end
)

:: KEMUNGKINAN 2: File bat ada DI LUAR (sebelahan sama folder WPy64...)
:: Kita cek nama foldernya secara manual satu per satu karena versi bisa beda
if exist "WPy64-313110\python\python.exe" (
    echo [INFO] Menggunakan Python di folder WPy64-313110.
    "WPy64-313110\python\python.exe" -m streamlit run V19_Scheduler.py
    goto end
)

:: KEMUNGKINAN 3: Mungkin nama foldernya ada embel-embel 'slim'
if exist "WPy64-313110slim\python\python.exe" (
    echo [INFO] Menggunakan Python di folder WPy64-313110slim.
    "WPy64-313110slim\python\python.exe" -m streamlit run V19_Scheduler.py
    goto end
)

:: JIKA SEMUA GAGAL
echo [ERROR] Gawat! File python.exe tidak ditemukan.
echo Pastikan struktur foldermu benar.
echo.
pause

:end
pause