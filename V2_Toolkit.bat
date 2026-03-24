@echo off
cd /d "%~dp0"
echo ====================================
echo   POKJA 2026 Toolkit
echo ====================================
echo.
python\python.exe -m streamlit run Home.py
pause
