@echo off
set "STARTUP_SUAMI=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Govem_Suami_Auto.bat"
if exist "%STARTUP_SUAMI%" (
    del "%STARTUP_SUAMI%"
    echo ✅ SUKSES! Auto-Start Suami Dimatikan.
) else (
    echo ⚠️ Auto-Start Suami tidak ditemukan.
)
pause
