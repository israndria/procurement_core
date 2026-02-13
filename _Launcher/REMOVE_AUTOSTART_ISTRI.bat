@echo off
set "STARTUP_ISTRI=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Govem_Istri_Auto.bat"
if exist "%STARTUP_ISTRI%" (
    del "%STARTUP_ISTRI%"
    echo ✅ SUKSES! Auto-Start Istri Dimatikan.
) else (
    echo ⚠️ Auto-Start Istri tidak ditemukan.
)
pause
