@echo off
set "TARGET_DIR=%~dp0"
if "%TARGET_DIR:~-1%"=="\" set "TARGET_DIR=%TARGET_DIR:~0,-1%"

set "STARTUP_ISTRI=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Govem_Istri_Auto.bat"

echo Menginstall Auto-Start untuk ISTRI...
(
echo @echo off
echo cd /d "%TARGET_DIR%"
echo start "" /MIN pythonw Govem_Istri.py --auto
) > "%STARTUP_ISTRI%"

echo.
echo ✅ SUKSES! Auto-Start Istri Terpasang.
echo File: %STARTUP_ISTRI%
pause
