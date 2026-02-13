@echo off
set "TARGET_DIR=%~dp0"
if "%TARGET_DIR:~-1%"=="\" set "TARGET_DIR=%TARGET_DIR:~0,-1%"

set "STARTUP_SUAMI=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup\Govem_Suami_Auto.bat"

echo Menginstall Auto-Start untuk SUAMI...
(
echo @echo off
echo cd /d "%TARGET_DIR%"
echo start "" /MIN pythonw Govem_Suami.py --auto
) > "%STARTUP_SUAMI%"

echo.
echo ✅ SUKSES! Auto-Start Suami Terpasang.
echo File: %STARTUP_SUAMI%
pause
