@echo off
REM =============================================
REM  Setup Autostart Govem Scheduler (Tray Icon)
REM =============================================

set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

REM Buat VBS launcher untuk pythonw (tidak muncul console)
echo Set WshShell = CreateObject("WScript.Shell") > "%STARTUP_FOLDER%\Govem_Scheduler.vbs"
echo WshShell.CurrentDirectory = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\_Govem" >> "%STARTUP_FOLDER%\Govem_Scheduler.vbs"
echo WshShell.Run "pythonw Govem_Tray.pyw", 0, False >> "%STARTUP_FOLDER%\Govem_Scheduler.vbs"

echo.
echo =============================================
echo   AUTOSTART BERHASIL DISETUP!
echo =============================================
echo.
echo File dibuat di:
echo %STARTUP_FOLDER%\Govem_Scheduler.vbs
echo.
echo Scheduler akan otomatis berjalan saat Windows startup
echo dengan icon di system tray (pojok kanan bawah).
echo.
pause
