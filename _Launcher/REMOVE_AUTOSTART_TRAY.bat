@echo off
REM =============================================
REM  Hapus Autostart Govem Scheduler
REM =============================================

set STARTUP_FOLDER=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup

if exist "%STARTUP_FOLDER%\Govem_Scheduler.vbs" (
    del "%STARTUP_FOLDER%\Govem_Scheduler.vbs"
    echo.
    echo =============================================
    echo   AUTOSTART BERHASIL DIHAPUS!
    echo =============================================
    echo.
    echo Scheduler tidak akan otomatis berjalan saat Windows startup.
    echo.
) else (
    echo.
    echo Tidak ada autostart yang perlu dihapus.
    echo.
)

pause
