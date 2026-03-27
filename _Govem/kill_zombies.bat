@echo off
echo ============================================
echo   KILL ZOMBIE PROCESSES (Sebelum Test Govem)
echo ============================================
echo.

echo [1/3] Killing semua python Govem...
taskkill /F /FI "WINDOWTITLE eq *Govem*" >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq *V23*" >nul 2>&1
taskkill /F /FI "WINDOWTITLE eq *aktivitas*" >nul 2>&1

echo [2/3] Killing LDPlayer emulators...
taskkill /F /IM dnplayer.exe >nul 2>&1
taskkill /F /IM Ld9BoxHeadless.exe >nul 2>&1
taskkill /F /IM LdVBoxHeadless.exe >nul 2>&1

echo [3/3] Reset ADB...
D:\LDPlayer\LDPlayer9\adb.exe kill-server >nul 2>&1
D:\LDPlayer\LDPlayer9\adb.exe start-server >nul 2>&1

echo.
echo ============================================
echo   BERSIH! Siap test Govem.
echo ============================================
pause
