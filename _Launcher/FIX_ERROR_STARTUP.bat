@echo off
echo SEDANG MEMBERSIHKAN SHORTCUT ERROR DARI STARTUP...
echo.

set "S_DIR=%APPDATA%\Microsoft\Windows\Start Menu\Programs\Startup"

:: Hapus Versi Lama (V22 Awal)
if exist "%S_DIR%\Govem_Auto.bat" del "%S_DIR%\Govem_Auto.bat"

:: Hapus Versi Intermediate (Yang bikin error 'Run_Suami not found')
if exist "%S_DIR%\Govem_Suami.bat" del "%S_DIR%\Govem_Suami.bat"
if exist "%S_DIR%\Govem_Istri.bat" del "%S_DIR%\Govem_Istri.bat"

:: Hapus Versi Baru (Biar bersih sekalian, nanti install ulang)
if exist "%S_DIR%\Govem_Suami_Auto.bat" del "%S_DIR%\Govem_Suami_Auto.bat"
if exist "%S_DIR%\Govem_Istri_Auto.bat" del "%S_DIR%\Govem_Istri_Auto.bat"

echo.
echo ✅ PEMBERSIHAN SELESAI.
echo Error "Run_Suami.bat not found" tidak akan muncul lagi.
echo.
echo Silakan jalankan 'SETUP_AUTOSTART_SUAMI.bat' lagi jika ingin mengaktifkan auto-start yang BENAR.
pause
