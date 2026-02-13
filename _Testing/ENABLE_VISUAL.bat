@echo off
cd /d "D:\LDPlayer\LDPlayer9"

echo [1/4] Connecting to Emulator 0...
adb.exe connect 127.0.0.1:5554

echo [2/4] ENABLING VISUAL TOUCHES (Munculkan Putih-Putih saat klik)...
adb.exe -s 127.0.0.1:5554 shell settings put system show_touches 1

echo [3/4] TEST CLICKS (Perhatikan Layar!)...
echo    - Klik Kiri Atas...
adb.exe -s 127.0.0.1:5554 shell input swipe 100 100 100 100 200
timeout /t 1
echo    - Klik Tengah...
adb.exe -s 127.0.0.1:5554 shell input swipe 800 450 800 450 200
timeout /t 1
echo    - Klik Kanan Bawah...
adb.exe -s 127.0.0.1:5554 shell input swipe 1500 800 1500 800 200

echo [4/4] Selesai.
echo APAKAH ANDA MELIHAT TITIK PUTIH MUNCUL DI LAYAR SAAT TEST DIATAS?
pause
