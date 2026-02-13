@echo off
cd /d "D:\LDPlayer\LDPlayer9"

echo [TEST 1] Clicking Top-Left (100, 100)...
adb.exe -s emulator-5554 shell input swipe 100 100 100 100 200
timeout /t 2

echo [TEST 2] Clicking CENTER (800, 450) - Assumes 1600x900...
adb.exe -s emulator-5554 shell input swipe 800 450 800 450 200
timeout /t 2

echo [TEST 3] Clicking RECORDED POINT (Step 1 approx: 672, 251)...
adb.exe -s emulator-5554 shell input swipe 672 251 672 251 200
timeout /t 2

echo Done.
pause
