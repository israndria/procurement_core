@echo off
cd /d "D:\LDPlayer\LDPlayer9"

echo [1/3] Copying Record File...
copy "vms\operationRecords\Step 1 (Menuju Dashboard Isian).record" "vms\operationRecords\test.record" /Y

echo [2/3] Executing LDConsole Action...
echo Command: ldconsole.exe action --index 0 --key call.keyboard --value test
ldconsole.exe action --index 0 --key call.keyboard --value test

echo [3/3] Done.
echo Jika layar LDPlayer tidak bergerak, berarti perintah 'action' gagal.
pause
