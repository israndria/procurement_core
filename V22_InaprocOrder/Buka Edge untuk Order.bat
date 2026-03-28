@echo off
echo Membuka Edge dengan mode debug untuk Order Bot...
echo Setelah Edge terbuka, login ke katalog.inaproc.id seperti biasa.
echo Kemudian kembali ke Streamlit dan klik "Hubungkan ke Edge".
echo.
start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --no-first-run https://katalog.inaproc.id
