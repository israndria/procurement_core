' sync_jadwal_silent.vbs — jalankan sync_jadwal.py tanpa jendela hitam
Dim objShell
Set objShell = CreateObject("WScript.Shell")
objShell.Run "cmd /c """ & Chr(34) & "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\python\python.exe" & Chr(34) & " """ & Chr(34) & "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\sync_jadwal.py" & Chr(34) & """", 0, False
Set objShell = Nothing
