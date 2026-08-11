' sync_jadwal_silent.vbs — sync Tender + PL tanpa jendela hitam
Dim objShell, objFSO, repoDir, pythonExe, scriptPath, command, exitCode
Set objShell = CreateObject("WScript.Shell")
Set objFSO = CreateObject("Scripting.FileSystemObject")

repoDir = objFSO.GetParentFolderName(WScript.ScriptFullName)
pythonExe = objShell.ExpandEnvironmentStrings("%POKJA_PYTHON%")
If pythonExe = "%POKJA_PYTHON%" Or Not objFSO.FileExists(pythonExe) Then
    pythonExe = objFSO.BuildPath(objFSO.GetParentFolderName(repoDir), "Runtime\WPy64-313110\python\python.exe")
End If
scriptPath = objFSO.BuildPath(repoDir, "sync_jadwal_all.py")

If Not objFSO.FileExists(pythonExe) Or Not objFSO.FileExists(scriptPath) Then
    WScript.Quit 2
End If

objShell.CurrentDirectory = repoDir
command = "cmd.exe /d /c " & Chr(34) & Chr(34) & pythonExe & Chr(34) & " " & Chr(34) & scriptPath & Chr(34) & Chr(34)
exitCode = objShell.Run(command, 0, True)
WScript.Quit exitCode
