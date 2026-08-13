Option Explicit

Dim shell, fso, repoDir, codeRoot, pythonExe, scriptPath, command, rcTender, rcNonTender
Set shell = CreateObject("WScript.Shell")
Set fso = CreateObject("Scripting.FileSystemObject")

repoDir = fso.GetParentFolderName(WScript.ScriptFullName)
codeRoot = fso.GetParentFolderName(repoDir)
pythonExe = shell.Environment("Process")("POKJA_PYTHON")
If Len(pythonExe) = 0 Or Not fso.FileExists(pythonExe) Then
    pythonExe = codeRoot & "\Runtime\WPy64-313110\python\pythonw.exe"
Else
    pythonExe = Replace(pythonExe, "\python.exe", "\pythonw.exe")
End If
scriptPath = repoDir & "\scrape_spse.py"

If Not fso.FileExists(pythonExe) Or Not fso.FileExists(scriptPath) Then
    WScript.Quit 2
End If

shell.Environment("Process")("SCRAPE_TAHUN") = "2026"
shell.Environment("Process")("SCRAPE_KODE_LPSE") = ""
shell.CurrentDirectory = repoDir
command = Chr(34) & pythonExe & Chr(34) & " " & Chr(34) & scriptPath & Chr(34)

shell.Environment("Process")("SCRAPE_KATEGORI") = "Tender"
rcTender = shell.Run(command, 0, True)

shell.Environment("Process")("SCRAPE_KATEGORI") = "Non Tender"
rcNonTender = shell.Run(command, 0, True)

If rcTender <> 0 Then
    WScript.Quit rcTender
End If
WScript.Quit rcNonTender
