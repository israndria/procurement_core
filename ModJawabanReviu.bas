Attribute VB_Name = "ModJawabanReviu"
Option Explicit

' ============================================================
' ModJawabanReviu — Simpan & Load jawaban reviu via XML
'
' Simpan: save + tutup -> Python baca ZIP -> JSON -> buka lagi
' Load  : tutup -> Python inject ZIP -> buka via shell
' ============================================================

Private Const PY_EXE    As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\python\python.exe"
Private Const PY_ENGINE As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\jawaban_reviu_engine.py"

Private Function JalankanEngine(mode As String, docPath As String) As Integer
    Dim cmd As String
    cmd = """" & PY_EXE & """ """ & PY_ENGINE & """ " & mode & " """ & docPath & """"
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    JalankanEngine = wsh.Run(cmd, 0, True)
End Function

Private Sub TungguFileBebas(docPath As String)
    Dim folder As String: folder = Left(docPath, InStrRev(docPath, "\"))
    Dim fname As String:  fname = Mid(docPath, InStrRev(docPath, "\") + 1)
    Dim lockFile As String: lockFile = folder & "~$" & fname
    Dim i As Integer
    For i = 1 To 20
        If Dir(lockFile) = "" Then Exit Sub
        Application.Wait Now + TimeValue("00:00:01")
    Next i
End Sub

' ============================================================
' PUBLIC: Simpan — save+tutup -> Python baca ZIP -> buka lagi
' ============================================================
Public Sub SimpanJawabanReviu()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim docPath As String
    docPath = doc.FullName

    doc.Save
    doc.Close SaveChanges:=False
    TungguFileBebas docPath

    Dim ret As Integer
    ret = JalankanEngine("simpan", docPath)

    Documents.Open docPath

    If ret = 0 Then
        MsgBox "Jawaban reviu berhasil disimpan.", vbInformation, "Simpan Jawaban"
    Else
        MsgBox "Simpan gagal (exit code " & ret & ").", vbCritical, "Simpan Jawaban"
    End If
End Sub

' ============================================================
' PUBLIC: Load — tutup -> Python inject ZIP + buka via shell
' ============================================================
Public Sub LoadJawabanReviu()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim docPath As String
    docPath = doc.FullName

    If Dir(doc.Path & "\jawaban_reviu.json") = "" Then
        MsgBox "File jawaban_reviu.json tidak ditemukan." & Chr(10) & _
               "Simpan dulu sebelum load.", vbCritical, "Load Jawaban"
        Exit Sub
    End If

    If MsgBox("Dokumen akan ditutup sementara untuk memuat data." & Chr(10) & _
              "Pastikan sudah simpan perubahan lain." & Chr(10) & Chr(10) & _
              "Lanjutkan?", vbYesNo + vbQuestion, "Load Jawaban") <> vbYes Then
        Exit Sub
    End If

    doc.Save
    doc.Close SaveChanges:=False
    TungguFileBebas docPath

    ' Python inject lalu buka sendiri via os.startfile
    JalankanEngine "load", docPath
End Sub
