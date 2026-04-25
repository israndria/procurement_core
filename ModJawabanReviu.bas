Attribute VB_Name = "ModJawabanReviu"
Option Explicit

' ============================================================
' ModJawabanReviu — Simpan & Load jawaban reviu via XML injection
'
' Format: jawaban_reviu.json di folder dokumen aktif
' Engine: jawaban_reviu_engine.py (Python, ZIP injection)
'
' Simpan: tutup doc → Python simpan XML → buka lagi
' Load  : tutup doc → Python inject XML → buka lagi
'
' Keuntungan: formatting 100% preserved (bold, strikethrough, indent, dll)
' ============================================================

Private Const PY_EXE    As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\python\python.exe"
Private Const PY_ENGINE As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\jawaban_reviu_engine.py"

' ============================================================
' INTERNAL: jalankan engine Python, tunggu selesai
' ============================================================
Private Function JalankanEngine(mode As String, docPath As String) As Integer
    Dim cmd As String
    cmd = """" & PY_EXE & """ """ & PY_ENGINE & """ " & mode & " """ & docPath & """"
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    JalankanEngine = wsh.Run(cmd, 1, True)
End Function

' ============================================================
' PUBLIC: Simpan jawaban — tutup, simpan XML via Python, buka lagi
' ============================================================
Public Sub SimpanJawabanReviu()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim docPath As String
    docPath = doc.FullName

    doc.Save
    doc.Close SaveChanges:=False

    Dim ret As Integer
    ret = JalankanEngine("simpan", docPath)

    Dim wdApp As Object
    Set wdApp = GetObject(, "Word.Application")
    wdApp.Documents.Open docPath

    If ret = 0 Then
        MsgBox "Jawaban reviu berhasil disimpan.", vbInformation, "Simpan Jawaban"
    Else
        MsgBox "Simpan selesai dengan exit code " & ret & "." & Chr(10) & _
               "Cek file jawaban_reviu.json di folder dokumen.", vbExclamation, "Simpan Jawaban"
    End If
End Sub

' ============================================================
' PUBLIC: Load jawaban — tutup, inject XML via Python, buka lagi
' ============================================================
Public Sub LoadJawabanReviu()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim docPath As String
    docPath = doc.FullName

    Dim jsonPath As String
    jsonPath = doc.Path & "\jawaban_reviu.json"
    If Dir(jsonPath) = "" Then
        MsgBox "File jawaban tidak ditemukan:" & Chr(10) & jsonPath & Chr(10) & Chr(10) & _
               "Simpan dulu sebelum load.", vbCritical, "Load Jawaban"
        Exit Sub
    End If

    If MsgBox("Dokumen akan ditutup sementara untuk inject data." & Chr(10) & _
              "Pastikan sudah simpan perubahan lain." & Chr(10) & Chr(10) & _
              "Lanjutkan?", vbYesNo + vbQuestion, "Load Jawaban") <> vbYes Then
        Exit Sub
    End If

    doc.Save
    doc.Close SaveChanges:=False

    Dim ret As Integer
    ret = JalankanEngine("load", docPath)

    Dim wdApp As Object
    Set wdApp = GetObject(, "Word.Application")
    wdApp.Documents.Open docPath

    If ret = 0 Then
        MsgBox "Jawaban reviu berhasil dimuat.", vbInformation, "Load Jawaban"
    Else
        MsgBox "Load selesai dengan exit code " & ret & "." & Chr(10) & _
               "Cek apakah data terisi dengan benar.", vbExclamation, "Load Jawaban"
    End If
End Sub
