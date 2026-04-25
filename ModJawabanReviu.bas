Attribute VB_Name = "ModJawabanReviu"
Option Explicit

' ============================================================
' ModJawabanReviu — Simpan & Load jawaban reviu via XML/ZIP
'
' Simpan: Python baca ZIP langsung (Word boleh tetap buka)
'         -> simpan ke JSON (merge, tidak timpa placeholder)
' Load  : Python inject XML ke ZIP -> doc.Close -> reopen
'         -> user bisa menyaksikan dokumen reload dengan jawaban baru
' ============================================================

Private Const PY_EXE    As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\python\python.exe"
Private Const PY_ENGINE As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\jawaban_reviu_engine.py"

Private Function JalankanEngine(mode As String, docPath As String) As Integer
    ' Jalankan Python engine, tunggu selesai, return exit code
    Dim cmd As String
    cmd = """" & PY_EXE & """ """ & PY_ENGINE & """ " & mode & " """ & docPath & """"
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    JalankanEngine = wsh.Run(cmd, 0, True)
End Function

' ============================================================
' PUBLIC: Simpan — Python baca ZIP langsung, dokumen tetap buka
' ============================================================
Public Sub SimpanJawabanReviu()
    Dim doc As Document
    Set doc = ActiveDocument

    ' Pastikan dokumen sudah tersimpan agar XML di ZIP up-to-date
    If doc.Saved = False Then
        doc.Save
    End If

    Dim docPath As String
    docPath = doc.FullName

    Dim ret As Integer
    ret = JalankanEngine("simpan", docPath)

    If ret = 0 Then
        MsgBox "Jawaban reviu berhasil disimpan.", vbInformation, "Simpan Jawaban"
    Else
        MsgBox "Simpan gagal (exit code " & ret & ").", vbCritical, "Simpan Jawaban"
    End If
End Sub

' ============================================================
' PUBLIC: Load — inject ZIP -> close -> reopen di Word yg sama
'         Dokumen akan tutup sebentar lalu buka kembali
'         dengan jawaban yang sudah terisi.
' ============================================================
Public Sub LoadJawabanReviu()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim docPath As String
    docPath = doc.FullName

    If Dir(doc.Path & "\jawaban_reviu.json") = "" Then
        MsgBox "File jawaban_reviu.json tidak ditemukan." & Chr(10) & _
               "Simpan dari file sumber dulu, lalu copy JSON ke folder ini.", _
               vbCritical, "Load Jawaban"
        Exit Sub
    End If

    MsgBox "Dokumen akan ditutup sebentar lalu dibuka kembali" & Chr(10) & _
           "dengan jawaban reviu yang sudah terisi." & Chr(10) & Chr(10) & _
           "Klik OK untuk melanjutkan.", vbInformation, "Load Jawaban"

    Dim ret As Integer
    ret = JalankanEngine("load", docPath)

    If ret <> 0 Then
        MsgBox "Load gagal (exit code " & ret & ")." & Chr(10) & _
               "Pastikan dokumen ini terbuka di Word.", vbCritical, "Load Jawaban"
    End If
    ' Jika sukses: Python sudah reopen dokumen, VBA tidak perlu tindakan lanjut
End Sub
