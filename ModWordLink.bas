Attribute VB_Name = "ModWordLink"
' ============================================================
' ModWordLink - Buka Word (merge) & Print (merge + Ctrl+P)
' ============================================================
' Semua merge dijalankan via Python (proses terpisah, Excel tidak hang)
' PENTING: File Word TIDAK PERNAH dimodifikasi - koneksi saat runtime saja

Private Const PY_SCRIPT As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\word_merge.py"
Private Const PY_IMPORT As String = "D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\import_web_data.py"

Private Const WORD_BA As String = "1. Full Dokumen BA PK v1.docx"
Private Const WORD_REVIU As String = "2. Isi Reviu PK v1.docx"
Private Const WORD_DOKPIL As String = "3. Dokpil Full PK v1.docx"

Private Const SHEET_BA As String = "satu_data"
Private Const SHEET_REVIU As String = "list_reviu"
Private Const SHEET_DOKPIL As String = "list_dokpil"

' ===== BUKA (merge + tampilkan Word) =====

Public Sub BukaBA()
    RunMerge "buka", WORD_BA, SHEET_BA
End Sub

Public Sub BukaReviu()
    RunMerge "buka", WORD_REVIU, SHEET_REVIU
End Sub

Public Sub BukaDokpil()
    RunMerge "buka", WORD_DOKPIL, SHEET_DOKPIL
End Sub

' ===== PRINT (merge + buka Print dialog) =====

Public Sub PrintBAReviuPDF()
    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & WORD_BA

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " pdf_bareviu " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(SHEET_BA) & " " & Q(kodePokja)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuat PDF BA_REVIU_DPP_" & kodePokja & ".pdf ..."
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Public Sub PrintIsiReviuPDF()
    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & WORD_REVIU

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " pdf_all " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(SHEET_REVIU) & " " & Q(kodePokja)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuat PDF Isi_Reviu_" & kodePokja & ".pdf ..."
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Public Sub PrintDokpilPDF()
    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & WORD_DOKPIL

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " pdf_dokpil " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(SHEET_DOKPIL) & " " & Q(kodePokja)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuat PDF DOKPIL_" & kodePokja & ".pdf ..."
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

' ===== PDF (merge + export halaman 1-2 ke PDF) =====

Public Sub PrintUndanganPDF()
    ' Ambil Kode Pokja dari cell E14 untuk nama file PDF
    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & WORD_BA

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " pdf " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(SHEET_BA) & " " & Q(kodePokja)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuat PDF Undangan_" & kodePokja & ".pdf ..."
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Public Sub PrintPembuktianPDF()
    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & WORD_BA

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " pdf_pembuktian " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(SHEET_BA) & " " & Q(kodePokja)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuat PDF BA Pembuktian & Nego_ " & kodePokja & " ..."
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Public Sub PrintREvaluasiPDF()
    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & WORD_BA

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " pdf_revaluasi " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(SHEET_BA) & " " & Q(kodePokja)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuat PDF REvaluasi_" & kodePokja & ".pdf ..."
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Public Sub PrintPembuktianTimpangPDF()
    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & WORD_BA

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " pdf_pembuktian_timpang " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(SHEET_BA) & " " & Q(kodePokja)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuat PDF BA Pembuktian Timpang_" & kodePokja & ".pdf ..."
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

' ===== CORE: Panggil Python script (non-blocking) =====

Private Sub RunMerge(mode As String, wordFile As String, sheetName As String)
    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & wordFile
    
    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If
    
    ' Simpan Excel dulu supaya Python baca data terbaru
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0
    
    Dim excelPath As String
    excelPath = ThisWorkbook.FullName
    
    ' Jalankan Python script (proses terpisah, Excel tidak hang)
    Dim cmd As String
    cmd = "pythonw " & Q(PY_SCRIPT) & " " & mode & " " & Q(wordPath) & " " & Q(excelPath) & " " & Q(sheetName)
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False  ' 0=hidden, False=non-blocking
    Set wsh = Nothing
    
    Application.StatusBar = "Merging " & wordFile & "... Word akan muncul sebentar lagi."
    
    ' Reset status bar setelah 5 detik
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Private Function Q(s As String) As String
    Q = Chr(34) & s & Chr(34)
End Function

Public Sub ResetStatusBar()
    Application.StatusBar = False
End Sub

' ===== IMPORT (baca file HTML LPSE dan isi ke excel) =====
Public Sub ImportHTML()
    Dim excelPath As String
    excelPath = ThisWorkbook.FullName
    
    Dim folder As String
    folder = ThisWorkbook.Path
    
    ' Hapus JSON lama dulu
    Dim jsonPath As String
    jsonPath = folder & "\_import_lpse.json"
    If Dir(jsonPath) <> "" Then Kill jsonPath
    
    ' Jalankan Python dan TUNGGU sampai selesai (blocking: True)
    Dim cmd As String
    cmd = "pythonw " & Q(PY_IMPORT) & " " & Q(excelPath)
    
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True  ' 0=hidden, True=TUNGGU (blocking)
    Set wsh = Nothing
    
    ' Cek apakah JSON sudah ada
    If Dir(jsonPath) = "" Then
        MsgBox "Tidak ada data LPSE yang berhasil dibaca." & vbCrLf & _
               "Pastikan file .html LPSE ada di folder ini: " & folder, vbExclamation
        Exit Sub
    End If
    
    ' Baca JSON secara manual (baris demi baris)
    Dim fNum As Integer
    fNum = FreeFile
    Dim jsonText As String
    Open jsonPath For Input As #fNum
    Dim lineText As String
    Do While Not EOF(fNum)
        Line Input #fNum, lineText
        jsonText = jsonText & lineText
    Loop
    Close #fNum
    
    ' Extract nilai dari JSON format sederhana
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("1. Input Data")
    
    ws.Range("E3").Value = ExtractJSON(jsonText, "MAK")
    ws.Range("E5").Value = ExtractJSON(jsonText, "Kode Tender")
    ws.Range("E6").Value = ExtractJSON(jsonText, "Nama Tender")
    ws.Range("E8").Value = ExtractJSON(jsonText, "Kode RUP")
    ws.Range("E10").Value = ExtractJSON(jsonText, "Nilai Pagu")
    ws.Range("E11").Value = ExtractJSON(jsonText, "Nilai HPS")
    
    ' Data dari PDF Lembar Disposisi Pokja
    ws.Range("E12").Value = ExtractJSON(jsonText, "Nomor Surat Dinas")
    ws.Range("E13").Value = ExtractJSON(jsonText, "Nomor PP")
    ws.Range("E14").NumberFormat = "@"  ' Format Text agar 075 tidak jadi 75
    ws.Range("E14").Value = ExtractJSON(jsonText, "Kode Pokja")
    ws.Range("F17").Value = ExtractJSON(jsonText, "Nama Dinas")
    
    ' Hapus file JSON sementara
    If Dir(jsonPath) <> "" Then Kill jsonPath
    
    MsgBox "Data LPSE berhasil diimport!", vbInformation, "Import Selesai"
End Sub

' Helper: ambil value dari JSON string sederhana (format "key": "value")
Private Function ExtractJSON(jsonText As String, key As String) As String
    Dim searchKey As String
    searchKey = """" & key & """: """
    Dim pos As Long
    pos = InStr(jsonText, searchKey)
    If pos = 0 Then
        ExtractJSON = ""
        Exit Function
    End If
    pos = pos + Len(searchKey)
    Dim endPos As Long
    endPos = InStr(pos, jsonText, """")
    If endPos = 0 Then
        ExtractJSON = ""
    Else
        ExtractJSON = Mid(jsonText, pos, endPos - pos)
    End If
End Function
