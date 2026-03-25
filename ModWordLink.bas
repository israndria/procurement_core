Attribute VB_Name = "ModWordLink"
' ============================================================
' ModWordLink - Buka Word (merge) & Print (merge + Ctrl+P)
' ============================================================
' Semua merge dijalankan via Python (proses terpisah, Excel tidak hang)
' PENTING: File Word TIDAK PERNAH dimodifikasi - koneksi saat runtime saja

' v1.4: Path dinamis - tidak perlu hardcode, otomatis detect dari lokasi Excel
' Struktur folder: @ POKJA 2026\<Paket>\file.xlsm  →  @ POKJA 2026\V19_Scheduler\WPy64-313110\

Private Const WORD_BA As String = "1. Full Dokumen BA PK v1.docx"
Private Const WORD_REVIU As String = "2. Isi Reviu PK v1.docx"
Private Const WORD_DOKPIL As String = "3. Dokpil Full PK v1.docx"

Private Const SHEET_BA As String = "satu_data"
Private Const SHEET_REVIU As String = "list_reviu"
Private Const SHEET_DOKPIL As String = "list_dokpil"

' Printer terakhir yang dipilih (reset saat Excel ditutup)
Private m_LastPrinter As String


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

' ===== PRINT PDF (semua pakai RunPDF) =====

Public Sub PrintBAReviuPDF()
    RunPDF "pdf_bareviu", WORD_BA, SHEET_BA, "BA_REVIU_DPP"
End Sub

Public Sub PrintIsiReviuPDF()
    RunPDF "pdf_all", WORD_REVIU, SHEET_REVIU, "Isi_Reviu"
End Sub

Public Sub PrintDokpilPDF()
    RunPDF "pdf_dokpil", WORD_DOKPIL, SHEET_DOKPIL, "DOKPIL"
End Sub

Public Sub PrintUndanganPDF()
    RunPDF "pdf", WORD_BA, SHEET_BA, "Undangan"
End Sub

Public Sub PrintPembuktianPDF()
    RunPDF "pdf_pembuktian", WORD_BA, SHEET_BA, "BA Pembuktian & Nego"
End Sub

Public Sub PrintREvaluasiPDF()
    RunPDF "pdf_revaluasi", WORD_BA, SHEET_BA, "REvaluasi"
End Sub

Public Sub PrintPembuktianTimpangPDF()
    RunPDF "pdf_pembuktian_timpang", WORD_BA, SHEET_BA, "BA Pembuktian Timpang"
End Sub

' ===== CORE: Panggil Python script (non-blocking) =====

Private Sub RunPDF(mode As String, wordFile As String, sheetName As String, statusLabel As String)
    ' Tentukan apakah mode ini mendukung printer langsung
    ' Mode kompleks (pembuktian) hanya PDF karena butuh stitching multi-source
    Dim supportsPrinter As Boolean
    supportsPrinter = (mode <> "pdf_pembuktian" And mode <> "pdf_pembuktian_timpang")

    ' Tanya user: PDF atau Printer?
    Dim outputMode As String
    Dim printerName As String
    If supportsPrinter Then
        outputMode = ChooseOutputMode(printerName)
        If outputMode = "" Then Exit Sub  ' cancelled
    Else
        outputMode = "pdf"
    End If

    Dim kodePokja As String
    kodePokja = Trim(CStr(ThisWorkbook.Sheets("1. Input Data").Range("E14").Value))
    If kodePokja = "" Then kodePokja = "000"

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & wordFile

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")

    If outputMode = "printer" Then
        ' Mapping page range per mode (0 = all pages)
        Dim fromPage As Long, toPage As Long
        fromPage = 0: toPage = 0
        Select Case mode
            Case "pdf_bareviu": fromPage = 3: toPage = 6
            Case "pdf_revaluasi": fromPage = 30: toPage = 37
            Case "pdf": fromPage = 1: toPage = 2  ' Undangan
            ' pdf_all, pdf_dokpil = all pages (0,0)
        End Select

        cmd = Q(PyExe()) & " " & Q(ScriptDir() & "\word_merge.py") & " printer " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(sheetName) & " " & Q(printerName) & " " & fromPage & " " & toPage
        wsh.Run cmd, 0, False
        Application.StatusBar = "Printing " & statusLabel & " ke " & printerName & " ..."
    Else
        cmd = Q(PyExe()) & " " & Q(ScriptDir() & "\word_merge.py") & " " & mode & " " & Q(wordPath) & " " & Q(ThisWorkbook.FullName) & " " & Q(sheetName) & " " & Q(kodePokja)
        wsh.Run cmd, 0, False
        Application.StatusBar = "Membuat PDF " & statusLabel & "_" & kodePokja & " ..."
    End If

    Set wsh = Nothing
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

' ===== OUTPUT MODE: Popup pilih PDF atau Printer =====

Private Function ChooseOutputMode(ByRef outPrinter As String) As String
    On Error GoTo ErrChoose
    Dim choice As VbMsgBoxResult
    choice = MsgBox("Pilih output:" & vbCrLf & vbCrLf & _
                     "YES = Export PDF (seperti biasa)" & vbCrLf & _
                     "NO = Print ke Printer", _
                     vbYesNoCancel + vbQuestion, "Pilih Output")

    If choice = vbCancel Then
        ChooseOutputMode = ""
        Exit Function
    ElseIf choice = vbYes Then
        ChooseOutputMode = "pdf"
        Exit Function
    End If

    ' User pilih Printer
    outPrinter = PickPhysicalPrinter()
    If outPrinter = "" Then
        ChooseOutputMode = ""  ' cancelled atau tidak ada printer
    Else
        ChooseOutputMode = "printer"
        m_LastPrinter = outPrinter
    End If
    Exit Function

ErrChoose:
    MsgBox "Error di ChooseOutputMode:" & vbCrLf & Err.Description, vbCritical, "Debug"
    ChooseOutputMode = ""
End Function

' ===== PRINTER ENUMERATION: Scan & filter printer fisik =====

Private Function PickPhysicalPrinter() As String
    On Error GoTo ErrHandler

    Dim blacklist As Variant
    blacklist = Array("pdf", "xps", "onenote", "fax", "send to", "writer", "print to", "microsoft print", "onedriver")

    ' WMI query semua printer
    Dim wmi As Object
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    Dim printers As Object
    Set printers = wmi.ExecQuery("SELECT Name FROM Win32_Printer")

    Dim names() As String
    Dim count As Long
    count = 0

    Dim p As Object
    For Each p In printers
        Dim pName As String
        pName = p.Name
        Dim isVirtual As Boolean
        isVirtual = False
        Dim j As Long
        For j = LBound(blacklist) To UBound(blacklist)
            If InStr(1, LCase(pName), blacklist(j)) > 0 Then
                isVirtual = True
                Exit For
            End If
        Next j
        If Not isVirtual Then
            count = count + 1
            ReDim Preserve names(1 To count)
            names(count) = pName
        End If
    Next p

    If count = 0 Then
        MsgBox "Tidak ada printer fisik ditemukan!" & vbCrLf & vbCrLf & _
               "Pastikan driver printer sudah terinstall.", vbExclamation, "Printer"
        PickPhysicalPrinter = ""
        Exit Function
    End If

    ' Jika printer terakhir masih tersedia, tawarkan
    If m_LastPrinter <> "" Then
        Dim found As Boolean: found = False
        For j = 1 To count
            If names(j) = m_LastPrinter Then found = True: Exit For
        Next j
        If found Then
            Dim reuse As VbMsgBoxResult
            reuse = MsgBox("Pakai printer terakhir?" & vbCrLf & vbCrLf & m_LastPrinter, _
                           vbYesNo + vbQuestion, "Printer")
            If reuse = vbYes Then
                PickPhysicalPrinter = m_LastPrinter
                Exit Function
            End If
        End If
    End If

    ' Jika hanya 1 printer fisik, konfirmasi dulu
    If count = 1 Then
        Dim confirmOne As VbMsgBoxResult
        confirmOne = MsgBox("Print ke printer ini?" & vbCrLf & vbCrLf & names(1), _
                            vbOKCancel + vbQuestion, "Printer")
        If confirmOne = vbOK Then
            PickPhysicalPrinter = names(1)
        Else
            PickPhysicalPrinter = ""
        End If
        Exit Function
    End If

    ' Jika >1, tampilkan list pilihan
    Dim listStr As String
    For j = 1 To count
        listStr = listStr & j & ". " & names(j) & vbCrLf
    Next j

    Dim ans As String
    ans = InputBox("Pilih nomor printer:" & vbCrLf & vbCrLf & listStr, "Pilih Printer", "1")
    If ans = "" Then
        PickPhysicalPrinter = ""
        Exit Function
    End If

    Dim idx As Long
    On Error Resume Next
    idx = CLng(ans)
    On Error GoTo 0
    If idx >= 1 And idx <= count Then
        PickPhysicalPrinter = names(idx)
    Else
        MsgBox "Nomor tidak valid.", vbExclamation
        PickPhysicalPrinter = ""
    End If
    Exit Function

ErrHandler:
    MsgBox "Error di PickPhysicalPrinter:" & vbCrLf & Err.Description, vbCritical, "Debug"
    PickPhysicalPrinter = ""
End Function

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
    cmd = Q(PyExe()) & " " & Q(ScriptDir() & "\word_merge.py") & " " & mode & " " & Q(wordPath) & " " & Q(excelPath) & " " & Q(sheetName)
    
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

Private Function ScriptDir() As String
    ' Naik folder sampai ketemu V19_Scheduler (support subfolder dalam berapa pun)
    Dim folder As String
    folder = ThisWorkbook.Path
    Dim i As Integer
    For i = 1 To 10
        If Dir(folder & "\V19_Scheduler\WPy64-313110\python\python.exe") <> "" Then
            ScriptDir = folder & "\V19_Scheduler\WPy64-313110"
            Exit Function
        End If
        ' Naik 1 level
        Dim pos As Long
        pos = InStrRev(folder, "\")
        If pos <= 3 Then Exit For
        folder = Left(folder, pos - 1)
    Next i
    MsgBox "Python tidak ditemukan!" & vbCrLf & "Pastikan folder V19_Scheduler ada di dalam @ POKJA 2026", vbCritical
    ScriptDir = ""
End Function

Private Function PyExe() As String
    PyExe = ScriptDir() & "\python\pythonw.exe"
End Function

' Public wrapper agar Workbook_Open bisa akses ScriptDir
Public Function ScriptDir_Public() As String
    ScriptDir_Public = ScriptDir()
End Function

Public Sub ResetStatusBar()
    Application.StatusBar = False
End Sub

' ===== RELINK (update data source di Word templates via .NET zip — no dialog) =====
Public Sub RelinkTemplate()
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim cmd As String
    cmd = "powershell -ExecutionPolicy Bypass -File " & Q(ScriptDir() & "\relink_dotnet.ps1") & " -ExcelPath " & Q(ThisWorkbook.FullName)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True  ' hidden, blocking
    Set wsh = Nothing
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
    cmd = Q(PyExe()) & " " & Q(ScriptDir() & "\import_web_data.py") & " " & Q(excelPath)
    
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
