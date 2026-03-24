Attribute VB_Name = "ModUtilitas"
' ============================================================
' MODUL UTILITAS - VBA untuk mempercepat kerja
' ============================================================

' Fungsi bantu konversi bulan Indonesia ke Inggris
Public Function KonversiBulan(bulanIndo As String) As String
    Select Case LCase(Trim(bulanIndo))
        Case "januari": KonversiBulan = "January"
        Case "februari": KonversiBulan = "February"
        Case "maret": KonversiBulan = "March"
        Case "april": KonversiBulan = "April"
        Case "mei": KonversiBulan = "May"
        Case "juni": KonversiBulan = "June"
        Case "juli": KonversiBulan = "July"
        Case "agustus": KonversiBulan = "August"
        Case "september": KonversiBulan = "September"
        Case "oktober": KonversiBulan = "October"
        Case "november": KonversiBulan = "November"
        Case "desember": KonversiBulan = "December"
        Case Else: KonversiBulan = ""
    End Select
End Function

' Fungsi bantu nomor bulan
Public Function NomorBulan(bulanIndo As String) As Integer
    Select Case LCase(Trim(bulanIndo))
        Case "januari": NomorBulan = 1
        Case "februari": NomorBulan = 2
        Case "maret": NomorBulan = 3
        Case "april": NomorBulan = 4
        Case "mei": NomorBulan = 5
        Case "juni": NomorBulan = 6
        Case "juli": NomorBulan = 7
        Case "agustus": NomorBulan = 8
        Case "september": NomorBulan = 9
        Case "oktober": NomorBulan = 10
        Case "november": NomorBulan = 11
        Case "desember": NomorBulan = 12
        Case Else: NomorBulan = 0
    End Select
End Function

' ============================================================
' EXPORT SHEETS KE PDF
' ============================================================
Public Sub ExportSheetKePDF()
    ' Export sheet aktif ke PDF
    Dim ws As Worksheet
    Set ws = ActiveSheet
    
    Dim filePath As String
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:=ws.Name & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Simpan sebagai PDF")
    
    If filePath = "False" Then Exit Sub
    
    ws.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    MsgBox "PDF berhasil disimpan!" & vbCrLf & filePath, vbInformation
End Sub

' ============================================================
' EXPORT BEBERAPA SHEET SEKALIGUS KE 1 PDF
' ============================================================
Public Sub ExportMultiSheetPDF()
    ' Export beberapa sheet Berita Acara ke 1 file PDF
    Dim sheetNames As Variant
    sheetNames = Array( _
        "3. BA Pembukaan Penawaran", _
        "7. BA Klarifikasi HS", _
        "7.1 Daftar Hadir Negosiasi", _
        "7.2 Dengan Nego", _
        "9. BA Penetapan Pemenang" _
    )
    
    ' Verifikasi sheet ada
    Dim validSheets() As String
    Dim count As Long
    count = 0
    
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            count = count + 1
            ReDim Preserve validSheets(1 To count)
            validSheets(count) = sheetNames(i)
        End If
        Set ws = Nothing
    Next i
    
    If count = 0 Then
        MsgBox "Tidak ada sheet yang ditemukan!", vbExclamation
        Exit Sub
    End If
    
    ' Pilih lokasi simpan
    Dim filePath As String
    filePath = Application.GetSaveAsFilename( _
        InitialFileName:="BA_Paket_" & Format(Now, "yyyymmdd") & ".pdf", _
        FileFilter:="PDF Files (*.pdf), *.pdf", _
        Title:="Simpan Berita Acara sebagai PDF")
    
    If filePath = "False" Then Exit Sub
    
    ' Select semua sheet valid
    ThisWorkbook.Sheets(validSheets).Select
    
    ActiveSheet.ExportAsFixedFormat _
        Type:=xlTypePDF, _
        Filename:=filePath, _
        Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, _
        IgnorePrintAreas:=False, _
        OpenAfterPublish:=True
    
    ' Kembali ke sheet pertama
    ThisWorkbook.Sheets(validSheets(1)).Select
    
    MsgBox count & " sheet berhasil diexport ke PDF!" & vbCrLf & filePath, vbInformation
End Sub

' ============================================================
' PRINT BERITA ACARA
' ============================================================
Public Sub PrintBeritaAcara()
    Dim sheetNames As Variant
    sheetNames = Array( _
        "3. BA Pembukaan Penawaran", _
        "5. BA Pembuktian Kualifikasi", _
        "7. BA Klarifikasi HS", _
        "7.1 Daftar Hadir Negosiasi", _
        "7.2 Dengan Nego", _
        "9. BA Penetapan Pemenang" _
    )
    
    Dim msg As String
    msg = "Sheet yang akan diprint:" & vbCrLf
    
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        msg = msg & "  - " & sheetNames(i) & vbCrLf
    Next i
    
    msg = msg & vbCrLf & "Lanjutkan?"
    
    If MsgBox(msg, vbYesNo + vbQuestion, "Print Berita Acara") = vbNo Then Exit Sub
    
    Dim printed As Long
    printed = 0
    
    For i = LBound(sheetNames) To UBound(sheetNames)
        Dim ws As Worksheet
        On Error Resume Next
        Set ws = ThisWorkbook.Sheets(sheetNames(i))
        On Error GoTo 0
        If Not ws Is Nothing Then
            ws.PrintOut
            printed = printed + 1
        End If
        Set ws = Nothing
    Next i
    
    MsgBox printed & " sheet telah dikirim ke printer!", vbInformation
End Sub

