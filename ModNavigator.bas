Attribute VB_Name = "ModNavigator"
' ============================================================
' NAVIGASI CEPAT - Lompat antar sheet dengan mudah
' ============================================================
Public Sub NavigasiSheet()
Attribute NavigasiSheet.VB_ProcData.VB_Invoke_Func = "N\n14"
    ' Tampilkan daftar sheet dan lompat ke sheet yang dipilih
    Dim sheetList As String
    Dim i As Long
    Dim maxDisplay As Long
    maxDisplay = 42
    
    ' Kelompokkan sheet berdasarkan kategori
    Dim categories As String
    categories = "=== SHEET UTAMA ===" & vbCrLf
    categories = categories & "  1. 1. Input Data" & vbCrLf
    categories = categories & "  2. 7. Input Data HPS + Penawar" & vbCrLf
    categories = categories & "  3. 7.2 Dengan Nego" & vbCrLf
    categories = categories & "  4. Klarifikasi Timpang Fix (2)" & vbCrLf
    categories = categories & vbCrLf
    categories = categories & "=== BERITA ACARA ===" & vbCrLf
    categories = categories & "  5. 3. BA Pembukaan Penawaran" & vbCrLf
    categories = categories & "  6. 5. BA Pembuktian Kualifikasi" & vbCrLf
    categories = categories & "  7. 7. BA Klarifikasi HS" & vbCrLf
    categories = categories & "  8. 9. BA Penetapan Pemenang" & vbCrLf
    categories = categories & "  9. 15. BA Timpang" & vbCrLf
    categories = categories & vbCrLf
    categories = categories & "=== EVALUASI ===" & vbCrLf
    categories = categories & " 10. 1.1 Eva. Adm" & vbCrLf
    categories = categories & " 11. 2. Eva. Teknis" & vbCrLf
    categories = categories & " 12. 3. KK Evaluasi Kualifikasi" & vbCrLf
    categories = categories & " 13. 4. Evaluasi Harga" & vbCrLf
    categories = categories & " 14. 4. Ringkasan Evaluasi Penawaran" & vbCrLf
    categories = categories & vbCrLf
    categories = categories & "=== LAINNYA ===" & vbCrLf
    categories = categories & " 15. 1. Surat Undangan" & vbCrLf
    categories = categories & " 16. 11. Surat Pengantar Hasil" & vbCrLf
    categories = categories & " 17. 2. Input Jadwal" & vbCrLf
    categories = categories & " 18. 3. Penjadwalan Print" & vbCrLf
    categories = categories & " 19. Daftar Isi" & vbCrLf
    
    Dim pilihan As String
    pilihan = InputBox(categories & vbCrLf & "Ketik nomor sheet (1-19):", "Navigasi Cepat", "")
    
    If pilihan = "" Then Exit Sub
    
    Dim targetName As String
    Select Case val(pilihan)
        Case 1: targetName = "1. Input Data"
        Case 2: targetName = "7. Input Data HPS + Penawar"
        Case 3: targetName = "7.2 Dengan Nego"
        Case 4: targetName = "Klarifikasi Timpang Fix (2)"
        Case 5: targetName = "3. BA Pembukaan Penawaran"
        Case 6: targetName = "5. BA Pembuktian Kualifikasi"
        Case 7: targetName = "7. BA Klarifikasi HS"
        Case 8: targetName = "9. BA Penetapan Pemenang"
        Case 9: targetName = "15. BA Timpang"
        Case 10: targetName = "1.1 Eva. Adm"
        Case 11: targetName = "2. Eva. Teknis"
        Case 12: targetName = "3. KK Evaluasi Kualifikasi"
        Case 13: targetName = "4. Evaluasi Harga"
        Case 14: targetName = "4. Ringkasan Evaluasi Penawaran"
        Case 15: targetName = "1. Surat Undangan"
        Case 16: targetName = "11. Surat Pengantar Hasil"
        Case 17: targetName = "2. Input Jadwal"
        Case 18: targetName = "3. Penjadwalan Print"
        Case 19: targetName = "Daftar Isi"
        Case Else
            MsgBox "Nomor tidak valid!", vbExclamation
            Exit Sub
    End Select
    
    On Error Resume Next
    ThisWorkbook.Sheets(targetName).Activate
    If Err.Number <> 0 Then
        MsgBox "Sheet '" & targetName & "' tidak ditemukan!", vbExclamation
    End If
    On Error GoTo 0
End Sub

' ============================================================
' QUICK DATA ENTRY - Input data tender cepat
' ============================================================
Public Sub InputDataTender()
    ' Form input data tender cepat via dialog
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("1. Input Data")
    
    Dim currentVal As String
    Dim newVal As String
    
    ' Array of input fields: Row, Col, Label
    Dim fields As Variant
    fields = Array( _
        Array(6, 6, "Nama Pekerjaan/Tender"), _
        Array(5, 6, "Sumber Dana"), _
        Array(11, 6, "Tahun Anggaran"), _
        Array(16, 6, "Nama PPK"), _
        Array(28, 5, "Tahun"), _
        Array(27, 5, "Bulan (angka)"), _
        Array(26, 5, "Tanggal") _
    )
    
    Dim i As Long
    For i = LBound(fields) To UBound(fields)
        Dim r As Long, c As Long, lbl As String
        r = fields(i)(0)
        c = fields(i)(1)
        lbl = fields(i)(2)
        
        currentVal = ""
        If Not IsError(ws.Cells(r, c).Value) Then
            currentVal = CStr(ws.Cells(r, c).Value)
        End If
        
        newVal = InputBox( _
            "Field: " & lbl & vbCrLf & _
            "Cell: " & ws.Cells(r, c).Address & vbCrLf & _
            "Nilai saat ini: " & currentVal & vbCrLf & vbCrLf & _
            "Masukkan nilai baru (kosongkan = skip):", _
            "Input Data Tender (" & (i + 1) & "/" & (UBound(fields) + 1) & ")", _
            currentVal)
        
        ' Jika user klik Cancel, keluar
        If StrPtr(newVal) = 0 Then
            MsgBox "Input dibatalkan.", vbInformation
            Exit Sub
        End If
        
        ' Jika ada nilai baru
        If newVal <> currentVal And newVal <> "" Then
            ws.Cells(r, c).Value = newVal
        End If
    Next i
    
    MsgBox "Data tender berhasil diupdate!", vbInformation
    ws.Activate
End Sub

' ============================================================
' GENERATE DAFTAR ISI (Table of Contents)
' ============================================================
Public Sub GenerateDaftarIsi()
    ' Buat/update sheet "Daftar Isi" dengan hyperlink ke semua sheet
    
    Dim tocSheet As Worksheet
    
    ' Cek apakah sheet Daftar Isi sudah ada
    On Error Resume Next
    Set tocSheet = ThisWorkbook.Sheets("Daftar Isi")
    On Error GoTo 0
    
    If tocSheet Is Nothing Then
        ' Buat sheet baru di posisi pertama
        Set tocSheet = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        tocSheet.Name = "Daftar Isi"
    Else
        tocSheet.Cells.Clear
    End If
    
    ' Header
    With tocSheet
        .Cells(1, 1).Value = "DAFTAR ISI"
        .Cells(1, 1).Font.Size = 16
        .Cells(1, 1).Font.Bold = True
        
        .Cells(2, 1).Value = "Klik nama sheet untuk langsung ke sheet tersebut"
        .Cells(2, 1).Font.Italic = True
        .Cells(2, 1).Font.Color = RGB(100, 100, 100)
        
        .Cells(4, 1).Value = "No."
        .Cells(4, 2).Value = "Nama Sheet"
        .Cells(4, 3).Value = "Kategori"
        .Cells(4, 1).Font.Bold = True
        .Cells(4, 2).Font.Bold = True
        .Cells(4, 3).Font.Bold = True
        
        ' Column widths
        .Columns(1).ColumnWidth = 6
        .Columns(2).ColumnWidth = 50
        .Columns(3).ColumnWidth = 20
    End With
    
    Dim row As Long
    row = 5
    Dim sheetNum As Long
    sheetNum = 0
    
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "Daftar Isi" Then
            sheetNum = sheetNum + 1
            
            tocSheet.Cells(row, 1).Value = sheetNum
            
            ' Hyperlink ke sheet
            tocSheet.Hyperlinks.Add _
                Anchor:=tocSheet.Cells(row, 2), _
                Address:="", _
                SubAddress:="'" & ws.Name & "'!A1", _
                TextToDisplay:=ws.Name
            
            ' Kategori
            Dim kategori As String
            Select Case True
                Case InStr(ws.Name, "Input Data") > 0: kategori = "Data Input"
                Case InStr(ws.Name, "BA ") > 0 Or InStr(ws.Name, "Berita") > 0: kategori = "Berita Acara"
                Case InStr(ws.Name, "Eva") > 0 Or InStr(ws.Name, "Evaluasi") > 0: kategori = "Evaluasi"
                Case InStr(ws.Name, "Daftar Hadir") > 0: kategori = "Daftar Hadir"
                Case InStr(ws.Name, "Surat") > 0: kategori = "Surat"
                Case InStr(ws.Name, "database") > 0 Or InStr(ws.Name, "data_") > 0 Or InStr(ws.Name, "list_") > 0: kategori = "Database"
                Case InStr(ws.Name, "HPS") > 0 Or InStr(ws.Name, "Harga") > 0 Or InStr(ws.Name, "Nego") > 0: kategori = "Harga/Nego"
                Case InStr(ws.Name, "Jadwal") > 0 Or InStr(ws.Name, "Penjadwalan") > 0: kategori = "Jadwal"
                Case InStr(ws.Name, "Timpang") > 0 Or InStr(ws.Name, "Klarifikasi") > 0: kategori = "Klarifikasi"
                Case Else: kategori = "Lainnya"
            End Select
            tocSheet.Cells(row, 3).Value = kategori
            
            row = row + 1
        End If
    Next ws
    
    ' Format tabel
    With tocSheet.Range("A4:C" & row - 1)
        .Borders.LineStyle = xlContinuous
    End With
    
    tocSheet.Activate
    MsgBox "Daftar Isi berhasil dibuat! (" & sheetNum & " sheet)", vbInformation
End Sub

