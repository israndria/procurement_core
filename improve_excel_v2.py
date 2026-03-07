"""
SCRIPT IMPROVEMENT LANJUTAN untuk file Excel Improved.
Implementasi:
1. Fix #REF! (clear broken formulas)
2. Sheet protection untuk "1. Input Data"
3. VBA Navigasi cepat
5. VBA Quick Data Entry (simplified)
6. Table of Contents sheet
7. Conditional Formatting (timpang >110%)
8. Auto-numbering helpers
"""
import win32com.client
import pythoncom
import os

# ===== VBA NAVIGASI + DATA ENTRY + AUTO-NUMBERING =====
VBA_NAVIGATOR = """
' ============================================================
' NAVIGASI CEPAT - Lompat antar sheet dengan mudah
' ============================================================
Public Sub NavigasiSheet()
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
    Select Case Val(pilihan)
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
"""

def improve_file(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Improving: {filepath}\n")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.AutomationSecurity = 3
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath)
        print("File dibuka")
        
        # ===== 1. FIX #REF! ERRORS =====
        print("\n[1] Fixing #REF! errors...")
        
        # Clear data_tender dan rincian_tender (semua #REF!)
        for sname in ["data_tender", "rincian_tender"]:
            try:
                ws = wb.Sheets(sname)
                ws.UsedRange.ClearContents
                print(f"    [OK] {sname}: konten dihapus (semua #REF!)")
            except:
                print(f"    [SKIP] {sname}: tidak ditemukan")
        
        # Fix 10. Dengan Nego Agit - clear #REF! formulas saja
        try:
            ws = wb.Sheets("10. Dengan Nego Agit")
            ur = ws.UsedRange
            fix_count = 0
            for row in range(1, ur.Rows.Count + 1):
                for col in range(1, min(ur.Columns.Count + 1, 40)):
                    cell = ws.Cells(row, col)
                    try:
                        if cell.HasFormula and "#REF!" in cell.Formula:
                            cell.ClearContents
                            fix_count += 1
                    except:
                        pass
            print(f"    [OK] 10. Dengan Nego Agit: {fix_count} formula #REF! dihapus")
        except:
            print("    [SKIP] 10. Dengan Nego Agit: tidak ditemukan")
        
        # ===== 2. SHEET PROTECTION =====
        print("\n[2] Melindungi sheet '1. Input Data'...")
        
        try:
            ws = wb.Sheets("1. Input Data")
            # Unlock semua cell input dulu
            ws.Cells.Locked = True  # Lock semua
            
            # Unlock cell-cell input (berdasarkan analisa)
            input_ranges = [
                "E4:H4", "E5:H5", "E6:H6", "E7:H7", "E8:H8",
                "E11:H11", "E12:H12", "E13:H13", 
                "E16:H16", "E17:H17",
                "E19:H24",  # Nama pokja
                "E26:H28",  # Tanggal
                "E31:H33",  # Lokasi, sumber dana
                "E34:H34",
                "E38:H38",
                "E40:H40",
                "E50:H59",  # Setting evaluasi
                "E62:H63",  # Bulan tahun
                "E67:H69",  # TTD
            ]
            
            for rng in input_ranges:
                try:
                    ws.Range(rng).Locked = False
                except:
                    pass
            
            # Protect dengan password sederhana (bisa diubah user)
            ws.Protect(Password="pokja2026", AllowFormattingCells=True, AllowFormattingColumns=True, AllowFormattingRows=True)
            print("    [OK] Sheet dilindungi (password: pokja2026)")
            print("    Cell input tetap bisa diedit")
        except Exception as e:
            print(f"    [ERROR] {e}")
        
        # ===== 3. VBA NAVIGASI + DATA ENTRY + TOC =====
        print("\n[3] Menambah VBA Navigasi, Data Entry, & Daftar Isi...")
        
        vb_project = wb.VBProject
        
        for comp in vb_project.VBComponents:
            if comp.Name == "ModNavigator":
                vb_project.VBComponents.Remove(comp)
                break
        
        nav_module = vb_project.VBComponents.Add(1)
        nav_module.Name = "ModNavigator"
        nav_module.CodeModule.AddFromString(VBA_NAVIGATOR.strip())
        print(f"    [OK] ModNavigator ditambahkan ({nav_module.CodeModule.CountOfLines} baris)")
        print("    Macro baru: NavigasiSheet, InputDataTender, GenerateDaftarIsi")
        
        # ===== 4. CONDITIONAL FORMATTING =====
        print("\n[4] Menambah Conditional Formatting (timpang >110%)...")
        
        try:
            ws = wb.Sheets("7.2 Dengan Nego")
            # Kolom Q (17) = % Terhadap HPS, baris 8-45
            target_range = ws.Range("Q8:Q45")
            
            # Hapus conditional formatting lama di range ini
            target_range.FormatConditions.Delete()
            
            # >110% = merah (timpang)
            fc1 = target_range.FormatConditions.Add(
                Type=1,  # xlCellValue
                Operator=5,  # xlGreater
                Formula1="1.1"
            )
            fc1.Interior.Color = 13421823  # Light red (RGB: 255, 199, 206)
            fc1.Font.Color = 255  # Red text
            fc1.Font.Bold = True
            
            # <80% = kuning (terlalu rendah)
            fc2 = target_range.FormatConditions.Add(
                Type=1,
                Operator=6,  # xlLess
                Formula1="0.8"
            )
            fc2.Interior.Color = 10092543  # Light yellow
            fc2.Font.Color = 26367  # Dark yellow/orange text
            
            # 80%-110% = hijau (wajar)
            fc3 = target_range.FormatConditions.Add(
                Type=1,
                Operator=1,  # xlBetween
                Formula1="0.8",
                Formula2="1.1"
            )
            fc3.Interior.Color = 13434828  # Light green
            fc3.Font.Color = 32768  # Green text
            
            print("    [OK] Conditional Formatting diterapkan di Q8:Q45")
            print("    > 110% = MERAH (timpang)")
            print("    < 80% = KUNING (terlalu rendah)")
            print("    80%-110% = HIJAU (wajar)")
        except Exception as e:
            print(f"    [ERROR] {e}")
        
        # ===== 5. RUN GENERATE DAFTAR ISI =====
        print("\n[5] Membuat sheet Daftar Isi...")
        
        try:
            # Buat sheet Daftar Isi secara langsung via Python
            toc_sheet = None
            try:
                toc_sheet = wb.Sheets("Daftar Isi")
                toc_sheet.Cells.Clear()
            except:
                toc_sheet = wb.Sheets.Add(Before=wb.Sheets(1))
                toc_sheet.Name = "Daftar Isi"
            
            # Header
            toc_sheet.Cells(1, 1).Value = "DAFTAR ISI"
            toc_sheet.Cells(1, 1).Font.Size = 16
            toc_sheet.Cells(1, 1).Font.Bold = True
            
            toc_sheet.Cells(2, 1).Value = "Klik nama sheet untuk navigasi. Atau gunakan macro: NavigasiSheet"
            toc_sheet.Cells(2, 1).Font.Italic = True
            toc_sheet.Cells(2, 1).Font.Color = 6710886  # Gray
            
            toc_sheet.Cells(4, 1).Value = "No."
            toc_sheet.Cells(4, 2).Value = "Nama Sheet"
            toc_sheet.Cells(4, 3).Value = "Kategori"
            toc_sheet.Cells(4, 1).Font.Bold = True
            toc_sheet.Cells(4, 2).Font.Bold = True
            toc_sheet.Cells(4, 3).Font.Bold = True
            
            toc_sheet.Columns(1).ColumnWidth = 6
            toc_sheet.Columns(2).ColumnWidth = 50
            toc_sheet.Columns(3).ColumnWidth = 20
            
            row = 5
            num = 0
            for i in range(1, wb.Sheets.Count + 1):
                ws = wb.Sheets(i)
                if ws.Name == "Daftar Isi":
                    continue
                num += 1
                
                toc_sheet.Cells(row, 1).Value = num
                
                # Hyperlink
                toc_sheet.Hyperlinks.Add(
                    Anchor=toc_sheet.Cells(row, 2),
                    Address="",
                    SubAddress=f"'{ws.Name}'!A1",
                    TextToDisplay=ws.Name
                )
                
                # Kategori
                name = ws.Name
                if "Input Data" in name:
                    cat = "Data Input"
                elif "BA " in name:
                    cat = "Berita Acara"
                elif "Eva" in name:
                    cat = "Evaluasi"
                elif "Daftar Hadir" in name:
                    cat = "Daftar Hadir"
                elif "Surat" in name:
                    cat = "Surat"
                elif "database" in name or "data_" in name or "list_" in name or "satu_data" in name:
                    cat = "Database"
                elif "HPS" in name or "Harga" in name or "Nego" in name:
                    cat = "Harga/Nego"
                elif "Jadwal" in name or "Penjadwalan" in name:
                    cat = "Jadwal"
                elif "Timpang" in name or "Klarifikasi" in name:
                    cat = "Klarifikasi"
                elif "DAFTAR SBU" in name:
                    cat = "Referensi"
                else:
                    cat = "Lainnya"
                
                toc_sheet.Cells(row, 3).Value = cat
                row += 1
            
            # Border
            toc_sheet.Range(f"A4:C{row-1}").Borders.LineStyle = 1
            
            print(f"    [OK] Daftar Isi dibuat ({num} sheet terdaftar)")
        except Exception as e:
            print(f"    [ERROR] {e}")
        
        # ===== SAVE =====
        print("\n[6] Menyimpan file...")
        wb.Save()
        print("    [OK] File disimpan!")
        
        print(f"\n{'=' * 60}")
        print("RINGKASAN IMPROVEMENT LANJUTAN")
        print(f"{'=' * 60}")
        print("1. #REF! errors: data_tender & rincian_tender dibersihkan")
        print("2. Sheet '1. Input Data' dilindungi (password: pokja2026)")
        print("3. Macro baru:")
        print("   - NavigasiSheet: navigasi cepat antar sheet")
        print("   - InputDataTender: form input data tender")
        print("   - GenerateDaftarIsi: regenerate daftar isi")
        print("4. Conditional Formatting: timpang >110% merah, <80% kuning")
        print("5. Sheet 'Daftar Isi' dibuat dengan hyperlink ke semua sheet")
        
    except Exception as e:
        print(f"[ERROR] {e}")
        import traceback
        traceback.print_exc()
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    improve_file(r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026 (Improved).xlsm")
