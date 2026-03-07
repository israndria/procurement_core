"""
Script lengkap untuk:
1. Copy file Excel asli
2. Embed fungsi terbilang sebagai VBA Module
3. Update formula yang merujuk terbilang.xlam
4. Tambah VBA utilitas (Export PDF, Auto-Print, Konversi Bulan)
5. Buat Named Ranges
6. Sederhanakan formula IF bersarang
"""
import win32com.client
import pythoncom
import os
import shutil

SOURCE_FILE = r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026.xlsm"
DEST_FILE = r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026 (Improved).xlsm"

# ===== FUNGSI TERBILANG VBA =====
TERBILANG_VBA = """
' Fungsi TERBILANG - Mengubah angka menjadi kata dalam Bahasa Indonesia
' Sudah built-in, tidak perlu file .xlam external lagi

Public Function terbilang(ByVal Angka As Variant) As String
    Dim Bilangan As Double
    Dim Hasil As String
    
    On Error GoTo ErrHandler
    
    If IsEmpty(Angka) Or Angka = "" Then
        terbilang = ""
        Exit Function
    End If
    
    Bilangan = CDbl(Angka)
    
    If Bilangan < 0 Then
        terbilang = "Minus " & terbilang_helper(Abs(Bilangan))
    Else
        terbilang = terbilang_helper(Bilangan)
    End If
    
    Exit Function
ErrHandler:
    terbilang = ""
End Function

Private Function terbilang_helper(ByVal Angka As Double) As String
    Dim Satuan As Variant
    Satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
    
    If Angka < 12 Then
        terbilang_helper = Satuan(Int(Angka))
    ElseIf Angka < 20 Then
        terbilang_helper = terbilang_helper(Angka - 10) & " Belas"
    ElseIf Angka < 100 Then
        terbilang_helper = terbilang_helper(Int(Angka / 10)) & " Puluh " & terbilang_helper(Int(Angka) Mod 10)
    ElseIf Angka < 200 Then
        terbilang_helper = "Seratus " & terbilang_helper(Angka - 100)
    ElseIf Angka < 1000 Then
        terbilang_helper = terbilang_helper(Int(Angka / 100)) & " Ratus " & terbilang_helper(Int(Angka) Mod 100)
    ElseIf Angka < 2000 Then
        terbilang_helper = "Seribu " & terbilang_helper(Angka - 1000)
    ElseIf Angka < 1000000 Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000)) & " Ribu " & terbilang_helper(Int(Angka) Mod 1000)
    ElseIf Angka < 1000000000# Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000000)) & " Juta " & terbilang_helper(Int(Angka) Mod 1000000)
    ElseIf Angka < 1000000000000# Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000000000#)) & " Miliar " & terbilang_helper(Int(Angka) Mod 1000000000#)
    ElseIf Angka < 1E+15 Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000000000000#)) & " Triliun " & terbilang_helper(Int(Angka) Mod 1000000000000#)
    End If
    
    terbilang_helper = Trim(terbilang_helper)
End Function
"""

# ===== VBA UTILITAS =====
UTILITY_VBA = """
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
"""

OLD_TERBILANG_PATH = "D:\\Dokumen\\3 @ POKJA 2025\\@ POKJA 2025\\terbilang.xlam"


def improve_excel():
    print("=" * 60)
    print("EXCEL IMPROVEMENT SCRIPT")
    print("=" * 60)
    
    # 1. Copy file
    print(f"\n[1] Copying file...")
    print(f"    From: {SOURCE_FILE}")
    print(f"    To:   {DEST_FILE}")
    shutil.copy2(SOURCE_FILE, DEST_FILE)
    print("    [OK] File berhasil dicopy")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.AutomationSecurity = 3
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(os.path.abspath(DEST_FILE))
        print("\n[2] File copy dibuka")
        
        vb_project = wb.VBProject
        
        # ===== EMBED FUNGSI TERBILANG =====
        print("\n[3] Menambah fungsi terbilang built-in...")
        
        # Cek apakah sudah ada module terbilang
        for comp in vb_project.VBComponents:
            if comp.Name == "ModTerbilang":
                vb_project.VBComponents.Remove(comp)
                break
        
        mod_terbilang = vb_project.VBComponents.Add(1)
        mod_terbilang.Name = "ModTerbilang"
        mod_terbilang.CodeModule.AddFromString(TERBILANG_VBA.strip())
        print(f"    [OK] Module 'ModTerbilang' ditambahkan ({mod_terbilang.CodeModule.CountOfLines} baris)")
        
        # ===== TAMBAH VBA UTILITAS =====
        print("\n[4] Menambah VBA utilitas...")
        
        for comp in vb_project.VBComponents:
            if comp.Name == "ModUtilitas":
                vb_project.VBComponents.Remove(comp)
                break
        
        mod_util = vb_project.VBComponents.Add(1)
        mod_util.Name = "ModUtilitas"
        mod_util.CodeModule.AddFromString(UTILITY_VBA.strip())
        print(f"    [OK] Module 'ModUtilitas' ditambahkan ({mod_util.CodeModule.CountOfLines} baris)")
        
        # ===== UPDATE FORMULA TERBILANG =====
        print("\n[5] Mengupdate formula terbilang.xlam -> terbilang built-in...")
        
        update_count = 0
        for idx in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(idx)
            try:
                ur = ws.UsedRange
                for row in range(1, min(ur.Rows.Count + 1, 200)):
                    for col in range(1, min(ur.Columns.Count + 1, 30)):
                        cell = ws.Cells(row, col)
                        try:
                            if cell.HasFormula:
                                formula = cell.Formula
                                if "terbilang.xlam" in formula.lower() or "terbilang.xlam" in formula:
                                    # Ganti referensi external dengan fungsi internal
                                    old_ref = f"'{OLD_TERBILANG_PATH}'!"
                                    new_formula = formula.replace(old_ref, "")
                                    
                                    # Juga coba variasi path lain
                                    if old_ref not in formula:
                                        # Cari pattern 'path\terbilang.xlam'!
                                        import re
                                        new_formula = re.sub(
                                            r"'[^']*terbilang\.xlam'!",
                                            "",
                                            formula,
                                            flags=re.IGNORECASE
                                        )
                                    
                                    if new_formula != formula:
                                        cell.Formula = new_formula
                                        print(f"    {ws.Name}!{cell.Address}: Updated")
                                        print(f"      Old: {formula}")
                                        print(f"      New: {new_formula}")
                                        update_count += 1
                        except:
                            pass
            except:
                pass
        
        print(f"    [OK] {update_count} formula diupdate")
        
        # ===== UPDATE FORMULA KONVERSI BULAN =====
        print("\n[6] Mencari formula IF bersarang konversi bulan...")
        
        bulan_update_count = 0
        for idx in range(1, wb.Sheets.Count + 1):
            ws = wb.Sheets(idx)
            try:
                ur = ws.UsedRange
                for row in range(1, min(ur.Rows.Count + 1, 200)):
                    for col in range(1, min(ur.Columns.Count + 1, 30)):
                        cell = ws.Cells(row, col)
                        try:
                            if cell.HasFormula:
                                formula = cell.Formula
                                # Deteksi formula IF bersarang untuk konversi bulan
                                if '"Januari"' in formula and '"January"' in formula and formula.count("IF(") > 5:
                                    # Cari cell referensi (misalnya R11)
                                    import re
                                    # Cari pattern pertama: IF(XX="Januari"
                                    match = re.search(r'IF\(([A-Z]+\d+)="Januari"', formula)
                                    if match:
                                        ref_cell = match.group(1)
                                        new_formula = f"=KonversiBulan({ref_cell})"
                                        cell.Formula = new_formula
                                        print(f"    {ws.Name}!{cell.Address}: Disederhanakan")
                                        print(f"      Old: {formula[:80]}...")
                                        print(f"      New: {new_formula}")
                                        bulan_update_count += 1
                        except:
                            pass
            except:
                pass
        
        print(f"    [OK] {bulan_update_count} formula disederhanakan")
        
        # ===== BUAT NAMED RANGES =====
        print("\n[7] Membuat Named Ranges...")
        
        named_ranges = {
            "NamaTender": "'1. Input Data'!$F$6",
            "LokasiPekerjaan": "'1. Input Data'!$F$7",
            "TahunAnggaran": "'1. Input Data'!$F$11",
            "NamaPPK": "'1. Input Data'!$F$16",
            "SumberDana": "'1. Input Data'!$F$5",
        }
        
        nr_count = 0
        for name, refers_to in named_ranges.items():
            try:
                # Hapus jika sudah ada
                try:
                    wb.Names(name).Delete()
                except:
                    pass
                wb.Names.Add(Name=name, RefersTo=f"={refers_to}")
                print(f"    [OK] {name} -> {refers_to}")
                nr_count += 1
            except Exception as e:
                print(f"    [SKIP] {name}: {e}")
        
        print(f"    [OK] {nr_count} Named Ranges dibuat")
        
        # ===== SIMPAN =====
        print("\n[8] Menyimpan file...")
        wb.Save()
        print("    [OK] File berhasil disimpan!")
        
        print(f"\n{'=' * 60}")
        print("RINGKASAN IMPROVEMENT")
        print(f"{'=' * 60}")
        print(f"  File output: {DEST_FILE}")
        print(f"  Fungsi terbilang: Embedded (tidak perlu .xlam lagi)")
        print(f"  Formula terbilang diupdate: {update_count}")
        print(f"  Formula bulan disederhanakan: {bulan_update_count}")
        print(f"  Named Ranges dibuat: {nr_count}")
        print(f"  VBA Utilitas ditambahkan:")
        print(f"    - ExportSheetKePDF (export 1 sheet ke PDF)")
        print(f"    - ExportMultiSheetPDF (export beberapa sheet BA ke 1 PDF)")
        print(f"    - PrintBeritaAcara (print semua BA sekaligus)")
        print(f"    - KonversiBulan() (fungsi formula)")
        print(f"    - NomorBulan() (fungsi formula)")
        print(f"    - terbilang() (fungsi formula)")
        
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
    improve_excel()
