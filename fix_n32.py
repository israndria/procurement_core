"""
Fix bug: row height kolom N override kolom B.
Solusi: FixCellHeight sekarang pakai MAX(current height, new height).
"""
import win32com.client
import pythoncom
import os

MOD_AUTOFIT_VBA = """
' ModAutoFit - Smart auto-fit row height v3
' Fix: pakai MAX height agar kolom N tidak override kolom B

Private Sub FixCellHeight(ws As Worksheet, cell As Range)
    On Error Resume Next
    
    If IsError(cell.Value) Then Exit Sub
    If IsEmpty(cell.Value) Then Exit Sub
    If Len(CStr(cell.Value)) < 3 Then Exit Sub
    
    If cell.MergeCells Then
        Dim mergeArea As Range
        Set mergeArea = cell.mergeArea
        mergeArea.WrapText = True
        
        Dim mergedWidth As Double
        mergedWidth = 0
        Dim col As Range
        For Each col In mergeArea.Columns
            mergedWidth = mergedWidth + col.ColumnWidth
        Next col
        
        Dim tempRow As Long
        Dim tempCol As Long
        tempRow = 50000
        tempCol = 100
        
        With ws.Cells(tempRow, tempCol)
            .Value = mergeArea.Cells(1, 1).Value
            .ColumnWidth = mergedWidth
            .WrapText = True
            .Font.Name = mergeArea.Cells(1, 1).Font.Name
            .Font.Size = mergeArea.Cells(1, 1).Font.Size
            .Font.Bold = mergeArea.Cells(1, 1).Font.Bold
            .EntireRow.AutoFit
            
            Dim newHeight As Double
            newHeight = .RowHeight
            If newHeight < 15 Then newHeight = 15
            
            .ClearContents
            .ColumnWidth = 8.43
            .EntireRow.RowHeight = 15
        End With
        
        ' KUNCI: pakai MAX agar tidak memperkecil row yang sudah benar
        Dim currentHeight As Double
        currentHeight = mergeArea.Rows(1).RowHeight
        If newHeight > currentHeight Then
            mergeArea.Rows(1).RowHeight = newHeight
        End If
    Else
        cell.WrapText = True
        ' Untuk non-merged, simpan height lama dulu
        Dim oldHeight As Double
        oldHeight = cell.EntireRow.RowHeight
        cell.EntireRow.AutoFit
        ' Pastikan tidak menyusut
        If cell.EntireRow.RowHeight < oldHeight Then
            cell.EntireRow.RowHeight = oldHeight
        End If
    End If
    
    On Error GoTo 0
End Sub

Private Sub FixSheet(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
    
    Dim fixCount As Long
    fixCount = 0
    
    ' Fix kolom B (2) - uraian pekerjaan
    Dim i As Long
    For i = 1 To lastRow
        Dim cell As Range
        Set cell = ws.Cells(i, 2)
        
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            If Len(CStr(cell.Value)) > 5 Then
                FixCellHeight ws, cell
                fixCount = fixCount + 1
            End If
        End If
    Next i
    
    ' Fix kolom N (14) - terbilang harga
    Dim lastRowN As Long
    lastRowN = ws.Cells(ws.Rows.Count, 14).End(-4162).Row
    If lastRowN < 50 Then lastRowN = 50
    
    For i = 1 To lastRowN
        Set cell = ws.Cells(i, 14)
        
        If Not IsEmpty(cell.Value) And Not IsError(cell.Value) Then
            If Len(CStr(cell.Value)) > 5 Then
                FixCellHeight ws, cell
                fixCount = fixCount + 1
            End If
        End If
    Next i
    
    Application.StatusBar = ws.Name & ": " & fixCount & " cell diperbaiki"
End Sub

Public Sub FixSemuaBaris()
    Dim sheetNames As Variant
    sheetNames = Array("7.2 Dengan Nego", "Klarifikasi Timpang Fix (2)")
    
    Dim activeSheetName As String
    activeSheetName = ActiveSheet.Name
    
    Dim isTargetSheet As Boolean
    isTargetSheet = False
    Dim i As Long
    For i = LBound(sheetNames) To UBound(sheetNames)
        If activeSheetName = sheetNames(i) Then
            isTargetSheet = True
            Exit For
        End If
    Next i
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    If isTargetSheet Then
        Application.StatusBar = "Memperbaiki: " & activeSheetName & "..."
        FixSheet ActiveSheet
        
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        MsgBox "Selesai! Sheet '" & activeSheetName & "' telah diperbaiki." & vbCrLf & _
               "(Kolom B + Kolom N/Terbilang)", vbInformation
    Else
        Dim jawab As VbMsgBoxResult
        jawab = MsgBox("Anda tidak di sheet target." & vbCrLf & vbCrLf & _
                        "Mau fix semua sheet?" & vbCrLf & _
                        "- 7.2 Dengan Nego" & vbCrLf & _
                        "- Klarifikasi Timpang Fix (2)", _
                        vbYesNo + vbQuestion, "Fix Semua Baris")
        
        If jawab = vbNo Then
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            Exit Sub
        End If
        
        Dim msg As String
        msg = "Hasil:" & vbCrLf
        
        For i = LBound(sheetNames) To UBound(sheetNames)
            Dim ws As Worksheet
            On Error Resume Next
            Set ws = ThisWorkbook.Sheets(sheetNames(i))
            On Error GoTo 0
            
            If Not ws Is Nothing Then
                Application.StatusBar = "Memperbaiki: " & sheetNames(i) & "..."
                FixSheet ws
                msg = msg & "- " & sheetNames(i) & ": OK" & vbCrLf
            Else
                msg = msg & "- " & sheetNames(i) & ": TIDAK DITEMUKAN" & vbCrLf
            End If
            Set ws = Nothing
        Next i
        
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        MsgBox msg & vbCrLf & "Selesai!", vbInformation
    End If
End Sub
"""

def fix_all(filepath):
    filepath = os.path.abspath(filepath)
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        try: excel.Visible = False
        except: pass
        try: excel.DisplayAlerts = False
        except: pass
        
        wb = excel.Workbooks.Open(filepath)
        vb_project = wb.VBProject
        
        # 1. Replace ModAutoFit
        print("[1] Replace ModAutoFit v3 (MAX height)...")
        for comp in vb_project.VBComponents:
            if comp.Name == "ModAutoFit":
                vb_project.VBComponents.Remove(comp)
                break
        
        new_mod = vb_project.VBComponents.Add(1)
        new_mod.Name = "ModAutoFit"
        new_mod.CodeModule.AddFromString(MOD_AUTOFIT_VBA.strip())
        print(f"    [OK] {new_mod.CodeModule.CountOfLines} baris")
        
        # 2. Jalankan FixSemuaBaris langsung via Python
        print("\n[2] Jalankan fix langsung pada '7.2 Dengan Nego'...")
        ws = wb.Sheets("7.2 Dengan Nego")
        
        last_row = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
        print(f"    Last row kolom B: {last_row}")
        
        fixed = 0
        for i in range(1, last_row + 1):
            cell = ws.Cells(i, 2)
            try:
                val = cell.Value
                if val and not isinstance(val, (int, float)):
                    if len(str(val)) > 5 and cell.MergeCells:
                        ma = cell.MergeArea
                        ma.WrapText = True
                        
                        total_w = sum(ma.Columns(c).ColumnWidth for c in range(1, ma.Columns.Count + 1))
                        
                        temp = ws.Cells(50000, 100)
                        temp.Value = ma.Cells(1, 1).Value
                        temp.ColumnWidth = total_w
                        temp.WrapText = True
                        temp.Font.Name = ma.Cells(1, 1).Font.Name
                        temp.Font.Size = ma.Cells(1, 1).Font.Size
                        temp.Font.Bold = ma.Cells(1, 1).Font.Bold
                        temp.EntireRow.AutoFit()
                        
                        new_h = max(temp.RowHeight, 15)
                        
                        temp.ClearContents()
                        temp.ColumnWidth = 8.43
                        temp.EntireRow.RowHeight = 15
                        
                        old_h = ma.Rows(1).RowHeight
                        if new_h > old_h:
                            ma.Rows(1).RowHeight = new_h
                            print(f"    Row {i}: {old_h} -> {new_h} | {str(val)[:50]}")
                            fixed += 1
            except:
                pass
        
        # Fix kolom N juga
        print(f"\n    Fix kolom N...")
        for i in range(1, 50):
            cell = ws.Cells(i, 14)
            try:
                val = cell.Value
                if val and len(str(val)) > 5:
                    if cell.MergeCells:
                        ma = cell.MergeArea
                        ma.WrapText = True
                        total_w = sum(ma.Columns(c).ColumnWidth for c in range(1, ma.Columns.Count + 1))
                        
                        temp = ws.Cells(50000, 100)
                        temp.Value = ma.Cells(1, 1).Value
                        temp.ColumnWidth = total_w
                        temp.WrapText = True
                        temp.Font.Name = ma.Cells(1, 1).Font.Name
                        temp.Font.Size = ma.Cells(1, 1).Font.Size
                        temp.Font.Bold = ma.Cells(1, 1).Font.Bold
                        temp.EntireRow.AutoFit()
                        
                        new_h = max(temp.RowHeight, 15)
                        temp.ClearContents()
                        temp.ColumnWidth = 8.43
                        temp.EntireRow.RowHeight = 15
                        
                        old_h = ma.Rows(1).RowHeight
                        if new_h > old_h:
                            ma.Rows(1).RowHeight = new_h
                            print(f"    Row {i}: {old_h} -> {new_h} | {str(val)[:50]}")
                            fixed += 1
            except:
                pass
        
        print(f"\n    Total {fixed} baris diperbaiki")
        
        wb.Save()
        print("\n[OK] File disimpan!")
        
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
    fix_all(r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026 (Improved).xlsm")
