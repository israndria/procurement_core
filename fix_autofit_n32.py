"""
Fix: Update VBA di '7.2 Dengan Nego' dan ModAutoFit agar
Worksheet_Change dan FixSemuaBaris juga menangani kolom N (terbilang).
"""
import win32com.client
import pythoncom
import os

# VBA sheet event yang sudah diupdate - sekarang juga handle kolom N
SHEET_VBA = """
Private Sub FixRowHeight(ByVal cell As Range)
    On Error Resume Next
    
    If IsError(cell.Value) Then Exit Sub
    
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
        
        Dim ws As Worksheet
        Set ws = Me
        
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
        
        mergeArea.Rows(1).RowHeight = newHeight
    Else
        cell.WrapText = True
        cell.EntireRow.AutoFit
    End If
    
    On Error GoTo 0
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    
    For Each cell In Target
        ' Kolom B-H (2-8) DAN kolom N (14) - untuk terbilang
        If (cell.Column >= 2 And cell.Column <= 8) Or cell.Column = 14 Then
            If cell.Column = 14 Then
                FixRowHeight Me.Cells(cell.Row, 14)
            Else
                FixRowHeight Me.Cells(cell.Row, 2)
            End If
        End If
    Next cell
    
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub
"""

# VBA ModAutoFit yang sudah diupdate - FixSheet juga handle kolom N
MOD_AUTOFIT_VBA = """
' ModAutoFit - Smart auto-fit row height
' Menangani kolom B-H DAN kolom N (terbilang) di sheet target

Private Sub FixCellHeight(ws As Worksheet, cell As Range)
    On Error Resume Next
    
    If IsError(cell.Value) Then Exit Sub
    If IsEmpty(cell.Value) Then Exit Sub
    
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
        
        mergeArea.Rows(1).RowHeight = newHeight
    Else
        cell.WrapText = True
        cell.EntireRow.AutoFit
    End If
    
    On Error GoTo 0
End Sub

Private Sub FixSheet(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 2).End(-4162).Row ' xlUp
    
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
    
    ' Deteksi sheet aktif
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
        ' Fix hanya sheet aktif
        Application.StatusBar = "Memperbaiki: " & activeSheetName & "..."
        FixSheet ActiveSheet
        
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        MsgBox "Tinggi baris di '" & activeSheetName & "' sudah diperbaiki!" & vbCrLf & _
               "(Termasuk kolom B dan kolom N/Terbilang)", vbInformation
    Else
        ' Tanya user
        Dim jawab As VbMsgBoxResult
        jawab = MsgBox("Anda tidak di sheet target." & vbCrLf & vbCrLf & _
                        "Mau fix semua sheet berikut?" & vbCrLf & _
                        "- 7.2 Dengan Nego" & vbCrLf & _
                        "- Klarifikasi Timpang Fix (2)" & vbCrLf & vbCrLf & _
                        "(Termasuk kolom B dan kolom N/Terbilang)", _
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
        
        MsgBox msg & vbCrLf & "Selesai! (Termasuk kolom N/Terbilang)", vbInformation
    End If
End Sub
"""


def fix_autofit(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Fixing: {filepath}")
    
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
        print("File dibuka")
        
        # 1. Update sheet VBA for "7.2 Dengan Nego"
        print("\n[1] Update Worksheet_Change di '7.2 Dengan Nego'...")
        vb_project = wb.VBProject
        
        target_sheets = ["7.2 Dengan Nego", "Klarifikasi Timpang Fix (2)"]
        for sname in target_sheets:
            for comp in vb_project.VBComponents:
                if comp.Type == 100:  # Sheet module
                    try:
                        if comp.Properties("Caption").Value == sname:
                            # Clear existing code
                            if comp.CodeModule.CountOfLines > 0:
                                comp.CodeModule.DeleteLines(1, comp.CodeModule.CountOfLines)
                            comp.CodeModule.AddFromString(SHEET_VBA.strip())
                            print(f"    [OK] {sname}: VBA diupdate ({comp.CodeModule.CountOfLines} baris)")
                            break
                    except:
                        pass
        
        # 2. Update ModAutoFit
        print("\n[2] Update ModAutoFit...")
        for comp in vb_project.VBComponents:
            if comp.Name == "ModAutoFit":
                vb_project.VBComponents.Remove(comp)
                break
        
        new_mod = vb_project.VBComponents.Add(1)
        new_mod.Name = "ModAutoFit"
        new_mod.CodeModule.AddFromString(MOD_AUTOFIT_VBA.strip())
        print(f"    [OK] ModAutoFit diupdate ({new_mod.CodeModule.CountOfLines} baris)")
        
        # 3. Unlock cell E9 in protection
        print("\n[3] Unlock cell E9 di sheet '1. Input Data'...")
        try:
            ws = wb.Sheets("1. Input Data")
            ws.Unprotect(Password="pokja2026")
            ws.Range("E9").Locked = False
            ws.Protect(Password="pokja2026", AllowFormattingCells=True, AllowFormattingColumns=True, AllowFormattingRows=True)
            print("    [OK] Cell E9 unlocked")
        except Exception as e:
            print(f"    [WARN] {e}")
        
        # Save
        wb.Save()
        print("\n[OK] File disimpan!")
        
        print("\nPerubahan:")
        print("- Worksheet_Change: sekarang juga trigger untuk kolom N (14)")
        print("- FixSemuaBaris: sekarang juga fix kolom N (terbilang)")
        print("- Cell E9 di-unlock untuk kode unik")
        
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
    fix_autofit(r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026 (Improved).xlsm")
