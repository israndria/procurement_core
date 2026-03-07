"""Fix macro agar cerdas: deteksi ActiveSheet, hanya fix sheet yang sedang aktif."""
import win32com.client
import pythoncom
import os

FIXED_MODULE_CODE = """
Private Sub FixSheet(ws As Worksheet)
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim tempRow As Long
    Dim tempCol As Long
    tempRow = 50000
    tempCol = 100
    
    Dim i As Long
    For i = 1 To lastRow
        Dim cell As Range
        Set cell = ws.Cells(i, 2)
        
        If IsError(cell.Value) Then GoTo NextRow
        If IsEmpty(cell.Value) Then GoTo NextRow
        If cell.Value = "" Then GoTo NextRow
        
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
            
            With ws.Cells(tempRow, tempCol)
                .Value = cell.Value
                .ColumnWidth = mergedWidth
                .WrapText = True
                .Font.Name = cell.Font.Name
                .Font.Size = cell.Font.Size
                .Font.Bold = cell.Font.Bold
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
        
NextRow:
    Next i
End Sub

Public Sub FixSemuaBaris()
    ' SMART MACRO: Otomatis deteksi sheet yang sedang aktif
    ' - Di sheet "7.2 Dengan Nego" -> fix sheet itu saja
    ' - Di sheet "Klarifikasi Timpang Fix (2)" -> fix sheet itu saja
    ' - Di sheet lain -> tanya mau fix yang mana
    
    Dim targetSheets As Variant
    targetSheets = Array("7.2 Dengan Nego", "Klarifikasi Timpang Fix (2)")
    
    Dim currentSheet As String
    currentSheet = ActiveSheet.Name
    
    Dim found As Boolean
    found = False
    
    Dim i As Long
    For i = LBound(targetSheets) To UBound(targetSheets)
        If currentSheet = targetSheets(i) Then
            found = True
            Exit For
        End If
    Next i
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    If found Then
        ' Fix hanya sheet yang sedang aktif
        Application.StatusBar = "Memperbaiki: " & currentSheet & "..."
        FixSheet ActiveSheet
        
        Application.StatusBar = False
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        
        MsgBox "Selesai! Sheet '" & currentSheet & "' telah diperbaiki.", vbInformation
    Else
        ' Bukan di sheet target, tanya mau fix yang mana
        Dim pilihan As Long
        pilihan = MsgBox( _
            "Kamu sedang di sheet '" & currentSheet & "'." & vbCrLf & vbCrLf & _
            "Pilih:" & vbCrLf & _
            "  [Ya] = Fix semua sheet (Dengan Nego + Timpang)" & vbCrLf & _
            "  [Tidak] = Batal", _
            vbYesNo + vbQuestion, "Fix Baris")
        
        If pilihan = vbYes Then
            Dim ws As Worksheet
            For i = LBound(targetSheets) To UBound(targetSheets)
                On Error Resume Next
                Set ws = ThisWorkbook.Sheets(targetSheets(i))
                On Error GoTo 0
                If Not ws Is Nothing Then
                    Application.StatusBar = "Memperbaiki: " & targetSheets(i) & "..."
                    FixSheet ws
                End If
                Set ws = Nothing
            Next i
            
            Application.StatusBar = False
            Application.EnableEvents = True
            Application.ScreenUpdating = True
            
            MsgBox "Selesai! Semua sheet telah diperbaiki.", vbInformation
        Else
            Application.EnableEvents = True
            Application.ScreenUpdating = True
        End If
    End If
End Sub
"""

def fix_vba(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Fixing: {filepath}")
    
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
        
        vb_project = wb.VBProject
        
        for comp in vb_project.VBComponents:
            if comp.Name == "ModAutoFit":
                vb_project.VBComponents.Remove(comp)
                break
        
        new_module = vb_project.VBComponents.Add(1)
        new_module.Name = "ModAutoFit"
        new_module.CodeModule.AddFromString(FIXED_MODULE_CODE.strip())
        print(f"[OK] ModAutoFit diperbarui ({new_module.CodeModule.CountOfLines} baris)")
        
        wb.Save()
        print("[OK] File disimpan!")
        print("\nPerilaku macro FixSemuaBaris:")
        print("  - Di '7.2 Dengan Nego' -> fix sheet itu saja")
        print("  - Di 'Klarifikasi Timpang Fix (2)' -> fix sheet itu saja")
        print("  - Di sheet lain -> tanya fix semua atau batal")
        
    except Exception as e:
        print(f"[ERROR] {e}")
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()

if __name__ == "__main__":
    fix_vba(r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026.xlsm")
