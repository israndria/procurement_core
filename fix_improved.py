"""Fix file Improved: hapus Worksheet_Calculate dari sheet VBA"""
import win32com.client
import pythoncom
import os

LIGHT_SHEET_VBA = """
Private Sub Worksheet_Change(ByVal Target As Range)
    Dim cell As Range
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    
    For Each cell In Target
        If cell.Column >= 2 And cell.Column <= 8 Then
            FixRowHeight Me.Cells(cell.Row, 2)
        End If
    Next cell
    
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub FixRowHeight(ByVal cell As Range)
    On Error Resume Next
    
    If IsError(cell.Value) Then Exit Sub
    If IsEmpty(cell.Value) Then Exit Sub
    If cell.Value = "" Then Exit Sub
    
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
        
        With Me.Cells(50000, 100)
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
"""

TARGET_SHEETS = ["7.2 Dengan Nego", "Klarifikasi Timpang Fix (2)"]

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
        
        for sheet_name in TARGET_SHEETS:
            target_sheet = None
            for i in range(1, wb.Sheets.Count + 1):
                if wb.Sheets(i).Name == sheet_name:
                    target_sheet = wb.Sheets(i)
                    break
            
            if not target_sheet:
                continue
            
            code_name = target_sheet.CodeName
            for comp in vb_project.VBComponents:
                if comp.Name == code_name:
                    old = comp.CodeModule.CountOfLines
                    if old > 0:
                        comp.CodeModule.DeleteLines(1, old)
                    comp.CodeModule.AddFromString(LIGHT_SHEET_VBA.strip())
                    print(f"[OK] {sheet_name}: Worksheet_Calculate dihapus ({comp.CodeModule.CountOfLines} baris)")
                    break
        
        wb.Save()
        print("[OK] File disimpan! Sekarang ringan.")
        
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
    fix_vba(r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026 (Improved).xlsm")
