"""Fix sheet VBA yang tidak terdeteksi - cari by CodeName dan sheet name."""
import win32com.client
import pythoncom
import os

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

def fix_sheets(filepath):
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
        
        # List semua VBComponents untuk debug
        print("=== VB Components ===")
        target_sheets = ["7.2 Dengan Nego", "Klarifikasi Timpang Fix (2)"]
        
        for comp in vb_project.VBComponents:
            comp_type = comp.Type  # 100 = Document (sheet/thisworkbook)
            name = comp.Name
            lines = comp.CodeModule.CountOfLines
            
            if comp_type == 100:
                # Coba baca caption/codename
                try:
                    sheet_name = ""
                    for prop_idx in range(1, comp.Properties.Count + 1):
                        try:
                            prop = comp.Properties(prop_idx)
                            if prop.Name == "Name":
                                sheet_name = prop.Value
                        except:
                            pass
                    print(f"  Sheet: CodeName={name}, SheetName={sheet_name}, Lines={lines}")
                    
                    # Match by sheet name
                    if sheet_name in target_sheets:
                        print(f"    -> MATCH! Updating VBA...")
                        if lines > 0:
                            comp.CodeModule.DeleteLines(1, lines)
                        comp.CodeModule.AddFromString(SHEET_VBA.strip())
                        print(f"    -> OK ({comp.CodeModule.CountOfLines} baris)")
                except Exception as e:
                    print(f"  Sheet: {name} (error reading props: {e})")
            else:
                type_names = {1: "Module", 2: "Class", 3: "UserForm"}
                print(f"  {type_names.get(comp_type, f'Type{comp_type}')}: {name}, Lines={lines}")
        
        wb.Save()
        print("\n[OK] Disimpan!")
        
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
    fix_sheets(r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026 (Improved).xlsm")
