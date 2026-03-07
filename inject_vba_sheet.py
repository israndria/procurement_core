"""
Script untuk inject VBA Auto-Fit ke beberapa sheet sekaligus.
Sheet target:
  1. "7.2 Dengan Nego" (kolom B)
  2. "Klarifikasi Timpang Fix (2)" (kolom B)
"""
import win32com.client
import pythoncom
import os

def get_vba_code(sheet_display_name):
    """Generate VBA code for a sheet's code module"""
    return """
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

Private Sub Worksheet_Calculate()
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    
    Dim lastRow As Long
    lastRow = Me.Cells(Me.Rows.Count, "B").End(xlUp).Row
    If lastRow < 1 Then lastRow = 1
    
    Dim i As Long
    For i = 1 To lastRow
        If Me.Cells(i, 2).Value <> "" Then
            FixRowHeight Me.Cells(i, 2)
        End If
    Next i
    
    On Error GoTo 0
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

Private Sub FixRowHeight(ByVal cell As Range)
    On Error Resume Next
    
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
""".strip()


MODULE_CODE = """
Private Sub FixSheet(ws As Worksheet)
    ' Fix semua baris kolom B di sheet tertentu
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Dim tempRow As Long
    Dim tempCol As Long
    tempRow = 50000
    tempCol = 100
    
    Dim i As Long
    Dim fixCount As Long
    fixCount = 0
    
    For i = 1 To lastRow
        Dim cell As Range
        Set cell = ws.Cells(i, 2)
        
        If cell.Value <> "" Then
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
                fixCount = fixCount + 1
            Else
                cell.WrapText = True
                cell.EntireRow.AutoFit
                fixCount = fixCount + 1
            End If
        End If
    Next i
    
    FixSheet = fixCount
End Sub

Public Sub FixSemuaBaris()
    ' Fix semua baris di SEMUA sheet yang terdaftar
    ' Panggil dari: Developer > Macros > FixSemuaBaris > Run
    
    Dim sheetNames As Variant
    sheetNames = Array("7.2 Dengan Nego", "Klarifikasi Timpang Fix (2)")
    
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.StatusBar = "Memperbaiki tinggi baris..."
    
    Dim totalFix As Long
    totalFix = 0
    Dim msg As String
    msg = "Hasil:" & vbCrLf
    
    Dim i As Long
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
End Sub
"""

# Daftar sheet yang akan dipasangi VBA
TARGET_SHEETS = [
    "7.2 Dengan Nego",
    "Klarifikasi Timpang Fix (2)",
]

def fix_and_inject(filepath):
    filepath = os.path.abspath(filepath)
    print(f"Target: {filepath}")
    
    pythoncom.CoInitialize()
    excel = None
    wb = None
    
    try:
        excel = win32com.client.Dispatch("Excel.Application")
        excel.AutomationSecurity = 3
        excel.Visible = False
        excel.DisplayAlerts = False
        
        wb = excel.Workbooks.Open(filepath)
        print("File dibuka\n")
        
        vb_project = wb.VBProject
        
        # ===== Inject VBA ke setiap sheet =====
        for sheet_name in TARGET_SHEETS:
            print(f"--- Sheet: {sheet_name} ---")
            
            # Cari sheet
            target_sheet = None
            for i in range(1, wb.Sheets.Count + 1):
                if wb.Sheets(i).Name == sheet_name:
                    target_sheet = wb.Sheets(i)
                    break
            
            if not target_sheet:
                print(f"  [SKIP] Sheet tidak ditemukan!\n")
                continue
            
            code_name = target_sheet.CodeName
            print(f"  CodeName: {code_name}")
            
            # Inject VBA
            for comp in vb_project.VBComponents:
                if comp.Name == code_name:
                    old = comp.CodeModule.CountOfLines
                    if old > 0:
                        comp.CodeModule.DeleteLines(1, old)
                    comp.CodeModule.AddFromString(get_vba_code(sheet_name))
                    print(f"  [OK] VBA dipasang ({comp.CodeModule.CountOfLines} baris)")
                    break
            
            # Fix baris yang ada
            ws = target_sheet
            last_row = ws.Cells(ws.Rows.Count, 2).End(-4162).Row
            print(f"  Memperbaiki baris 1-{last_row}...")
            
            TEMP_ROW = 50000
            TEMP_COL = 100
            fix_count = 0
            
            for i in range(1, last_row + 1):
                cell = ws.Cells(i, 2)
                val = cell.Value
                if val is not None and str(val).strip() != "":
                    if cell.MergeCells:
                        merge_area = cell.MergeArea
                        merge_area.WrapText = True
                        
                        merged_width = 0
                        for c in range(1, merge_area.Columns.Count + 1):
                            merged_width += merge_area.Columns(c).ColumnWidth
                        
                        temp_cell = ws.Cells(TEMP_ROW, TEMP_COL)
                        temp_cell.Value = val
                        temp_cell.ColumnWidth = merged_width
                        temp_cell.WrapText = True
                        temp_cell.Font.Name = cell.Font.Name
                        temp_cell.Font.Size = cell.Font.Size
                        temp_cell.Font.Bold = cell.Font.Bold
                        temp_cell.EntireRow.AutoFit
                        
                        new_height = temp_cell.RowHeight
                        if new_height < 15:
                            new_height = 15
                        
                        temp_cell.ClearContents
                        temp_cell.ColumnWidth = 8.43
                        temp_cell.EntireRow.RowHeight = 15
                        
                        old_height = merge_area.Rows(1).RowHeight
                        merge_area.Rows(1).RowHeight = new_height
                        
                        if abs(old_height - new_height) > 0.5:
                            text_preview = str(val)[:55]
                            print(f"    Baris {i}: {old_height:.0f} -> {new_height:.0f}px | {text_preview}")
                        fix_count += 1
                    else:
                        cell.WrapText = True
                        cell.EntireRow.AutoFit
                        fix_count += 1
            
            print(f"  [OK] {fix_count} baris diproses\n")
        
        # ===== Update module macro =====
        for comp in vb_project.VBComponents:
            if comp.Name == "ModAutoFit":
                vb_project.VBComponents.Remove(comp)
                break
        
        new_module = vb_project.VBComponents.Add(1)
        new_module.Name = "ModAutoFit"
        new_module.CodeModule.AddFromString(MODULE_CODE.strip())
        print(f"[OK] Module 'ModAutoFit' diperbarui ({new_module.CodeModule.CountOfLines} baris)")
        
        wb.Save()
        print("[OK] File disimpan!")
        
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
    target = r"D:\Dokumen\@ POKJA 2026\@ BA PK 2026.xlsm"
    fix_and_inject(target)
