Attribute VB_Name = "ModAutoFit"
' ModAutoFit - Smart auto-fit row height v4
' Ukur semua target cell per row, lalu pakai tinggi terbesar.

Private Function CellHeight(ws As Worksheet, cell As Range) As Double
    On Error GoTo gagal
    If IsError(cell.Value) Or IsEmpty(cell.Value) Then GoTo gagal
    If Len(CStr(cell.Value)) < 3 Then GoTo gagal

    Dim source As Range
    Set source = cell
    If cell.MergeCells Then Set source = cell.MergeArea
    source.WrapText = True

    Dim width As Double
    width = 0
    Dim col As Range
    For Each col In source.Columns
        width = width + col.ColumnWidth
    Next col
    If width <= 0 Then GoTo gagal

    Dim probe As Range
    Set probe = ws.Cells(50000, 100)
    With probe
        .ClearContents
        .Value = source.Cells(1, 1).Value
        .ColumnWidth = width
        .WrapText = True
        .Font.Name = source.Cells(1, 1).Font.Name
        .Font.Size = source.Cells(1, 1).Font.Size
        .Font.Bold = source.Cells(1, 1).Font.Bold
        .Font.Italic = source.Cells(1, 1).Font.Italic
        .EntireRow.RowHeight = 15
        .EntireRow.AutoFit
        CellHeight = .RowHeight
        .ClearContents
        .ColumnWidth = 8.43
        .EntireRow.RowHeight = 15
    End With
    If CellHeight < 15 Then CellHeight = 15
    Exit Function

gagal:
    CellHeight = 15
End Function

Private Sub FixSheet(ws As Worksheet)
    Dim lastRow As Long
    Dim lastRowB As Long
    lastRowB = ws.Cells(ws.Rows.count, 2).End(-4162).row
    Dim lastRowN As Long
    lastRowN = ws.Cells(ws.Rows.count, 14).End(-4162).row
    If lastRowN < 50 Then lastRowN = 50
    lastRow = IIf(lastRowB > lastRowN, lastRowB, lastRowN)
    
    Dim fixCount As Long
    fixCount = 0
    
    Dim i As Long
    For i = 1 To lastRow
        Dim requiredHeight As Double
        requiredHeight = 15
        Dim measuredHeight As Double
        Dim touched As Boolean
        touched = False

        Dim cellB As Range
        Set cellB = ws.Cells(i, 2)
        If Not IsEmpty(cellB.Value) And Not IsError(cellB.Value) Then
            If Len(CStr(cellB.Value)) > 5 Then
                measuredHeight = CellHeight(ws, cellB)
                If measuredHeight > requiredHeight Then requiredHeight = measuredHeight
                touched = True
            End If
        End If

        Dim cellN As Range
        Set cellN = ws.Cells(i, 14)
        If Not IsEmpty(cellN.Value) And Not IsError(cellN.Value) Then
            If Len(CStr(cellN.Value)) > 5 Then
                measuredHeight = CellHeight(ws, cellN)
                If measuredHeight > requiredHeight Then requiredHeight = measuredHeight
                touched = True
            End If
        End If

        If touched Then
            ws.Rows(i).RowHeight = requiredHeight
            fixCount = fixCount + 1
        End If
    Next i
    
    Application.StatusBar = ws.Name & ": " & fixCount & " cell diperbaiki"
End Sub

Public Sub FixSheetByName(sheetName As String)
    On Error GoTo selesai
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(sheetName)
    FixSheet ws
selesai:
    Application.StatusBar = False
End Sub

Public Sub FixSemuaBaris()
Attribute FixSemuaBaris.VB_ProcData.VB_Invoke_Func = "Q\n14"
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

