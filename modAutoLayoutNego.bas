Attribute VB_Name = "modAutoLayoutNego"
Option Explicit

Private Const TARGET_SHEET As String = "7.2 Dengan Nego"
Private Const HELPER_SHEET As String = "_LayoutHelper"
Private Const MAX_ITEMS As Long = 500
Private Const MIN_ROW_HEIGHT As Double = 25
Private Const MAX_ROW_HEIGHT As Double = 409
Private Const EXTRA_HEIGHT As Double = 5

Private mIsRunning As Boolean
Private mLastSignature As String

Public Sub RapikanDaftarNego( _
    Optional ByVal ForceRefresh As Boolean = True, _
    Optional ByVal ShowWarnings As Boolean = True)
    Dim ws As Worksheet
    Dim helperWs As Worksheet
    Dim oldScreenUpdating As Boolean
    Dim oldEnableEvents As Boolean
    Dim oldCalculation As XlCalculation
    Dim skippedRows As String
    Dim activeCount As Long
    Dim hiddenCount As Long
    Dim firstItemRow As Long
    Dim lastItemRow As Long

    If mIsRunning Then Exit Sub
    mIsRunning = True

    On Error GoTo CleanFail

    Set ws = ThisWorkbook.Worksheets(TARGET_SHEET)
    Set helperWs = GetOrCreateHelperSheet()
    GetItemBounds ws, firstItemRow, lastItemRow
    If firstItemRow <= 0 Or lastItemRow < firstItemRow Then
        mIsRunning = False
        Exit Sub
    End If

    ' Modul lama pernah menganggap footer sebagai slot item. Pulihkan blok
    ' Total/Dibulatkan/Terbilang/TTD sebelum evaluasi signature.
    EnsureFooterVisible ws, lastItemRow + 1

    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldCalculation = Application.Calculation

    ' Worksheet_Calculate dapat memanggil macro ini lagi setelah sumber HPS
    ' dihitung. Jika signature item belum berubah, jangan ukur/merge 500 baris
    ' sekali lagi; ini mengurangi waktu buka dan cetak tanpa mengubah hasil.
    If Not ForceRefresh Then
        If BuildLayoutSignature(ws) = mLastSignature Then GoTo CleanExit
    End If

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Merapikan uraian pekerjaan..."

    ' Pastikan hanya sumber dan target yang dihitung. Application.Calculate
    ' menghitung semua workbook terbuka dan membuat Workbook_Open lambat.
    ThisWorkbook.Worksheets("5. HPS").Calculate
    ws.Calculate

    ProcessItemRows ws, helperWs, firstItemRow, lastItemRow, skippedRows, activeCount, hiddenCount
    mLastSignature = BuildLayoutSignature(ws)

CleanExit:
    Application.StatusBar = False
    Application.Calculation = oldCalculation
    Application.EnableEvents = oldEnableEvents
    Application.ScreenUpdating = oldScreenUpdating
    mIsRunning = False

    If Len(skippedRows) > 0 And ShowWarnings Then
        MsgBox _
            "Perapian selesai." & vbCrLf & vbCrLf & _
            "Baris aktif: " & activeCount & vbCrLf & _
            "Baris kosong tersembunyi: " & hiddenCount & vbCrLf & vbCrLf & _
            "Baris berikut tidak digabung karena C:H berisi data atau formula:" & _
            vbCrLf & skippedRows, _
            vbExclamation, _
            "Periksa Data di C:H"
    End If
    Exit Sub

CleanFail:
    MsgBox _
        "Macro perapian gagal." & vbCrLf & _
        "Nomor error: " & Err.Number & vbCrLf & _
        "Keterangan: " & Err.Description, _
        vbCritical, _
        "Rapikan Daftar Nego"
    Resume CleanExit
End Sub

Public Sub AutoRapikanJikaPerlu(ByVal ws As Worksheet)
    Dim currentSignature As String

    If mIsRunning Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If ws.Name <> TARGET_SHEET Then Exit Sub

    currentSignature = BuildLayoutSignature(ws)
    If currentSignature = mLastSignature Then Exit Sub

    RapikanDaftarNego False, False
End Sub

Public Sub PasangShortcutRapikan()
    Application.OnKey "%+q", "RapikanDaftarNego"
End Sub

Public Sub LepasShortcutRapikan()
    Application.OnKey "%+q"
End Sub

Public Sub ResetCacheLayout()
    mLastSignature = vbNullString
End Sub

Private Sub ProcessItemRows( _
    ByVal ws As Worksheet, _
    ByVal helperWs As Worksheet, _
    ByVal firstItemRow As Long, _
    ByVal lastItemRow As Long, _
    ByRef skippedRows As String, _
    ByRef activeCount As Long, _
    ByRef hiddenCount As Long)

    Dim rowNumber As Long
    Dim numberValue As Variant
    Dim descriptionText As String
    Dim descriptionArea As Range
    Dim sideArea As Range
    Dim calculatedHeight As Double

    For rowNumber = firstItemRow To lastItemRow
        numberValue = ws.Cells(rowNumber, "A").Value2
        descriptionText = SafeTrimmedText(ws.Cells(rowNumber, "B").Value2)
        Set descriptionArea = ws.Range("B" & rowNumber & ":H" & rowNumber)
        Set sideArea = ws.Range("C" & rowNumber & ":H" & rowNumber)

        If IsEmptyItem(numberValue, descriptionText) Then
            ws.Rows(rowNumber).Hidden = True
            hiddenCount = hiddenCount + 1
        Else
            ws.Rows(rowNumber).Hidden = False
            activeCount = activeCount + 1

            If RangeContainsDataOrFormula(sideArea) Then
                If Len(skippedRows) > 0 Then skippedRows = skippedRows & ", "
                skippedRows = skippedRows & CStr(rowNumber)
            Else
                EnsureCorrectMerge descriptionArea
                With descriptionArea
                    .WrapText = True
                    .ShrinkToFit = False
                    .HorizontalAlignment = xlLeft
                    .VerticalAlignment = xlCenter
                End With
                calculatedHeight = MeasureRequiredHeight(ws.Cells(rowNumber, "B"), descriptionArea, helperWs)
                ws.Rows(rowNumber).RowHeight = calculatedHeight
            End If
        End If
    Next rowNumber
End Sub

Private Sub GetItemBounds(ByVal ws As Worksheet, ByRef firstRow As Long, ByRef lastRow As Long)
    Dim rowNumber As Long
    Dim formulaText As String
    Dim footerRow As Long

    firstRow = 0
    For rowNumber = 1 To 100
        formulaText = CStr(ws.Cells(rowNumber, "A").Formula)
        If InStr(1, formulaText, "5. HPS", vbTextCompare) > 0 _
           And InStr(1, formulaText, "A2", vbTextCompare) > 0 Then
            firstRow = rowNumber
            Exit For
        End If
    Next rowNumber

    If firstRow = 0 Then
        If ws.Cells(8, "A").HasFormula Or Len(Trim$(CStr(ws.Cells(8, "A").Value2))) > 0 Then
            firstRow = 8
        Else
            firstRow = 9
        End If
    End If
    footerRow = FindFooterRow(ws, firstRow)
    If footerRow <= firstRow Then
        ' Fail closed: tanpa marker Total, jangan menyentuh visibilitas baris.
        firstRow = 0
        lastRow = 0
        Exit Sub
    End If
    lastRow = footerRow - 1
End Sub

Private Function FindFooterRow(ByVal ws As Worksheet, ByVal firstItemRow As Long) As Long
    Dim rowNumber As Long
    Dim cellValue As Variant
    Dim finalProbeRow As Long

    finalProbeRow = firstItemRow + MAX_ITEMS
    If finalProbeRow > ws.Rows.Count Then finalProbeRow = ws.Rows.Count

    For rowNumber = firstItemRow + 1 To finalProbeRow
        cellValue = ws.Cells(rowNumber, "M").Value2
        If Not IsError(cellValue) Then
            If StrComp(Trim$(CStr(cellValue)), "Total", vbTextCompare) = 0 Then
                FindFooterRow = rowNumber
                Exit Function
            End If
        End If
    Next rowNumber
End Function

Private Sub EnsureFooterVisible(ByVal ws As Worksheet, ByVal footerRow As Long)
    Dim lastCell As Range
    Dim lastContentRow As Long

    On Error Resume Next
    Set lastCell = ws.Range("A:R").Find( _
        What:="*", _
        After:=ws.Range("A1"), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    lastContentRow = footerRow
    If Not lastCell Is Nothing Then
        If lastCell.Row > lastContentRow Then lastContentRow = lastCell.Row
    End If
    ws.Rows(footerRow & ":" & lastContentRow).Hidden = False
End Sub

Private Function IsEmptyItem(ByVal numberValue As Variant, ByVal descriptionText As String) As Boolean
    Dim numberText As String

    If IsError(numberValue) Or IsNull(numberValue) Then
        IsEmptyItem = True
        Exit Function
    End If

    If IsNumeric(numberValue) Then
        If CDbl(numberValue) = 0 Then
            IsEmptyItem = True
            Exit Function
        End If
    End If

    numberText = SafeTrimmedText(numberValue)
    IsEmptyItem = Len(numberText) = 0 Or Len(descriptionText) = 0
End Function

Private Function SafeTrimmedText(ByVal value As Variant) As String
    If IsError(value) Or IsNull(value) Or IsEmpty(value) Then Exit Function
    SafeTrimmedText = Trim$(CStr(value))
End Function

Private Function RangeContainsDataOrFormula(ByVal targetRange As Range) As Boolean
    Dim cell As Range

    For Each cell In targetRange.Cells
        If cell.HasFormula Then
            RangeContainsDataOrFormula = True
            Exit Function
        End If
        If IsError(cell.Value2) Then
            RangeContainsDataOrFormula = True
            Exit Function
        End If
        If Len(SafeTrimmedText(cell.Value2)) > 0 Then
            RangeContainsDataOrFormula = True
            Exit Function
        End If
    Next cell
End Function

Private Sub EnsureCorrectMerge(ByVal targetRange As Range)
    Dim currentMerge As Range

    If targetRange.MergeCells Then
        Set currentMerge = targetRange.Cells(1, 1).MergeArea
        If currentMerge.Address = targetRange.Address Then Exit Sub
        currentMerge.UnMerge
    End If
    targetRange.Merge
End Sub

Private Function MeasureRequiredHeight( _
    ByVal sourceCell As Range, _
    ByVal mergedArea As Range, _
    ByVal helperWs As Worksheet) As Double

    Dim helperCell As Range
    Dim targetWidthPoints As Double
    Dim measuredHeight As Double
    Dim attempt As Long
    Dim currentWidth As Double
    Dim newColumnWidth As Double

    Set helperCell = helperWs.Range("A1")
    targetWidthPoints = mergedArea.Width
    helperWs.Cells.Clear

    With helperWs.Columns("A")
        .ColumnWidth = 8.43
        For attempt = 1 To 20
            currentWidth = .Width
            If currentWidth <= 0 Then Exit For
            If Abs(currentWidth - targetWidthPoints) <= 0.75 Then Exit For
            newColumnWidth = .ColumnWidth * targetWidthPoints / currentWidth
            If newColumnWidth < 0.1 Then newColumnWidth = 0.1
            If newColumnWidth > 255 Then newColumnWidth = 255
            .ColumnWidth = newColumnWidth
        Next attempt
    End With

    With helperCell
        .Value2 = sourceCell.Value2
        .WrapText = True
        .ShrinkToFit = False
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlTop
        .Font.Name = sourceCell.Font.Name
        .Font.Size = sourceCell.Font.Size
        .Font.Bold = sourceCell.Font.Bold
        .Font.Italic = sourceCell.Font.Italic
        .Font.Underline = sourceCell.Font.Underline
        .IndentLevel = sourceCell.IndentLevel
        .EntireRow.RowHeight = helperWs.StandardHeight
        .EntireRow.AutoFit
        measuredHeight = .RowHeight + EXTRA_HEIGHT
    End With

    If measuredHeight < MIN_ROW_HEIGHT Then measuredHeight = MIN_ROW_HEIGHT
    If measuredHeight > MAX_ROW_HEIGHT Then measuredHeight = MAX_ROW_HEIGHT
    MeasureRequiredHeight = measuredHeight
    helperCell.Clear
End Function

Private Function GetOrCreateHelperSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HELPER_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HELPER_SHEET
    End If
    ws.Visible = xlSheetVeryHidden
    Set GetOrCreateHelperSheet = ws
End Function

Private Function BuildLayoutSignature(ByVal ws As Worksheet) As String
    Dim valuesA As Variant
    Dim valuesB As Variant
    Dim parts() As String
    Dim rowIndex As Long
    Dim firstItemRow As Long
    Dim lastItemRow As Long
    Dim itemCount As Long
    Dim valueA As String
    Dim valueB As String

    GetItemBounds ws, firstItemRow, lastItemRow
    valuesA = ws.Range("A" & firstItemRow & ":A" & lastItemRow).Value2
    valuesB = ws.Range("B" & firstItemRow & ":B" & lastItemRow).Value2
    ReDim parts(1 To lastItemRow - firstItemRow + 1)

    For rowIndex = 1 To UBound(parts)
        If IsError(valuesA(rowIndex, 1)) Then
            valueA = "#ERROR"
        Else
            valueA = SafeTrimmedText(valuesA(rowIndex, 1))
        End If
        If IsError(valuesB(rowIndex, 1)) Then
            valueB = "#ERROR"
        Else
            valueB = SafeTrimmedText(valuesB(rowIndex, 1))
        End If
        parts(rowIndex) = valueA & ChrW$(30) & valueB
        If Len(valueA) > 0 Or Len(valueB) > 0 Then itemCount = itemCount + 1
    Next rowIndex

    BuildLayoutSignature = CStr(itemCount) & ChrW$(29) & Join(parts, ChrW$(31))
End Function
