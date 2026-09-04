Attribute VB_Name = "modBarisItem"
Option Explicit

Private Const SOURCE_SHEET As String = "5. HPS"
Private Const OUTPUT_SHEET As String = "7.2 Dengan Nego"
Private Const FIRST_SOURCE_ROW As Long = 2
Private Const LAST_SOURCE_ROW As Long = 501
Private Const MAX_ITEMS As Long = 500

Private m_IsRunning As Boolean
Private m_LastSignature As String

Public Sub RefreshBarisItem(Optional ByVal ForceRefresh As Boolean = True)
    Dim oldEvents As Boolean
    Dim oldScreenUpdating As Boolean
    Dim wsSource As Worksheet
    Dim wsOutput As Worksheet
    Dim cell As Range
    Dim value As Variant
    Dim validCount As Long
    Dim lastValidRow As Long
    Dim outputLastRow As Long
    Dim firstItemRow As Long
    Dim lastItemRow As Long
    Dim signature As String

    If m_IsRunning Then Exit Sub
    m_IsRunning = True
    On Error GoTo ErrorHandler

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    Set wsSource = ThisWorkbook.Worksheets(SOURCE_SHEET)
    Set wsOutput = ThisWorkbook.Worksheets(OUTPUT_SHEET)
    GetItemBounds wsOutput, firstItemRow, lastItemRow
    If firstItemRow <= 0 Or lastItemRow < firstItemRow Then GoTo SafeExit

    For Each cell In wsSource.Range("A" & FIRST_SOURCE_ROW & ":A" & LAST_SOURCE_ROW).Cells
        value = cell.Value2
        If IsValidItemValue(value) Then
            validCount = validCount + 1
            lastValidRow = cell.Row
        End If
    Next cell

    If validCount > MAX_ITEMS Then
        MsgBox "Jumlah item melebihi kapasitas template (maksimal " & MAX_ITEMS & " item).", vbExclamation, "Kapasitas Item"
        GoTo SafeExit
    End If

    If lastValidRow = 0 Then
        outputLastRow = firstItemRow - 1
    Else
        outputLastRow = firstItemRow + (lastValidRow - FIRST_SOURCE_ROW)
    End If

    signature = CStr(validCount) & ":" & CStr(lastValidRow)
    If Not ForceRefresh And signature = m_LastSignature Then GoTo SafeExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    ' Footer Total/Dibulatkan/Terbilang/TTD bukan slot item. Pulihkan dulu
    ' bila workbook pernah diproses modul lama yang menyembunyikannya.
    EnsureFooterVisible wsOutput, lastItemRow + 1
    wsOutput.Rows(firstItemRow & ":" & lastItemRow).Hidden = False
    If outputLastRow < lastItemRow Then
        wsOutput.Rows((outputLastRow + 1) & ":" & lastItemRow).Hidden = True
    End If

    m_LastSignature = signature

SafeExit:
    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreenUpdating
    m_IsRunning = False
    Exit Sub

ErrorHandler:
    MsgBox "RefreshBarisItem gagal: " & Err.Description, vbExclamation, "Penyesuaian Baris Item"
    Resume SafeExit
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
        ' Fail closed: tanpa marker Total, jangan sembunyikan baris apa pun.
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

Private Function IsValidItemValue(ByVal value As Variant) As Boolean
    Dim textValue As String

    If IsError(value) Or IsEmpty(value) Then Exit Function

    If VarType(value) = vbString Then
        textValue = Trim$(CStr(value))
        If Len(textValue) = 0 Then Exit Function
        If IsNumeric(textValue) Then
            If CDbl(textValue) = 0 Then Exit Function
        End If
    ElseIf IsNumeric(value) Then
        If CDbl(value) = 0 Then Exit Function
    End If

    IsValidItemValue = True
End Function
