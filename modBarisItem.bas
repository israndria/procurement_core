Attribute VB_Name = "modBarisItem"
Option Explicit

Private Const SOURCE_SHEET As String = "5. HPS"
Private Const OUTPUT_SHEET As String = "7.2 Dengan Nego"
Private Const FIRST_SOURCE_ROW As Long = 2
Private Const LAST_SOURCE_ROW As Long = 501
Private Const FIRST_ITEM_ROW As Long = 9
Private Const LAST_ITEM_ROW As Long = 508
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
    Dim signature As String

    If m_IsRunning Then Exit Sub
    m_IsRunning = True
    On Error GoTo ErrorHandler

    oldEvents = Application.EnableEvents
    oldScreenUpdating = Application.ScreenUpdating

    Set wsSource = ThisWorkbook.Worksheets(SOURCE_SHEET)
    Set wsOutput = ThisWorkbook.Worksheets(OUTPUT_SHEET)

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
        outputLastRow = FIRST_ITEM_ROW - 1
    Else
        outputLastRow = FIRST_ITEM_ROW + (lastValidRow - FIRST_SOURCE_ROW)
    End If

    signature = CStr(validCount) & ":" & CStr(lastValidRow)
    If Not ForceRefresh And signature = m_LastSignature Then GoTo SafeExit

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    wsOutput.Rows(FIRST_ITEM_ROW & ":" & LAST_ITEM_ROW).Hidden = False
    If outputLastRow < LAST_ITEM_ROW Then
        wsOutput.Rows((outputLastRow + 1) & ":" & LAST_ITEM_ROW).Hidden = True
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
