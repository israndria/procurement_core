Attribute VB_Name = "ModSearchDrop"
' ============================================================
' Searchable Dropdown Helper untuk Sheet "1. Input Data"
' Menyediakan fungsi untuk mengisi ComboBox dengan data dari
' sheet referensi saat user klik cell E17 atau E19.
' ============================================================

Public Sub PopulateComboBox(cbx As Object, sourceRange As Range)
    cbx.Clear
    Dim cell As Range
    For Each cell In sourceRange
        If Trim(cell.Value) <> "" And cell.Value <> "-" Then
            cbx.AddItem cell.Value
        End If
    Next cell
End Sub

Public Function GetSourceRange(cellAddr As String) As Range
    Dim refSheet As Worksheet
    On Error Resume Next
    Set refSheet = ThisWorkbook.Sheets("0. Data Nama Pokja & PPK")
    On Error GoTo 0
    
    If refSheet Is Nothing Then
        Set GetSourceRange = Nothing
        Exit Function
    End If
    
    Select Case cellAddr
        Case "E17"
            Set GetSourceRange = refSheet.Range("B2:B200")
        Case "E19"
            Set GetSourceRange = refSheet.Range("C2:C200")
        Case Else
            Set GetSourceRange = Nothing
    End Select
End Function
