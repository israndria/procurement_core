Attribute VB_Name = "ModSyncDraft"
' ============================================================
' ModSyncDraft - Sync & Diff Highlight untuk @ Master Data
' ============================================================
' Dua fungsi utama:
'   SyncDataDraft()   — baca semua field dari Excel → upsert ke Supabase data_snapshot
'   DiffHighlight()   — load snapshot dari Supabase → highlight KUNING sel yang berbeda
'
' Dipanggil dari tombol di sheet @ Master Data.
' Python helper: sync_draft.py (pola sync_kalender.py)
' ============================================================

Private Const MD_SHEET As String = "@ Master Data"
Private Const INPUT_SHEET As String = "1. Input Data"

' Warna highlight
Private Const CLR_DIFF     As Long = 16776960   ' Kuning
Private Const CLR_SAME     As Long = -4142       ' xlNone (no fill)

' ── Row constants (sama dengan ModDraftPaket) ──────────────
' INPUT DATA section (baris 3-22)
Private Const ROW_INPUT_START  As Integer = 3
Private Const ROW_INPUT_END    As Integer = 22
' REVIU section (baris 25-52)
Private Const ROW_REVIU_START  As Integer = 25
Private Const ROW_REVIU_END    As Integer = 52
' DOKPIL section (baris 56-66)
Private Const ROW_DOKPIL_START As Integer = 56
Private Const ROW_DOKPIL_END   As Integer = 66


' ============================================================
' SYNC: Baca semua field dari Excel → upsert data_snapshot
' ============================================================
Public Sub SyncDataDraft()
    Dim wsMD As Worksheet, wsInput As Worksheet
    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsMD    = ThisWorkbook.Sheets(MD_SHEET)
    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)
    On Error GoTo ErrHandler

    If wsMD Is Nothing Then
        MsgBox "Sheet '" & MD_SHEET & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If

    ' Ambil kode_tender dari 1. Input Data E5
    Dim kodeTender As String
    If Not wsInput Is Nothing Then
        kodeTender = Trim(CStr(wsInput.Range("E5").Value))
    End If
    If kodeTender = "" Then
        MsgBox "Kode Tender belum diisi di sheet '1. Input Data' cell E5.", vbExclamation
        Exit Sub
    End If

    ' Bangun snapshot JSON dari semua baris kolom C
    Dim snapshot As String
    snapshot = BuildSnapshotJSON(wsMD)

    ' Tulis input ke file temp
    Dim sd As String
    sd = ScriptDir()
    If sd = "" Then MsgBox "Script dir tidak ditemukan.", vbExclamation: Exit Sub

    Dim inputFile As String
    inputFile = sd & "\_sync_draft_input.json"
    Dim outputFile As String
    outputFile = sd & "\_sync_draft_output.json"

    ' Hapus output lama
    On Error Resume Next
    Kill outputFile
    On Error GoTo ErrHandler

    ' Tulis input JSON
    Dim payload As String
    payload = "{""kode_tender"":""" & kodeTender & """,""snapshot"":" & snapshot & "}"
    WriteUTF8 inputFile, payload

    ' Panggil Python
    Dim pyExe As String, pyScript As String
    pyExe    = sd & "\python\python.exe"
    pyScript = sd & "\sync_draft.py"

    Dim cmd As String
    cmd = """" & pyExe & """ """ & pyScript & """ save"

    Application.StatusBar = "Sync Data Draft: menyimpan ke Supabase..."
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True
    Application.StatusBar = False

    ' Baca hasil
    If Dir(outputFile) = "" Then
        MsgBox "sync_draft.py tidak menghasilkan output.", vbExclamation
        Exit Sub
    End If

    Dim result As String
    result = ReadUTF8(outputFile)

    Dim okVal As String
    okVal = ExtractVal(result, "ok")
    If okVal = "true" Then
        MsgBox "Data snapshot berhasil disimpan ke Supabase." & vbCrLf & _
               "Kode Tender: " & kodeTender, vbInformation, "Sync Data Draft"
        ' Reset highlight setelah sync — semua jadi sama dengan snapshot baru
        ClearHighlight wsMD
    Else
        Dim errMsg As String
        errMsg = ExtractVal(result, "error")
        MsgBox "Gagal sync: " & errMsg, vbExclamation, "Sync Data Draft"
    End If
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error SyncDataDraft: " & Err.Description, vbCritical
End Sub


' ============================================================
' DIFF HIGHLIGHT: Load snapshot → highlight sel yang berbeda
' Dipanggil otomatis setelah ParseDraftTerpilih selesai
' ============================================================
Public Sub DiffHighlight(Optional kodeTenderOverride As String = "")
    Dim wsMD As Worksheet, wsInput As Worksheet
    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsMD    = ThisWorkbook.Sheets(MD_SHEET)
    Set wsInput = ThisWorkbook.Sheets(INPUT_SHEET)
    On Error GoTo ErrHandler

    If wsMD Is Nothing Then Exit Sub

    ' Ambil kode_tender
    Dim kodeTender As String
    kodeTender = kodeTenderOverride
    If kodeTender = "" And Not wsInput Is Nothing Then
        kodeTender = Trim(CStr(wsInput.Range("E5").Value))
    End If
    If kodeTender = "" Then Exit Sub

    ' Tulis input
    Dim sd As String
    sd = ScriptDir()
    If sd = "" Then Exit Sub

    Dim inputFile As String:  inputFile  = sd & "\_sync_draft_input.json"
    Dim outputFile As String: outputFile = sd & "\_sync_draft_output.json"

    On Error Resume Next
    Kill outputFile
    On Error GoTo ErrHandler

    WriteUTF8 inputFile, "{""kode_tender"":""" & kodeTender & """}"

    Dim pyExe As String, pyScript As String
    pyExe    = sd & "\python\python.exe"
    pyScript = sd & "\sync_draft.py"

    Dim cmd As String
    cmd = """" & pyExe & """ """ & pyScript & """ load"

    Application.StatusBar = "Diff Highlight: memuat snapshot dari Supabase..."
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True
    Application.StatusBar = False

    If Dir(outputFile) = "" Then Exit Sub

    Dim result As String
    result = ReadUTF8(outputFile)

    Dim okVal As String
    okVal = ExtractVal(result, "ok")
    If okVal <> "true" Then Exit Sub

    ' Ekstrak blok snapshot
    Dim snapStart As Long
    snapStart = InStr(result, """snapshot""")
    If snapStart = 0 Then Exit Sub

    Dim bracePos As Long
    bracePos = InStr(snapStart, result, "{")
    If bracePos = 0 Then
        ' snapshot kosong {} atau null — tidak ada yang dibandingkan
        ClearHighlight wsMD
        Exit Sub
    End If

    ' Bandingkan tiap baris kolom C vs snapshot
    Dim r As Integer
    Dim ranges(2) As String
    ranges(0) = CStr(ROW_INPUT_START) & "-" & CStr(ROW_INPUT_END)
    ranges(1) = CStr(ROW_REVIU_START) & "-" & CStr(ROW_REVIU_END)
    ranges(2) = CStr(ROW_DOKPIL_START) & "-" & CStr(ROW_DOKPIL_END)

    Dim ri As Integer
    For ri = 0 To 2
        Dim parts() As String
        parts = Split(ranges(ri), "-")
        Dim rStart As Integer: rStart = CInt(parts(0))
        Dim rEnd   As Integer: rEnd   = CInt(parts(1))
        For r = rStart To rEnd
            Dim cellVal As String
            cellVal = Trim(CStr(wsMD.Cells(r, 3).Value))
            Dim key As String
            key = "r" & r
            Dim snapVal As String
            snapVal = ExtractVal(result, key)
            ' Highlight jika berbeda DAN snapshot ada (tidak kosong snapshot = belum pernah sync)
            If snapVal <> cellVal Then
                wsMD.Cells(r, 3).Interior.Color = CLR_DIFF
            Else
                wsMD.Cells(r, 3).Interior.ColorIndex = CLR_SAME
            End If
        Next r
    Next ri
    Exit Sub

ErrHandler:
    Application.StatusBar = False
End Sub


' ============================================================
' CLEAR HIGHLIGHT: hapus semua warna kuning di kolom C
' ============================================================
Public Sub ClearHighlight(wsMD As Worksheet)
    Dim r As Integer
    For r = ROW_INPUT_START To ROW_INPUT_END
        wsMD.Cells(r, 3).Interior.ColorIndex = CLR_SAME
    Next r
    For r = ROW_REVIU_START To ROW_REVIU_END
        wsMD.Cells(r, 3).Interior.ColorIndex = CLR_SAME
    Next r
    For r = ROW_DOKPIL_START To ROW_DOKPIL_END
        wsMD.Cells(r, 3).Interior.ColorIndex = CLR_SAME
    Next r
End Sub


' ============================================================
' BUILD SNAPSHOT JSON: semua nilai kolom C → {"r3":"val","r4":"val",...}
' ============================================================
Private Function BuildSnapshotJSON(wsMD As Worksheet) As String
    Dim sb As String
    sb = "{"
    Dim r As Integer
    Dim first As Boolean: first = True

    Dim i As Integer
    ' INPUT DATA
    For r = ROW_INPUT_START To ROW_INPUT_END
        If Not first Then sb = sb & ","
        sb = sb & """r" & r & """:""" & EscapeJSON(CStr(wsMD.Cells(r, 3).Value)) & """"
        first = False
    Next r
    ' REVIU
    For r = ROW_REVIU_START To ROW_REVIU_END
        sb = sb & ",""r" & r & """:""" & EscapeJSON(CStr(wsMD.Cells(r, 3).Value)) & """"
    Next r
    ' DOKPIL
    For r = ROW_DOKPIL_START To ROW_DOKPIL_END
        sb = sb & ",""r" & r & """:""" & EscapeJSON(CStr(wsMD.Cells(r, 3).Value)) & """"
    Next r

    sb = sb & "}"
    BuildSnapshotJSON = sb
End Function


' ============================================================
' HELPERS
' ============================================================
Private Function ScriptDir() As String
    ' Cari base dir WPy64-313110 dari lokasi workbook
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pokjaRoot As String
    pokjaRoot = fso.GetParentFolderName(ThisWorkbook.Path)
    Dim candidate As String
    candidate = pokjaRoot & "\V19_Scheduler\WPy64-313110"
    If Dir(candidate, vbDirectory) <> "" Then
        ScriptDir = candidate
        Exit Function
    End If
    ' Fallback: workbook langsung di WPy64-313110
    If Dir(ThisWorkbook.Path & "\python\python.exe") <> "" Then
        ScriptDir = ThisWorkbook.Path
    End If
End Function

Private Sub WriteUTF8(path As String, content As String)
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2: ado.Charset = "UTF-8": ado.Open
    ado.WriteText content
    ado.SaveToFile path, 2
    ado.Close
End Sub

Private Function ReadUTF8(path As String) As String
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2: ado.Charset = "UTF-8": ado.Open
    ado.LoadFromFile path
    ReadUTF8 = ado.ReadText
    ado.Close
End Function

Private Function EscapeJSON(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "")
    EscapeJSON = s
End Function

Private Function ExtractVal(json As String, key As String) As String
    Dim searchKey As String
    searchKey = """" & key & """:"
    Dim pos As Long
    pos = InStr(1, json, searchKey)
    If pos = 0 Then ExtractVal = "": Exit Function

    pos = pos + Len(searchKey)
    Do While Mid(json, pos, 1) = " ": pos = pos + 1: Loop

    If Mid(json, pos, 1) = """" Then
        pos = pos + 1
        Dim endPos As Long: endPos = pos
        Do While endPos <= Len(json)
            If Mid(json, endPos, 1) = """" And Mid(json, endPos - 1, 1) <> "\" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractVal = Mid(json, pos, endPos - pos)
    ElseIf Mid(json, pos, 4) = "null" Then
        ExtractVal = ""
    Else
        Dim endN As Long: endN = pos
        Do While endN <= Len(json) And InStr(",}]", Mid(json, endN, 1)) = 0
            endN = endN + 1
        Loop
        ExtractVal = Trim(Mid(json, pos, endN - pos))
    End If
End Function
