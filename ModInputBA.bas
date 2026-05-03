Attribute VB_Name = "ModInputBA"
' ============================================================
' ModInputBA - Isi sheet "0. Input BA" dari berbagai sumber:
'   1. Supabase: peserta_identitas, dokumen_penawaran
'   2. Sheet KK Evaluasi: SKP (C33), Hasil Pembuktian (C54)
'   3. Google Calendar: tanggal BA via V19 token
' ============================================================
' Trigger: Workbook_Open (auto) + tombol manual "Muat Input BA"

Private Const SB_URL As String = "%%SUPABASE_URL%%"
Private Const SB_KEY As String = "%%SUPABASE_KEY%%"

Private Const SHEET_INPUT    As String = "1. Input Data"
Private Const SHEET_BA       As String = "0. Input BA"
Private Const SHEET_KK       As String = "3. KK Evaluasi Kualifikasi"
Private Const CELL_KODE      As String = "E5"

' ── Row layout sheet "0. Input BA" ──────────────────────────
' Kolom B = label, Kolom C = nilai peserta 1, D = peserta 2
Private Const ROW_HEADER          As Integer = 1
' Blok Tanggal BA (kolom C = tanggal, D = hari otomatis)
Private Const ROW_TGL_PEMBUKAAN   As Integer = 3
Private Const ROW_TGL_PEMBUKTIAN  As Integer = 4   ' juga untuk klarif SKP, HS, penetapan
' Blok Identitas Peserta
Private Const ROW_NAMA_PERUSAHAAN As Integer = 7
Private Const ROW_NPWP            As Integer = 8
Private Const ROW_ALAMAT          As Integer = 9
Private Const ROW_DIREKTUR        As Integer = 10
' Blok Personel & Alat (dari PDF, diisi manual atau Tab 7)
Private Const ROW_PERSONEL_1      As Integer = 13
Private Const ROW_PERSONEL_2      As Integer = 14
Private Const ROW_ALAT_1          As Integer = 17
Private Const ROW_ALAT_2          As Integer = 18
Private Const ROW_ALAT_3          As Integer = 19
Private Const ROW_ALAT_4          As Integer = 20
Private Const ROW_ALAT_5          As Integer = 21
Private Const ROW_ALAT_6          As Integer = 22
' Blok Dokumen Penawaran
Private Const ROW_JML_DAFTAR      As Integer = 25
Private Const ROW_JML_KIRIM       As Integer = 26
Private Const ROW_JML_TDK_KIRIM   As Integer = 27
Private Const ROW_JML_TDK_LENGKAP As Integer = 28
Private Const ROW_JML_TDK_BUKA    As Integer = 29
' Blok Hasil (dari KK Evaluasi)
Private Const ROW_SKP             As Integer = 32
Private Const ROW_HASIL_PEMBUKTIAN As Integer = 33


' ============================================================
' FUNGSI UTAMA: dipanggil dari tombol + Workbook_Open
' ============================================================
Public Sub MuatInputBA()
    Dim wsInput As Worksheet, wsBA As Worksheet, wsKK As Worksheet
    On Error GoTo ErrHandler

    ' Pastikan sheet ada
    On Error Resume Next
    Set wsBA = ThisWorkbook.Sheets(SHEET_BA)
    On Error GoTo ErrHandler
    If wsBA Is Nothing Then
        MsgBox "Sheet '" & SHEET_BA & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If

    Set wsInput = ThisWorkbook.Sheets(SHEET_INPUT)
    Dim kodeTender As String
    kodeTender = Trim(wsInput.Range(CELL_KODE).Value)

    If kodeTender = "" Then
        MsgBox "Kode Tender belum diisi di sheet '1. Input Data' cell E5.", vbExclamation
        Exit Sub
    End If

    wsBA.Unprotect

    ' ── 1. Identitas peserta dari Supabase ──────────────────
    Call IsiIdentitasPeserta(wsBA, kodeTender)

    ' ── 2. Dokumen penawaran dari Supabase ──────────────────
    Call IsiDokumenPenawaran(wsBA, kodeTender)

    ' ── 3. SKP + Hasil Pembuktian dari KK Evaluasi ──────────
    On Error Resume Next
    Set wsKK = ThisWorkbook.Sheets(SHEET_KK)
    If Not wsKK Is Nothing Then
        ' SKP dari C33 (peserta 1) → isi ke "0. Input BA" DAN langsung ke sheet 6 W29
        ' (tidak pakai formula referensi di sheet 6 karena akan circular reference)
        Dim skpVal As String
        skpVal = Trim(CStr(wsKK.Cells(33, 3).Value))   ' C33
        If skpVal <> "" And skpVal <> "0" Then
            wsBA.Cells(ROW_SKP, 3).Value = skpVal
            ' Ekstrak angka dari "5 SKP" → isi W29 di sheet 6 langsung
            Dim wsKlarif As Worksheet
            Set wsKlarif = Nothing
            On Error Resume Next
            Set wsKlarif = ThisWorkbook.Sheets("6. BA KLARIF SKP ALAT")
            On Error GoTo ErrHandler
            If Not wsKlarif Is Nothing Then
                Dim skpAngka As Long
                Dim spacePos As Integer
                spacePos = InStr(skpVal, " ")
                If spacePos > 0 Then
                    skpAngka = CLng(Val(Left(skpVal, spacePos - 1)))
                Else
                    skpAngka = CLng(Val(skpVal))
                End If
                ' W29 = jumlah pekerjaan sedang berjalan (SKP paket = 5 - W29)
                ' Jadi W29 = 5 - skpAngka
                wsKlarif.Unprotect
                wsKlarif.Range("W29").Value = 5 - skpAngka
            End If
        End If

        ' Hasil pembuktian dari C54 (MS/TMS → teks lengkap)
        Dim hasilMs As String
        hasilMs = Trim(CStr(wsKK.Cells(54, 3).Value))  ' C54
        Select Case UCase(hasilMs)
            Case "MS":  wsBA.Cells(ROW_HASIL_PEMBUKTIAN, 3).Value = "Memenuhi"
            Case "TMS": wsBA.Cells(ROW_HASIL_PEMBUKTIAN, 3).Value = "Tidak Memenuhi"
            Case Else:  wsBA.Cells(ROW_HASIL_PEMBUKTIAN, 3).Value = hasilMs
        End Select
    End If
    On Error GoTo ErrHandler

    ' ── 4. Tanggal dari Google Calendar ─────────────────────
    ' (dinonaktifkan sementara — isi manual di sheet "0. Input BA" C3/C4)
    ' Call IsiTanggalDariGCal(wsBA, kodeTender)

    MsgBox "Sheet '0. Input BA' berhasil diperbarui.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error MuatInputBA: " & Err.Description, vbCritical
End Sub


' ============================================================
' Isi identitas peserta dari tabel peserta_identitas
' ============================================================
Private Sub IsiIdentitasPeserta(wsBA As Worksheet, kodeTender As String)
    Dim url As String
    url = SB_URL & "/rest/v1/peserta_identitas" & _
          "?kode_tender=eq." & kodeTender & _
          "&order=peserta_id.asc" & _
          "&limit=3"

    Dim json As String
    json = HttpGet(url)
    If json = "" Or json = "[]" Then Exit Sub

    ' Parse peserta: ambil sampai 3 peserta, isi kolom C/D/E
    Dim i As Integer
    Dim startPos As Long
    startPos = 1
    Dim col As Integer
    col = 3  ' kolom C = peserta pertama

    Do While col <= 5  ' maks 3 peserta (C, D, E)
        Dim objStart As Long, objEnd As Long
        objStart = InStr(startPos, json, "{")
        If objStart = 0 Then Exit Do
        objEnd = InStr(objStart, json, "}")
        If objEnd = 0 Then Exit Do

        Dim obj As String
        obj = Mid(json, objStart, objEnd - objStart + 1)

        wsBA.Cells(ROW_NAMA_PERUSAHAAN, col).Value = ExtractVal(obj, "nama_perusahaan")
        wsBA.Cells(ROW_NPWP, col).Value            = FormatNPWP(ExtractVal(obj, "npwp_raw"))
        wsBA.Cells(ROW_ALAMAT, col).Value           = ExtractVal(obj, "alamat")
        wsBA.Cells(ROW_DIREKTUR, col).Value         = ExtractVal(obj, "nama_direktur")
        wsBA.Cells(ROW_PERSONEL_1, col).Value       = ExtractVal(obj, "personel_1")
        wsBA.Cells(ROW_PERSONEL_2, col).Value       = ExtractVal(obj, "personel_2")
        wsBA.Cells(ROW_ALAT_1, col).Value           = ExtractVal(obj, "alat_1")
        wsBA.Cells(ROW_ALAT_2, col).Value           = ExtractVal(obj, "alat_2")
        wsBA.Cells(ROW_ALAT_3, col).Value           = ExtractVal(obj, "alat_3")
        wsBA.Cells(ROW_ALAT_4, col).Value           = ExtractVal(obj, "alat_4")
        wsBA.Cells(ROW_ALAT_5, col).Value           = ExtractVal(obj, "alat_5")
        wsBA.Cells(ROW_ALAT_6, col).Value           = ExtractVal(obj, "alat_6")

        startPos = objEnd + 1
        col = col + 1
    Loop
End Sub


' ============================================================
' Isi jumlah dokumen penawaran dari tabel dokumen_penawaran
' ============================================================
Private Sub IsiDokumenPenawaran(wsBA As Worksheet, kodeTender As String)
    Dim url As String
    url = SB_URL & "/rest/v1/dokumen_penawaran" & _
          "?kode_tender=eq." & kodeTender & _
          "&limit=1"

    Dim json As String
    json = HttpGet(url)
    If json = "" Or json = "[]" Then Exit Sub

    Dim objStart As Long
    objStart = InStr(1, json, "{")
    If objStart = 0 Then Exit Sub
    Dim objEnd As Long
    objEnd = InStr(objStart, json, "}")
    If objEnd = 0 Then Exit Sub

    Dim obj As String
    obj = Mid(json, objStart, objEnd - objStart + 1)

    Dim jmlDaftar As Integer, jmlKirim As Integer, jmlTdk As Integer
    jmlDaftar = CInt(Val(ExtractVal(obj, "jml_daftar")))
    jmlKirim  = CInt(Val(ExtractVal(obj, "jml_kirim")))
    jmlTdk    = CInt(Val(ExtractVal(obj, "jml_tidak_kirim")))

    wsBA.Cells(ROW_JML_DAFTAR, 3).Value      = jmlDaftar
    wsBA.Cells(ROW_JML_KIRIM, 3).Value       = jmlKirim
    wsBA.Cells(ROW_JML_TDK_KIRIM, 3).Value   = jmlTdk
    wsBA.Cells(ROW_JML_TDK_LENGKAP, 3).Value = CInt(Val(ExtractVal(obj, "jml_tidak_lengkap")))
    wsBA.Cells(ROW_JML_TDK_BUKA, 3).Value    = CInt(Val(ExtractVal(obj, "jml_tidak_dapat_dibuka")))
End Sub


' ============================================================
' Isi tanggal dari Google Calendar via V19 token
' Cari event berdasarkan kata kunci di title
' ============================================================
Private Sub IsiTanggalDariGCal(wsBA As Worksheet, kodeTender As String)
    Dim scriptDir As String
    scriptDir = ModWordLink.ScriptDir_Public()
    If scriptDir = "" Then Exit Sub

    Dim tokenPath As String
    tokenPath = scriptDir & "\token.json"
    If Dir(tokenPath) = "" Then Exit Sub

    ' Baca access_token dari token.json
    Dim fso As Object, f As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set f = fso.OpenTextFile(tokenPath, 1)
    Dim tokenJson As String
    tokenJson = f.ReadAll
    f.Close

    Dim accessToken As String
    accessToken = ExtractVal(tokenJson, "token")
    If accessToken = "" Then accessToken = ExtractVal(tokenJson, "access_token")
    If accessToken = "" Then Exit Sub

    ' Query GCal untuk event yang mengandung kode tender atau "Pembukaan"
    ' timeMin = 30 hari lalu, timeMax = 90 hari kedepan
    Dim timeMin As String, timeMax As String
    timeMin = Format(Now - 30, "yyyy-mm-dd") & "T00:00:00Z"
    timeMax = Format(Now + 90, "yyyy-mm-dd") & "T00:00:00Z"

    Dim gcUrl As String
    gcUrl = "https://www.googleapis.com/calendar/v3/calendars/primary/events" & _
            "?maxResults=50" & _
            "&singleEvents=true" & _
            "&orderBy=startTime" & _
            "&timeMin=" & timeMin & _
            "&timeMax=" & timeMax & _
            "&q=" & kodeTender

    Dim gcJson As String
    gcJson = HttpGetWithToken(gcUrl, accessToken)
    If gcJson = "" Then Exit Sub

    ' Parse event: cari "Pembukaan Dokumen Penawaran" dan "Pembuktian"
    Dim pos As Long
    pos = 1
    Dim tglPembukaan As String, tglPembuktian As String
    tglPembukaan = ""
    tglPembuktian = ""

    Do
        Dim itemStart As Long
        itemStart = InStr(pos, gcJson, """summary""")
        If itemStart = 0 Then Exit Do

        ' Ambil summary
        Dim sumStart As Long, sumEnd As Long
        sumStart = InStr(itemStart, gcJson, ":""") + 2
        sumEnd = InStr(sumStart, gcJson, """")
        Dim summary As String
        summary = Mid(gcJson, sumStart, sumEnd - sumStart)

        ' Cari tanggal mulai (dateTime atau date)
        Dim dtStart As Long
        dtStart = InStr(itemStart, gcJson, """start""")
        Dim tgl As String
        tgl = ""
        If dtStart > 0 And dtStart < InStr(pos + 1, gcJson, """summary""") + 500 Then
            Dim dtPos As Long
            dtPos = InStr(dtStart, gcJson, """date")  ' match "date" atau "dateTime"
            If dtPos > 0 Then
                Dim valS As Long, valE As Long
                valS = InStr(dtPos, gcJson, ":""") + 2
                valE = InStr(valS, gcJson, """")
                tgl = Left(Mid(gcJson, valS, valE - valS), 10)  ' ambil YYYY-MM-DD
            End If
        End If

        ' Cocokkan keyword
        If tglPembukaan = "" And (InStr(summary, "Pembukaan") > 0 Or InStr(summary, "pembukaan") > 0) Then
            tglPembukaan = tgl
        End If
        If tglPembuktian = "" And (InStr(summary, "Pembuktian") > 0 Or InStr(summary, "pembuktian") > 0 Or _
                                    InStr(summary, "Penetapan") > 0 Or InStr(summary, "penetapan") > 0) Then
            tglPembuktian = tgl
        End If

        pos = sumEnd + 1
        If tglPembukaan <> "" And tglPembuktian <> "" Then Exit Do
    Loop

    ' Tulis ke sheet: format date serial Excel
    If tglPembukaan <> "" Then
        wsBA.Cells(ROW_TGL_PEMBUKAAN, 3).Value = CDate(tglPembukaan)
        wsBA.Cells(ROW_TGL_PEMBUKAAN, 3).NumberFormat = "dd/mm/yyyy"
    End If
    If tglPembuktian <> "" Then
        wsBA.Cells(ROW_TGL_PEMBUKTIAN, 3).Value = CDate(tglPembuktian)
        wsBA.Cells(ROW_TGL_PEMBUKTIAN, 3).NumberFormat = "dd/mm/yyyy"
    End If
End Sub


' ============================================================
' Helper: HTTP GET dengan Supabase auth
' ============================================================
Private Function HttpGet(url As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 10000, 10000  ' resolve, connect, send, receive (ms)
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    On Error Resume Next
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        HttpGet = http.ResponseText
    Else
        HttpGet = ""
    End If
    On Error GoTo 0
End Function


' ============================================================
' Helper: HTTP GET dengan Bearer token (GCal)
' ============================================================
Private Function HttpGetWithToken(url As String, token As String) As String
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.SetRequestHeader "Authorization", "Bearer " & token
    http.SetRequestHeader "Accept", "application/json"
    On Error Resume Next
    http.Send
    If Err.Number = 0 And http.Status = 200 Then
        HttpGetWithToken = http.ResponseText
    Else
        HttpGetWithToken = ""
    End If
    On Error GoTo 0
End Function


' ============================================================
' Helper: Format NPWP raw 16 digit → XX.XXX.XXX.X-XXX.XXX
' ============================================================
Private Function FormatNPWP(raw As String) As String
    Dim d As String
    d = raw
    ' Hapus karakter non-digit
    Dim i As Integer, cleaned As String
    cleaned = ""
    For i = 1 To Len(d)
        If Mid(d, i, 1) >= "0" And Mid(d, i, 1) <= "9" Then
            cleaned = cleaned & Mid(d, i, 1)
        End If
    Next i
    If Len(cleaned) < 15 Then
        FormatNPWP = raw
        Exit Function
    End If
    ' Format: skip digit pertama (0), lalu XX.XXX.XXX.X-XXX.XXX
    Dim s As String
    s = cleaned
    If Left(s, 1) = "0" Then s = Mid(s, 2)  ' strip leading 0
    If Len(s) >= 15 Then
        FormatNPWP = Mid(s, 1, 2) & "." & Mid(s, 3, 3) & "." & _
                     Mid(s, 6, 3) & "." & Mid(s, 9, 1) & "-" & _
                     Mid(s, 10, 3) & "." & Mid(s, 13, 3)
    Else
        FormatNPWP = raw
    End If
End Function


' ============================================================
' Helper: ExtractVal JSON sederhana
' ============================================================
Private Function ExtractVal(json As String, key As String) As String
    Dim searchKey As String
    searchKey = """" & key & """:"
    Dim pos As Long
    pos = InStr(1, json, searchKey)
    If pos = 0 Then ExtractVal = "": Exit Function

    pos = pos + Len(searchKey)
    ' Skip whitespace
    Do While Mid(json, pos, 1) = " "
        pos = pos + 1
    Loop

    If Mid(json, pos, 1) = """" Then
        ' String value
        pos = pos + 1
        Dim endPos As Long
        endPos = pos
        Do While endPos <= Len(json)
            If Mid(json, endPos, 1) = """" And Mid(json, endPos - 1, 1) <> "\" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractVal = Mid(json, pos, endPos - pos)
    ElseIf Mid(json, pos, 4) = "null" Then
        ExtractVal = ""
    Else
        ' Numeric / boolean
        Dim endN As Long
        endN = pos
        Do While endN <= Len(json) And InStr(",}]", Mid(json, endN, 1)) = 0
            endN = endN + 1
        Loop
        ExtractVal = Trim(Mid(json, pos, endN - pos))
    End If
End Function
