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
' Matrix kanonik peserta berada di C37:L53 (maksimal 10 peserta).
' C/D/E/I di baris lama hanya compatibility view; G = peserta terpilih.
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
Private Const MAX_PARTICIPANTS     As Integer = 10
Private Const MATRIX_HEADER_ROW    As Integer = 37
Private Const MATRIX_FIRST_ROW     As Integer = 38
Private Const MATRIX_LAST_ROW      As Integer = 53


' ============================================================
' FUNGSI UTAMA: dipanggil dari tombol + Workbook_Open
' ============================================================
Public Sub MuatInputBA(Optional tampilPesan As Boolean = True)
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
    EnsureInputBASlots wsBA

    ' ── 1. Identitas peserta dari Supabase ──────────────────
    Call IsiIdentitasPeserta(wsBA, kodeTender)

    ' ── 2. Dokumen penawaran dari Supabase ──────────────────
    Call IsiDokumenPenawaran(wsBA, kodeTender)

    ' ── 3. Conflict check personil/alat ─────────────────────
    Call IsiConflictCheck(wsBA, kodeTender)

    ' ── 4. SKP + Hasil Pembuktian dari KK Evaluasi ──────────
    On Error Resume Next
    Set wsKK = ThisWorkbook.Sheets(SHEET_KK)
    If Not wsKK Is Nothing Then
        ' Salin SKP + hasil dari KK matrix C:L ke Input BA matrix C:L.
        Dim kkCol As Integer, baCol As Integer
        Dim pesertaIdx As Integer
        For pesertaIdx = 1 To MAX_PARTICIPANTS
            kkCol = pesertaIdx + 2
            baCol = pesertaIdx + 2
            wsBA.Cells(52, baCol).Value = wsKK.Cells(33, kkCol).Value

            Dim hasilMs As String
            hasilMs = Trim(CStr(wsKK.Cells(54, kkCol).Value))
            Select Case UCase(hasilMs)
                Case "MS":  wsBA.Cells(53, baCol).Value = "Memenuhi"
                Case "TMS": wsBA.Cells(53, baCol).Value = "Tidak Memenuhi"
                Case Else:  wsBA.Cells(53, baCol).Value = hasilMs
            End Select
        Next pesertaIdx

        ' Sheet klarifikasi tetap winner-oriented, tetapi mengambil SKP dari
        ' peserta terpilih, bukan selalu dari kolom C.
        Dim selectedCol As Integer
        Dim selectedIdx As Integer
        selectedIdx = Val(wsBA.Range("F5").Value)
        If selectedIdx < 1 Or selectedIdx > MAX_PARTICIPANTS Then selectedIdx = 1
        selectedCol = selectedIdx + 2
        Dim skpVal As String
        skpVal = Trim(CStr(wsBA.Cells(52, selectedCol).Value))
        If skpVal <> "" And skpVal <> "0" Then
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
                wsKlarif.Unprotect
                wsKlarif.Range("W29").Value = 5 - skpAngka
            End If
        End If
    End If
    On Error GoTo ErrHandler

    If tampilPesan Then MsgBox "Sheet '0. Input BA' berhasil diperbarui.", vbInformation
    Exit Sub

ErrHandler:
    MsgBox "Error MuatInputBA: " & Err.Description, vbCritical
End Sub


' ============================================================
' REFRESH PAKET: satu tombol untuk sinkronisasi data utama.
' Semua operasi hanya GET/baca; penyimpanan ke Supabase tetap dilakukan
' lebih dulu melalui Streamlit Asisten Pokja.
' ============================================================
Public Sub RefreshPaket()
    Dim wsInput As Worksheet, wsMD As Worksheet
    On Error GoTo ErrHandler

    Set wsInput = ThisWorkbook.Sheets(SHEET_INPUT)
    Set wsMD = ThisWorkbook.Sheets("@ Master Data")
    Dim kodeTender As String
    kodeTender = Trim(CStr(wsInput.Range(CELL_KODE).Value))
    If kodeTender = "" Then
        MsgBox "Kode Tender belum diisi di sheet '1. Input Data' cell E5.", vbExclamation
        Exit Sub
    End If

    wsMD.Unprotect
    SiapkanPanelRefresh wsMD, kodeTender, "Memulai..."
    Application.ScreenUpdating = False

    Application.StatusBar = "Refresh Paket: memuat KK Evaluasi..."
    SiapkanPanelRefresh wsMD, kodeTender, "Memuat KK Evaluasi..."
    MuatKKEvaluasi False

    Application.StatusBar = "Refresh Paket: memuat Harga Penawaran..."
    SiapkanPanelRefresh wsMD, kodeTender, "Memuat Harga Penawaran..."
    MuatHargaPenawaran False

    Application.StatusBar = "Refresh Paket: memuat Input BA..."
    SiapkanPanelRefresh wsMD, kodeTender, "Memuat Input BA..."
    MuatInputBA False

    Application.StatusBar = False
    Application.ScreenUpdating = True
    SiapkanPanelRefresh wsMD, kodeTender, "Selesai"
    MsgBox "Refresh paket selesai." & vbCrLf & _
           "Data sumber tetap berasal dari Streamlit/Supabase; Excel hanya memuat ulang data terbaru.", _
           vbInformation, "Refresh Paket"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Error RefreshPaket: " & Err.Description, vbCritical
End Sub


Private Sub SiapkanPanelRefresh(wsMD As Worksheet, kodeTender As String, status As String)
    Dim wsKK As Worksheet, wsHP As Worksheet, wsBA As Worksheet
    Dim publishAt As String, priorRefresh As Variant, publishStatus As String
    On Error Resume Next
    Set wsKK = ThisWorkbook.Sheets(SHEET_KK)
    Set wsHP = ThisWorkbook.Sheets("6. Harga Penawaran")
    Set wsBA = ThisWorkbook.Sheets(SHEET_BA)
    On Error GoTo 0
    publishAt = BacaPublishTimestamp()
    priorRefresh = wsMD.Range("K25").Value
    If publishAt = "" Then
        publishStatus = "Belum publish"
    ElseIf IsDate(priorRefresh) Then
        If Replace(Left(publishAt, 19), "T", " ") > Format(CDate(priorRefresh), "yyyy-mm-dd hh:nn:ss") Then
            publishStatus = "STALE - publish lebih baru"
        Else
            publishStatus = "OK"
        End If
    Else
        publishStatus = "OK"
    End If

    With wsMD.Range("F24:I27")
        .UnMerge
        .ClearContents
        .Interior.Color = RGB(242, 242, 242)
        .Font.Bold = False
        .Borders.LineStyle = xlContinuous
    End With
    With wsMD.Range("F24:I24")
        .Merge
        .Value = "STATUS REFRESH PAKET"
        .Font.Bold = True
        .Interior.Color = RGB(0, 128, 128)
        .Font.Color = RGB(255, 255, 255)
    End With
    wsMD.Range("F25").Value = "Kode Tender": wsMD.Range("G25").Value = kodeTender
    wsMD.Range("F26").Value = "KK Evaluasi": wsMD.Range("G26").Value = IIf(Not wsKK Is Nothing And Trim(CStr(wsKK.Range("C6").Value)) <> "", "OK", "Belum ada")
    wsMD.Range("F27").Value = "Harga Penawaran": wsMD.Range("G27").Value = IIf(Not wsHP Is Nothing And Trim(CStr(wsHP.Range("A3").Value)) <> "", "OK", "Belum ada")
    wsMD.Range("H25").Value = "Last publish": wsMD.Range("I25").Value = IIf(publishAt = "", "Belum pernah", publishAt)
    wsMD.Range("H26").Value = "BA / HPS": wsMD.Range("I26").Value = IIf(Not wsBA Is Nothing And Trim(CStr(wsBA.Range("C7").Value)) <> "", "BA OK", "BA belum ada") & " / " & IIf(Trim(CStr(wsMD.Range("C8").Value)) <> "", "HPS OK", "HPS belum ada")
    wsMD.Range("H27").Value = "Status": wsMD.Range("I27").Value = status & " | " & publishStatus & " | " & Format(Now, "dd/mm/yyyy hh:nn")
    wsMD.Range("F25:F27,H25:H27").Font.Bold = True
    wsMD.Range("G25:G27,I25:I27").Interior.Color = RGB(255, 255, 204)
    If status = "Selesai" Then wsMD.Range("K25").Value = Now
End Sub


Private Function BacaPublishTimestamp() As String
    Dim path As String
    path = ThisWorkbook.Path & "\_workflow_publish.json"
    If Dir(path) = "" Then Exit Function
    On Error Resume Next
    BacaPublishTimestamp = ExtractVal(ReadFileBA(path), "last_published_at")
    On Error GoTo 0
End Function


' ============================================================
' MUAT & SYNC: gabungan MuatInputBA + SyncKalender dalam 1 tombol
' ============================================================
Public Sub MuatDanSync()
    MuatInputBA
    SyncKalender
End Sub


' SYNC KALENDER: isi C3/C4 dari Google Calendar
' ============================================================
Public Sub SyncKalender()
    Dim wsBA As Worksheet
    On Error GoTo ErrHandler

    On Error Resume Next
    Set wsBA = ThisWorkbook.Sheets(SHEET_BA)
    On Error GoTo ErrHandler
    If wsBA Is Nothing Then
        MsgBox "Sheet '" & SHEET_BA & "' tidak ditemukan.", vbExclamation
        Exit Sub
    End If

    ' Ambil nama_tender dari @ Master Data C3
    Dim wsMD As Worksheet
    On Error Resume Next
    Set wsMD = ThisWorkbook.Sheets("@ Master Data")
    On Error GoTo ErrHandler
    If wsMD Is Nothing Then
        MsgBox "Sheet '@ Master Data' tidak ditemukan.", vbExclamation
        Exit Sub
    End If

    Dim namaTender As String
    namaTender = Trim(CStr(wsMD.Cells(5, 3).Value))  ' C5 = nama_tender (C3=MAK, C4=kode_tender)
    If namaTender = "" Then
        MsgBox "Nama tender belum terisi di '@ Master Data' C5." & vbCrLf & _
               "Klik 'Muat Draft Paket' terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    ' Panggil Python (bisa refresh token, tidak blocking Excel)
    Dim sd As String
    sd = ModWordLink.ScriptDir_Public()
    If sd = "" Then MsgBox "Script dir tidak ditemukan.", vbExclamation: Exit Sub

    Dim pyExe As String
    pyExe = sd & "\python\python.exe"
    Dim pyScript As String
    pyScript = sd & "\sync_kalender.py"
    Dim outJson As String
    outJson = sd & "\_sync_kalender.json"

    ' Hapus output lama
    On Error Resume Next
    Kill outJson
    On Error GoTo ErrHandler

    ' Tulis nama tender ke file input (hindari masalah special char di argumen)
    Dim inpFile As String
    inpFile = sd & "\_sync_kalender_input.txt"
    Dim ado2 As Object
    Set ado2 = CreateObject("ADODB.Stream")
    ado2.Type = 2: ado2.Charset = "UTF-8": ado2.Open
    ado2.WriteText namaTender
    ado2.SaveToFile inpFile, 2
    ado2.Close

    Dim cmd As String
    cmd = """" & pyExe & """ """ & pyScript & """"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Application.StatusBar = "Sync Kalender: mencari jadwal di Google Calendar..."
    wsh.Run cmd, 0, True  ' hidden, blocking — Python selesai dulu baru Excel lanjut
    Set wsh = Nothing
    Application.StatusBar = False

    ' Baca hasil JSON
    If Dir(outJson) = "" Then
        MsgBox "sync_kalender.py tidak menghasilkan output." & vbCrLf & _
               "Pastikan token.json valid (login dulu di V19 Scheduler).", vbExclamation
        Exit Sub
    End If

    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2: ado.Charset = "UTF-8": ado.Open
    ado.LoadFromFile outJson
    Dim jsonStr As String
    jsonStr = ado.ReadText
    ado.Close

    Dim errMsg As String
    errMsg = ExtractVal(jsonStr, "error")
    If errMsg <> "" And errMsg <> "null" Then
        MsgBox "Error dari GCal: " & errMsg & vbCrLf & _
               "Coba login ulang di V19 Scheduler.", vbExclamation
        Exit Sub
    End If

    Dim tglPembukaan As String
    Dim tglPembuktian As String
    Dim keyword As String
    tglPembukaan  = ExtractVal(jsonStr, "tgl_pembukaan")
    tglPembuktian = ExtractVal(jsonStr, "tgl_pembuktian")
    keyword       = ExtractVal(jsonStr, "keyword")

    wsBA.Unprotect

    Dim pesanHasil As String
    pesanHasil = "Hasil Sync Kalender:" & vbCrLf & _
                 "Paket: " & namaTender & vbCrLf & vbCrLf

    If tglPembukaan <> "" And tglPembukaan <> "null" Then
        wsBA.Cells(ROW_TGL_PEMBUKAAN, 3).Value = CDate(tglPembukaan)
        wsBA.Cells(ROW_TGL_PEMBUKAAN, 3).NumberFormat = "dd mmmm yyyy"
        pesanHasil = pesanHasil & Chr(10) & "Pembukaan Penawaran: " & tglPembukaan
    Else
        pesanHasil = pesanHasil & Chr(10) & "Pembukaan Penawaran: tidak ditemukan"
    End If

    If tglPembuktian <> "" And tglPembuktian <> "null" Then
        wsBA.Cells(ROW_TGL_PEMBUKTIAN, 3).Value = CDate(tglPembuktian)
        wsBA.Cells(ROW_TGL_PEMBUKTIAN, 3).NumberFormat = "dd mmmm yyyy"
        pesanHasil = pesanHasil & Chr(10) & "Pembuktian/Penetapan: " & tglPembuktian
    Else
        pesanHasil = pesanHasil & Chr(10) & "Pembuktian/Penetapan: tidak ditemukan"
    End If

    If tglPembukaan = "" And tglPembuktian = "" Then
        pesanHasil = pesanHasil & vbCrLf & vbCrLf & "Keyword: """ & keyword & """"
    End If

    MsgBox pesanHasil, vbInformation, "Sync Kalender"
    Exit Sub

ErrHandler:
    Application.StatusBar = False
    MsgBox "Error SyncKalender: " & Err.Description, vbCritical
End Sub


' ============================================================
' Provision matrix peserta 1..10 + compatibility view.
' ============================================================
Private Sub EnsureInputBASlots(wsBA As Worksheet)
    Dim i As Integer, r As Integer, c As Integer
    Dim viewRows As Variant, matrixRows As Variant, labels As Variant

    wsBA.Cells(36, 2).Value = "DATABASE PESERTA (MAKSIMAL 10)"
    wsBA.Cells(37, 2).Value = "Data"
    If Trim(CStr(wsBA.Range("C5").Value)) = "" Then wsBA.Range("C5").Value = "Peserta 1"
    labels = Array("Nama Perusahaan", "NPWP", "Alamat", "Direktur / Pemilik", _
                   "Harga Penawaran", "Harga Penawaran Terkoreksi", _
                   "Personel Manajerial 1", "Personel Manajerial 2", _
                   "Alat 1", "Alat 2", "Alat 3", "Alat 4", "Alat 5", "Alat 6", _
                   "SKP", "Hasil Pembuktian Kualifikasi")
    For i = LBound(labels) To UBound(labels)
        wsBA.Cells(MATRIX_FIRST_ROW + i, 2).Value = labels(i)
    Next i
    For i = 1 To MAX_PARTICIPANTS
        c = i + 2
        wsBA.Cells(MATRIX_HEADER_ROW, c).Value = "Peserta " & i
        wsBA.Cells(MATRIX_HEADER_ROW, c).Font.Bold = True
    Next i

    wsBA.Cells(5, 6).Formula = "=IFERROR(MATCH($C$5,$C$37:$L$37,0),1)"
    On Error Resume Next
    wsBA.Range("C5").Validation.Delete
    wsBA.Range("C5").Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
        Operator:=xlBetween, Formula1:="=$C$37:$L$37"
    On Error GoTo 0

    viewRows = Array(7, 8, 9, 10, 11, 12, 13, 14, 17, 18, 19, 20, 21, 22, 32, 33)
    matrixRows = Array(38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 50, 51, 52, 53)
    For i = LBound(viewRows) To UBound(viewRows)
        r = CInt(viewRows(i))
        wsBA.Cells(r, 7).Formula = _
            "=IF(INDEX($C$" & matrixRows(i) & ":$L$" & matrixRows(i) & ",1,$F$5)=\"\",\"\",INDEX($C$" & matrixRows(i) & ":$L$" & matrixRows(i) & ",1,$F$5))"
        wsBA.Cells(r, 3).Formula = "=C$" & matrixRows(i)
        wsBA.Cells(r, 4).Formula = "=D$" & matrixRows(i)
        wsBA.Cells(r, 5).Formula = "=E$" & matrixRows(i)
        wsBA.Cells(r, 9).Formula = "=F$" & matrixRows(i)
    Next i
End Sub


' ============================================================
' Isi identitas peserta dari tabel peserta_identitas
' ============================================================
Private Sub IsiIdentitasPeserta(wsBA As Worksheet, kodeTender As String)
    Dim url As String
    url = SB_URL & "/rest/v1/peserta_identitas" & _
          "?kode_tender=eq." & kodeTender & _
          "&order=peserta_id.asc" & _
          "&limit=10"

    Dim json As String
    json = HttpGet(url)
    If json = "" Or json = "[]" Then Exit Sub

    ' Parse maksimal 10 peserta ke matrix C:L.
    Dim i As Integer
    Dim startPos As Long
    startPos = 1
    Dim col As Integer
    Dim pesertaIdx As Integer
    pesertaIdx = 1

    ' Identitas, alat/personel, dan hasil lama dibersihkan; harga 42:43
    ' dipertahankan karena sumbernya Sheet 6/evaluasi harga.
    wsBA.Range("C38:L41").ClearContents
    wsBA.Range("C44:L53").ClearContents

    Do While pesertaIdx <= MAX_PARTICIPANTS
        col = pesertaIdx + 2
        Dim objStart As Long, objEnd As Long
        objStart = InStr(startPos, json, "{")
        If objStart = 0 Then Exit Do
        objEnd = InStr(objStart, json, "}")
        If objEnd = 0 Then Exit Do

        Dim obj As String
        obj = Mid(json, objStart, objEnd - objStart + 1)

        wsBA.Cells(38, col).Value = ExtractVal(obj, "nama_perusahaan")
        wsBA.Cells(39, col).NumberFormat = "@"
        wsBA.Cells(39, col).Value = FormatNPWP(ExtractVal(obj, "npwp_raw"))
        wsBA.Cells(40, col).Value = ExtractVal(obj, "alamat")
        wsBA.Cells(41, col).Value = ExtractVal(obj, "nama_direktur")
        wsBA.Cells(44, col).Value = ExtractVal(obj, "personel_1")
        wsBA.Cells(45, col).Value = ExtractVal(obj, "personel_2")
        wsBA.Cells(46, col).Value = ExtractVal(obj, "alat_1")
        wsBA.Cells(47, col).Value = ExtractVal(obj, "alat_2")
        wsBA.Cells(48, col).Value = ExtractVal(obj, "alat_3")
        wsBA.Cells(49, col).Value = ExtractVal(obj, "alat_4")
        wsBA.Cells(50, col).Value = ExtractVal(obj, "alat_5")
        wsBA.Cells(51, col).Value = ExtractVal(obj, "alat_6")

        startPos = objEnd + 1
        pesertaIdx = pesertaIdx + 1
    Loop

    ' Harga slot di luar peserta aktif tidak boleh tertinggal dari refresh lama.
    Dim hargaClearStart As Integer
    hargaClearStart = pesertaIdx + 2
    If hargaClearStart <= 12 Then
        wsBA.Range(wsBA.Cells(42, hargaClearStart), wsBA.Cells(43, 12)).ClearContents
    End If
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
' Conflict check personil/alat via Python conflict_check.py
' Isi helper M:V (peserta 1..10) untuk baris personil & alat
' ============================================================
Private Sub IsiConflictCheck(wsBA As Worksheet, kodeTender As String)
    Dim sd As String
    sd = ScriptDirBA()
    If sd = "" Then Exit Sub

    Dim pyExe    As String: pyExe    = sd & "\python\python.exe"
    Dim pyScript As String: pyScript = sd & "\conflict_check.py"
    Dim outJson  As String: outJson  = sd & "\_conflict_check.json"

    If Not FileExistsBA(pyExe) Or Not FileExistsBA(pyScript) Then Exit Sub

    ' Hapus output lama
    On Error Resume Next
    Kill outJson
    On Error GoTo 0

    ' Panggil Python: conflict_check.py <kode_tender>
    Dim cmd As String
    cmd = """" & pyExe & """ """ & pyScript & """ " & kodeTender

    Application.StatusBar = "Conflict Check: memeriksa riwayat personil/alat..."
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True
    Application.StatusBar = False

    If Not FileExistsBA(outJson) Then Exit Sub

    ' Baca JSON output
    Dim result As String
    result = ReadFileBA(outJson)

    If ExtractVal(result, "error") <> "" And ExtractVal(result, "error") <> "null" Then
        Exit Sub  ' silent — conflict check opsional
    End If

    ' Conflict helper dipindah ke M:V agar tidak menimpa G (view peserta
    ' terpilih) atau matrix kanonik C:L.
    Dim r As Integer
    For r = ROW_PERSONEL_1 To ROW_ALAT_6
        wsBA.Range(wsBA.Cells(r, 13), wsBA.Cells(r, 12 + MAX_PARTICIPANTS)).ClearContents
        wsBA.Range(wsBA.Cells(r, 13), wsBA.Cells(r, 12 + MAX_PARTICIPANTS)).Interior.ColorIndex = -4142
    Next r

    ' Personil/alat peserta 1..10 → helper M:V.
    Dim pesertaCols(1 To MAX_PARTICIPANTS) As Integer
    Dim pIdx As Integer
    For pIdx = 1 To MAX_PARTICIPANTS
        pesertaCols(pIdx) = 12 + pIdx
    Next pIdx

    ' Ekstrak status per peserta per field — format JSON flat per peserta_id
    ' Karena peserta_id dari Supabase urut asc, kita map index ke kolom helper.
    For pIdx = 1 To MAX_PARTICIPANTS
        ' Cari blok peserta ke-pIdx di "personil"
        Dim p1Key As String: p1Key = "personil_1"
        Dim p2Key As String: p2Key = "personil_2"

        ' Ekstrak status personil_1 peserta ke-pIdx
        Dim st1 As String: st1 = GetConflictPesan(result, "personil", pIdx, "personil_1")
        Dim st2 As String: st2 = GetConflictPesan(result, "personil", pIdx, "personil_2")

        If st1 <> "" Then
            wsBA.Cells(ROW_PERSONEL_1, pesertaCols(pIdx)).Value = st1
            TerapkanWarna wsBA.Cells(ROW_PERSONEL_1, pesertaCols(pIdx)), GetConflictStatus(result, "personil", pIdx, "personil_1")
        End If
        If st2 <> "" Then
            wsBA.Cells(ROW_PERSONEL_2, pesertaCols(pIdx)).Value = st2
            TerapkanWarna wsBA.Cells(ROW_PERSONEL_2, pesertaCols(pIdx)), GetConflictStatus(result, "personil", pIdx, "personil_2")
        End If

        ' Alat: alat_1..6 → ROW_ALAT_1..6
        Dim aIdx As Integer
        Dim aRows(1 To 6) As Integer
        aRows(1) = ROW_ALAT_1: aRows(2) = ROW_ALAT_2: aRows(3) = ROW_ALAT_3
        aRows(4) = ROW_ALAT_4: aRows(5) = ROW_ALAT_5: aRows(6) = ROW_ALAT_6
        For aIdx = 1 To 6
            Dim aKey As String: aKey = "alat_" & aIdx
            Dim ast As String: ast = GetConflictPesan(result, "alat", pIdx, aKey)
            If ast <> "" Then
                wsBA.Cells(aRows(aIdx), pesertaCols(pIdx)).Value = ast
                TerapkanWarna wsBA.Cells(aRows(aIdx), pesertaCols(pIdx)), GetConflictStatus(result, "alat", pIdx, aKey)
            End If
        Next aIdx
    Next pIdx
End Sub


' Warna sel berdasarkan status
Private Sub TerapkanWarna(cel As Range, status As String)
    Select Case status
        Case "warning", "proses"
            cel.Interior.Color = RGB(255, 200, 0)   ' Kuning peringatan
        Case "ok"
            cel.Interior.Color = RGB(198, 239, 206)  ' Hijau muda
        Case Else
            cel.Interior.ColorIndex = -4142
    End Select
End Sub


' Ekstrak "pesan" dari nested JSON: result["personil"][peserta_idx]["key"]["pesan"]
' JSON flat — kita parse manual dengan InStr bertingkat
Private Function GetConflictPesan(json As String, tipe As String, pesertaIdx As Integer, key As String) As String
    GetConflictPesan = GetConflictField(json, tipe, pesertaIdx, key, "pesan")
End Function

Private Function GetConflictStatus(json As String, tipe As String, pesertaIdx As Integer, key As String) As String
    GetConflictStatus = GetConflictField(json, tipe, pesertaIdx, key, "status")
End Function

Private Function GetConflictField(json As String, tipe As String, pesertaIdx As Integer, itemKey As String, field As String) As String
    GetConflictField = ""

    ' Cari blok tipe ("personil" atau "alat")
    Dim pTipe As Long: pTipe = InStr(json, """" & tipe & """")
    If pTipe = 0 Then Exit Function

    ' Cari objek peserta ke-pesertaIdx (skip pesertaIdx-1 buka kurung kurawal dalam blok tipe)
    Dim braceStart As Long: braceStart = InStr(pTipe, json, "{")
    If braceStart = 0 Then Exit Function

    ' Masuk ke dalam blok tipe — cari peserta ke-pesertaIdx (key = peserta_id string)
    ' Format: "personil": { "pid1": { "personil_1": {...} }, "pid2": {...} }
    ' Kita skip pesertaIdx-1 pasang key:{ untuk sampai ke peserta yang dimaksud
    Dim depth As Integer: depth = 0
    Dim pos As Long: pos = braceStart
    Dim pesertaCount As Integer: pesertaCount = 0
    Dim pesertaBlokStart As Long: pesertaBlokStart = 0

    ' Masuk brace pertama (blok tipe)
    depth = 1
    pos = braceStart + 1

    Do While pos <= Len(json) And depth > 0
        Dim ch As String: ch = Mid(json, pos, 1)
        If ch = "{" Then
            depth = depth + 1
            If depth = 2 Then
                ' Ini buka blok peserta baru
                pesertaCount = pesertaCount + 1
                If pesertaCount = pesertaIdx Then
                    pesertaBlokStart = pos
                    Exit Do
                End If
            End If
        ElseIf ch = "}" Then
            depth = depth - 1
        End If
        pos = pos + 1
    Loop

    If pesertaBlokStart = 0 Then Exit Function

    ' Cari itemKey di dalam blok peserta
    Dim pItem As Long: pItem = InStr(pesertaBlokStart, json, """" & itemKey & """")
    If pItem = 0 Then Exit Function

    ' Cari blok { setelah itemKey
    Dim pBrace As Long: pBrace = InStr(pItem, json, "{")
    If pBrace = 0 Then Exit Function

    ' Ekstrak blok item
    Dim d2 As Integer: d2 = 0
    Dim p2 As Long: p2 = pBrace
    Dim itemBlok As String
    Do While p2 <= Len(json)
        Dim c2 As String: c2 = Mid(json, p2, 1)
        If c2 = "{" Then d2 = d2 + 1
        If c2 = "}" Then
            d2 = d2 - 1
            If d2 = 0 Then
                itemBlok = Mid(json, pBrace, p2 - pBrace + 1)
                Exit Do
            End If
        End If
        p2 = p2 + 1
    Loop

    If itemBlok = "" Then Exit Function
    GetConflictField = ExtractVal(itemBlok, field)
End Function


' Helper: cari script dir (WPy64-313110)
Private Function ScriptDirBA() As String
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pokjaRoot As String
    pokjaRoot = fso.GetParentFolderName(ThisWorkbook.Path)
    Dim candidate As String
    candidate = pokjaRoot & "\V19_Scheduler\WPy64-313110"
    If Dir(candidate, vbDirectory) <> "" Then
        ScriptDirBA = candidate
    End If
End Function

Private Function FileExistsBA(path As String) As Boolean
    On Error Resume Next
    FileExistsBA = (Dir(path) <> "")
    On Error GoTo 0
End Function

Private Function ReadFileBA(path As String) As String
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2: ado.Charset = "UTF-8": ado.Open
    ado.LoadFromFile path
    ReadFileBA = ado.ReadText
    ado.Close
End Function


' ============================================================
' Isi tanggal dari Google Calendar via V19 token
' Cari event berdasarkan kata kunci di title
' ============================================================
Private Sub IsiTanggalDariGCal_UNUSED(wsBA As Worksheet, namaTender As String)
    ' Format event GCal V19: "{Tahap} - {Nama_Paket}"
    ' Cari berdasarkan kata kunci pertama nama_tender (agar match partial)
    Dim scriptDir As String
    scriptDir = ModWordLink.ScriptDir_Public()
    If scriptDir = "" Then
        MsgBox "Script dir tidak ditemukan. Pastikan ModWordLink ter-inject.", vbExclamation
        Exit Sub
    End If

    Dim tokenPath As String
    tokenPath = scriptDir & "\token.json"
    If Dir(tokenPath) = "" Then
        MsgBox "token.json tidak ditemukan di: " & tokenPath & vbCrLf & _
               "Login Google Calendar dulu di V19 Scheduler.", vbExclamation
        Exit Sub
    End If

    ' Baca token via ADODB (UTF-8 safe)
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2: ado.Charset = "UTF-8": ado.Open
    ado.LoadFromFile tokenPath
    Dim tokenJson As String
    tokenJson = ado.ReadText
    ado.Close

    Dim accessToken As String
    accessToken = ExtractVal(tokenJson, "token")
    If accessToken = "" Then accessToken = ExtractVal(tokenJson, "access_token")
    If accessToken = "" Then
        MsgBox "Access token tidak valid. Login ulang di V19 Scheduler.", vbExclamation
        Exit Sub
    End If

    ' Gunakan 3 kata pertama nama_tender sebagai query (lebih presisi)
    Dim kataPertama As String
    Dim parts() As String
    parts = Split(namaTender, " ")
    Dim nKata As Integer: nKata = IIf(UBound(parts) >= 3, 3, UBound(parts) + 1)
    Dim k As Integer
    kataPertama = ""
    For k = 0 To nKata - 1
        If kataPertama <> "" Then kataPertama = kataPertama & " "
        kataPertama = kataPertama & parts(k)
    Next k

    ' Cakup seluruh tahun 2025-2026 agar paket selesai pun ketemu
    Dim timeMin As String, timeMax As String
    timeMin = "2025-01-01T00:00:00Z"
    timeMax = "2026-12-31T00:00:00Z"

    ' Encode spasi sebagai %20 untuk URL query
    Dim qParam As String
    qParam = Replace(kataPertama, " ", "%20")

    Dim gcUrl As String
    gcUrl = "https://www.googleapis.com/calendar/v3/calendars/primary/events" & _
            "?maxResults=100" & _
            "&singleEvents=true" & _
            "&orderBy=startTime" & _
            "&timeMin=" & timeMin & _
            "&timeMax=" & timeMax & _
            "&q=" & qParam

    Dim gcJson As String
    gcJson = HttpGetWithToken(gcUrl, accessToken)
    If gcJson = "" Then
        MsgBox "Gagal mengambil data dari Google Calendar." & vbCrLf & _
               "Cek koneksi internet atau login ulang di V19 Scheduler.", vbExclamation
        Exit Sub
    End If

    ' Parse event: cari "Pembukaan" dan "Pembuktian"/"Penetapan"/"Klarifikasi"
    Dim pos As Long
    pos = 1
    Dim tglPembukaan As String, tglPembuktian As String
    tglPembukaan = ""
    tglPembuktian = ""

    Do
        Dim itemStart As Long
        itemStart = InStr(pos, gcJson, """summary""")
        If itemStart = 0 Then Exit Do

        Dim sumStart As Long, sumEnd As Long
        sumStart = InStr(itemStart, gcJson, ":""") + 2
        sumEnd = InStr(sumStart, gcJson, """")
        Dim summary As String
        summary = Mid(gcJson, sumStart, sumEnd - sumStart)

        ' Cari tanggal mulai event ini
        Dim dtStart As Long
        dtStart = InStr(itemStart, gcJson, """start""")
        Dim tgl As String
        tgl = ""
        If dtStart > 0 Then
            Dim dtPos As Long
            dtPos = InStr(dtStart, gcJson, """date")
            If dtPos > 0 And dtPos < itemStart + 500 Then
                Dim valS As Long, valE As Long
                valS = InStr(dtPos, gcJson, ":""") + 2
                valE = InStr(valS, gcJson, """")
                tgl = Left(Mid(gcJson, valS, valE - valS), 10)
            End If
        End If

        Dim sumLow As String
        sumLow = LCase(summary)

        If tglPembukaan = "" And InStr(sumLow, "pembukaan") > 0 Then
            tglPembukaan = tgl
        End If
        If tglPembuktian = "" And (InStr(sumLow, "pembuktian") > 0 Or _
                                    InStr(sumLow, "penetapan") > 0 Or _
                                    InStr(sumLow, "klarifikasi") > 0) Then
            tglPembuktian = tgl
        End If

        pos = sumEnd + 1
        If tglPembukaan <> "" And tglPembuktian <> "" Then Exit Do
    Loop

    ' Tulis ke sheet dengan feedback
    Dim pesanHasil As String
    pesanHasil = "Hasil Sync Kalender untuk:" & vbCrLf & namaTender & vbCrLf & vbCrLf

    If tglPembukaan <> "" Then
        wsBA.Cells(ROW_TGL_PEMBUKAAN, 3).Value = CDate(tglPembukaan)
        wsBA.Cells(ROW_TGL_PEMBUKAAN, 3).NumberFormat = "dd mmmm yyyy"
        pesanHasil = pesanHasil & "✔ Pembukaan Penawaran: " & tglPembukaan & vbCrLf
    Else
        pesanHasil = pesanHasil & "✘ Pembukaan Penawaran: tidak ditemukan" & vbCrLf
    End If

    If tglPembuktian <> "" Then
        wsBA.Cells(ROW_TGL_PEMBUKTIAN, 3).Value = CDate(tglPembuktian)
        wsBA.Cells(ROW_TGL_PEMBUKTIAN, 3).NumberFormat = "dd mmmm yyyy"
        pesanHasil = pesanHasil & "✔ Pembuktian/Penetapan: " & tglPembuktian & vbCrLf
    Else
        pesanHasil = pesanHasil & "✘ Pembuktian/Penetapan: tidak ditemukan" & vbCrLf
    End If

    If tglPembukaan = "" And tglPembuktian = "" Then
        pesanHasil = pesanHasil & vbCrLf & "Keyword pencarian: """ & kataPertama & """"
    End If

    MsgBox pesanHasil, vbInformation, "Sync Kalender"
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
