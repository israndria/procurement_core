Attribute VB_Name = "ModDraftPaket"
' ============================================================
' ModDraftPaket - Load data Draft Paket dari Supabase
' ============================================================
' Flow:
'   1. MuatDraftPaket()   → GET Supabase REST → isi Named Range "DaftarDraftPaket"
'                           → populate dropdown di E2 sheet "1. Input Data"
'   2. PilihDraftPaket()  → dipanggil dari Worksheet_Change saat E2 berubah
'                           → autofill semua field ke cell yang sesuai

' Konfigurasi Supabase
Private Const SB_URL As String = "%%SUPABASE_URL%%"
Private Const SB_KEY As String = "%%SUPABASE_KEY%%"
Private Const SB_TABLE As String = "draft_paket"

' Sheet & Cell target
Private Const SHEET_INPUT As String = "1. Input Data"
Private Const CELL_SELECTOR As String = "F1"   ' Dropdown pilih paket (di @ Master Data)

' Kolom yang di-fetch (efisien, tidak perlu semua)
Private Const SB_SELECT As String = "kode_tender,nama_tender,mak,kode_rup,nilai_pagu,nilai_hps,kode_pokja,nomor_pp,nomor_surat_dinas,nama_dinas,nama_ppk,jangka_waktu,sumber_anggaran,anggota_1,anggota_2,anggota_3,bidang,sbu_baru,sbu_lama,nip_ppk,sk_ppk"

' ── @ Master Data row constants (fixed layout) ──
Private Const MD_SHEET As String = "@ Master Data"
' INPUT DATA section (row 2 = header)
Private Const MD_E3 As Integer = 3
Private Const MD_E5 As Integer = 4
Private Const MD_E6 As Integer = 5
Private Const MD_E8 As Integer = 6
Private Const MD_E10 As Integer = 7
Private Const MD_E11 As Integer = 8
Private Const MD_E12 As Integer = 9
Private Const MD_E13 As Integer = 10
Private Const MD_E14 As Integer = 11
Private Const MD_E15 As Integer = 12
Private Const MD_E16 As Integer = 13
Private Const MD_E17 As Integer = 14
Private Const MD_E19 As Integer = 15
Private Const MD_E20 As Integer = 16
Private Const MD_E21 As Integer = 17
Private Const MD_E22 As Integer = 18
Private Const MD_E23 As Integer = 19
Private Const MD_E24 As Integer = 20
Private Const MD_E32 As Integer = 21
Private Const MD_E33 As Integer = 22
' REVIU section (row 24 = header)
Private Const MD_R_E2 As Integer = 25
Private Const MD_R_E6 As Integer = 26
Private Const MD_R_E7 As Integer = 27
Private Const MD_R_E9 As Integer = 28
Private Const MD_R_E10 As Integer = 29
Private Const MD_R_E11 As Integer = 30
Private Const MD_R_E12 As Integer = 31
Private Const MD_R_E13 As Integer = 32
Private Const MD_R_E14 As Integer = 33
Private Const MD_R_E15 As Integer = 34
Private Const MD_R_E16 As Integer = 35
Private Const MD_R_E17 As Integer = 36
Private Const MD_R_E18 As Integer = 37
Private Const MD_R_E19 As Integer = 38
Private Const MD_R_E20 As Integer = 39
Private Const MD_R_E21 As Integer = 40
Private Const MD_R_E22 As Integer = 41
Private Const MD_R_E23 As Integer = 42
Private Const MD_R_E24 As Integer = 43
Private Const MD_R_E25 As Integer = 44
Private Const MD_R_E26 As Integer = 45
Private Const MD_R_E27 As Integer = 46
Private Const MD_R_E28 As Integer = 47
Private Const MD_R_E29 As Integer = 48
Private Const MD_R_E30 As Integer = 49
Private Const MD_R_E31 As Integer = 50
Private Const MD_R_E32 As Integer = 51
Private Const MD_R_E33 As Integer = 52
Private Const MD_R_E34 As Integer = 53
' DOKPIL section (row 55 = header)
Private Const MD_D_E6 As Integer = 56
Private Const MD_D_E7 As Integer = 57
Private Const MD_D_E8 As Integer = 58
Private Const MD_D_E9 As Integer = 59
Private Const MD_D_E10 As Integer = 60
Private Const MD_D_E11 As Integer = 61
Private Const MD_D_E12 As Integer = 62
Private Const MD_D_E13 As Integer = 63
Private Const MD_D_E14 As Integer = 64
Private Const MD_D_E15 As Integer = 65
Private Const MD_D_E16 As Integer = 66

' Cache data in-memory (Collection of Dictionary-like arrays)
Private m_DataCache As Collection
Private m_LastLoad As Date


' ============================================================
' FUNGSI UTAMA: Load dari Supabase, buat dropdown
' ============================================================
Public Sub MuatDraftPaket(Optional bFromOpen As Boolean = False)
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    ' Dropdown sekarang di @ Master Data
    Set ws = ThisWorkbook.Sheets(MD_SHEET)
    ' Pastikan sheet ada
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        ws.Name = MD_SHEET
    End If

    ' Fetch JSON dari Supabase
    Dim json As String
    json = FetchSupabase()
    If json = "" Then
        MsgBox "Gagal mengambil data dari Supabase.", vbExclamation
        Exit Sub
    End If

    ' Parse JSON → Collection
    Set m_DataCache = ParseDraftJSON(json)
    m_LastLoad = Now

    If m_DataCache.Count = 0 Then
        MsgBox "Tidak ada data Draft Paket di database." & vbCrLf & _
               "Jalankan 'Serap Inbox' di Asisten Pokja terlebih dahulu.", vbInformation
        Exit Sub
    End If

    ' Tulis label ke sheet tersembunyi "_DraftPaketList" kolom A
    Dim wsHidden As Worksheet
    On Error Resume Next
    Set wsHidden = ThisWorkbook.Sheets("_DraftPaketList")
    On Error GoTo ErrHandler
    If wsHidden Is Nothing Then
        Set wsHidden = ThisWorkbook.Sheets.Add
        wsHidden.Name = "_DraftPaketList"
    End If
    wsHidden.Visible = xlSheetVeryHidden

    ' Bersihkan isi lama
    wsHidden.Columns(1).ClearContents

    Dim i As Integer
    For i = 1 To m_DataCache.Count
        Dim item As Variant
        item = m_DataCache(i)
        Dim pokja As String: pokja = CStr(item(0))  ' kode_pokja
        Dim nama As String:  nama  = CStr(item(1))  ' nama_tender
        Dim dinas As String: dinas = CStr(item(9))  ' nama_dinas
        Do While InStr(nama, "  ") > 0: nama = Join(Split(nama, "  "), " "): Loop
        nama = Trim(nama)
        ' Ambil tahun dari nomor_pp (4 digit terakhir)
        Dim npp As String: npp = Trim(CStr(item(7)))
        Dim thn As String: thn = ""
        If Len(npp) >= 4 Then thn = Right(npp, 4)
        Dim label As String
        label = pokja & " - " & Left(nama, 60)
        If thn <> "" Then label = label & " - " & thn
        wsHidden.Cells(i, 1).Value = label
    Next i

    ' Buat Named Range "DaftarDraftPaket" merujuk ke list di sheet tersembunyi
    Dim rngList As Range
    Set rngList = wsHidden.Range(wsHidden.Cells(1, 1), wsHidden.Cells(m_DataCache.Count, 1))
    On Error Resume Next
    ThisWorkbook.Names("DaftarDraftPaket").Delete
    On Error GoTo ErrHandler
    ThisWorkbook.Names.Add Name:="DaftarDraftPaket", RefersTo:=rngList

    ' Validation merujuk ke Named Range (tidak ada limit 8192)
    With ws.Range(CELL_SELECTOR).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=DaftarDraftPaket"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With

    ' Simpan nilai dropdown yang sudah dipilih sebelum rebuild
    Dim savedLabel As String
    Dim selCell As Range: Set selCell = ws.Range(CELL_SELECTOR)
    If selCell.MergeCells Then
        savedLabel = Trim(CStr(selCell.MergeArea.Cells(1, 1).Value))
    Else
        savedLabel = Trim(CStr(selCell.Value))
    End If

    ' Restore pilihan dropdown tanpa trigger event
    If savedLabel <> "" Then
        Application.EnableEvents = False
        ws.Range(CELL_SELECTOR).Value = savedLabel
        Application.EnableEvents = True
    End If

    MsgBox m_DataCache.Count & " paket berhasil dimuat." & vbCrLf & _
           "Pilih paket di cell " & CELL_SELECTOR & ", lalu klik 'Parse Draft' untuk mengisi data.", _
           vbInformation, "Draft Paket Dimuat"
    Exit Sub

ErrHandler:
    MsgBox "Error MuatDraftPaket: " & Err.Description, vbCritical
End Sub


' ============================================================
' PARSE DRAFT: Dipanggil tombol "Parse Draft" — baca pilihan dropdown lalu parse
' ============================================================
Public Sub ParseDraftTerpilih()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(MD_SHEET)
    Dim selCell As Range: Set selCell = ws.Range(CELL_SELECTOR)
    Dim val As String
    If selCell.MergeCells Then
        val = Trim(CStr(selCell.MergeArea.Cells(1, 1).Value))
    Else
        val = Trim(CStr(selCell.Value))
    End If
    If val = "" Then
        MsgBox "Pilih paket dari dropdown " & CELL_SELECTOR & " terlebih dahulu.", vbExclamation, "Belum Ada Pilihan"
        Exit Sub
    End If
    ' Auto-load jika cache kosong (misal setelah reopen workbook)
    If m_DataCache Is Nothing Then
        MuatDraftPaket
    ElseIf m_DataCache.Count = 0 Then
        MuatDraftPaket
    End If
    ' Cek ulang setelah load
    If m_DataCache Is Nothing Then Exit Sub
    If m_DataCache.Count = 0 Then Exit Sub
    PilihDraftPaket val
End Sub


' ============================================================
' AUTOFILL: Dipanggil dari Worksheet_Change
' ============================================================
Public Sub PilihDraftPaket(selectedLabel As String)
    If m_DataCache Is Nothing Then Exit Sub
    If m_DataCache.Count = 0 Then Exit Sub

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(SHEET_INPUT)

    ' Cari item yang labelnya cocok (bandingkan prefix kode_pokja + potongan nama)
    Dim i As Integer
    For i = 1 To m_DataCache.Count
        Dim item As Variant
        item = m_DataCache(i)
        Dim pokja As String: pokja = CStr(item(0))
        Dim nama As String:  nama  = CStr(item(1))
        Dim dinas As String: dinas = CStr(item(9))

        Do While InStr(nama, "  ") > 0: nama = Join(Split(nama, "  "), " "): Loop
        nama = Trim(nama)
        Dim npp2 As String: npp2 = Trim(CStr(item(7)))
        Dim thn2 As String: thn2 = ""
        If Len(npp2) >= 4 Then thn2 = Right(npp2, 4)
        Dim label As String
        label = pokja & " - " & Left(nama, 60)
        If thn2 <> "" Then label = label & " - " & thn2

        If label = selectedLabel Then
            ' Mapping: tulis ke @ Master Data kolom C
            Dim wsMD As Worksheet
            On Error Resume Next
            Set wsMD = ThisWorkbook.Sheets(MD_SHEET)
            On Error GoTo 0
            If wsMD Is Nothing Then
                Set wsMD = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
                wsMD.Name = MD_SHEET
            End If
            With wsMD
                .Cells(MD_E3, 3).Value  = CStr(item(2))   ' MAK
                .Cells(MD_E5, 3).Value  = CStr(item(6))   ' Kode Tender
                .Cells(MD_E6, 3).Value  = CStr(item(1))   ' Nama Tender
                .Cells(MD_E8, 3).Value  = CStr(item(3))   ' Kode RUP
                .Cells(MD_E10, 3).Value = CStr(item(4))   ' Nilai Pagu
                .Cells(MD_E11, 3).Value = CStr(item(5))   ' Nilai HPS
                .Cells(MD_E12, 3).Value = CStr(item(8))   ' Nomor Surat Dinas
                .Cells(MD_E13, 3).Value = CStr(item(7))   ' Nomor PP
                .Cells(MD_E14, 3).NumberFormat = "@"
                .Cells(MD_E14, 3).Value = CStr(item(0))   ' Kode Pokja
                ' Jangka Waktu
                Dim jw As String: jw = Trim(CStr(item(11)))
                Dim spPos As Long: spPos = InStr(jw, " ")
                If spPos > 0 Then jw = Left(jw, spPos - 1)
                If IsNumeric(jw) Then
                    .Cells(MD_E15, 3).Value = CLng(jw)
                Else
                    .Cells(MD_E15, 3).Value = CStr(item(11))
                End If
                ' SKPD/OPD
                .Cells(MD_E17, 3).Value = CStr(item(9))
            End With
            ' Helper cells tetap di sheet asli
            ws.Range("F17").Value = CStr(item(9))
            ' Nama PPK
            Dim ppkNama As String: ppkNama = Trim(CStr(item(10)))
            Dim komaPos As Long: komaPos = InStr(ppkNama, ",")
            Dim ppkNamaBersih As String
            If komaPos > 0 Then
                ppkNamaBersih = Trim(Left(ppkNama, komaPos - 1))
                wsMD.Cells(MD_E19, 3).Value = ppkNamaBersih
                ws.Range("F19").Value = Trim(Mid(ppkNama, komaPos))
            Else
                ppkNamaBersih = ppkNama
                wsMD.Cells(MD_E19, 3).Value = ppkNama
            End If
            ' NIP PPK + Nomor SK: dari Supabase dulu, fallback ke sheet lokal
            Dim nipPPK As String, skPPK As String, gelarPPK As String
            nipPPK = Trim(CStr(item(19)))
            skPPK  = Trim(CStr(item(20)))
            If nipPPK = "" Or skPPK = "" Then
                Dim nipLookup As String, skLookup As String, gelarLookup As String
                LookupPPK ppkNamaBersih, nipLookup, skLookup, gelarLookup
                If nipPPK = "" Then nipPPK = nipLookup
                If skPPK  = "" Then skPPK  = skLookup
                If gelarPPK = "" Then gelarPPK = gelarLookup
            End If
            If nipPPK <> "" Then
                wsMD.Cells(MD_E20, 3).NumberFormat = "@"
                wsMD.Cells(MD_E20, 3).Value = nipPPK
            End If
            If skPPK <> "" Then wsMD.Cells(MD_E21, 3).Value = skPPK
            If gelarPPK <> "" And komaPos = 0 Then ws.Range("F19").Value = gelarPPK
            ' Sumber Anggaran
            wsMD.Cells(MD_E33, 3).Value = CStr(item(12))
            ' Isi anggota pokja
            IsiAnggotaPokjaToMaster wsMD, CStr(item(13)), CStr(item(14)), CStr(item(15))

            ' ── Parse PDF → isi database_reviu + database_dokpil ──────────────
            Dim kodeTender As String: kodeTender = CStr(item(6))
            Dim kodePokja2 As String: kodePokja2 = CStr(item(0))
            Dim bidang2 As String: bidang2 = CStr(item(16))
            Dim namaTender2 As String: namaTender2 = CStr(item(1))
            ParsaDanIsiDariPDF kodeTender, kodePokja2, bidang2, namaTender2

            ' ── SBU → @ Master Data DATABASE REVIU: E6=SBU Baru (baris 26), E7=SBU Lama (baris 27) ──
            Dim sbuBaru As String: sbuBaru = CStr(item(17))
            Dim sbuLama As String: sbuLama = CStr(item(18))
            If sbuBaru <> "" Then
                wsMD.Cells(MD_R_E6, 3).Value = sbuBaru  ' SBU Baru (KBLI 2020) → baris 26
                wsMD.Cells(MD_R_E6, 2).Value = "Terisi (Supabase)"
            End If
            If sbuLama <> "" Then
                wsMD.Cells(MD_R_E7, 3).Value = sbuLama  ' SBU Lama (KBLI 2015) → baris 27
                wsMD.Cells(MD_R_E7, 2).Value = "Terisi (Supabase)"
            End If
            ' Bersihkan baris 39 (E20) jika isinya adalah kode SBU yang salah tempat
            Dim c39 As String: c39 = Trim(CStr(wsMD.Cells(MD_R_E20, 3).Value))
            If Len(c39) = 5 And (Left(c39, 2) = "BS" Or Left(c39, 2) = "SI" Or Left(c39, 2) = "BG") Then
                wsMD.Cells(MD_R_E20, 3).Value = ""
            End If
            ' Override baris SBU di _HasilParse agar status sinkron dengan Supabase
            OverrideSBUdiHasilParse sbuBaru, sbuLama

            ' ── Muat HPS dari Supabase → isi sheet "5. HPS" ───────────────────
            MuatHPS kodeTender

            ' ── Diff Highlight: bandingkan vs snapshot Supabase ───────────────
            ModSyncDraft.DiffHighlight kodeTender
            Exit For
        End If
    Next i
End Sub


' ============================================================
' PARSE PDF: Cari Draft_Pokja_XXX.pdf → panggil Python → isi sheet
' ============================================================
Public Sub ParsaDanIsiDariPDF(kodeTender As String, kodePokja As String, Optional bidang As String = "", Optional namaTender As String = "")
    Dim folderWorkbook As String
    folderWorkbook = ThisWorkbook.Path

    ' Cari file Draft_Pokja_*.pdf di folder workbook
    Dim pathPDF As String
    pathPDF = CariDraftPDF(folderWorkbook, kodePokja)
    If pathPDF = "" Then
        ' Coba subfolder satu level di atas
        pathPDF = CariDraftPDF(folderWorkbook & "\..", kodePokja)
    End If

    If pathPDF = "" Then
        ' Tidak ada PDF — tetap tampilkan _HasilParse dengan info PDF tidak ditemukan
        TampilkanHasilParseNoPDF folderWorkbook
        Exit Sub
    End If

    ' Cek apakah nama PDF cocok dengan kode pokja yang dipilih
    Dim namaFilePDF As String
    namaFilePDF = Mid(pathPDF, InStrRev(pathPDF, "\") + 1)
    Dim kodePokjaFormatted As String
    On Error Resume Next
    kodePokjaFormatted = Format(CLng(kodePokja), "000")
    On Error GoTo 0
    Dim pdfCocok As Boolean
    pdfCocok = (InStr(LCase(namaFilePDF), LCase(kodePokja)) > 0) Or _
               (kodePokjaFormatted <> "" And InStr(namaFilePDF, kodePokjaFormatted) > 0)
    If Not pdfCocok Then
        Dim konfirmasi As Integer
        konfirmasi = MsgBox("Paket yang dipilih: " & kodePokja & vbCrLf & _
                            "PDF yang ditemukan: " & namaFilePDF & vbCrLf & vbCrLf & _
                            "PDF ini kemungkinan bukan untuk paket yang dipilih." & vbCrLf & _
                            "Lanjutkan parse dengan PDF ini?", _
                            vbQuestion + vbYesNo, "Konfirmasi PDF")
        If konfirmasi = vbNo Then Exit Sub
    End If

    ' Panggil Python parse_reviu.py
    Dim pythonExe As String
    pythonExe = folderWorkbook & "\..\..\V19_Scheduler\WPy64-313110\python\python.exe"
    If Not FileExists(pythonExe) Then
        ' Coba path relatif jika workbook ada di subfolder paket
        pythonExe = folderWorkbook & "\..\V19_Scheduler\WPy64-313110\python\python.exe"
    End If

    Dim scriptPY As String
    scriptPY = folderWorkbook & "\parse_reviu.py"
    If Not FileExists(scriptPY) Then
        scriptPY = folderWorkbook & "\..\V19_Scheduler\WPy64-313110\parse_reviu.py"
    End If

    ' Tentukan python exe dari BASE_DIR standar
    Dim baseDir As String
    baseDir = CariBaseDir(folderWorkbook)
    pythonExe = baseDir & "\python\python.exe"
    scriptPY  = baseDir & "\parse_reviu.py"

    If Not FileExists(pythonExe) Or Not FileExists(scriptPY) Then
        MsgBox "Python atau parse_reviu.py tidak ditemukan." & vbCrLf & _
               "Python: " & pythonExe & vbCrLf & "Script: " & scriptPY, vbExclamation
        Exit Sub
    End If

    Dim outJSON As String
    outJSON = folderWorkbook & "\_parse_reviu.json"

    ' Hapus JSON lama
    If FileExists(outJSON) Then Kill outJSON

    ' Tulis argumen ke file temp pakai forward slash
    ' (backslash di dalam string Python bisa jadi escape, misal \1 → karakter kontrol)
    Dim argFile As String
    argFile = Environ("TEMP") & "\_parse_reviu_args.txt"
    Dim fArg As Integer: fArg = FreeFile
    Open argFile For Output As #fArg
    Print #fArg, Replace(pathPDF, "\", "/")
    Print #fArg, Replace(folderWorkbook, "\", "/")
    Print #fArg, bidang
    Print #fArg, namaTender
    Close #fArg

    ' Gunakan junction C:\pokja2026 agar path tidak mengandung @
    ' (@ menyebabkan cmd.exe gagal parse; junction dibuat sekali saat pertama kali)
    Dim junctionBase As String
    junctionBase = "C:\pokja2026"
    If Dir(junctionBase, vbDirectory) = "" Then
        CreateObject("WScript.Shell").Run "cmd /c mklink /J """ & junctionBase & """ ""D:\Dokumen\@ POKJA 2026""", 0, True
    End If
    Dim pyExeJunction As String: pyExeJunction = junctionBase & "\V19_Scheduler\WPy64-313110\python\python.exe"
    Dim pyScriptJunction As String: pyScriptJunction = junctionBase & "\V19_Scheduler\WPy64-313110\parse_reviu.py"

    ' Jalankan via PowerShell (bukan cmd.exe) agar path dengan spasi dan angka aman
    Dim psCmd As String
    psCmd = "powershell.exe -NoProfile -WindowStyle Hidden -Command """ & _
            "& '" & pyExeJunction & "' '" & pyScriptJunction & "' --argfile '" & argFile & "'"""

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim ret As Long
    ret = wsh.Run(psCmd, 0, True)

    If Not FileExists(outJSON) Then
        MsgBox "Parse PDF gagal. Cek apakah pdfplumber terinstall." & vbCrLf & "Kode return: " & ret, vbExclamation
        Exit Sub
    End If

    ' Baca JSON
    Dim jsonTeks As String
    jsonTeks = BacaFile(outJSON)

    ' Cek error
    Dim errMsg As String
    errMsg = ExtractJSONVal(jsonTeks, "error")
    If errMsg <> "" Then
        MsgBox "Error saat parsing PDF: " & errMsg, vbExclamation
        Exit Sub
    End If

    ' Isi input_data (1. Input Data)
    IsiInputDataDariPDF jsonTeks

    ' Isi database_reviu
    IsiDatabaseReviu jsonTeks

    ' Isi database_dokpil
    IsiDatabaseDokpil jsonTeks

    ' Tampilkan sheet _HasilParse (hanya 1x MsgBox)
    TampilkanHasilParse jsonTeks, folderWorkbook
End Sub


' ============================================================
' ISI 1. Input Data dari JSON (hasil parse PDF)
' ============================================================
Private Sub IsiInputDataDariPDF(jsonTeks As String)
    Dim wsMD As Worksheet
    On Error Resume Next
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)
    On Error GoTo 0
    If wsMD Is Nothing Then Exit Sub

    ' E16: Kegiatan/Sub Kegiatan
    Dim blok16 As String: blok16 = ExtractJSONBlok(jsonTeks, "input_data", "E16")
    If ExtractJSONVal(blok16, "status") = "terisi" Then
        If Trim(CStr(wsMD.Cells(MD_E16, 3).Value)) = "" Then
            wsMD.Cells(MD_E16, 3).Value = ExtractJSONVal(blok16, "nilai")
        End If
    End If

    ' E32: Lokasi
    Dim blok32 As String: blok32 = ExtractJSONBlok(jsonTeks, "input_data", "E32")
    If ExtractJSONVal(blok32, "status") = "terisi" Then
        If Trim(CStr(wsMD.Cells(MD_E32, 3).Value)) = "" Then
            wsMD.Cells(MD_E32, 3).Value = ExtractJSONVal(blok32, "nilai")
        End If
    End If

    ' E33: Sumber Dana
    Dim blok33 As String: blok33 = ExtractJSONBlok(jsonTeks, "input_data", "E33")
    If ExtractJSONVal(blok33, "status") = "terisi" Then
        If Trim(CStr(wsMD.Cells(MD_E33, 3).Value)) = "" Then
            wsMD.Cells(MD_E33, 3).Value = ExtractJSONVal(blok33, "nilai")
        End If
    End If
End Sub


' ============================================================
' ISI database_reviu dari JSON
' ============================================================
Private Sub IsiDatabaseReviu(jsonTeks As String)
    Dim wsMD As Worksheet
    On Error Resume Next
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)
    On Error GoTo 0
    If wsMD Is Nothing Then Exit Sub

    ' Mapping cell → path JSON (format: "reviu.EXX.nilai")
    Dim cells(25) As String, paths(25) As String
    cells(0)  = "E2":  paths(0)  = "E2"
    cells(1)  = "E6":  paths(1)  = "E6"
    cells(2)  = "E7":  paths(2)  = "E7"
    cells(3)  = "E9":  paths(3)  = "E9"
    cells(4)  = "E10": paths(4)  = "E10"
    cells(5)  = "E11": paths(5)  = "E11"
    cells(6)  = "E12": paths(6)  = "E12"
    cells(7)  = "E13": paths(7)  = "E13"
    cells(8)  = "E14": paths(8)  = "E14"
    cells(9)  = "E15": paths(9)  = "E15"
    cells(10) = "E16": paths(10) = "E16"
    cells(11) = "E17": paths(11) = "E17"
    cells(12) = "E18": paths(12) = "E18"
    cells(13) = "E19": paths(13) = "E19"
    cells(14) = "E20": paths(14) = "E20"
    cells(15) = "E21": paths(15) = "E21"
    cells(16) = "E22": paths(16) = "E22"
    cells(17) = "E23": paths(17) = "E23"
    cells(18) = "E24": paths(18) = "E24"
    cells(19) = "E25": paths(19) = "E25"
    cells(20) = "E26": paths(20) = "E26"
    cells(21) = "E27": paths(21) = "E27"
    cells(22) = "E28": paths(22) = "E28"
    cells(23) = "E29": paths(23) = "E29"
    cells(24) = "E30": paths(24) = "E30"
    cells(25) = "E31": paths(25) = "E31"

    Dim i As Integer
    For i = 0 To 25
        Dim blok As String
        blok = ExtractJSONBlok(jsonTeks, "reviu", paths(i))
        Dim status As String: status = ExtractJSONVal(blok, "status")
        Dim nilai As String:  nilai  = ExtractJSONVal(blok, "nilai")
        ' Hanya isi jika terisi
        If status = "terisi" And nilai <> "" Then
            ' Write to @ Master Data kolom C at mapped row
            Dim mdRow As Integer
            Select Case cells(i)
                Case "E2": mdRow = 25
                Case "E6": mdRow = 26
                Case "E7": mdRow = 27
                Case "E9": mdRow = 28
                Case "E10": mdRow = 29
                Case "E11": mdRow = 30
                Case "E12": mdRow = 31
                Case "E13": mdRow = 32
                Case "E14": mdRow = 33
                Case "E15": mdRow = 34
                Case "E16": mdRow = 35
                Case "E17": mdRow = 36
                Case "E18": mdRow = 37
                Case "E19": mdRow = 38
                Case "E20": mdRow = 39
                Case "E21": mdRow = 40
                Case "E22": mdRow = 41
                Case "E23": mdRow = 42
                Case "E24": mdRow = 43
                Case "E25": mdRow = 44
                Case "E26": mdRow = 45
                Case "E27": mdRow = 46
                Case "E28": mdRow = 47
                Case "E29": mdRow = 48
                Case "E30": mdRow = 49
                Case "E31": mdRow = 50
                Case Else: mdRow = 0
            End Select
            If mdRow > 0 Then wsMD.Cells(mdRow, 3).Value = nilai
        End If
    Next i

    ' E32, E33, E34
    Dim blok32 As String: blok32 = ExtractJSONBlok(jsonTeks, "reviu", "E32")
    If ExtractJSONVal(blok32, "status") = "terisi" Then
        wsMD.Cells(MD_R_E32, 3).Value = ExtractJSONVal(blok32, "nilai")
    End If
    Dim blok33 As String: blok33 = ExtractJSONBlok(jsonTeks, "reviu", "E33")
    If ExtractJSONVal(blok33, "status") = "terisi" Then
        Dim uraian As String: uraian = ExtractJSONVal(blok33, "nilai")
        uraian = Join(Split(uraian, "\n"), Chr(10))
        wsMD.Cells(MD_R_E33, 3).Value = uraian
    End If
    Dim blok34 As String: blok34 = ExtractJSONBlok(jsonTeks, "reviu", "E34")
    If ExtractJSONVal(blok34, "status") = "terisi" Then
        wsMD.Cells(MD_R_E34, 3).Value = ExtractJSONVal(blok34, "nilai")
    End If
End Sub


' ============================================================
' ISI database_dokpil dari JSON
' ============================================================
Private Sub IsiDatabaseDokpil(jsonTeks As String)
    Dim wsMD As Worksheet
    On Error Resume Next
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)
    On Error GoTo 0
    If wsMD Is Nothing Then Exit Sub

    Dim dokpilCells(9) As String
    dokpilCells(0) = "E6":  dokpilCells(1) = "E7"
    dokpilCells(2) = "E8":  dokpilCells(3) = "E9"
    dokpilCells(4) = "E10": dokpilCells(5) = "E11"
    dokpilCells(6) = "E12": dokpilCells(7) = "E13"
    dokpilCells(8) = "E14": dokpilCells(9) = "E15"

    Dim i As Integer
    For i = 0 To 9
        Dim blok As String
        blok = ExtractJSONBlok(jsonTeks, "dokpil", dokpilCells(i))
        Dim status As String: status = ExtractJSONVal(blok, "status")
        Dim nilai As String:  nilai  = ExtractJSONVal(blok, "nilai")
        If status = "terisi" And nilai <> "" Then
            Dim mdRowD As Integer
            Select Case dokpilCells(i)
                Case "E6": mdRowD = 56
                Case "E7": mdRowD = 57
                Case "E8": mdRowD = 58
                Case "E9": mdRowD = 59
                Case "E10": mdRowD = 60
                Case "E11": mdRowD = 61
                Case "E12": mdRowD = 62
                Case "E13": mdRowD = 63
                Case "E14": mdRowD = 64
                Case "E15": mdRowD = 65
                Case Else: mdRowD = 0
            End Select
            If mdRowD > 0 Then wsMD.Cells(mdRowD, 3).Value = nilai
        ElseIf status = "tidak_ada" Then
            ' Clear di master data juga
            Select Case dokpilCells(i)
                Case "E6": wsMD.Cells(56, 3).Value = ""
                Case "E7": wsMD.Cells(57, 3).Value = ""
                Case "E8": wsMD.Cells(58, 3).Value = ""
                Case "E9": wsMD.Cells(59, 3).Value = ""
                Case "E10": wsMD.Cells(60, 3).Value = ""
                Case "E11": wsMD.Cells(61, 3).Value = ""
                Case "E12": wsMD.Cells(62, 3).Value = ""
                Case "E13": wsMD.Cells(63, 3).Value = ""
                Case "E14": wsMD.Cells(64, 3).Value = ""
                Case "E15": wsMD.Cells(65, 3).Value = ""
            End Select
        End If
    Next i

    ' E16: Cara Pembayaran
    Dim blok16 As String: blok16 = ExtractJSONBlok(jsonTeks, "dokpil", "E16")
    If ExtractJSONVal(blok16, "status") = "terisi" Then
        wsMD.Cells(MD_D_E16, 3).Value = ExtractJSONVal(blok16, "nilai")
    End If
End Sub


' ============================================================
' TAMPILKAN _HasilParse — sheet checklist, MsgBox hanya 1x
' ============================================================
Public Sub TampilkanHasilParse(jsonTeks As String, folderWb As String)
    ' Buat/ambil sheet _HasilParse
    Dim wsHP As Worksheet
    On Error Resume Next
    Set wsHP = ThisWorkbook.Sheets(MD_SHEET)
    On Error GoTo 0
    If wsHP Is Nothing Then
        Set wsHP = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsHP.Name = MD_SHEET
    Else
        ' Pindah ke posisi paling kiri jika belum di sana
        If wsHP.Index <> 1 Then
            wsHP.Move Before:=ThisWorkbook.Sheets(1)
        End If
    End If
    wsHP.Visible = xlSheetVisible
    ' Simpan nilai kolom C (INPUT DATA, baris 3-22) sebelum clear
    Dim savedC(3 To 22) As Variant
    Dim r As Integer
    For r = 3 To 22
        savedC(r) = wsHP.Cells(r, 3).Value
    Next r
    ' Clear layout A:D semua baris (kolom C diisi ulang: INPUT DATA dari saved, REVIU/DOKPIL dari JSON)
    wsHP.Range("A2:E70").UnMerge
    wsHP.Range("A2:E70").ClearContents
    wsHP.Range("A2:E70").Interior.Pattern = xlNone
    wsHP.Range("A2:E70").Font.Bold = False
    ' Kembalikan nilai INPUT DATA ke kolom C (akan di-override oleh TulisBarisHasil jika JSON punya nilai)
    For r = 3 To 22
        If savedC(r) <> "" Then wsHP.Cells(r, 3).Value = savedC(r)
    Next r

    ' Header
    With wsHP
        .Range("A1").Value = "Field"
        .Range("B1").Value = "Status"
        .Range("C1").Value = "Nilai / Keterangan"
        .Range("D1").Value = "Navigasi"
        ' E1 = dropdown pilih paket (JANGAN dihapus/tulis ulang)
        .Range("E1").Value = "Sumber"
        .Range("A1:E1").Font.Bold = True
        .Range("A1:E1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:E1").Font.Color = RGB(255, 255, 255)
    End With

    Dim baris As Integer: baris = 2
    Dim jmlTerisi As Integer: jmlTerisi = 0
    Dim jmlKosong As Integer: jmlKosong = 0
    Dim jmlKeputusan As Integer: jmlKeputusan = 0
    Dim jmlTidakAda As Integer: jmlTidakAda = 0

    ' ─── Bagian 1. Input Data ────────────────────────────────────────────────
    TulisHeaderBagian wsHP, 2, "1. INPUT DATA"

    ' Field dari Supabase — baca nilai langsung dari sheet "1. Input Data"
    Dim wsInp As Worksheet
    On Error Resume Next
    Set wsInp = ThisWorkbook.Sheets("1. Input Data")
    On Error GoTo 0

    ' Pasangan: cell address, label
    Dim inpCells(16) As String, inpLabels(16) As String
    inpCells(0)  = "E3":  inpLabels(0)  = "Kode Rekening (MAK)"
    inpCells(1)  = "E5":  inpLabels(1)  = "Kode Tender"
    inpCells(2)  = "E6":  inpLabels(2)  = "Nama Tender"
    inpCells(3)  = "E8":  inpLabels(3)  = "Kode RUP"
    inpCells(4)  = "E10": inpLabels(4)  = "Nilai Pagu"
    inpCells(5)  = "E11": inpLabels(5)  = "Nilai HPS"
    inpCells(6)  = "E12": inpLabels(6)  = "Nomor Surat Permohonan"
    inpCells(7)  = "E13": inpLabels(7)  = "Nomor Surat Tugas"
    inpCells(8)  = "E14": inpLabels(8)  = "Kode Pokja"
    inpCells(9)  = "E15": inpLabels(9)  = "Masa Pelaksanaan (Hari)"
    inpCells(10) = "E17": inpLabels(10) = "SKPD/OPD"
    inpCells(11) = "E19": inpLabels(11) = "Nama PPK"
    inpCells(12) = "E20": inpLabels(12) = "NIP PPK"
    inpCells(13) = "E21": inpLabels(13) = "Nomor SK PPK"
    inpCells(14) = "E22": inpLabels(14) = "Anggota 1"
    inpCells(15) = "E23": inpLabels(15) = "Anggota 2"
    inpCells(16) = "E24": inpLabels(16) = "Anggota 3"

    Dim si As Integer
    For si = 0 To 16
        Dim nilaiInp As String
        ' Baca langsung dari @ Master Data kolom C (bukan dari sheet target)
        Dim mdRowInp As Integer
        Select Case inpCells(si)
            Case "E3": mdRowInp = MD_E3
            Case "E5": mdRowInp = MD_E5
            Case "E6": mdRowInp = MD_E6
            Case "E8": mdRowInp = MD_E8
            Case "E10": mdRowInp = MD_E10
            Case "E11": mdRowInp = MD_E11
            Case "E12": mdRowInp = MD_E12
            Case "E13": mdRowInp = MD_E13
            Case "E14": mdRowInp = MD_E14
            Case "E15": mdRowInp = MD_E15
            Case "E17": mdRowInp = MD_E17
            Case "E19": mdRowInp = MD_E19
            Case "E20": mdRowInp = MD_E20
            Case "E21": mdRowInp = MD_E21
            Case "E22": mdRowInp = MD_E22
            Case "E23": mdRowInp = MD_E23
            Case "E24": mdRowInp = MD_E24
            Case Else: mdRowInp = 0
        End Select
        If mdRowInp > 0 Then
            nilaiInp = Trim(CStr(wsHP.Cells(mdRowInp, 3).Value))
        Else
            nilaiInp = ""
        End If
        Dim statusInp As String
        If nilaiInp <> "" And nilaiInp <> "0" Then
            statusInp = "terisi"
        Else
            statusInp = ""
        End If
        Dim blokFake As String
        blokFake = "{""label"":""" & inpLabels(si) & """,""nilai"":""" & nilaiInp & """,""status"":""" & statusInp & """,""sumber"":""Supabase""}"
        If mdRowInp > 0 Then
            TulisBarisHasil wsHP, mdRowInp, blokFake, inpCells(si), "1. Input Data", _
                            jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda
        End If
    Next si

    ' Field dari PDF (E16 Kegiatan, E32 Lokasi, E33 Sumber Dana)
    Dim idKeys(2) As String
    idKeys(0) = "E16": idKeys(1) = "E32": idKeys(2) = "E33"

    Dim idIdx As Integer
    For idIdx = 0 To 2
        Dim blokID As String: blokID = ExtractJSONBlok(jsonTeks, "input_data", idKeys(idIdx))
        Dim mdRowPdf As Integer
        Select Case idKeys(idIdx)
            Case "E16": mdRowPdf = MD_E16
            Case "E32": mdRowPdf = MD_E32
            Case "E33": mdRowPdf = MD_E33
        End Select
        If mdRowPdf > 0 Then
            TulisBarisHasil wsHP, mdRowPdf, blokID, idKeys(idIdx), "1. Input Data", _
                            jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda
        End If
    Next idIdx

    baris = baris + 1

    ' ─── Bagian database_reviu ────────────────────────────────────────────────
    TulisHeaderBagian wsHP, 24, "DATABASE REVIU"

    Dim revKeys(27) As String
    revKeys(0)  = "E2":  revKeys(1)  = "E6":  revKeys(2)  = "E7"
    revKeys(3)  = "E9":  revKeys(4)  = "E10": revKeys(5)  = "E11"
    revKeys(6)  = "E12": revKeys(7)  = "E13": revKeys(8)  = "E14"
    revKeys(9)  = "E15": revKeys(10) = "E16": revKeys(11) = "E17"
    revKeys(12) = "E18": revKeys(13) = "E19": revKeys(14) = "E20"
    revKeys(15) = "E21": revKeys(16) = "E22": revKeys(17) = "E23"
    revKeys(18) = "E24": revKeys(19) = "E25": revKeys(20) = "E26"
    revKeys(21) = "E27": revKeys(22) = "E28": revKeys(23) = "E29"
    revKeys(24) = "E30": revKeys(25) = "E31": revKeys(26) = "E32"
    revKeys(27) = "E33" ' E34 ditambah di bawah

    Dim k As Integer
    For k = 0 To 27
        Dim blok As String: blok = ExtractJSONBlok(jsonTeks, "reviu", revKeys(k))
        Dim mdRowRev As Integer
        Select Case revKeys(k)
            Case "E2": mdRowRev = MD_R_E2
            Case "E6": mdRowRev = MD_R_E6
            Case "E7": mdRowRev = MD_R_E7
            Case "E9": mdRowRev = MD_R_E9
            Case "E10": mdRowRev = MD_R_E10
            Case "E11": mdRowRev = MD_R_E11
            Case "E12": mdRowRev = MD_R_E12
            Case "E13": mdRowRev = MD_R_E13
            Case "E14": mdRowRev = MD_R_E14
            Case "E15": mdRowRev = MD_R_E15
            Case "E16": mdRowRev = MD_R_E16
            Case "E17": mdRowRev = MD_R_E17
            Case "E18": mdRowRev = MD_R_E18
            Case "E19": mdRowRev = MD_R_E19
            Case "E20": mdRowRev = MD_R_E20
            Case "E21": mdRowRev = MD_R_E21
            Case "E22": mdRowRev = MD_R_E22
            Case "E23": mdRowRev = MD_R_E23
            Case "E24": mdRowRev = MD_R_E24
            Case "E25": mdRowRev = MD_R_E25
            Case "E26": mdRowRev = MD_R_E26
            Case "E27": mdRowRev = MD_R_E27
            Case "E28": mdRowRev = MD_R_E28
            Case "E29": mdRowRev = MD_R_E29
            Case "E30": mdRowRev = MD_R_E30
            Case "E31": mdRowRev = MD_R_E31
            Case "E32": mdRowRev = MD_R_E32
            Case "E33": mdRowRev = MD_R_E33
            Case Else: mdRowRev = 0
        End Select
        If mdRowRev > 0 Then
            TulisBarisHasil wsHP, mdRowRev, blok, revKeys(k), "database_reviu", _
                            jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda
        End If
    Next k
    ' E34
    Dim blok34 As String: blok34 = ExtractJSONBlok(jsonTeks, "reviu", "E34")
    TulisBarisHasil wsHP, MD_R_E34, blok34, "E34", "database_reviu", _
                    jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda

    ' ─── Bagian database_dokpil ───────────────────────────────────────────────
    TulisHeaderBagian wsHP, 55, "DATABASE DOKPIL"

    Dim dpKeys(10) As String
    dpKeys(0) = "E6":  dpKeys(1) = "E7":  dpKeys(2) = "E8"
    dpKeys(3) = "E9":  dpKeys(4) = "E10": dpKeys(5) = "E11"
    dpKeys(6) = "E12": dpKeys(7) = "E13": dpKeys(8) = "E14"
    dpKeys(9) = "E15": dpKeys(10) = "E16"

    Dim j As Integer
    For j = 0 To 10
        Dim blokD As String: blokD = ExtractJSONBlok(jsonTeks, "dokpil", dpKeys(j))
        Dim mdRowDok As Integer
        Select Case dpKeys(j)
            Case "E6": mdRowDok = MD_D_E6
            Case "E7": mdRowDok = MD_D_E7
            Case "E8": mdRowDok = MD_D_E8
            Case "E9": mdRowDok = MD_D_E9
            Case "E10": mdRowDok = MD_D_E10
            Case "E11": mdRowDok = MD_D_E11
            Case "E12": mdRowDok = MD_D_E12
            Case "E13": mdRowDok = MD_D_E13
            Case "E14": mdRowDok = MD_D_E14
            Case "E15": mdRowDok = MD_D_E15
            Case "E16": mdRowDok = MD_D_E16
            Case Else: mdRowDok = 0
        End Select
        If mdRowDok > 0 Then
            TulisBarisHasil wsHP, mdRowDok, blokD, dpKeys(j), "database_dokpil", _
                            jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda
        End If
    Next j


    ' Format kolom
    wsHP.Columns("A").ColumnWidth = 35
    wsHP.Columns("B").ColumnWidth = 20
    wsHP.Columns("C").ColumnWidth = 60
    wsHP.Columns("D").ColumnWidth = 15
    wsHP.Range("C2:C70").WrapText = False

    ' Cek apakah MsgBox sudah pernah ditampilkan (flag di A1 comment)
    Dim sudahTampil As Boolean
    sudahTampil = False
    If wsHP.Range("A1").Comment Is Nothing Then
        sudahTampil = False
    Else
        sudahTampil = (wsHP.Range("A1").Comment.Text = "SHOWN")
    End If

    If Not sudahTampil Then
        ' Ringkasan MsgBox 1x
        Dim msg As String
        msg = "Hasil parsing Draft PDF selesai:" & vbCrLf & vbCrLf
        msg = msg & "  Terisi otomatis : " & jmlTerisi & " field" & vbCrLf
        msg = msg & "  Perlu keputusan : " & jmlKeputusan & " field (SBU)" & vbCrLf
        msg = msg & "  Tidak ada di PDF: " & jmlTidakAda & " field" & vbCrLf
        msg = msg & "  Belum terisi    : " & jmlKosong & " field" & vbCrLf & vbCrLf
        msg = msg & "Lihat sheet '_HasilParse' untuk detail dan navigasi ke cell yang perlu diisi manual."
        MsgBox msg, vbInformation, "Parse Draft PDF"

        ' Tandai sudah ditampilkan
        On Error Resume Next
        wsHP.Range("A1").AddComment "SHOWN"
        On Error GoTo 0
    End If

    ' Aktifkan sheet _HasilParse
    wsHP.Activate
    wsHP.Range("A1").Select
End Sub

Private Sub TampilkanHasilParseNoPDF(folderWb As String)
    Dim wsHP As Worksheet
    On Error Resume Next
    Set wsHP = ThisWorkbook.Sheets(MD_SHEET)
    On Error GoTo 0
    If wsHP Is Nothing Then
        Set wsHP = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsHP.Name = MD_SHEET
    Else
        If wsHP.Index <> 1 Then wsHP.Move Before:=ThisWorkbook.Sheets(1)
    End If
    wsHP.Visible = xlSheetVisible
    wsHP.Cells.Clear
    wsHP.Range("A1").Value = "INFO"
    wsHP.Range("B1").Value = "File Draft_Pokja_XXX.pdf tidak ditemukan."
    wsHP.Range("C1").Value = "Jalankan 'Download Dokumen' di Asisten Pokja terlebih dahulu, atau salin PDF ke folder yang sama dengan file Excel."
    wsHP.Columns("A:C").AutoFit
    wsHP.Activate
End Sub

Private Sub TulisHeaderBagian(ws As Worksheet, baris As Integer, judul As String)
    ws.Range("A" & baris & ":D" & baris).Merge
    ws.Range("A" & baris).Value = judul
    ws.Range("A" & baris).Font.Bold = True
    ws.Range("A" & baris).Interior.Color = RGB(189, 215, 238)
End Sub

Private Sub TulisBarisHasil(ws As Worksheet, baris As Integer, blok As String, _
                             cellAddr As String, namaSheet As String, _
                             ByRef jmlTerisi As Integer, ByRef jmlKosong As Integer, _
                             ByRef jmlKeputusan As Integer, ByRef jmlTidakAda As Integer)
    Dim status As String: status = ExtractJSONVal(blok, "status")
    Dim label  As String: label  = ExtractJSONVal(blok, "label")
    Dim nilai  As String: nilai  = ExtractJSONVal(blok, "nilai")

    Dim statusTeks As String
    Dim warna As Long

    Select Case status
        Case "terisi"
            statusTeks = "Terisi Otomatis"
            warna = RGB(198, 239, 206)  ' Hijau muda
            jmlTerisi = jmlTerisi + 1
        Case "perlu_keputusan"
            statusTeks = "Perlu Keputusan"
            warna = RGB(255, 235, 156)  ' Kuning
            jmlKeputusan = jmlKeputusan + 1
        Case "tidak_ada"
            statusTeks = "Tidak Ada di PDF"
            warna = RGB(217, 217, 217)  ' Abu
            jmlTidakAda = jmlTidakAda + 1
        Case Else
            statusTeks = "Belum Terisi"
            warna = RGB(255, 199, 206)  ' Merah muda
            jmlKosong = jmlKosong + 1
    End Select

    With ws
        .Cells(baris, 1).Value = label & " (" & cellAddr & ")"
        .Cells(baris, 2).Value = statusTeks
        .Cells(baris, 2).Interior.Color = warna
        ' Kolom C: tulis nilai dari JSON (kolom sudah di-clear sebelum TampilkanHasilParse)
        If nilai <> "" And nilai <> "0" Then
            .Cells(baris, 3).Value = nilai
        End If

        ' Hyperlink navigasi ke cell target
        If status <> "tidak_ada" Then
            On Error Resume Next
            Dim wsTarget As Worksheet
            Set wsTarget = ThisWorkbook.Sheets(namaSheet)
            If Not wsTarget Is Nothing Then
                ws.Hyperlinks.Add Anchor:=ws.Cells(baris, 4), _
                    Address:="", _
                    SubAddress:="'" & namaSheet & "'!" & cellAddr, _
                    TextToDisplay:="Buka " & cellAddr
            End If
            On Error GoTo 0
        End If

        ' Kolom E: sumber data
        Dim sumber As String: sumber = ExtractJSONVal(blok, "sumber")
        If sumber <> "" Then
            .Cells(baris, 5).Value = sumber
        End If
    End With
End Sub


' ============================================================
' HELPERS: Cari file PDF, path, baca file
' ============================================================
Private Function CariDraftPDF(folder As String, kodePokja As String) As String
    CariDraftPDF = ""
    ' Pattern: Draft_Pokja_086.pdf atau Draft_Pokja_86.pdf
    Dim patterns(2) As String
    patterns(0) = folder & "\Draft_Pokja_" & Format(CLng(kodePokja), "000") & ".pdf"
    patterns(1) = folder & "\Draft_Pokja_" & kodePokja & ".pdf"
    patterns(2) = folder & "\Draft_Pokja*.pdf"

    Dim i As Integer
    For i = 0 To 1
        If FileExists(patterns(i)) Then
            CariDraftPDF = patterns(i)
            Exit Function
        End If
    Next i

    ' Wildcard: ambil file pertama yang match
    Dim found As String
    found = Dir(patterns(2))
    If found <> "" Then
        CariDraftPDF = folder & "\" & found
    End If
End Function

Private Function CariBaseDir(folderWorkbook As String) As String
    ' Struktur: D:\Dokumen\@ POKJA 2026\{N}. Pokja XXX\  ← folderWorkbook
    '           D:\Dokumen\@ POKJA 2026\V19_Scheduler\WPy64-313110\
    ' Naik 1 level dari folderWorkbook → @ POKJA 2026, lalu masuk V19_Scheduler\WPy64-313110
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim pokjaRoot As String
    pokjaRoot = fso.GetParentFolderName(folderWorkbook)
    CariBaseDir = pokjaRoot & "\V19_Scheduler\WPy64-313110"
End Function

Private Function FileExists(path As String) As Boolean
    On Error Resume Next
    FileExists = (Dir(path) <> "")
    On Error GoTo 0
End Function

Private Function BacaFile(path As String) As String
    ' Baca file UTF-8 via ADODB.Stream agar karakter Unicode tidak rusak
    On Error GoTo fallback
    Dim stm As Object
    Set stm = CreateObject("ADODB.Stream")
    stm.Type = 2      ' adTypeText
    stm.Charset = "utf-8"
    stm.Open
    stm.LoadFromFile path
    BacaFile = stm.ReadText()
    stm.Close
    Exit Function
fallback:
    ' Fallback: VBA default (ANSI)
    Dim f As Integer: f = FreeFile
    Dim teks As String, baris As String
    Open path For Input As #f
    Do While Not EOF(f)
        Line Input #f, baris
        teks = teks & baris & Chr(10)
    Loop
    Close #f
    BacaFile = teks
End Function


' ============================================================
' HELPER: Ekstrak blok JSON untuk field tertentu di nested object
' Format: {"reviu": {"E2": {"nilai":"...", "status":"..."}}}
' ============================================================
Private Function ExtractJSONBlok(jsonTeks As String, bagian As String, cellKey As String) As String
    ' Cari posisi bagian (reviu/dokpil)
    Dim pBagian As Long
    pBagian = InStr(jsonTeks, """" & bagian & """")
    If pBagian = 0 Then ExtractJSONBlok = "{}": Exit Function

    ' Cari posisi cellKey dalam bagian itu
    Dim pKey As Long
    pKey = InStr(pBagian, jsonTeks, """" & cellKey & """")
    If pKey = 0 Then ExtractJSONBlok = "{}": Exit Function

    ' Ambil blok { } setelah cellKey
    Dim pBrace As Long
    pBrace = InStr(pKey, jsonTeks, "{")
    If pBrace = 0 Then ExtractJSONBlok = "{}": Exit Function

    Dim depth As Integer: depth = 0
    Dim pos As Long: pos = pBrace
    Do While pos <= Len(jsonTeks)
        Dim c As String: c = Mid(jsonTeks, pos, 1)
        If c = "{" Then depth = depth + 1
        If c = "}" Then
            depth = depth - 1
            If depth = 0 Then
                ExtractJSONBlok = Mid(jsonTeks, pBrace, pos - pBrace + 1)
                Exit Function
            End If
        End If
        pos = pos + 1
    Loop
    ExtractJSONBlok = "{}"
End Function


' ============================================================
' LOOKUP PPK: Cari NIP + Nomor SK dari sheet "0. Data Nama Pokja & PPK"
' Kolom: A=Nama Pokja, C=Nama PPK, D=Gelar PPK, E=NIP PPK, F=Nomor SK
' ============================================================
Private Sub LookupPPK(namaPPK As String, ByRef nip As String, ByRef noSK As String, ByRef gelar As String)
    nip = "": noSK = "": gelar = ""
    If namaPPK = "" Then Exit Sub

    Dim wsRef As Worksheet
    On Error Resume Next
    Set wsRef = ThisWorkbook.Sheets("0. Data Nama Pokja & PPK")
    On Error GoTo 0
    If wsRef Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = wsRef.Cells(wsRef.Rows.Count, 3).End(xlUp).Row

    Dim r As Long
    For r = 2 To lastRow
        Dim namaCel As String
        namaCel = Trim(CStr(wsRef.Cells(r, 3).Value))  ' Kolom C = Nama PPK
        ' Bandingkan case-insensitive, abaikan gelar di belakang koma
        Dim namaRef As String
        Dim komRef As Long: komRef = InStr(namaCel, ",")
        If komRef > 0 Then
            namaRef = Trim(Left(namaCel, komRef - 1))
        Else
            namaRef = namaCel
        End If
        If UCase(namaRef) = UCase(namaPPK) Or UCase(namaCel) = UCase(namaPPK) Then
            nip   = Trim(CStr(wsRef.Cells(r, 5).Value))  ' Kolom E = NIP PPK
            noSK  = Trim(CStr(wsRef.Cells(r, 6).Value))  ' Kolom F = Nomor SK
            gelar = Trim(CStr(wsRef.Cells(r, 4).Value))  ' Kolom D = Gelar PPK
            Exit For
        End If
    Next r
End Sub


' ============================================================
' ISI ANGGOTA POKJA: Isi E22/E23/E24 dari sheet "0. Data Nama Pokja & PPK"
' Logika: cari semua baris di kolom A yang nama pokjanya MIRIP kode pokja
' Fallback: isi nama anggota yang sering muncul di paket dengan kode pokja sama
' ============================================================
Private Sub IsiAnggotaPokja(wsInput As Worksheet, a1 As String, a2 As String, a3 As String)
    ' Hanya isi jika cell masih kosong (tidak menimpa input manual)
    If Trim(CStr(wsInput.Range("E22").Value)) = "" And a1 <> "" Then
        wsInput.Range("E22").Value = a1
    End If
    If Trim(CStr(wsInput.Range("E23").Value)) = "" And a2 <> "" Then
        wsInput.Range("E23").Value = a2
    End If
    If Trim(CStr(wsInput.Range("E24").Value)) = "" And a3 <> "" Then
        wsInput.Range("E24").Value = a3
    End If
End Sub


' ============================================================
' MUAT HPS: GET hps_items dari Supabase → isi sheet "5. HPS"
' ============================================================
Public Sub MuatHPS(kodeTender As String)
    If kodeTender = "" Then Exit Sub

    Dim wsHPS As Worksheet
    On Error Resume Next
    Set wsHPS = ThisWorkbook.Sheets("5. HPS")
    On Error GoTo 0
    If wsHPS Is Nothing Then Exit Sub

    ' Fetch JSON dari Supabase
    Dim url As String
    url = SB_URL & "/rest/v1/hps_items" & _
          "?kode_tender=eq." & kodeTender & _
          "&order=urutan.asc" & _
          "&select=urutan,jenis_bj,satuan,vol,harga,pajak_pct,total_spse,total_hitung,is_divisi,selisih,selisih_ok,total_nilai,total_nilai_bulat"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo ErrHPS
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send

    If http.Status <> 200 Then
        MsgBox "Gagal fetch HPS: HTTP " & http.Status, vbExclamation, "Muat HPS"
        Exit Sub
    End If

    Dim json As String: json = http.ResponseText

    ' Cek kosong
    If json = "[]" Or json = "" Then
        ' HPS belum tersedia di Supabase — tidak perlu MsgBox, silent
        Exit Sub
    End If

    ' ── Bersihkan sheet "5. HPS" baris 2 ke bawah (A–J) ──────────────────
    Dim lastRow As Long
    lastRow = wsHPS.Cells(wsHPS.Rows.Count, 2).End(xlUp).Row
    If lastRow >= 2 Then wsHPS.Range("A2:J" & lastRow).ClearContents

    ' ── Parse JSON array → isi baris ──────────────────────────────────────
    Dim pos As Long: pos = 1
    Dim baris As Long: baris = 2
    Dim adaSelisih As Boolean: adaSelisih = False
    Dim selisihMsg As String: selisihMsg = ""
    Dim totalNilai As String: totalNilai = ""
    Dim totalBulat As String: totalBulat = ""
    Dim totalHitungAll As Double: totalHitungAll = 0

    Do
        Dim bs As Long: bs = InStr(pos, json, "{")
        If bs = 0 Then Exit Do
        Dim depth As Integer: depth = 1
        Dim p As Long: p = bs + 1
        Do While p <= Len(json) And depth > 0
            Dim ch As String: ch = Mid(json, p, 1)
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then depth = depth - 1
            p = p + 1
        Loop
        Dim be As Long: be = p - 1
        Dim obj As String: obj = Mid(json, bs, be - bs + 1)

        Dim isDivisi As String: isDivisi = ExtractJSONVal(obj, "is_divisi")
        Dim jenisBJ  As String: jenisBJ  = ExtractJSONVal(obj, "jenis_bj")
        Dim satuan   As String: satuan   = ExtractJSONVal(obj, "satuan")
        Dim volS     As String: volS     = ExtractJSONVal(obj, "vol")
        Dim hargaS   As String: hargaS   = ExtractJSONVal(obj, "harga")
        Dim pajakS   As String: pajakS   = ExtractJSONVal(obj, "pajak_pct")
        Dim totSpseS As String: totSpseS = ExtractJSONVal(obj, "total_spse")
        Dim totHitS  As String: totHitS  = ExtractJSONVal(obj, "total_hitung")
        Dim selisihS As String: selisihS = ExtractJSONVal(obj, "selisih")
        Dim selOkS   As String: selOkS   = ExtractJSONVal(obj, "selisih_ok")
        Dim urutan   As String: urutan   = ExtractJSONVal(obj, "urutan")

        ' Simpan total nilai dari item pertama (sama di semua baris)
        If totalNilai = "" Then
            totalNilai = ExtractJSONVal(obj, "total_nilai")
            totalBulat = ExtractJSONVal(obj, "total_nilai_bulat")
        End If

        With wsHPS
            .Cells(baris, 1).Value = CDblSafe(urutan)  ' A: nomor urut
            .Cells(baris, 2).Value = jenisBJ            ' B: Jenis B/J

            If isDivisi <> "true" Then
                .Cells(baris, 3).Value = satuan          ' C: Satuan
                If volS <> "" Then
                    .Cells(baris, 4).Value = CDblSafe(volS)
                    .Cells(baris, 4).NumberFormat = "#,##0.00"
                End If
                If hargaS <> "" Then
                    .Cells(baris, 5).Value = CDblSafe(hargaS)
                    .Cells(baris, 5).NumberFormat = "#,##0.00"
                End If
                If pajakS <> "" Then
                    .Cells(baris, 6).Value = CDblSafe(pajakS)
                    .Cells(baris, 6).NumberFormat = "0.00"
                End If
                If totSpseS <> "" Then
                    .Cells(baris, 7).Value = CDblSafe(totSpseS)
                    .Cells(baris, 7).NumberFormat = "#,##0.00"
                End If
                ' Kolom H: total hitung manual
                If totHitS <> "" Then
                    .Cells(baris, 8).Value = CDblSafe(totHitS)
                    .Cells(baris, 8).NumberFormat = "#,##0.00"
                End If
                ' Kolom I: selisih
                Dim selisihVal As Double: selisihVal = CDblSafe(selisihS)
                .Cells(baris, 9).Value = selisihVal
                .Cells(baris, 9).NumberFormat = "#,##0.00"

                totalHitungAll = totalHitungAll + CDblSafe(totHitS)

                ' Highlight baris jika selisih > 1
                If selOkS = "false" Then
                    .Range(.Cells(baris, 1), .Cells(baris, 9)).Interior.Color = RGB(255, 255, 0)
                    adaSelisih = True
                    selisihMsg = selisihMsg & "  Baris " & baris & ": " & jenisBJ & _
                                 " (selisih Rp " & Format(selisihVal, "#,##0.00") & ")" & vbCrLf
                End If
            End If
        End With

        baris = baris + 1
        pos = be + 1
    Loop

    ' ── Tulis header kolom H–I jika belum ada ─────────────────────────────
    If wsHPS.Cells(1, 8).Value = "" Then
        wsHPS.Cells(1, 8).Value = "Total (Hitung)"
        wsHPS.Cells(1, 8).Font.Bold = True
    End If
    If wsHPS.Cells(1, 9).Value = "" Then
        wsHPS.Cells(1, 9).Value = "Selisih"
        wsHPS.Cells(1, 9).Font.Bold = True
    End If

    ' ── Tulis total di baris terakhir+1 ───────────────────────────────────
    Dim barisTotal As Long: barisTotal = baris + 1
    With wsHPS
        .Cells(barisTotal, 2).Value = "TOTAL NILAI (SPSE)"
        .Cells(barisTotal, 7).Value = CDblSafe(totalNilai)
        .Cells(barisTotal, 7).NumberFormat = "#,##0.00"
        .Cells(barisTotal, 2).Font.Bold = True

        .Cells(barisTotal + 1, 2).Value = "TOTAL NILAI (Setelah Pembulatan SPSE)"
        .Cells(barisTotal + 1, 7).Value = CDblSafe(totalBulat)
        .Cells(barisTotal + 1, 7).NumberFormat = "#,##0.00"
        .Cells(barisTotal + 1, 2).Font.Bold = True

        .Cells(barisTotal + 2, 2).Value = "TOTAL HITUNG MANUAL"
        .Cells(barisTotal + 2, 8).Value = totalHitungAll
        .Cells(barisTotal + 2, 8).NumberFormat = "#,##0.00"
        .Cells(barisTotal + 2, 2).Font.Bold = True
    End With

    ' ── MsgBox ringkasan ──────────────────────────────────────────────────
    Dim itemCount As Long: itemCount = baris - 2
    Dim msg As String
    msg = "HPS berhasil dimuat: " & itemCount & " baris."
    If adaSelisih Then
        msg = msg & vbCrLf & vbCrLf & _
              "⚠️ Terdapat selisih antara total SPSE vs hitung manual (highlight kuning):" & vbCrLf & selisihMsg
        MsgBox msg, vbExclamation, "Muat HPS"
    End If
    ' Jika tidak ada selisih — silent (tidak ganggu alur pilih paket)
    Exit Sub

ErrHPS:
    ' Gagal muat HPS — silent, tidak mengganggu alur utama
End Sub


' ============================================================
' ISI ANGGOTA POKJA ke @ Master Data
' ============================================================
Private Sub IsiAnggotaPokjaToMaster(wsMD As Worksheet, anggota1 As String, anggota2 As String, anggota3 As String)
    If anggota1 <> "" Then wsMD.Cells(MD_E22, 3).Value = anggota1
    If anggota2 <> "" Then wsMD.Cells(MD_E23, 3).Value = anggota2
    If anggota3 <> "" Then wsMD.Cells(MD_E24, 3).Value = anggota3
End Sub


' Helper: konversi string angka ke Double (toleran koma/titik)
Private Function CDblSafe(s As String) As Double
    If s = "" Or s = "null" Then CDblSafe = 0: Exit Function
    On Error Resume Next
    CDblSafe = CDbl(s)
    On Error GoTo 0
End Function


' ============================================================
' HTTP: GET ke Supabase REST API
' ============================================================
Private Function FetchSupabase() As String
    Dim url As String
    url = SB_URL & "/rest/v1/" & SB_TABLE & "?select=" & SB_SELECT & "&nomor_pp=ilike.*" & Year(Now) & "*&order=kode_pokja.asc"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo ErrHTTP
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send

    If http.Status = 200 Then
        FetchSupabase = http.ResponseText
    Else
        MsgBox "HTTP Error " & http.Status & ": " & http.StatusText & vbCrLf & _
               "URL: " & url, vbExclamation, "FetchSupabase Error"
        FetchSupabase = ""
    End If
    Exit Function

ErrHTTP:
    MsgBox "WinHTTP Error: " & Err.Number & " - " & Err.Description & vbCrLf & _
           "URL: " & url, vbCritical, "FetchSupabase Error"
    FetchSupabase = ""
End Function


' ============================================================
' PARSER: JSON array sederhana → Collection of Variant arrays
' ============================================================
' Format JSON: [{"kode_tender":"...","nama_tender":"...",...}, ...]
' Setiap item di-store sebagai Variant array sesuai urutan SB_SELECT:
'   0=kode_pokja, 1=nama_tender, 2=mak, 3=kode_rup, 4=nilai_pagu (NOTE: bukan kode_tender)
'   Urutan: kode_tender,nama_tender,mak,kode_rup,nilai_pagu,nilai_hps,
'           kode_pokja,nomor_pp,nomor_surat_dinas,nama_dinas,nama_ppk,jangka_waktu,sumber_anggaran

Private Function ParseDraftJSON(json As String) As Collection
    Dim col As New Collection

    ' Ekstrak setiap objek {...} dalam array
    Dim pos As Long: pos = 1
    Dim braceStart As Long, braceEnd As Long
    Dim depth As Long
    Dim inQuote As Boolean
    Dim scanPos As Long
    Dim ch As String
    Dim obj As String
    Do
        Dim item(20) As Variant
        braceStart = InStr(pos, json, "{")
        If braceStart = 0 Then Exit Do
        ' Cari matching } dengan brace counting (skip isi string)
        depth = 0
        inQuote = False
        braceEnd = 0
        For scanPos = braceStart To Len(json)
            ch = Mid(json, scanPos, 1)
            If ch = """" And (scanPos = 1 Or Mid(json, scanPos - 1, 1) <> "\") Then
                inQuote = Not inQuote
            ElseIf Not inQuote Then
                If ch = "{" Then depth = depth + 1
                If ch = "}" Then
                    depth = depth - 1
                    If depth = 0 Then
                        braceEnd = scanPos
                        Exit For
                    End If
                End If
            End If
        Next scanPos
        If braceEnd = 0 Then Exit Do

        obj = Mid(json, braceStart, braceEnd - braceStart + 1)

        item(0)  = ExtractJSONVal(obj, "kode_pokja")
        item(1)  = ExtractJSONVal(obj, "nama_tender")
        item(2)  = ExtractJSONVal(obj, "mak")
        item(3)  = ExtractJSONVal(obj, "kode_rup")
        item(4)  = ExtractJSONVal(obj, "nilai_pagu")
        item(5)  = ExtractJSONVal(obj, "nilai_hps")
        item(6)  = ExtractJSONVal(obj, "kode_tender")
        item(7)  = ExtractJSONVal(obj, "nomor_pp")
        item(8)  = ExtractJSONVal(obj, "nomor_surat_dinas")
        item(9)  = ExtractJSONVal(obj, "nama_dinas")
        item(10) = ExtractJSONVal(obj, "nama_ppk")
        item(11) = ExtractJSONVal(obj, "jangka_waktu")
        item(12) = ExtractJSONVal(obj, "sumber_anggaran")
        item(13) = ExtractJSONVal(obj, "anggota_1")
        item(14) = ExtractJSONVal(obj, "anggota_2")
        item(15) = ExtractJSONVal(obj, "anggota_3")
        item(16) = ExtractJSONVal(obj, "bidang")
        item(17) = ExtractJSONVal(obj, "sbu_baru")
        item(18) = ExtractJSONVal(obj, "sbu_lama")
        item(19) = ExtractJSONVal(obj, "nip_ppk")
        item(20) = ExtractJSONVal(obj, "sk_ppk")

        col.Add item
        pos = braceEnd + 1
    Loop

    Set ParseDraftJSON = col
End Function


' ============================================================
' HELPER: Ekstrak nilai string dari pasangan key:value JSON
' ============================================================
Private Function ExtractJSONVal(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """" & ":"

    Dim p As Long
    p = InStr(json, pattern)
    If p = 0 Then
        ExtractJSONVal = ""
        Exit Function
    End If

    p = p + Len(pattern)
    ' Skip spasi
    Do While Mid(json, p, 1) = " ": p = p + 1: Loop

    If Mid(json, p, 1) = """" Then
        ' String value
        p = p + 1
        Dim q As Long
        q = p
        Do While q <= Len(json)
            If Mid(json, q, 1) = """" And (q = 1 Or Mid(json, q - 1, 1) <> "\") Then Exit Do
            q = q + 1
        Loop
        ExtractJSONVal = Mid(json, p, q - p)
    ElseIf Mid(json, p, 4) = "null" Then
        ExtractJSONVal = ""
    Else
        ' Number/boolean: ambil sampai koma atau kurung tutup
        Dim endPos As Long
        endPos = p
        Do While endPos <= Len(json)
            Dim c As String: c = Mid(json, endPos, 1)
            If c = "," Or c = "}" Or c = "]" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractJSONVal = Trim(Mid(json, p, endPos - p))
    End If
End Function

' ============================================================
' Override baris SBU di _HasilParse setelah diisi dari Supabase
' ============================================================
Private Sub OverrideSBUdiHasilParse(sbuBaru As String, sbuLama As String)
    Dim wsHP As Worksheet
    On Error Resume Next
    Set wsHP = ThisWorkbook.Sheets(MD_SHEET)
    On Error GoTo 0
    If wsHP Is Nothing Then Exit Sub

    ' E6 = baris MD_R_E6 = 26 → override status + nilai dari Supabase
    If sbuBaru <> "" Then
        wsHP.Cells(MD_R_E6, 2).Value = "Terisi (Supabase)"
        wsHP.Cells(MD_R_E6, 2).Interior.Color = RGB(198, 239, 206)
        wsHP.Cells(MD_R_E6, 3).Value = sbuBaru
    End If

    ' E7 = baris MD_R_E7 = 27
    If sbuLama <> "" Then
        wsHP.Cells(MD_R_E7, 2).Value = "Terisi (Supabase)"
        wsHP.Cells(MD_R_E7, 2).Interior.Color = RGB(198, 239, 206)
        wsHP.Cells(MD_R_E7, 3).Value = sbuLama
    End If
End Sub

' ============================================================
' UPDATE HPS SAJA (Tanpa Parse Draft)
' Dipanggil dari tombol "Update HPS Saja" di @ Master Data
' ============================================================
Public Sub UpdateHPSSaja()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("1. Input Data")
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet '1. Input Data' tidak ditemukan.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Dim kodeTender As String
    kodeTender = Trim(CStr(ws.Range("B5").Value))
    
    If kodeTender = "" Then
        MsgBox "Kode Tender tidak ditemukan di '1. Input Data' B5." & vbCrLf & _
               "Pastikan data paket sudah dimuat/dipilih.", vbExclamation, "Error"
        Exit Sub
    End If
    
    MuatHPS kodeTender
    MsgBox "Proses fetch dan update data HPS selesai." & vbCrLf & _
           "Silakan cek sheet '5. HPS'.", vbInformation, "Update HPS Saja"
End Sub
