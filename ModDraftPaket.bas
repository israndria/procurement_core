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
Private Const CELL_SELECTOR As String = "E2"   ' Dropdown pilih paket

' Kolom yang di-fetch (efisien, tidak perlu semua)
Private Const SB_SELECT As String = "kode_tender,nama_tender,mak,kode_rup,nilai_pagu,nilai_hps,kode_pokja,nomor_pp,nomor_surat_dinas,nama_dinas,nama_ppk,jangka_waktu,sumber_anggaran,anggota_1,anggota_2,anggota_3,bidang"

' Cache data in-memory (Collection of Dictionary-like arrays)
Private m_DataCache As Collection
Private m_LastLoad As Date


' ============================================================
' FUNGSI UTAMA: Load dari Supabase, buat dropdown
' ============================================================
Public Sub MuatDraftPaket(Optional bFromOpen As Boolean = False)
    Dim ws As Worksheet
    On Error GoTo ErrHandler
    Set ws = ThisWorkbook.Sheets(SHEET_INPUT)

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

    ' Simpan nilai E2 yang sudah dipilih sebelum rebuild dropdown
    Dim savedLabel As String
    savedLabel = Trim(CStr(ws.Range(CELL_SELECTOR).Value))

    If savedLabel = "" Then
        ' Belum ada pilihan — tampilkan MsgBox panduan
        ws.Range(CELL_SELECTOR).Value = ""
        MsgBox m_DataCache.Count & " paket berhasil dimuat." & vbCrLf & _
               "Pilih paket di cell " & CELL_SELECTOR & " untuk mengisi data.", vbInformation, "Draft Paket Dimuat"
    ElseIf bFromOpen Then
        ' Dipanggil dari Workbook_Open — restore dropdown saja, tidak parse ulang
        ' (data sudah ada di sheet dari sesi sebelumnya)
        Application.EnableEvents = False
        ws.Range(CELL_SELECTOR).Value = savedLabel
        Application.EnableEvents = True
    Else
        ' Dipanggil manual (klik tombol) — parse ulang untuk refresh data
        Application.EnableEvents = False
        ws.Range(CELL_SELECTOR).Value = savedLabel
        Application.EnableEvents = True
        PilihDraftPaket savedLabel
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error MuatDraftPaket: " & Err.Description, vbCritical
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
            ' Mapping cell berdasarkan sheet "1. Input Data" v1.4
            With ws
                .Range("E3").Value  = CStr(item(2))   ' MAK
                .Range("E5").Value  = CStr(item(6))   ' Kode Tender
                .Range("E6").Value  = CStr(item(1))   ' Nama Tender
                .Range("E8").Value  = CStr(item(3))   ' Kode RUP
                .Range("E10").Value = CStr(item(4))   ' Nilai Pagu
                .Range("E11").Value = CStr(item(5))   ' Nilai HPS
                .Range("E12").Value = CStr(item(8))   ' Nomor Surat Dinas
                .Range("E13").Value = CStr(item(7))   ' Nomor PP
                .Range("E14").NumberFormat = "@"
                .Range("E14").Value = CStr(item(0))   ' Kode Pokja
                ' Jangka Waktu → E15 (angka hari saja)
                Dim jw As String: jw = Trim(CStr(item(11)))
                Dim spPos As Long: spPos = InStr(jw, " ")
                If spPos > 0 Then jw = Left(jw, spPos - 1)
                If IsNumeric(jw) Then
                    .Range("E15").Value = CLng(jw)
                Else
                    .Range("E15").Value = CStr(item(11))
                End If
                ' Nama Dinas lengkap → F17 (SKPD/OPD)
                .Range("F17").Value = CStr(item(9))
                ' Bidang (khusus PUPR: Bina Marga/SDA/CK) → E17, kosong jika bukan PUPR
                .Range("E17").Value = CStr(item(16))
                ' Nama PPK → E19 + F19 (gelar), NIP → E20, Nomor SK → E21
                Dim ppkNama As String: ppkNama = Trim(CStr(item(10)))
                Dim komaPos As Long: komaPos = InStr(ppkNama, ",")
                Dim ppkNamaBersih As String
                If komaPos > 0 Then
                    ppkNamaBersih = Trim(Left(ppkNama, komaPos - 1))
                    .Range("E19").Value = ppkNamaBersih
                    .Range("F19").Value = Trim(Mid(ppkNama, komaPos))
                Else
                    ppkNamaBersih = ppkNama
                    .Range("E19").Value = ppkNama
                End If
                ' Lookup NIP PPK + Nomor SK dari sheet "0. Data Nama Pokja & PPK"
                Dim nipPPK As String, skPPK As String, gelarPPK As String
                LookupPPK ppkNamaBersih, nipPPK, skPPK, gelarPPK
                If nipPPK <> "" Then
                    .Range("E20").NumberFormat = "@"
                    .Range("E20").Value = nipPPK
                End If
                If skPPK <> "" Then .Range("E21").Value = skPPK
                If gelarPPK <> "" And komaPos = 0 Then .Range("F19").Value = gelarPPK
                ' Sumber Anggaran → E33
                .Range("E33").Value = CStr(item(12))
            End With
            ' Isi anggota pokja dari sheet "0. Data Nama Pokja & PPK" berdasarkan Kode Pokja
            IsiAnggotaPokja ws, CStr(item(13)), CStr(item(14)), CStr(item(15))

            ' ── Parse PDF → isi database_reviu + database_dokpil ──────────────
            Dim kodeTender As String: kodeTender = CStr(item(6))
            Dim kodePokja2 As String: kodePokja2 = CStr(item(0))
            Dim bidang2 As String: bidang2 = CStr(item(16))
            ParsaDanIsiDariPDF kodeTender, kodePokja2, bidang2
            Exit For
        End If
    Next i
End Sub


' ============================================================
' PARSE PDF: Cari Draft_Pokja_XXX.pdf → panggil Python → isi sheet
' ============================================================
Public Sub ParsaDanIsiDariPDF(kodeTender As String, kodePokja As String, Optional bidang As String = "")
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

    ' Isi database_reviu
    IsiDatabaseReviu jsonTeks

    ' Isi database_dokpil
    IsiDatabaseDokpil jsonTeks

    ' Tampilkan sheet _HasilParse (hanya 1x MsgBox)
    TampilkanHasilParse jsonTeks, folderWorkbook
End Sub


' ============================================================
' ISI database_reviu dari JSON
' ============================================================
Private Sub IsiDatabaseReviu(jsonTeks As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("database_reviu")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

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
        ' Hanya isi jika terisi (tidak timpa perlu_keputusan / tidak_ada)
        If status = "terisi" And nilai <> "" Then
            ws.Range(cells(i)).Value = nilai
        End If
    Next i

    ' E32, E33, E34 (string panjang, bisa ada newline)
    Dim blok32 As String: blok32 = ExtractJSONBlok(jsonTeks, "reviu", "E32")
    If ExtractJSONVal(blok32, "status") = "terisi" Then
        ws.Range("E32").Value = ExtractJSONVal(blok32, "nilai")
    End If
    Dim blok33 As String: blok33 = ExtractJSONBlok(jsonTeks, "reviu", "E33")
    If ExtractJSONVal(blok33, "status") = "terisi" Then
        ' Ganti \n literal dengan newline Excel
        Dim uraian As String: uraian = ExtractJSONVal(blok33, "nilai")
        uraian = Join(Split(uraian, "\n"), Chr(10))
        ws.Range("E33").Value = uraian
    End If
    Dim blok34 As String: blok34 = ExtractJSONBlok(jsonTeks, "reviu", "E34")
    If ExtractJSONVal(blok34, "status") = "terisi" Then
        ws.Range("E34").Value = ExtractJSONVal(blok34, "nilai")
    End If
End Sub


' ============================================================
' ISI database_dokpil dari JSON
' ============================================================
Private Sub IsiDatabaseDokpil(jsonTeks As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("database_dokpil")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

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
            ws.Range(dokpilCells(i)).Value = nilai
        ElseIf status = "tidak_ada" Then
            ws.Range(dokpilCells(i)).Value = ""
        End If
    Next i

    ' E16: Cara Pembayaran
    Dim blok16 As String: blok16 = ExtractJSONBlok(jsonTeks, "dokpil", "E16")
    If ExtractJSONVal(blok16, "status") = "terisi" Then
        ws.Range("E16").Value = ExtractJSONVal(blok16, "nilai")
    End If
End Sub


' ============================================================
' TAMPILKAN _HasilParse — sheet checklist, MsgBox hanya 1x
' ============================================================
Public Sub TampilkanHasilParse(jsonTeks As String, folderWb As String)
    ' Buat/ambil sheet _HasilParse
    Dim wsHP As Worksheet
    On Error Resume Next
    Set wsHP = ThisWorkbook.Sheets("_HasilParse")
    On Error GoTo 0
    If wsHP Is Nothing Then
        Set wsHP = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsHP.Name = "_HasilParse"
    Else
        ' Pindah ke posisi paling kiri jika belum di sana
        If wsHP.Index <> 1 Then
            wsHP.Move Before:=ThisWorkbook.Sheets(1)
        End If
    End If
    wsHP.Visible = xlSheetVisible
    wsHP.Cells.Clear

    ' Header
    With wsHP
        .Range("A1").Value = "Field"
        .Range("B1").Value = "Status"
        .Range("C1").Value = "Nilai / Keterangan"
        .Range("D1").Value = "Navigasi"
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Interior.Color = RGB(68, 114, 196)
        .Range("A1:D1").Font.Color = RGB(255, 255, 255)
    End With

    Dim baris As Integer: baris = 2
    Dim jmlTerisi As Integer: jmlTerisi = 0
    Dim jmlKosong As Integer: jmlKosong = 0
    Dim jmlKeputusan As Integer: jmlKeputusan = 0
    Dim jmlTidakAda As Integer: jmlTidakAda = 0

    ' ─── Bagian database_reviu ────────────────────────────────────────────────
    TulisHeaderBagian wsHP, baris, "DATABASE REVIU"
    baris = baris + 1

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
        TulisBarisHasil wsHP, baris, blok, revKeys(k), "database_reviu", _
                        jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda
        baris = baris + 1
    Next k
    ' E34
    Dim blok34 As String: blok34 = ExtractJSONBlok(jsonTeks, "reviu", "E34")
    TulisBarisHasil wsHP, baris, blok34, "E34", "database_reviu", _
                    jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda
    baris = baris + 1

    ' ─── Bagian database_dokpil ───────────────────────────────────────────────
    baris = baris + 1
    TulisHeaderBagian wsHP, baris, "DATABASE DOKPIL"
    baris = baris + 1

    Dim dpKeys(10) As String
    dpKeys(0) = "E6":  dpKeys(1) = "E7":  dpKeys(2) = "E8"
    dpKeys(3) = "E9":  dpKeys(4) = "E10": dpKeys(5) = "E11"
    dpKeys(6) = "E12": dpKeys(7) = "E13": dpKeys(8) = "E14"
    dpKeys(9) = "E15": dpKeys(10) = "E16"

    Dim j As Integer
    For j = 0 To 10
        Dim blokD As String: blokD = ExtractJSONBlok(jsonTeks, "dokpil", dpKeys(j))
        TulisBarisHasil wsHP, baris, blokD, dpKeys(j), "database_dokpil", _
                        jmlTerisi, jmlKosong, jmlKeputusan, jmlTidakAda
        baris = baris + 1
    Next j

    ' Format kolom
    wsHP.Columns("A").ColumnWidth = 35
    wsHP.Columns("B").ColumnWidth = 20
    wsHP.Columns("C").ColumnWidth = 60
    wsHP.Columns("D").ColumnWidth = 15
    wsHP.Range("C2:C" & baris).WrapText = False

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
    Set wsHP = ThisWorkbook.Sheets("_HasilParse")
    On Error GoTo 0
    If wsHP Is Nothing Then
        Set wsHP = ThisWorkbook.Sheets.Add(Before:=ThisWorkbook.Sheets(1))
        wsHP.Name = "_HasilParse"
    Else
        If wsHP.Index <> 1 Then wsHP.Move Before:=ThisWorkbook.Sheets(1)
    End If
    wsHP.Visible = xlSheetVisible
    wsHP.Cells.Clear
    wsHP.Range("A1").Value = "INFO"
    wsHP.Range("B1").Value = "File Draft_Pokja_XXX.pdf tidak ditemukan di folder paket."
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
        .Cells(baris, 3).Value = Left(nilai, 120)

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
' HTTP: GET ke Supabase REST API
' ============================================================
Private Function FetchSupabase() As String
    Dim url As String
    url = SB_URL & "/rest/v1/" & SB_TABLE & "?select=" & SB_SELECT & "&order=kode_pokja.asc"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo ErrHTTP
    http.Open "GET", url, False
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

    Do
        braceStart = InStr(pos, json, "{")
        If braceStart = 0 Then Exit Do
        braceEnd = InStr(braceStart, json, "}")
        If braceEnd = 0 Then Exit Do

        Dim obj As String
        obj = Mid(json, braceStart, braceEnd - braceStart + 1)

        Dim item(16) As Variant
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
