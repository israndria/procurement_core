Attribute VB_Name = "ModDraftPaketPL"
' ============================================================
' ModDraftPaketPL - Load data Draft Paket PL dari Supabase
' ============================================================
' Bisnis proses PL sepenuhnya terpisah dari tender (PK).
' Tidak ada dependency ke ModDraftPaket, ModInputBA, atau ModKKEvaluasi.
'
' Flow:
'   1. MuatDraftPaketPL()   -> GET Supabase draft_paket_pl -> dropdown PL_F1
'   2. PilihDraftPaketPL()  -> dipanggil tombol "Pilih" -> autofill @ Master Data col C
'
' @ Master Data layout (BAPLJKK):
'   R3  Kode Paket       R14 Nilai HPS         R25 Kode Rekening (MAK)
'   R4  Kode RUP         R15 Jangka Waktu      R26 No BA Reviu
'   R5  Nama Pekerjaan   R16 Sumber Dana       R27 Alamat PP (lookup master_dinas)
'   R6  Nama SKPD        R17 Lokasi            R29 SBU Baru
'   R7  Sub Kegiatan     R18 Jenis Kontrak     R30 SBU Lama
'   R8  Nama PPK         R19 Uraian Singkat    R32-R49 Personil 1-6 (stride=3)
'   R9  NIP PPK          R20 Nomor Dokpil        Rn   Jabatan
'   R10 No SK PPK        R21 Tanggal Dokpil      Rn+1 Pengalaman
'   R11 Nama PP          R22 No Undangan PL      Rn+2 Sertifikat
'   R12 NIP PP           R23 Tahun Anggaran    R51 Nama Peserta
'   R13 Pagu Anggaran    R24 Tanggal Undangan  R52 NPWP Peserta

' Konfigurasi Supabase
Private Const SB_URL As String = "%%SUPABASE_URL%%"
Private Const SB_KEY As String = "%%SUPABASE_KEY%%"

' Word Merge — pattern nama file + sheet name
Private Const WM_SHEET_BA     As String = "satu_data"
Private Const WM_SHEET_REVIU  As String = "list_reviu"
Private Const WM_SHEET_DOKPIL As String = "list_dokpil"
Private Const WM_PAT_BA       As String = "1. Full Dokumen BA PLJKK -"
Private Const WM_PAT_REVIU    As String = "2. Isi Reviu PLJKK -"
Private Const WM_PAT_DOKPIL   As String = "3. Dokpil Full PLJKK -"

' Sheet & Cell selector
Private Const MD_SHEET As String = "@ Master Data"
Private Const CELL_SELECTOR As String = "F1"

' Kolom yang di-fetch dari draft_paket_pl
' index: 0=kode_paket, 1=nama_paket, 2=satker, 3=kode_rup, 4=nilai_hps,
'        5=jenis_pl, 6=jenis_kontrak, 7=status, 8=nama_ppk, 9=nip_ppk,
'        10=no_sk_ppk, 11=nilai_pagu, 12=jangka_waktu, 13=sumber_anggaran,
'        14=lokasi, 15=sbu_baru, 16=sbu_lama, 17=jabatan_teknis, 18=skk_teknis,
'        19=jabatan_k3, 20=skk_k3, 21=dpa_nomor, 22=sub_kegiatan, 23=nama_file_uraian,
'        24=mak, 25=nama_penyedia, 26=npwp_penyedia, 27=personil_json, 28=tgl_dokpil, 29=nomor_dokpil, 30=kode_unik
Private Const SB_SELECT As String = "kode_paket,nama_paket,satker,kode_rup,nilai_hps,jenis_pl,jenis_kontrak,status,nama_ppk,nip_ppk,no_sk_ppk,nilai_pagu,jangka_waktu,sumber_anggaran,lokasi,sbu_baru,sbu_lama,jabatan_teknis,skk_teknis,jabatan_k3,skk_k3,dpa_nomor,sub_kegiatan,nama_file_uraian,mak,nama_penyedia,npwp_penyedia,personil_json,tgl_dokpil,nomor_dokpil,kode_unik"

' Row constants di @ Master Data (kolom C = nilai)
Private Const PLR_KODE_PAKET      As Integer = 3
Private Const PLR_KODE_RUP        As Integer = 4
Private Const PLR_NAMA_PEKERJAAN  As Integer = 5
Private Const PLR_NAMA_SKPD       As Integer = 6
Private Const PLR_SUB_KEGIATAN    As Integer = 7
Private Const PLR_NAMA_PPK        As Integer = 8
Private Const PLR_NIP_PPK         As Integer = 9
Private Const PLR_NO_SK_PPK       As Integer = 10
Private Const PLR_NAMA_PP         As Integer = 11  ' hardcoded, tidak ditimpa
Private Const PLR_NIP_PP          As Integer = 12  ' hardcoded, tidak ditimpa
Private Const PLR_PAGU            As Integer = 13
Private Const PLR_HPS             As Integer = 14
Private Const PLR_JANGKA_WAKTU    As Integer = 15
Private Const PLR_SUMBER_DANA     As Integer = 16
Private Const PLR_LOKASI          As Integer = 17
Private Const PLR_JENIS_KONTRAK   As Integer = 18
Private Const PLR_URAIAN_SINGKAT  As Integer = 19
Private Const PLR_NOMOR_DOKPIL    As Integer = 20
Private Const PLR_TANGGAL_DOKPIL  As Integer = 21
Private Const PLR_NO_UNDANGAN     As Integer = 22
Private Const PLR_TAHUN_ANGGARAN  As Integer = 23
Private Const PLR_TGL_UNDANGAN    As Integer = 24
Private Const PLR_KODE_REKENING   As Integer = 25
Private Const PLR_NO_BA_REVIU     As Integer = 26
Private Const PLR_ALAMAT_PP       As Integer = 59  ' lookup master_dinas.alamat_pp_bertugas
Private Const PLR_TELEPON_PP      As Integer = 58  ' lookup master_dinas.telepon_pp
Private Const PLR_SBU_BARU        As Integer = 29
Private Const PLR_SBU_LAMA        As Integer = 30
Private Const PLR_PERSONIL_BASE   As Integer = 32  ' R32-R49: stride=3 (jabatan/pengalaman/sertifikat)
Private Const PLR_PERSONIL_STRIDE As Integer = 3   ' 3 row per personil
Private Const PLR_NAMA_PESERTA    As Integer = 51  ' geser dari 45 (6 personil * 3 row = 18 row, 32+18=50, +1=51)
Private Const PLR_NPWP_PESERTA    As Integer = 52

' Cache data in-memory
Private m_DataCache As Collection
Private m_LastLoad As Date


' ============================================================
' FUNGSI UTAMA: Load dari Supabase, buat dropdown
' ============================================================
Public Sub MuatDraftPaketPL()
    Dim wsMD As Worksheet
    On Error GoTo ErrHandler
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)

    ' Fetch JSON dari Supabase
    Dim json As String
    json = FetchSupabasePL()
    If json = "" Then
        MsgBox "Gagal mengambil data PL dari Supabase.", vbExclamation, "Muat Draft PL"
        Exit Sub
    End If

    ' Parse JSON -> Collection
    Set m_DataCache = ParsePLJSON(json)
    m_LastLoad = Now

    If m_DataCache.Count = 0 Then
        MsgBox "Tidak ada data Draft Paket PL di database." & vbCrLf & _
               "Jalankan 'Serap Paket PL' di Asisten Pokja terlebih dahulu.", _
               vbInformation, "Muat Draft PL"
        Exit Sub
    End If

    ' Tulis label ke sheet tersembunyi "_DraftPaketPLList"
    Dim wsHidden As Worksheet
    On Error Resume Next
    Set wsHidden = ThisWorkbook.Sheets("_DraftPaketPLList")
    On Error GoTo ErrHandler
    If wsHidden Is Nothing Then
        Set wsHidden = ThisWorkbook.Sheets.Add
        wsHidden.Name = "_DraftPaketPLList"
    End If
    wsHidden.Visible = xlSheetVeryHidden
    wsHidden.Columns(1).ClearContents

    Dim i As Integer
    For i = 1 To m_DataCache.Count
        Dim item As Variant
        item = m_DataCache(i)
        ' item(0)=kode_paket, item(1)=nama_paket, item(2)=satker, item(5)=jenis_pl
        Dim label As String
        label = CStr(item(5)) & " - " & Left(Trim(CStr(item(1))), 55)
        wsHidden.Cells(i, 1).Value = label
    Next i

    ' Named Range "DaftarDraftPaketPL"
    Dim rngList As Range
    Set rngList = wsHidden.Range(wsHidden.Cells(1, 1), wsHidden.Cells(m_DataCache.Count, 1))
    On Error Resume Next
    ThisWorkbook.Names("DaftarDraftPaketPL").Delete
    On Error GoTo ErrHandler
    ThisWorkbook.Names.Add Name:="DaftarDraftPaketPL", RefersTo:=rngList

    ' Dropdown validation di @ Master Data F1
    With wsMD.Range(CELL_SELECTOR).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, _
             Operator:=xlBetween, Formula1:="=DaftarDraftPaketPL"
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowInput = False
        .ShowError = False
    End With

    ' Restore pilihan yang sudah ada sebelumnya
    Dim savedLabel As String
    savedLabel = Trim(CStr(wsMD.Range(CELL_SELECTOR).Value))
    If savedLabel <> "" Then
        Application.EnableEvents = False
        wsMD.Range(CELL_SELECTOR).Value = savedLabel
        Application.EnableEvents = True
    End If

    MsgBox m_DataCache.Count & " paket PL berhasil dimuat." & vbCrLf & _
           "Pilih paket dari dropdown " & CELL_SELECTOR & ", lalu klik 'Isi Data PL'.", _
           vbInformation, "Draft Paket PL"
    Exit Sub

ErrHandler:
    MsgBox "Error MuatDraftPaketPL: " & Err.Description, vbCritical, "Muat Draft PL"
End Sub


' ============================================================
' ISI DATA: Dipanggil tombol "Isi Data PL" di @ Master Data
' ============================================================
Public Sub IsiDataPL()
    Dim wsMD As Worksheet
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)

    Dim selVal As String
    selVal = Trim(CStr(wsMD.Range(CELL_SELECTOR).Value))
    If selVal = "" Then
        MsgBox "Pilih paket dari dropdown " & CELL_SELECTOR & " terlebih dahulu.", _
               vbExclamation, "Belum Ada Pilihan"
        Exit Sub
    End If

    ' Auto-load jika cache kosong
    If m_DataCache Is Nothing Then MuatDraftPaketPL
    If m_DataCache Is Nothing Then Exit Sub
    If m_DataCache.Count = 0 Then MuatDraftPaketPL
    If m_DataCache.Count = 0 Then Exit Sub

    ' Cari item yang label-nya cocok
    Dim i As Integer
    For i = 1 To m_DataCache.Count
        Dim item As Variant
        item = m_DataCache(i)
        Dim label As String
        label = CStr(item(5)) & " - " & Left(Trim(CStr(item(1))), 55)
        If label = selVal Then
            IsiMasterDataPL wsMD, item
            ' Auto diff highlight setelah isi data
            DiffHighlightPL CStr(item(0))
            Exit For
        End If
    Next i
End Sub


' ============================================================
' AUTOFILL: Tulis nilai dari item ke @ Master Data kolom C
' ============================================================
Private Sub IsiMasterDataPL(wsMD As Worksheet, item As Variant)
    ' item index sesuai SB_SELECT:
    ' 0=kode_paket, 1=nama_paket, 2=satker, 3=kode_rup, 4=nilai_hps,
    ' 5=jenis_pl, 6=jenis_kontrak, 7=status, 8=nama_ppk, 9=nip_ppk,
    ' 10=no_sk_ppk, 11=pagu_anggaran, 12=jangka_waktu, 13=sumber_dana,
    ' 14=lokasi, 15=sbu_baru, 16=sbu_lama, 17=jabatan_teknis, 18=skk_teknis,
    ' 19=jabatan_k3, 20=skk_k3, 21=dpa_nomor, 22=sub_kegiatan, 23=nama_file_uraian,
    ' 24=mak, 25=nama_penyedia, 26=npwp_penyedia

    With wsMD
        ' ── INPUT DATA dari Supabase ──────────────────────────────────────
        .Cells(PLR_KODE_PAKET, 3).NumberFormat = "@"
        .Cells(PLR_KODE_PAKET, 3).Value = CStr(item(0))     ' kode_paket
        .Cells(PLR_KODE_RUP, 3).Value   = CStr(item(3))     ' kode_rup
        .Cells(PLR_NAMA_PEKERJAAN, 3).Value = CStr(item(1)) ' nama_paket
        .Cells(PLR_NAMA_SKPD, 3).Value  = CStr(item(2))     ' satker
        .Cells(PLR_SUB_KEGIATAN, 3).Value = CStr(item(22))  ' sub_kegiatan
        .Cells(PLR_NAMA_PPK, 3).Value   = CStr(item(8))     ' nama_ppk

        ' Lookup NIP + SK dari master_kpa berdasarkan nama PPK
        Dim kpaData As String: kpaData = LookupKPA(CStr(item(8)))
        Dim kpaNip As String: kpaNip = ExtractJSONValPL(kpaData, "nip")
        Dim kpaSK As String:  kpaSK  = ExtractJSONValPL(kpaData, "nomor_sk")
        Dim kpaTglSK As String: kpaTglSK = ExtractJSONValPL(kpaData, "tanggal_sk")

        ' NIP PPK: prefer master_kpa, fallback Supabase draft
        If kpaNip <> "" Then
            .Cells(PLR_NIP_PPK, 3).NumberFormat = "@"
            .Cells(PLR_NIP_PPK, 3).Value = kpaNip
        ElseIf CStr(item(9)) <> "" Then
            .Cells(PLR_NIP_PPK, 3).NumberFormat = "@"
            .Cells(PLR_NIP_PPK, 3).Value = CStr(item(9))
        End If

        ' No SK PPK: prefer master_kpa
        If kpaSK <> "" Then
            Dim skFmt As String: skFmt = kpaSK
            If kpaTglSK <> "" Then
                ' Format tanggal "YYYY-MM-DD" -> "DD Bulan YYYY"
                Dim tglFmt As String: tglFmt = FormatTanggalIndo(kpaTglSK)
                skFmt = skFmt & ", Tanggal " & tglFmt
            End If
            .Cells(PLR_NO_SK_PPK, 3).Value = skFmt
        ElseIf CStr(item(10)) <> "" Then
            .Cells(PLR_NO_SK_PPK, 3).Value = CStr(item(10))
        End If

        ' Nama PP + NIP PP: TIDAK ditimpa jika sudah ada (hardcoded)
        ' Biarkan nilai default di R11/R12 tetap

        ' Nilai finansial
        Dim paguVal As String: paguVal = CStr(item(11))
        If Left(paguVal, 3) <> "Rp." Then paguVal = "Rp. " & paguVal
        .Cells(PLR_PAGU, 3).Value = paguVal
        Dim hpsVal As String: hpsVal = CStr(item(4))
        If Left(hpsVal, 3) <> "Rp." Then hpsVal = "Rp. " & hpsVal
        .Cells(PLR_HPS, 3).Value = hpsVal
        .Cells(PLR_SUMBER_DANA, 3).Value = CStr(item(13))   ' sumber_dana
        .Cells(PLR_LOKASI, 3).Value      = CStr(item(14))   ' lokasi

        ' Jangka waktu: extract angka pertama saja (handles "30 (Tiga Puluh) Hari Kalender")
        Dim jw As String: jw = Trim(CStr(item(12)))
        Dim jwAngka As String: jwAngka = ""
        Dim jwk As Long
        For jwk = 1 To Len(jw)
            Dim jwc As String: jwc = Mid(jw, jwk, 1)
            If jwc >= "0" And jwc <= "9" Then
                jwAngka = jwAngka & jwc
            ElseIf jwAngka <> "" Then
                Exit For
            End If
        Next jwk
        If jwAngka <> "" Then
            .Cells(PLR_JANGKA_WAKTU, 3).Value = CLng(jwAngka)
        ElseIf jw <> "" Then
            .Cells(PLR_JANGKA_WAKTU, 3).Value = jw
        End If

        ' Kontrak & dokpil
        .Cells(PLR_JENIS_KONTRAK, 3).Value = CStr(item(6))  ' jenis_kontrak

        ' Tahun anggaran: dari sumber_anggaran "APBD YYYY", fallback Year(Now)
        Dim tahun As String: tahun = ""
        Dim srcAng As String: srcAng = Trim(CStr(item(13)))
        If srcAng <> "" Then
            Dim posSpasi As Long: posSpasi = InStrRev(srcAng, " ")
            If posSpasi > 0 Then
                Dim tStr As String: tStr = Mid(srcAng, posSpasi + 1)
                If Len(tStr) = 4 And IsNumeric(tStr) Then tahun = tStr
            End If
        End If
        If tahun = "" Then tahun = CStr(Year(Now))
        .Cells(PLR_TAHUN_ANGGARAN, 3).Value = tahun

        ' Kode Rekening (MAK) — sumber: inbox PL field MAK
        ' Fallback: dpa_nomor (legacy) jika mak kosong
        Dim makVal As String: makVal = CStr(item(24))       ' mak
        If makVal = "" Then makVal = CStr(item(21))         ' fallback dpa_nomor
        .Cells(PLR_KODE_REKENING, 3).NumberFormat = "@"
        .Cells(PLR_KODE_REKENING, 3).Value = makVal

        ' Tanggal Dokpil dari Supabase tgl_dokpil (format YYYY-MM-DD) → string Indonesia
        Dim tglDokpilStr As String: tglDokpilStr = CStr(item(28))
        If tglDokpilStr <> "" And tglDokpilStr <> "null" Then
            Dim tglParts() As String: tglParts = Split(Left(tglDokpilStr, 10), "-")
            If UBound(tglParts) >= 2 Then
                Dim tglYr As Integer: tglYr = CInt(tglParts(0))
                Dim tglMo As Integer: tglMo = CInt(tglParts(1))
                Dim tglDy As Integer: tglDy = CInt(tglParts(2))
                Dim arrBulan(12) As String
                arrBulan(1)  = "Januari":  arrBulan(2)  = "Februari": arrBulan(3)  = "Maret"
                arrBulan(4)  = "April":    arrBulan(5)  = "Mei":      arrBulan(6)  = "Juni"
                arrBulan(7)  = "Juli":     arrBulan(8)  = "Agustus":  arrBulan(9)  = "September"
                arrBulan(10) = "Oktober":  arrBulan(11) = "November": arrBulan(12) = "Desember"
                .Cells(PLR_TANGGAL_DOKPIL, 3).NumberFormat = "@"
                .Cells(PLR_TANGGAL_DOKPIL, 3).Value = tglDy & " " & arrBulan(tglMo) & " " & tglYr
            End If
        End If

        ' Nama + NPWP Peserta — sumber: parse Draft_PL PDF
        If CStr(item(25)) <> "" Then
            .Cells(PLR_NAMA_PESERTA, 3).Value = CStr(item(25))  ' nama_penyedia
        End If
        If CStr(item(26)) <> "" Then
            .Cells(PLR_NPWP_PESERTA, 3).NumberFormat = "@"
            .Cells(PLR_NPWP_PESERTA, 3).Value = CStr(item(26))  ' npwp_penyedia
        End If

        ' ── URAIAN SINGKAT (hardcode template) ───────────────────────────
        Dim namaUraian As String: namaUraian = Trim(CStr(item(23)))
        If namaUraian = "" Then namaUraian = "Uraian Singkat Pekerjaan"
        ' Buang ekstensi .pdf di display
        If LCase(Right(namaUraian, 4)) = ".pdf" Then
            namaUraian = Left(namaUraian, Len(namaUraian) - 4)
        End If
        .Cells(PLR_URAIAN_SINGKAT, 3).Value = _
            "Mengerjakan " & CStr(item(1)) & " sesuai dengan KAK/Dokumen " & namaUraian

        ' ── NOMOR DOKPIL: 000.3.3/01/PL/PP-NN/KodeUnik/SKPD/Tahun ────────
        ' PP-NN dari angka terakhir di nama_paket (Paket 2 -> 02, Paket 12 -> 12)
        Dim npaket As String: npaket = CStr(item(1))
        Dim numStr As String: numStr = ""
        Dim ii As Long
        For ii = Len(npaket) To 1 Step -1
            Dim chr As String: chr = Mid(npaket, ii, 1)
            If chr >= "0" And chr <= "9" Then
                numStr = chr & numStr
            ElseIf numStr <> "" Then
                Exit For  ' sudah kumpul digit, ketemu non-digit -> stop
            End If
        Next ii
        If numStr = "" Then numStr = "1"
        If Len(numStr) = 1 Then numStr = "0" & numStr

        ' Singkatan SKPD dari master_dinas via lookup
        Dim singkatan As String: singkatan = LookupSingkatanDinas(CStr(item(2)))
        If singkatan = "" Then singkatan = "DPUPR"

        ' Kode unik: prefer Supabase item(30), fallback G2
        Dim koUnik As String: koUnik = CStr(item(30))
        If koUnik = "" Or koUnik = "null" Then koUnik = CStr(wsMD.Range("G2").Value)
        If koUnik = "" Or koUnik = "null" Then koUnik = "KodeUnik"
        wsMD.Range("G2").Value = koUnik

        ' Tahun: ambil dari tahun anggaran yang sudah diisi
        Dim tahunDokpil As String: tahunDokpil = CStr(wsMD.Cells(PLR_TAHUN_ANGGARAN, 3).Value)
        If tahunDokpil = "" Then tahunDokpil = CStr(Year(Now))

        ' Nomor Dokpil: prefer Supabase (sudah di-upload via Tab 4), fallback compose
        Dim noDokpilSupa As String: noDokpilSupa = CStr(item(29))
        If noDokpilSupa <> "" And noDokpilSupa <> "null" Then
            .Cells(PLR_NOMOR_DOKPIL, 3).Value = noDokpilSupa
        Else
            .Cells(PLR_NOMOR_DOKPIL, 3).Value = _
                "000.3.3/01/PL/PP-" & numStr & "/" & koUnik & "/" & singkatan & "/" & tahunDokpil
        End If

        ' Auto-generate No Undangan dari Nomor Dokpil (ganti /01/PL/ -> /02/PL/)
        Dim noUndangan As String
        noUndangan = Replace(.Cells(PLR_NOMOR_DOKPIL, 3).Value, "/01/PL/", "/02/PL/")
        .Cells(PLR_NO_UNDANGAN, 3).Value = noUndangan

        ' ── NOMOR BA REVIU: 000.3.3/PP{NN}/02/SKPD/Reviu-KodeUnik/Tahun ────
        .Cells(PLR_NO_BA_REVIU, 3).Value = _
            "000.3.3/PP" & numStr & "/02/" & singkatan & "/Reviu-" & koUnik & "/" & tahunDokpil

        ' ── ALAMAT PP: lookup master_dinas.alamat_pp_bertugas via satker ───
        Dim alamatPP As String: alamatPP = LookupAlamatPP(CStr(item(2)))
        If alamatPP <> "" And alamatPP <> "null" Then
            .Cells(PLR_ALAMAT_PP, 3).Value = alamatPP
        End If

        ' ── TELEPON PP: lookup master_dinas.telepon_pp via satker ───────
        Dim telPP As String: telPP = LookupTeleponPP(CStr(item(2)))
        If telPP <> "" And telPP <> "null" Then
            .Cells(PLR_TELEPON_PP, 3).Value = telPP
        End If

        ' ── SBU ──────────────────────────────────────────────────────────
        If CStr(item(15)) <> "" Then
            .Cells(PLR_SBU_BARU, 3).Value = CStr(item(15))  ' sbu_baru
        End If
        If CStr(item(16)) <> "" Then
            .Cells(PLR_SBU_LAMA, 3).Value = CStr(item(16))  ' sbu_lama
        End If

        ' ── PERSONIL (R32-R49): jabatan/pengalaman/sertifikat P1-P6, stride=3 ──
        Dim personilJsonStr As String: personilJsonStr = CStr(item(27))
        If personilJsonStr <> "" And personilJsonStr <> "null" Then
            Dim personilArr() As String
            personilArr = ParsePersonilArrayPL(personilJsonStr)
            Dim nP As Long: nP = (UBound(personilArr) + 1) \ PLR_PERSONIL_STRIDE
            Dim iP As Long
            For iP = 0 To nP - 1
                If iP >= 6 Then Exit For
                Dim baseRow As Long: baseRow = PLR_PERSONIL_BASE + iP * PLR_PERSONIL_STRIDE
                .Cells(baseRow, 3).Value     = personilArr(iP * PLR_PERSONIL_STRIDE)       ' jabatan
                .Cells(baseRow + 1, 3).Value = personilArr(iP * PLR_PERSONIL_STRIDE + 1)   ' pengalaman
                .Cells(baseRow + 2, 3).Value = personilArr(iP * PLR_PERSONIL_STRIDE + 2)   ' sertifikat
            Next iP
        End If
    End With

    ' Konfirmasi
    Dim namaP As String: namaP = Left(Trim(CStr(item(1))), 50)
    MsgBox "Data PL berhasil diisi:" & vbCrLf & vbCrLf & _
           "  Paket  : " & namaP & vbCrLf & _
           "  Jenis  : " & CStr(item(5)) & vbCrLf & _
           "  HPS    : Rp " & Format(CDblSafe(CStr(item(4))), "#,##0") & vbCrLf & _
           "  SBU    : " & CStr(item(15)) & vbCrLf & vbCrLf & _
           "Cek sheet @ Master Data, lalu isi field yang masih kosong " & vbCrLf & _
           "(Nomor Dokpil, Tanggal, No Undangan, Uraian Singkat).", _
           vbInformation, "Data PL Terisi"
End Sub


' ============================================================
' HTTP: GET draft_paket_pl dari Supabase
' ============================================================
Private Function FetchSupabasePL() As String
    Dim url As String
    url = SB_URL & "/rest/v1/draft_paket_pl" & _
          "?select=" & SB_SELECT & _
          "&status=neq.selesai" & _
          "&order=diambil_pada.desc"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo ErrHTTP
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 15000, 15000
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send

    If http.Status = 200 Then
        FetchSupabasePL = http.ResponseText
    Else
        MsgBox "HTTP Error " & http.Status & ": " & http.StatusText & vbCrLf & _
               "URL: " & url, vbExclamation, "FetchSupabasePL"
        FetchSupabasePL = ""
    End If
    Exit Function

ErrHTTP:
    MsgBox "WinHTTP Error: " & Err.Number & " - " & Err.Description, vbCritical, "FetchSupabasePL"
    FetchSupabasePL = ""
End Function


' ============================================================
' PARSER: JSON array -> Collection of Variant arrays
' ============================================================
Private Function ParsePLJSON(json As String) As Collection
    Dim col As New Collection

    Dim pos As Long: pos = 1
    Do
        Dim braceStart As Long: braceStart = InStr(pos, json, "{")
        If braceStart = 0 Then Exit Do

        ' Cari matching }
        Dim depth As Long: depth = 0
        Dim inQuote As Boolean: inQuote = False
        Dim scanPos As Long
        Dim braceEnd As Long: braceEnd = 0
        For scanPos = braceStart To Len(json)
            Dim ch As String: ch = Mid(json, scanPos, 1)
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

        Dim obj As String: obj = Mid(json, braceStart, braceEnd - braceStart + 1)

        Dim item(30) As Variant
        item(0)  = ExtractJSONValPL(obj, "kode_paket")
        item(1)  = ExtractJSONValPL(obj, "nama_paket")
        item(2)  = ExtractJSONValPL(obj, "satker")
        item(3)  = ExtractJSONValPL(obj, "kode_rup")
        item(4)  = ExtractJSONValPL(obj, "nilai_hps")
        item(5)  = ExtractJSONValPL(obj, "jenis_pl")
        item(6)  = ExtractJSONValPL(obj, "jenis_kontrak")
        item(7)  = ExtractJSONValPL(obj, "status")
        item(8)  = ExtractJSONValPL(obj, "nama_ppk")
        item(9)  = ExtractJSONValPL(obj, "nip_ppk")
        item(10) = ExtractJSONValPL(obj, "no_sk_ppk")
        item(11) = ExtractJSONValPL(obj, "nilai_pagu")
        item(12) = ExtractJSONValPL(obj, "jangka_waktu")
        item(13) = ExtractJSONValPL(obj, "sumber_anggaran")
        item(14) = ExtractJSONValPL(obj, "lokasi")
        item(15) = ExtractJSONValPL(obj, "sbu_baru")
        item(16) = ExtractJSONValPL(obj, "sbu_lama")
        item(17) = ExtractJSONValPL(obj, "jabatan_teknis")
        item(18) = ExtractJSONValPL(obj, "skk_teknis")
        item(19) = ExtractJSONValPL(obj, "jabatan_k3")
        item(20) = ExtractJSONValPL(obj, "skk_k3")
        item(21) = ExtractJSONValPL(obj, "dpa_nomor")
        item(22) = ExtractJSONValPL(obj, "sub_kegiatan")
        item(23) = ExtractJSONValPL(obj, "nama_file_uraian")
        item(24) = ExtractJSONValPL(obj, "mak")
        item(25) = ExtractJSONValPL(obj, "nama_penyedia")
        item(26) = ExtractJSONValPL(obj, "npwp_penyedia")
        item(27) = ExtractJSONValPL(obj, "personil_json")
        item(28) = ExtractJSONValPL(obj, "tgl_dokpil")
        item(29) = ExtractJSONValPL(obj, "nomor_dokpil")
        item(30) = ExtractJSONValPL(obj, "kode_unik")

        col.Add item
        pos = braceEnd + 1
    Loop

    Set ParsePLJSON = col
End Function


' ============================================================
' WORD MERGE — Buka / PDF dokumen PLJKK
' ============================================================

Public Sub BukaBAPlJkk()
    RunMergePL "buka", WM_PAT_BA, WM_SHEET_BA
End Sub

Public Sub BukaReviuPlJkk()
    RunMergePL "buka", WM_PAT_REVIU, WM_SHEET_REVIU
End Sub

Public Sub BukaDokpilPlJkk()
    RunMergePL "buka", WM_PAT_DOKPIL, WM_SHEET_DOKPIL
End Sub

Public Sub CetakDokpilPlJkkPDF()
    RunMergePL "pdf_dokpil", WM_PAT_DOKPIL, WM_SHEET_DOKPIL
End Sub

Public Sub CetakReviuPlJkkPDF()
    RunMergePL "pdf_all", WM_PAT_REVIU, WM_SHEET_REVIU
End Sub

Public Sub CetakBAReviuPLPDF()
    ' Cetak BA Reviu PL halaman 1-3 ke PDF
    Dim wordFile As String
    wordFile = FindWordFilePL(WM_PAT_BA)
    If wordFile = "" Then Exit Sub

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & wordFile
    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim scriptDir As String
    scriptDir = ScriptDirPL()
    If scriptDir = "" Then
        MsgBox "Python tidak ditemukan.", vbCritical
        Exit Sub
    End If
    Dim pyExe As String: pyExe = scriptDir & "\python\python.exe"

    Dim outMode As String, printerName As String
    outMode = ChooseOutputModePL(printerName)
    If outMode = "" Then Exit Sub

    Dim kodePkt As String
    kodePkt = CStr(ThisWorkbook.Sheets("@ Master Data").Cells(3, 3).Value)
    If kodePkt = "" Then kodePkt = "000"

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim cmd As String

    If outMode = "printer" Then
        cmd = Chr(34) & pyExe & Chr(34) & " " & _
              Chr(34) & scriptDir & "\word_merge.py" & Chr(34) & " printer " & _
              Chr(34) & wordPath & Chr(34) & " " & _
              Chr(34) & ThisWorkbook.FullName & Chr(34) & " " & _
              Chr(34) & WM_SHEET_BA & Chr(34) & " " & _
              Chr(34) & printerName & Chr(34) & " 1 3"
        wsh.Run cmd, 0, False
        Application.StatusBar = "Printing BA Reviu PL ke " & printerName & " ..."
    Else
        cmd = Chr(34) & pyExe & Chr(34) & " " & _
              Chr(34) & scriptDir & "\word_merge.py" & Chr(34) & " pdf_bareviu_pl " & _
              Chr(34) & wordPath & Chr(34) & " " & _
              Chr(34) & ThisWorkbook.FullName & Chr(34) & " " & _
              Chr(34) & WM_SHEET_BA & Chr(34) & " " & _
              Chr(34) & kodePkt & Chr(34)
        wsh.Run cmd, 0, False
        Application.StatusBar = "Membuat PDF BA_REVIU_PL_" & kodePkt & " ..."
    End If

    Set wsh = Nothing
    Application.OnTime Now + TimeValue("00:00:05"), "ResetStatusBar"
End Sub

Private Function ChooseOutputModePL(ByRef outPrinter As String) As String
    Dim choice As VbMsgBoxResult
    choice = MsgBox("Pilih output:" & vbCrLf & vbCrLf & _
                   "YES = Export PDF" & vbCrLf & _
                   "NO = Print ke Printer", _
                   vbYesNoCancel + vbQuestion, "Pilih Output BA Reviu PL")
    If choice = vbCancel Then ChooseOutputModePL = "": Exit Function
    If choice = vbYes Then ChooseOutputModePL = "pdf": Exit Function

    outPrinter = PickPrinterPL()
    If outPrinter = "" Then
        ChooseOutputModePL = ""
    Else
        ChooseOutputModePL = "printer"
    End If
End Function

Private Function PickPrinterPL() As String
    Dim printers() As String
    Dim i As Integer
    Dim count As Integer
    count = 0
    ReDim printers(0)

    On Error Resume Next
    Dim pName As String
    pName = Application.Printers(0).Name
    If Err.Number <> 0 Then
        ' Fallback: pakai printer aktif
        PickPrinterPL = Application.ActivePrinter
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    For i = 0 To Application.Printers.Count - 1
        ReDim Preserve printers(count)
        printers(count) = Application.Printers(i).Name
        count = count + 1
    Next i

    If count = 0 Then PickPrinterPL = "": Exit Function

    Dim listStr As String: listStr = ""
    For i = 0 To count - 1
        listStr = listStr & i + 1 & ". " & printers(i) & vbCrLf
    Next i

    Dim inp As String
    inp = InputBox("Pilih nomor printer:" & vbCrLf & vbCrLf & listStr, "Pilih Printer")
    If inp = "" Then PickPrinterPL = "": Exit Function

    Dim idx As Integer
    idx = CInt(inp) - 1
    If idx >= 0 And idx < count Then
        PickPrinterPL = printers(idx)
    Else
        PickPrinterPL = ""
    End If
End Function

Public Sub RelinkPL()
    ' Relink semua Word PL ke Excel ini via relink_pl.py.
    ' Pakai WScript.Shell synchronous (blocking, no manual polling = no UI freeze loop).
    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim scriptDir As String
    scriptDir = ScriptDirPL()
    If scriptDir = "" Then
        MsgBox "Python tidak ditemukan.", vbCritical
        Exit Sub
    End If
    Dim pyExe As String: pyExe = scriptDir & "\python\python.exe"
    Dim relinkScript As String: relinkScript = scriptDir & "\relink_pl.py"

    Dim outFile As String
    outFile = scriptDir & "\_relink_pl_output.txt"
    On Error Resume Next
    If Dir(outFile) <> "" Then Kill outFile
    On Error GoTo 0

    Dim cmd As String
    cmd = "cmd /c """ & Chr(34) & pyExe & Chr(34) & " " & _
          Chr(34) & relinkScript & Chr(34) & " " & _
          Chr(34) & ThisWorkbook.FullName & Chr(34) & _
          " > " & Chr(34) & outFile & Chr(34) & " 2>&1"""

    ' WScript.Shell sync mode: arg3=True -> tunggu sampai selesai, no manual loop
    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    Dim rc As Long
    rc = wsh.Run(cmd, 0, True)  ' 0=hidden, True=wait
    Set wsh = Nothing

    DoEvents

    Dim output As String
    If Dir(outFile) <> "" Then
        Dim fNum As Integer
        fNum = FreeFile
        Open outFile For Input As #fNum
        Do While Not EOF(fNum)
            Dim lineText As String
            Line Input #fNum, lineText
            output = output & lineText & vbCrLf
        Loop
        Close #fNum
        Kill outFile
    End If

    MsgBox "Relink selesai!" & vbCrLf & vbCrLf & output, vbInformation, "Relink Word PL"
End Sub

' ============================================================
' SYNC DATA DRAFT PL: Baca @ Master Data kolom C → upsert data_snapshot
' ============================================================
Public Sub SyncDataDraftPL()
    Dim wsMD As Worksheet
    On Error GoTo ErrSync
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)

    Dim kodePaket As String
    kodePaket = Trim(CStr(wsMD.Cells(PLR_KODE_PAKET, 3).Value))
    If kodePaket = "" Then
        MsgBox "Kode Paket (C3) kosong. Pilih dan isi data PL terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    Dim sd As String: sd = ScriptDirPL()
    If sd = "" Then MsgBox "Python tidak ditemukan.", vbExclamation: Exit Sub

    Dim inputFile As String:  inputFile  = sd & "\_sync_draft_pl_input.json"
    Dim outputFile As String: outputFile = sd & "\_sync_draft_pl_output.json"

    On Error Resume Next: Kill outputFile: On Error GoTo ErrSync

    ' Pastikan input file tidak terkunci proses sebelumnya
    On Error Resume Next: Kill inputFile: On Error GoTo ErrSync
    Application.Wait Now + TimeSerial(0, 0, 0) + 0.0001  ' yield ~100ms

    Dim snapshot As String: snapshot = BuildSnapshotPL(wsMD)
    Dim payload As String
    payload = "{""kode_paket"":""" & kodePaket & """,""snapshot"":" & snapshot & "}"
    WriteUTF8PL inputFile, payload

    Dim pyExe As String: pyExe = sd & "\python\python.exe"
    Dim cmd As String
    cmd = """" & pyExe & """ """ & sd & "\sync_draft_pl.py"" save"

    Application.StatusBar = "Sync Data Draft PL: menyimpan..."
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True
    Application.StatusBar = False

    If Dir(outputFile) = "" Then
        MsgBox "sync_draft_pl.py tidak menghasilkan output.", vbExclamation
        Exit Sub
    End If

    Dim result As String: result = ReadUTF8PL(outputFile)
    Dim okVal As String:  okVal  = ExtractValPL2(result, "ok")
    If okVal = "true" Then
        MsgBox "Snapshot berhasil disimpan." & vbCrLf & "Kode Paket: " & kodePaket, _
               vbInformation, "Sync Data Draft PL"
        ClearHighlightPL wsMD
    Else
        MsgBox "Gagal sync: " & ExtractValPL2(result, "error"), vbExclamation, "Sync Data Draft PL"
    End If
    Exit Sub
ErrSync:
    Application.StatusBar = False
    MsgBox "Error SyncDataDraftPL: " & Err.Description, vbCritical
End Sub


' ============================================================
' DIFF HIGHLIGHT PL: Load snapshot → highlight sel yang berbeda
' Dipanggil otomatis setelah IsiMasterDataPL selesai
' ============================================================
Public Sub DiffHighlightPL(Optional kodePaketOverride As String = "")
    Dim wsMD As Worksheet
    On Error GoTo ErrDiff
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)

    Dim kodePaket As String
    kodePaket = kodePaketOverride
    If kodePaket = "" Then
        kodePaket = Trim(CStr(wsMD.Cells(PLR_KODE_PAKET, 3).Value))
    End If
    If kodePaket = "" Then Exit Sub

    Dim sd As String: sd = ScriptDirPL()
    If sd = "" Then Exit Sub

    Dim inputFile As String:  inputFile  = sd & "\_sync_draft_pl_input.json"
    Dim outputFile As String: outputFile = sd & "\_sync_draft_pl_output.json"

    On Error Resume Next: Kill outputFile: On Error GoTo ErrDiff

    WriteUTF8PL inputFile, "{""kode_paket"":""" & kodePaket & """}"

    Dim pyExe As String: pyExe = sd & "\python\python.exe"
    Dim cmd As String
    cmd = """" & pyExe & """ """ & sd & "\sync_draft_pl.py"" load"

    Application.StatusBar = "Diff Highlight PL: memuat snapshot..."
    Dim wsh As Object: Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, True
    Application.StatusBar = False

    If Dir(outputFile) = "" Then Exit Sub

    Dim result As String: result = ReadUTF8PL(outputFile)
    If ExtractValPL2(result, "ok") <> "true" Then Exit Sub

    ' snapshot kosong = belum pernah sync
    Dim snapStart As Long: snapStart = InStr(result, """snapshot""")
    If snapStart = 0 Then Exit Sub
    Dim bracePos As Long: bracePos = InStr(snapStart, result, "{")
    If bracePos = 0 Then
        ClearHighlightPL wsMD
        Exit Sub
    End If

    ' Bandingkan semua baris PLR kolom C
    Dim rowList As Variant
    rowList = Array(PLR_KODE_PAKET, PLR_KODE_RUP, PLR_NAMA_PEKERJAAN, PLR_NAMA_SKPD, _
                    PLR_SUB_KEGIATAN, PLR_NAMA_PPK, PLR_NIP_PPK, PLR_NO_SK_PPK, _
                    PLR_NAMA_PP, PLR_NIP_PP, PLR_PAGU, PLR_HPS, PLR_JANGKA_WAKTU, _
                    PLR_SUMBER_DANA, PLR_LOKASI, PLR_JENIS_KONTRAK, PLR_URAIAN_SINGKAT, _
                    PLR_NOMOR_DOKPIL, PLR_TANGGAL_DOKPIL, PLR_NO_UNDANGAN, _
                    PLR_TAHUN_ANGGARAN, PLR_TGL_UNDANGAN, PLR_KODE_REKENING, _
                    PLR_NO_BA_REVIU, PLR_ALAMAT_PP, PLR_SBU_BARU, PLR_SBU_LAMA, _
                    PLR_NAMA_PESERTA, PLR_NPWP_PESERTA)

    Dim ri As Long
    For ri = 0 To UBound(rowList)
        Dim r As Long: r = rowList(ri)
        Dim cellVal As String: cellVal = Trim(CStr(wsMD.Cells(r, 3).Value))
        Dim snapVal As String: snapVal = ExtractValPL2(result, "r" & r)
        If snapVal <> cellVal Then
            wsMD.Cells(r, 3).Interior.Color = 16776960  ' Kuning
        Else
            wsMD.Cells(r, 3).Interior.ColorIndex = -4142  ' xlNone
        End If
    Next ri

    ' Personil baris R32-R49
    Dim rp As Long
    For rp = PLR_PERSONIL_BASE To PLR_PERSONIL_BASE + 17
        Dim cpVal As String: cpVal = Trim(CStr(wsMD.Cells(rp, 3).Value))
        Dim spVal As String: spVal = ExtractValPL2(result, "r" & rp)
        If spVal <> cpVal Then
            wsMD.Cells(rp, 3).Interior.Color = 16776960
        Else
            wsMD.Cells(rp, 3).Interior.ColorIndex = -4142
        End If
    Next rp
    Exit Sub
ErrDiff:
    Application.StatusBar = False
End Sub


' ============================================================
' CLEAR HIGHLIGHT PL
' ============================================================
Public Sub ClearHighlightPL(Optional wsMDArg As Worksheet = Nothing)
    Dim wsMD As Worksheet
    If wsMDArg Is Nothing Then
        On Error Resume Next
        Set wsMD = ThisWorkbook.Sheets(MD_SHEET)
        On Error GoTo 0
    Else
        Set wsMD = wsMDArg
    End If
    If wsMD Is Nothing Then Exit Sub

    Dim r As Long
    ' PLR rows
    Dim rowList As Variant
    rowList = Array(PLR_KODE_PAKET, PLR_KODE_RUP, PLR_NAMA_PEKERJAAN, PLR_NAMA_SKPD, _
                    PLR_SUB_KEGIATAN, PLR_NAMA_PPK, PLR_NIP_PPK, PLR_NO_SK_PPK, _
                    PLR_NAMA_PP, PLR_NIP_PP, PLR_PAGU, PLR_HPS, PLR_JANGKA_WAKTU, _
                    PLR_SUMBER_DANA, PLR_LOKASI, PLR_JENIS_KONTRAK, PLR_URAIAN_SINGKAT, _
                    PLR_NOMOR_DOKPIL, PLR_TANGGAL_DOKPIL, PLR_NO_UNDANGAN, _
                    PLR_TAHUN_ANGGARAN, PLR_TGL_UNDANGAN, PLR_KODE_REKENING, _
                    PLR_NO_BA_REVIU, PLR_ALAMAT_PP, PLR_SBU_BARU, PLR_SBU_LAMA, _
                    PLR_NAMA_PESERTA, PLR_NPWP_PESERTA)
    Dim ri As Long
    For ri = 0 To UBound(rowList)
        wsMD.Cells(rowList(ri), 3).Interior.ColorIndex = -4142
    Next ri
    For r = PLR_PERSONIL_BASE To PLR_PERSONIL_BASE + 17
        wsMD.Cells(r, 3).Interior.ColorIndex = -4142
    Next r
End Sub


' ============================================================
' BUILD SNAPSHOT JSON PL: semua nilai kolom C → {"r3":"val",...}
' ============================================================
Private Function BuildSnapshotPL(wsMD As Worksheet) As String
    Dim sb As String: sb = "{"
    Dim first As Boolean: first = True

    Dim rowList As Variant
    rowList = Array(PLR_KODE_PAKET, PLR_KODE_RUP, PLR_NAMA_PEKERJAAN, PLR_NAMA_SKPD, _
                    PLR_SUB_KEGIATAN, PLR_NAMA_PPK, PLR_NIP_PPK, PLR_NO_SK_PPK, _
                    PLR_NAMA_PP, PLR_NIP_PP, PLR_PAGU, PLR_HPS, PLR_JANGKA_WAKTU, _
                    PLR_SUMBER_DANA, PLR_LOKASI, PLR_JENIS_KONTRAK, PLR_URAIAN_SINGKAT, _
                    PLR_NOMOR_DOKPIL, PLR_TANGGAL_DOKPIL, PLR_NO_UNDANGAN, _
                    PLR_TAHUN_ANGGARAN, PLR_TGL_UNDANGAN, PLR_KODE_REKENING, _
                    PLR_NO_BA_REVIU, PLR_ALAMAT_PP, PLR_SBU_BARU, PLR_SBU_LAMA, _
                    PLR_NAMA_PESERTA, PLR_NPWP_PESERTA)

    Dim ri As Long
    For ri = 0 To UBound(rowList)
        Dim r As Long: r = rowList(ri)
        If Not first Then sb = sb & ","
        sb = sb & """r" & r & """:""" & EscapeJSONPL(CStr(wsMD.Cells(r, 3).Value)) & """"
        first = False
    Next ri

    ' Personil R32-R49
    Dim rp As Long
    For rp = PLR_PERSONIL_BASE To PLR_PERSONIL_BASE + 17
        sb = sb & ",""r" & rp & """:""" & EscapeJSONPL(CStr(wsMD.Cells(rp, 3).Value)) & """"
    Next rp

    sb = sb & "}"
    BuildSnapshotPL = sb
End Function


' ============================================================
' HELPERS I/O untuk Sync PL
' ============================================================
Private Sub WriteUTF8PL(path As String, content As String)
    Dim ado As Object
    Dim tries As Integer
    For tries = 1 To 5
        On Error Resume Next
        Set ado = CreateObject("ADODB.Stream")
        ado.Type = 2: ado.Charset = "UTF-8": ado.Open
        ado.WriteText content
        ado.SaveToFile path, 2
        ado.Close
        If Err.Number = 0 Then Exit For
        Err.Clear
        Application.Wait Now + TimeSerial(0, 0, 1)  ' tunggu 1 detik, retry
    Next tries
    On Error GoTo 0
End Sub

Private Function ReadUTF8PL(path As String) As String
    Dim ado As Object
    Set ado = CreateObject("ADODB.Stream")
    ado.Type = 2: ado.Charset = "UTF-8": ado.Open
    ado.LoadFromFile path
    ReadUTF8PL = ado.ReadText
    ado.Close
End Function

Private Function EscapeJSONPL(s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, Chr(10), "\n")
    s = Replace(s, Chr(13), "")
    EscapeJSONPL = s
End Function

Private Function ExtractValPL2(json As String, key As String) As String
    Dim searchKey As String: searchKey = """" & key & """:"
    Dim pos As Long: pos = InStr(1, json, searchKey)
    If pos = 0 Then ExtractValPL2 = "": Exit Function
    pos = pos + Len(searchKey)
    Do While Mid(json, pos, 1) = " ": pos = pos + 1: Loop
    If Mid(json, pos, 1) = """" Then
        pos = pos + 1
        Dim endPos As Long: endPos = pos
        Do While endPos <= Len(json)
            If Mid(json, endPos, 1) = """" And Mid(json, endPos - 1, 1) <> "\" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractValPL2 = Mid(json, pos, endPos - pos)
    ElseIf Mid(json, pos, 4) = "null" Then
        ExtractValPL2 = ""
    Else
        Dim endN As Long: endN = pos
        Do While endN <= Len(json) And InStr(",}]", Mid(json, endN, 1)) = 0
            endN = endN + 1
        Loop
        ExtractValPL2 = Trim(Mid(json, pos, endN - pos))
    End If
End Function


' ============================================================
' RUN MERGE PL
' ============================================================
Private Sub RunMergePL(ByVal mode As String, ByVal wordPattern As String, ByVal sheetName As String)
    Dim wordFile As String
    wordFile = FindWordFilePL(wordPattern)
    If wordFile = "" Then Exit Sub

    Dim wordPath As String
    wordPath = ThisWorkbook.Path & "\" & wordFile

    If Dir(wordPath) = "" Then
        MsgBox "File Word tidak ditemukan:" & vbCrLf & wordPath, vbExclamation
        Exit Sub
    End If

    On Error Resume Next
    ThisWorkbook.Save
    On Error GoTo 0

    Dim excelPath As String
    excelPath = ThisWorkbook.FullName

    Dim scriptDir As String
    scriptDir = ScriptDirPL()
    If scriptDir = "" Then
        MsgBox "Python tidak ditemukan. Pastikan V19_Scheduler ada di folder POKJA.", vbCritical
        Exit Sub
    End If

    Dim pyExe As String
    pyExe = scriptDir & "\python\python.exe"

    Dim cmd As String
    cmd = Chr(34) & pyExe & Chr(34) & " " & _
          Chr(34) & scriptDir & "\word_merge.py" & Chr(34) & " " & _
          mode & " " & _
          Chr(34) & wordPath & Chr(34) & " " & _
          Chr(34) & excelPath & Chr(34) & " " & _
          Chr(34) & sheetName & Chr(34)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 0, False
    Set wsh = Nothing

    Application.StatusBar = "Membuka " & wordFile & "... Word akan muncul sebentar lagi."
End Sub

Private Function FindWordFilePL(ByVal pattern As String) As String
    Dim folder As String
    folder = ThisWorkbook.Path
    Dim f As String
    Dim ext As Variant
    For Each ext In Array("*.docx", "*.docm")
        f = Dir(folder & "\" & ext)
        Do While f <> ""
            If Left(f, Len(pattern)) = pattern Then
                FindWordFilePL = f
                Exit Function
            End If
            f = Dir()
        Loop
    Next ext
    MsgBox "File Word tidak ditemukan." & vbCrLf & _
           "Pastikan ada file yang diawali: """ & pattern & """", _
           vbExclamation, "File Tidak Ditemukan"
    FindWordFilePL = ""
End Function

Private Function ScriptDirPL() As String
    Dim folder As String
    folder = ThisWorkbook.Path
    Dim i As Integer
    For i = 1 To 10
        If Dir(folder & "\V19_Scheduler\WPy64-313110\python\python.exe") <> "" Then
            ScriptDirPL = folder & "\V19_Scheduler\WPy64-313110"
            Exit Function
        End If
        folder = Left(folder, InStrRev(folder, "\") - 1)
        If Len(folder) < 3 Then Exit For
    Next i
    ScriptDirPL = ""
End Function


' ============================================================
' HELPERS
' ============================================================
Private Function ExtractJSONValPL(json As String, key As String) As String
    Dim pattern As String
    pattern = """" & key & """" & ":"
    Dim p As Long
    p = InStr(json, pattern)
    If p = 0 Then ExtractJSONValPL = "": Exit Function
    p = p + Len(pattern)
    Do While Mid(json, p, 1) = " ": p = p + 1: Loop
    If Mid(json, p, 1) = """" Then
        p = p + 1
        Dim q As Long: q = p
        Do While q <= Len(json)
            If Mid(json, q, 1) = """" And (q = 1 Or Mid(json, q - 1, 1) <> "\") Then Exit Do
            q = q + 1
        Loop
        ExtractJSONValPL = Mid(json, p, q - p)
    ElseIf Mid(json, p, 4) = "null" Then
        ExtractJSONValPL = ""
    ElseIf Mid(json, p, 1) = "[" Then
        ' JSON array value — cari bracket matching
        Dim depth As Long: depth = 0
        Dim aEnd As Long: aEnd = p
        Do While aEnd <= Len(json)
            Dim ac As String: ac = Mid(json, aEnd, 1)
            If ac = "[" Then depth = depth + 1
            If ac = "]" Then
                depth = depth - 1
                If depth = 0 Then Exit Do
            End If
            aEnd = aEnd + 1
        Loop
        ExtractJSONValPL = Mid(json, p, aEnd - p + 1)
    Else
        Dim endPos As Long: endPos = p
        Do While endPos <= Len(json)
            Dim c As String: c = Mid(json, endPos, 1)
            If c = "," Or c = "}" Or c = "]" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractJSONValPL = Trim(Mid(json, p, endPos - p))
    End If
End Function

Private Function CDblSafe(s As String) As Double
    If s = "" Or s = "null" Then CDblSafe = 0: Exit Function
    On Error Resume Next
    CDblSafe = CDbl(s)
    On Error GoTo 0
End Function


Private Function ParsePersonilArrayPL(jsonArr As String) As String()
    ' Parse JSON array of {jabatan, pengalaman, sertifikat} objects.
    ' Returns flat string array stride=3: [jabatan0, pengalaman0, sertifikat0, jabatan1, ...]
    ' Maximum 6 personil = 18 elemen. sertifikat boleh "" (backward compatible).
    Dim result(17) As String  ' max 6 personil * 3
    Dim count As Long: count = 0
    Dim pos As Long: pos = 1
    Dim arrLen As Long: arrLen = Len(jsonArr)

    Do While pos <= arrLen And count < 6
        Dim bStart As Long: bStart = InStr(pos, jsonArr, "{")
        If bStart = 0 Then Exit Do
        Dim bEnd As Long: bEnd = InStr(bStart, jsonArr, "}")
        If bEnd = 0 Then Exit Do

        Dim obj As String: obj = Mid(jsonArr, bStart, bEnd - bStart + 1)
        Dim jabatan As String: jabatan = ExtractJSONValPL(obj, "jabatan")
        Dim pengalaman As String: pengalaman = ExtractJSONValPL(obj, "pengalaman")
        Dim sertifikat As String: sertifikat = ExtractJSONValPL(obj, "sertifikat")

        result(count * 3) = jabatan
        result(count * 3 + 1) = pengalaman
        result(count * 3 + 2) = sertifikat
        count = count + 1
        pos = bEnd + 1
    Loop

    ' Resize to actual count
    If count = 0 Then
        Dim emptyArr(2) As String
        ParsePersonilArrayPL = emptyArr
        Exit Function
    End If

    ReDim resized(count * 3 - 1) As String
    Dim ii As Long
    For ii = 0 To count * 3 - 1
        resized(ii) = result(ii)
    Next ii
    ParsePersonilArrayPL = resized
End Function


' ============================================================
' Lookup KPA (NIP + nomor SK + tanggal) dari master_kpa via nama
' ============================================================
Private Function LookupKPA(namaPPK As String) As String
    LookupKPA = ""
    If namaPPK = "" Then Exit Function
    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' Strip gelar (S.T., M.M., S.E., dll) untuk lookup ilike
    Dim namaClean As String: namaClean = namaPPK
    Dim p As Long: p = InStr(namaClean, ",")
    If p > 0 Then namaClean = Trim(Left(namaClean, p - 1))
    ' ilike *namaClean*
    Dim q As String: q = Replace(namaClean, " ", "%20")
    Dim url As String
    url = SB_URL & "/rest/v1/master_kpa?nama=ilike.*" & q & "*&select=nip,nomor_sk,tanggal_sk&limit=1"
    http.Open "GET", url, False
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send
    If http.Status = 200 Then LookupKPA = http.ResponseText
    On Error GoTo 0
End Function


' ============================================================
' Format "2026-02-06" -> "06 Februari 2026"
' ============================================================
Private Function FormatTanggalIndo(tglISO As String) As String
    FormatTanggalIndo = tglISO
    If Len(tglISO) < 10 Then Exit Function
    Dim parts() As String
    parts = Split(Left(tglISO, 10), "-")
    If UBound(parts) < 2 Then Exit Function
    Dim tahun As String, bulan As String, hari As String
    tahun = parts(0): hari = parts(2)
    Dim bulanArr As Variant
    bulanArr = Array("Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember")
    Dim idx As Long
    On Error Resume Next
    idx = CLng(parts(1)) - 1
    On Error GoTo 0
    If idx >= 0 And idx <= 11 Then bulan = bulanArr(idx)
    FormatTanggalIndo = hari & " " & bulan & " " & tahun
End Function


' ============================================================
' Lookup singkatan dari master_dinas Supabase via nama_dinas (ilike)
' ============================================================
Private Function LookupSingkatanDinas(namaDinas As String) As String
    LookupSingkatanDinas = ""
    If namaDinas = "" Then Exit Function

    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    ' Query: ilike "*nama_dinas*"
    Dim q As String: q = Replace(namaDinas, " ", "%20")
    Dim url As String
    url = SB_URL & "/rest/v1/master_dinas?nama_dinas=ilike.*" & q & "*&select=singkatan&limit=1"
    http.Open "GET", url, False
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send
    If http.Status = 200 Then
        Dim resp As String: resp = http.ResponseText
        LookupSingkatanDinas = ExtractJSONValPL(resp, "singkatan")
    End If
    On Error GoTo 0
End Function


' ============================================================
' Lookup alamat_pp_bertugas dari master_dinas Supabase via nama_dinas (ilike)
' ============================================================
Private Function LookupAlamatPP(namaDinas As String) As String
    LookupAlamatPP = ""
    If namaDinas = "" Then Exit Function

    On Error Resume Next
    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim q As String: q = Replace(namaDinas, " ", "%20")
    Dim url As String
    url = SB_URL & "/rest/v1/master_dinas?nama_dinas=ilike.*" & q & "*&select=alamat_pp_bertugas&limit=1"
    http.Open "GET", url, False
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send
    If http.Status = 200 Then
        Dim resp As String: resp = http.ResponseText
        LookupAlamatPP = ExtractJSONValPL(resp, "alamat_pp_bertugas")
    End If
    On Error GoTo 0
End Function

Private Function LookupTeleponPP(namaDinas As String) As String
    LookupTeleponPP = ""
    If namaDinas = "" Then Exit Function

    On Error Resume Next
    Dim http2 As Object
    Set http2 = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim q2 As String: q2 = Replace(namaDinas, " ", "%20")
    Dim url2 As String
    url2 = SB_URL & "/rest/v1/master_dinas?nama_dinas=ilike.*" & q2 & "*&select=telepon_pp&limit=1"
    http2.Open "GET", url2, False
    http2.SetRequestHeader "apikey", SB_KEY
    http2.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http2.SetRequestHeader "Accept", "application/json"
    http2.Send
    If http2.Status = 200 Then
        Dim resp2 As String: resp2 = http2.ResponseText
        LookupTeleponPP = ExtractJSONValPL(resp2, "telepon_pp")
    End If
    On Error GoTo 0
End Function


' ============================================================
' MUAT HPS PL: GET hps_items_pl dari Supabase -> sheet "5. HPS"
' ============================================================
Public Sub MuatHPSPL()
    Dim wsMD As Worksheet
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)
    Dim kodePaket As String
    kodePaket = CStr(wsMD.Cells(PLR_KODE_PAKET, 3).Value)
    If kodePaket = "" Then
        MsgBox "Kode Paket di @ Master Data C3 kosong. Pilih paket dulu.", vbExclamation, "Muat HPS PL"
        Exit Sub
    End If

    Dim wsHPS As Worksheet
    On Error Resume Next
    Set wsHPS = ThisWorkbook.Sheets("5. HPS")
    On Error GoTo 0
    If wsHPS Is Nothing Then
        MsgBox "Sheet '5. HPS' tidak ditemukan.", vbCritical, "Muat HPS PL"
        Exit Sub
    End If

    Dim url As String
    url = SB_URL & "/rest/v1/hps_items_pl" & _
          "?kode_paket=eq." & kodePaket & _
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
        MsgBox "Gagal fetch HPS PL: HTTP " & http.Status, vbExclamation, "Muat HPS PL"
        Exit Sub
    End If

    Dim json As String: json = http.ResponseText
    If json = "[]" Or json = "" Then
        MsgBox "HPS PL belum ada di Supabase untuk kode " & kodePaket & "." & vbCrLf & _
               "Jalankan tombol scrape HPS di Streamlit dulu.", vbInformation, "Muat HPS PL"
        Exit Sub
    End If

    ' Bersihkan A2:I bawah
    Dim lastRow As Long
    lastRow = wsHPS.Cells(wsHPS.Rows.Count, 2).End(xlUp).Row
    If lastRow >= 2 Then wsHPS.Range("A2:I" & lastRow).ClearContents
    ' Bersihkan baris total lama (50-52)
    wsHPS.Range("A50:I52").ClearContents
    ' Reset highlight
    wsHPS.Range("A2:I" & (lastRow + 5)).Interior.ColorIndex = xlNone

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

        Dim isDivisi As String: isDivisi = ExtractJSONValPL(obj, "is_divisi")
        Dim jenisBJ  As String: jenisBJ  = ExtractJSONValPL(obj, "jenis_bj")
        Dim satuan   As String: satuan   = ExtractJSONValPL(obj, "satuan")
        Dim volS     As String: volS     = ExtractJSONValPL(obj, "vol")
        Dim hargaS   As String: hargaS   = ExtractJSONValPL(obj, "harga")
        Dim pajakS   As String: pajakS   = ExtractJSONValPL(obj, "pajak_pct")
        Dim totSpseS As String: totSpseS = ExtractJSONValPL(obj, "total_spse")
        Dim totHitS  As String: totHitS  = ExtractJSONValPL(obj, "total_hitung")
        Dim selisihS As String: selisihS = ExtractJSONValPL(obj, "selisih")
        Dim selOkS   As String: selOkS   = ExtractJSONValPL(obj, "selisih_ok")
        Dim urutan   As String: urutan   = ExtractJSONValPL(obj, "urutan")

        If totalNilai = "" Then
            totalNilai = ExtractJSONValPL(obj, "total_nilai")
            totalBulat = ExtractJSONValPL(obj, "total_nilai_bulat")
        End If

        With wsHPS
            .Cells(baris, 1).Value = CDblSafe(urutan)
            .Cells(baris, 2).Value = jenisBJ

            If isDivisi <> "true" Then
                .Cells(baris, 3).Value = satuan
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
                If totHitS <> "" Then
                    .Cells(baris, 8).Value = CDblSafe(totHitS)
                    .Cells(baris, 8).NumberFormat = "#,##0.00"
                End If
                Dim selisihVal As Double: selisihVal = CDblSafe(selisihS)
                .Cells(baris, 9).Value = selisihVal
                .Cells(baris, 9).NumberFormat = "#,##0.00"

                totalHitungAll = totalHitungAll + CDblSafe(totHitS)

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

    Dim itemCount As Long: itemCount = baris - 2
    Dim msg As String
    msg = "HPS PL berhasil dimuat: " & itemCount & " baris."
    If adaSelisih Then
        msg = msg & vbCrLf & vbCrLf & _
              "Selisih SPSE vs hitung manual (highlight kuning):" & vbCrLf & selisihMsg
        MsgBox msg, vbExclamation, "Muat HPS PL"
    Else
        MsgBox msg, vbInformation, "Muat HPS PL"
    End If
    Exit Sub

ErrHPS:
    MsgBox "Gagal muat HPS PL: " & Err.Description, vbCritical, "Muat HPS PL"
End Sub


' ============================================================
' MUAT KODE UNIK PL: GET kode_unik dari Supabase -> @ Master Data G2
' ============================================================
Public Sub MuatKodeUnikPL()
    Dim wsMD As Worksheet
    Set wsMD = ThisWorkbook.Sheets(MD_SHEET)
    Dim kodePaket As String
    kodePaket = CStr(wsMD.Cells(PLR_KODE_PAKET, 3).Value)
    If kodePaket = "" Then
        MsgBox "Kode Paket di @ Master Data C3 kosong. Pilih paket dulu.", vbExclamation, "Muat Kode Unik PL"
        Exit Sub
    End If

    Dim url As String
    url = SB_URL & "/rest/v1/draft_paket_pl" & _
          "?kode_paket=eq." & kodePaket & _
          "&select=kode_unik"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")
    On Error GoTo ErrKU
    http.Open "GET", url, False
    http.SetTimeouts 5000, 5000, 10000, 10000
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send

    If http.Status <> 200 Then
        MsgBox "Gagal fetch Kode Unik: HTTP " & http.Status, vbExclamation, "Muat Kode Unik PL"
        Exit Sub
    End If

    Dim json As String: json = http.ResponseText
    Dim ku As String: ku = ExtractJSONValPL(json, "kode_unik")
    If ku = "" Or ku = "null" Then
        MsgBox "Kode Unik belum ada di Supabase untuk paket ini." & vbCrLf & _
               "Jalankan tombol 'Kode Unik PL' atau Streamlit dulu.", vbInformation, "Muat Kode Unik PL"
        Exit Sub
    End If

    wsMD.Range("G2").Value = ku
    MsgBox "Kode Unik diisi ke G2: " & ku, vbInformation, "Muat Kode Unik PL"
    Exit Sub

ErrKU:
    MsgBox "Gagal muat Kode Unik: " & Err.Description, vbCritical, "Muat Kode Unik PL"
End Sub
