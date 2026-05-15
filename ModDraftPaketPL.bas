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
'   R3  Kode Paket       R13 Pagu Anggaran     R23 Tahun Anggaran
'   R4  Kode RUP         R14 Nilai HPS         R24 Tanggal Undangan PL
'   R5  Nama Pekerjaan   R15 Jangka Waktu      R25 Nomor DPA
'   R6  Nama SKPD        R16 Sumber Dana       R27 SBU Baru
'   R7  Sub Kegiatan     R17 Lokasi            R28 SBU Lama
'   R8  Nama PPK         R18 Jenis Kontrak     R30 Jabatan Teknis
'   R9  NIP PPK          R19 Uraian Singkat    R31 SKK Teknis
'   R10 No SK PPK        R20 Nomor Dokpil      R32 Pengalaman Teknis
'   R11 Nama PP          R21 Tanggal Dokpil    R33 Jabatan K3
'   R12 NIP PP           R22 No Undangan PL    R34 SKK K3 / R35 Pengalaman K3

' Konfigurasi Supabase
Private Const SB_URL As String = "%%SUPABASE_URL%%"
Private Const SB_KEY As String = "%%SUPABASE_KEY%%"

' Word Merge — pattern nama file + sheet name
Private Const WM_SHEET_BA     As String = "satu_data"
Private Const WM_SHEET_REVIU  As String = "list_reviu"
Private Const WM_SHEET_DOKPIL As String = "list_dokpil"
Private Const WM_PAT_BA       As String = "1. Full Dokumen BA PLJKK"
Private Const WM_PAT_REVIU    As String = "2. Isi Reviu PLJKK"
Private Const WM_PAT_DOKPIL   As String = "3. Dokpil Full PLJKK"

' Sheet & Cell selector
Private Const MD_SHEET As String = "@ Master Data"
Private Const CELL_SELECTOR As String = "F1"

' Kolom yang di-fetch dari draft_paket_pl
Private Const SB_SELECT As String = "kode_paket,nama_paket,satker,kode_rup,nilai_hps,jenis_pl,jenis_kontrak,status,nama_ppk,nip_ppk,no_sk_ppk,nilai_pagu,jangka_waktu,sumber_anggaran,lokasi,sbu_baru,sbu_lama,jabatan_teknis,skk_teknis,jabatan_k3,skk_k3,dpa_nomor,sub_kegiatan"

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
Private Const PLR_NOMOR_DPA       As Integer = 25
Private Const PLR_SBU_BARU        As Integer = 27
Private Const PLR_SBU_LAMA        As Integer = 28
Private Const PLR_JABATAN_TEKNIS  As Integer = 30
Private Const PLR_SKK_TEKNIS      As Integer = 31
Private Const PLR_PENGALAMAN_TKN  As Integer = 32
Private Const PLR_JABATAN_K3      As Integer = 33
Private Const PLR_SKK_K3          As Integer = 34
Private Const PLR_PENGALAMAN_K3   As Integer = 35

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
    ' 19=jabatan_k3, 20=skk_k3, 21=dpa_nomor, 22=sub_kegiatan

    With wsMD
        ' ── INPUT DATA dari Supabase ──────────────────────────────────────
        .Cells(PLR_KODE_PAKET, 3).NumberFormat = "@"
        .Cells(PLR_KODE_PAKET, 3).Value = CStr(item(0))     ' kode_paket
        .Cells(PLR_KODE_RUP, 3).Value   = CStr(item(3))     ' kode_rup
        .Cells(PLR_NAMA_PEKERJAAN, 3).Value = CStr(item(1)) ' nama_paket
        .Cells(PLR_NAMA_SKPD, 3).Value  = CStr(item(2))     ' satker
        .Cells(PLR_SUB_KEGIATAN, 3).Value = CStr(item(22))  ' sub_kegiatan
        .Cells(PLR_NAMA_PPK, 3).Value   = CStr(item(8))     ' nama_ppk
        If CStr(item(9)) <> "" Then
            .Cells(PLR_NIP_PPK, 3).NumberFormat = "@"
            .Cells(PLR_NIP_PPK, 3).Value = CStr(item(9))   ' nip_ppk
        End If
        If CStr(item(10)) <> "" Then
            .Cells(PLR_NO_SK_PPK, 3).Value = CStr(item(10)) ' no_sk_ppk
        End If

        ' Nama PP + NIP PP: TIDAK ditimpa jika sudah ada (hardcoded)
        ' Biarkan nilai default di R11/R12 tetap

        ' Nilai finansial
        .Cells(PLR_PAGU, 3).Value        = CStr(item(11))   ' pagu_anggaran
        .Cells(PLR_HPS, 3).Value         = CStr(item(4))    ' nilai_hps
        .Cells(PLR_SUMBER_DANA, 3).Value = CStr(item(13))   ' sumber_dana
        .Cells(PLR_LOKASI, 3).Value      = CStr(item(14))   ' lokasi

        ' Jangka waktu: ambil angka saja
        Dim jw As String: jw = Trim(CStr(item(12)))
        Dim sp As Long: sp = InStr(jw, " ")
        If sp > 0 Then jw = Left(jw, sp - 1)
        If IsNumeric(jw) Then
            .Cells(PLR_JANGKA_WAKTU, 3).Value = CLng(jw)
        Else
            .Cells(PLR_JANGKA_WAKTU, 3).Value = CStr(item(12))
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

        ' DPA nomor
        .Cells(PLR_NOMOR_DPA, 3).Value = CStr(item(21))     ' dpa_nomor

        ' ── SBU ──────────────────────────────────────────────────────────
        If CStr(item(15)) <> "" Then
            .Cells(PLR_SBU_BARU, 3).Value = CStr(item(15))  ' sbu_baru
        End If
        If CStr(item(16)) <> "" Then
            .Cells(PLR_SBU_LAMA, 3).Value = CStr(item(16))  ' sbu_lama
        End If

        ' ── PERSONIL TEKNIS ──────────────────────────────────────────────
        If CStr(item(17)) <> "" Then
            .Cells(PLR_JABATAN_TEKNIS, 3).Value = CStr(item(17)) ' jabatan_teknis
        End If
        If CStr(item(18)) <> "" Then
            .Cells(PLR_SKK_TEKNIS, 3).Value = CStr(item(18))     ' skk_teknis
        End If
        If CStr(item(19)) <> "" Then
            .Cells(PLR_JABATAN_K3, 3).Value = CStr(item(19))     ' jabatan_k3
        End If
        If CStr(item(20)) <> "" Then
            .Cells(PLR_SKK_K3, 3).Value = CStr(item(20))         ' skk_k3
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

        Dim item(22) As Variant
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

Public Sub RelinkPL()
    ' Relink semua Word PL ke Excel ini (update path data source)
    Dim scriptDir As String
    scriptDir = ScriptDirPL()
    If scriptDir = "" Then
        MsgBox "Python tidak ditemukan.", vbCritical
        Exit Sub
    End If
    Dim pyExe As String: pyExe = scriptDir & "\python\python.exe"
    Dim setupScript As String: setupScript = scriptDir & "\setup_paket_baru.py"
    Dim folderPath As String: folderPath = ThisWorkbook.Path

    ' Deteksi output_base = parent folder dari folder paket
    Dim outputBase As String
    outputBase = Left(folderPath, InStrRev(folderPath, "\") - 1)

    Dim folderName As String
    folderName = Mid(folderPath, InStrRev(folderPath, "\") + 1)

    Dim cmd As String
    cmd = Chr(34) & pyExe & Chr(34) & " " & _
          Chr(34) & setupScript & Chr(34) & " --mode pl" & _
          " --output-dir " & Chr(34) & outputBase & Chr(34) & _
          " " & Chr(34) & folderName & Chr(34)

    Dim wsh As Object
    Set wsh = CreateObject("WScript.Shell")
    wsh.Run cmd, 1, True  ' tunggu selesai, tampilkan jendela
    Set wsh = Nothing

    MsgBox "Relink selesai! Tutup dan buka ulang file Word untuk melihat hasilnya.", vbInformation
End Sub

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
