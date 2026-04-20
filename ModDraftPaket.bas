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
Public Sub MuatDraftPaket()
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

    ws.Range(CELL_SELECTOR).Value = ""  ' Reset pilihan
    MsgBox m_DataCache.Count & " paket berhasil dimuat." & vbCrLf & _
           "Pilih paket di cell " & CELL_SELECTOR & " untuk mengisi data.", vbInformation, "Draft Paket Dimuat"
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
            Exit For
        End If
    Next i
End Sub


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
