Attribute VB_Name = "ModKKEvaluasi"
' ============================================================
' ModKKEvaluasi - Isi sheet "3. KK Evaluasi Kualifikasi"
'                 dari data Supabase tabel kk_evaluasi_peserta
' ============================================================
' Flow:
'   1. Tombol "Muat KK Evaluasi" di sheet KK Evaluasi di-klik
'   2. MuatKKEvaluasi() baca kode_tender dari sheet "1. Input Data" cell E5
'   3. GET Supabase → filter kode_tender, order urutan asc
'   4. Isi kolom C (urutan 1), D (urutan 2), E (urutan 3) di sheet KK Evaluasi

Private Const SB_URL As String = "%%SUPABASE_URL%%"
Private Const SB_KEY As String = "%%SUPABASE_KEY%%"
Private Const SB_TABLE As String = "kk_evaluasi_peserta"

Private Const SHEET_INPUT As String = "1. Input Data"
Private Const SHEET_KK As String = "3. KK Evaluasi Kualifikasi"
Private Const CELL_KODE_TENDER As String = "E5"

' ── Row constants (sheet v1.4 verified) ──────────────────────
' Poin 1: Perizinan Berusaha
Private Const ROW_POIN1_KESIMPULAN As Integer = 6
Private Const ROW_NIB_ADA As Integer = 7
Private Const ROW_NIB_NOMOR As Integer = 8
Private Const ROW_SS_TERVERIFIKASI As Integer = 9
Private Const ROW_SS_NOMOR As Integer = 10
Private Const ROW_TANGKAPAN_OSS As Integer = 11
Private Const ROW_SBU_2015 As Integer = 12
' Poin 2: SBU
Private Const ROW_SBU_KESIMPULAN As Integer = 13
Private Const ROW_SBU_NOMOR As Integer = 14
Private Const ROW_SBU_BERLAKU As Integer = 15
Private Const ROW_SBU_KUALIFIKASI As Integer = 16
Private Const ROW_SBU_SUBKLAS As Integer = 17
' Poin 3: Pengalaman
Private Const ROW_PGL_KESIMPULAN As Integer = 19
Private Const ROW_PGL1_NAMA As Integer = 20
Private Const ROW_PGL1_INSTANSI As Integer = 21
Private Const ROW_PGL1_NILAI As Integer = 22
Private Const ROW_PGL1_TANGGAL As Integer = 23
Private Const ROW_PGL1_NOMOR As Integer = 24
Private Const ROW_PGL2_NAMA As Integer = 25
Private Const ROW_PGL2_INSTANSI As Integer = 26
Private Const ROW_PGL2_NILAI As Integer = 27
Private Const ROW_PGL2_TANGGAL As Integer = 28
Private Const ROW_PGL2_NOMOR As Integer = 29
' Poin 4: SKP
Private Const ROW_SKP_KESIMPULAN As Integer = 30
Private Const ROW_SKP As Integer = 33
' Poin 5: NPWP/KSWP
Private Const ROW_NPWP_KESIMPULAN As Integer = 34
Private Const ROW_NPWP As Integer = 35
Private Const ROW_KSWP As Integer = 36
' Poin 6: Akta
Private Const ROW_AKTA_P_KESIMPULAN As Integer = 38
Private Const ROW_AKTA_P_NOMOR As Integer = 39
Private Const ROW_AKTA_P_TANGGAL As Integer = 40
Private Const ROW_AKTA_P_NOTARIS As Integer = 41
Private Const ROW_AKTA_K_KESIMPULAN As Integer = 42
Private Const ROW_AKTA_K_NOMOR As Integer = 43
Private Const ROW_AKTA_K_TANGGAL As Integer = 44
Private Const ROW_AKTA_K_NOTARIS As Integer = 45
Private Const ROW_PEMILIK_HEADER As Integer = 46
Private Const ROW_PEMILIK_1 As Integer = 47
Private Const ROW_PEMILIK_2 As Integer = 48
Private Const ROW_PEMILIK_3 As Integer = 49
Private Const ROW_PEMILIK_4 As Integer = 50
' Poin 7: Kinerja
Private Const ROW_KINERJA_ADA As Integer = 51
Private Const ROW_KINERJA_NILAI As Integer = 52
' Poin 8: Daftar Hitam
Private Const ROW_DAFTAR_HITAM As Integer = 53
' Hasil
Private Const ROW_HASIL_MS As Integer = 54


' ============================================================
' FUNGSI UTAMA: dipanggil dari tombol di sheet KK Evaluasi
' ============================================================
Public Sub MuatKKEvaluasi()
    Dim wsInput As Worksheet, wsKK As Worksheet
    On Error GoTo ErrHandler

    Set wsInput = ThisWorkbook.Sheets(SHEET_INPUT)
    Dim kodeTender As String
    kodeTender = Trim(wsInput.Range(CELL_KODE_TENDER).Value)

    If kodeTender = "" Then
        MsgBox "Kode Tender belum diisi di sheet '" & SHEET_INPUT & "' cell " & CELL_KODE_TENDER & "." & vbCrLf & _
               "Pilih paket terlebih dahulu.", vbExclamation
        Exit Sub
    End If

    Dim json As String
    json = FetchKKEvaluasi(kodeTender)
    If json = "" Then Exit Sub

    Dim items() As Variant
    Dim itemCount As Integer
    ParseKKJSON json, items, itemCount

    If itemCount = 0 Then
        MsgBox "Tidak ada data KK Evaluasi untuk kode tender " & kodeTender & "." & vbCrLf & _
               "Jalankan 'Parse & Simpan KK Evaluasi' di Asisten Pokja terlebih dahulu.", vbInformation
        Exit Sub
    End If

    Set wsKK = ThisWorkbook.Sheets(SHEET_KK)
    wsKK.Unprotect

    ' Kolom: urutan 1=C(3), 2=D(4), 3=E(5)
    Dim colMap(1 To 3) As Integer
    colMap(1) = 3: colMap(2) = 4: colMap(3) = 5

    ' Bersihkan kolom C/D/E baris 6-54: hapus isi + Data Validation + highlight
    Dim r As Integer, c As Integer
    For r = 6 To 54
        For c = 3 To 5
            With wsKK.Cells(r, c)
                .ClearContents
                .Validation.Delete
                .Interior.ColorIndex = xlNone
            End With
        Next c
    Next r

    ' Kumpulkan peringatan field perlu cek manual (per peserta)
    Dim warnMsg As String: warnMsg = ""

    Dim i As Integer
    For i = 1 To itemCount
        If i > 3 Then Exit For
        Dim it As Variant
        it = items(i - 1)
        Dim col As Integer
        col = colMap(i)
        Dim namaP As String: namaP = GetF(it, "nama")
        If namaP = "" Then namaP = "Peserta " & i
        Dim warnPeserta As String: warnPeserta = ""

        ' ── Poin 1: Perizinan Berusaha ──────────────────────
        Dim nibNomor As String: nibNomor = GetF(it, "nib_nomor")
        Dim ssTerverifikasi As String: ssTerverifikasi = GetF(it, "ss_terverifikasi")

        ' Kesimpulan poin 1: tentukan sub-poin yang dipenuhi
        Dim poin1Val As String
        If nibNomor <> "" And ssTerverifikasi = "Terverifikasi" Then
            poin1Val = "Memenuhi Syarat Kualifikasi pada Poin a)."
        ElseIf nibNomor <> "" And ssTerverifikasi = "Belum Terverifikasi" Then
            poin1Val = "Memenuhi Syarat Kualifikasi pada Poin b)."
        ElseIf nibNomor <> "" Then
            poin1Val = "Memenuhi Syarat Kualifikasi pada Poin c)."
        Else
            poin1Val = "Tidak Memenuhi"
        End If
        wsKK.Cells(ROW_POIN1_KESIMPULAN, col).Value = poin1Val

        wsKK.Cells(ROW_NIB_ADA, col).Value = IIf(nibNomor <> "", "Ada", "Tidak Ada")
        wsKK.Cells(ROW_NIB_NOMOR, col).Value = nibNomor
        wsKK.Cells(ROW_SS_TERVERIFIKASI, col).Value = ssTerverifikasi
        ' Highlight jika SS terverifikasi tidak dikenali (bukan nilai baku)
        If ssTerverifikasi <> "Terverifikasi" And ssTerverifikasi <> "Belum Terverifikasi" Then
            wsKK.Cells(ROW_SS_TERVERIFIKASI, col).Interior.Color = RGB(255, 255, 0)
            warnPeserta = warnPeserta & "  - SS Terverifikasi: """ & ssTerverifikasi & """ (tidak dikenali)" & vbCrLf
        End If
        wsKK.Cells(ROW_SS_NOMOR, col).Value = GetF(it, "ss_nomor")
        wsKK.Cells(ROW_TANGKAPAN_OSS, col).Value = "-"
        wsKK.Cells(ROW_SBU_2015, col).Value = "-"

        ' ── Poin 2: SBU ──────────────────────────────────────
        Dim sbuNomor As String: sbuNomor = GetF(it, "sbu_nomor")
        Dim sbuSubklas As String: sbuSubklas = GetF(it, "sbu_subklas_label")
        Dim sbuKualifikasi As String: sbuKualifikasi = GetF(it, "sbu_kualifikasi")

        ' Kesimpulan SBU: teks deskriptif subklasifikasi
        Dim sbuKesimpulan As String
        If sbuNomor <> "" And sbuSubklas <> "" Then
            sbuKesimpulan = sbuSubklas
        ElseIf sbuNomor <> "" Then
            sbuKesimpulan = "Ada (subklasifikasi tidak terdeteksi)"
        Else
            sbuKesimpulan = "Tidak Ada"
        End If
        wsKK.Cells(ROW_SBU_KESIMPULAN, col).Value = sbuKesimpulan
        wsKK.Cells(ROW_SBU_NOMOR, col).Value = sbuNomor
        wsKK.Cells(ROW_SBU_BERLAKU, col).Value = GetF(it, "sbu_berlaku")
        wsKK.Cells(ROW_SBU_KUALIFIKASI, col).Value = sbuKualifikasi
        wsKK.Cells(ROW_SBU_SUBKLAS, col).Value = sbuSubklas

        ' ── Poin 3: Pengalaman ───────────────────────────────
        Dim pgl1Nama As String: pgl1Nama = GetF(it, "pgl1_nama")
        Dim pglKesimpulan As String
        pglKesimpulan = IIf(pgl1Nama <> "", "Memenuhi", "Tidak Memenuhi")
        wsKK.Cells(ROW_PGL_KESIMPULAN, col).Value = pglKesimpulan
        wsKK.Cells(ROW_PGL1_NAMA, col).Value = pgl1Nama
        wsKK.Cells(ROW_PGL1_INSTANSI, col).Value = GetF(it, "pgl1_instansi")
        wsKK.Cells(ROW_PGL1_NILAI, col).Value = GetF(it, "pgl1_nilai")
        wsKK.Cells(ROW_PGL1_TANGGAL, col).Value = GetF(it, "pgl1_tanggal")
        wsKK.Cells(ROW_PGL1_NOMOR, col).Value = GetF(it, "pgl1_nomor")
        wsKK.Cells(ROW_PGL2_NAMA, col).Value = GetF(it, "pgl2_nama")
        wsKK.Cells(ROW_PGL2_INSTANSI, col).Value = GetF(it, "pgl2_instansi")
        wsKK.Cells(ROW_PGL2_NILAI, col).Value = GetF(it, "pgl2_nilai")
        wsKK.Cells(ROW_PGL2_TANGGAL, col).Value = GetF(it, "pgl2_tanggal")
        wsKK.Cells(ROW_PGL2_NOMOR, col).Value = GetF(it, "pgl2_nomor")

        ' ── Poin 4: SKP ──────────────────────────────────────
        Dim skpVal As String: skpVal = GetF(it, "skp")
        Dim skpCatatan As String: skpCatatan = GetF(it, "skp_catatan")
        wsKK.Cells(ROW_SKP_KESIMPULAN, col).Value = "Memenuhi"
        ' Format "5 SKP" — ambil angka dari skp lalu tambah " SKP"
        Dim skpNum As String
        If skpVal <> "" Then
            skpNum = skpVal & " SKP"
        Else
            skpNum = skpCatatan
        End If
        If GetF(it, "skp_berbeda") = "true" Then
            skpNum = skpNum & " [!]"
        End If
        wsKK.Cells(ROW_SKP, col).Value = skpNum
        ' Highlight jika SKP kosong (tidak terdeteksi)
        If skpNum = "" Then
            wsKK.Cells(ROW_SKP, col).Interior.Color = RGB(255, 255, 0)
            warnPeserta = warnPeserta & "  - SKP: tidak terdeteksi" & vbCrLf
        End If

        ' ── Poin 5: NPWP / KSWP ──────────────────────────────
        Dim kswpVal As String: kswpVal = GetF(it, "kswp_status")
        Dim npwpVal As String: npwpVal = GetF(it, "npwp")
        Dim npwpKesimpulan As String
        If kswpVal = "VALID" Then
            npwpKesimpulan = "Memenuhi"
        ElseIf kswpVal = "TIDAK VALID" Then
            npwpKesimpulan = "Tidak Memenuhi"
        Else
            npwpKesimpulan = "-"
        End If
        wsKK.Cells(ROW_NPWP_KESIMPULAN, col).Value = npwpKesimpulan
        wsKK.Cells(ROW_NPWP, col).Value = npwpVal
        ' KSWP: jika tidak diketahui/kosong → kosongkan + highlight kuning (isi manual)
        If kswpVal = "TIDAK DIKETAHUI" Or kswpVal = "" Then
            wsKK.Cells(ROW_KSWP, col).Value = ""
            wsKK.Cells(ROW_KSWP, col).Interior.Color = RGB(255, 255, 0)
            warnPeserta = warnPeserta & "  - KSWP: tidak terdeteksi (OCR gagal atau tidak ada PDF)" & vbCrLf
        ElseIf kswpVal <> "VALID" And kswpVal <> "TIDAK VALID" Then
            wsKK.Cells(ROW_KSWP, col).Value = kswpVal
            wsKK.Cells(ROW_KSWP, col).Interior.Color = RGB(255, 255, 0)
            warnPeserta = warnPeserta & "  - KSWP: """ & kswpVal & """ (nilai tidak baku)" & vbCrLf
        Else
            wsKK.Cells(ROW_KSWP, col).Value = kswpVal
        End If

        ' ── Poin 6: Akta ─────────────────────────────────────
        Dim aktaPNomor As String: aktaPNomor = GetF(it, "akta_p_nomor")
        Dim aktaKNomor As String: aktaKNomor = GetF(it, "akta_k_nomor")
        wsKK.Cells(ROW_AKTA_P_KESIMPULAN, col).Value = IIf(aktaPNomor <> "", "Memenuhi", "-")
        wsKK.Cells(ROW_AKTA_P_NOMOR, col).Value = aktaPNomor
        wsKK.Cells(ROW_AKTA_P_TANGGAL, col).Value = GetF(it, "akta_p_tanggal")
        wsKK.Cells(ROW_AKTA_P_NOTARIS, col).Value = GetF(it, "akta_p_notaris")
        wsKK.Cells(ROW_AKTA_K_KESIMPULAN, col).Value = IIf(aktaKNomor <> "", "Memenuhi", "-")
        wsKK.Cells(ROW_AKTA_K_NOMOR, col).Value = aktaKNomor
        wsKK.Cells(ROW_AKTA_K_TANGGAL, col).Value = GetF(it, "akta_k_tanggal")
        wsKK.Cells(ROW_AKTA_K_NOTARIS, col).Value = GetF(it, "akta_k_notaris")
        wsKK.Cells(ROW_PEMILIK_HEADER, col).Value = "-"
        wsKK.Cells(ROW_PEMILIK_1, col).Value = GetF(it, "pemilik_1")
        wsKK.Cells(ROW_PEMILIK_2, col).Value = GetF(it, "pemilik_2")
        wsKK.Cells(ROW_PEMILIK_3, col).Value = GetF(it, "pemilik_3")
        wsKK.Cells(ROW_PEMILIK_4, col).Value = GetF(it, "pemilik_4")

        ' ── Poin 7: Kinerja ──────────────────────────────────
        Dim kinerjaAda As String: kinerjaAda = GetF(it, "kinerja_ada")
        Dim kinerjaVal As String: kinerjaVal = GetF(it, "kinerja_nilai")
        If kinerjaAda = "true" Then
            wsKK.Cells(ROW_KINERJA_ADA, col).Value = "ADA"
            wsKK.Cells(ROW_KINERJA_NILAI, col).Value = kinerjaVal
            ' Highlight jika kinerja ada tapi nilai tidak terdeteksi
            If kinerjaVal = "" Then
                wsKK.Cells(ROW_KINERJA_NILAI, col).Interior.Color = RGB(255, 255, 0)
                warnPeserta = warnPeserta & "  - Kinerja: ADA tapi nilai tidak terdeteksi (cek PDF kinerja)" & vbCrLf
            End If
        Else
            wsKK.Cells(ROW_KINERJA_ADA, col).Value = "TIDAK MENYAMPAIKAN"
            wsKK.Cells(ROW_KINERJA_NILAI, col).Value = "-"
        End If

        ' ── Poin 8: Daftar Hitam (terikat SPSE) ─────────────
        wsKK.Cells(ROW_DAFTAR_HITAM, col).Value = "Memenuhi"

        ' ── Hasil MS/TMS ─────────────────────────────────────
        ' TMS jika: KSWP tidak valid, atau tidak ada NIB, atau tidak ada SBU
        Dim msVal As String
        If UCase(kswpVal) = "TIDAK VALID" Or nibNomor = "" Or sbuNomor = "" Then
            msVal = "TMS"
        Else
            msVal = "MS"
        End If
        wsKK.Cells(ROW_HASIL_MS, col).Value = msVal

        ' Akumulasi peringatan peserta ini ke pesan global
        If warnPeserta <> "" Then
            warnMsg = warnMsg & namaP & ":" & vbCrLf & warnPeserta
        End If

    Next i

    ' Tampilkan ringkasan: sukses + peringatan field perlu cek manual
    Dim finalMsg As String
    finalMsg = "Data " & itemCount & " peserta berhasil dimuat dari Supabase."
    If warnMsg <> "" Then
        finalMsg = finalMsg & vbCrLf & vbCrLf & _
                   "⚠️ Field berikut perlu dicek manual (highlight kuning):" & vbCrLf & warnMsg
        MsgBox finalMsg, vbExclamation, "Muat KK Evaluasi"
    Else
        MsgBox finalMsg, vbInformation, "Muat KK Evaluasi"
    End If
    Exit Sub

ErrHandler:
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "MuatKKEvaluasi"
End Sub


' ============================================================
' HTTP: Fetch data dari Supabase (filter kode_tender)
' ============================================================
Private Function FetchKKEvaluasi(kodeTender As String) As String
    Dim url As String
    url = SB_URL & "/rest/v1/" & SB_TABLE & _
          "?kode_tender=eq." & kodeTender & _
          "&order=urutan.asc"

    Dim http As Object
    Set http = CreateObject("WinHttp.WinHttpRequest.5.1")

    On Error GoTo ErrHTTP
    http.Open "GET", url, False
    http.SetRequestHeader "apikey", SB_KEY
    http.SetRequestHeader "Authorization", "Bearer " & SB_KEY
    http.SetRequestHeader "Accept", "application/json"
    http.Send

    If http.Status = 200 Then
        FetchKKEvaluasi = http.ResponseText
    Else
        MsgBox "HTTP Error " & http.Status & ": " & http.StatusText & vbCrLf & _
               "URL: " & url, vbExclamation, "FetchKKEvaluasi Error"
        FetchKKEvaluasi = ""
    End If
    Exit Function

ErrHTTP:
    MsgBox "WinHTTP Error: " & Err.Number & " - " & Err.Description, vbCritical, "FetchKKEvaluasi Error"
    FetchKKEvaluasi = ""
End Function


' ============================================================
' PARSER: JSON array → array of field arrays
' ============================================================
Private Sub ParseKKJSON(json As String, ByRef items() As Variant, ByRef itemCount As Integer)
    Dim col As New Collection
    Dim pos As Long: pos = 1

    Do
        Dim braceStart As Long: braceStart = InStr(pos, json, "{")
        If braceStart = 0 Then Exit Do
        Dim depth As Integer: depth = 1
        Dim p As Long: p = braceStart + 1
        Do While p <= Len(json) And depth > 0
            Dim ch As String: ch = Mid(json, p, 1)
            If ch = "{" Then depth = depth + 1
            If ch = "}" Then depth = depth - 1
            p = p + 1
        Loop
        Dim braceEnd As Long: braceEnd = p - 1

        Dim obj As String
        obj = Mid(json, braceStart, braceEnd - braceStart + 1)

        Dim fields(40) As String
        fields(0)  = ExtractJSONVal(obj, "kode_tender")
        fields(1)  = ExtractJSONVal(obj, "urutan")
        fields(2)  = ExtractJSONVal(obj, "nama")
        fields(3)  = ExtractJSONVal(obj, "npwp")
        fields(4)  = ExtractJSONVal(obj, "nib_nomor")
        fields(5)  = ExtractJSONVal(obj, "nib_berlaku")
        fields(6)  = ExtractJSONVal(obj, "ss_nomor")
        fields(7)  = ExtractJSONVal(obj, "ss_berlaku")
        fields(8)  = ExtractJSONVal(obj, "ss_terverifikasi")
        fields(9)  = ExtractJSONVal(obj, "sbu_nomor")
        fields(10) = ExtractJSONVal(obj, "sbu_berlaku")
        fields(11) = ExtractJSONVal(obj, "sbu_kualifikasi")
        fields(12) = ExtractJSONVal(obj, "sbu_klasifikasi")
        fields(13) = ExtractJSONVal(obj, "sbu_subklas_label")
        fields(14) = ExtractJSONVal(obj, "pgl1_nama")
        fields(15) = ExtractJSONVal(obj, "pgl1_instansi")
        fields(16) = ExtractJSONVal(obj, "pgl1_nilai")
        fields(17) = ExtractJSONVal(obj, "pgl1_tanggal")
        fields(18) = ExtractJSONVal(obj, "pgl1_nomor")
        fields(19) = ExtractJSONVal(obj, "pgl2_nama")
        fields(20) = ExtractJSONVal(obj, "pgl2_instansi")
        fields(21) = ExtractJSONVal(obj, "pgl2_nilai")
        fields(22) = ExtractJSONVal(obj, "pgl2_tanggal")
        fields(23) = ExtractJSONVal(obj, "pgl2_nomor")
        fields(24) = ExtractJSONVal(obj, "skp")
        fields(25) = ExtractJSONVal(obj, "skp_catatan")
        fields(26) = ExtractJSONVal(obj, "skp_berbeda")
        fields(27) = ExtractJSONVal(obj, "kswp_status")
        fields(28) = ExtractJSONVal(obj, "akta_p_nomor")
        fields(29) = ExtractJSONVal(obj, "akta_p_tanggal")
        fields(30) = ExtractJSONVal(obj, "akta_p_notaris")
        fields(31) = ExtractJSONVal(obj, "akta_k_nomor")
        fields(32) = ExtractJSONVal(obj, "akta_k_tanggal")
        fields(33) = ExtractJSONVal(obj, "akta_k_notaris")
        fields(34) = ExtractJSONVal(obj, "pemilik_1")
        fields(35) = ExtractJSONVal(obj, "pemilik_2")
        fields(36) = ExtractJSONVal(obj, "pemilik_3")
        fields(37) = ExtractJSONVal(obj, "pemilik_4")
        fields(38) = ExtractJSONVal(obj, "kinerja_ada")
        fields(39) = ExtractJSONVal(obj, "kinerja_nilai")
        fields(40) = ExtractJSONVal(obj, "kinerja_kategori")

        col.Add fields
        pos = braceEnd + 1
    Loop

    itemCount = col.Count
    If itemCount = 0 Then Exit Sub

    ReDim items(itemCount - 1)
    Dim idx As Integer: idx = 0
    Dim itm As Variant
    For Each itm In col
        items(idx) = itm
        idx = idx + 1
    Next itm
End Sub


' ============================================================
' Helper: Ambil nilai field dari array parsed
' ============================================================
Private Function GetF(fields As Variant, fieldName As String) As String
    Select Case fieldName
        Case "kode_tender":       GetF = fields(0)
        Case "urutan":            GetF = fields(1)
        Case "nama":              GetF = fields(2)
        Case "npwp":              GetF = fields(3)
        Case "nib_nomor":         GetF = fields(4)
        Case "nib_berlaku":       GetF = fields(5)
        Case "ss_nomor":          GetF = fields(6)
        Case "ss_berlaku":        GetF = fields(7)
        Case "ss_terverifikasi":  GetF = fields(8)
        Case "sbu_nomor":         GetF = fields(9)
        Case "sbu_berlaku":       GetF = fields(10)
        Case "sbu_kualifikasi":   GetF = fields(11)
        Case "sbu_klasifikasi":   GetF = fields(12)
        Case "sbu_subklas_label": GetF = fields(13)
        Case "pgl1_nama":         GetF = fields(14)
        Case "pgl1_instansi":     GetF = fields(15)
        Case "pgl1_nilai":        GetF = fields(16)
        Case "pgl1_tanggal":      GetF = fields(17)
        Case "pgl1_nomor":        GetF = fields(18)
        Case "pgl2_nama":         GetF = fields(19)
        Case "pgl2_instansi":     GetF = fields(20)
        Case "pgl2_nilai":        GetF = fields(21)
        Case "pgl2_tanggal":      GetF = fields(22)
        Case "pgl2_nomor":        GetF = fields(23)
        Case "skp":               GetF = fields(24)
        Case "skp_catatan":       GetF = fields(25)
        Case "skp_berbeda":       GetF = fields(26)
        Case "kswp_status":       GetF = fields(27)
        Case "akta_p_nomor":      GetF = fields(28)
        Case "akta_p_tanggal":    GetF = fields(29)
        Case "akta_p_notaris":    GetF = fields(30)
        Case "akta_k_nomor":      GetF = fields(31)
        Case "akta_k_tanggal":    GetF = fields(32)
        Case "akta_k_notaris":    GetF = fields(33)
        Case "pemilik_1":         GetF = fields(34)
        Case "pemilik_2":         GetF = fields(35)
        Case "pemilik_3":         GetF = fields(36)
        Case "pemilik_4":         GetF = fields(37)
        Case "kinerja_ada":       GetF = fields(38)
        Case "kinerja_nilai":     GetF = fields(39)
        Case "kinerja_kategori":  GetF = fields(40)
        Case Else:                GetF = ""
    End Select
End Function


' ============================================================
' PARSER: Ekstrak nilai dari JSON string (single-level)
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
    Do While Mid(json, p, 1) = " ": p = p + 1: Loop

    If Mid(json, p, 1) = """" Then
        p = p + 1
        Dim q As Long: q = p
        Do While q <= Len(json)
            If Mid(json, q, 1) = """" And (q = 1 Or Mid(json, q - 1, 1) <> "\") Then Exit Do
            q = q + 1
        Loop
        ExtractJSONVal = Mid(json, p, q - p)
    ElseIf Mid(json, p, 4) = "null" Then
        ExtractJSONVal = ""
    Else
        Dim endPos As Long: endPos = p
        Do While endPos <= Len(json)
            Dim c As String: c = Mid(json, endPos, 1)
            If c = "," Or c = "}" Or c = "]" Then Exit Do
            endPos = endPos + 1
        Loop
        ExtractJSONVal = Trim(Mid(json, p, endPos - p))
    End If
End Function
