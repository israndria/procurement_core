Attribute VB_Name = "ModKodeUnik"
' ============================================================
' GENERATE KODE UNIK SURAT - Full Otomatis v2
' ============================================================
' Pola: [Pekerjaan+Objek].[TipeLokasi]_[NamaLokasi][Detail][Kecamatan]
' Contoh: PJ.D_HatungunRT06, RJ.GerilyaRKTU, PMJ.LimpanaKB

Private Function IsVowel(ch As String) As Boolean
    IsVowel = InStr(1, "aeiouAEIOU", ch) > 0
End Function

' Singkatkan nama lokasi (proper name panjang)
Private Function SingkatNama(nama As String) As String
    If Len(nama) <= 8 Then
        SingkatNama = nama
        Exit Function
    End If
    
    ' Ambil huruf pertama + konsonan utama
    Dim result As String
    result = Left(nama, 1)
    Dim i As Long
    Dim cnt As Long
    cnt = 1
    
    For i = 2 To Len(nama)
        Dim ch As String
        ch = Mid(nama, i, 1)
        If Not IsVowel(ch) Then
            result = result & LCase(ch)
            cnt = cnt + 1
        ElseIf i >= Len(nama) - 1 Then
            ' Tambah vokal di dekat akhir kata supaya terbaca
            result = result & LCase(ch)
            cnt = cnt + 1
        End If
        If cnt >= 6 Then Exit For
    Next i
    
    SingkatNama = result
End Function

' FUNGSI UTAMA: Generate kode unik dari nama paket
Public Function GenerateKodeUnikOtomatis(namaPaket As String) As String
    If Trim(namaPaket) = "" Then
        GenerateKodeUnikOtomatis = ""
        Exit Function
    End If
    
    ' Bersihkan input
    Dim s As String
    s = Replace(namaPaket, ",", " ")
    s = Replace(s, "  ", " ")
    s = Trim(s)
    
    Dim words() As String
    words = Split(s, " ")
    
    Dim partPekerjaan As String  ' PJ, Rehab, Pel, dll
    Dim partObjek As String      ' J, BJ, G, dll
    Dim partTipeLok As String    ' D, Ds, Kel
    Dim partNamaLok As String    ' Hatungun, Gerilya, dll
    Dim partDetail As String     ' RT06, RW01
    Dim partKec As String        ' TS, TU, CLU, B
    
    Dim idx As Long
    idx = LBound(words)
    Dim totalWords As Long
    totalWords = UBound(words)
    
    ' ===== FASE 1: Extract Pekerjaan =====
    ' Gabungkan kata kerja berurutan (Perbaikan/Peningkatan -> cukup 1 huruf P)
    Dim pekerjaanDone As Boolean
    pekerjaanDone = False
    Dim lastPekerjaanInitial As String
    
    Do While idx <= totalWords And Not pekerjaanDone
        Dim w As String
        w = CleanWord(words(idx))
        Dim lw As String
        lw = LCase(w)
        
        ' Skip separator
        If lw = "/" Or lw = "&" Or lw = "-" Then
            idx = idx + 1
            GoTo ContinuePekerjaan
        End If
        
        ' Angka di depan (3 Ruang Kelas)
        If IsNumeric(w) Then
            partPekerjaan = partPekerjaan & w
            idx = idx + 1
            GoTo ContinuePekerjaan
        End If
        
        Select Case lw
            Case "perbaikan", "peningkatan", "pelebaran", "rehabilitasi", _
                 "pembangunan", "pemeliharaan", "normalisasi", "pengaspalan", _
                 "lanjutan"
                
                Dim abbr As String
                Select Case lw
                    Case "perbaikan": abbr = "P"
                    Case "peningkatan": abbr = "P"
                    Case "pelebaran": abbr = "Pel"
                    Case "rehabilitasi": abbr = "R"
                    Case "pembangunan": abbr = "Pbn"
                    Case "pemeliharaan": abbr = "PM"
                    Case "normalisasi": abbr = "N"
                    Case "pengaspalan": abbr = "P"
                    Case "lanjutan": abbr = "L"
                End Select
                
                ' Hindari duplikat huruf (Perbaikan / Peningkatan -> P, bukan PP)
                If abbr <> lastPekerjaanInitial Or Len(abbr) > 1 Then
                    partPekerjaan = partPekerjaan & abbr
                End If
                lastPekerjaanInitial = abbr
                idx = idx + 1
            Case Else
                pekerjaanDone = True
        End Select
ContinuePekerjaan:
    Loop
    
    ' ===== FASE 2: Extract Objek =====
    If idx <= totalWords Then
        w = CleanWord(words(idx))
        lw = LCase(w)
        
        Select Case lw
            Case "jalan": partObjek = "J": idx = idx + 1
                ' Cek "Bahu Jalan" -> sudah lewat, handle "Jalan" saja
            Case "bahu"
                partObjek = "BJ"
                idx = idx + 1
                ' Skip "Jalan" berikutnya
                If idx <= totalWords And LCase(CleanWord(words(idx))) = "jalan" Then idx = idx + 1
            Case "gedung": partObjek = "G": idx = idx + 1
            Case "siring": partObjek = "S": idx = idx + 1
            Case "drainase": partObjek = "Dr": idx = idx + 1
            Case "jembatan": partObjek = "Jbtn": idx = idx + 1
            Case "irigasi": partObjek = "Iri": idx = idx + 1
            Case "gorong-gorong", "gorong": partObjek = "GG": idx = idx + 1
            Case "saluran": partObjek = "Sal": idx = idx + 1
            Case "ruang"
                partObjek = "R"
                idx = idx + 1
                If idx <= totalWords And LCase(CleanWord(words(idx))) = "kelas" Then
                    partObjek = partObjek & ".KLS"
                    idx = idx + 1
                End If
            Case "kantor": partObjek = "Kntr": idx = idx + 1
            Case "trotoar": partObjek = "Trt": idx = idx + 1
            Case "talud": partObjek = "Tld": idx = idx + 1
            Case "turap": partObjek = "Trp": idx = idx + 1
            Case "bendung": partObjek = "Bnd": idx = idx + 1
            Case "sungai": partObjek = "S": idx = idx + 1
        End Select
    End If
    
    ' ===== FASE 3: Scan sisa kata =====
    Dim inKecamatan As Boolean
    inKecamatan = False
    
    Do While idx <= totalWords
        w = CleanWord(words(idx))
        lw = LCase(w)
        idx = idx + 1
        
        If Len(w) = 0 Then GoTo ContinueScan
        
        ' Skip separator
        If lw = "/" Or lw = "&" Or lw = "-" Then GoTo ContinueScan
        
        ' Skip kata penghubung
        If lw = "menuju" Or lw = "dari" Or lw = "sampai" Or lw = "s/d" Or _
           lw = "ke" Or lw = "di" Or lw = "dan" Or lw = "yang" Or _
           lw = "dengan" Or lw = "melalui" Or lw = "baru" Then
            GoTo ContinueScan
        End If
        
        ' Skip kata deskriptif panjang
        If lw = "usulan" Or lw = "fisik" Or lw = "prioritas" Or _
           lw = "intervensi" Or lw = "kebakaran" Or lw = "kabupaten" Then
            GoTo ContinueScan
        End If
        
        ' Nama kabupaten (skip 1 kata berikutnya)
        If lw = "kab." Or lw = "kab" Then
            ' Skip nama kab
            If idx <= totalWords Then idx = idx + 1
            GoTo ContinueScan
        End If
        
        ' Kecamatan
        If lw = "kecamatan" Or lw = "kec." Or lw = "kec" Then
            inKecamatan = True
            GoTo ContinueScan
        End If
        
        If inKecamatan Then
            ' Ambil huruf pertama setiap kata kecamatan
            If lw <> "kabupaten" And lw <> "kab." And lw <> "kab" Then
                partKec = partKec & UCase(Left(w, 1))
            Else
                If idx <= totalWords Then idx = idx + 1
            End If
            GoTo ContinueScan
        End If
        
        ' Tipe lokasi
        If lw = "desa" Then
            partTipeLok = "D"
            GoTo ContinueScan
        End If
        If lw = "kelurahan" Or lw = "kel." Or lw = "kel" Then
            partTipeLok = "Kel"
            GoTo ContinueScan
        End If
        
        ' RT/RW
        If Left(lw, 2) = "rt" Then
            Dim rtNum As String
            rtNum = Replace(Mid(w, 3), ".", "")
            rtNum = Replace(rtNum, " ", "")
            If IsNumeric(rtNum) And rtNum <> "" Then
                partDetail = "RT" & Format(CInt(rtNum), "0")
            Else
                partDetail = "RT"
            End If
            GoTo ContinueScan
        End If
        If Left(lw, 2) = "rw" Then
            Dim rwNum As String
            rwNum = Replace(Mid(w, 3), ".", "")
            If IsNumeric(rwNum) And rwNum <> "" Then
                partDetail = partDetail & "RW" & Format(CInt(rwNum), "0")
            Else
                partDetail = partDetail & "RW"
            End If
            GoTo ContinueScan
        End If
        
        ' Angka setelah RT/RW
        If IsNumeric(Replace(w, ".", "")) And partDetail <> "" Then
            partDetail = partDetail & Replace(w, ".", "")
            GoTo ContinueScan
        End If
        
        ' Kata umum yang punya singkatan
        Dim ku As String
        ku = ""
        Select Case lw
            Case "pondok": ku = "Pndk"
            Case "pesantren": ku = "Pst"
            Case "sungai": ku = "S"
            Case "simpang": ku = "Smpg"
            Case "datu": ku = "D"
            Case "gunung": ku = "G"
            Case "tanjung": ku = "Tjg"
            Case "kubah": ku = "Kbh"
            Case "very": ku = "V"
            Case "lingkar": ku = "Lnkr"
            Case "lima": ku = ""
        End Select
        
        If ku <> "" Then
            partNamaLok = partNamaLok & ku
            GoTo ContinueScan
        ElseIf ku = "" And ( _
            lw = "lima") Then
            GoTo ContinueScan
        End If
        
        ' Proper name (nama lokasi)
        partNamaLok = partNamaLok & SingkatNama(w)
        
ContinueScan:
    Loop
    
    ' ===== GABUNGKAN =====
    Dim hasil As String
    hasil = partPekerjaan & partObjek
    
    ' Separator titik
    If partTipeLok <> "" Or partNamaLok <> "" Then
        hasil = hasil & "."
    End If
    
    ' Tipe lokasi
    If partTipeLok <> "" Then
        hasil = hasil & partTipeLok & "_"
    End If
    
    ' Nama lokasi
    hasil = hasil & partNamaLok
    
    ' Detail RT/RW
    hasil = hasil & partDetail
    
    ' Kecamatan
    If partKec <> "" Then
        hasil = hasil & partKec
    End If
    
    GenerateKodeUnikOtomatis = hasil
End Function

Private Function CleanWord(w As String) As String
    Dim result As String
    result = Trim(w)
    ' Hapus trailing koma
    If Len(result) > 0 Then
        If Right(result, 1) = "," Then
            result = Left(result, Len(result) - 1)
        End If
    End If
    CleanWord = result
End Function

' MACRO: Generate kode unik dari nama paket
Public Sub GenerateKodeUnik()
Attribute GenerateKodeUnik.VB_ProcData.VB_Invoke_Func = "R\n14"
    Dim wsInput As Worksheet
    On Error Resume Next
    Set wsInput = ThisWorkbook.Sheets("1. Input Data")
    On Error GoTo 0
    
    If wsInput Is Nothing Then
        MsgBox "Sheet '1. Input Data' tidak ditemukan!", vbExclamation
        Exit Sub
    End If
    
    Dim namaPaket As String
    If Not IsError(wsInput.Range("F6").Value) Then
        namaPaket = CStr(wsInput.Range("F6").Value)
    End If
    
    If namaPaket = "" Then
        namaPaket = InputBox("Masukkan nama paket:", "Generate Kode Unik")
        If namaPaket = "" Then Exit Sub
    End If
    
    ' Generate
    Dim kodeUnik As String
    kodeUnik = GenerateKodeUnikOtomatis(namaPaket)
    
    ' Tampilkan untuk review/edit
    Dim hasil As String
    hasil = InputBox( _
        "Nama Paket:" & vbCrLf & namaPaket & vbCrLf & vbCrLf & _
        "Kode Unik (edit jika perlu):", _
        "Generate Kode Unik Surat", _
        kodeUnik)
    
    If StrPtr(hasil) = 0 Then Exit Sub
    If hasil = "" Then Exit Sub
    
    ' Simpan ke cell
    On Error Resume Next
    wsInput.Unprotect Password:="pokja2026"
    On Error GoTo 0
    
    wsInput.Range("E9").Value = hasil
    
    On Error Resume Next
    wsInput.Protect Password:="pokja2026", AllowFormattingCells:=True
    On Error GoTo 0
    
    MsgBox "Kode unik disimpan di E9:" & vbCrLf & hasil, vbInformation
    
    wsInput.Activate
    wsInput.Range("E9").Select
End Sub

' Fungsi formula: bisa dipanggil dari cell
' Contoh: =GenerateKodeUnikOtomatis(F6)

