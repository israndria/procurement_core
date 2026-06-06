Attribute VB_Name = "ModKodeUnikPL"
' ============================================================
' GENERATE KODE UNIK SURAT - Mode Pengadaan Langsung (PL)
' ============================================================
' Identik dengan ModKodeUnik (Tender) tapi:
'   - Baca nama_paket dari @ Master Data C5 (bukan 1. Input Data F6)
'   - Baca kode_paket dari @ Master Data C3 (bukan 1. Input Data E5)
'   - Simpan kode_unik ke @ Master Data F2
'   - Upsert ke Supabase draft_paket_pl via upsert_kode_unik_pl.py

Private Function IsVowelPL(ch As String) As Boolean
    IsVowelPL = InStr(1, "aeiouAEIOU", ch) > 0
End Function

Private Function SingkatNamaPL(nama As String) As String
    If Len(nama) <= 8 Then
        SingkatNamaPL = nama
        Exit Function
    End If

    Dim result As String
    result = Left(nama, 1)
    Dim i As Long
    Dim cnt As Long
    cnt = 1

    For i = 2 To Len(nama)
        Dim ch As String
        ch = Mid(nama, i, 1)
        If Not IsVowelPL(ch) Then
            result = result & LCase(ch)
            cnt = cnt + 1
        ElseIf i >= Len(nama) - 1 Then
            result = result & LCase(ch)
            cnt = cnt + 1
        End If
        If cnt >= 6 Then Exit For
    Next i

    SingkatNamaPL = result
End Function

Private Function CleanWordPL(w As String) As String
    Dim result As String
    result = Trim(w)
    If Len(result) > 0 Then
        If Right(result, 1) = "," Then
            result = Left(result, Len(result) - 1)
        End If
    End If
    CleanWordPL = result
End Function

Public Function GenerateKodeUnikPL(namaPaket As String) As String
    If Trim(namaPaket) = "" Then
        GenerateKodeUnikPL = ""
        Exit Function
    End If

    Dim s As String
    s = Replace(namaPaket, ",", " ")
    s = Replace(s, "  ", " ")
    s = Trim(s)

    Dim words() As String
    words = Split(s, " ")

    Dim partPekerjaan As String
    Dim partObjek As String
    Dim partTipeLok As String
    Dim partNamaLok As String
    Dim partDetail As String
    Dim partKec As String

    Dim idx As Long
    idx = LBound(words)
    Dim totalWords As Long
    totalWords = UBound(words)

    ' ===== FASE 1: Extract Pekerjaan =====
    Dim pekerjaanDone As Boolean
    pekerjaanDone = False
    Dim lastPekerjaanInitial As String

    Do While idx <= totalWords And Not pekerjaanDone
        Dim w As String
        w = CleanWordPL(words(idx))
        Dim lw As String
        lw = LCase(w)

        If lw = "/" Or lw = "&" Or lw = "-" Then
            idx = idx + 1
            GoTo ContinuePekerjaan
        End If

        If IsNumeric(w) Then
            partPekerjaan = partPekerjaan & w
            idx = idx + 1
            GoTo ContinuePekerjaan
        End If

        Select Case lw
            Case "perbaikan", "peningkatan", "pelebaran", "rehabilitasi", _
                 "pembangunan", "pemeliharaan", "normalisasi", "pengaspalan", _
                 "lanjutan", "pengawasan", "perencanaan", "studi", "kajian", _
                 "konsultansi", "konsultan"

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
                    Case "pengawasan": abbr = "Pgw"
                    Case "perencanaan": abbr = "Prn"
                    Case "studi": abbr = "St"
                    Case "kajian": abbr = "Kjn"
                    Case "konsultansi", "konsultan": abbr = "Ksl"
                End Select

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
        w = CleanWordPL(words(idx))
        lw = LCase(w)

        Select Case lw
            Case "jalan": partObjek = "J": idx = idx + 1
            Case "bahu"
                partObjek = "BJ"
                idx = idx + 1
                If idx <= totalWords And LCase(CleanWordPL(words(idx))) = "jalan" Then idx = idx + 1
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
                If idx <= totalWords And LCase(CleanWordPL(words(idx))) = "kelas" Then
                    partObjek = partObjek & ".KLS"
                    idx = idx + 1
                End If
            Case "kantor": partObjek = "Kntr": idx = idx + 1
            Case "trotoar": partObjek = "Trt": idx = idx + 1
            Case "talud": partObjek = "Tld": idx = idx + 1
            Case "turap": partObjek = "Trp": idx = idx + 1
            Case "bendung": partObjek = "Bnd": idx = idx + 1
            Case "sungai": partObjek = "S": idx = idx + 1
            Case "kak", "dokumen": partObjek = "": idx = idx + 1
        End Select
    End If

    ' ===== FASE 3: Scan sisa kata =====
    Dim inKecamatan As Boolean
    inKecamatan = False

    Do While idx <= totalWords
        w = CleanWordPL(words(idx))
        lw = LCase(w)
        idx = idx + 1

        If Len(w) = 0 Then GoTo ContinueScan

        If lw = "/" Or lw = "&" Or lw = "-" Then GoTo ContinueScan

        If lw = "menuju" Or lw = "dari" Or lw = "sampai" Or lw = "s/d" Or _
           lw = "ke" Or lw = "di" Or lw = "dan" Or lw = "yang" Or _
           lw = "dengan" Or lw = "melalui" Or lw = "baru" Or lw = "untuk" Or _
           lw = "teknis" Or lw = "persiapan" Or lw = "pengadaan" Then
            GoTo ContinueScan
        End If

        If lw = "usulan" Or lw = "fisik" Or lw = "prioritas" Or _
           lw = "intervensi" Or lw = "kebakaran" Or lw = "kabupaten" Then
            GoTo ContinueScan
        End If

        If lw = "kab." Or lw = "kab" Then
            If idx <= totalWords Then idx = idx + 1
            GoTo ContinueScan
        End If

        If lw = "kecamatan" Or lw = "kec." Or lw = "kec" Then
            inKecamatan = True
            GoTo ContinueScan
        End If

        If inKecamatan Then
            If lw <> "kabupaten" And lw <> "kab." And lw <> "kab" Then
                partKec = partKec & UCase(Left(w, 1))
            Else
                If idx <= totalWords Then idx = idx + 1
            End If
            GoTo ContinueScan
        End If

        If lw = "desa" Then
            partTipeLok = "D"
            GoTo ContinueScan
        End If
        If lw = "kelurahan" Or lw = "kel." Or lw = "kel" Then
            partTipeLok = "Kel"
            GoTo ContinueScan
        End If

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

        If IsNumeric(Replace(w, ".", "")) And partDetail <> "" Then
            partDetail = partDetail & Replace(w, ".", "")
            GoTo ContinueScan
        End If

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
            Case "rsud", "puskesmas": ku = UCase(lw)
            Case "dinas": ku = "Din"
        End Select

        If ku <> "" Then
            partNamaLok = partNamaLok & ku
            GoTo ContinueScan
        ElseIf ku = "" And lw = "lima" Then
            GoTo ContinueScan
        End If

        partNamaLok = partNamaLok & SingkatNamaPL(w)

ContinueScan:
    Loop

    ' ===== GABUNGKAN =====
    Dim hasil As String
    hasil = partPekerjaan & partObjek

    If partTipeLok <> "" Or partNamaLok <> "" Then
        hasil = hasil & "."
    End If

    If partTipeLok <> "" Then
        hasil = hasil & partTipeLok & "_"
    End If

    hasil = hasil & partNamaLok
    hasil = hasil & partDetail

    If partKec <> "" Then
        hasil = hasil & partKec
    End If

    GenerateKodeUnikPL = hasil
End Function

' MACRO UTAMA: Generate + simpan kode unik PL
Public Sub GenerateKodeUnikPaketPL()
Attribute GenerateKodeUnikPaketPL.VB_ProcData.VB_Invoke_Func = "R\n14"
    Dim wsMD As Worksheet
    On Error Resume Next
    Set wsMD = ThisWorkbook.Sheets("@ Master Data")
    On Error GoTo 0

    If wsMD Is Nothing Then
        MsgBox "Sheet '@ Master Data' tidak ditemukan!", vbExclamation
        Exit Sub
    End If

    ' Baca nama paket dari C5 dan kode paket dari C3
    Dim namaPaket As String
    Dim kodePaket As String
    On Error Resume Next
    namaPaket = CStr(wsMD.Cells(5, 3).Value)  ' C5
    kodePaket  = CStr(wsMD.Cells(3, 3).Value)  ' C3
    On Error GoTo 0

    If namaPaket = "" Then
        namaPaket = InputBox("Masukkan nama paket:", "Generate Kode Unik PL")
        If namaPaket = "" Then Exit Sub
    End If

    ' Generate kode unik
    Dim kodeUnik As String
    kodeUnik = GenerateKodeUnikPL(namaPaket)

    ' Tampilkan untuk review/edit
    Dim hasil As String
    hasil = InputBox( _
        "Nama Paket:" & vbCrLf & namaPaket & vbCrLf & vbCrLf & _
        "Kode Unik (edit jika perlu):", _
        "Generate Kode Unik Surat PL", _
        kodeUnik)

    If StrPtr(hasil) = 0 Then Exit Sub
    If hasil = "" Then Exit Sub

    ' Simpan ke F2 @ Master Data
    On Error Resume Next
    wsMD.Unprotect Password:="pokja2026"
    On Error GoTo 0

    wsMD.Range("F2").Value = hasil

    On Error Resume Next
    wsMD.Protect Password:="pokja2026", AllowFormattingCells:=True
    On Error GoTo 0

    ' Upsert ke Supabase via Python (non-blocking)
    If kodePaket <> "" Then
        ' Cari WPy64 dengan naik ke atas dari lokasi Excel sampai ketemu V19_Scheduler
        Dim xlDir As String
        xlDir = ThisWorkbook.Path
        Dim wpy64Dir As String
        Dim testPath As String
        Dim lvl As Integer
        Dim curDir As String
        curDir = xlDir
        wpy64Dir = ""
        For lvl = 1 To 6
            testPath = curDir & "\V19_Scheduler\WPy64-313110\python\python.exe"
            If Dir(testPath) <> "" Then
                wpy64Dir = curDir & "\V19_Scheduler\WPy64-313110"
                Exit For
            End If
            ' Naik satu level
            Dim parentPos As Long
            parentPos = InStrRev(curDir, "\")
            If parentPos <= 3 Then Exit For
            curDir = Left(curDir, parentPos - 1)
        Next lvl

        If wpy64Dir = "" Then
            MsgBox "Python WPy64 tidak ditemukan. Kode unik sudah tersimpan di Excel saja.", vbExclamation
        Else
            Dim wshUp As Object
            Set wshUp = CreateObject("WScript.Shell")
            Dim cmdUp As String
            cmdUp = """" & wpy64Dir & "\python\python.exe"" """ & wpy64Dir & "\upsert_kode_unik_pl.py"" """ & kodePaket & """ """ & hasil & """"
            wshUp.Run cmdUp, 0, False
            Set wshUp = Nothing
        End If
    End If

    MsgBox "Kode unik PL disimpan di F2 @ Master Data:" & vbCrLf & hasil, vbInformation

    wsMD.Activate
    wsMD.Range("F2").Select
End Sub
