Attribute VB_Name = "ModTerbilang"
' Fungsi TERBILANG - Mengubah angka menjadi kata dalam Bahasa Indonesia
' Sudah built-in, tidak perlu file .xlam external lagi

Public Function terbilang(ByVal Angka As Variant) As String
    Dim Bilangan As Double
    Dim hasil As String
    
    On Error GoTo ErrHandler
    
    If IsEmpty(Angka) Or Angka = "" Then
        terbilang = ""
        Exit Function
    End If
    
    Bilangan = CDbl(Angka)
    
    If Bilangan < 0 Then
        terbilang = "Minus " & terbilang_helper(Abs(Bilangan))
    Else
        terbilang = terbilang_helper(Bilangan)
    End If
    
    Exit Function
ErrHandler:
    terbilang = ""
End Function

Private Function terbilang_helper(ByVal Angka As Double) As String
    Dim Satuan As Variant
    Satuan = Array("", "Satu", "Dua", "Tiga", "Empat", "Lima", "Enam", "Tujuh", "Delapan", "Sembilan", "Sepuluh", "Sebelas")
    
    If Angka < 12 Then
        terbilang_helper = Satuan(Int(Angka))
    ElseIf Angka < 20 Then
        terbilang_helper = terbilang_helper(Angka - 10) & " Belas"
    ElseIf Angka < 100 Then
        terbilang_helper = terbilang_helper(Int(Angka / 10)) & " Puluh " & terbilang_helper(Int(Angka) Mod 10)
    ElseIf Angka < 200 Then
        terbilang_helper = "Seratus " & terbilang_helper(Angka - 100)
    ElseIf Angka < 1000 Then
        terbilang_helper = terbilang_helper(Int(Angka / 100)) & " Ratus " & terbilang_helper(Int(Angka) Mod 100)
    ElseIf Angka < 2000 Then
        terbilang_helper = "Seribu " & terbilang_helper(Angka - 1000)
    ElseIf Angka < 1000000 Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000)) & " Ribu " & terbilang_helper(Int(Angka) Mod 1000)
    ElseIf Angka < 1000000000# Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000000)) & " Juta " & terbilang_helper(Int(Angka) Mod 1000000)
    ElseIf Angka < 1000000000000# Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000000000#)) & " Miliar " & terbilang_helper(Int(Angka) Mod 1000000000#)
    ElseIf Angka < 1E+15 Then
        terbilang_helper = terbilang_helper(Int(Angka / 1000000000000#)) & " Triliun " & terbilang_helper(Int(Angka) Mod 1000000000000#)
    End If
    
    terbilang_helper = Trim(terbilang_helper)
End Function
