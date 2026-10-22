Attribute VB_Name = "modKS_AKTAR"
Option Explicit

' TOPLU AKTAR: geçmiþ yýllarýn verisini Excel'den yapýþtýrýp form kurallarýyla doðrulayarak aktarýr.
' Hatalý satýrlar aktarýlmaz; SONUÇ sütununda nedeni yazar. Aktarýlan satýr ikinci kez aktarýlmaz.

Private Const T_BASLIK As Long = 5
Private Const T_ILK As Long = 6
Private Const T_KOLON As Long = 2
Private Const T_GENIS As Long = 20
Private Const T_SONSATIR As Long = 60000

Private Function KS_Sablon(ByVal tur As String) As Variant
    Select Case tur
        Case "MALZEME"
            KS_Sablon = Array("Malzeme Kodu *", "Malzeme Adý *", "Kategori", "Birim", "Üretici", _
                "Birim Fiyat", "Para Birimi", "Garanti (Ay)", "Not")
        Case "SEVKÝYAT"
            KS_Sablon = Array("Sevkçi *", "Sevk Yeri *", "Malzeme Kodu *", "Adet *", "Birim Fiyat (boþ = liste)", _
                "Para Birimi", "Planlanan Teslim *", "Gerçek Geliþ", "Garanti (Ay)", "Belge / SAS No", _
                "Seri / Lot No", "Not")
        Case Else
            KS_Sablon = Array("Arýza Tarihi *", "Malzeme Kodu *", "Seri No", "Arýzalý Adet *", "Arýza Tipi", _
                "Arýza Sebebi / Açýklama", "Baðlý Sevk No (boþ = otomatik)", "Garanti Devam? (EVET/HAYIR)", _
                "Üreticiye Bildirim", "Mail Yazýþmasý", "Mail Dosyasý (yol)", "Üretici Dönüþ", _
                "Arýza Devam Ediyor mu?", "Çözüldü mü?", "Çözüm Tarihi", "Çözüm Þekli", "Tamir Maliyeti (TL)", _
                "Maliyeti Karþýlayan", "Not")
    End Select
End Function

Private Function KS_TarihKolonlari(ByVal tur As String) As Variant
    Select Case tur
        Case "MALZEME": KS_TarihKolonlari = Array()
        Case "SEVKÝYAT": KS_TarihKolonlari = Array(6, 7)
        Case Else: KS_TarihKolonlari = Array(0, 8, 11, 14)
    End Select
End Function

Public Sub KS_AktarKur()
    Dim ws As Worksheet, b As Variant, i As Long, sol As Double
    Set ws = KS_WS(SH_AKTAR)
    KS_SayfaSifirla ws
    ws.Columns("B:U").ColumnWidth = 15
    ws.Columns("V").ColumnWidth = 2
    KS_VeriSayfaUst ws, "TOPLU AKTARIM", _
        "Þablonu seçin, verileri 6. satýrdan itibaren DEÐER olarak yapýþtýrýn, KONTROL ET + AKTAR'a basýn.", _
        Array("SAP / YILLIK EXCEL", "KS_SAP_AKTAR", "yesil", "KONTROL ET + AKTAR", "KS_AKTAR_CALISTIR", "ana", _
            "TEMÝZLE", "KS_AKTAR_TEMIZLE", "cizgi")
    ws.Rows(3).RowHeight = 34
    b = Array("MALZEME ÞABLONU", "KS_SABLON_MALZEME", "SEVKÝYAT ÞABLONU", "KS_SABLON_SEVK", _
        "ARIZA ÞABLONU", "KS_SABLON_ARIZA")
    sol = ws.Cells(1, 2).Left
    For i = 0 To 2
        KS_Buton ws, "ksA_S" & i, CStr(b(i * 2)), sol + i * 146, ws.Rows(3).Top + 5, 140, 24, CStr(b(i * 2 + 1)), "cizgi"
    Next i
    With ws.Range("F3")
        .Value = "TÜR"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("G3")
        .Value = "SEVKÝYAT"
        .Font.Bold = True
        .Font.Color = C_LAC()
        .Interior.Color = C_ACIKMAVI()
        .HorizontalAlignment = xlCenter
    End With
    KS_AdHucre "aTur", ws.Range("G3")
    With ws.Range("I3")
        .Value = "EKSÝK TANIMLARI EKLE"
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("J3")
        .Value = "HAYIR"
        .Font.Bold = True
        .Font.Color = C_LAC()
        .Interior.Color = vbWhite
        .HorizontalAlignment = xlCenter
        .Locked = False
        .BorderAround xlContinuous, xlThin, , RGB(190, 204, 222)
    End With
    KS_AdHucre "aOto", ws.Range("J3")
    KS_Liste ws.Range("J3"), "lstEvetHayir", True
    With ws.Range("K3")
        .Value = "EVET: listede olmayan sevkçi / sevk yerini TANIMLAR'a ekler."
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = C_GRI()
    End With
    ws.Rows(4).RowHeight = 6
    KS_SablonCiz ws, "SEVKÝYAT"
    KS_Pencere ws, 90, T_BASLIK
End Sub

Private Sub KS_SablonCiz(ByVal ws As Worksheet, ByVal tur As String)
    Dim h As Variant, n As Long, alan As Range, fc As Object, sonuc As Range, k As Variant
    h = KS_Sablon(tur)
    n = UBound(h) + 1
    With ws.Range(ws.Cells(T_BASLIK, T_KOLON), ws.Cells(T_SONSATIR, T_KOLON + T_GENIS))
        .FormatConditions.Delete
        .Clear
        .Font.Name = KS_FONT
        .Font.Size = 9.5
    End With
    KS_TabloBaslik ws.Range(ws.Cells(T_BASLIK, T_KOLON), ws.Cells(T_BASLIK, T_KOLON + n - 1)), h
    Set sonuc = ws.Cells(T_BASLIK, T_KOLON + n)
    KS_TabloBaslik sonuc, Array("SONUÇ")
    sonuc.Interior.Color = C_TEAL()
    ws.Range(ws.Columns(T_KOLON), ws.Columns(T_KOLON + T_GENIS)).ColumnWidth = 15
    ws.Columns(T_KOLON + n).ColumnWidth = 60
    Set alan = ws.Range(ws.Cells(T_ILK, T_KOLON), ws.Cells(T_SONSATIR, T_KOLON + n - 1))
    alan.Interior.Color = vbWhite
    alan.Locked = False
    alan.NumberFormat = "@"
    For Each k In KS_TarihKolonlari(tur)
        ws.Range(ws.Cells(T_ILK, T_KOLON + k), ws.Cells(T_SONSATIR, T_KOLON + k)).NumberFormat = KS_FmtTarih()
    Next k
    Set sonuc = ws.Range(ws.Cells(T_ILK, T_KOLON + n), ws.Cells(T_SONSATIR, T_KOLON + n))
    sonuc.Interior.Color = RGB(240, 246, 250)
    sonuc.Font.Size = 9
    Set fc = sonuc.FormatConditions.Add(Type:=xlTextString, String:="AKTARILDI", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_YESIL(): fc.Font.Bold = True
    Set fc = sonuc.FormatConditions.Add(Type:=xlTextString, String:="HATA", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_KIRMIZI(): fc.Font.Bold = True
    ws.Range(ws.Cells(T_ILK, T_KOLON), ws.Cells(T_SONSATIR, T_KOLON + n)).Borders(xlInsideHorizontal).Color = RGB(230, 236, 243)
    ws.Rows(T_BASLIK).RowHeight = 36
    KS_F("aTur").Value = tur
End Sub

Public Sub KS_SABLON_MALZEME()
    KS_SablonDegistir "MALZEME"
End Sub

Public Sub KS_SABLON_SEVK()
    KS_SablonDegistir "SEVKÝYAT"
End Sub

Public Sub KS_SABLON_ARIZA()
    KS_SablonDegistir "ARIZA"
End Sub

Private Sub KS_SablonDegistir(ByVal tur As String)
    Dim ws As Worksheet
    On Error GoTo Hata
    Set ws = KS_WS(SH_AKTAR)
    If KS_SonSatir(ws) >= T_ILK Then
        If MsgBox("Aktarým alanýnda veri var. Þablon deðiþince alan temizlenecek. Devam edilsin mi?", _
            vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    KS_Basla
    KS_Ac ws
    KS_SablonCiz ws, tur
    KS_Kapat ws
    KS_Bitir
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Þablon deðiþtirilemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Private Function KS_SonSatir(ByVal ws As Worksheet) As Long
    Dim c As Long, r As Long
    For c = T_KOLON To T_KOLON + T_GENIS
        r = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If r > KS_SonSatir Then KS_SonSatir = r
    Next c
    If KS_SonSatir < T_ILK Then KS_SonSatir = T_ILK - 1
End Function

Public Sub KS_AKTAR_TEMIZLE()
    Dim ws As Worksheet, son As Long
    On Error GoTo Hata
    Set ws = KS_WS(SH_AKTAR)
    son = KS_SonSatir(ws)
    If son < T_ILK Then Exit Sub
    If MsgBox("Aktarým alanýndaki " & son - T_ILK + 1 & " satýr temizlensin mi?" & vbCrLf & _
        "(Aktarýlmýþ kayýtlar tablolarda kalýr.)", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    KS_Ac ws
    ws.Range(ws.Cells(T_ILK, T_KOLON), ws.Cells(son, T_KOLON + T_GENIS)).ClearContents
    KS_Kapat ws
    Exit Sub
Hata:
    MsgBox "Temizlenemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Public Sub KS_AKTAR_CALISTIR()
    Dim ws As Worksheet, tur As String, h As Variant, n As Long, son As Long, v As Variant, g() As Variant
    Dim i As Long, j As Long, m As Long, hataN As Long, atla As Long, satir As Variant, hata As String, uyari As String
    Dim cikti() As Variant, sonuc() As Variant, kol As Long, lo As ListObject, tablo As String, bos As Boolean
    Dim adim As String
    On Error GoTo Hata
    Set ws = KS_WS(SH_AKTAR)
    tur = KS_Metin(KS_FV("aTur"))
    h = KS_Sablon(tur)
    n = UBound(h) + 1
    son = KS_SonSatir(ws)
    If son < T_ILK Then
        MsgBox "Aktarýlacak satýr yok. Verileri " & T_ILK & ". satýrdan itibaren yapýþtýrýn.", vbInformation, KS_AD
        Exit Sub
    End If
    Select Case tur
        Case "MALZEME": tablo = T_MALZ: kol = M_KOLON
        Case "SEVKÝYAT": tablo = T_SEVK: kol = S_KOLON
        Case Else: tablo = T_ARIZA: kol = A_KOLON
    End Select
    If MsgBox(son - T_ILK + 1 & " satýr " & tur & " olarak kontrol edilip aktarýlacak." & vbCrLf & _
        "Önce dosyanýn yedeði alýnacak. Devam edilsin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub

    adim = "Yedek"
    KS_YedekKopya "AKTARIM_ONCESI"
    adim = "Hazýrlýk"
    KS_Basla
    KS_BaglamYukle
    KS_OtoTanimAyarla (KS_Metin(KS_FV("aOto")) = "EVET")
    v = ws.Range(ws.Cells(T_ILK, T_KOLON), ws.Cells(son, T_KOLON + n)).Value
    ReDim cikti(1 To UBound(v, 1), 1 To kol)
    ReDim sonuc(1 To UBound(v, 1), 1 To 1)
    ReDim g(0 To n - 1)

    For i = 1 To UBound(v, 1)
        adim = "Satýr " & (T_ILK + i - 1)
        If i Mod 200 = 0 Then KS_Durum "Kontrol: " & i & " / " & UBound(v, 1)
        sonuc(i, 1) = v(i, n + 1)
        bos = True
        For j = 0 To n - 1
            g(j) = v(i, j + 1)
            If Not KS_Bos(g(j)) Then bos = False
        Next j
        If bos Then
            sonuc(i, 1) = Empty
        ElseIf Left$(KS_Metin(v(i, n + 1)), 9) = "AKTARILDI" Then
            atla = atla + 1
        Else
            hata = "": uyari = ""
            Select Case tur
                Case "MALZEME"
                    If KS_MalzemeHazirla(g, 0, satir, hata) Then KS_MalzemeKaydiEkle satir(1, M_KOD)
                Case "SEVKÝYAT"
                    If KS_SevkHazirla(g, 0, satir, hata, uyari) Then
                        If Len(uyari) > 0 Then
                            hata = "Olasý mükerrer. " & uyari
                            KS_SevkSayacGeri
                        End If
                    End If
                Case Else
                    KS_ArizaHazirla g, 0, Empty, satir, hata
            End Select
            If Len(hata) > 0 Then
                hataN = hataN + 1
                sonuc(i, 1) = "HATA: " & hata
            Else
                m = m + 1
                For j = 1 To kol
                    cikti(m, j) = satir(1, j)
                Next j
                sonuc(i, 1) = "AKTARILDI  " & KS_Metin(satir(1, 1)) & "  (" & Format$(Now, "dd.mm.yyyy hh:nn") & ")"
            End If
        End If
    Next i

    adim = "Tabloya yazma"
    KS_Durum "Kayýtlar yazýlýyor..."
    Set lo = KS_LOZ(tablo)
    If m > 0 Then
        KS_Ac lo.Parent
        KS_TopluEkle lo, cikti, m
        KS_Kapat lo.Parent
    End If
    KS_Ac ws
    ws.Range(ws.Cells(T_ILK, T_KOLON + n), ws.Cells(son, T_KOLON + n)).Value = sonuc
    KS_Kapat ws
    adim = "Yeniden hesaplama"
    KS_Durum "Yeniden hesaplanýyor..."
    KS_HesaplaTumu
    KS_OtoTanimAyarla False
    KS_Bitir True
    KS_Log "TOPLU AKTAR", tur, m & " aktarýldý, " & hataN & " hatalý, " & atla & " daha önce aktarýlmýþ"
    MsgBox "Toplu aktarým tamamlandý (" & tur & ")." & vbCrLf & vbCrLf & _
        "Aktarýlan: " & m & vbCrLf & _
        "Hatalý (aktarýlmadý): " & hataN & vbCrLf & _
        "Daha önce aktarýlmýþ (atlandý): " & atla & vbCrLf & vbCrLf & _
        IIf(hataN > 0, "Hatalý satýrlarý SONUÇ sütununa göre düzeltip tekrar çalýþtýrýn.", ""), _
        IIf(hataN > 0, vbExclamation, vbInformation), KS_AD
    Exit Sub
Hata:
    KS_OtoTanimAyarla False
    KS_Bitir True
    MsgBox "Toplu aktarým durdu. Tablolara hiçbir satýr yazýlmadýysa veri deðiþmedi." & vbCrLf & _
        "Adým: " & adim & vbCrLf & "Hata: " & Err.Description, vbCritical, KS_AD
End Sub
