Attribute VB_Name = "modKS_KAYIT"
Option Explicit

' KAYIT sayfasý: üç kartlý giriþ formu (Malzeme / Sevkiyat / Arýza) ve ortak doðrulama motoru.
' Toplu aktarým da ayný Hazirla fonksiyonlarýný kullanýr; kurallar tek yerde durur.

Private dMalz As Object, aMalz As Variant
Private dSevk As Object, aSevk As Variant
Private dSevkci As Object, dYer As Object, dIrs As Object
Private mSevkSayac As Long, mArizaSayac As Long
Private mOtoTanim As Boolean

Private Const KART_UST As Long = 7
Private Const ALAN_ILK As Long = 9
Private Const BUTON_SATIR As Long = 28

' ================================================================ baðlam

Public Sub KS_BaglamYukle()
    Dim i As Long, k As String, loS As ListObject
    KS_ParamYukle
    Set dMalz = CreateObject("Scripting.Dictionary")
    aMalz = KS_Veri(KS_LOZ(T_MALZ))
    If IsArray(aMalz) Then
        For i = 1 To UBound(aMalz, 1)
            dMalz(KS_Anahtar(aMalz(i, M_KOD))) = i
        Next i
    End If
    Set dSevk = CreateObject("Scripting.Dictionary")
    Set dIrs = CreateObject("Scripting.Dictionary")
    Set loS = KS_LOZ(T_SEVK)
    aSevk = KS_Veri(loS)
    If IsArray(aSevk) Then
        For i = 1 To UBound(aSevk, 1)
            dSevk(KS_Anahtar(aSevk(i, S_NO))) = i
            k = KS_Anahtar(aSevk(i, S_IRS))
            If Len(k) > 0 Then dIrs(k & "|" & KS_Anahtar(aSevk(i, S_KOD))) = KS_Metin(aSevk(i, S_NO))
        Next i
    End If
    Set dSevkci = KS_Sozluk(T_SEVKCI)
    Set dYer = KS_Sozluk(T_YER)
    mSevkSayac = KS_EnBuyukNo(loS, "SVK-")
    mArizaSayac = KS_EnBuyukNo(KS_LOZ(T_ARIZA), "ARZ-")
    mOtoTanim = False
End Sub

' Toplu aktarýmda reddedilen yeni sevkiyatýn numarasý bir sonrakine býrakýlýr.
Public Sub KS_SevkSayacGeri()
    mSevkSayac = mSevkSayac - 1
End Sub

Public Sub KS_OtoTanimAyarla(ByVal acik As Boolean)
    mOtoTanim = acik
End Sub

Public Function KS_MalzemeVar(ByVal kod As Variant) As Boolean
    KS_MalzemeVar = dMalz.Exists(KS_Anahtar(kod))
End Function

Public Sub KS_MalzemeKaydiEkle(ByVal kod As Variant)
    dMalz(KS_Anahtar(kod)) = 0
End Sub

' Listede olmayan tanýmý toplu aktarýmda isteðe baðlý ekler.
Private Function KS_TanimKontrol(ByVal d As Object, ByVal tablo As String, ByVal deger As String, _
        ByVal alan As String, ByRef hata As String) As Boolean
    Dim lo As ListObject, v As Variant, korumali As Boolean
    If d.Exists(KS_Anahtar(deger)) Then
        KS_TanimKontrol = True
        Exit Function
    End If
    If Not mOtoTanim Then
        hata = alan & " TANIMLAR listesinde yok: " & deger
        Exit Function
    End If
    Set lo = KS_LOZ(tablo)
    v = KS_BosSatir(lo.ListColumns.Count)
    v(1, 1) = deger
    If tablo = T_SEVKCI And Len(deger) > 11 Then
        If Mid$(deger, 11, 1) = " " And IsNumeric(Left$(deger, 10)) Then
            v(1, 2) = Left$(deger, 10)
            v(1, 3) = Mid$(deger, 12)
        End If
    End If
    v(1, lo.ListColumns.Count) = "Toplu aktarýmda eklendi"
    korumali = lo.Parent.ProtectContents
    KS_Ac lo.Parent
    KS_SatirYaz lo, 0, v
    If korumali Then KS_Kapat lo.Parent
    d(KS_Anahtar(deger)) = lo.ListRows.Count
    KS_TanimKontrol = True
End Function

' ================================================================ doðrulama motoru

' g: Kod, Ad, Kategori, Birim, Üretici, Birim Fiyat, Garanti (Ay), Not
Public Function KS_MalzemeHazirla(ByVal g As Variant, ByVal idx As Long, ByRef satir As Variant, _
        ByRef hata As String) As Boolean
    Dim kod As String, fiyat As Double, gar As Double
    hata = ""
    kod = KS_Metin(g(0))
    If Len(kod) = 0 Then hata = "Malzeme kodu boþ olamaz.": Exit Function
    If Len(kod) > 40 Then hata = "Malzeme kodu 40 karakteri geçemez.": Exit Function
    If idx = 0 And dMalz.Exists(KS_Anahtar(kod)) Then hata = "Bu malzeme kodu zaten kayýtlý: " & kod: Exit Function
    If Len(KS_Metin(g(1))) = 0 Then hata = "Malzeme adý boþ olamaz.": Exit Function
    If Not KS_Sayi(g(5), fiyat) Then hata = "Birim fiyat sayý olmalý.": Exit Function
    If fiyat < 0 Then hata = "Birim fiyat negatif olamaz.": Exit Function
    If Not KS_Sayi(g(6), gar) Then hata = "Garanti süresi (ay) girilmeli.": Exit Function
    If gar < 0 Or gar > 240 Or gar <> Int(gar) Then hata = "Garanti süresi 0-240 arasý tam ay olmalý.": Exit Function
    If idx = 0 Then
        satir = KS_BosSatir(M_KOLON)
        satir(1, M_KAYIT) = Date
    Else
        satir = KS_SatirKopya(aMalz, idx)
    End If
    satir(1, M_KOD) = kod
    satir(1, M_AD) = KS_Metin(g(1))
    satir(1, M_KAT) = KS_Metin(g(2))
    satir(1, M_BIRIM) = KS_Metin(g(3))
    satir(1, M_URETICI) = KS_Metin(g(4))
    satir(1, M_FIYAT) = fiyat
    satir(1, M_GARANTI) = CLng(gar)
    satir(1, M_NOT) = KS_Metin(g(7))
    KS_MalzemeHazirla = True
End Function

' g: Sevkçi, Sevk Yeri, Malzeme Kodu, Adet, Birim Fiyat, Planlanan, Gerçek Geliþ, Ýrsaliye, Seri, Not
Public Function KS_SevkHazirla(ByVal g As Variant, ByVal idx As Long, ByRef satir As Variant, _
        ByRef hata As String, ByRef uyari As String) As Boolean
    Dim mi As Long, adet As Double, fiyat As Double, plan As Date, gelis As Date, k As String, no As String
    hata = "": uyari = ""
    If Len(KS_Metin(g(0))) = 0 Then hata = "Sevkçi seçilmeli.": Exit Function
    If Not KS_TanimKontrol(dSevkci, T_SEVKCI, KS_Metin(g(0)), "Sevkçi", hata) Then Exit Function
    If Len(KS_Metin(g(1))) = 0 Then hata = "Sevk yeri seçilmeli.": Exit Function
    If Not KS_TanimKontrol(dYer, T_YER, KS_Metin(g(1)), "Sevk yeri", hata) Then Exit Function
    If Len(KS_Metin(g(2))) = 0 Then hata = "Malzeme kodu girilmeli.": Exit Function
    If Not dMalz.Exists(KS_Anahtar(g(2))) Then hata = "Malzeme kodu kayýtlý deðil: " & KS_Metin(g(2)) & _
        ". Önce malzemeyi tanýmlayýn.": Exit Function
    mi = dMalz(KS_Anahtar(g(2)))
    If Not KS_Sayi(g(3), adet) Then hata = "Adet / miktar sayý olmalý.": Exit Function
    If adet <= 0 Then hata = "Adet / miktar sýfýrdan büyük olmalý.": Exit Function
    If KS_Bos(g(4)) Then
        If Not KS_Sayi(aMalz(mi, M_FIYAT), fiyat) Then hata = "Malzemenin liste fiyatý yok; birim fiyat girin.": Exit Function
    Else
        If Not KS_Sayi(g(4), fiyat) Then hata = "Birim fiyat sayý olmalý.": Exit Function
        If fiyat < 0 Then hata = "Birim fiyat negatif olamaz.": Exit Function
    End If
    If Not KS_Tarih(g(5), plan) Then hata = "Planlanan teslim tarihi geçerli deðil (gg.aa.yyyy).": Exit Function
    If Not KS_Bos(g(6)) Then
        If Not KS_Tarih(g(6), gelis) Then hata = "Gerçek geliþ tarihi geçerli deðil (gg.aa.yyyy).": Exit Function
        If gelis > Date Then hata = "Gerçek geliþ tarihi bugünden ileri olamaz.": Exit Function
    End If

    If idx = 0 Then
        satir = KS_BosSatir(S_KOLON)
        mSevkSayac = mSevkSayac + 1
        no = KS_NoBicim("SVK-", mSevkSayac)
        satir(1, S_NO) = no
        satir(1, S_ZAMAN) = Now
        satir(1, S_KULL) = KS_Kullanici()
        satir(1, S_GARAY) = aMalz(mi, M_GARANTI)
    Else
        satir = KS_SatirKopya(aSevk, idx)
        no = KS_Metin(satir(1, S_NO))
        ' Garanti süresi sevk anýndaki sözleþmeden gelir; yalnýz malzeme deðiþirse yenilenir.
        If KS_Anahtar(satir(1, S_KOD)) <> KS_Anahtar(g(2)) Then satir(1, S_GARAY) = aMalz(mi, M_GARANTI)
    End If
    satir(1, S_SEVKCI) = KS_Metin(g(0))
    satir(1, S_YER) = KS_Metin(g(1))
    satir(1, S_KOD) = KS_Metin(aMalz(mi, M_KOD))
    satir(1, S_AD) = aMalz(mi, M_AD)
    satir(1, S_ADET) = adet
    satir(1, S_FIYAT) = fiyat
    satir(1, S_PLAN) = plan
    If KS_Bos(g(6)) Then satir(1, S_GELIS) = Empty Else satir(1, S_GELIS) = gelis
    satir(1, S_IRS) = KS_Metin(g(7))
    satir(1, S_SERI) = KS_Metin(g(8))
    satir(1, S_NOT) = KS_Metin(g(9))
    KS_SevkHesapla satir, 1

    k = KS_Anahtar(g(7))
    If Len(k) > 0 Then
        k = k & "|" & KS_Anahtar(g(2))
        If dIrs.Exists(k) Then
            If dIrs(k) <> no Then uyari = "Ayný irsaliye ve malzeme ile kayýt var: " & dIrs(k)
        Else
            dIrs(k) = no
        End If
    End If
    KS_SevkHazirla = True
End Function

' g: Sevk No, Seri No, Arýza Tarihi, Adet, Tip, Açýklama, Bildirim, Dönüþ, Çözüm, Çözüm Þekli,
'    Tamir Maliyeti, Karþýlayan, Not
Public Function KS_ArizaHazirla(ByVal g As Variant, ByVal idx As Long, ByVal mevcut As Variant, _
        ByRef satir As Variant, ByRef hata As String) As Boolean
    Dim si As Long, mi As Long, t As Date, gelis As Date, bil As Date, don As Date, coz As Date
    Dim adet As Double, sevkAdet As Double, maliyet As Double, hb As Boolean
    hata = ""
    If Len(KS_Metin(g(0))) = 0 Then hata = "Sevk no seçilmeli.": Exit Function
    If Not dSevk.Exists(KS_Anahtar(g(0))) Then hata = "Sevk no bulunamadý: " & KS_Metin(g(0)): Exit Function
    si = dSevk(KS_Anahtar(g(0)))
    If Not KS_Tarih(g(2), t) Then hata = "Arýza tarihi geçerli deðil (gg.aa.yyyy).": Exit Function
    If t > Date Then hata = "Arýza tarihi bugünden ileri olamaz.": Exit Function
    If KS_Tarih(aSevk(si, S_GELIS), gelis) Then
        If t < gelis Then hata = "Arýza tarihi, malzemenin geliþ tarihinden (" & _
            Format$(gelis, "dd.mm.yyyy") & ") önce olamaz.": Exit Function
    End If
    If Not KS_Sayi(g(3), adet) Then hata = "Arýzalý adet sayý olmalý.": Exit Function
    If adet <= 0 Then hata = "Arýzalý adet sýfýrdan büyük olmalý.": Exit Function
    If KS_Sayi(aSevk(si, S_ADET), sevkAdet) Then
        If adet > sevkAdet Then hata = "Arýzalý adet (" & adet & "), sevk edilen miktarý (" & sevkAdet & ") aþamaz.": Exit Function
    End If
    If Not KS_Bos(g(6)) Then
        If Not KS_Tarih(g(6), bil) Then hata = "Üreticiye bildirim tarihi geçerli deðil.": Exit Function
        If bil < t Then hata = "Bildirim tarihi arýza tarihinden önce olamaz.": Exit Function
        If bil > Date Then hata = "Bildirim tarihi bugünden ileri olamaz.": Exit Function
        hb = True
    End If
    If Not KS_Bos(g(7)) Then
        If Not hb Then hata = "Üretici dönüþ tarihi için önce bildirim tarihi girilmeli.": Exit Function
        If Not KS_Tarih(g(7), don) Then hata = "Üretici dönüþ tarihi geçerli deðil.": Exit Function
        If don < bil Then hata = "Dönüþ tarihi bildirim tarihinden önce olamaz.": Exit Function
        If don > Date Then hata = "Dönüþ tarihi bugünden ileri olamaz.": Exit Function
    End If
    If Not KS_Bos(g(8)) Then
        If Not KS_Tarih(g(8), coz) Then hata = "Çözüm tarihi geçerli deðil.": Exit Function
        If coz < t Then hata = "Çözüm tarihi arýza tarihinden önce olamaz.": Exit Function
        If coz > Date Then hata = "Çözüm tarihi bugünden ileri olamaz.": Exit Function
    End If
    If Not KS_Bos(g(10)) Then
        If Not KS_Sayi(g(10), maliyet) Then hata = "Tamir maliyeti sayý olmalý.": Exit Function
        If maliyet < 0 Then hata = "Tamir maliyeti negatif olamaz.": Exit Function
    End If

    If idx = 0 Then
        satir = KS_BosSatir(A_KOLON)
        mArizaSayac = mArizaSayac + 1
        satir(1, A_NO) = KS_NoBicim("ARZ-", mArizaSayac)
        satir(1, A_ZAMAN) = Now
        satir(1, A_KULL) = KS_Kullanici()
    Else
        satir = mevcut
    End If
    satir(1, A_SEVK) = KS_Metin(aSevk(si, S_NO))
    satir(1, A_SERI) = KS_Metin(g(1))
    satir(1, A_KOD) = aSevk(si, S_KOD)
    satir(1, A_AD) = aSevk(si, S_AD)
    satir(1, A_SEVKCI) = aSevk(si, S_SEVKCI)
    satir(1, A_YER) = aSevk(si, S_YER)
    satir(1, A_GARBIT) = aSevk(si, S_GARBIT)
    If dMalz.Exists(KS_Anahtar(aSevk(si, S_KOD))) Then
        mi = dMalz(KS_Anahtar(aSevk(si, S_KOD)))
        If mi > 0 Then satir(1, A_URETICI) = aMalz(mi, M_URETICI)
    End If
    satir(1, A_TARIH) = t
    satir(1, A_ADET) = adet
    satir(1, A_TIP) = KS_Metin(g(4))
    satir(1, A_ACIK) = KS_Metin(g(5))
    If KS_Bos(g(6)) Then satir(1, A_BILDIRIM) = Empty Else satir(1, A_BILDIRIM) = bil
    If KS_Bos(g(7)) Then satir(1, A_DONUS) = Empty Else satir(1, A_DONUS) = don
    If KS_Bos(g(8)) Then satir(1, A_COZUM) = Empty Else satir(1, A_COZUM) = coz
    satir(1, A_SEKIL) = KS_Metin(g(9))
    If KS_Bos(g(10)) Then satir(1, A_MALIYET) = Empty Else satir(1, A_MALIYET) = maliyet
    satir(1, A_KARS) = KS_Metin(g(11))
    satir(1, A_NOT) = KS_Metin(g(12))
    KS_ArizaHesapla satir, 1
    KS_ArizaHazirla = True
End Function

' ================================================================ form düzeni

Private Function KS_MalzemeAlanlari() As Variant
    KS_MalzemeAlanlari = Array( _
        Array("Malzeme Kodu", "fM_KOD", "kod", True, ""), _
        Array("Malzeme Adý", "fM_AD", "metin", True, ""), _
        Array("Kategori", "fM_KAT", "listeS", False, "lstKategori"), _
        Array("Birim", "fM_BIRIM", "listeS", False, "lstBirim"), _
        Array("Üretici", "fM_URETICI", "listeS", False, "lstSevkci"), _
        Array("Birim Fiyat (TL)", "fM_FIYAT", "para", True, ""), _
        Array("Garanti Süresi (Ay)", "fM_GARANTI", "sayi", True, ""), _
        Array("Not", "fM_NOT", "metin", False, ""))
End Function

Private Function KS_SevkAlanlari() As Variant
    Dim m As String
    m = "MATCH(fS_KOD,INDEX(" & T_MALZ & ",0,1),0)"
    KS_SevkAlanlari = Array( _
        Array("Sevk No  (güncellemede)", "fS_NO", "kod", False, ""), _
        Array("Sevkçi", "fS_SEVKCI", "liste", True, "lstSevkci"), _
        Array("Sevk Yeri", "fS_YER", "liste", True, "lstYer"), _
        Array("Malzeme Kodu", "fS_KOD", "listeK", True, "lstMalzeme"), _
        Array("» Malzeme", "fS_OTO", "oto", False, _
            "=IFERROR(INDEX(" & T_MALZ & "," & m & "," & M_AD & ")&""   |   Garanti ""&INDEX(" & T_MALZ & _
            "," & m & "," & M_GARANTI & ")&"" ay"","""")"), _
        Array("» Liste Fiyatý", "fS_LFIYAT", "otoPara", False, _
            "=IFERROR(INDEX(" & T_MALZ & "," & m & "," & M_FIYAT & "),"""")"), _
        Array("Adet / Miktar", "fS_ADET", "sayi", True, ""), _
        Array("Birim Fiyat (boþ = liste)", "fS_FIYAT", "para", False, ""), _
        Array("Planlanan Teslim Tarihi", "fS_PLAN", "tarih", True, ""), _
        Array("Gerçek Geliþ Tarihi", "fS_GELIS", "tarih", False, ""), _
        Array("Ýrsaliye / Belge No", "fS_IRS", "kod", False, ""), _
        Array("Seri / Lot No", "fS_SERI", "kod", False, ""), _
        Array("Not", "fS_NOT", "metin", False, ""))
End Function

Private Function KS_ArizaAlanlari() As Variant
    Dim m As String
    m = "MATCH(fA_SEVK,INDEX(" & T_SEVK & ",0,1),0)"
    KS_ArizaAlanlari = Array( _
        Array("Arýza No  (güncellemede)", "fA_NO", "kod", False, ""), _
        Array("Sevk No", "fA_SEVK", "listeK", True, "lstSevk"), _
        Array("» Malzeme", "fA_OTO1", "oto", False, _
            "=IFERROR(INDEX(" & T_SEVK & "," & m & "," & S_KOD & ")&""  -  ""&INDEX(" & T_SEVK & "," & m & "," & S_AD & "),"""")"), _
        Array("» Sevk Yeri / Sevkçi", "fA_OTO2", "oto", False, _
            "=IFERROR(INDEX(" & T_SEVK & "," & m & "," & S_YER & ")&""  /  ""&INDEX(" & T_SEVK & "," & m & "," & S_SEVKCI & "),"""")"), _
        Array("» Garanti Bitiþ", "fA_GARBIT", "otoTarih", False, _
            "=IFERROR(IF(INDEX(" & T_SEVK & "," & m & "," & S_GARBIT & ")="""",""Geliþ tarihi yok"",INDEX(" & T_SEVK & "," & m & "," & S_GARBIT & ")),"""")"), _
        Array("Seri No (varsa)", "fA_SERI", "kod", False, ""), _
        Array("Arýza Tarihi", "fA_TARIH", "tarih", True, ""), _
        Array("» Garanti Durumu", "fA_KAPSAM", "oto", False, _
            "=IF(OR(fA_TARIH="""",NOT(ISNUMBER(fA_GARBIT))),"""",IF(fA_TARIH<=fA_GARBIT,""GARANTÝ KAPSAMINDA"",""GARANTÝ DIÞI""))"), _
        Array("Arýzalý Adet", "fA_ADET", "sayi", True, ""), _
        Array("Arýza Tipi", "fA_TIP", "listeS", False, "lstTip"), _
        Array("Arýza Açýklamasý", "fA_ACIK", "metin", False, ""), _
        Array("Üreticiye Bildirim Tarihi", "fA_BILDIRIM", "tarih", False, ""), _
        Array("Üretici Dönüþ Tarihi", "fA_DONUS", "tarih", False, ""), _
        Array("Çözüm Tarihi", "fA_COZUM", "tarih", False, ""), _
        Array("Çözüm Þekli", "fA_SEKIL", "listeS", False, "lstCozum"), _
        Array("Tamir Maliyeti (TL)", "fA_MALIYET", "para", False, ""), _
        Array("Maliyeti Karþýlayan", "fA_KARS", "listeS", False, "lstKarsilayan"), _
        Array("Not", "fA_NOT", "metin", False, ""))
End Function

Public Sub KS_KayitFormKur()
    Dim ws As Worksheet
    Set ws = KS_WS(SH_KAYIT)
    KS_SayfaSifirla ws
    ws.Columns("B").ColumnWidth = 23
    ws.Columns("C").ColumnWidth = 30
    ws.Columns("D").ColumnWidth = 3
    ws.Columns("E").ColumnWidth = 23
    ws.Columns("F").ColumnWidth = 34
    ws.Columns("G").ColumnWidth = 3
    ws.Columns("H").ColumnWidth = 24
    ws.Columns("I").ColumnWidth = 34
    ws.Columns("J").ColumnWidth = 3
    KS_Bant ws, 10, "KAYIT MERKEZÝ", _
        "Sýrayla: 1 Malzeme > 2 Sevkiyat > 3 Arýza.  * zorunlu   » otomatik   GETÝR: numarayla kaydý forma çaðýrýr."
    KS_FormKart ws, 2, "01   MALZEME TANIMI", "Ana kayýt: kod, fiyat, garanti süresi", KS_MalzemeAlanlari(), _
        Array("KAYDET", "KS_MALZEME_KAYDET", "GÜNCELLE", "KS_MALZEME_GUNCELLE", "GETÝR", "KS_MALZEME_GETIR", _
            "TEMÝZLE", "KS_MALZEME_TEMIZLE"), "kM"
    KS_FormKart ws, 5, "02   SEVKÝYAT KAYDI", "Hangi sevkçi, nereye, ne zaman, kaç adet", KS_SevkAlanlari(), _
        Array("KAYDET", "KS_SEVK_KAYDET", "GÜNCELLE", "KS_SEVK_GUNCELLE", "GETÝR", "KS_SEVK_GETIR", _
            "TEMÝZLE", "KS_SEVK_TEMIZLE"), "kS"
    KS_FormKart ws, 8, "03   ARIZA KAYDI", "Garanti kapsamý ve üretici dönüþü otomatik izlenir", KS_ArizaAlanlari(), _
        Array("KAYDET", "KS_ARIZA_KAYDET", "GÜNCELLE", "KS_ARIZA_GUNCELLE", "GETÝR", "KS_ARIZA_GETIR", _
            "TEMÝZLE", "KS_ARIZA_TEMIZLE"), "kA"
    KS_KapsamRenk ws
    KS_Pencere ws, 90, 0, False
    ws.Range("C" & ALAN_ILK).Select
End Sub

Private Sub KS_FormKart(ByVal ws As Worksheet, ByVal kol As Long, ByVal baslik As String, _
        ByVal alt As String, ByVal alanlar As Variant, ByVal butonlar As Variant, ByVal onek As String)
    Dim i As Long, r As Long, e As Range, v As Range, a As Variant, kart As Range
    Dim sol As Double, gen As Double, ust As Double, b As Long

    Set kart = ws.Range(ws.Cells(KART_UST, kol), ws.Cells(BUTON_SATIR + 2, kol + 1))
    kart.Interior.Color = C_KART()
    With kart
        .Borders(xlEdgeLeft).Color = C_CIZGI()
        .Borders(xlEdgeRight).Color = C_CIZGI()
        .Borders(xlEdgeBottom).Color = C_CIZGI()
    End With
    With ws.Range(ws.Cells(KART_UST, kol), ws.Cells(KART_UST, kol + 1))
        .Interior.Color = C_LAC()
        .Borders(xlEdgeTop).Color = C_CAMGOBEGI()
        .Borders(xlEdgeTop).Weight = xlThick
    End With
    ws.Rows(KART_UST).RowHeight = 28
    With ws.Cells(KART_UST, kol)
        .Value = baslik
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 11
        .IndentLevel = 1
    End With
    With ws.Cells(KART_UST + 1, kol)
        .Value = alt
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = C_GRI()
        .IndentLevel = 1
    End With
    ws.Rows(KART_UST + 1).RowHeight = 20

    For i = 0 To UBound(alanlar)
        a = alanlar(i)
        r = ALAN_ILK + i
        ws.Rows(r).RowHeight = 21
        Set e = ws.Cells(r, kol)
        Set v = ws.Cells(r, kol + 1)
        e.Value = a(0) & IIf(a(3), "  *", "")
        e.Font.Size = 9
        e.IndentLevel = 1
        If Left$(a(2), 3) = "oto" Then
            e.Font.Color = C_TURKUAZ()
            e.Font.Italic = True
        Else
            e.Font.Color = C_METIN()
        End If
        If a(3) Then e.Characters(Len(e.Value) - 0, 1).Font.Color = C_CAMGOBEGI()
        KS_AdHucre CStr(a(1)), v
        With v
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
            .IndentLevel = 1
            .Borders(xlEdgeBottom).Color = C_CIZGI()
        End With
        Select Case a(2)
            Case "oto", "otoPara", "otoTarih"
                v.Interior.Color = C_OTO()
                v.Font.Color = C_TURKUAZ()
                v.Font.Bold = True
                v.Locked = True
                If a(2) = "otoPara" Then v.NumberFormat = KS_FmtTL()
                If a(2) = "otoTarih" Then v.NumberFormat = KS_FmtTarih()
                v.Formula = a(4)
            Case Else
                v.Interior.Color = C_GIRIS()
                v.Font.Color = C_LAC()
                v.Locked = False
                v.Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
                Select Case a(2)
                    Case "kod", "metin": v.NumberFormat = "@"
                    Case "liste": v.NumberFormat = "@": KS_Liste v, CStr(a(4)), True
                    Case "listeK": v.NumberFormat = "@": KS_Liste v, CStr(a(4)), True
                    Case "listeS": v.NumberFormat = "@": KS_Liste v, CStr(a(4)), False
                    Case "para": v.NumberFormat = KS_FmtTL(): KS_SayiDogrulama v
                    Case "sayi": v.NumberFormat = "General": KS_SayiDogrulama v
                    Case "tarih": v.NumberFormat = KS_FmtTarih(): KS_TarihDogrulama v
                End Select
        End Select
    Next i

    ' Butonlar: kartýn altýnda eþit aralýklý dört düðme.
    sol = ws.Cells(BUTON_SATIR, kol).Left + 6
    gen = (ws.Cells(BUTON_SATIR, kol + 2).Left - ws.Cells(BUTON_SATIR, kol).Left - 12 - 3 * 5) / 4
    ust = ws.Cells(BUTON_SATIR, kol).Top + 2
    ws.Rows(BUTON_SATIR).RowHeight = 28
    For b = 0 To 3
        KS_Buton ws, onek & "_btn" & b, CStr(butonlar(b * 2)), sol + b * (gen + 5), ust, gen, 24, _
            CStr(butonlar(b * 2 + 1)), Choose(b + 1, "ana", "koyu", "cizgi", "gri")
    Next b
End Sub

Private Sub KS_KapsamRenk(ByVal ws As Worksheet)
    Dim r As Range, fc As Object
    Set r = KS_F("fA_KAPSAM")
    r.FormatConditions.Delete
    Set fc = r.FormatConditions.Add(Type:=xlTextString, String:="GARANTÝ KAPSAMINDA", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_YESIL()
    fc.Interior.Color = RGB(220, 245, 232)
    Set fc = r.FormatConditions.Add(Type:=xlTextString, String:="GARANTÝ DIÞI", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_KIRMIZI()
    fc.Interior.Color = RGB(253, 226, 228)
End Sub

' ================================================================ form yardýmcýlarý

Private Sub KS_Temizle(ByVal adlar As Variant)
    Dim a As Variant
    For Each a In adlar
        KS_F(CStr(a)).ClearContents
    Next a
End Sub

Private Function KS_Degerler(ByVal adlar As Variant) As Variant
    Dim v() As Variant, i As Long
    ReDim v(0 To UBound(adlar))
    For i = 0 To UBound(adlar)
        v(i) = KS_FV(CStr(adlar(i)))
    Next i
    KS_Degerler = v
End Function

Private Sub KS_Doldur(ByVal adlar As Variant, ByVal degerler As Variant)
    Dim i As Long
    For i = 0 To UBound(adlar)
        KS_F(CStr(adlar(i))).Value = degerler(i)
    Next i
End Sub

Private Function KS_AlanM() As Variant
    KS_AlanM = Array("fM_KOD", "fM_AD", "fM_KAT", "fM_BIRIM", "fM_URETICI", "fM_FIYAT", "fM_GARANTI", "fM_NOT")
End Function

Private Function KS_AlanS() As Variant
    KS_AlanS = Array("fS_SEVKCI", "fS_YER", "fS_KOD", "fS_ADET", "fS_FIYAT", "fS_PLAN", "fS_GELIS", _
        "fS_IRS", "fS_SERI", "fS_NOT")
End Function

Private Function KS_AlanA() As Variant
    KS_AlanA = Array("fA_SEVK", "fA_SERI", "fA_TARIH", "fA_ADET", "fA_TIP", "fA_ACIK", "fA_BILDIRIM", _
        "fA_DONUS", "fA_COZUM", "fA_SEKIL", "fA_MALIYET", "fA_KARS", "fA_NOT")
End Function

' Ortak kayýt yazma: veriyi yazar, tümünü yeniden hesaplar, günlüðe iþler.
Private Sub KS_Yaz(ByVal tablo As String, ByVal idx As Long, ByVal satir As Variant, _
        ByVal islem As String, ByVal no As String, ByVal detay As String)
    Dim lo As ListObject
    Set lo = KS_LOZ(tablo)
    KS_Basla
    KS_Ac lo.Parent
    KS_SatirYaz lo, idx, satir
    KS_Kapat lo.Parent
    KS_HesaplaTumu
    KS_Bitir
    KS_Log islem, no, detay
End Sub

Private Sub KS_HataGoster(ByVal baslik As String)
    Dim no As Long, ac As String
    no = Err.Number: ac = Err.Description
    KS_Bitir True
    MsgBox baslik & vbCrLf & "Hata: " & no & " / " & ac, vbExclamation, KS_AD
End Sub

' ================================================================ MALZEME

Public Sub KS_MALZEME_KAYDET()
    KS_MalzemeIslem False
End Sub
Public Sub KS_MALZEME_GUNCELLE()
    KS_MalzemeIslem True
End Sub

Private Sub KS_MalzemeIslem(ByVal guncelle As Boolean)
    Dim g As Variant, satir As Variant, hata As String, idx As Long
    On Error GoTo Hata
    KS_BaglamYukle
    g = KS_Degerler(KS_AlanM())
    If guncelle Then
        If dMalz.Exists(KS_Anahtar(g(0))) Then idx = dMalz(KS_Anahtar(g(0)))
        If idx = 0 Then
            MsgBox "Güncellenecek malzeme bulunamadý. Kodu yazýp önce GETÝR'e basýn.", vbExclamation, KS_AD
            Exit Sub
        End If
    End If
    If Not KS_MalzemeHazirla(g, idx, satir, hata) Then
        MsgBox hata, vbExclamation, KS_AD
        Exit Sub
    End If
    If guncelle Then
        If MsgBox(KS_Metin(g(0)) & " güncellensin mi?" & vbCrLf & vbCrLf & _
            "Not: Garanti süresi deðiþikliði yalnýz bundan sonraki sevkiyatlara uygulanýr.", _
            vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    KS_Yaz T_MALZ, idx, satir, IIf(guncelle, "MALZEME GÜNCELLE", "MALZEME EKLE"), _
        KS_Metin(satir(1, M_KOD)), KS_Metin(satir(1, M_AD))
    KS_Temizle KS_AlanM()
    MsgBox "Malzeme " & IIf(guncelle, "güncellendi: ", "kaydedildi: ") & KS_Metin(satir(1, M_KOD)), vbInformation, KS_AD
    Exit Sub
Hata:
    KS_HataGoster "Malzeme kaydedilemedi."
End Sub

Public Sub KS_MALZEME_GETIR()
    Dim idx As Long, k As Variant
    On Error GoTo Hata
    KS_BaglamYukle
    k = KS_FV("fM_KOD")
    If dMalz.Exists(KS_Anahtar(k)) Then idx = dMalz(KS_Anahtar(k))
    If idx = 0 Then
        MsgBox "Malzeme kodu bulunamadý: " & KS_Metin(k), vbExclamation, KS_AD
        Exit Sub
    End If
    KS_Doldur KS_AlanM(), Array(aMalz(idx, M_KOD), aMalz(idx, M_AD), aMalz(idx, M_KAT), aMalz(idx, M_BIRIM), _
        aMalz(idx, M_URETICI), aMalz(idx, M_FIYAT), aMalz(idx, M_GARANTI), aMalz(idx, M_NOT))
    Exit Sub
Hata:
    KS_HataGoster "Malzeme getirilemedi."
End Sub

Public Sub KS_MALZEME_TEMIZLE()
    KS_Temizle KS_AlanM()
End Sub

' ================================================================ SEVKÝYAT

Public Sub KS_SEVK_KAYDET()
    KS_SevkIslem False
End Sub
Public Sub KS_SEVK_GUNCELLE()
    KS_SevkIslem True
End Sub

Private Sub KS_SevkIslem(ByVal guncelle As Boolean)
    Dim g As Variant, satir As Variant, hata As String, uyari As String, idx As Long, no As Variant
    On Error GoTo Hata
    KS_BaglamYukle
    g = KS_Degerler(KS_AlanS())
    If guncelle Then
        no = KS_FV("fS_NO")
        If dSevk.Exists(KS_Anahtar(no)) Then idx = dSevk(KS_Anahtar(no))
        If idx = 0 Then
            MsgBox "Güncellenecek sevk no bulunamadý. Sevk no yazýp önce GETÝR'e basýn.", vbExclamation, KS_AD
            Exit Sub
        End If
    End If
    If Not KS_SevkHazirla(g, idx, satir, hata, uyari) Then
        MsgBox hata, vbExclamation, KS_AD
        Exit Sub
    End If
    If Len(uyari) > 0 Then
        If MsgBox(uyari & vbCrLf & vbCrLf & "Yine de kaydedilsin mi?", vbExclamation + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    If guncelle Then
        If MsgBox(KS_Metin(satir(1, S_NO)) & " güncellensin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    KS_Yaz T_SEVK, idx, satir, IIf(guncelle, "SEVK GÜNCELLE", "SEVK EKLE"), KS_Metin(satir(1, S_NO)), _
        KS_Metin(satir(1, S_KOD)) & " x " & satir(1, S_ADET) & " > " & KS_Metin(satir(1, S_YER))
    KS_Temizle KS_AlanS()
    KS_F("fS_NO").ClearContents
    MsgBox "Sevkiyat " & IIf(guncelle, "güncellendi: ", "kaydedildi: ") & KS_Metin(satir(1, S_NO)) & vbCrLf & _
        "Garanti: " & KS_Metin(satir(1, S_GDURUM)), vbInformation, KS_AD
    Exit Sub
Hata:
    KS_HataGoster "Sevkiyat kaydedilemedi."
End Sub

Public Sub KS_SEVK_GETIR()
    Dim idx As Long, k As Variant
    On Error GoTo Hata
    KS_BaglamYukle
    k = KS_FV("fS_NO")
    If dSevk.Exists(KS_Anahtar(k)) Then idx = dSevk(KS_Anahtar(k))
    If idx = 0 Then
        MsgBox "Sevk no bulunamadý: " & KS_Metin(k), vbExclamation, KS_AD
        Exit Sub
    End If
    KS_Doldur KS_AlanS(), Array(aSevk(idx, S_SEVKCI), aSevk(idx, S_YER), aSevk(idx, S_KOD), aSevk(idx, S_ADET), _
        aSevk(idx, S_FIYAT), aSevk(idx, S_PLAN), aSevk(idx, S_GELIS), aSevk(idx, S_IRS), aSevk(idx, S_SERI), _
        aSevk(idx, S_NOT))
    Exit Sub
Hata:
    KS_HataGoster "Sevkiyat getirilemedi."
End Sub

Public Sub KS_SEVK_TEMIZLE()
    KS_Temizle KS_AlanS()
    KS_F("fS_NO").ClearContents
End Sub

' ================================================================ ARIZA

Public Sub KS_ARIZA_KAYDET()
    KS_ArizaIslem False
End Sub
Public Sub KS_ARIZA_GUNCELLE()
    KS_ArizaIslem True
End Sub

Private Sub KS_ArizaIslem(ByVal guncelle As Boolean)
    Dim g As Variant, satir As Variant, hata As String, idx As Long, lo As ListObject, mevcut As Variant
    On Error GoTo Hata
    KS_BaglamYukle
    g = KS_Degerler(KS_AlanA())
    Set lo = KS_LOZ(T_ARIZA)
    If guncelle Then
        idx = KS_SatirBul(lo, A_NO, KS_FV("fA_NO"))
        If idx = 0 Then
            MsgBox "Güncellenecek arýza no bulunamadý. Arýza no yazýp önce GETÝR'e basýn.", vbExclamation, KS_AD
            Exit Sub
        End If
        mevcut = lo.ListRows(idx).Range.Value
    End If
    If Not KS_ArizaHazirla(g, idx, mevcut, satir, hata) Then
        MsgBox hata, vbExclamation, KS_AD
        Exit Sub
    End If
    If guncelle Then
        If MsgBox(KS_Metin(satir(1, A_NO)) & " güncellensin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    KS_Yaz T_ARIZA, idx, satir, IIf(guncelle, "ARIZA GÜNCELLE", "ARIZA EKLE"), KS_Metin(satir(1, A_NO)), _
        KS_Metin(satir(1, A_SEVK)) & " / " & KS_Metin(satir(1, A_KOD)) & " / " & KS_Metin(satir(1, A_DURUM))
    KS_Temizle KS_AlanA()
    KS_F("fA_NO").ClearContents
    MsgBox "Arýza " & IIf(guncelle, "güncellendi: ", "kaydedildi: ") & KS_Metin(satir(1, A_NO)) & vbCrLf & _
        "Garanti kapsamý: " & KS_Metin(satir(1, A_KAPSAM)) & "   |   Durum: " & KS_Metin(satir(1, A_DURUM)), _
        vbInformation, KS_AD
    Exit Sub
Hata:
    KS_HataGoster "Arýza kaydedilemedi."
End Sub

Public Sub KS_ARIZA_GETIR()
    Dim idx As Long, lo As ListObject, a As Variant
    On Error GoTo Hata
    Set lo = KS_LOZ(T_ARIZA)
    idx = KS_SatirBul(lo, A_NO, KS_FV("fA_NO"))
    If idx = 0 Then
        MsgBox "Arýza no bulunamadý: " & KS_Metin(KS_FV("fA_NO")), vbExclamation, KS_AD
        Exit Sub
    End If
    a = lo.ListRows(idx).Range.Value
    KS_Doldur KS_AlanA(), Array(a(1, A_SEVK), a(1, A_SERI), a(1, A_TARIH), a(1, A_ADET), a(1, A_TIP), _
        a(1, A_ACIK), a(1, A_BILDIRIM), a(1, A_DONUS), a(1, A_COZUM), a(1, A_SEKIL), a(1, A_MALIYET), _
        a(1, A_KARS), a(1, A_NOT))
    Exit Sub
Hata:
    KS_HataGoster "Arýza getirilemedi."
End Sub

Public Sub KS_ARIZA_TEMIZLE()
    KS_Temizle KS_AlanA()
    KS_F("fA_NO").ClearContents
End Sub
