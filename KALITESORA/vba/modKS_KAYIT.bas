Attribute VB_Name = "modKS_KAYIT"
Option Explicit

' KAYIT sayfasý: sekmeli giriþ formu (Malzeme / Sevkiyat / Arýza), hafýzadan otomatik doldurma,
' ortak doðrulama motoru ve sayfa deðiþiklik yönlendiricisi (Workbook_SheetChange -> KS_DEGISIKLIK).

Private dMalz As Object, aMalz As Variant
Private dSevk As Object, aSevk As Variant, dKodSevk As Object
Private dSevkci As Object, dYer As Object, dIrs As Object
Private mSevkSayac As Long, mArizaSayac As Long
Private mOtoTanim As Boolean
Private mOlayda As Boolean

' Form bloklarý (satýr aralýklarý)
Private Const B1_ILK As Long = 4
Private Const B1_SON As Long = 12
Private Const B2_ILK As Long = 13
Private Const B2_SON As Long = 24
Private Const B3_ILK As Long = 25
Private Const B3_SON As Long = 45

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
    Set dKodSevk = KS_KodSevkleri(aSevk)
    Set dSevkci = KS_Sozluk(T_SEVKCI)
    Set dYer = KS_Sozluk(T_YER)
    mSevkSayac = KS_EnBuyukNo(loS, "SVK-")
    mArizaSayac = KS_EnBuyukNo(KS_LOZ(T_ARIZA), "ARZ-")
    mOtoTanim = False
End Sub

Public Sub KS_OtoTanimAyarla(ByVal acik As Boolean)
    mOtoTanim = acik
End Sub

' Toplu aktarýmda reddedilen yeni sevkiyatýn numarasý bir sonrakine býrakýlýr.
Public Sub KS_SevkSayacGeri()
    mSevkSayac = mSevkSayac - 1
End Sub

Public Function KS_MalzemeVar(ByVal kod As Variant) As Boolean
    KS_MalzemeVar = dMalz.Exists(KS_Anahtar(kod))
End Function

Public Sub KS_MalzemeKaydiEkle(ByVal kod As Variant)
    dMalz(KS_Anahtar(kod)) = 0
End Sub

' Kod ya da adla malzeme bulur: tam kod, tam ad, sonra tek "içerir" eþleþmesi.
' Bulunamazsa 0, birden çok eþleþmede -(eþleþme sayýsý) döndürür.
Public Function KS_MalzemeAra(ByVal metin As Variant) As Long
    Dim k As String, i As Long, n As Long, son As Long
    k = KS_Anahtar(metin)
    If Len(k) = 0 Or Not IsArray(aMalz) Then Exit Function
    If dMalz.Exists(k) Then
        KS_MalzemeAra = dMalz(k)
        Exit Function
    End If
    For i = 1 To UBound(aMalz, 1)
        If KS_Anahtar(aMalz(i, M_AD)) = k Then
            KS_MalzemeAra = i
            Exit Function
        End If
    Next i
    For i = 1 To UBound(aMalz, 1)
        If InStr(1, KS_Anahtar(aMalz(i, M_AD)), k, vbBinaryCompare) > 0 Or _
            InStr(1, KS_Anahtar(aMalz(i, M_KOD)), k, vbBinaryCompare) > 0 Then
            n = n + 1
            son = i
        End If
    Next i
    If n = 1 Then KS_MalzemeAra = son Else KS_MalzemeAra = -n
End Function

' Listede olmayan tanýmý toplu aktarýmda isteðe baðlý ekler.
Private Function KS_TanimKontrol(ByVal d As Object, ByVal tablo As String, ByVal deger As String, _
        ByVal alan As String, ByRef hata As String) As Boolean
    Dim lo As ListObject, v As Variant, korumali As Boolean
    If d.Exists(KS_Anahtar(deger)) Then
        KS_TanimKontrol = True
        Exit Function
    End If
    If Not mOtoTanim Then
        hata = alan & " TANIMLAR listesinde yok: " & deger & ". YÖNETÝM > TANIMLAR sayfasýndan ekleyin."
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
    If tablo = T_SEVKCI Then v(1, lo.ListColumns.Count) = "Aktarýmda eklendi"
    korumali = lo.Parent.ProtectContents
    KS_Ac lo.Parent
    KS_SatirYaz lo, 0, v
    If korumali Then KS_Kapat lo.Parent
    d(KS_Anahtar(deger)) = lo.ListRows.Count
    KS_TanimKontrol = True
End Function

' ================================================================ doðrulama motoru

' g: Kod, Ad, Kategori, Birim, Üretici, Birim Fiyat, Para Birimi, Garanti (Ay), Not
Public Function KS_MalzemeHazirla(ByVal g As Variant, ByVal idx As Long, ByRef satir As Variant, _
        ByRef hata As String) As Boolean
    Dim kod As String, fiyat As Double, gar As Double
    hata = ""
    kod = KS_Metin(g(0))
    If Len(kod) = 0 Then hata = "Malzeme kodu boþ olamaz.": Exit Function
    If Len(kod) > 40 Then hata = "Malzeme kodu 40 karakteri geçemez.": Exit Function
    If idx = 0 And dMalz.Exists(KS_Anahtar(kod)) Then hata = "Bu malzeme kodu zaten kayýtlý: " & kod: Exit Function
    If Len(KS_Metin(g(1))) = 0 Then hata = "Malzeme adý boþ olamaz.": Exit Function
    If Not KS_Bos(g(5)) Then
        If Not KS_Sayi(g(5), fiyat) Then hata = "Birim fiyat sayý olmalý.": Exit Function
        If fiyat < 0 Then hata = "Birim fiyat negatif olamaz.": Exit Function
    End If
    If Not KS_Bos(g(7)) Then
        If Not KS_Sayi(g(7), gar) Then hata = "Garanti süresi (ay) sayý olmalý.": Exit Function
        If gar < 0 Or gar > 240 Or gar <> Int(gar) Then hata = "Garanti süresi 0-240 arasý tam ay olmalý.": Exit Function
    End If
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
    If KS_Bos(g(5)) Then satir(1, M_FIYAT) = Empty Else satir(1, M_FIYAT) = fiyat
    satir(1, M_PB) = UCase$(KS_Metin(g(6)))
    If Len(satir(1, M_PB)) = 0 Then satir(1, M_PB) = "TRY"
    If KS_Bos(g(7)) Then satir(1, M_GARANTI) = Empty Else satir(1, M_GARANTI) = CLng(gar)
    satir(1, M_NOT) = KS_Metin(g(8))
    KS_MalzemeHazirla = True
End Function

' g: Sevkçi, Sevk Yeri, Malzeme Kodu, Adet, Birim Fiyat, Para Birimi, Planlanan, Gerçek Geliþ,
'    Garanti (Ay), Belge No, Seri, Not
Public Function KS_SevkHazirla(ByVal g As Variant, ByVal idx As Long, ByRef satir As Variant, _
        ByRef hata As String, ByRef uyari As String) As Boolean
    Dim mi As Long, adet As Double, fiyat As Double, plan As Date, gelis As Date, k As String, no As String
    Dim gar As Double, pb As String
    hata = "": uyari = ""
    If Len(KS_Metin(g(0))) = 0 Then hata = "Sevkçi seçilmeli.": Exit Function
    If Not KS_TanimKontrol(dSevkci, T_SEVKCI, KS_Metin(g(0)), "Sevkçi", hata) Then Exit Function
    If Len(KS_Metin(g(1))) = 0 Then hata = "Sevk yeri seçilmeli.": Exit Function
    If Not KS_TanimKontrol(dYer, T_YER, KS_Metin(g(1)), "Sevk yeri", hata) Then Exit Function
    If Len(KS_Metin(g(2))) = 0 Then hata = "Malzeme kodu girilmeli.": Exit Function
    If Not dMalz.Exists(KS_Anahtar(g(2))) Then hata = "Malzeme kodu kayýtlý deðil: " & KS_Metin(g(2)) & _
        ". Önce MALZEME sekmesinden tanýmlayýn.": Exit Function
    mi = dMalz(KS_Anahtar(g(2)))
    If Not KS_Sayi(g(3), adet) Then hata = "Adet / miktar sayý olmalý.": Exit Function
    If adet <= 0 Then hata = "Adet / miktar sýfýrdan büyük olmalý.": Exit Function
    pb = UCase$(KS_Metin(g(5)))
    If KS_Bos(g(4)) Then
        If mi > 0 Then
            If KS_Sayi(aMalz(mi, M_FIYAT), fiyat) Then
                If Len(pb) = 0 Then pb = UCase$(KS_Metin(aMalz(mi, M_PB)))
            End If
        End If
    Else
        If Not KS_Sayi(g(4), fiyat) Then hata = "Birim fiyat sayý olmalý.": Exit Function
        If fiyat < 0 Then hata = "Birim fiyat negatif olamaz.": Exit Function
    End If
    If Len(pb) = 0 Then pb = "TRY"
    If Not KS_Tarih(g(6), plan) Then hata = "Planlanan teslim tarihi geçerli deðil (gg.aa.yyyy).": Exit Function
    If Not KS_Bos(g(7)) Then
        If Not KS_Tarih(g(7), gelis) Then hata = "Gerçek geliþ tarihi geçerli deðil (gg.aa.yyyy).": Exit Function
        If gelis > Date Then hata = "Gerçek geliþ tarihi bugünden ileri olamaz.": Exit Function
    End If
    If Not KS_Bos(g(8)) Then
        If Not KS_Sayi(g(8), gar) Then hata = "Garanti süresi (ay) sayý olmalý.": Exit Function
        If gar < 0 Or gar > 240 Then hata = "Garanti süresi 0-240 ay arasý olmalý.": Exit Function
    End If

    If idx = 0 Then
        satir = KS_BosSatir(S_KOLON)
        mSevkSayac = mSevkSayac + 1
        no = KS_NoBicim("SVK-", mSevkSayac)
        satir(1, S_NO) = no
        satir(1, S_ZAMAN) = Now
        satir(1, S_KULL) = KS_Kullanici()
    Else
        satir = KS_SatirKopya(aSevk, idx)
        no = KS_Metin(satir(1, S_NO))
    End If
    satir(1, S_SEVKCI) = KS_Metin(g(0))
    satir(1, S_YER) = KS_Metin(g(1))
    If mi > 0 Then
        satir(1, S_KOD) = KS_Metin(aMalz(mi, M_KOD))
        satir(1, S_AD) = aMalz(mi, M_AD)
    Else
        satir(1, S_KOD) = KS_Metin(g(2))
    End If
    satir(1, S_ADET) = adet
    If KS_Bos(g(4)) And fiyat = 0 Then satir(1, S_FIYAT) = Empty Else satir(1, S_FIYAT) = fiyat
    satir(1, S_PB) = pb
    satir(1, S_PLAN) = plan
    If KS_Bos(g(7)) Then satir(1, S_GELIS) = Empty Else satir(1, S_GELIS) = gelis
    If KS_Bos(g(8)) Then
        satir(1, S_GARAY) = Empty
        If mi > 0 Then
            If Not KS_Bos(aMalz(mi, M_GARANTI)) Then satir(1, S_GARAY) = aMalz(mi, M_GARANTI)
        End If
    Else
        satir(1, S_GARAY) = gar
    End If
    satir(1, S_IRS) = KS_Metin(g(9))
    satir(1, S_SERI) = KS_Metin(g(10))
    satir(1, S_NOT) = KS_Metin(g(11))
    KS_SevkHesapla satir, 1

    k = KS_Anahtar(g(9))
    If Len(k) > 0 Then
        k = k & "|" & KS_Anahtar(g(2))
        If dIrs.Exists(k) Then
            If dIrs(k) <> no Then uyari = "Ayný belge no ve malzeme ile kayýt var: " & dIrs(k)
        Else
            dIrs(k) = no
        End If
    End If
    KS_SevkHazirla = True
End Function

' g: Arýza Tarihi, Malzeme Kodu, Seri No, Adet, Tip, Açýklama, Baðlý Sevk, Garanti Devam (EVET/HAYIR),
'    Bildirim, Mail Yazýþmasý, Mail Dosyasý, Dönüþ, Devam Ediyor mu, Çözüldü mü, Çözüm Tarihi,
'    Çözüm Þekli, Tamir Maliyeti, Karþýlayan, Not
Public Function KS_ArizaHazirla(ByVal g As Variant, ByVal idx As Long, ByVal mevcut As Variant, _
        ByRef satir As Variant, ByRef hata As String) As Boolean
    Dim si As Long, mi As Long, t As Date, gelis As Date, bil As Date, don As Date, coz As Date
    Dim adet As Double, sevkAdet As Double, maliyet As Double, hb As Boolean, kap As String, oto As String
    Dim bit As Date
    hata = ""
    If Not KS_Tarih(g(0), t) Then hata = "Arýza tarihi geçerli deðil (gg.aa.yyyy).": Exit Function
    If t > Date Then hata = "Arýza tarihi bugünden ileri olamaz.": Exit Function
    If Len(KS_Metin(g(1))) = 0 Then hata = "Malzeme kodu girilmeli.": Exit Function
    If Not dMalz.Exists(KS_Anahtar(g(1))) Then hata = "Malzeme kodu kayýtlý deðil: " & KS_Metin(g(1)): Exit Function
    mi = dMalz(KS_Anahtar(g(1)))
    If Not KS_Sayi(g(3), adet) Then hata = "Arýzalý adet sayý olmalý.": Exit Function
    If adet <= 0 Then hata = "Arýzalý adet sýfýrdan büyük olmalý.": Exit Function
    If Len(KS_Metin(g(6))) > 0 Then
        If Not dSevk.Exists(KS_Anahtar(g(6))) Then hata = "Baðlý sevkiyat bulunamadý: " & KS_Metin(g(6)): Exit Function
        si = dSevk(KS_Anahtar(g(6)))
        If KS_Anahtar(aSevk(si, S_KOD)) <> KS_Anahtar(g(1)) Then
            hata = "Baðlý sevkiyat (" & KS_Metin(g(6)) & ") baþka bir malzemeye ait.": Exit Function
        End If
    ElseIf dKodSevk.Exists(KS_Anahtar(g(1))) Then
        si = KS_SevkSec(aSevk, dKodSevk(KS_Anahtar(g(1))), KS_Metin(g(2)), t)
    End If
    If si > 0 Then
        If KS_Tarih(aSevk(si, S_GELIS), gelis) Then
            If t < gelis Then hata = "Arýza tarihi, malzemenin geliþ tarihinden (" & _
                Format$(gelis, "dd.mm.yyyy") & ") önce olamaz.": Exit Function
        End If
        If KS_Sayi(aSevk(si, S_ADET), sevkAdet) Then
            If adet > sevkAdet Then hata = "Arýzalý adet (" & adet & "), sevk edilen miktarý (" & sevkAdet & ") aþamaz.": Exit Function
        End If
    End If
    If Not KS_Bos(g(8)) Then
        If Not KS_Tarih(g(8), bil) Then hata = "Üreticiye bildirim tarihi geçerli deðil.": Exit Function
        If bil < t Then hata = "Bildirim tarihi arýza tarihinden önce olamaz.": Exit Function
        If bil > Date Then hata = "Bildirim tarihi bugünden ileri olamaz.": Exit Function
        hb = True
    End If
    If Not KS_Bos(g(11)) Then
        If Not hb Then hata = "Üretici dönüþ tarihi için önce bildirim tarihi girilmeli.": Exit Function
        If Not KS_Tarih(g(11), don) Then hata = "Üretici dönüþ tarihi geçerli deðil.": Exit Function
        If don < bil Then hata = "Dönüþ tarihi bildirim tarihinden önce olamaz.": Exit Function
        If don > Date Then hata = "Dönüþ tarihi bugünden ileri olamaz.": Exit Function
    End If
    If Not KS_Bos(g(14)) Then
        If Not KS_Tarih(g(14), coz) Then hata = "Çözüm tarihi geçerli deðil.": Exit Function
        If coz < t Then hata = "Çözüm tarihi arýza tarihinden önce olamaz.": Exit Function
        If coz > Date Then hata = "Çözüm tarihi bugünden ileri olamaz.": Exit Function
    End If
    If Not KS_Bos(g(16)) Then
        If Not KS_Sayi(g(16), maliyet) Then hata = "Tamir maliyeti sayý olmalý.": Exit Function
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
    satir(1, A_TARIH) = t
    satir(1, A_KOD) = KS_Metin(aMalz(mi, M_KOD))
    satir(1, A_AD) = aMalz(mi, M_AD)
    satir(1, A_URETICI) = aMalz(mi, M_URETICI)
    satir(1, A_SERI) = KS_Metin(g(2))
    satir(1, A_ADET) = adet
    satir(1, A_TIP) = KS_Metin(g(4))
    satir(1, A_ACIK) = KS_Metin(g(5))
    If si > 0 Then
        satir(1, A_SEVK) = KS_Metin(aSevk(si, S_NO))
        satir(1, A_SEVKCI) = aSevk(si, S_SEVKCI)
        satir(1, A_YER) = aSevk(si, S_YER)
        satir(1, A_GARBIT) = aSevk(si, S_GARBIT)
    Else
        satir(1, A_SEVK) = Empty
        satir(1, A_GARBIT) = Empty
    End If
    ' Garanti: form deðeri otomatik hesaptan farklýysa "ELLE" olarak saklanýr.
    oto = "BELÝRSÝZ"
    If KS_Tarih(satir(1, A_GARBIT), bit) Then
        If t <= bit Then oto = "EVET" Else oto = "HAYIR"
    End If
    kap = UCase$(KS_Metin(g(7)))
    If Len(kap) > 0 And kap <> oto Then
        satir(1, A_KAPSAM) = kap
        satir(1, A_KAPKAYNAK) = "ELLE"
    Else
        satir(1, A_KAPSAM) = oto
        satir(1, A_KAPKAYNAK) = "OTOMATÝK"
    End If
    If KS_Bos(g(8)) Then satir(1, A_BILDIRIM) = Empty Else satir(1, A_BILDIRIM) = bil
    satir(1, A_MAIL) = Left$(KS_Metin(g(9)), 32000)
    satir(1, A_MAILDOSYA) = KS_Metin(g(10))
    If KS_Bos(g(11)) Then satir(1, A_DONUS) = Empty Else satir(1, A_DONUS) = don
    satir(1, A_DEVAM) = UCase$(KS_Metin(g(12)))
    satir(1, A_COZULDU) = UCase$(KS_Metin(g(13)))
    If KS_Bos(g(14)) Then satir(1, A_COZUM) = Empty Else satir(1, A_COZUM) = coz
    satir(1, A_SEKIL) = KS_Metin(g(15))
    If KS_Bos(g(16)) Then satir(1, A_MALIYET) = Empty Else satir(1, A_MALIYET) = maliyet
    satir(1, A_KARS) = KS_Metin(g(17))
    satir(1, A_NOT) = KS_Metin(g(18))
    KS_ArizaHesapla satir, 1
    KS_ArizaHazirla = True
End Function

' ================================================================ form düzeni

' Alan tanýmý: etiket, ad, tür, zorunlu, liste, sütun (1 sol / 2 sað / 3 geniþ), satýr, yükseklik (satýr)
Private Function KS_Alanlar(ByVal blok As Long) As Variant
    Select Case blok
        Case 1
            KS_Alanlar = Array( _
                Array("Malzeme Kodu", "fM_KOD", "listeS", True, "lstMalzeme", 1, 5, 1), _
                Array("Malzeme Adý", "fM_AD", "listeS", True, "lstMalzemeAd", 1, 6, 1), _
                Array("Kategori", "fM_KAT", "listeS", False, "lstKategori", 1, 7, 1), _
                Array("Birim", "fM_BIRIM", "listeS", False, "lstBirim", 1, 8, 1), _
                Array("Üretici", "fM_URETICI", "listeS", False, "lstSevkci", 2, 5, 1), _
                Array("Birim Fiyat", "fM_FIYAT", "para", False, "", 2, 6, 1), _
                Array("Para Birimi", "fM_PB", "liste", False, "lstPB", 2, 7, 1), _
                Array("Garanti Süresi (Ay)", "fM_GARANTI", "sayi", False, "", 2, 8, 1), _
                Array("Not", "fM_NOT", "metin", False, "", 3, 9, 1))
        Case 2
            KS_Alanlar = Array( _
                Array("Sevk No  (boþ = yeni)", "fS_NO", "listeS", False, "lstSevk", 1, 14, 1), _
                Array("Malzeme Kodu", "fS_KOD", "listeS", True, "lstMalzeme", 1, 15, 1), _
                Array("Malzeme Adý", "fS_AD", "listeS", True, "lstMalzemeAd", 1, 16, 1), _
                Array("Sevkçi", "fS_SEVKCI", "listeS", True, "lstSevkci", 1, 17, 1), _
                Array("Sevk Yeri", "fS_YER", "liste", True, "lstYer", 1, 18, 1), _
                Array("Adet / Miktar", "fS_ADET", "sayi", True, "", 1, 19, 1), _
                Array("Seri / Lot No", "fS_SERI", "kod", False, "", 1, 20, 1), _
                Array("Birim Fiyat  (boþ = liste)", "fS_FIYAT", "para", False, "", 2, 14, 1), _
                Array("Para Birimi", "fS_PB", "liste", False, "lstPB", 2, 15, 1), _
                Array("Planlanan Teslim Tarihi", "fS_PLAN", "tarih", True, "", 2, 16, 1), _
                Array("Gerçek Geliþ Tarihi", "fS_GELIS", "tarih", False, "", 2, 17, 1), _
                Array("Garanti Süresi (Ay)", "fS_GARAY", "sayi", False, "", 2, 18, 1), _
                Array("Belge / SAS No", "fS_IRS", "kod", False, "", 2, 19, 1), _
                Array("» Garanti Bitiþ", "fS_GARBIT", "otoTarih", False, "", 2, 20, 1), _
                Array("Not", "fS_NOT", "metin", False, "", 3, 21, 1))
        Case Else
            KS_Alanlar = Array( _
                Array("Arýza No  (boþ = yeni)", "fA_NO", "listeS", False, "lstAriza", 1, 26, 1), _
                Array("Arýza Tarihi", "fA_TARIH", "tarih", True, "", 1, 27, 1), _
                Array("Malzeme Kodu", "fA_KOD", "listeS", True, "lstMalzeme", 1, 28, 1), _
                Array("Malzeme Adý", "fA_AD", "listeS", True, "lstMalzemeAd", 1, 29, 1), _
                Array("Seri No", "fA_SERI", "kod", False, "", 1, 30, 1), _
                Array("Arýzalý Adet", "fA_ADET", "sayi", True, "", 1, 31, 1), _
                Array("Arýza Tipi", "fA_TIP", "listeS", False, "lstTip", 1, 32, 1), _
                Array("Üreticiye Bildirim Tarihi", "fA_BILDIRIM", "tarih", False, "", 1, 33, 1), _
                Array("Üretici Dönüþ Tarihi", "fA_DONUS", "tarih", False, "", 1, 34, 1), _
                Array("Tamir Maliyeti (TL)", "fA_MALIYET", "para", False, "", 1, 35, 1), _
                Array("» Üretici", "fA_URETICI", "oto", False, "", 2, 26, 1), _
                Array("Baðlý Sevkiyat  (otomatik)", "fA_SEVK", "listeS", False, "lstSevk", 2, 27, 1), _
                Array("» Sevk Yeri / Sevkçi", "fA_YER", "oto", False, "", 2, 28, 1), _
                Array("» Garanti Bitiþ", "fA_GARBIT", "otoTarih", False, "", 2, 29, 1), _
                Array("Garanti devam ediyor mu?", "fA_KAPSAM", "liste", False, "lstEvetHayir", 2, 30, 1), _
                Array("Arýza devam ediyor mu?", "fA_DEVAM", "liste", False, "lstEvetHayir", 2, 31, 1), _
                Array("Çözüldü mü?", "fA_COZULDU", "liste", False, "lstCozuldu", 2, 32, 1), _
                Array("Çözüm Tarihi", "fA_COZUM", "tarih", False, "", 2, 33, 1), _
                Array("Çözüm Þekli", "fA_SEKIL", "listeS", False, "lstCozum", 2, 34, 1), _
                Array("Maliyeti Karþýlayan", "fA_KARS", "listeS", False, "lstKarsilayan", 2, 35, 1), _
                Array("Arýza Sebebi / Açýklama", "fA_ACIK", "metin", False, "", 3, 36, 2), _
                Array("Üretici ile Mail Yazýþmasý", "fA_MAIL", "metin", False, "", 3, 38, 3), _
                Array("Mail Dosyasý", "fA_MAILDOSYA", "dosya", False, "", 3, 41, 1), _
                Array("Not", "fA_NOT", "metin", False, "", 3, 42, 1))
    End Select
End Function

Public Sub KS_KayitFormKur()
    Dim ws As Worksheet, alt As Double, r As Long
    Set ws = KS_WS(SH_KAYIT)
    KS_SayfaSifirla ws
    ws.Columns("B").ColumnWidth = 1
    ws.Columns("C").ColumnWidth = 25
    ws.Columns("D").ColumnWidth = 36
    ws.Columns("E").ColumnWidth = 3
    ws.Columns("F").ColumnWidth = 25
    ws.Columns("G").ColumnWidth = 36
    ws.Columns("H").ColumnWidth = 3
    ws.Columns("I").ColumnWidth = 44
    ws.Columns("J").ColumnWidth = 2
    alt = KS_SayfaCerceve(ws, False)
    ws.Rows(1).RowHeight = alt + 6
    ws.Rows(2).RowHeight = 48
    ws.Rows(3).RowHeight = 10
    KS_SayfaBaslik ws, ws.Rows(2).Top + 6, "KAYIT MERKEZÝ", _
        "Kodu ya da adý yazýn; hafýzadaki bilgiler kendiliðinden dolar. Boþ kalanlarý siz tamamlayýn.  * zorunlu"
    For r = 4 To B3_SON
        ws.Rows(r).RowHeight = 24
    Next r
    KS_FormBlok ws, 1, B1_ILK, B1_SON, "MALZEME BÝLGÝLERÝ", "Malzeme ana kaydý: kod, ad, üretici, fiyat ve garanti süresi."
    KS_FormBlok ws, 2, B2_ILK, B2_SON, "SEVKÝYAT BÝLGÝLERÝ", "Kim, nereye, ne zaman, kaç adet gönderdi; garanti geliþ tarihinden baþlar."
    KS_FormBlok ws, 3, B3_ILK, B3_SON, "ARIZA BÝLGÝLERÝ", "Malzeme ve seri no ile baðlý sevkiyat ve garanti durumu otomatik bulunur."
    KS_KapsamRenk ws
    KS_Sekmeler ws
    KS_ZeminUygula ws, True
    KS_SekmeGoster 1
    KS_Pencere ws, 90, 0
End Sub

Private Sub KS_Sekmeler(ByVal ws As Worksheet)
    Dim et As Variant, mk As Variant, i As Long, x As Double, gen As Double
    et = Array("01   MALZEME", "02   SEVKÝYAT", "03   ARIZA")
    mk = Array("KS_SEKME_MALZ", "KS_SEKME_SEVK", "KS_SEKME_ARIZA")
    gen = 130
    x = ws.Cells(1, 2).Left + KS_GENISLIK - 3 * gen - 2 * 6
    For i = 0 To 2
        KS_Buton ws, "ksT_" & i, CStr(et(i)), x + i * (gen + 6), ws.Rows(2).Top + 10, gen, 30, CStr(mk(i)), "cizgi"
    Next i
End Sub

Private Sub KS_FormBlok(ByVal ws As Worksheet, ByVal blok As Long, ByVal ilk As Long, ByVal son As Long, _
        ByVal baslik As String, ByVal alt As String)
    Dim alanlar As Variant, a As Variant, i As Long, e As Range, v As Range, kL As Long, kV As Long
    Dim kart As Range, bilgi As Range, x As Double, ust As Double, bb As Variant, b As Long, onek As String, r As Long

    onek = "ksF" & blok & "_"
    Set kart = ws.Range(ws.Cells(ilk, 3), ws.Cells(son - 1, 7))
    kart.Interior.Color = vbWhite
    KS_Cerceve kart
    ' baþlýk bandý
    With ws.Range(ws.Cells(ilk, 3), ws.Cells(ilk, 7))
        .Interior.Color = C_LAC()
        .Font.Color = vbWhite
    End With
    ws.Rows(ilk).RowHeight = 30
    With ws.Cells(ilk, 3)
        .Value = baslik
        .Font.Name = KS_FONT_KALIN
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = vbWhite
        .IndentLevel = 1
    End With
    With ws.Cells(ilk, 5)
        .Value = alt
        .Font.Size = 8.5
        .Font.Italic = True
        .Font.Color = RGB(170, 205, 235)
    End With

    alanlar = KS_Alanlar(blok)
    For i = 0 To UBound(alanlar)
        a = alanlar(i)
        If a(5) = 2 Then kL = 6 Else kL = 3
        kV = kL + 1
        Set e = ws.Cells(a(6), kL)
        If a(5) = 3 Then
            Set v = ws.Range(ws.Cells(a(6), 4), ws.Cells(a(6) + a(7) - 1, 7))
        Else
            Set v = ws.Range(ws.Cells(a(6), kV), ws.Cells(a(6) + a(7) - 1, kV))
        End If
        e.Value = a(0) & IIf(a(3), "  *", "")
        e.Font.Size = 9
        e.IndentLevel = 1
        e.VerticalAlignment = xlCenter
        If a(7) > 1 Then e.VerticalAlignment = xlTop
        If Left$(a(0), 1) = "»" Then
            e.Font.Color = C_TEAL()
        Else
            e.Font.Color = RGB(70, 85, 105)
        End If
        If a(3) Then e.Characters(Len(e.Value), 1).Font.Color = C_KIRMIZI()
        If v.Cells.Count > 1 Then v.Merge
        KS_AdHucre CStr(a(1)), v.Cells(1, 1)
        With v
            .Font.Size = 10
            .HorizontalAlignment = xlLeft
            .VerticalAlignment = IIf(a(7) > 1, xlTop, xlCenter)
            .WrapText = (a(7) > 1)
            .IndentLevel = 1
            With .Borders(xlEdgeBottom)
                .LineStyle = xlContinuous
                .Color = RGB(176, 194, 216)
                .Weight = xlThin
            End With
            If a(7) > 1 Then KS_Cerceve v
        End With
        Select Case a(2)
            Case "oto", "otoTarih", "dosya"
                v.Interior.Color = RGB(243, 248, 251)
                v.Font.Color = C_TEAL()
                v.Font.Bold = True
                v.Locked = True
                If a(2) = "otoTarih" Then v.NumberFormat = KS_FmtTarih()
                If a(2) = "dosya" Then v.Font.Bold = False: v.Font.Underline = xlUnderlineStyleSingle
            Case Else
                v.Interior.Color = RGB(250, 252, 255)
                v.Font.Color = C_LAC()
                v.Locked = False
                Select Case a(2)
                    Case "kod", "metin": v.NumberFormat = "@"
                    Case "liste": v.NumberFormat = "@": KS_Liste v, CStr(a(4)), True
                    Case "listeS": v.NumberFormat = "@": KS_Liste v, CStr(a(4)), False
                    Case "para": v.NumberFormat = "#,##0.00": KS_SayiDogrulama v
                    Case "sayi": v.NumberFormat = "General": KS_SayiDogrulama v
                    Case "tarih": v.NumberFormat = KS_FmtTarih(): KS_TarihDogrulama v
                End Select
        End Select
    Next i

    ' Hafýza paneli (sað kart)
    Set bilgi = ws.Range(ws.Cells(ilk, 9), ws.Cells(son - 1, 9))
    bilgi.Interior.Color = vbWhite
    KS_Cerceve bilgi
    With ws.Cells(ilk, 9)
        .Value = "HAFIZA  ·  kayýt geçmiþi"
        .Font.Name = KS_FONT_KALIN
        .Font.Bold = True
        .Font.Size = 10.5
        .Font.Color = vbWhite
        .Interior.Color = C_TEAL()
        .IndentLevel = 1
    End With
    Set bilgi = ws.Range(ws.Cells(ilk + 1, 9), ws.Cells(son - 1, 9))
    bilgi.Merge
    With bilgi
        .WrapText = True
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Font.Size = 9.5
        .Font.Color = RGB(60, 75, 95)
        .Value = "Kod ya da ad yazýldýðýnda bu alanda kayýt geçmiþi görünür."
    End With
    KS_AdHucre "fBilgi" & blok, bilgi.Cells(1, 1)

    ' Düðmeler
    If blok = 1 Then
        bb = Array("KAYDET", "KS_KAYIT_MALZEME", "ana", "TEMÝZLE", "KS_TEMIZLE_MALZ", "cizgi", "MALZEMELERÝ AÇ", "KS_GIT_MALZ", "teal")
    ElseIf blok = 2 Then
        bb = Array("KAYDET", "KS_KAYIT_SEVK", "ana", "TEMÝZLE", "KS_TEMIZLE_SEVK", "cizgi", "SEVKÝYATLARI AÇ", "KS_GIT_SEVK", "teal")
    Else
        bb = Array("KAYDET", "KS_KAYIT_ARIZA", "ana", "TEMÝZLE", "KS_TEMIZLE_ARIZA", "cizgi", "ARIZALARI AÇ", "KS_GIT_ARIZA", "teal", _
            "ÜRETÝCÝYE MAÝL", "KS_MAIL_TASLAK", "koyu")
    End If
    ust = ws.Rows(son - 1).Top + (ws.Rows(son - 1).RowHeight - 26) / 2
    x = ws.Cells(1, 4).Left
    For b = 0 To (UBound(bb) + 1) \ 3 - 1
        With KS_Buton(ws, onek & b, CStr(bb(b * 3)), x + b * 136, ust, 128, 26, CStr(bb(b * 3 + 1)), CStr(bb(b * 3 + 2)))
            .Placement = xlMove
        End With
    Next b
    If blok = 3 Then
        r = KS_F("fA_MAILDOSYA").Row
        With KS_Buton(ws, onek & "Mail", "DOSYA EKLE", ws.Cells(1, 8).Left - 150, ws.Rows(r).Top + 1, 82, 22, _
                "KS_MAIL_EKLE", "koyu")
            .Placement = xlMove
        End With
        With KS_Buton(ws, onek & "MailAc", "AÇ", ws.Cells(1, 8).Left - 64, ws.Rows(r).Top + 1, 60, 22, _
                "KS_MAIL_AC", "cizgi")
            .Placement = xlMove
        End With
        ws.Range(ws.Cells(r, 4), ws.Cells(r, 7)).UnMerge
        ws.Range(ws.Cells(r, 4), ws.Cells(r, 6)).Merge
    End If
    ws.Rows(son - 1).RowHeight = 34
    ws.Rows(son).RowHeight = 14
End Sub

' Ýnce, açýk gri dýþ çerçeve.
Private Sub KS_Cerceve(ByVal r As Range)
    Dim k As Variant
    For Each k In Array(xlEdgeLeft, xlEdgeTop, xlEdgeRight, xlEdgeBottom)
        With r.Borders(k)
            .LineStyle = xlContinuous
            .Color = RGB(208, 219, 232)
            .Weight = xlThin
        End With
    Next k
End Sub

Private Sub KS_KapsamRenk(ByVal ws As Worksheet)
    Dim r As Range, fc As Object
    Set r = KS_F("fA_KAPSAM")
    r.FormatConditions.Delete
    Set fc = r.FormatConditions.Add(Type:=xlTextString, String:="EVET", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_YESIL()
    fc.Interior.Color = RGB(220, 245, 232)
    Set fc = r.FormatConditions.Add(Type:=xlTextString, String:="HAYIR", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_KIRMIZI()
    fc.Interior.Color = RGB(253, 226, 228)
End Sub

' Sekme: yalnýz seçili formun satýrlarý ve düðmeleri görünür.
Public Sub KS_SekmeGoster(ByVal blok As Long)
    Dim ws As Worksheet, korumali As Boolean, i As Long, s As Shape, sekme As Shape
    Set ws = KS_WS(SH_KAYIT)
    korumali = ws.ProtectContents
    KS_Ac ws
    ws.Rows(B1_ILK & ":" & B1_SON).Hidden = (blok <> 1)
    ws.Rows(B2_ILK & ":" & B2_SON).Hidden = (blok <> 2)
    ws.Rows(B3_ILK & ":" & B3_SON).Hidden = (blok <> 3)
    For Each s In ws.Shapes
        If Left$(s.Name, 3) = "ksF" Then s.Visible = (Mid$(s.Name, 4, 1) = CStr(blok))
    Next s
    For i = 0 To 2
        Set sekme = ws.Shapes("ksT_" & i)
        If i + 1 = blok Then
            sekme.Fill.ForeColor.RGB = C_LAC()
            sekme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
            sekme.Line.Visible = msoFalse
        Else
            sekme.Fill.ForeColor.RGB = vbWhite
            sekme.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = C_LAC()
            sekme.Line.Visible = msoTrue
        End If
    Next i
    If korumali Then KS_Kapat ws
    ws.Activate
    Select Case blok
        Case 1: KS_F("fM_KOD").Select
        Case 2: KS_F("fS_KOD").Select
        Case Else: KS_F("fA_TARIH").Select
    End Select
End Sub

Public Sub KS_SEKME_MALZ()
    KS_SekmeGoster 1
End Sub

Public Sub KS_SEKME_SEVK()
    KS_SekmeGoster 2
End Sub

Public Sub KS_SEKME_ARIZA()
    KS_SekmeGoster 3
End Sub

Public Sub KS_GIT_KAYIT_MALZ()
    KS_SekmeGoster 1
End Sub

Public Sub KS_GIT_KAYIT_SEVK()
    KS_SekmeGoster 2
End Sub

Public Sub KS_GIT_KAYIT_ARIZA()
    KS_SekmeGoster 3
End Sub

' ================================================================ form yardýmcýlarý

Private Function KS_AlanM() As Variant
    KS_AlanM = Array("fM_KOD", "fM_AD", "fM_KAT", "fM_BIRIM", "fM_URETICI", "fM_FIYAT", "fM_PB", "fM_GARANTI", "fM_NOT")
End Function

Private Function KS_AlanS() As Variant
    KS_AlanS = Array("fS_SEVKCI", "fS_YER", "fS_KOD", "fS_ADET", "fS_FIYAT", "fS_PB", "fS_PLAN", "fS_GELIS", _
        "fS_GARAY", "fS_IRS", "fS_SERI", "fS_NOT")
End Function

Private Function KS_AlanA() As Variant
    KS_AlanA = Array("fA_TARIH", "fA_KOD", "fA_SERI", "fA_ADET", "fA_TIP", "fA_ACIK", "fA_SEVK", "fA_KAPSAM", _
        "fA_BILDIRIM", "fA_MAIL", "fA_MAILDOSYA", "fA_DONUS", "fA_DEVAM", "fA_COZULDU", "fA_COZUM", "fA_SEKIL", _
        "fA_MALIYET", "fA_KARS", "fA_NOT")
End Function

Private Sub KS_Temizle(ByVal adlar As Variant)
    Dim a As Variant
    For Each a In adlar
        KS_Yaz1 CStr(a), Empty
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

Private Sub KS_Yaz1(ByVal ad As String, ByVal deger As Variant)
    Dim ws As Worksheet, korumali As Boolean, r As Range
    Set r = KS_F(ad)
    Set ws = r.Worksheet
    If r.Locked Then
        korumali = ws.ProtectContents
        KS_Ac ws
        r.MergeArea.Cells(1, 1).Value = deger
        If korumali Then KS_Kapat ws
    Else
        r.MergeArea.Cells(1, 1).Value = deger
    End If
End Sub

Private Sub KS_Bilgi(ByVal blok As Long, ByVal metin As String)
    KS_Yaz1 "fBilgi" & blok, metin
End Sub

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
    mOlayda = False
    Application.EnableEvents = True
    KS_Bitir True
    MsgBox baslik & vbCrLf & "Hata: " & no & " / " & ac, vbExclamation, KS_AD
End Sub

Private Function KS_Tarihce(ByVal v As Variant) As String
    Dim d As Date
    If KS_Tarih(v, d) Then KS_Tarihce = Format$(d, "dd.mm.yyyy") Else KS_Tarihce = "-"
End Function

' Malzemenin sevkiyat / arýza geçmiþini kýsa metin olarak verir.
Private Function KS_MalzemeOzet(ByVal mi As Long) As String
    Dim k As String, aA As Variant, i As Long, n As Long, mik As Double, x As Double, sonS As Date, d As Date
    Dim arzN As Long, acikN As Long, sonA As Date, sonANo As String, s As String
    k = KS_Anahtar(aMalz(mi, M_KOD))
    If IsArray(aSevk) Then
        For i = 1 To UBound(aSevk, 1)
            If KS_Anahtar(aSevk(i, S_KOD)) = k Then
                n = n + 1
                If KS_Sayi(aSevk(i, S_ADET), x) Then mik = mik + x
                If KS_Tarih(aSevk(i, S_GELIS), d) Then If d > sonS Then sonS = d
            End If
        Next i
    End If
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Anahtar(aA(i, A_KOD)) = k Then
                arzN = arzN + 1
                If KS_ArizaAcik(aA(i, A_DURUM)) Then acikN = acikN + 1
                If KS_Tarih(aA(i, A_TARIH), d) Then
                    If d > sonA Then sonA = d: sonANo = KS_Metin(aA(i, A_NO))
                End If
            End If
        Next i
    End If
    s = KS_Metin(aMalz(mi, M_KOD)) & "  " & KS_Metin(aMalz(mi, M_AD)) & vbLf & vbLf
    s = s & "Üretici: " & IIf(KS_Bos(aMalz(mi, M_URETICI)), "-", KS_Metin(aMalz(mi, M_URETICI))) & vbLf
    s = s & "Garanti süresi: " & IIf(KS_Bos(aMalz(mi, M_GARANTI)), "GÝRÝLMEMÝÞ", KS_Metin(aMalz(mi, M_GARANTI)) & " ay") & vbLf
    s = s & "Liste fiyatý: " & IIf(KS_Bos(aMalz(mi, M_FIYAT)), "-", Format$(aMalz(mi, M_FIYAT), "#,##0.00") & " " & KS_Metin(aMalz(mi, M_PB))) & vbLf & vbLf
    s = s & "Sevkiyat: " & n & " kayýt, " & Format$(mik, "#,##0") & " birim" & vbLf
    s = s & "Son teslim: " & IIf(sonS > 0, Format$(sonS, "dd.mm.yyyy"), "-") & vbLf
    s = s & "Arýza: " & arzN & " kayýt" & IIf(acikN > 0, "  (" & acikN & " açýk)", "") & vbLf
    s = s & "Son arýza: " & IIf(sonA > 0, Format$(sonA, "dd.mm.yyyy") & "  " & sonANo, "-")
    KS_MalzemeOzet = s
End Function

' ================================================================ otomatik doldurma (sayfa olaylarý)

' ThisWorkbook.Workbook_SheetChange buradan geçer.
Public Sub KS_DEGISIKLIK(ByVal Sh As Object, ByVal Target As Range)
    If mOlayda Then Exit Sub
    If Not Application.EnableEvents Then Exit Sub
    If Target.Cells.CountLarge > 200 Then Exit Sub
    On Error GoTo Son
    mOlayda = True
    Select Case Sh.Name
        Case SH_KAYIT: KS_KayitDegisti Target
        Case SH_SEVK, SH_MALZ: KS_TabloDegisti Sh, Target
        Case SH_KARNE: KS_KarneDegisti Target
    End Select
Son:
    mOlayda = False
    Application.EnableEvents = True
End Sub

Private Function KS_AlanAdi(ByVal hucre As Range) As String
    Dim n As Name, ad As String
    For Each n In ThisWorkbook.Names
        ad = n.Name
        If Left$(ad, 2) = "fM" Or Left$(ad, 2) = "fS" Or Left$(ad, 2) = "fA" Then
            On Error Resume Next
            If Not Intersect(n.RefersToRange.MergeArea, hucre) Is Nothing Then
                If Err.Number = 0 Then
                    KS_AlanAdi = ad
                    Exit Function
                End If
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Next n
End Function

Private Sub KS_KayitDegisti(ByVal Target As Range)
    Dim ad As String
    If Target.Cells.CountLarge > 1 Then Exit Sub
    ad = KS_AlanAdi(Target)
    If Len(ad) = 0 Then Exit Sub
    Application.EnableEvents = False
    KS_BaglamYukle
    Select Case ad
        Case "fM_KOD", "fM_AD": KS_DoldurMalzeme ad
        Case "fS_NO": KS_DoldurSevkNo
        Case "fS_KOD", "fS_AD": KS_DoldurSevkMalzeme ad
        Case "fS_GELIS", "fS_GARAY": KS_SevkGarantiOnizle
        Case "fA_NO": KS_DoldurArizaNo
        Case "fA_KOD", "fA_AD", "fA_SERI", "fA_TARIH": KS_DoldurArizaMalzeme ad
        Case "fA_SEVK": KS_ArizaSevkGoster
    End Select
    Application.EnableEvents = True
End Sub

Private Sub KS_DoldurMalzeme(ByVal ad As String)
    Dim mi As Long
    mi = KS_MalzemeAra(KS_FV(ad))
    If mi > 0 Then
        KS_Yaz1 "fM_KOD", aMalz(mi, M_KOD)
        KS_Yaz1 "fM_AD", aMalz(mi, M_AD)
        KS_Yaz1 "fM_KAT", aMalz(mi, M_KAT)
        KS_Yaz1 "fM_BIRIM", aMalz(mi, M_BIRIM)
        KS_Yaz1 "fM_URETICI", aMalz(mi, M_URETICI)
        KS_Yaz1 "fM_FIYAT", aMalz(mi, M_FIYAT)
        KS_Yaz1 "fM_PB", aMalz(mi, M_PB)
        KS_Yaz1 "fM_GARANTI", aMalz(mi, M_GARANTI)
        KS_Yaz1 "fM_NOT", aMalz(mi, M_NOT)
        KS_Bilgi 1, "KAYITLI MALZEME  -  deðiþiklik yapýp KAYDET'e basarsanýz güncellenir." & vbLf & vbLf & KS_MalzemeOzet(mi)
    ElseIf mi < 0 Then
        KS_Bilgi 1, -mi & " malzeme bu ifadeyi içeriyor. Listeden seçin ya da tam kodu yazýn."
    ElseIf ad = "fM_KOD" Then
        KS_Bilgi 1, "YENÝ MALZEME" & vbLf & vbLf & "Bu kod kayýtlý deðil. Bilgileri doldurup KAYDET'e basýn."
    End If
End Sub

Private Sub KS_DoldurSevkMalzeme(ByVal ad As String)
    Dim mi As Long, i As Long, k As String, sonSevkci As Variant, sonYer As Variant, d As Date, enSon As Date
    mi = KS_MalzemeAra(KS_FV(ad))
    If mi <= 0 Then
        If mi < 0 Then KS_Bilgi 2, -mi & " malzeme bu ifadeyi içeriyor. Listeden seçin." _
            Else KS_Bilgi 2, "Malzeme bulunamadý. Önce MALZEME sekmesinden tanýmlayýn."
        Exit Sub
    End If
    k = KS_Anahtar(aMalz(mi, M_KOD))
    If IsArray(aSevk) Then
        For i = 1 To UBound(aSevk, 1)
            If KS_Anahtar(aSevk(i, S_KOD)) = k Then
                If KS_Tarih(aSevk(i, S_PLAN), d) Then
                    If d >= enSon Then
                        enSon = d
                        sonSevkci = aSevk(i, S_SEVKCI)
                        sonYer = aSevk(i, S_YER)
                    End If
                End If
            End If
        Next i
    End If
    KS_Yaz1 "fS_KOD", aMalz(mi, M_KOD)
    KS_Yaz1 "fS_AD", aMalz(mi, M_AD)
    If KS_Bos(KS_FV("fS_FIYAT")) Then KS_Yaz1 "fS_FIYAT", aMalz(mi, M_FIYAT)
    If KS_Bos(KS_FV("fS_PB")) Then KS_Yaz1 "fS_PB", aMalz(mi, M_PB)
    If KS_Bos(KS_FV("fS_GARAY")) Then KS_Yaz1 "fS_GARAY", aMalz(mi, M_GARANTI)
    If KS_Bos(KS_FV("fS_SEVKCI")) Then
        If Not IsEmpty(sonSevkci) Then
            KS_Yaz1 "fS_SEVKCI", sonSevkci
        ElseIf Not KS_Bos(aMalz(mi, M_URETICI)) Then
            KS_Yaz1 "fS_SEVKCI", aMalz(mi, M_URETICI)
        End If
    End If
    If KS_Bos(KS_FV("fS_YER")) And Not IsEmpty(sonYer) Then KS_Yaz1 "fS_YER", sonYer
    KS_SevkGarantiOnizle
    KS_Bilgi 2, "Hafýzadan dolduruldu; boþ alanlarý tamamlayýn." & vbLf & vbLf & KS_MalzemeOzet(mi)
End Sub

Private Sub KS_SevkGarantiOnizle()
    Dim gelis As Date, gar As Double
    If KS_Tarih(KS_FV("fS_GELIS"), gelis) And KS_Sayi(KS_FV("fS_GARAY"), gar) Then
        KS_Yaz1 "fS_GARBIT", DateAdd("m", CLng(gar), gelis)
    Else
        KS_Yaz1 "fS_GARBIT", Empty
    End If
End Sub

Private Sub KS_DoldurSevkNo()
    Dim idx As Long, k As Variant, mi As Long
    k = KS_FV("fS_NO")
    If KS_Bos(k) Then
        KS_Bilgi 2, "Yeni sevkiyat: KAYDET'e basýnca numara verilir."
        Exit Sub
    End If
    If dSevk.Exists(KS_Anahtar(k)) Then idx = dSevk(KS_Anahtar(k))
    If idx = 0 Then
        KS_Bilgi 2, "Sevk no bulunamadý: " & KS_Metin(k)
        Exit Sub
    End If
    KS_Yaz1 "fS_SEVKCI", aSevk(idx, S_SEVKCI)
    KS_Yaz1 "fS_YER", aSevk(idx, S_YER)
    KS_Yaz1 "fS_KOD", aSevk(idx, S_KOD)
    KS_Yaz1 "fS_AD", aSevk(idx, S_AD)
    KS_Yaz1 "fS_ADET", aSevk(idx, S_ADET)
    KS_Yaz1 "fS_FIYAT", aSevk(idx, S_FIYAT)
    KS_Yaz1 "fS_PB", aSevk(idx, S_PB)
    KS_Yaz1 "fS_PLAN", aSevk(idx, S_PLAN)
    KS_Yaz1 "fS_GELIS", aSevk(idx, S_GELIS)
    KS_Yaz1 "fS_GARAY", aSevk(idx, S_GARAY)
    KS_Yaz1 "fS_IRS", aSevk(idx, S_IRS)
    KS_Yaz1 "fS_SERI", aSevk(idx, S_SERI)
    KS_Yaz1 "fS_NOT", aSevk(idx, S_NOT)
    KS_Yaz1 "fS_GARBIT", aSevk(idx, S_GARBIT)
    mi = KS_MalzemeAra(aSevk(idx, S_KOD))
    KS_Bilgi 2, "KAYITLI SEVKÝYAT  " & KS_Metin(aSevk(idx, S_NO)) & "  -  deðiþtirip KAYDET'e basýn." & vbLf & _
        "Teslim: " & KS_Metin(aSevk(idx, S_TDURUM)) & "   |   Garanti: " & KS_Metin(aSevk(idx, S_GDURUM)) & _
        IIf(mi > 0, vbLf & vbLf & KS_MalzemeOzet(mi), "")
End Sub

Private Sub KS_DoldurArizaMalzeme(ByVal ad As String)
    Dim mi As Long, si As Long, t As Variant, oncekiOto As String, yeniOto As String, bit As Date, d As Date
    If ad = "fA_KOD" Or ad = "fA_AD" Then
        mi = KS_MalzemeAra(KS_FV(ad))
        If mi <= 0 Then
            If mi < 0 Then KS_Bilgi 3, -mi & " malzeme bu ifadeyi içeriyor. Listeden seçin." _
                Else KS_Bilgi 3, "Malzeme bulunamadý. Önce MALZEME sekmesinden tanýmlayýn."
            Exit Sub
        End If
        KS_Yaz1 "fA_KOD", aMalz(mi, M_KOD)
        KS_Yaz1 "fA_AD", aMalz(mi, M_AD)
        KS_Yaz1 "fA_SEVK", Empty
    Else
        mi = KS_MalzemeAra(KS_FV("fA_KOD"))
        If mi <= 0 Then Exit Sub
    End If
    KS_Yaz1 "fA_URETICI", aMalz(mi, M_URETICI)
    t = KS_FV("fA_TARIH")
    If KS_Bos(KS_FV("fA_SEVK")) Or ad = "fA_SERI" Or ad = "fA_TARIH" Then
        If dKodSevk.Exists(KS_Anahtar(aMalz(mi, M_KOD))) Then
            si = KS_SevkSec(aSevk, dKodSevk(KS_Anahtar(aMalz(mi, M_KOD))), KS_Metin(KS_FV("fA_SERI")), t)
        End If
        If si > 0 Then KS_Yaz1 "fA_SEVK", aSevk(si, S_NO) Else KS_Yaz1 "fA_SEVK", Empty
    End If
    KS_ArizaSevkGoster
    KS_Bilgi 3, "Hafýzadan dolduruldu; baðlý sevkiyat " & IIf(KS_Bos(KS_FV("fA_SEVK")), "bulunamadý.", _
        "otomatik seçildi (deðiþtirebilirsiniz).") & vbLf & vbLf & KS_MalzemeOzet(mi)
End Sub

' Baðlý sevkiyata göre yer, garanti bitiþi ve garanti önerisini gösterir.
Private Sub KS_ArizaSevkGoster()
    Dim si As Long, bit As Date, t As Date, oto As String
    If dSevk.Exists(KS_Anahtar(KS_FV("fA_SEVK"))) Then si = dSevk(KS_Anahtar(KS_FV("fA_SEVK")))
    If si > 0 Then
        KS_Yaz1 "fA_YER", KS_Metin(aSevk(si, S_YER)) & "  /  " & KS_Metin(aSevk(si, S_SEVKCI))
        KS_Yaz1 "fA_GARBIT", aSevk(si, S_GARBIT)
        If KS_Tarih(aSevk(si, S_GARBIT), bit) Then
            If KS_Tarih(KS_FV("fA_TARIH"), t) Then
                If t <= bit Then oto = "EVET" Else oto = "HAYIR"
            ElseIf bit >= Date Then
                oto = "EVET"
            Else
                oto = "HAYIR"
            End If
        End If
    Else
        KS_Yaz1 "fA_YER", Empty
        KS_Yaz1 "fA_GARBIT", Empty
    End If
    KS_Yaz1 "fA_KAPSAM", IIf(Len(oto) > 0, oto, Empty)
End Sub

Private Sub KS_DoldurArizaNo()
    Dim lo As ListObject, idx As Long, a As Variant
    Set lo = KS_LOZ(T_ARIZA)
    idx = KS_SatirBul(lo, A_NO, KS_FV("fA_NO"))
    If idx = 0 Then
        If Not KS_Bos(KS_FV("fA_NO")) Then KS_Bilgi 3, "Arýza no bulunamadý: " & KS_Metin(KS_FV("fA_NO"))
        Exit Sub
    End If
    a = lo.ListRows(idx).Range.Value
    KS_Yaz1 "fA_TARIH", a(1, A_TARIH)
    KS_Yaz1 "fA_KOD", a(1, A_KOD)
    KS_Yaz1 "fA_AD", a(1, A_AD)
    KS_Yaz1 "fA_SERI", a(1, A_SERI)
    KS_Yaz1 "fA_ADET", a(1, A_ADET)
    KS_Yaz1 "fA_TIP", a(1, A_TIP)
    KS_Yaz1 "fA_URETICI", a(1, A_URETICI)
    KS_Yaz1 "fA_SEVK", a(1, A_SEVK)
    KS_Yaz1 "fA_YER", KS_Metin(a(1, A_YER)) & "  /  " & KS_Metin(a(1, A_SEVKCI))
    KS_Yaz1 "fA_GARBIT", a(1, A_GARBIT)
    KS_Yaz1 "fA_KAPSAM", IIf(KS_Metin(a(1, A_KAPSAM)) = "BELÝRSÝZ", Empty, a(1, A_KAPSAM))
    KS_Yaz1 "fA_BILDIRIM", a(1, A_BILDIRIM)
    KS_Yaz1 "fA_DONUS", a(1, A_DONUS)
    KS_Yaz1 "fA_ACIK", a(1, A_ACIK)
    KS_Yaz1 "fA_MAIL", a(1, A_MAIL)
    KS_Yaz1 "fA_MAILDOSYA", a(1, A_MAILDOSYA)
    KS_Yaz1 "fA_DEVAM", a(1, A_DEVAM)
    KS_Yaz1 "fA_COZULDU", a(1, A_COZULDU)
    KS_Yaz1 "fA_COZUM", a(1, A_COZUM)
    KS_Yaz1 "fA_SEKIL", a(1, A_SEKIL)
    KS_Yaz1 "fA_MALIYET", a(1, A_MALIYET)
    KS_Yaz1 "fA_KARS", a(1, A_KARS)
    KS_Yaz1 "fA_NOT", a(1, A_NOT)
    KS_Bilgi 3, "KAYITLI ARIZA  " & KS_Metin(a(1, A_NO)) & "  -  deðiþtirip KAYDET'e basýn." & vbLf & vbLf & _
        "Durum: " & KS_Metin(a(1, A_DURUM)) & vbLf & _
        "Garanti: " & KS_Metin(a(1, A_KAPSAM)) & " (" & KS_Metin(a(1, A_KAPKAYNAK)) & ")" & vbLf & _
        "Üretici dönüþ süresi: " & IIf(KS_Bos(a(1, A_DONUSGUN)), "-", KS_Metin(a(1, A_DONUSGUN)) & " gün") & vbLf & _
        "Önceki arýza: " & KS_Tarihce(a(1, A_ONCEKI)) & "   |   Tekrar: " & KS_Metin(a(1, A_TEKRAR))
End Sub

' SEVKÝYATLAR / MALZEMELER tablosunda elle düzenlenen sütunlar sonrasý yeniden hesap.
Private Sub KS_TabloDegisti(ByVal Sh As Object, ByVal Target As Range)
    Dim lo As ListObject
    If Sh.Name = SH_SEVK Then Set lo = KS_LO(T_SEVK) Else Set lo = KS_LO(T_MALZ)
    If lo Is Nothing Then Exit Sub
    If lo.DataBodyRange Is Nothing Then Exit Sub
    If Intersect(Target, lo.DataBodyRange) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    KS_HesaplaTumu
    Application.EnableEvents = True
End Sub

' ================================================================ KAYDET / TEMÝZLE

Public Sub KS_KAYIT_MALZEME()
    Dim g As Variant, satir As Variant, hata As String, idx As Long
    On Error GoTo Hata
    KS_BaglamYukle
    g = KS_Degerler(KS_AlanM())
    If dMalz.Exists(KS_Anahtar(g(0))) Then idx = dMalz(KS_Anahtar(g(0)))
    If Not KS_MalzemeHazirla(g, idx, satir, hata) Then
        MsgBox hata, vbExclamation, KS_AD
        Exit Sub
    End If
    If idx > 0 Then
        If MsgBox(KS_Metin(g(0)) & " kayýtlý. Bilgileri güncellensin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    KS_Yaz T_MALZ, idx, satir, IIf(idx > 0, "MALZEME GÜNCELLE", "MALZEME EKLE"), _
        KS_Metin(satir(1, M_KOD)), KS_Metin(satir(1, M_AD))
    Application.EnableEvents = False
    KS_Temizle KS_AlanM()
    KS_Bilgi 1, "Kaydedildi: " & KS_Metin(satir(1, M_KOD)) & "  " & KS_Metin(satir(1, M_AD))
    Application.EnableEvents = True
    MsgBox "Malzeme " & IIf(idx > 0, "güncellendi: ", "kaydedildi: ") & KS_Metin(satir(1, M_KOD)), vbInformation, KS_AD
    Exit Sub
Hata:
    KS_HataGoster "Malzeme kaydedilemedi."
End Sub

Public Sub KS_KAYIT_SEVK()
    Dim g As Variant, satir As Variant, hata As String, uyari As String, idx As Long, no As Variant
    On Error GoTo Hata
    KS_BaglamYukle
    g = KS_Degerler(KS_AlanS())
    no = KS_FV("fS_NO")
    If Not KS_Bos(no) Then
        If dSevk.Exists(KS_Anahtar(no)) Then idx = dSevk(KS_Anahtar(no))
        If idx = 0 Then
            MsgBox "Sevk no bulunamadý: " & KS_Metin(no) & vbCrLf & "Yeni kayýt için Sevk No alanýný boþ býrakýn.", vbExclamation, KS_AD
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
    If idx > 0 Then
        If MsgBox(KS_Metin(satir(1, S_NO)) & " güncellensin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    KS_Yaz T_SEVK, idx, satir, IIf(idx > 0, "SEVK GÜNCELLE", "SEVK EKLE"), KS_Metin(satir(1, S_NO)), _
        KS_Metin(satir(1, S_KOD)) & " x " & satir(1, S_ADET) & " > " & KS_Metin(satir(1, S_YER))
    Application.EnableEvents = False
    KS_Temizle KS_AlanS()
    KS_Temizle Array("fS_NO", "fS_AD", "fS_GARBIT")
    KS_Bilgi 2, "Kaydedildi: " & KS_Metin(satir(1, S_NO)) & vbLf & "Garanti: " & KS_Metin(satir(1, S_GDURUM))
    Application.EnableEvents = True
    MsgBox "Sevkiyat " & IIf(idx > 0, "güncellendi: ", "kaydedildi: ") & KS_Metin(satir(1, S_NO)) & vbCrLf & _
        "Garanti: " & KS_Metin(satir(1, S_GDURUM)), vbInformation, KS_AD
    Exit Sub
Hata:
    KS_HataGoster "Sevkiyat kaydedilemedi."
End Sub

Public Sub KS_KAYIT_ARIZA()
    Dim g As Variant, satir As Variant, hata As String, idx As Long, lo As ListObject, mevcut As Variant
    On Error GoTo Hata
    KS_BaglamYukle
    g = KS_Degerler(KS_AlanA())
    Set lo = KS_LOZ(T_ARIZA)
    If Not KS_Bos(KS_FV("fA_NO")) Then
        idx = KS_SatirBul(lo, A_NO, KS_FV("fA_NO"))
        If idx = 0 Then
            MsgBox "Arýza no bulunamadý: " & KS_Metin(KS_FV("fA_NO")) & vbCrLf & _
                "Yeni kayýt için Arýza No alanýný boþ býrakýn.", vbExclamation, KS_AD
            Exit Sub
        End If
        mevcut = lo.ListRows(idx).Range.Value
    End If
    If Not KS_ArizaHazirla(g, idx, mevcut, satir, hata) Then
        MsgBox hata, vbExclamation, KS_AD
        Exit Sub
    End If
    If idx > 0 Then
        If MsgBox(KS_Metin(satir(1, A_NO)) & " güncellensin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    satir(1, A_MAILDOSYA) = KS_MailTasi(KS_Metin(satir(1, A_MAILDOSYA)), KS_Metin(satir(1, A_NO)))
    KS_Yaz T_ARIZA, idx, satir, IIf(idx > 0, "ARIZA GÜNCELLE", "ARIZA EKLE"), KS_Metin(satir(1, A_NO)), _
        KS_Metin(satir(1, A_KOD)) & " / " & KS_Metin(satir(1, A_DURUM))
    Application.EnableEvents = False
    KS_Temizle KS_AlanA()
    KS_Temizle Array("fA_NO", "fA_AD", "fA_URETICI", "fA_YER", "fA_GARBIT")
    KS_Bilgi 3, "Kaydedildi: " & KS_Metin(satir(1, A_NO)) & vbLf & "Durum: " & KS_Metin(satir(1, A_DURUM))
    Application.EnableEvents = True
    MsgBox "Arýza " & IIf(idx > 0, "güncellendi: ", "kaydedildi: ") & KS_Metin(satir(1, A_NO)) & vbCrLf & _
        "Garanti: " & KS_Metin(satir(1, A_KAPSAM)) & "   |   Durum: " & KS_Metin(satir(1, A_DURUM)), _
        vbInformation, KS_AD
    Exit Sub
Hata:
    KS_HataGoster "Arýza kaydedilemedi."
End Sub

Public Sub KS_TEMIZLE_MALZ()
    Application.EnableEvents = False
    KS_Temizle KS_AlanM()
    KS_Bilgi 1, "Kod ya da ad yazýldýðýnda bu alanda kayýt geçmiþi görünür."
    Application.EnableEvents = True
End Sub

Public Sub KS_TEMIZLE_SEVK()
    Application.EnableEvents = False
    KS_Temizle KS_AlanS()
    KS_Temizle Array("fS_NO", "fS_AD", "fS_GARBIT")
    KS_Bilgi 2, "Kod ya da ad yazýldýðýnda bu alanda kayýt geçmiþi görünür."
    Application.EnableEvents = True
End Sub

Public Sub KS_TEMIZLE_ARIZA()
    Application.EnableEvents = False
    KS_Temizle KS_AlanA()
    KS_Temizle Array("fA_NO", "fA_AD", "fA_URETICI", "fA_YER", "fA_GARBIT")
    KS_Bilgi 3, "Kod ya da ad yazýldýðýnda bu alanda kayýt geçmiþi görünür."
    Application.EnableEvents = True
End Sub

' ================================================================ mail dosyasý

Public Sub KS_MAIL_EKLE()
    Dim f As Variant
    On Error GoTo Hata
    f = Application.GetOpenFilename("E-posta ve belgeler (*.msg;*.eml;*.pdf;*.docx;*.xlsx;*.png;*.jpg),*.msg;*.eml;*.pdf;*.docx;*.xlsx;*.png;*.jpg", , _
        "Üretici ile yazýþma dosyasýný seçin")
    If VarType(f) = vbBoolean Then Exit Sub
    Application.EnableEvents = False
    KS_Yaz1 "fA_MAILDOSYA", CStr(f)
    Application.EnableEvents = True
    KS_Bilgi 3, "Mail dosyasý seçildi. KAYDET'e basýnca KALITESORA_MAILLER klasörüne kopyalanýr:" & vbLf & CStr(f)
    Exit Sub
Hata:
    Application.EnableEvents = True
    MsgBox "Dosya eklenemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

' Dosyayý <dosya klasörü>\KALITESORA_MAILLER\<Arýza No>\ altýna kopyalar; yeni yolu döndürür.
Private Function KS_MailTasi(ByVal kaynak As String, ByVal no As String) As String
    Dim fso As Object, klasor As String, hedef As String
    KS_MailTasi = kaynak
    If Len(kaynak) = 0 Then Exit Function
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Function
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FileExists(kaynak) Then Exit Function
    klasor = fso.BuildPath(ThisWorkbook.Path, "KALITESORA_MAILLER")
    If InStr(1, kaynak, klasor, vbTextCompare) = 1 Then Exit Function
    If Not fso.FolderExists(klasor) Then fso.CreateFolder klasor
    klasor = fso.BuildPath(klasor, no)
    If Not fso.FolderExists(klasor) Then fso.CreateFolder klasor
    hedef = fso.BuildPath(klasor, fso.GetFileName(kaynak))
    fso.CopyFile kaynak, hedef, True
    KS_MailTasi = hedef
End Function

Public Sub KS_MAIL_AC()
    Dim yol As String
    yol = KS_Metin(KS_FV("fA_MAILDOSYA"))
    If Len(yol) = 0 Then
        MsgBox "Bu arýzaya eklenmiþ mail dosyasý yok.", vbInformation, KS_AD
        Exit Sub
    End If
    If Len(Dir$(yol)) = 0 Then
        MsgBox "Dosya bulunamadý:" & vbCrLf & yol, vbExclamation, KS_AD
        Exit Sub
    End If
    On Error GoTo Hata
    ThisWorkbook.FollowHyperlink yol
    Exit Sub
Hata:
    MsgBox "Dosya açýlamadý: " & Err.Description, vbExclamation, KS_AD
End Sub

' Formdaki arýza bilgileriyle üreticiye Outlook mail taslaðý hazýrlar; bildirim tarihi boþsa bugün yazýlýr.
Public Sub KS_MAIL_TASLAK()
    Dim ol As Object, m As Object, konu As String, govde As String, t As Date, satirlar As String, kayit As String
    If KS_Bos(KS_FV("fA_KOD")) Then
        MsgBox "Önce arýzalý malzemeyi seçin.", vbExclamation, KS_AD
        Exit Sub
    End If
    konu = "Arýza Bildirimi - " & KS_Metin(KS_FV("fA_KOD")) & " " & KS_Metin(KS_FV("fA_AD"))
    If Not KS_Bos(KS_FV("fA_NO")) Then konu = konu & "  (" & KS_Metin(KS_FV("fA_NO")) & ")"
    satirlar = "Malzeme kodu      : " & KS_Metin(KS_FV("fA_KOD")) & vbCrLf & _
        "Malzeme adý       : " & KS_Metin(KS_FV("fA_AD")) & vbCrLf
    If Not KS_Bos(KS_FV("fA_SERI")) Then satirlar = satirlar & "Seri no           : " & KS_Metin(KS_FV("fA_SERI")) & vbCrLf
    If KS_Tarih(KS_FV("fA_TARIH"), t) Then satirlar = satirlar & "Arýza tarihi      : " & Format$(t, "dd.mm.yyyy") & vbCrLf
    If Not KS_Bos(KS_FV("fA_ADET")) Then satirlar = satirlar & "Arýzalý adet      : " & KS_Metin(KS_FV("fA_ADET")) & vbCrLf
    If Not KS_Bos(KS_FV("fA_YER")) Then satirlar = satirlar & "Sevk yeri / sevkçi: " & KS_Metin(KS_FV("fA_YER")) & vbCrLf
    If Not KS_Bos(KS_FV("fA_SEVK")) Then satirlar = satirlar & "Ýlgili sevkiyat   : " & KS_Metin(KS_FV("fA_SEVK")) & vbCrLf
    If KS_Tarih(KS_FV("fA_GARBIT"), t) Then satirlar = satirlar & "Garanti bitiþi    : " & Format$(t, "dd.mm.yyyy") & vbCrLf
    If Not KS_Bos(KS_FV("fA_TIP")) Then satirlar = satirlar & "Arýza tipi        : " & KS_Metin(KS_FV("fA_TIP")) & vbCrLf
    govde = "Sayýn Yetkili," & vbCrLf & vbCrLf & _
        "Firmanýzca tedarik edilen aþaðýdaki malzemede arýza tespit edilmiþtir." & vbCrLf & vbCrLf & satirlar & vbCrLf & _
        "Arýza açýklamasý:" & vbCrLf & KS_Metin(KS_FV("fA_ACIK")) & vbCrLf & vbCrLf & _
        IIf(KS_Metin(KS_FV("fA_KAPSAM")) = "EVET", "Malzeme garanti süresi içindedir; garanti kapsamýnda deðerlendirilmesini rica ederiz." & vbCrLf, "") & _
        "En kýsa sürede dönüþ yapmanýzý rica ederiz." & vbCrLf & vbCrLf & "Saygýlarýmýzla," & vbCrLf & _
        "Çoruh Elektrik Daðýtým A.Þ."
    On Error Resume Next
    Set ol = CreateObject("Outlook.Application")
    On Error GoTo Hata
    Application.EnableEvents = False
    If KS_Bos(KS_FV("fA_BILDIRIM")) Then KS_Yaz1 "fA_BILDIRIM", Date
    kayit = Format$(Now, "dd.mm.yyyy hh:nn") & " - Üreticiye arýza bildirimi hazýrlandý."
    If KS_Bos(KS_FV("fA_MAIL")) Then KS_Yaz1 "fA_MAIL", kayit Else KS_Yaz1 "fA_MAIL", KS_Metin(KS_FV("fA_MAIL")) & vbLf & kayit
    Application.EnableEvents = True
    If ol Is Nothing Then
        KS_Yaz1 "fA_MAIL", KS_Metin(KS_FV("fA_MAIL")) & vbLf & vbLf & konu & vbLf & govde
        MsgBox "Outlook bulunamadý. Mail metni 'Üretici ile Mail Yazýþmasý' alanýna yazýldý; kopyalayýp gönderebilirsiniz.", _
            vbInformation, KS_AD
        Exit Sub
    End If
    Set m = ol.CreateItem(0)
    m.Subject = konu
    m.Body = govde
    If Len(KS_Metin(KS_FV("fA_MAILDOSYA"))) > 0 Then
        If Len(Dir$(KS_Metin(KS_FV("fA_MAILDOSYA")))) > 0 Then m.Attachments.Add KS_Metin(KS_FV("fA_MAILDOSYA"))
    End If
    m.Display
    KS_Bilgi 3, "Outlook taslaðý açýldý; alýcýyý ekleyip gönderin. Bildirim tarihi yazýldý, KAYDET'e basmayý unutmayýn."
    Exit Sub
Hata:
    Application.EnableEvents = True
    MsgBox "Mail taslaðý hazýrlanamadý: " & Err.Description, vbExclamation, KS_AD
End Sub
