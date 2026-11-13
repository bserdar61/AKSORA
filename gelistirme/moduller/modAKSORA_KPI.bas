Attribute VB_Name = "modAKSORA_KPI"
Option Explicit

'==========================================================
' AKSORA v3.1  |  modAKSORA_KPI
'
' YONETIM KARNESI (trend) ve YONETICI OZETI (tek sayfa PDF)
'
' Yonetim Merkezi bugune kadar yalniz "o anin" fotografini gosteriyordu.
' Bu modul her rapor yenilemesinde temel gostergeleri gunde bir satir olarak
' cok gizli "_AKSORA_KPI" sayfasina kaydeder ve Yonetim Merkezi'ne
' "onceki kayda gore / 30 gun oncesine gore" degisim tablosu ekler.
'
' Baglanti noktalari (kurulum betigi ekler):
'   ADM_KUR            : r = AKSORA_KPI_YER_AYIR(ws, r)      -> tablo icin satir ayirir
'   ADM_KUR (sonu)     : AKSORA_KPI_TAMAMLA ...              -> kaydeder ve cizer
'   ADM_ANALIZ_UYARI   : AKSORA_KPI_ANALIZ ...               -> kritik stok / transfer sayilari
'   ADM_ATIL_CIZ       : AKSORA_KPI_ATIL v, kapsam           -> atil stok sayilari
'
' Kural: bu moduldeki hicbir hata Yonetim Merkezi'nin acilmasini engellemez.
'==========================================================

Private Const KPI_SAYFA As String = "_AKSORA_KPI"
Private Const KPI_SUTUN As Long = 24
Private Const KPI_AZAMI As Long = 4000
Private Const KPI_ADET As Long = 8          ' gosterge sayisi
Private Const KPI_SERI As Long = 12         ' egilim cizgisindeki kayit sayisi
Private Const KPI_SERI_SUTUN As Long = 27   ' egilim verisi: _AKSORA_KPI!AA..AL

' Kayit sutunlari
Private Const K_TARIH As Long = 1, K_KAPSAM As Long = 2, K_SKOR As Long = 3, K_KRITIK As Long = 4, K_UYARI As Long = 5
Private Const K_TEKIL As Long = 6, K_MALZEME As Long = 7, K_YATIRIM As Long = 8, K_MUSTERI As Long = 9, K_ILCE As Long = 10
Private Const K_ATIL_KALEM As Long = 11, K_ATIL_TL As Long = 12, K_ATIL_ORAN As Long = 13
Private Const K_KS_KALEM As Long = 14, K_KS_TL As Long = 15, K_TR_ADET As Long = 16, K_TR_TL As Long = 17
Private Const K_IP_ADET As Long = 18, K_IP_TL As Long = 19, K_FT_ADET As Long = 20, K_FT_TL As Long = 21
Private Const K_DEMONTAJ As Long = 22, K_ISG As Long = 23, K_KULLANICI As Long = 24

' Rapor kurulurken toplanan ara degerler
Private mYerSatir As Long
Private mAnalizVar As Boolean
Private mAnK As Long, mAnTK As Double, mAnT As Long, mAnTT As Double
Private mAnIp As Long, mAnTip As Double, mAnF As Long, mAnTF As Double
Private mAtilVar As Boolean
Private mAtil(0 To 4, 0 To 2) As Double     ' il: atil kalem, atil deger, kapsam (toplam stok degeri)

' Son cizilen karne (PDF ozeti icin)
Private mOzVar As Boolean
Private mOzIl As String
Private mOzOnceki As String
Private mOzTablo(1 To 8, 1 To 7) As String  ' gosterge, bugun, onceki, degisim, 30 gun once, degisim, degerlendirme
Private mOzYon(1 To 8) As Long              ' degerlendirme rengi: 1 iyi, -1 kotu, 0 notr
Private mOzIller(0 To 4, 0 To 3) As Double  ' tekil deger, atil kalem, atil deger, kapsam
Private mOzAtilVar As Boolean
Private mOzUyari() As String
Private mOzUyariSay As Long
Private mOzIsler As Collection
Private mOzKapsam As String
Private mOzBirim(0 To 3) As Double
Private mOzHam(1 To 8, 1 To 3) As Variant   ' gosterge: bugun, onceki, 30 gun once (ham deger)

'==========================================================
' BAGLANTI NOKTALARI
'==========================================================

' Yonetim Merkezi'nde "BUGUN YAPILACAK ISLER" tablosunun altinda karne icin yer ayirir. Donus: ayrilan son satir.
Public Function AKSORA_KPI_YER_AYIR(ByVal ws As Worksheet, ByVal r As Long) As Long
    AKSORA_KPI_YER_AYIR = r + KPI_ADET + 4
    On Error Resume Next
    mYerSatir = r + 2
    ws.Cells.SparklineGroups.Clear
End Function

' Malzeme analizlerinden gelen sayilar (ADM_ANALIZ_UYARI icinden).
Public Sub AKSORA_KPI_ANALIZ(ByVal nKritik As Long, ByVal tKritik As Double, ByVal nTransfer As Long, ByVal tTransfer As Double, _
    ByVal nIptal As Long, ByVal tIptal As Double, ByVal nFazla As Long, ByVal tFazla As Double)
    mAnK = nKritik: mAnTK = tKritik: mAnT = nTransfer: mAnTT = tTransfer
    mAnIp = nIptal: mAnTip = tIptal: mAnF = nFazla: mAnTF = tFazla
    mAnalizVar = True
End Sub

' Atil stok ozetinden gelen il bazli sayilar (ADM_ATIL_CIZ icinden). v: (il, 0 atil kalem / 1 atil deger ...), kapsam: il toplam deger.
Public Sub AKSORA_KPI_ATIL(ByRef v() As Double, ByRef kapsam() As Double)
    Dim p As Long
    On Error Resume Next
    For p = 0 To 4
        mAtil(p, 0) = v(p, 0): mAtil(p, 1) = v(p, 1): mAtil(p, 2) = kapsam(p)
    Next p
    mAtilVar = (Err.Number = 0)
End Sub

' Rapor kurulumunun sonunda: gostergeleri kaydeder, karneyi ve dugmeleri cizer.
Public Sub AKSORA_KPI_TAMAMLA(ByVal ws As Worksheet, ByVal secIl As String, ByVal ilMetin As String, ByVal tekilToplam As Double, _
    ByRef ilTekil() As Double, ByRef deger() As Double, ByVal skor As Long, ByVal krit As Long, ByVal uyari As Long, _
    ByVal demTl As Double, ByVal isgZ As Double, ByVal uy As Collection)
    Dim kayit(1 To 1, 1 To 24) As Variant, iller As Variant, p As Long, kapsam As String
    Dim aKalem As Double, aTl As Double, aKapsam As Double, gecmis As Variant, x As Variant, sev As Long
    Dim hataMetni As String
    On Error GoTo Hata
    If mYerSatir <= 0 Then GoTo Temizle
    iller = ADM_ILLER()
    kapsam = IIf(Len(secIl) = 0, "TUMU", secIl)

    kayit(1, K_TARIH) = CDbl(Now)
    kayit(1, K_KAPSAM) = kapsam
    kayit(1, K_SKOR) = skor
    kayit(1, K_KRITIK) = krit
    kayit(1, K_UYARI) = uyari
    kayit(1, K_TEKIL) = Round(tekilToplam, 2)
    kayit(1, K_MALZEME) = Round(deger(0), 2)
    kayit(1, K_YATIRIM) = Round(deger(1), 2)
    kayit(1, K_MUSTERI) = Round(deger(2), 2)
    kayit(1, K_ILCE) = Round(deger(3), 2)
    If mAtilVar Then
        For p = 0 To 4
            If Len(secIl) = 0 Or CStr(iller(p)) = secIl Then
                aKalem = aKalem + mAtil(p, 0): aTl = aTl + mAtil(p, 1): aKapsam = aKapsam + mAtil(p, 2)
            End If
        Next p
        kayit(1, K_ATIL_KALEM) = aKalem
        kayit(1, K_ATIL_TL) = Round(aTl, 2)
        If aKapsam > 0 Then kayit(1, K_ATIL_ORAN) = Round(aTl / aKapsam, 4)
    End If
    If mAnalizVar Then
        kayit(1, K_KS_KALEM) = mAnK: kayit(1, K_KS_TL) = Round(mAnTK, 2)
        kayit(1, K_TR_ADET) = mAnT: kayit(1, K_TR_TL) = Round(mAnTT, 2)
        kayit(1, K_IP_ADET) = mAnIp: kayit(1, K_IP_TL) = Round(mAnTip, 2)
        kayit(1, K_FT_ADET) = mAnF: kayit(1, K_FT_TL) = Round(mAnTF, 2)
    End If
    kayit(1, K_DEMONTAJ) = Round(demTl, 2)
    kayit(1, K_ISG) = isgZ
    kayit(1, K_KULLANICI) = AKSORA_KULLANICI()

    ' Izleme (yayin) kopyasinda kayit tutulmaz; gecmis ana dosyadan gelir.
    If Not AKSORA_YAYIN_MODU() Then KPI_KAYDET kayit
    gecmis = KPI_GECMIS(kapsam)

    ' PDF ozeti icin il tablosu ve uyarilar
    For p = 0 To 4
        mOzIller(p, 0) = ilTekil(p)
        mOzIller(p, 1) = mAtil(p, 0): mOzIller(p, 2) = mAtil(p, 1): mOzIller(p, 3) = mAtil(p, 2)
    Next p
    mOzAtilVar = mAtilVar
    mOzUyariSay = 0
    ReDim mOzUyari(1 To 8)
    If Not uy Is Nothing Then
        For sev = 1 To 2
            For Each x In uy
                If x(0) = sev And mOzUyariSay < 8 Then
                    mOzUyariSay = mOzUyariSay + 1
                    mOzUyari(mOzUyariSay) = IIf(sev = 1, "KRİTİK", "UYARI") & vbTab & CStr(x(1))
                End If
            Next x
        Next sev
    End If
    mOzIl = ilMetin
    mOzKapsam = kapsam
    For p = 0 To 3
        mOzBirim(p) = deger(p)
    Next p

    KPI_CIZ ws, mYerSatir, ilMetin, kayit, gecmis
    Set mOzIsler = KPI_ISLERI_OKU(ws)
    KPI_DUGMELER ws
    mOzVar = True

Temizle:
    mAnalizVar = False: mAtilVar = False: mYerSatir = 0
    Exit Sub
Hata:
    hataMetni = Err.Description
    On Error Resume Next
    AKSORA_GUNLUK_YAZ "HATA", "Yönetim karnesi çizilemedi: " & hataMetni
    mAnalizVar = False: mAtilVar = False: mYerSatir = 0
End Sub

Public Function AKSORA_KPI_KAYIT_SAYISI() As Long
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(KPI_SAYFA)
    If ws Is Nothing Then Exit Function
    AKSORA_KPI_KAYIT_SAYISI = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row - 1
    If AKSORA_KPI_KAYIT_SAYISI < 0 Then AKSORA_KPI_KAYIT_SAYISI = 0
End Function

'==========================================================
' KAYIT
'==========================================================
Private Function KPI_BASLIKLAR() As Variant
    KPI_BASLIKLAR = Array("Tarih", "Kapsam", "Skor", "Kritik uyarı", "Uyarı", "Tekil stok değeri", "Malzeme", "Yatırım", "Müşteri", "İlçe", _
        "Atıl kalem", "Atıl değer", "Atıl oran", "Kritik stok kalem", "Kritik stok TL", "Transfer adet", "Transfer TL", _
        "İptal adet", "İptal TL", "Fazla teslim adet", "Fazla teslim TL", "Demontaj eksik TL", "İSG zimmetsiz", "Kullanıcı")
End Function

' Ayni gun + ayni kapsam icin tek satir: gun icindeki son rapor oncekinin uzerine yazilir.
Private Sub KPI_KAYDET(ByRef kayit() As Variant)
    Dim kws As Worksheet, son As Long, v As Variant, i As Long, hedef As Long, gun As Long
    Set kws = AKSORA_GIZLI_SAYFA(KPI_SAYFA, KPI_BASLIKLAR())
    If kws Is Nothing Then Exit Sub
    son = kws.Cells(kws.Rows.Count, 1).End(xlUp).Row
    gun = CLng(Int(CDbl(kayit(1, K_TARIH))))
    If son >= 2 Then
        v = kws.Range(kws.Cells(2, 1), kws.Cells(son, 2)).Value2
        For i = UBound(v, 1) To 1 Step -1
            If KPI_SAYI_MI(v(i, 1)) Then
                If CLng(Int(CDbl(v(i, 1)))) < gun Then Exit For
                If CLng(Int(CDbl(v(i, 1)))) = gun And CStr(v(i, 2)) = CStr(kayit(1, K_KAPSAM)) Then hedef = i + 1: Exit For
            End If
        Next i
    End If
    If hedef = 0 Then hedef = son + 1
    kws.Range(kws.Cells(hedef, 1), kws.Cells(hedef, KPI_SUTUN)).Value2 = kayit
    kws.Cells(hedef, 1).NumberFormat = "dd.mm.yyyy hh:mm"
    If hedef > KPI_AZAMI Then kws.Range(kws.Cells(2, 1), kws.Cells(201, KPI_SUTUN)).Delete Shift:=xlUp
End Sub

' Kapsamin (TUMU ya da il) gecmisi, eskiden yeniye. Donus: Array(n, veri(1..n, 1..KPI_SUTUN)); kayit yoksa Array(0, Empty).
Private Function KPI_GECMIS(ByVal kapsam As String) As Variant
    Dim kws As Worksheet, son As Long, v As Variant, i As Long, j As Long, n As Long, d() As Variant
    KPI_GECMIS = Array(0, Empty)
    On Error Resume Next
    Set kws = ThisWorkbook.Worksheets(KPI_SAYFA)
    On Error GoTo 0
    If kws Is Nothing Then Exit Function
    son = kws.Cells(kws.Rows.Count, 1).End(xlUp).Row
    If son < 2 Then Exit Function
    v = kws.Range(kws.Cells(2, 1), kws.Cells(son, KPI_SUTUN)).Value2
    For i = 1 To UBound(v, 1)
        If CStr(v(i, K_KAPSAM)) = kapsam And KPI_SAYI_MI(v(i, K_TARIH)) Then n = n + 1
    Next i
    If n = 0 Then Exit Function
    ReDim d(1 To n, 1 To KPI_SUTUN)
    n = 0
    For i = 1 To UBound(v, 1)
        If CStr(v(i, K_KAPSAM)) = kapsam And KPI_SAYI_MI(v(i, K_TARIH)) Then
            n = n + 1
            For j = 1 To KPI_SUTUN
                d(n, j) = v(i, j)
            Next j
        End If
    Next i
    KPI_GECMIS = Array(n, d)
End Function

'==========================================================
' HESAP (saf islevler; AKSORA_KPI_SINAMA ile dogrulanir)
'==========================================================
Private Function KPI_SAYI_MI(ByVal a As Variant) As Boolean
    If IsEmpty(a) Then Exit Function
    If IsError(a) Then Exit Function
    If VarType(a) = vbString Then Exit Function
    KPI_SAYI_MI = IsNumeric(a)
End Function

' Gosterge tanimi: Array(ad, tur, yon)
'   tur: 1 puan, 2 TL, 3 oran, 4 adet        yon: 1 artmasi iyi, -1 azalmasi iyi, 0 notr
Private Function KPI_TANIM(ByVal k As Long) As Variant
    Select Case k
        Case 1: KPI_TANIM = Array("Sağlık skoru", 1, 1)
        Case 2: KPI_TANIM = Array("Tekil stok değeri", 2, 0)
        Case 3: KPI_TANIM = Array("Atıl stok değeri", 2, -1)
        Case 4: KPI_TANIM = Array("Atıl stok oranı", 3, -1)
        Case 5: KPI_TANIM = Array("Kritik stok kalemi", 4, -1)
        Case 6: KPI_TANIM = Array("Bekleyen tasarruf fırsatı", 2, 0)
        Case 7: KPI_TANIM = Array("Demontaj girişi eksiği", 2, -1)
        Case Else: KPI_TANIM = Array("Zimmetsiz İSG malzemesi", 4, -1)
    End Select
End Function

' Kayit satirindan gosterge degeri; veri yoksa Empty.
Private Function KPI_DEGER(ByRef v As Variant, ByVal i As Long, ByVal k As Long) As Variant
    Dim a As Variant, b As Variant
    KPI_DEGER = Empty
    Select Case k
        Case 1: a = v(i, K_SKOR)
        Case 2: a = v(i, K_TEKIL)
        Case 3: a = v(i, K_ATIL_TL)
        Case 4: a = v(i, K_ATIL_ORAN)
        Case 5: a = v(i, K_KS_KALEM)
        Case 6
            a = v(i, K_TR_TL): b = v(i, K_IP_TL)
            If Not KPI_SAYI_MI(a) And Not KPI_SAYI_MI(b) Then Exit Function
            If Not KPI_SAYI_MI(a) Then a = 0#
            If Not KPI_SAYI_MI(b) Then b = 0#
            a = CDbl(a) + CDbl(b)
        Case 7: a = v(i, K_DEMONTAJ)
        Case Else: a = v(i, K_ISG)
    End Select
    If Not KPI_SAYI_MI(a) Then Exit Function
    KPI_DEGER = CDbl(a)
End Function

Private Function KPI_METIN(ByVal a As Variant, ByVal tur As Long) As String
    If IsEmpty(a) Then KPI_METIN = "veri yok": Exit Function
    Select Case tur
        Case 1: KPI_METIN = CStr(CLng(a)) & " / 100"
        Case 2: KPI_METIN = ADM_TL_KISA(CDbl(a))
        Case 3: KPI_METIN = "%" & AKSORA_SAYI_METNI(CDbl(a) * 100, "0.0")
        Case Else: KPI_METIN = ADM_ADET(CDbl(a))
    End Select
End Function

' Degisim metni: "(yukari ok) 1,2 mn TL (%0,9)". Karsilastirilacak deger yoksa "-".
Private Function KPI_FARK(ByVal simdi As Variant, ByVal once As Variant, ByVal tur As Long) As String
    Dim d As Double, ok As String
    If IsEmpty(simdi) Or IsEmpty(once) Then KPI_FARK = "-": Exit Function
    d = CDbl(simdi) - CDbl(once)
    If Abs(d) < 0.00005 Then KPI_FARK = "değişmedi": Exit Function
    If tur = 2 And Abs(d) < 0.005 Then KPI_FARK = "değişmedi": Exit Function
    ok = IIf(d > 0, ChrW(9650), ChrW(9660)) & " "
    Select Case tur
        Case 1
            KPI_FARK = ok & CStr(CLng(Abs(d))) & " puan"
        Case 2
            KPI_FARK = ok & ADM_TL_KISA(Abs(d))
            If Abs(CDbl(once)) > 0.005 Then KPI_FARK = KPI_FARK & "  (%" & AKSORA_SAYI_METNI(Abs(d) / Abs(CDbl(once)) * 100, "0.0") & ")"
        Case 3
            KPI_FARK = ok & AKSORA_SAYI_METNI(Abs(d) * 100, "0.0") & " puan"
        Case Else
            KPI_FARK = ok & ADM_ADET(Abs(d))
    End Select
End Function

' Degisimin yonu: 1 iyilesti, -1 kotulesti, 0 notr / degismedi / karsilastirma yok.
Private Function KPI_YON(ByVal simdi As Variant, ByVal once As Variant, ByVal yon As Long) As Long
    Dim d As Double
    If IsEmpty(simdi) Or IsEmpty(once) Or yon = 0 Then Exit Function
    d = CDbl(simdi) - CDbl(once)
    If Abs(d) < 0.00005 Then Exit Function
    If (d > 0 And yon > 0) Or (d < 0 And yon < 0) Then KPI_YON = 1 Else KPI_YON = -1
End Function

Private Function KPI_YORUM(ByVal simdi As Variant, ByVal once As Variant, ByVal yon As Long) As String
    Dim d As Double
    If IsEmpty(simdi) Then KPI_YORUM = "VERİ YOK": Exit Function
    If IsEmpty(once) Then KPI_YORUM = "İLK KAYIT": Exit Function
    d = CDbl(simdi) - CDbl(once)
    If Abs(d) < 0.00005 Then KPI_YORUM = "AYNI": Exit Function
    Select Case KPI_YON(simdi, once, yon)
        Case 1: KPI_YORUM = "İYİLEŞTİ"
        Case -1: KPI_YORUM = "KÖTÜLEŞTİ"
        Case Else: KPI_YORUM = IIf(d > 0, "ARTTI", "AZALDI")
    End Select
End Function

Private Function KPI_YON_RENK(ByVal yon As Long) As Long
    Select Case yon
        Case 1: KPI_YON_RENK = RGB(21, 128, 61)
        Case -1: KPI_YON_RENK = RGB(185, 28, 28)
        Case Else: KPI_YON_RENK = RGB(71, 85, 105)
    End Select
End Function

' Gecmiste, bugunden en az "gun" gun onceki en yeni kaydin sirasi (yoksa 0).
Private Function KPI_ONCEKI_SIRA(ByRef d As Variant, ByVal n As Long, ByVal bugun As Long, ByVal gun As Long) As Long
    Dim i As Long
    For i = n To 1 Step -1
        If CLng(Int(CDbl(d(i, K_TARIH)))) <= bugun - gun Then KPI_ONCEKI_SIRA = i: Exit Function
    Next i
End Function

'==========================================================
' CIZIM (Yonetim Merkezi)
'==========================================================
Private Sub KPI_CIZ(ByVal ws As Worksheet, ByVal r0 As Long, ByVal ilMetin As String, ByRef kayit As Variant, ByVal gecmis As Variant)
    Dim n As Long, d As Variant, iOnce As Long, iOtuz As Long, bugun As Long, k As Long, r As Long, t As Variant
    Dim simdi As Variant, once As Variant, otuz As Variant, yon As Long, yorum As String, tOnce As String, tOtuz As String
    Dim dolgu As Long, yazi As Long, ilk As String

    n = gecmis(0)
    If n > 0 Then d = gecmis(1)
    bugun = CLng(Int(CDbl(kayit(1, K_TARIH))))
    If n > 0 Then
        iOnce = KPI_ONCEKI_SIRA(d, n, bugun, 1)
        iOtuz = KPI_ONCEKI_SIRA(d, n, bugun, 28)
        ilk = Format$(CDate(d(1, K_TARIH)), "dd.mm.yyyy")
    End If
    If iOnce > 0 Then tOnce = Format$(CDate(d(iOnce, K_TARIH)), "dd.mm.yyyy")
    If iOtuz > 0 Then tOtuz = Format$(CDate(d(iOtuz, K_TARIH)), "dd.mm.yyyy")
    mOzOnceki = tOnce

    ADM_BOLUM ws, r0, ADM_ADRES(r0, 2, r0, 13), "YÖNETİM KARNESİ  •  TREND   •   " & ilMetin & _
        IIf(Len(tOnce) > 0, "   (önceki kayıt: " & tOnce & ")", "   (ilk kayıt bugün alındı; karşılaştırma yarından itibaren)")
    r = r0 + 1
    ADM_BASLIKLAR ws, r, 2, Array("Gösterge", "", "Bugün", "Önceki kayıt" & IIf(Len(tOnce) > 0, vbLf & tOnce, ""), "Değişim", "", _
        "30 gün önce" & IIf(Len(tOtuz) > 0, vbLf & tOtuz, ""), "30 günlük değişim", "", "Eğilim", "Değerlendirme", "")
    ws.Range(ADM_ADRES(r, 2, r, 3)).Merge: ws.Range(ADM_ADRES(r, 6, r, 7)).Merge
    ws.Range(ADM_ADRES(r, 9, r, 10)).Merge: ws.Range(ADM_ADRES(r, 12, r, 13)).Merge
    ws.Rows(r).RowHeight = 30

    For k = 1 To KPI_ADET
        r = r0 + 1 + k
        t = KPI_TANIM(k)
        simdi = KPI_DEGER(kayit, 1, k)
        once = Empty: otuz = Empty
        If iOnce > 0 Then once = KPI_DEGER(d, iOnce, k)
        If iOtuz > 0 Then otuz = KPI_DEGER(d, iOtuz, k)
        ws.Rows(r).RowHeight = 22
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
        ADM_BLOK ws, ADM_ADRES(r, 2, r, 3), CStr(t(0)), vbWhite, RGB(30, 41, 59), 10, True, xlLeft
        ws.Cells(r, 2).ShrinkToFit = True
        ADM_BLOK ws, ADM_ADRES(r, 4, r, 4), KPI_METIN(simdi, CLng(t(1))), RGB(226, 236, 247), RGB(12, 45, 79), 11, True, xlCenter
        ADM_BLOK ws, ADM_ADRES(r, 5, r, 5), IIf(iOnce > 0, KPI_METIN(once, CLng(t(1))), "-"), vbWhite, RGB(71, 85, 105), 10, False, xlCenter
        yon = KPI_YON(simdi, once, CLng(t(2)))
        ADM_BLOK ws, ADM_ADRES(r, 6, r, 7), KPI_FARK(simdi, once, CLng(t(1))), vbWhite, KPI_YON_RENK(yon), 10, True, xlCenter
        ADM_BLOK ws, ADM_ADRES(r, 8, r, 8), IIf(iOtuz > 0, KPI_METIN(otuz, CLng(t(1))), "-"), vbWhite, RGB(71, 85, 105), 10, False, xlCenter
        ADM_BLOK ws, ADM_ADRES(r, 9, r, 10), KPI_FARK(simdi, otuz, CLng(t(1))), vbWhite, _
            KPI_YON_RENK(KPI_YON(simdi, otuz, CLng(t(2)))), 10, True, xlCenter
        ws.Cells(r, 4).ShrinkToFit = True: ws.Cells(r, 5).ShrinkToFit = True: ws.Cells(r, 8).ShrinkToFit = True
        ws.Cells(r, 6).ShrinkToFit = True: ws.Cells(r, 9).ShrinkToFit = True
        yorum = KPI_YORUM(simdi, once, CLng(t(2)))
        Select Case yon
            Case 1: dolgu = RGB(220, 252, 231): yazi = RGB(21, 128, 61)
            Case -1: dolgu = RGB(254, 226, 226): yazi = RGB(185, 28, 28)
            Case Else: dolgu = RGB(241, 245, 249): yazi = RGB(100, 116, 139)
        End Select
        ADM_BLOK ws, ADM_ADRES(r, 12, r, 13), yorum, dolgu, yazi, 9, True, xlCenter
        ws.Cells(r, 11).Interior.Color = vbWhite
        KPI_EGILIM ws, r, k, d, n

        mOzTablo(k, 1) = CStr(t(0))
        mOzTablo(k, 2) = KPI_METIN(simdi, CLng(t(1)))
        mOzTablo(k, 3) = IIf(iOnce > 0, KPI_METIN(once, CLng(t(1))), "-")
        mOzTablo(k, 4) = KPI_FARK(simdi, once, CLng(t(1)))
        mOzTablo(k, 5) = IIf(iOtuz > 0, KPI_METIN(otuz, CLng(t(1))), "-")
        mOzTablo(k, 6) = KPI_FARK(simdi, otuz, CLng(t(1)))
        mOzTablo(k, 7) = yorum
        mOzYon(k) = yon
        mOzHam(k, 1) = simdi: mOzHam(k, 2) = once: mOzHam(k, 3) = otuz
    Next k
    ADM_KENAR ws.Range(ADM_ADRES(r0 + 1, 2, r0 + 1 + KPI_ADET, 13)), True

    r = r0 + KPI_ADET + 2
    ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Her rapor yenilemesinde günde bir kayıt tutulur (" & n & " kayıt" & _
        IIf(Len(ilk) > 0, ", ilk kayıt " & ilk, "") & "). Yeşil: iyileşme, kırmızı: kötüleşme. " & _
        "Bekleyen tasarruf fırsatı = uygulanmamış transfer önerileri + iptal edilebilir talep / siparişler. " & _
        "Tek sayfalık özet için üstteki YÖNETİCİ ÖZETİ (PDF) düğmesini kullanın.", RGB(244, 248, 252), RGB(100, 116, 139), 8, False, xlLeft
    ws.Cells(r, 2).Font.Italic = True
End Sub

' Egilim cizgisi (son 12 kayit). Veri, kayit sayfasinin AA:AL sutunlarina yazilir; cizgi oradan beslenir.
Private Sub KPI_EGILIM(ByVal ws As Worksheet, ByVal r As Long, ByVal k As Long, ByRef d As Variant, ByVal n As Long)
    Dim kws As Worksheet, ilk As Long, i As Long, j As Long, a As Variant, dolu As Long, grp As Object, alan As Range
    On Error Resume Next
    Set kws = ThisWorkbook.Worksheets(KPI_SAYFA)
    If kws Is Nothing Or n < 2 Then
        ws.Cells(r, 11).NumberFormat = "@": ws.Cells(r, 11).Value2 = "-"
        ws.Cells(r, 11).HorizontalAlignment = xlCenter: ws.Cells(r, 11).Font.Color = RGB(148, 163, 184)
        Exit Sub
    End If
    Set alan = kws.Range(kws.Cells(k, KPI_SERI_SUTUN), kws.Cells(k, KPI_SERI_SUTUN + KPI_SERI - 1))
    alan.ClearContents
    ilk = n - KPI_SERI + 1
    If ilk < 1 Then ilk = 1
    For i = ilk To n
        a = KPI_DEGER(d, i, k)
        If Not IsEmpty(a) Then
            kws.Cells(k, KPI_SERI_SUTUN + j).Value2 = a
            dolu = dolu + 1
        End If
        j = j + 1
    Next i
    If dolu < 2 Then
        ws.Cells(r, 11).NumberFormat = "@": ws.Cells(r, 11).Value2 = "-"
        ws.Cells(r, 11).HorizontalAlignment = xlCenter: ws.Cells(r, 11).Font.Color = RGB(148, 163, 184)
        Exit Sub
    End If
    ws.Cells(r, 11).NumberFormat = "General"
    ' Type 1 = cizgi (xlSparkLine)
    Set grp = ws.Cells(r, 11).SparklineGroups.Add(1, "'" & KPI_SAYFA & "'!" & alan.Address)
    If Not grp Is Nothing Then
        grp.SeriesColor.Color = RGB(20, 103, 173)
        grp.LineWeight = 1.75
        grp.Points.Markers.Visible = True
        grp.Points.Markers.Color.Color = RGB(20, 103, 173)
        grp.Points.Lastpoint.Visible = True
        grp.Points.Lastpoint.Color.Color = RGB(185, 28, 28)
    End If
End Sub

' Ust bolumde (3. satir seridinin saginda) dort dugme: web panosu, PDF ozeti, islem gunlugu, saglik kontrolu.
Private Sub KPI_DUGMELER(ByVal ws As Worksheet)
    Dim x As Double, y As Double, h As Double, w As Double
    On Error Resume Next
    y = ws.Rows(3).Top + 2: h = ws.Rows(3).Height - 4
    w = ws.Columns("B").Width
    x = ws.Range("H1").Left + 20
    KPI_DUGME ws, "ADM_KPI_PANO", "YÖNETİCİ PANOSU (WEB)", x, y, w * 1.5, h, RGB(0, 125, 110), "AKSORA_PANO_AC"
    KPI_DUGME ws, "ADM_KPI_PDF", "YÖNETİCİ ÖZETİ (PDF)", x + w * 1.5 + 6, y, w * 1.4, h, RGB(185, 28, 28), "AKSORA_KPI_OZET_PDF"
    KPI_DUGME ws, "ADM_KPI_GUNLUK", "İŞLEM GÜNLÜĞÜ", x + w * 2.9 + 12, y, w, h, RGB(51, 65, 85), "AKSORA_GUNLUK_EXCEL"
    KPI_DUGME ws, "ADM_KPI_SAGLIK", "SAĞLIK KONTROLÜ", x + w * 3.9 + 18, y, w * 1.1, h, RGB(51, 65, 85), "AKSORA_SAGLIK_KONTROL"
End Sub

Private Sub KPI_DUGME(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, ByVal x As Double, ByVal y As Double, _
    ByVal w As Double, ByVal h As Double, ByVal renk As Long, ByVal makro As String)
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(ad).Delete
    Err.Clear
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    If s Is Nothing Then Exit Sub
    With s
        .Name = ad
        .Fill.ForeColor.RGB = renk
        .Line.ForeColor.RGB = RGB(116, 192, 226): .Line.Weight = 0.75
        .Adjustments.Item(1) = 0.25
        .Placement = xlFreeFloating
        .OnAction = BIRIM_MAKRO(makro)
        With .TextFrame2
            .MarginLeft = 3: .MarginRight = 3: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
            .TextRange.Text = metin
            .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 8: .TextRange.Font.Bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub

' "BUGUN YAPILACAK ISLER" tablosu (karnenin hemen ustunde) PDF ozeti icin okunur.
' Donus: Collection of Array(is metni, tutar ya da Empty, nerede)
Private Function KPI_ISLERI_OKU(ByVal ws As Worksheet) As Collection
    Dim c As Collection, r As Long, son As Long, t As String, v As Variant
    Set c = New Collection
    Set KPI_ISLERI_OKU = c
    On Error GoTo Bitti
    son = mYerSatir
    If son <= 0 Then son = 600
    For r = son To 1 Step -1
        v = ws.Cells(r, 2).Value2
        If Not IsError(v) Then
            t = CStr(v)
            If Left$(t, 15) = "BUGÜN YAPILACAK" Then Exit For
        End If
    Next r
    If r < 1 Then Exit Function
    r = r + 2
    Do While r < son
        v = ws.Cells(r, 2).Value2
        If Not KPI_SAYI_MI(v) Then Exit Do
        c.Add Array(CStr(ws.Cells(r, 3).Value2), ws.Cells(r, 10).Value2, CStr(ws.Cells(r, 12).Value2))
        r = r + 1
    Loop
Bitti:
End Function

'==========================================================
' YONETICI OZETI (tek sayfa PDF, masaustune)
'==========================================================
Public Sub AKSORA_KPI_OZET_PDF()
    Dim wb As Workbook, c As Worksheet, r As Long, i As Long, k As Long, ev As Boolean, yol As String
    Dim no As Long, metin As String, kart As Boolean, x As Variant, iller As Variant, toplam As Double, ilk As Long
    Dim parca() As String, t As String
    If Not mOzVar Then
        MsgBox "Özet için önce RAPORU YENİLE düğmesine basın.", vbInformation, "AKSORA | YÖNETİCİ ÖZETİ"
        Exit Sub
    End If
    kart = AKSORA_BEKLE_AC("AKSORA  |  YÖNETİCİ ÖZETİ", "PDF hazırlanıyor...", 0.3)
    ev = Application.EnableEvents
    On Error GoTo Hata
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set c = wb.Worksheets(1)
    c.Cells.Font.Name = "Aptos": c.Cells.Font.Size = 10: c.Cells.Font.Color = RGB(30, 41, 59)
    c.Cells.VerticalAlignment = xlCenter
    c.Columns("A").ColumnWidth = 30: c.Columns("B").ColumnWidth = 17: c.Columns("C").ColumnWidth = 17
    c.Columns("D").ColumnWidth = 26: c.Columns("E").ColumnWidth = 17: c.Columns("F").ColumnWidth = 26
    c.Columns("G").ColumnWidth = 18

    c.Rows(1).RowHeight = 36: c.Rows(2).RowHeight = 20
    ADM_BLOK c, "A1:G1", "AKSORA  |  YÖNETİCİ ÖZETİ", RGB(12, 45, 79), vbWhite, 20, True, xlLeft
    ADM_BLOK c, "A2:G2", "Çoruh EDAŞ  •  " & mOzIl & "  •  " & Format$(Now, "dd.mm.yyyy hh:nn") & _
        "  •  Hazırlayan: " & AKSORA_KULLANICI(), RGB(23, 66, 107), RGB(198, 240, 97), 10, True, xlLeft

    ' 1) Gostergeler
    r = 4
    ADM_BOLUM c, r, ADM_ADRES(r, 1, r, 7), "1.  GÖSTERGELER" & IIf(Len(mOzOnceki) > 0, "   (önceki kayıt: " & mOzOnceki & ")", "")
    r = r + 1: ilk = r
    ADM_BASLIKLAR c, r, 1, Array("Gösterge", "Bugün", "Önceki kayıt", "Değişim", "30 gün önce", "30 günlük değişim", "Değerlendirme")
    For k = 1 To KPI_ADET
        r = r + 1
        c.Rows(r).RowHeight = 20
        For i = 1 To 7
            c.Cells(r, i).NumberFormat = "@"
            c.Cells(r, i).Value2 = mOzTablo(k, i)
            c.Cells(r, i).HorizontalAlignment = IIf(i = 1, xlLeft, xlCenter)
        Next i
        c.Cells(r, 1).Font.Bold = True: c.Cells(r, 1).IndentLevel = 1
        c.Cells(r, 2).Font.Bold = True: c.Cells(r, 2).Interior.Color = RGB(226, 236, 247)
        c.Cells(r, 4).Font.Color = KPI_YON_RENK(mOzYon(k)): c.Cells(r, 4).Font.Bold = True
        Select Case mOzYon(k)
            Case 1: c.Cells(r, 7).Interior.Color = RGB(220, 252, 231): c.Cells(r, 7).Font.Color = RGB(21, 128, 61)
            Case -1: c.Cells(r, 7).Interior.Color = RGB(254, 226, 226): c.Cells(r, 7).Font.Color = RGB(185, 28, 28)
            Case Else: c.Cells(r, 7).Interior.Color = RGB(241, 245, 249): c.Cells(r, 7).Font.Color = RGB(100, 116, 139)
        End Select
        c.Cells(r, 7).Font.Bold = True
    Next k
    ADM_KENAR c.Range(ADM_ADRES(ilk, 1, r, 7)), True

    ' 2) Bugun yapilacak isler
    r = r + 2
    ADM_BOLUM c, r, ADM_ADRES(r, 1, r, 7), "2.  BUGÜN YAPILACAK İŞLER   (en önemliden başlayarak)"
    r = r + 1: ilk = r
    ADM_BASLIKLAR c, r, 1, Array("Yapılacak iş", "", "", "", "", "Tutar", "Nerede")
    c.Range(ADM_ADRES(r, 1, r, 5)).Merge
    i = 0
    If Not mOzIsler Is Nothing Then
        For Each x In mOzIsler
            r = r + 1: i = i + 1
            c.Range(ADM_ADRES(r, 1, r, 5)).Merge
            c.Cells(r, 1).NumberFormat = "@": c.Cells(r, 1).Value2 = i & ")  " & CStr(x(0))
            c.Cells(r, 1).WrapText = True: c.Cells(r, 1).HorizontalAlignment = xlLeft: c.Cells(r, 1).IndentLevel = 1
            c.Rows(r).RowHeight = 16 * (1 + (Len(CStr(x(0))) \ 100))
            If KPI_SAYI_MI(x(1)) Then
                c.Cells(r, 6).Value2 = CDbl(x(1)): c.Cells(r, 6).NumberFormat = "#,##0.00 ""TL"""
                c.Cells(r, 6).HorizontalAlignment = xlRight: c.Cells(r, 6).Font.Bold = True
            End If
            c.Cells(r, 7).NumberFormat = "@": c.Cells(r, 7).Value2 = CStr(x(2))
            c.Cells(r, 7).HorizontalAlignment = xlCenter: c.Cells(r, 7).ShrinkToFit = True
            If i = 1 Then c.Range(ADM_ADRES(r, 1, r, 7)).Interior.Color = RGB(254, 242, 242): c.Cells(r, 1).Font.Bold = True
        Next x
    End If
    If i = 0 Then
        r = r + 1
        ADM_BLOK c, ADM_ADRES(r, 1, r, 7), "Bekleyen önemli iş yok.", RGB(220, 252, 231), RGB(21, 128, 61), 10, True, xlLeft
    End If
    ADM_KENAR c.Range(ADM_ADRES(ilk, 1, r, 7)), True

    ' 3) Il karsilastirmasi
    r = r + 2
    ADM_BOLUM c, r, ADM_ADRES(r, 1, r, 7), "3.  İL KARŞILAŞTIRMASI   (tekil stok değeri: ortak depolar bir kez sayılır)"
    r = r + 1: ilk = r
    ADM_BASLIKLAR c, r, 1, Array("İl", "Tekil stok değeri", "Pay", "Atıl stok değeri", "Atıl kalem", "Atıl oranı", "")
    iller = ADM_ILLER()
    For i = 0 To 4
        toplam = toplam + mOzIller(i, 0)
    Next i
    For i = 0 To 4
        r = r + 1
        c.Cells(r, 1).Value2 = Y_IL_ADI(CStr(iller(i))): c.Cells(r, 1).Font.Bold = True: c.Cells(r, 1).IndentLevel = 1
        c.Cells(r, 2).Value2 = Round(mOzIller(i, 0), 2): c.Cells(r, 2).NumberFormat = "#,##0 ""TL"""
        If toplam > 0 Then c.Cells(r, 3).Value2 = mOzIller(i, 0) / toplam
        c.Cells(r, 3).NumberFormat = "0.0%"
        If mOzAtilVar Then
            c.Cells(r, 4).Value2 = Round(mOzIller(i, 2), 2): c.Cells(r, 4).NumberFormat = "#,##0 ""TL"""
            c.Cells(r, 5).Value2 = mOzIller(i, 1): c.Cells(r, 5).NumberFormat = "#,##0"
            If mOzIller(i, 3) > 0 Then c.Cells(r, 6).Value2 = mOzIller(i, 2) / mOzIller(i, 3)
            c.Cells(r, 6).NumberFormat = "0.0%"
        Else
            c.Cells(r, 4).Value2 = "-": c.Cells(r, 5).Value2 = "-": c.Cells(r, 6).Value2 = "-"
        End If
        c.Range(ADM_ADRES(r, 2, r, 7)).HorizontalAlignment = xlCenter
    Next i
    r = r + 1
    c.Cells(r, 1).Value2 = "TOPLAM": c.Cells(r, 1).IndentLevel = 1
    c.Cells(r, 2).Value2 = Round(toplam, 2): c.Cells(r, 2).NumberFormat = "#,##0 ""TL"""
    c.Range(ADM_ADRES(r, 1, r, 7)).Font.Bold = True: c.Range(ADM_ADRES(r, 1, r, 7)).Interior.Color = RGB(226, 236, 247)
    c.Range(ADM_ADRES(r, 2, r, 7)).HorizontalAlignment = xlCenter
    ADM_KENAR c.Range(ADM_ADRES(ilk, 1, r, 7)), True
    If Not mOzAtilVar Then
        r = r + 1
        ADM_BLOK c, ADM_ADRES(r, 1, r, 7), "Atıl stok sütunları için Malzeme > ATIL STOK ekranında SAP okuması (MB51) yapılmış olmalıdır.", _
            vbWhite, RGB(100, 116, 139), 8, False, xlLeft
    End If

    ' 4) Oncelikli uyarilar
    r = r + 2
    ADM_BOLUM c, r, ADM_ADRES(r, 1, r, 7), "4.  ÖNCELİKLİ UYARILAR"
    ilk = r + 1
    If mOzUyariSay = 0 Then
        r = r + 1
        ADM_BLOK c, ADM_ADRES(r, 1, r, 7), "Kritik veya dikkat gerektiren uyarı yok.", RGB(220, 252, 231), RGB(21, 128, 61), 10, True, xlLeft
    Else
        For i = 1 To mOzUyariSay
            r = r + 1
            parca = Split(mOzUyari(i), vbTab)
            t = parca(1)
            c.Cells(r, 1).NumberFormat = "@": c.Cells(r, 1).Value2 = parca(0)
            c.Cells(r, 1).HorizontalAlignment = xlCenter: c.Cells(r, 1).Font.Bold = True
            If parca(0) = "KRİTİK" Then
                c.Cells(r, 1).Interior.Color = RGB(254, 226, 226): c.Cells(r, 1).Font.Color = RGB(185, 28, 28)
            Else
                c.Cells(r, 1).Interior.Color = RGB(254, 243, 199): c.Cells(r, 1).Font.Color = RGB(146, 64, 14)
            End If
            c.Range(ADM_ADRES(r, 2, r, 7)).Merge
            c.Cells(r, 2).NumberFormat = "@": c.Cells(r, 2).Value2 = t
            c.Cells(r, 2).WrapText = True: c.Cells(r, 2).HorizontalAlignment = xlLeft: c.Cells(r, 2).IndentLevel = 1
            c.Rows(r).RowHeight = 16 * (1 + (Len(t) \ 105))
        Next i
    End If
    ADM_KENAR c.Range(ADM_ADRES(ilk, 1, r, 7)), True

    r = r + 2
    ADM_BLOK c, ADM_ADRES(r, 1, r, 7), "Bu özet AKSORA Yönetim Merkezi verilerinden otomatik üretilmiştir. Değerler rapor anındaki SAP okumalarına dayanır.   " & _
        AKSORA_IMZA_METNI(), vbWhite, RGB(100, 116, 139), 8, False, xlLeft
    c.Cells(r, 1).Font.Italic = True

    On Error Resume Next
    Application.PrintCommunication = False
    With c.PageSetup
        .PrintArea = "$A$1:$G$" & r
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = 1
        .LeftMargin = Application.CentimetersToPoints(1.2): .RightMargin = Application.CentimetersToPoints(1.2)
        .TopMargin = Application.CentimetersToPoints(1.2): .BottomMargin = Application.CentimetersToPoints(1.2)
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True
    On Error GoTo Hata

    yol = AKSORA_MASAUSTU() & Application.PathSeparator & _
        AKSORA_DOSYA_ADI("AKSORA Yönetici Özeti - " & mOzIl & " - " & Format$(Now, "dd.mm.yyyy hh.nn")) & ".pdf"
    c.ExportAsFixedFormat Type:=xlTypePDF, Filename:=yol, Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    wb.Close SaveChanges:=False
    Set wb = Nothing
    ThisWorkbook.Activate
    Application.EnableEvents = ev: Application.ScreenUpdating = True
    If kart Then AKSORA_BEKLE_KAPAT
    AKSORA_GUNLUK_YAZ "YÖNETİCİ", "Yönetici özeti PDF: " & mOzIl
    Application.StatusBar = "Masaüstüne kaydedildi: " & yol
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.EnableEvents = ev: Application.ScreenUpdating = True
    If kart Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    MsgBox "Yönetici özeti oluşturulamadı." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

'==========================================================
' YONETICI PANOSU (WEB) icin veri: son cizilen karne + gecmis, JSON metni olarak.
' Rapor henuz kurulmadiysa bos metin doner. (Sayfayi modAKSORA_PANO yazar.)
'==========================================================
Public Function AKSORA_KPI_PANO_JSON() As String
    Dim s As String, k As Long, t As Variant, g As Variant, n As Long, d As Variant, i As Long, ilk As Long
    Dim iller As Variant, x As Variant, parca() As String, adlar As Variant
    If Not mOzVar Then Exit Function
    g = KPI_GECMIS(mOzKapsam)
    n = g(0)
    If n > 0 Then d = g(1)
    ilk = 1
    If n > 120 Then ilk = n - 119
    s = "{""ornek"":false,""il"":" & AKSORA_PANO_METIN(mOzIl) & ",""tarih"":" & AKSORA_PANO_METIN(Format$(Now, "dd.mm.yyyy hh:nn")) & _
        ",""hazirlayan"":" & AKSORA_PANO_METIN(AKSORA_KULLANICI()) & ",""onceki"":" & AKSORA_PANO_METIN(mOzOnceki) & ",""seriTarih"":["
    For i = ilk To n
        s = s & IIf(i > ilk, ",", "") & AKSORA_PANO_METIN(Format$(CDate(d(i, K_TARIH)), "dd.mm.yyyy"))
    Next i
    s = s & "],""kpi"":["
    For k = 1 To KPI_ADET
        t = KPI_TANIM(k)
        s = s & IIf(k > 1, ",", "") & "{""ad"":" & AKSORA_PANO_METIN(CStr(t(0))) & ",""tur"":" & CStr(CLng(t(1))) & ",""yon"":" & CStr(CLng(t(2))) & _
            ",""bugun"":" & AKSORA_PANO_SAYI(mOzHam(k, 1)) & ",""onceki"":" & AKSORA_PANO_SAYI(mOzHam(k, 2)) & _
            ",""otuz"":" & AKSORA_PANO_SAYI(mOzHam(k, 3)) & ",""seri"":["
        For i = ilk To n
            s = s & IIf(i > ilk, ",", "") & AKSORA_PANO_SAYI(KPI_DEGER(d, i, k))
        Next i
        s = s & "]}"
    Next k
    s = s & "],""atilVar"":" & IIf(mOzAtilVar, "true", "false") & ",""iller"":["
    iller = ADM_ILLER()
    For i = 0 To 4
        s = s & IIf(i > 0, ",", "") & "{""ad"":" & AKSORA_PANO_METIN(Y_IL_ADI(CStr(iller(i)))) & ",""tekil"":" & AKSORA_PANO_SAYI(mOzIller(i, 0)) & _
            ",""atil"":" & AKSORA_PANO_SAYI(mOzIller(i, 2)) & ",""kapsam"":" & AKSORA_PANO_SAYI(mOzIller(i, 3)) & "}"
    Next i
    s = s & "],""birimler"":["
    adlar = Array("Malzeme Stok", "Yatırım", "Müşteri", "İlçe Yönetim")
    For i = 0 To 3
        s = s & IIf(i > 0, ",", "") & "{""ad"":" & AKSORA_PANO_METIN(CStr(adlar(i))) & ",""deger"":" & AKSORA_PANO_SAYI(mOzBirim(i)) & "}"
    Next i
    s = s & "],""isler"":["
    i = 0
    If Not mOzIsler Is Nothing Then
        For Each x In mOzIsler
            s = s & IIf(i > 0, ",", "") & "{""is"":" & AKSORA_PANO_METIN(CStr(x(0))) & ",""tutar"":" & AKSORA_PANO_SAYI(x(1)) & _
                ",""nerede"":" & AKSORA_PANO_METIN(CStr(x(2))) & "}"
            i = i + 1
        Next x
    End If
    s = s & "],""uyarilar"":["
    For i = 1 To mOzUyariSay
        parca = Split(mOzUyari(i), vbTab)
        s = s & IIf(i > 1, ",", "") & "{""sev"":" & AKSORA_PANO_METIN(parca(0)) & ",""metin"":" & AKSORA_PANO_METIN(parca(1)) & "}"
    Next i
    AKSORA_KPI_PANO_JSON = s & "]}"
End Function

'==========================================================
' SINAMA: saf hesap islevleri (kurulum betigi ve AKSORA_GELISTIRME_SINAMA cagirir).
' Donus: "OK" ya da ilk basarisiz kontrolun adi.
'==========================================================
Public Function AKSORA_KPI_SINAMA() As String
    Dim v(1 To 2, 1 To 24) As Variant, g As Variant, bos As Variant
    On Error GoTo Hata
    v(1, K_TARIH) = 46000#: v(1, K_SKOR) = 70: v(1, K_TEKIL) = 1000000#: v(1, K_TR_TL) = 100#
    v(2, K_TARIH) = 46030#: v(2, K_SKOR) = 82: v(2, K_TEKIL) = 900000#: v(2, K_TR_TL) = 250#: v(2, K_IP_TL) = 50#
    g = v
    If KPI_DEGER(g, 1, 1) <> 70 Then AKSORA_KPI_SINAMA = "deger-skor": Exit Function
    If KPI_DEGER(g, 2, 6) <> 300 Then AKSORA_KPI_SINAMA = "deger-tasarruf": Exit Function
    If KPI_DEGER(g, 1, 6) <> 100 Then AKSORA_KPI_SINAMA = "deger-tasarruf-tek": Exit Function
    If Not IsEmpty(KPI_DEGER(g, 1, 3)) Then AKSORA_KPI_SINAMA = "deger-bos": Exit Function
    If KPI_YON(82, 70, 1) <> 1 Then AKSORA_KPI_SINAMA = "yon-iyi": Exit Function
    If KPI_YON(900000#, 1000000#, -1) <> 1 Then AKSORA_KPI_SINAMA = "yon-azalan-iyi": Exit Function
    If KPI_YON(5, 3, -1) <> -1 Then AKSORA_KPI_SINAMA = "yon-kotu": Exit Function
    If KPI_YON(5, 3, 0) <> 0 Then AKSORA_KPI_SINAMA = "yon-notr": Exit Function
    If KPI_YON(5, bos, 1) <> 0 Then AKSORA_KPI_SINAMA = "yon-bos": Exit Function
    If KPI_YORUM(82, 70, 1) <> "İYİLEŞTİ" Then AKSORA_KPI_SINAMA = "yorum-iyi": Exit Function
    If KPI_YORUM(5, 3, -1) <> "KÖTÜLEŞTİ" Then AKSORA_KPI_SINAMA = "yorum-kotu": Exit Function
    If KPI_YORUM(5, 5, -1) <> "AYNI" Then AKSORA_KPI_SINAMA = "yorum-ayni": Exit Function
    If KPI_YORUM(5, bos, -1) <> "İLK KAYIT" Then AKSORA_KPI_SINAMA = "yorum-ilk": Exit Function
    If KPI_YORUM(5, 3, 0) <> "ARTTI" Then AKSORA_KPI_SINAMA = "yorum-notr": Exit Function
    If KPI_FARK(5, bos, 4) <> "-" Then AKSORA_KPI_SINAMA = "fark-bos": Exit Function
    If KPI_FARK(5, 5, 4) <> "değişmedi" Then AKSORA_KPI_SINAMA = "fark-ayni": Exit Function
    If Left$(KPI_FARK(82, 70, 1), 1) <> ChrW(9650) Then AKSORA_KPI_SINAMA = "fark-ok": Exit Function
    If InStr(KPI_FARK(82, 70, 1), "12 puan") = 0 Then AKSORA_KPI_SINAMA = "fark-puan": Exit Function
    If KPI_ONCEKI_SIRA(g, 2, 46030, 1) <> 1 Then AKSORA_KPI_SINAMA = "onceki-sira": Exit Function
    If KPI_ONCEKI_SIRA(g, 2, 46030, 28) <> 1 Then AKSORA_KPI_SINAMA = "otuz-sira": Exit Function
    If KPI_ONCEKI_SIRA(g, 2, 46030, 31) <> 0 Then AKSORA_KPI_SINAMA = "otuz-yok": Exit Function
    If KPI_SAYI_MI("12") Or KPI_SAYI_MI(bos) Or Not KPI_SAYI_MI(12.5) Then AKSORA_KPI_SINAMA = "sayi-mi": Exit Function
    If CStr(KPI_TANIM(4)(0)) <> "Atıl stok oranı" Then AKSORA_KPI_SINAMA = "tanim": Exit Function
    AKSORA_KPI_SINAMA = "OK"
    Exit Function
Hata:
    AKSORA_KPI_SINAMA = "beklenmeyen hata: " & Err.Description
End Function
