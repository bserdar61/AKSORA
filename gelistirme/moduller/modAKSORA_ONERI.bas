Attribute VB_Name = "modAKSORA_ONERI"
Option Explicit

'==========================================================
' AKSORA v3.1  |  modAKSORA_ONERI
'
' ONERI TAKIBI: AKSORA'nin verdigi transfer / atil kullanim onerilerinin akibeti.
'
' Bugune kadar sistem "286 oneri, 7,6 mn TL kazanc" diyordu ama
'   - bir oneri kac gundur bekliyor?
'   - kac oneri kapandi, toplam kac TL?
' sorularini cevaplayamiyordu. Bu modul her yonetim raporunda (Tum iller) oneri
' listesini cok gizli "_AKSORA_ONERI" sayfasindaki kayitla eslestirir:
'   - listede olan oneri  : ACIK kalir, "son gorulme" guncellenir
'   - listeden cikan oneri: KAPANDI olarak isaretlenir (kapanis tarihi ve son kazanc tutariyla)
'   - yeni oneri          : ACIK olarak eklenir ("ilk gorulme" = bugun)
'
' "Kapandi" = oneri artik listede yok: uygulanmis ya da ihtiyac ortadan kalkmistir.
' Kesin "uygulandi" bilgisi icin SAP hareketleriyle eslestirme gerekir (sonraki adim).
'
' Baglanti noktasi (kurulum betigi ekler):
'   ADM_ANALIZ_UYARI : AKSORA_ONERI_KAYDET liste, secIl
'==========================================================

Private Const ONR_SAYFA As String = "_AKSORA_ONERI"
Private Const ONR_SUTUN As Long = 14
Private Const ONR_AZAMI As Long = 6000
Private Const ONR_SAKLA_GUN As Long = 400      ' kapanan oneriler bu kadar gun saklanir

' Tablo sutunlari
Private Const O_ANAHTAR As Long = 1, O_NEREDEN As Long = 2, O_NEREYE As Long = 3, O_SISTEM As Long = 4, O_MALZEME As Long = 5
Private Const O_TANIM As Long = 6, O_MIKTAR As Long = 7, O_BIRIM As Long = 8, O_TL As Long = 9, O_ILK As Long = 10
Private Const O_SON As Long = 11, O_DURUM As Long = 12, O_KAPANIS As Long = 13, O_YAPILACAK As Long = 14

Private Const D_ACIK As String = "AÇIK"
Private Const D_KAPANDI As String = "KAPANDI"

'==========================================================
' SAF ISLEVLER (sayfaya dokunmaz; AKSORA_ONERI_SINAMA ile dogrulanir)
'==========================================================

' Anahtar -> sira esleme (Collection; anahtar yoksa 0).
Private Function ONR_BUL(ByVal c As Collection, ByVal anahtar As String) As Long
    On Error Resume Next
    ONR_BUL = c(anahtar)
End Function

Private Function ONR_GUN(ByVal v As Variant) As Long
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If IsNumeric(v) Then ONR_GUN = CLng(Int(CDbl(v)))
End Function

Private Function ONR_SAYI(ByVal v As Variant) As Double
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If IsNumeric(v) Then ONR_SAYI = CDbl(v)
End Function

' Kayitli tablo (eski) ile bugunun oneri listesini (yeni) eslestirir.
'   eski : (1..nEski, 1..14) ya da Empty        yeni : (1..nYeni, 1..14) - yalniz 1..9 ve 14. sutunlar dolu
'   bugun: gun numarasi (CLng(Date))
' Donus: yeni tablo (1..nSonuc, 1..14); nSonuc = 0 ise Empty.
Public Function AKSORA_ONERI_ESLE(ByVal eski As Variant, ByVal nEski As Long, ByVal yeni As Variant, ByVal nYeni As Long, _
    ByVal bugun As Long, ByRef nSonuc As Long) As Variant
    Dim acik As Collection, gorulen() As Boolean, s() As Variant, i As Long, j As Long, k As Long, n As Long, tut As Boolean
    Set acik = New Collection
    nSonuc = 0
    If nEski + nYeni = 0 Then Exit Function
    ReDim s(1 To nEski + nYeni, 1 To ONR_SUTUN)
    If nYeni > 0 Then ReDim gorulen(1 To nYeni)
    ' Bugunun listesi: anahtar -> sira
    For j = 1 To nYeni
        If ONR_BUL(acik, CStr(yeni(j, O_ANAHTAR))) = 0 Then acik.Add j, CStr(yeni(j, O_ANAHTAR))
    Next j
    ' 1) Eski kayitlar
    For i = 1 To nEski
        tut = True
        If CStr(eski(i, O_DURUM)) = D_ACIK Then
            j = ONR_BUL(acik, CStr(eski(i, O_ANAHTAR)))
            If j > 0 Then
                If gorulen(j) Then j = 0     ' ayni anahtarli ikinci acik kayit (olmamali): kapatilir
            End If
        Else
            j = -1                           ' kapali kayit: oldugu gibi tasinir
            If bugun - ONR_GUN(eski(i, O_KAPANIS)) > ONR_SAKLA_GUN Then tut = False
        End If
        If tut Then
            n = n + 1
            For k = 1 To ONR_SUTUN
                s(n, k) = eski(i, k)
            Next k
            If j > 0 Then
                ' hala listede: guncel miktar / tutar, son gorulme bugun
                gorulen(j) = True
                s(n, O_MIKTAR) = yeni(j, O_MIKTAR): s(n, O_TL) = yeni(j, O_TL)
                s(n, O_TANIM) = yeni(j, O_TANIM): s(n, O_BIRIM) = yeni(j, O_BIRIM): s(n, O_YAPILACAK) = yeni(j, O_YAPILACAK)
                s(n, O_SON) = bugun
            ElseIf j = 0 Then
                ' listeden cikti: kapandi
                s(n, O_DURUM) = D_KAPANDI
                s(n, O_KAPANIS) = bugun
            End If
        End If
    Next i
    ' 2) Yeni oneriler
    For j = 1 To nYeni
        If Not gorulen(j) Then
            If ONR_BUL(acik, CStr(yeni(j, O_ANAHTAR))) = j Then
                n = n + 1
                For k = 1 To 9
                    s(n, k) = yeni(j, k)
                Next k
                s(n, O_ILK) = bugun: s(n, O_SON) = bugun: s(n, O_DURUM) = D_ACIK
                s(n, O_YAPILACAK) = yeni(j, O_YAPILACAK)
            End If
        End If
    Next j
    nSonuc = n
    If n > 0 Then AKSORA_ONERI_ESLE = s
End Function

' Tablo ozeti. Donus: Array(0 acik adet, 1 acik TL, 2 yedi gunu gecen adet, 3 yedi gunu gecen TL,
'   4 son 30 gunde kapanan adet, 5 son 30 gunde kapanan TL, 6 toplam kapanan adet, 7 toplam kapanan TL, 8 ilk kayit gunu)
Public Function AKSORA_ONERI_OZET_HESAP(ByVal t As Variant, ByVal n As Long, ByVal bugun As Long) As Variant
    Dim o(0 To 8) As Double, i As Long, g As Long
    For i = 1 To n
        g = ONR_GUN(t(i, O_ILK))
        If g > 0 And (o(8) = 0 Or g < o(8)) Then o(8) = g
        If CStr(t(i, O_DURUM)) = D_ACIK Then
            o(0) = o(0) + 1: o(1) = o(1) + ONR_SAYI(t(i, O_TL))
            If g > 0 And bugun - g >= 7 Then o(2) = o(2) + 1: o(3) = o(3) + ONR_SAYI(t(i, O_TL))
        Else
            o(6) = o(6) + 1: o(7) = o(7) + ONR_SAYI(t(i, O_TL))
            If bugun - ONR_GUN(t(i, O_KAPANIS)) <= 30 Then o(4) = o(4) + 1: o(5) = o(5) + ONR_SAYI(t(i, O_TL))
        End If
    Next i
    AKSORA_ONERI_OZET_HESAP = o
End Function

' En uzun bekleyen acik oneriler (once bekleme gunu, esitse tutar). Donus: sira numaralari (1..adet); yoksa adet = 0.
Public Function AKSORA_ONERI_BEKLEYENLER(ByVal t As Variant, ByVal n As Long, ByVal istenen As Long, ByRef adet As Long) As Variant
    Dim sec() As Long, alindi() As Boolean, i As Long, k As Long, en As Long
    adet = 0
    If n = 0 Or istenen <= 0 Then Exit Function
    ReDim sec(1 To istenen): ReDim alindi(1 To n)
    For k = 1 To istenen
        en = 0
        For i = 1 To n
            If Not alindi(i) Then
                If CStr(t(i, O_DURUM)) = D_ACIK Then
                    If en = 0 Then
                        en = i
                    ElseIf ONR_GUN(t(i, O_ILK)) < ONR_GUN(t(en, O_ILK)) Then
                        en = i
                    ElseIf ONR_GUN(t(i, O_ILK)) = ONR_GUN(t(en, O_ILK)) And ONR_SAYI(t(i, O_TL)) > ONR_SAYI(t(en, O_TL)) Then
                        en = i
                    End If
                End If
            End If
        Next i
        If en = 0 Then Exit For
        alindi(en) = True
        adet = adet + 1
        sec(adet) = en
    Next k
    If adet > 0 Then AKSORA_ONERI_BEKLEYENLER = sec
End Function

'==========================================================
' KAYIT (sayfa)
'==========================================================
Private Function ONR_BASLIKLAR() As Variant
    ONR_BASLIKLAR = Array("Anahtar", "Nereden", "Nereye", "Sistem", "Malzeme", "Tanım", "Miktar", "Birim", "Kazanç TL", _
        "İlk görülme", "Son görülme", "Durum", "Kapanış", "Ne yapılacak")
End Function

Private Function ONR_TABLO(ByRef n As Long) As Variant
    Dim ws As Worksheet, son As Long
    n = 0
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ONR_SAYFA)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    son = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If son < 2 Then Exit Function
    n = son - 1
    ONR_TABLO = ws.Range(ws.Cells(2, 1), ws.Cells(son, ONR_SUTUN)).Value2
End Function

' Yonetim raporu kurulurken cagrilir. liste: AN_TRANSFERLER sonucu
' (0 nereden, 1 nereye, 2 sistem, 3 malzeme, 4 tanim, 5 miktar, 6 birim, 7 kazanc TL, 8 ne yapilacak, ...).
Public Sub AKSORA_ONERI_KAYDET(ByVal liste As Collection, ByVal secIl As String)
    Dim ws As Worksheet, yeni() As Variant, m As Long, x As Variant, k As String, sira As Collection, j As Long
    Dim eski As Variant, nEski As Long, sonuc As Variant, nSonuc As Long, son As Long, hataMetni As String
    On Error GoTo Hata
    ' Yalniz "Tum iller" raporunda ve ana dosyada: il secili raporda liste eksik gorunmesin.
    If Len(secIl) > 0 Then Exit Sub
    If AKSORA_YAYIN_MODU() Then Exit Sub
    If liste Is Nothing Then Exit Sub
    If liste.Count > 0 Then
        ReDim yeni(1 To liste.Count, 1 To ONR_SUTUN)
        Set sira = New Collection
        For Each x In liste
            k = CStr(x(0)) & "|" & CStr(x(1)) & "|" & CStr(x(2)) & "|" & CStr(x(3))
            j = ONR_BUL(sira, k)
            If j = 0 Then
                m = m + 1: j = m
                sira.Add j, k
                yeni(j, O_ANAHTAR) = k
                yeni(j, O_NEREDEN) = CStr(x(0)): yeni(j, O_NEREYE) = CStr(x(1)): yeni(j, O_SISTEM) = CStr(x(2))
                yeni(j, O_MALZEME) = CStr(x(3)): yeni(j, O_TANIM) = CStr(x(4)): yeni(j, O_BIRIM) = CStr(x(6))
                yeni(j, O_MIKTAR) = 0#: yeni(j, O_TL) = 0#
                yeni(j, O_YAPILACAK) = CStr(x(8))
            End If
            ' Ayni malzeme icin birden cok oneri (talep + kritik stok) tek kayitta toplanir.
            yeni(j, O_MIKTAR) = yeni(j, O_MIKTAR) + ONR_SAYI(x(5))
            yeni(j, O_TL) = Round(yeni(j, O_TL) + ONR_SAYI(x(7)), 2)
        Next x
    End If
    eski = ONR_TABLO(nEski)
    ' Liste bossa (SAP okumasi yok / eski) acik oneriler kapatilmaz: yanlis "kapandi" uretmemek icin beklenir.
    If m = 0 Then Exit Sub
    sonuc = AKSORA_ONERI_ESLE(eski, nEski, yeni, m, CLng(Date), nSonuc)
    If nSonuc = 0 Then Exit Sub
    Set ws = AKSORA_GIZLI_SAYFA(ONR_SAYFA, ONR_BASLIKLAR())
    If ws Is Nothing Then Exit Sub
    son = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If son >= 2 Then ws.Range(ws.Cells(2, 1), ws.Cells(son, ONR_SUTUN)).ClearContents
    If nSonuc > ONR_AZAMI Then nSonuc = ONR_AZAMI
    ws.Range(ws.Cells(2, 1), ws.Cells(nSonuc + 1, ONR_SUTUN)).Value2 = sonuc
    ws.Columns(O_ILK).NumberFormat = "dd.mm.yyyy": ws.Columns(O_SON).NumberFormat = "dd.mm.yyyy"
    ws.Columns(O_KAPANIS).NumberFormat = "dd.mm.yyyy": ws.Columns(O_MALZEME).NumberFormat = "@"
    Exit Sub
Hata:
    hataMetni = Err.Description
    On Error Resume Next
    AKSORA_GUNLUK_YAZ "HATA", "Öneri takibi kaydedilemedi: " & hataMetni
End Sub

' Yonetim Merkezi icin tek satir ozet; kayit yoksa bos metin.
Public Function AKSORA_ONERI_SATIR() As String
    Dim t As Variant, n As Long, o As Variant
    On Error GoTo Cikis
    t = ONR_TABLO(n)
    If n = 0 Then Exit Function
    o = AKSORA_ONERI_OZET_HESAP(t, n, CLng(Date))
    AKSORA_ONERI_SATIR = "ÖNERİ TAKİBİ   •   " & ADM_ADET(o(0)) & " açık öneri (" & ADM_TL_KISA(o(1)) & ")   •   " & _
        ADM_ADET(o(2)) & " tanesi 7 günden uzun bekliyor (" & ADM_TL_KISA(o(3)) & ")   •   son 30 günde " & _
        ADM_ADET(o(4)) & " öneri kapandı (" & ADM_TL_KISA(o(5)) & ")"
Cikis:
End Function

' Web panosu icin JSON; kayit yoksa "null".
Public Function AKSORA_ONERI_JSON() As String
    Dim t As Variant, n As Long, o As Variant, sec As Variant, adet As Long, i As Long, r As Long, s As String, bugun As Long
    AKSORA_ONERI_JSON = "null"
    On Error GoTo Cikis
    t = ONR_TABLO(n)
    If n = 0 Then Exit Function
    bugun = CLng(Date)
    o = AKSORA_ONERI_OZET_HESAP(t, n, bugun)
    s = "{""acik"":" & AKSORA_PANO_SAYI(o(0)) & ",""acikTl"":" & AKSORA_PANO_SAYI(o(1)) & ",""bekleyen7"":" & AKSORA_PANO_SAYI(o(2)) & _
        ",""bekleyen7Tl"":" & AKSORA_PANO_SAYI(o(3)) & ",""kapanan30"":" & AKSORA_PANO_SAYI(o(4)) & ",""kapanan30Tl"":" & AKSORA_PANO_SAYI(o(5)) & _
        ",""kapananToplam"":" & AKSORA_PANO_SAYI(o(6)) & ",""kapananToplamTl"":" & AKSORA_PANO_SAYI(o(7)) & _
        ",""ilk"":" & AKSORA_PANO_METIN(IIf(o(8) > 0, Format$(CDate(o(8)), "dd.mm.yyyy"), "")) & ",""liste"":["
    sec = AKSORA_ONERI_BEKLEYENLER(t, n, 5, adet)
    For i = 1 To adet
        r = sec(i)
        s = s & IIf(i > 1, ",", "") & "{""gun"":" & AKSORA_PANO_SAYI(bugun - ONR_GUN(t(r, O_ILK))) & _
            ",""nereden"":" & AKSORA_PANO_METIN(Y_IL_ADI(CStr(t(r, O_NEREDEN)))) & ",""nereye"":" & AKSORA_PANO_METIN(Y_IL_ADI(CStr(t(r, O_NEREYE)))) & _
            ",""tanim"":" & AKSORA_PANO_METIN(CStr(t(r, O_TANIM))) & ",""miktar"":" & AKSORA_PANO_SAYI(t(r, O_MIKTAR)) & _
            ",""birim"":" & AKSORA_PANO_METIN(CStr(t(r, O_BIRIM))) & ",""tl"":" & AKSORA_PANO_SAYI(t(r, O_TL)) & "}"
    Next i
    AKSORA_ONERI_JSON = s & "]}"
Cikis:
End Function

'==========================================================
' SINAMA: eslestirme ve ozet (kurulum betigi cagirir). Donus: "OK" ya da basarisiz kontrolun adi.
'==========================================================
Public Function AKSORA_ONERI_SINAMA() As String
    Dim y1(1 To 2, 1 To 14) As Variant, y2(1 To 2, 1 To 14) As Variant, t As Variant, n As Long, o As Variant
    Dim bos As Variant, sec As Variant, adet As Long, i As Long, nA As Long, nK As Long
    On Error GoTo Hata
    ' 1. gun (100): A ve B onerileri
    y1(1, O_ANAHTAR) = "A": y1(1, O_TL) = 1000#: y1(1, O_MIKTAR) = 5#
    y1(2, O_ANAHTAR) = "B": y1(2, O_TL) = 400#: y1(2, O_MIKTAR) = 2#
    t = AKSORA_ONERI_ESLE(bos, 0, y1, 2, 100, n)
    If n <> 2 Then AKSORA_ONERI_SINAMA = "ilk-adet": Exit Function
    If CStr(t(1, O_DURUM)) <> D_ACIK Or t(1, O_ILK) <> 100 Then AKSORA_ONERI_SINAMA = "ilk-durum": Exit Function
    ' 10. gun (110): A suruyor (tutar degisti), B yok, C yeni
    y2(1, O_ANAHTAR) = "A": y2(1, O_TL) = 900#: y2(1, O_MIKTAR) = 4#
    y2(2, O_ANAHTAR) = "C": y2(2, O_TL) = 250#: y2(2, O_MIKTAR) = 1#
    t = AKSORA_ONERI_ESLE(t, n, y2, 2, 110, n)
    If n <> 3 Then AKSORA_ONERI_SINAMA = "ikinci-adet": Exit Function
    For i = 1 To n
        Select Case CStr(t(i, O_ANAHTAR))
            Case "A"
                If CStr(t(i, O_DURUM)) <> D_ACIK Or t(i, O_ILK) <> 100 Or t(i, O_SON) <> 110 Or t(i, O_TL) <> 900 Then _
                    AKSORA_ONERI_SINAMA = "suren": Exit Function
            Case "B"
                If CStr(t(i, O_DURUM)) <> D_KAPANDI Or t(i, O_KAPANIS) <> 110 Or t(i, O_TL) <> 400 Then _
                    AKSORA_ONERI_SINAMA = "kapanan": Exit Function
            Case "C"
                If CStr(t(i, O_DURUM)) <> D_ACIK Or t(i, O_ILK) <> 110 Then AKSORA_ONERI_SINAMA = "yeni": Exit Function
        End Select
    Next i
    o = AKSORA_ONERI_OZET_HESAP(t, n, 110)
    If o(0) <> 2 Or o(1) <> 1150 Then AKSORA_ONERI_SINAMA = "ozet-acik": Exit Function
    If o(2) <> 1 Or o(3) <> 900 Then AKSORA_ONERI_SINAMA = "ozet-bekleyen": Exit Function
    If o(4) <> 1 Or o(5) <> 400 Or o(7) <> 400 Then AKSORA_ONERI_SINAMA = "ozet-kapanan": Exit Function
    If o(8) <> 100 Then AKSORA_ONERI_SINAMA = "ozet-ilk": Exit Function
    sec = AKSORA_ONERI_BEKLEYENLER(t, n, 5, adet)
    If adet <> 2 Then AKSORA_ONERI_SINAMA = "bekleyen-adet": Exit Function
    If CStr(t(sec(1), O_ANAHTAR)) <> "A" Then AKSORA_ONERI_SINAMA = "bekleyen-sira": Exit Function
    ' 11. gun: B yeniden onerildi -> kapali kayit durur, yeni acik kayit eklenir
    y2(2, O_ANAHTAR) = "B"
    t = AKSORA_ONERI_ESLE(t, n, y2, 2, 111, n)
    For i = 1 To n
        If CStr(t(i, O_ANAHTAR)) = "B" Then
            If CStr(t(i, O_DURUM)) = D_ACIK Then nA = nA + 1 Else nK = nK + 1
        End If
    Next i
    If nA <> 1 Or nK <> 1 Then AKSORA_ONERI_SINAMA = "yeniden-acilan": Exit Function
    ' 600. gun: eski kapali kayitlar atilir
    t = AKSORA_ONERI_ESLE(t, n, y2, 2, 700, n)
    For i = 1 To n
        If CStr(t(i, O_DURUM)) = D_KAPANDI And t(i, O_KAPANIS) < 200 Then AKSORA_ONERI_SINAMA = "eski-kapali": Exit Function
    Next i
    AKSORA_ONERI_SINAMA = "OK"
    Exit Function
Hata:
    AKSORA_ONERI_SINAMA = "beklenmeyen hata: " & Err.Description
End Function
