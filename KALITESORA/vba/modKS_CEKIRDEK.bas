Attribute VB_Name = "modKS_CEKIRDEK"
Option Explicit

' KALÝTESORA çekirdek modülü: sabitler, renkler ve tüm modüllerin ortak yardýmcýlarý.

Public Const KS_AD As String = "KALÝTESORA"
Public Const KS_SURUM As String = "2.4"
Public Const KS_SIFRE As String = "KS2026"
Public Const KS_ORNEK As String = "ÖRNEK VERÝ"

' Sayfa adlarý
Public Const SH_PANEL As String = "PANEL"
Public Const SH_KAYIT As String = "KAYIT"
Public Const SH_SEVK As String = "SEVKÝYATLAR"
Public Const SH_ARIZA As String = "ARIZALAR"
Public Const SH_MALZ As String = "MALZEMELER"
Public Const SH_KARNE As String = "KARNE"
Public Const SH_RAPOR As String = "RAPORLAR"
Public Const SH_AKTAR As String = "TOPLU AKTAR"
Public Const SH_TANIM As String = "TANIMLAR"
Public Const SH_LOG As String = "GÜNLÜK"

' Tablo adlarý
Public Const T_MALZ As String = "tblMalzeme"
Public Const T_SEVK As String = "tblSevk"
Public Const T_ARIZA As String = "tblAriza"
Public Const T_SEVKCI As String = "tblSevkci"
Public Const T_YER As String = "tblSevkYeri"
Public Const T_KAT As String = "tblKategori"
Public Const T_BIRIM As String = "tblBirim"
Public Const T_TIP As String = "tblArizaTipi"
Public Const T_COZUM As String = "tblCozum"
Public Const T_KARS As String = "tblKarsilayan"
Public Const T_KUR As String = "tblKur"
Public Const T_LOG As String = "tblGunluk"

' Veri tablolarýnýn baþlýk satýrý
Public Const KS_BASLIK_SATIR As Long = 4

' MALZEMELER sütunlarý
Public Const M_KOD As Long = 1
Public Const M_AD As Long = 2
Public Const M_KAT As Long = 3
Public Const M_BIRIM As Long = 4
Public Const M_URETICI As Long = 5
Public Const M_FIYAT As Long = 6
Public Const M_PB As Long = 7
Public Const M_GARANTI As Long = 8
Public Const M_SEVKMIK As Long = 9
Public Const M_ARIZAMIK As Long = 10
Public Const M_ORAN As Long = 11
Public Const M_SONARIZA As Long = 12
Public Const M_NOT As Long = 13
Public Const M_KAYIT As Long = 14
Public Const M_KOLON As Long = 14

' SEVKÝYATLAR sütunlarý
Public Const S_NO As Long = 1
Public Const S_SEVKCI As Long = 2
Public Const S_YER As Long = 3
Public Const S_KOD As Long = 4
Public Const S_AD As Long = 5
Public Const S_ADET As Long = 6
Public Const S_FIYAT As Long = 7
Public Const S_PB As Long = 8
Public Const S_TUTAR As Long = 9
Public Const S_PLAN As Long = 10
Public Const S_GELIS As Long = 11
Public Const S_GECIKME As Long = 12
Public Const S_TDURUM As Long = 13
Public Const S_GARAY As Long = 14
Public Const S_GARBIT As Long = 15
Public Const S_GDURUM As Long = 16
Public Const S_IRS As Long = 17
Public Const S_SERI As Long = 18
Public Const S_ARIZA As Long = 19
Public Const S_NOT As Long = 20
Public Const S_ZAMAN As Long = 21
Public Const S_KULL As Long = 22
Public Const S_KOLON As Long = 22

' ARIZALAR sütunlarý
Public Const A_NO As Long = 1
Public Const A_TARIH As Long = 2
Public Const A_KOD As Long = 3
Public Const A_AD As Long = 4
Public Const A_SERI As Long = 5
Public Const A_URETICI As Long = 6
Public Const A_SEVKCI As Long = 7
Public Const A_YER As Long = 8
Public Const A_SEVK As Long = 9
Public Const A_GARBIT As Long = 10
Public Const A_KAPSAM As Long = 11
Public Const A_KAPKAYNAK As Long = 12
Public Const A_ADET As Long = 13
Public Const A_TIP As Long = 14
Public Const A_ACIK As Long = 15
Public Const A_BILDIRIM As Long = 16
Public Const A_MAIL As Long = 17
Public Const A_MAILDOSYA As Long = 18
Public Const A_DONUS As Long = 19
Public Const A_DONUSGUN As Long = 20
Public Const A_DEVAM As Long = 21
Public Const A_COZULDU As Long = 22
Public Const A_DURUM As Long = 23
Public Const A_COZUM As Long = 24
Public Const A_COZUMGUN As Long = 25
Public Const A_SEKIL As Long = 26
Public Const A_MALIYET As Long = 27
Public Const A_KARS As Long = 28
Public Const A_ONCEKI As Long = 29
Public Const A_TEKRAR As Long = 30
Public Const A_NOT As Long = 31
Public Const A_ZAMAN As Long = 32
Public Const A_KULL As Long = 33
Public Const A_KOLON As Long = 33

Private mDerinlik As Long
Private mEv As Boolean
Private mSc As Boolean
Private mCalc As Long

' ---------------------------------------------------------------- renkler

Public Function C_LAC() As Long
    C_LAC = RGB(12, 45, 79)
End Function

Public Function C_LAC2() As Long
    C_LAC2 = RGB(20, 58, 99)
End Function

Public Function C_CAMGOBEGI() As Long
    C_CAMGOBEGI = RGB(0, 180, 216)
End Function

Public Function C_TURKUAZ() As Long
    C_TURKUAZ = RGB(0, 118, 150)
End Function

Public Function C_ZEMIN() As Long
    C_ZEMIN = RGB(241, 245, 249)
End Function

Public Function C_KART() As Long
    C_KART = RGB(255, 255, 255)
End Function

Public Function C_CIZGI() As Long
    C_CIZGI = RGB(219, 227, 236)
End Function

Public Function C_METIN() As Long
    C_METIN = RGB(30, 41, 59)
End Function

Public Function C_GRI() As Long
    C_GRI = RGB(100, 116, 139)
End Function

Public Function C_GIRIS() As Long
    C_GIRIS = RGB(240, 249, 255)
End Function

Public Function C_OTO() As Long
    C_OTO = RGB(236, 246, 248)
End Function

Public Function C_YESIL() As Long
    C_YESIL = RGB(16, 150, 100)
End Function

Public Function C_KIRMIZI() As Long
    C_KIRMIZI = RGB(220, 53, 69)
End Function

Public Function C_TURUNCU() As Long
    C_TURUNCU = RGB(230, 140, 10)
End Function

Public Function C_SERIT() As Long
    C_SERIT = RGB(246, 249, 252)
End Function


' ---------------------------------------------------------------- biçimler

Public Function KS_FmtTL() As String
    KS_FmtTL = "#,##0.00 """ & ChrW(8378) & """"
End Function

Public Function KS_FmtTLKisa() As String
    KS_FmtTLKisa = "#,##0 """ & ChrW(8378) & """"
End Function

Public Function KS_FmtTarih() As String
    KS_FmtTarih = "dd.mm.yyyy"
End Function

' ---------------------------------------------------------------- uygulama durumu

Public Sub KS_Basla()
    If mDerinlik = 0 Then
        mEv = Application.EnableEvents
        mSc = Application.ScreenUpdating
        mCalc = Application.Calculation
        Application.EnableEvents = False
        Application.ScreenUpdating = False
        Application.Calculation = xlCalculationManual
    End If
    mDerinlik = mDerinlik + 1
End Sub

Public Sub KS_Bitir(Optional ByVal zorla As Boolean = False)
    If zorla Then mDerinlik = 1
    If mDerinlik > 0 Then mDerinlik = mDerinlik - 1
    If mDerinlik = 0 Then
        On Error Resume Next
        If mCalc = 0 Then mCalc = xlCalculationAutomatic
        Application.Calculation = mCalc
        Application.EnableEvents = True
        Application.ScreenUpdating = True
        Application.StatusBar = False
        On Error GoTo 0
    End If
End Sub

Public Sub KS_Durum(ByVal metin As String)
    Application.StatusBar = KS_AD & " | " & metin
End Sub

' ---------------------------------------------------------------- sayfa / tablo eriþimi

Public Function KS_WS(ByVal ad As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, ad, vbTextCompare) = 0 Then
            Set KS_WS = ws
            Exit Function
        End If
    Next ws
End Function

Public Function KS_LO(ByVal ad As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, ad, vbTextCompare) = 0 Then
                Set KS_LO = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

' Tablo yoksa açýklayýcý hata verir.
Public Function KS_LOZ(ByVal ad As String) As ListObject
    Set KS_LOZ = KS_LO(ad)
    If KS_LOZ Is Nothing Then
        Err.Raise vbObjectError + 7001, KS_AD, "Tablo bulunamadý: " & ad & _
            ". Önce KALITESORA_KUR makrosunu çalýþtýrýn."
    End If
End Function

Public Function KS_F(ByVal ad As String) As Range
    Set KS_F = ThisWorkbook.Names(ad).RefersToRange
End Function

Public Function KS_FV(ByVal ad As String) As Variant
    KS_FV = ThisWorkbook.Names(ad).RefersToRange.Value
End Function

Public Function KS_AdVar(ByVal ad As String) As Boolean
    Dim n As Name
    On Error Resume Next
    Set n = ThisWorkbook.Names(ad)
    KS_AdVar = Not n Is Nothing
    On Error GoTo 0
End Function

Public Sub KS_AdTanimla(ByVal ad As String, ByVal formul As String)
    If KS_AdVar(ad) Then ThisWorkbook.Names(ad).Delete
    ThisWorkbook.Names.Add Name:=ad, RefersTo:=formul
End Sub

Public Sub KS_AdHucre(ByVal ad As String, ByVal hucre As Range)
    KS_AdTanimla ad, "='" & hucre.Worksheet.Name & "'!" & hucre.Address
End Sub

' Tablonun veri gövdesini 2 boyutlu dizi olarak verir; boþ tabloda Empty.
Public Function KS_Veri(ByVal lo As ListObject) As Variant
    Dim a As Variant, b() As Variant
    If lo.DataBodyRange Is Nothing Then Exit Function
    a = lo.DataBodyRange.Value
    If IsArray(a) Then
        KS_Veri = a
    Else
        ReDim b(1 To 1, 1 To 1)
        b(1, 1) = a
        KS_Veri = b
    End If
End Function

Public Function KS_Satir(ByVal a As Variant) As Long
    If IsArray(a) Then KS_Satir = UBound(a, 1)
End Function

' Dizinin i. satýrýný 1 satýrlýk yeni bir 2 boyutlu diziye kopyalar.
Public Function KS_SatirKopya(ByRef a As Variant, ByVal i As Long) As Variant
    Dim b() As Variant, j As Long
    ReDim b(1 To 1, 1 To UBound(a, 2))
    For j = 1 To UBound(a, 2)
        b(1, j) = a(i, j)
    Next j
    KS_SatirKopya = b
End Function

Public Function KS_BosSatir(ByVal kolon As Long) As Variant
    Dim b() As Variant
    ReDim b(1 To 1, 1 To kolon)
    KS_BosSatir = b
End Function

' satir = 0 ise sona yeni satýr ekler. Yazýlan ListRow sýrasýný döndürür.
Public Function KS_SatirYaz(ByVal lo As ListObject, ByVal satir As Long, ByRef v As Variant) As Long
    If satir = 0 Then
        lo.ListRows.Add
        satir = lo.ListRows.Count
    End If
    lo.ListRows(satir).Range.Value = v
    KS_SatirYaz = satir
End Function

' Çok satýrý tek seferde tablonun sonuna ekler (toplu aktarým ve örnek veri için).
Public Sub KS_TopluEkle(ByVal lo As ListObject, ByRef v As Variant, ByVal adet As Long)
    Dim w() As Variant, i As Long, j As Long, c As Long, ilk As Long
    Dim ws As Worksheet, bas As Range
    If adet <= 0 Then Exit Sub
    c = lo.ListColumns.Count
    ReDim w(1 To adet, 1 To c)
    For i = 1 To adet
        For j = 1 To c
            w(i, j) = v(i, j)
        Next j
    Next i
    Set ws = lo.Parent
    Set bas = lo.HeaderRowRange.Cells(1, 1)
    ilk = bas.Row + lo.ListRows.Count + 1
    ws.Cells(ilk, bas.Column).Resize(adet, c).Value = w
    lo.Resize ws.Range(bas, ws.Cells(ilk + adet - 1, bas.Column + c - 1))
End Sub

Public Function KS_SatirBul(ByVal lo As ListObject, ByVal kolon As Long, ByVal anahtar As Variant) As Long
    Dim a As Variant, i As Long, k As String
    k = KS_Anahtar(anahtar)
    If Len(k) = 0 Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    a = lo.ListColumns(kolon).DataBodyRange.Value
    If Not IsArray(a) Then
        If KS_Anahtar(a) = k Then KS_SatirBul = 1
        Exit Function
    End If
    For i = 1 To UBound(a, 1)
        If KS_Anahtar(a(i, 1)) = k Then
            KS_SatirBul = i
            Exit Function
        End If
    Next i
End Function

Public Function KS_YeniNo(ByVal lo As ListObject, ByVal onek As String) As String
    KS_YeniNo = onek & Format$(KS_EnBuyukNo(lo, onek) + 1, "000000")
End Function

Public Function KS_EnBuyukNo(ByVal lo As ListObject, ByVal onek As String) As Long
    Dim a As Variant, i As Long, s As String, m As Long
    a = KS_Veri(lo)
    If Not IsArray(a) Then Exit Function
    For i = 1 To UBound(a, 1)
        If Not IsError(a(i, 1)) Then
            s = UCase$(Trim$(CStr(a(i, 1))))
            If Left$(s, Len(onek)) = onek Then
                If IsNumeric(Mid$(s, Len(onek) + 1)) Then
                    If CLng(Mid$(s, Len(onek) + 1)) > m Then m = CLng(Mid$(s, Len(onek) + 1))
                End If
            End If
        End If
    Next i
    KS_EnBuyukNo = m
End Function

Public Function KS_NoBicim(ByVal onek As String, ByVal n As Long) As String
    KS_NoBicim = onek & Format$(n, "000000")
End Function

' Tek sütunluk taným tablosundan büyük harfli anahtar sözlüðü.
Public Function KS_Sozluk(ByVal tablo As String, Optional ByVal kolon As Long = 1) As Object
    Dim d As Object, a As Variant, i As Long, k As String, lo As ListObject
    Set d = CreateObject("Scripting.Dictionary")
    Set lo = KS_LO(tablo)
    If Not lo Is Nothing Then
        a = KS_Veri(lo)
        If IsArray(a) Then
            For i = 1 To UBound(a, 1)
                k = KS_Anahtar(a(i, kolon))
                If Len(k) > 0 Then d(k) = i
            Next i
        End If
    End If
    Set KS_Sozluk = d
End Function

' ---------------------------------------------------------------- koruma

Public Function KS_KorumaAktif() As Boolean
    KS_KorumaAktif = True
    On Error Resume Next
    KS_KorumaAktif = (ThisWorkbook.Names("_KS_KORUMA").RefersTo <> "=0")
    On Error GoTo 0
End Function

Public Sub KS_Ac(ByVal ws As Worksheet)
    If ws.ProtectContents Or ws.ProtectDrawingObjects Then ws.Unprotect KS_SIFRE
End Sub

Public Sub KS_Kapat(ByVal ws As Worksheet)
    If Not KS_KorumaAktif() Then Exit Sub
    If ws.ProtectContents Then Exit Sub
    ' TANIMLAR açýk kalýr: kullanýcý listelere satýr ekleyebilmeli.
    If StrComp(ws.Name, SH_TANIM, vbTextCompare) = 0 Then Exit Sub
    ws.Protect Password:=KS_SIFRE, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        AllowFormattingColumns:=True, AllowFiltering:=True, AllowSorting:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Public Sub KS_KORUMA_DEGISTIR()
    Dim s As String, ws As Worksheet, ac As Boolean
    s = InputBox("Koruma þifresini girin:", KS_AD)
    If Len(s) = 0 Then Exit Sub
    If s <> KS_SIFRE Then
        MsgBox "Þifre hatalý.", vbExclamation, KS_AD
        Exit Sub
    End If
    ac = KS_KorumaAktif()
    KS_AdTanimla "_KS_KORUMA", IIf(ac, "=0", "=1")
    ThisWorkbook.Names("_KS_KORUMA").Visible = False
    For Each ws In ThisWorkbook.Worksheets
        If KS_BizimSayfa(ws.Name) Then
            If ac Then KS_Ac ws Else KS_Kapat ws
        End If
    Next ws
    KS_Log "KORUMA", "", IIf(ac, "Sayfa korumasý kapatýldý", "Sayfa korumasý açýldý")
    MsgBox "Sayfa korumasý " & IIf(ac, "KAPATILDI. Doðrudan düzenleme açýk; iþiniz bitince tekrar açýn.", "AÇILDI."), _
        vbInformation, KS_AD
End Sub

Public Function KS_Sayfalar() As Variant
    KS_Sayfalar = Array(SH_PANEL, SH_KAYIT, SH_SEVK, SH_ARIZA, SH_MALZ, SH_KARNE, _
        SH_RAPOR, SH_AKTAR, SH_TANIM, SH_LOG)
End Function

Public Function KS_BizimSayfa(ByVal ad As String) As Boolean
    Dim s As Variant
    For Each s In KS_Sayfalar()
        If StrComp(CStr(s), ad, vbTextCompare) = 0 Then
            KS_BizimSayfa = True
            Exit Function
        End If
    Next s
End Function

' ---------------------------------------------------------------- deðer çözümleme

Public Function KS_Anahtar(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    KS_Anahtar = UCase$(Trim$(CStr(v)))
End Function

' SAP sevkçi biçimi "1300000329 HALÝÇ TEKNÝK ..." -> "HALÝÇ TEKNÝK ...". Ayný firma bölge þirketine göre
' farklý satýcý koduyla gelebilir; karne ve sýralamada firma adý birleþtirir.
Public Function KS_FirmaAd(ByVal v As Variant) As String
    Dim s As String, p As Long
    s = Application.WorksheetFunction.Trim(KS_Metin(v))
    p = InStr(1, s, " ")
    If p > 1 Then
        If IsNumeric(Left$(s, p - 1)) Then s = Mid$(s, p + 1)
    End If
    KS_FirmaAd = s
End Function

Public Function KS_FirmaAnahtar(ByVal v As Variant) As String
    KS_FirmaAnahtar = KS_Anahtar(KS_FirmaAd(v))
End Function

Public Function KS_Bos(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Or IsNull(v) Then
        KS_Bos = True
    Else
        KS_Bos = (Len(Trim$(CStr(v))) = 0)
    End If
End Function

Public Function KS_Metin(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or IsNull(v) Then Exit Function
    KS_Metin = Trim$(CStr(v))
End Function

' Tarih, seri numarasý veya gg.aa.yyyy / yyyy-aa-gg metnini kabul eder.
Public Function KS_Tarih(ByVal v As Variant, ByRef d As Date) As Boolean
    Dim s As String, p() As String, g As Long, a As Long, y As Long, t As Date
    If IsError(v) Then Exit Function
    If KS_Bos(v) Then Exit Function
    Select Case VarType(v)
        Case vbDate
            t = v
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal
            If v < 32874 Or v > 73051 Then Exit Function
            t = CDate(Int(CDbl(v)))
        Case vbString
            s = Trim$(CStr(v))
            If InStr(s, " ") > 0 Then s = Left$(s, InStr(s, " ") - 1)
            s = Replace(Replace(s, "/", "."), "-", ".")
            p = Split(s, ".")
            If UBound(p) <> 2 Then Exit Function
            If Not (KS_Rakam(p(0)) And KS_Rakam(p(1)) And KS_Rakam(p(2))) Then Exit Function
            If Len(p(0)) = 4 Then
                y = CLng(p(0)): a = CLng(p(1)): g = CLng(p(2))
            Else
                g = CLng(p(0)): a = CLng(p(1)): y = CLng(p(2))
                If y < 100 Then y = 2000 + y
            End If
            If a < 1 Or a > 12 Or g < 1 Or g > 31 Then Exit Function
            t = DateSerial(y, a, g)
            If Day(t) <> g Then Exit Function
        Case Else
            Exit Function
    End Select
    If Year(t) < 1990 Or Year(t) > 2100 Then Exit Function
    d = Int(t)
    KS_Tarih = True
End Function

Private Function KS_Rakam(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Or Len(s) > 4 Then Exit Function
    For i = 1 To Len(s)
        If Mid$(s, i, 1) < "0" Or Mid$(s, i, 1) > "9" Then Exit Function
    Next i
    KS_Rakam = True
End Function

Public Function KS_Sayi(ByVal v As Variant, ByRef x As Double) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    If KS_Bos(v) Then Exit Function
    Select Case VarType(v)
        Case vbDouble, vbSingle, vbLong, vbInteger, vbCurrency, vbDecimal, vbByte
            x = CDbl(v)
            KS_Sayi = True
        Case vbString
            s = Replace(CStr(v), ChrW(8378), "")
            s = Trim$(Replace(s, "TL", "", , , vbTextCompare))
            If IsNumeric(s) Then
                x = CDbl(s)
                KS_Sayi = True
            End If
    End Select
End Function

Public Function KS_Kullanici() As String
    KS_Kullanici = Environ$("USERNAME")
    If Len(KS_Kullanici) = 0 Then KS_Kullanici = Application.UserName
End Function

Public Function KS_ParamSayi(ByVal ad As String, ByVal varsayilan As Double) As Double
    Dim v As Variant, x As Double
    KS_ParamSayi = varsayilan
    On Error GoTo Son
    v = KS_FV(ad)
    If KS_Sayi(v, x) Then KS_ParamSayi = x
Son:
End Function

' ---------------------------------------------------------------- günlük

Public Sub KS_Log(ByVal islem As String, ByVal kayit As String, ByVal detay As String)
    Dim lo As ListObject, v As Variant
    On Error GoTo Son
    Set lo = KS_LO(T_LOG)
    If lo Is Nothing Then Exit Sub
    KS_Ac lo.Parent
    v = KS_BosSatir(5)
    v(1, 1) = Now
    v(1, 2) = KS_Kullanici()
    v(1, 3) = islem
    v(1, 4) = kayit
    v(1, 5) = Left$(detay, 250)
    KS_SatirYaz lo, 0, v
Son:
    If Not lo Is Nothing Then KS_Kapat lo.Parent
End Sub

' ---------------------------------------------------------------- sýralama

' a dizisinin satýr sýralarýný, kolon deðerine göre sýralý indeks dizisi olarak verir.
Public Function KS_SiraIndeks(ByRef a As Variant, ByVal kolon As Long, ByVal azalan As Boolean, _
        Optional ByVal adet As Long = -1) As Variant
    Dim idx() As Long, n As Long, i As Long, j As Long, t As Long
    If adet < 0 Then adet = UBound(a, 1)
    n = adet
    If n = 0 Then Exit Function
    ReDim idx(1 To n)
    For i = 1 To n
        idx(i) = i
    Next i
    For i = 2 To n
        t = idx(i)
        j = i - 1
        Do While j >= 1
            If Not KS_Once(a(t, kolon), a(idx(j), kolon), azalan) Then Exit Do
            idx(j + 1) = idx(j)
            j = j - 1
        Loop
        idx(j + 1) = t
    Next i
    KS_SiraIndeks = idx
End Function

Private Function KS_Once(ByVal x As Variant, ByVal y As Variant, ByVal azalan As Boolean) As Boolean
    Dim bosX As Boolean, bosY As Boolean, k As Long
    bosX = KS_Bos(x) Or IsError(x)
    bosY = KS_Bos(y) Or IsError(y)
    If bosX Or bosY Then
        KS_Once = (Not bosX And bosY)
        Exit Function
    End If
    If KS_SayisalTur(x) And KS_SayisalTur(y) Then
        If CDbl(x) > CDbl(y) Then
            k = 1
        ElseIf CDbl(x) < CDbl(y) Then
            k = -1
        End If
    Else
        k = StrComp(CStr(x), CStr(y), vbTextCompare)
    End If
    If azalan Then KS_Once = (k > 0) Else KS_Once = (k < 0)
End Function

Private Function KS_SayisalTur(ByVal v As Variant) As Boolean
    Select Case VarType(v)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbDate, vbDecimal, vbByte
            KS_SayisalTur = True
    End Select
End Function
