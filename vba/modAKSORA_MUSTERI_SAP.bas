Attribute VB_Name = "modAKSORA_MUSTERI_SAP"
Option Explicit

'==========================================================
' AKSORA - MUSTERI SAP (7155)  | Surum: MB52'siz, fiyat listeli
' Sorgu 1: ZMMRMC001 -> CDP (E001-E999,D003,D005,S001,Z002,Z003,Z007,EM14) + Yuklenici (M132)
' Sorgu 2: ZMMRMC001 -> Ilce (K001-K999 + 3 mal grubu)
' Deger = Miktar x Birim fiyat (modAKSORA_MUSTERI_FIYAT, Malzeme+Parti)
'==========================================================

#If VBA7 Then
Private Declare PtrSafe Sub MS_Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
Private Declare PtrSafe Function MS_CoRegisterMessageFilter Lib "ole32" Alias "CoRegisterMessageFilter" _
    (ByVal yeni As LongPtr, ByRef onceki As LongPtr) As Long
Private mOleFiltre As LongPtr
Private mOleBos As LongPtr
Private Declare PtrSafe Function MS_PID Lib "kernel32" Alias "GetCurrentProcessId" () As Long
#Else
Private Declare Function MS_PID Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare Sub MS_Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
Private Declare Function MS_CoRegisterMessageFilter Lib "ole32" Alias "CoRegisterMessageFilter" _
    (ByVal yeni As Long, ByRef onceki As Long) As Long
Private mOleFiltre As Long
Private mOleBos As Long
#End If
Private mOleKapali As Boolean
Private mHizliSebep As String
Private mHizliSatir As Long
Private mStandartSatir As Long
Private mSapSatir As Long
Private mListeSatir As Long
Private mStandartOnay As Boolean
Private mDisaDosyalar As Collection
Private mBasZamanMetin As String
Private mFiyatsiz As String
Private mFiyatsizAdet As Long
Private mCalisiyor As Boolean
Private mBas As Double
Private mDurum As String
Private mFiyatTur As Boolean
Private mSeritZaman As Double

Private Function MS_Gecen(ByVal bas As Double) As Double
    MS_Gecen = Timer - bas
    If MS_Gecen < 0 Then MS_Gecen = MS_Gecen + 86400#
End Function

Private Sub MS_Durum(ByVal metin As String)
    mDurum = metin
    application.StatusBar = "AKSORA | Müþteri | " & metin
    If Abs(Timer - mSeritZaman) >= 1 Then
        mSeritZaman = Timer
        AKSORA_MUSTERI_SERIT "Güncelleniyor  •  " & metin, 1
    End If
    DoEvents
End Sub

Private Sub MS_Bekle(ByVal ses As Object)
    Dim bas As Double
    bas = Timer
    Do While ses.Busy
        If MS_Gecen(bas) > 900 Then Err.Raise vbObjectError + 1900, , "SAP 15 dakika içinde yanýt vermedi."
        DoEvents: MS_Sleep 50
    Loop
    If ses.FindById("wnd[0]/sbar").MessageType = "E" Then _
        Err.Raise vbObjectError + 1901, , ses.FindById("wnd[0]/sbar").Text
End Sub









Private Sub MS_OleUyariKapat()
    If mOleKapali Then Exit Sub
    On Error Resume Next
    MS_CoRegisterMessageFilter 0, mOleFiltre
    mOleKapali = (Err.Number = 0)
    On Error GoTo 0
End Sub

Private Sub MS_OleUyariAc()
    If Not mOleKapali Then Exit Sub
    On Error Resume Next
    MS_CoRegisterMessageFilter mOleFiltre, mOleBos
    mOleKapali = False
    On Error GoTo 0
End Sub

Private Function MS_SayiDene(ByVal v As Variant, ByRef sonuc As Double) As Boolean
    On Error GoTo Cikis
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If Len(Trim$(CStr(v))) = 0 Then Exit Function
    sonuc = MS_Sayi(v)
    MS_SayiDene = True
Cikis:
End Function

' SAP raporundaki deger sutunu satir degeri olarak alinir (Yeni / Cikma / Hurda ayri degerlenir).
Private Sub MS_SapDeger(ByRef buf As Variant, ByVal r As Long)
    Dim d As Double
    If MS_SayiDene(buf(r, 11), d) Then
        buf(r, 11) = d: buf(r, 13) = "SAP"
    Else
        buf(r, 11) = vbNullString
    End If
End Sub

'----------------------------------------------------------
' SAP TESHIS VE PENCERE YARDIMCILARI
'----------------------------------------------------------
Private Function MS_TeshisYolu() As String
    Dim k As String
    k = ThisWorkbook.Path
    If Len(k) = 0 Or Left$(k, 4) = "http" Then k = Environ$("TEMP")
    MS_TeshisYolu = k & application.PathSeparator & "AKSORA_SAP_TESHIS.txt"
End Function

Private Sub MS_Teshis(ByVal metin As String)
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open MS_TeshisYolu() For Append As #f
    Print #f, metin
    Close #f
    On Error GoTo 0
End Sub

Private Sub MS_Agac(ByVal kok As Object, ByVal derinlik As Long)
    Dim i As Long, c As Object, t As String, adet As Long
    If derinlik > 7 Then Exit Sub
    On Error Resume Next
    adet = kok.Children.Count
    For i = 0 To adet - 1
        Set c = Nothing
        Set c = kok.Children(i)
        If Not c Is Nothing Then
            t = vbNullString: t = c.Text
            MS_Teshis String$(derinlik * 2, " ") & c.Type & " | " & c.Id & " | " & c.Name & " | " & Left$(t, 60)
            If c.ContainerType Then MS_Agac c, derinlik + 1
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function MS_AlanBul(ByVal kok As Object, ByVal ad As String) As Object
    Dim i As Long, c As Object, alt As Object, adet As Long
    On Error Resume Next
    adet = kok.Children.Count
    For i = 0 To adet - 1
        Set c = Nothing
        Set c = kok.Children(i)
        If Not c Is Nothing Then
            If UCase$(c.Name) = ad Then Set MS_AlanBul = c: Exit Function
            If c.ContainerType Then
                Set alt = Nothing
                Set alt = MS_AlanBul(c, ad)
                If Not alt Is Nothing Then Set MS_AlanBul = alt: Exit Function
            End If
        End If
    Next i
    On Error GoTo 0
End Function

Private Sub MS_Topla(ByVal kok As Object, ByVal tip As String, ByVal liste As Collection)
    Dim i As Long, c As Object, adet As Long
    On Error Resume Next
    adet = kok.Children.Count
    For i = 0 To adet - 1
        Set c = Nothing
        Set c = kok.Children(i)
        If Not c Is Nothing Then
            If c.Type = tip Then liste.Add c
            If c.ContainerType Then MS_Topla c, tip, liste
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function MS_PencereBekle(ByVal ses As Object, ByVal saniye As Double) As Object
    Dim bas As Double, w As Object
    bas = Timer
    Do
        Set w = Nothing
        On Error Resume Next
        Set w = ses.FindById("wnd[1]", False)
        On Error GoTo 0
        If Not w Is Nothing Then
            If Not ses.Busy Then Set MS_PencereBekle = w: Exit Function
        End If
        If MS_Gecen(bas) > saniye Then Exit Function
        DoEvents: MS_Sleep 100
    Loop
End Function

'----------------------------------------------------------
' Disa aktarma penceresini acar: arac cubugu, menu cubugu, %PC.
'----------------------------------------------------------
Private Sub MS_MenuTopla(ByVal kok As Object, ByVal yol As String, ByVal liste As Collection, ByVal derinlik As Long)
    Dim i As Long, c As Object, adet As Long, t As String, alt As Long
    If derinlik > 4 Then Exit Sub
    On Error Resume Next
    adet = kok.Children.Count
    For i = 0 To adet - 1
        Set c = Nothing
        Set c = kok.Children(i)
        If Not c Is Nothing Then
            t = vbNullString: t = c.Text
            alt = -1: alt = c.Children.Count
            If alt > 0 Then
                MS_MenuTopla c, yol & " > " & t, liste, derinlik + 1
            Else
                liste.Add Array(c.Id, yol & " > " & t)
            End If
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function MS_DisaAktarAc(ByVal ses As Object, ByVal alv As Object) As Boolean
    Dim i As Long, adet As Long, bid As String, ipucu As String, tip As String, t As String
    Dim menuler As Collection, m As Variant, secim As String, puan As Long, enIyi As Long
    ' 1) ALV arac cubugu
    On Error Resume Next
    adet = alv.ToolbarButtonCount
    For i = 0 To adet - 1
        bid = vbNullString: ipucu = vbNullString: tip = vbNullString
        bid = alv.GetToolbarButtonId(i): ipucu = alv.GetToolbarButtonTooltip(i): tip = alv.GetToolbarButtonType(i)
        MS_Teshis "Araç düðmesi: " & bid & " | " & tip & " | " & ipucu
        t = AKFIX_NORMAL(bid & " " & ipucu)
        If InStr(t, "EXPORT") > 0 Or InStr(t, "DISA AKTAR") > 0 Or InStr(t, "&PC") > 0 Then
            Err.Clear
            If tip = "Button" Then
                alv.PressToolbarButton bid
            Else
                alv.PressToolbarContextButton bid
                alv.SelectContextMenuItem "&PC"
                If Err.Number <> 0 Then Err.Clear: alv.SelectContextMenuItemByText "Yerel dosya..."
                If Err.Number <> 0 Then Err.Clear: alv.SelectContextMenuItemByText "Local File..."
            End If
            If Err.Number = 0 Then
                If Not MS_PencereBekle(ses, 6) Is Nothing Then MS_DisaAktarAc = True: Exit Function
            End If
            MS_Teshis "Araç düðmesi denemesi sonuçsuz: " & bid & " " & Err.Description
            Err.Clear
        End If
    Next i
    ' 2) Menu cubugu: Yerel dosya / Kaydet > Dosya
    Set menuler = New Collection
    MS_MenuTopla ses.FindById("wnd[0]/mbar"), "Menü", menuler, 0
    For Each m In menuler
        MS_Teshis "Menü öðesi: " & m(0) & " | " & m(1)
        t = AKFIX_NORMAL(CStr(m(1))): puan = 0
        If InStr(t, "YEREL DOSYA") > 0 Or InStr(t, "LOCAL FILE") > 0 Then puan = 3
        If puan = 0 And (InStr(t, "KAYDET") > 0 Or InStr(t, "SAVE") > 0) And (InStr(t, "> DOSYA") > 0 Or InStr(t, "> FILE") > 0) Then puan = 2
        If puan = 0 And (InStr(t, "DISA AKTAR") > 0 Or InStr(t, "EXPORT") > 0) And (InStr(t, "DOSYA") > 0 Or InStr(t, "FILE") > 0) Then puan = 1
        If puan > enIyi Then enIyi = puan: secim = CStr(m(0))
    Next m
    If Len(secim) > 0 Then
        Err.Clear
        MS_Teshis "Menü seçiliyor: " & secim
        ses.FindById(secim).Select
        If Err.Number = 0 Then
            If Not MS_PencereBekle(ses, 6) Is Nothing Then MS_DisaAktarAc = True: Exit Function
        End If
        MS_Teshis "Menü denemesi sonuçsuz: " & Err.Description
        Err.Clear
    End If
    ' 3) %PC komutu
    ses.FindById("wnd[0]/tbar[0]/okcd").Text = "%PC"
    ses.FindById("wnd[0]").SendVKey 0
    If Not MS_PencereBekle(ses, 6) Is Nothing Then MS_DisaAktarAc = True: Exit Function
    MS_Teshis "%PC sonuçsuz " & Err.Description
    On Error GoTo 0
End Function

'==========================================================
' HIZLI OKUMA (Malzeme birimindeki yontem): ALV -> Ctrl+Shift+F7
' ile Excel'e aktarilir, acilan kitap tek seferde okunur.
'==========================================================
Private Function MS_AcikKitaplar() As Object
    Dim d As Object, wb As Workbook
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For Each wb In application.Workbooks: d(wb.Name) = True: Next wb
    Set MS_AcikKitaplar = d
End Function

Private Function MS_YeniKitapBekle(ByVal eski As Object, ByVal saniye As Double, ByVal ses As Object) As Workbook
    Dim bas As Double, wb As Workbook, w As Object, onay As Boolean
    bas = Timer
    Do
        DoEvents
        For Each wb In application.Workbooks
            If Not wb Is ThisWorkbook Then
                If Not eski.Exists(wb.Name) Then Set MS_YeniKitapBekle = wb: Exit Function
            End If
        Next wb
        If Not onay Then
            Set w = Nothing
            On Error Resume Next
            Set w = ses.FindById("wnd[1]", False)
            If Not w Is Nothing Then
                MS_Teshis "Aktarým penceresi: " & w.Text
                MS_Agac w, 0
                ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
                onay = True
            End If
            On Error GoTo 0
        End If
        If MS_Gecen(bas) > saniye Then Exit Function
        MS_Sleep 100
    Loop
End Function

Private Function MS_Metin(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbDecimal Then
        If v = Fix(v) And Abs(v) < 1E+15 Then MS_Metin = Format$(v, "0"): Exit Function
    End If
    MS_Metin = Trim$(CStr(v))
End Function

Private Function MS_ALV_Excel(ByVal ses As Object, ByVal alv As Object, ByRef ids() As String, ByVal n As Long) As Variant
    Dim eski As Object, wb As Workbook, sh As Worksheet, en As Worksheet, v As Variant
    Dim h As Long, r As Long, c As Long, k As Long, dolu As Long, no As Long, adet As Long
    Dim pos(1 To 11) As Long, buf() As Variant, sonuc() As Variant, eksikAlan As String
    Dim uy As String, mz As String, a As String, b As String, x As Variant, ornek As Variant
    Dim j As Long, bulundu As Boolean, uygun As Boolean
    On Error GoTo Hata
    mHizliSebep = "Excel'e aktarým"
    Set eski = MS_AcikKitaplar()
    MS_Teshis "---- " & Format$(Now, "dd.mm.yyyy hh:nn:ss") & " | Excel'e aktarým (Ctrl+Shift+F7) | ALV satýr " & CStr(n)
    ses.FindById("wnd[0]").SendVKey 43
    Set wb = MS_YeniKitapBekle(eski, 90, ses)
    If wb Is Nothing Then mHizliSebep = "Excel'e aktarýlan liste açýlmadý": GoTo Temizle
    For Each sh In wb.Worksheets
        If en Is Nothing Then
            Set en = sh
        ElseIf sh.UsedRange.Rows.Count > en.UsedRange.Rows.Count Then
            Set en = sh
        End If
    Next sh
    v = en.UsedRange.Value2
    If Not IsArray(v) Then mHizliSebep = "aktarýlan liste boþ": GoTo Temizle
    For r = 1 To application.Min(10, UBound(v, 1))
        dolu = 0
        For c = 1 To UBound(v, 2)
            If Len(MS_Metin(v(r, c))) > 0 Then dolu = dolu + 1
        Next c
        If dolu >= 5 Then h = r: Exit For
    Next r
    If h = 0 Then mHizliSebep = "aktarýlan listede baþlýk bulunamadý": GoTo Temizle
    For c = 1 To UBound(v, 2)
        no = MS_Alan(MS_Norm(MS_Metin(v(h, c))), False)
        MS_Teshis "Excel kolon " & CStr(c) & ": " & MS_Metin(v(h, c)) & " -> alan " & CStr(no)
        If no > 0 Then
            If pos(no) = 0 Then pos(no) = c
        End If
    Next c
    For c = 1 To 10
        If pos(c) = 0 Then eksikAlan = eksikAlan & " " & CStr(c)
    Next c
    If Len(eksikAlan) > 0 Then mHizliSebep = "aktarýlan listede eksik sütun:" & eksikAlan: GoTo Temizle
    mHizliSebep = "aktarýlan veri"
    ReDim buf(1 To UBound(v, 1), 1 To 14)
    For r = h + 1 To UBound(v, 1)
        uy = MS_Metin(v(r, pos(1))): mz = MS_Metin(v(r, pos(5)))
        If Len(uy) > 0 Or Len(mz) > 0 Then
            If uy <> "7155" Or Len(mz) = 0 Then
                If Len(mz) > 0 Then mHizliSebep = "beklenmeyen satýr (ÜY " & uy & ")": GoTo Temizle
            Else
                adet = adet + 1
                For c = 1 To 10
                    buf(adet, c) = MS_Metin(v(r, pos(c)))
                Next c
                buf(adet, 7) = MS_Sayi(v(r, pos(7)))
                If pos(11) > 0 Then buf(adet, 11) = v(r, pos(11))
            End If
        End If
    Next r
    If adet = 0 Or adet > n Or adet < n - 50 Then
        mHizliSebep = "satýr sayýsý (Excel " & CStr(adet) & " / SAP " & CStr(n) & ")": GoTo Temizle
    End If
    ' Ornek satirlar SAP ekranindaki degerlerle karsilastirilir.
    mHizliSebep = "örnek kontrol"
    For Each ornek In Array(0, adet \ 2, adet - 1)
        bulundu = False
        For j = CLng(ornek) To application.Min(n - 1, CLng(ornek) + 60)
            If j Mod 20 = CLng(ornek) Mod 20 Then
                alv.FirstVisibleRow = application.Min(j, application.Max(0, n - CLng(alv.VisibleRowCount)))
                DoEvents
            End If
            uygun = True
            For Each x In Array(2, 5, 7)
                a = CStr(buf(CLng(ornek) + 1, CLng(x)))
                b = vbNullString
                On Error Resume Next
                b = Trim$(CStr(alv.GetCellValue(j, ids(CLng(x)))))
                On Error GoTo Hata
                If CLng(x) = 5 Then a = MS_Malzeme(a): b = MS_Malzeme(b)
                If CLng(x) = 7 Then
                    a = CStr(CDbl(a))
                    If Len(b) > 0 Then b = CStr(MS_Sayi(b))
                End If
                If UCase$(a) <> UCase$(b) Then uygun = False: Exit For
            Next x
            If uygun Then bulundu = True: Exit For
        Next j
        If Not bulundu Then
            mHizliSebep = "örnek kontrol (Excel satýr " & CStr(CLng(ornek) + 1) & " SAP ekranýnda bulunamadý)"
            GoTo Temizle
        End If
    Next ornek
    ReDim sonuc(1 To adet, 1 To 14)
    For r = 1 To adet
        For c = 1 To 14: sonuc(r, c) = buf(r, c): Next c
        MS_SapDeger sonuc, r
    Next r
    On Error Resume Next
    wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0
    MS_Teshis "Excel'e aktarým baþarýlý: " & CStr(adet) & " satýr"
    mHizliSebep = vbNullString
    MS_ALV_Excel = sonuc
    Exit Function
Hata:
    mHizliSebep = mHizliSebep & " / " & Err.Description
    Resume Temizle
Temizle:
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    MS_Teshis "Excel'e aktarým kullanýlamadý: " & mHizliSebep
    MS_ALV_Excel = Empty
End Function

'==========================================================
' MB52 DEGERLERI: Depo + Malzeme + Degerleme turu bazinda gercek
' SAP degeri alinir; seri numarali satirlara birim deger olarak dagitilir.
'==========================================================
Private Function MS_Mb52Anahtar(ByVal depo As Variant, ByVal malzeme As Variant, ByVal parti As Variant) As String
    MS_Mb52Anahtar = UCase$(MS_Metin(depo)) & "|" & MS_Malzeme(MS_Metin(malzeme)) & "|" & AKFIX_NORMAL(MS_Metin(parti))
End Function

Private Function MS_MB52_SOZLUK(ByRef hMesaj As String) As Object
    Dim veri As Variant, satir As Long, sutun As Long, sure As Double, d As Object
    Dim r As Long, k As String, t As Variant, q As Double, deger As Double, i As Long
    On Error GoTo Hata
    If Not STOK_EXPORT_AL("CDP", "7155", "*", "*", veri, satir, sutun, sure, hMesaj) Then
        MS_Teshis "MB52 alýnamadý: " & hMesaj
        Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = 1 To satir
        q = 0: deger = 0
        If IsNumeric(veri(r, 9)) Then q = CDbl(veri(r, 9))
        If IsNumeric(veri(r, 11)) Then deger = CDbl(veri(r, 11))
        For i = 0 To 1
            If i = 0 Then
                k = MS_Mb52Anahtar(veri(r, 2), veri(r, 7), veri(r, 6))
            Else
                k = MS_Mb52Anahtar("*", veri(r, 7), veri(r, 6))
            End If
            If d.Exists(k) Then
                t = d(k): t(0) = t(0) + q: t(1) = t(1) + deger: d(k) = t
            Else
                d.Add k, Array(q, deger)
            End If
        Next i
    Next r
    MS_Teshis "MB52 alýndý: " & CStr(satir) & " satýr, " & AKSORA_SAYI_METNI(sure, "0.0") & " sn"
    Set MS_MB52_SOZLUK = d
    Exit Function
Hata:
    hMesaj = "MB52 okunamadý: " & Err.Description
    MS_Teshis hMesaj
End Function

' Satir degeri = satir miktari x (MB52 deger / MB52 miktar). Bulunamazsa fiyat listesi kullanilir.
Private Sub MS_Mb52Uygula(ByRef a As Variant, ByVal d As Object)
    Dim r As Long, k As String, t As Variant, birim As Double, q As Double
    If d Is Nothing Or IsEmpty(a) Then Exit Sub
    For r = 1 To UBound(a, 1)
        k = MS_Mb52Anahtar(a(r, 2), a(r, 5), a(r, 9))
        If Not d.Exists(k) Then k = MS_Mb52Anahtar("*", a(r, 5), a(r, 9))
        If d.Exists(k) Then
            t = d(k)
            If CDbl(t(0)) <> 0 Then
                birim = CDbl(t(1)) / CDbl(t(0))
                q = 0
                If IsNumeric(a(r, 7)) Then q = CDbl(a(r, 7))
                a(r, 11) = Round(q * birim, 2): a(r, 13) = "SAP"
            End If
        End If
    Next r
End Sub

Private Function MS_DepoAdi(ByVal uy As String, ByVal depo As String) As String
    ' Depo hafizasi: 7151-7155 tum depolar (kod -> ad).
    Static d As Object
    Dim s As String, p As Variant, x As Variant
    If d Is Nothing Then
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
        s = s & "0001|0001|Lager 0001;0001|0088|Lager 0088 (WM);0001|0100|Lagerort WM&HU;0001|0200|SCM-EWM;0003|0001|Lager 0001;0003|0088|Lager 0088 (WM);7151|B201|Kullaným dýþý.;7151|B301|Kullaným dýþý;7151|BO01|14.53.OB.01;7151|BO02|15.53.OB.01;7151|BO03|15.01.ÝNAL;7151|BO04|18.01.KORDON;7151|BO05|21.53.OB.01;7151|D001|Ýþletme Ambarý;7151|D002|Yatýrým Ambarý;7151|D003|Müþ.Bað.Ambarý;7151|D004|Atýl Malzeme Amb;7151|D006|Arýza Mlz. Ambar;7151|D007|Demontaj Ambarý;7151|D008|Hurda Ambarý;7151|D009|ÝSG-HIRDAVAT-ÞEB;7151|D010|ÝSG-HIRDAVAT-MÜÞ;7151|E001|BAÐLANTI Ekip1;7151|E002|BAÐLANTI Ekip2;7151|E003|BAÐLANTI Ekip3;7151|E004|BAÐLANTI Ekip4;7151|E005|BAÐLANTI Ekip5;7151|E006|BAÐLANTI Ekip6;"
        s = s & "7151|EM23|23.53.EMANET;7151|EM24|24.53.EMANET;7151|G001|Ýl Müdürlüðü Dp;7151|G002|ÝDARÝ ÝÞ-FÝLO-BT;7151|G004|EDAÞ EMANET ÞEB;7151|G005|EMANET DEPO YTRM;7151|G006|AFAD;7151|K001|BAÐLANTI TES. KA;7151|K002|ARDEÞEN KA;7151|K003|ÇAMLIHEMÞÝN KA;7151|K004|ÇAYELÝ KA;7151|K005|DEREPAZARI KA;7151|K006|FINDIKLI KA;7151|K007|GÜNDOÐDU KA;7151|K008|GÜNEYSU KA;7151|K009|HAT BAKIM KA;7151|K010|HEMÞÝN KA;7151|K011|ÝKÝZDERE KA;7151|K012|ÝYÝDERE KA;7151|K013|KALKANDERE KA;7151|K014|KÖY ARIZA KA;7151|K015|PAZAR KA;7151|K016|ÞEHÝR ARIZA KA;7151|K017|Trafo Bakým KA;7151|M001|19.GR.02.MARMRÝS;7151|M002|20.GR.01.OKAN;7151|M003|20.GR.02.KORDON;7151|M004|22.53.EMANET;7151|M005|21.GR.01.KAMERÝM;"
        s = s & "7151|M006|Kullaným dýþý;7151|M007|21.KET1.SEKMAN;7151|M008|21.KET2.TUGAY;7151|M009|21.KET4.DÝLMAÇ;7151|M010|21.53.AYKOL.E06;7151|M011|22.53.KET.01;7151|M012|22.53.KET.02;7151|M014|22.GR.01.KORD(A);7151|M015|22.GR.02.ENER(A);7151|M016|22.GR.03.TUGA(A);7151|M017|22.GR.04.DÝLM(A);7151|M018|22.GR.05.ENER(A);7151|M019|22.KET6.ENER(A);7151|M020|23.GR.01.ENER;7151|M021|23.GR.02.ENER;7151|M022|23.KET.4.DÝLMAÇ;7151|M023|23.KET.1.TUGAY;7151|M024|23.KET.2.ENER;7151|M027|13.GR.01 MERCAN;7151|M028|13.KET.DOÐUSU;7151|M029|14.GR.01 KORDON;7151|M030|19.GR.01 TUGAY;7151|M031|19.KET1.SEKMAN;7151|M032|20.53.BAÐL.ADA;7151|M033|20.GR.02 KORD(Y);7151|M034|20.KET1.SEKMAN;7151|M035|20.KET2.TÜRKMEN;"
        s = s & "7151|M036|21.53.AYKOL.E01;7151|M037|21.53.AYKOL.E02;7151|M038|21.53.AYKOL.E03;7151|M039|21.53.AYKOL.E04;7151|M040|21.53.AYKOL.E05;7151|M041|21.GR.01.KAME(Y);7151|M042|21.GR.02.KAME(Y);7151|M043|21.KET1.SEKMA(Y);7151|M044|21.KET2.TUGAY(Y);7151|M045|21.KET4.DÝLM(Y);7151|M046|21.KET5.TUG(Y);7151|M047|22.GR.01.KORDO(Y;7151|M048|22.GR.02.ENER(Y);7151|M049|22.GR.03.TUGA(Y);7151|M050|22.GR.05.ENER(Y);7151|M051|22.KET2.TUGAY(Y);7151|M052|22.KET3.TUYLU(Y);7151|M053|22.KET6.ENER(Y);7151|M054|21.53.AYKOL;7151|M055|13.BAÐL.ERSAN;7151|M056|13.GR.01 MERCAN;7151|M057|13.GR.02 AKSU;7151|M058|13.GR.03 ELKAR;7151|M059|13.BAÐL.ERSAN;7151|M060|22.GR04.DÝLMAÇ;7151|M096|24.KET.01 KORDON;7151|M097|24.KET.02 TUGAY;"
        s = s & "7151|M098|24.GR.01 TUGAY;7151|M142|23.KET.08 TUGAY;7151|M143|23.KET.03 KORDON;7151|M144|22.GR.05.ENER;7151|M145|21.KET.1 SEKMAN2;7151|P001|12.08.OSOS;7151|P002|Kullaným dýþý;7151|P003|17.53.SCADA;7151|P004|KOMP.53.YEDEK;7151|PB01|PB MERKEZ KA;7151|PB02|PLANLIBAKIM S3-1;7151|S001|S.AYAR MASASI D.;7152|B201|Kullaným dýþý.;7152|B202|Kullaným dýþý;7152|B203|Kullaným dýþý;7152|B204|Kullaným dýþý;7152|B205|Kullaným dýþý;7152|B206|Kullaným dýþý;7152|BO01|18.OB.PNRALTI;7152|BO02|Kullaným Dýþý;7152|BO03|21.61.OB.01;7152|BO04|19.OB.PNRALTI;7152|BO05|20.OB.PNARALT;7152|BO06|21.61.OB.02;7152|BO07|22.61.OB.01;7152|BO08|22.61.OB.02;7152|BO09|23.61.OB.01;7152|D001|Ýþletme Ambarý;7152|D002|Yatýrým Ambarý;"
        s = s & "7152|D003|Müþ.Bað.Ambarý;7152|D004|Atýl Malzeme Amb;7152|D006|Arýza Mlz. Ambar;7152|D007|Demontaj Ambarý;7152|D008|Hurda Ambarý;7152|D009|ÝSG-HIRDAVAT-ÞEB;7152|D010|ÝSG-HIRDAVAT-MÜÞ;7152|E001|BAÐLANTI Ekip1;7152|E002|BAÐLANTI Ekip2;7152|E003|BAÐLANTI Ekip3;7152|E004|BAÐLANTI Ekip4;7152|E005|BAÐLANTI Ekip5;7152|E006|BAÐLANTI Ekip6;7152|E007|BAÐLANTI Ekip7;7152|E008|BAÐLANTI Ekip8;7152|EM21|21.61.EMANET;7152|EM23|23.61.EMANET;7152|EM24|24.61.EMANET;7152|G001|Þirket-Ýl Müd.Dp;7152|G002|ÝDARÝ ÝÞ-FÝLO-BT;7152|G003|Ýstanbul Deposu;7152|G004|EDAÞ EMANET ÞEB;7152|G005|EMANET DEPO YTRM;7152|G006|AFAD;7152|K001|BAÐLANTI TES. KA;7152|K002|AKÇAABAT KA;7152|K003|ARAKLI KA;7152|K004|ARSÝN KA;"
        s = s & "7152|K005|BEÞÝKDÜZÜ KA;7152|K006|ÇARÞIBAÞI KA;7152|K007|ÇAYKARA KA;7152|K008|DERNEKPAZARI KA;7152|K009|DÜZKÖY KA;7152|K010|HAT BAKIM KA;7152|K011|HAYRAT KA;7152|K012|KÖPRÜBAÞI KA;7152|K013|KÖY ARIZA KA;7152|K014|MAÇKA KA;7152|K015|OF KA;7152|K016|SÜRMENE KA;7152|K017|ÞALPAZARI KA;7152|K018|ÞEHÝR ARIZA KA;7152|K019|TONYA KA;7152|K020|TRAFO BAKIM KA;7152|K021|VAKFIKEBÝR KA;7152|K022|YOMRA KA;7152|M001|17.61.GR.04.PINA;7152|M002|Kullaným dýþý;7152|M003|18.GR.03 DEMÝRAY;7152|M004|18.KET.1.SEKMAN;7152|M005|Kullaným dýþý;7152|M007|19.GR.02.ÇINARAL;7152|M008|19.GR.02.ENGÝND.;7152|M009|19.GR.03.ÇINARAL;7152|M010|19.GR.04.TEKÝNAL;7152|M011|Kullaným dýþý;7152|M012|Kullaným dýþý;"
        s = s & "7152|M013|19.KET.3.PINARAL;7152|M014|20.GR.01.KAMERÝM;7152|M015|20.GR.02.PINARAL;7152|M016|20.GR.03.PINARAL;7152|M017|20.GR.04.TÜRKMEN;7152|M018|20.KET.2.TEKINAL;7152|M019|20.KET.3.DEMIR;7152|M020|21.61.GR.02.PINA;7152|M021|Kullaným dýþý;7152|M022|Kullaným dýþý;7152|M026|Kullaným dýþý;7152|M027|21.KET.1.EKÝM;7152|M028|21.KET2.TEKÝNAL;7152|M029|21.KET3.DEMÝR;7152|M033|Kullaným dýþý;7152|M035|Kullaným dýþý;7152|M036|22.GR.03.ORÝM;7152|M038|Kullaným dýþý;7152|M039|22.KET.3 HASANOK;7152|M040|22.KET.4 ORÝM;7152|M041|22.KET.5.KUBAY;7152|M042|23.GR.02.FÝBER;7152|M043|23.GR.03.FÝBER;7152|M044|23.GR.04.TÜRKMEN;7152|M045|23.GR.05.KUBAY;7152|M046|23.GR.07.KUBAY;7152|M047|23.61.GR-08;"
        s = s & "7152|M048|23.GR.10.ENGÝNDE;7152|M049|23.GR.12.DEMÝR;7152|M050|23.KET.2.ERKE;7152|M051|23.KET.3.ORÝM;7152|M052|23.KET.4.ERKE;7152|M053|23.KET.5.KUBAY;7152|M054|23.GR.11 ORÝM;7152|M055|12.GR.03.DOÐUSU;7152|M056|12.KET Doðusu;7152|M057|19.GR.01PINARALT;7152|M058|20.KET1.EKÝM;7152|M059|21.61.AYKOL;7152|M060|21.61.AYKOL.E01;7152|M061|21.61.AYKOL.E02;7152|M062|21.61.AYKOL.E03;7152|M063|21.61.AYKOL.E04;7152|M064|21.61.AYKOL.E05;7152|M065|21.61.AYKOL.E06;7152|M067|21.61.AYKOL.E07;7152|M068|21.61.AYKOL.E08;7152|M069|21.61.AYKOL.E09;7152|M070|21.61.AYKOL.E10;7152|M071|21.61.AYKOL.E11;7152|M072|21.61.AYKOL.E12;7152|M077|22.61.GR.01(S);7152|M078|22.61.GR.02(S);7152|M079|22.61.GR.03(S);"
        s = s & "7152|M080|22.61.KET.01(S);7152|M081|22.61.KET.02(S);7152|M082|22.61.KET.04(S);7152|M083|22.61.KET.05(S);7152|M084|22.61.KET/06;7152|M085|22.GR.01.ERKE(A);7152|M086|22.GR.02.DEMÝ(A);7152|M089|22.KET.5 KUBAY;7152|M090|21.61.AYKOL.ÖZAY;7152|M091|21.61.AYKOL.E13;7152|M092|23.GR.13 PINARAL;7152|M096|24.KET.01 DEMÝR;7152|M097|24.KET.02 ONENER;7152|M098|24.KET.03 KUBAY;7152|M099|21.GR.04.KUBAY;7152|P001|12.08.OSOS;7152|P002|15.GUC KALITE IZ;7152|P003|17.61.SCADA;7152|P004|KOMP.61.YEDEK;7152|PB01|PB OF KA;7152|PB02|PB SÜRMENE KA;7152|PB03|PB ARSÝN KA;7152|PB04|PB MRKZ ÞEHÝR KA;7152|PB05|PB AKÇAABAT KA;7152|PB06|PB VAKFIKEBÝR KA;7152|PB07|PLANLIBAKIM S3-1;7152|S001|S.AYAR MASASI D.;"
        s = s & "7153|B201|Kullaným dýþý;7153|BO01|15.29.OB.01;7153|BO02|23.29.OB.01;7153|D001|Ýþletme Ambarý;7153|D002|Yatýrým Ambarý;7153|D003|Müþ.Bað.Ambarý;7153|D004|Atýl Malzeme Amb;7153|D006|Arýza Mlz. Ambar;7153|D007|Demontaj Ambarý;7153|D008|Hurda Ambarý;7153|D009|ÝSG-HIRDAVAT-ÞEB;7153|D010|ÝSG-HIRDAVAT-MÜÞ;7153|E001|BAÐLANTI Ekip1;7153|E002|BAÐLANTI Ekip2;7153|E003|BAÐLANTI Ekip3;7153|E004|BAÐLANTI Ekip4;7153|EM23|23.29.EMANET;7153|EM24|24.29.EMANET;7153|G001|Ýl Müdürlüðü Dp;7153|G002|ÝDARÝ ÝÞ-FÝLO-BT;7153|G004|EDAÞ EMANET ÞEB;7153|G005|EMANET DEPO YTRM;7153|K001|BAÐLANTI TES. KA;7153|K002|HAT BAKIM KA;7153|K003|KELKÝT KA;7153|K004|KÖSE KA;7153|K005|KÜRTÜN KA;7153|K006|MERKEZ KA;7153|K007|ÞÝRAN KA;"
        s = s & "7153|K008|TORUL KA;7153|K009|TRAFO BAKIM KA;7153|K010|MERKEZ BAGLANTI;7153|M001|22.KET.1 HASANOK;7153|M002|23.GR.01.ERKYAP;7153|M003|23.KET.1.ERKYAP;7153|M004|23.GR.01 ERKYAP;7153|M005|23.29 KET.02;7153|M007|20.29.GR.02 TÜRK;7153|M008|21.29.AYKOL.E01;7153|M009|21.29.AYKOL.E02;7153|M010|22.KET1.HASAN(Y);7153|M011|22.GR.01.ÝÞBÝ(A);7153|M012|21.29.AYKOL;7153|M013|21.29.AYKOL.E03;7153|M096|24.KET.01 ENGÝND;7153|P001|12.08.OSOS;7153|P002|Kullaným dýþý;7153|P003|17.29.SCADA;7153|P004|KOMP.29.YEDEK;7153|PB01|PB MERKEZ KA;7153|PB02|PLANLIBAKIM S3-1;7153|S001|S.AYAR MASASI D.;7154|B201|Kullaným dýþý;7154|B202|Kullaným dýþý;7154|B203|Kullaným dýþý;7154|B204|Kullaným dýþý;7154|B205|Kullaným dýþý;"
        s = s & "7154|B206|Kullaným dýþý;7154|B301|Kullaným dýþý;7154|BO01|13.01.NETKAYA;7154|BO02|14.28.OB.01;7154|BO03|19.28.OB.3SEKMAN;7154|BO04|Kullaným dýþý;7154|BO05|22.28.OB.01 KAME;7154|BO06|23.OB.01.AHMET;7154|BO08|23.28.HEYELAN.OB;7154|D001|Ýþletme Ambarý;7154|D002|Yatýrým Ambarý;7154|D003|Müþ.Bað.Ambarý;7154|D004|Atýl Malzeme Amb;7154|D006|Arýza Mlz. Ambar;7154|D007|Demontaj Ambarý;7154|D008|Hurda Ambarý;7154|D009|ÝSG-HIRDAVAT-ÞEB;7154|D010|ÝSG-HIRDAVAT-MÜÞ;7154|E001|BAÐLANTI Ekip1;7154|E002|BAÐLANTI Ekip2;7154|E003|BAÐLANTI Ekip3;7154|E004|BAÐLANTI Ekip4;7154|E005|BAÐLANTI Ekip5;7154|E006|BAÐLANTI Ekip6;7154|E007|BAÐLANTI Ekip7;7154|EM23|23.28.EMANET;7154|EM24|24.28.EMANET;"
        s = s & "7154|G001|Ýl Müdürlüðü Dp;7154|G002|ÝDARÝ ÝÞ-FÝLO-BT;7154|G004|EDAÞ EMANET ÞEB;7154|G005|EMANET DEPO YTRM;7154|K001|BAÐLANTI TES. KA;7154|K002|ALUCRA KA;7154|K003|BULANCAK KA;7154|K004|ÇAMOLUK KA;7154|K005|ÇANAKÇI KA;7154|K006|DERELÝ KA;7154|K007|DOÐANKENT KA;7154|K008|ESPÝYE KA;7154|K009|EYNESÝL KA;7154|K010|GÖRELE KA;7154|K011|GÜCE KA;7154|K012|KEÞAP KA;7154|K013|KÖY ARIZA KA;7154|K014|PÝRAZÝZ KA;7154|K015|Þ.KARAHÝSAR KA;7154|K016|ÞEHÝR ARIZA KA;7154|K017|TÝREBOLU KA;7154|K018|TRAFO BAKIM KA;7154|K019|YAÐLIDERE KA;7154|M001|07.YAP SANEL;7154|M002|20.GR.03 TEKÝNAL;7154|M003|21.GR.03.ÝÞBÝL;7154|M005|Kullaným dýþý;7154|M006|Kullaným dýþý;7154|M007|22.GR.03.ON E(A);7154|M008|Kullaným dýþý;"
        s = s & "7154|M009|Kullaným dýþý;7154|M010|23.GR.01.KURALAY;7154|M011|23.GR.03.KURALAY;7154|M012|23.GR.04.ELSAN;7154|M013|23.KET.1 KURALAY;7154|M014|23.KET.2 KURALAY;7154|M015|Kullaným dýþý;7154|M016|Kullaným dýþý;7154|M017|Kullaným dýþý;7154|M018|13.GR.01_DEMÝRYÜ;7154|M019|20.GR.02 ELSAN;7154|M020|20.GR.03 TEKÝ(Y);7154|M021|21.28.AYKOL.E01;7154|M022|21.28.AYKOL.E02;7154|M023|21.28.AYKOL.E03;7154|M024|21.28.AYKOL.E04;7154|M025|21.28.AYKOL.E05;7154|M026|21.28.AYKOL.E06;7154|M027|21.GR.01 ONEN(Y);7154|M028|21.GR.03.ALÖZ;7154|M029|22.GR.03 ONEN(Y);7154|M030|POWERLINE 477;7154|M031|21.28.AYKOL;7154|M096|24.28.KET.01 KUR;7154|M097|24.GR.01 KURALAY;7154|P001|12.08.OSOS;7154|P002|15.GUC KALITE IZ;"
        s = s & "7154|P003|17.28.SCADA;7154|P004|KOMP.28.YEDEK;7154|PB01|PB Þ.KARAHÝSR KA;7154|PB02|PB MERKEZ KA;7154|PB03|PB BULANCAK KA;7154|PB04|PB TÝREBOLU KA;7154|PB05|PB ESPÝYE KA;7154|PB06|PB GÖRELE KA;7154|PB07|PLANLIBAKIM S3-1;7154|S001|S.AYAR MASASI D.;7155|B301|Kullaným dýþý;7155|BO01|21.08.OB.01;7155|BO02|22.08.OB.01GUKAS;7155|BO03|23.OB.01.ENER;7155|D001|Ýþletme Ambarý;7155|D002|Yatýrým Ambarý;7155|D003|Müþ.Bað.Ambarý;7155|D004|Atýl Malzeme Amb;7155|D006|Arýza Mlz. Ambar;7155|D007|Demontaj Ambarý;7155|D008|Hurda Ambarý;7155|D009|ÝSG-HIRDAVAT-ÞEB;7155|D010|ÝSG-HIRDAVAT-MÜÞ;7155|E001|BAÐLANTI EKÝP1;7155|E002|BAÐLANTI EKÝP2;7155|E003|BAÐLANTI EKÝP3;7155|E004|BAÐLANTI Ekip4;7155|EM14|14.08.EMANET;"
        s = s & "7155|EM21|21.08.EMANET;7155|EM22|22.08.EMANET;7155|EM23|23.08.EMANET;7155|EM24|24.08.EMANET;7155|G001|Ýl Müdürlüðü Dp;7155|G002|ÝDARÝ ÝÞ-FÝLO-BT;7155|G004|EDAÞ EMANET ÞEB;7155|G005|EMANET DEPO YTRM;7155|K001|Ardanuç KA;7155|K002|Arhavi KA;7155|K003|Borçka KA;7155|K004|Hopa KA;7155|K005|Kemalpaþa KA;7155|K006|Merkez KA;7155|K007|Murgul KA;7155|K008|Þavþat KA;7155|K009|Trafo Bakým KA;7155|K010|Yusufeli KA;7155|K011|BAÐLANTI TES KA;7155|K012|MERKEZ BAGLANTI;7155|M001|20.GR.01.ASÝÞ;7155|M002|20.GR.02 KAMERÝM;7155|M003|Kullaným dýþý;7155|M004|Kullaným dýþý;7155|M005|Kullaným dýþý;7155|M006|22.GR.01.F.TÜRK;7155|M007|22.KET.1 TÜRKMEN;7155|M008|23.KET.01 GÜKAÞ;7155|M011|17.GR.02NETKAYA;"
        s = s & "7155|M012|20.GR.01 ASÝÞ(Y);7155|M013|20.GR.02 KAME(Y);7155|M014|21.08.AYKOL;7155|M015|21.GR.01.KAME(Y);7155|M016|22.GR.01.FÝBE(S);7155|M017|22.KET.1 TÜRK(S);7155|M018|21.08.AYKOL.E01;7155|M019|21.08.AYKOL.E02;7155|M020|23.08.GR.01;7155|M021|23.GR.02 KABADAY;7155|M096|24.KET.01.JYN;7155|M114|kullaným dýþý;7155|P001|12.08.OSOS;7155|P002|Kullaným dýþý;7155|P003|17.08.SCADA;7155|P004|KOMP.08.YEDEK;7155|PB01|PB MERKEZ KA;7155|PB02|PLANLIBAKIM S3-1;7155|S001|S.AYAR MASASI D.;7155|S002|Kullaným Dýþý;7155|S003|Kullaným Dýþý;7155|S004|Kullaným Dýþý;7155|S005|Kullaným Dýþý;7155|S006|Kullaným Dýþý;"
        For Each p In Split(s, ";")
            If Len(p) > 0 Then
                x = Split(p, "|")
                d(x(0) & "|" & x(1)) = x(2)
            End If
        Next p
    End If
    If d.Exists(Trim$(uy) & "|" & UCase$(Trim$(depo))) Then MS_DepoAdi = d(Trim$(uy) & "|" & UCase$(Trim$(depo)))
End Function
'==========================================================
' YENI MOTOR: IQ09 (seri numarali stok) + MB52 (toplu stok ve deger)
' Liste SAP'den "Tablo hesaplama" (&XXL) ile Excel'e aktarilip okunur.
'==========================================================
Private Function MS_GridBekle(ByVal ses As Object, ByVal saniye As Double) As Object
    Dim bas As Double, g As Object
    bas = Timer
    Do
        Set g = Nothing
        On Error Resume Next
        Set g = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
        On Error GoTo 0
        If Not g Is Nothing Then
            If Not ses.Busy Then Set MS_GridBekle = g: Exit Function
        End If
        If MS_Gecen(bas) > saniye Then Exit Function
        DoEvents: MS_Sleep 100
    Loop
End Function

Private Function MS_IlkKolon(ByVal alv As Object) As String
    On Error Resume Next
    MS_IlkKolon = CStr(alv.ColumnOrder.Item(0))
    On Error GoTo 0
End Function

' Duzen secimi: aday adlardan biri bulunursa o, bulunamazsa SAP'nin varsayilan secili satiri.
Private Sub MS_DuzenSec(ByVal ses As Object, ByVal alv As Object, ByVal adaylar As Variant, ByVal etiket As String)
    Dim g As Object, r As Long, ad As String, tx As String, secim As Long, kol As String, y As Variant
    On Error Resume Next
    kol = MS_IlkKolon(alv)
    alv.SetCurrentCell -1, kol
    alv.SelectColumn kol
    alv.ContextMenu
    alv.SelectContextMenuItem "&COL0"
    ses.FindById("wnd[1]/tbar[0]/btn[36]").Press
    Set g = ses.FindById("wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R9/ssubSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell", False)
    If g Is Nothing Then
        Err.Clear
        If Not ses.FindById("wnd[1]", False) Is Nothing Then ses.FindById("wnd[1]/tbar[0]/btn[12]").Press
        ses.FindById("wnd[0]/tbar[1]/btn[33]").Press
        Set g = ses.FindById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell", False)
    End If
    If g Is Nothing Then
        MS_Teshis etiket & " düzen penceresi bulunamadý"
        If Not ses.FindById("wnd[1]", False) Is Nothing Then ses.FindById("wnd[1]/tbar[0]/btn[12]").Press
        Exit Sub
    End If
    secim = -1
    For r = 0 To CLng(g.RowCount) - 1
        ad = vbNullString: tx = vbNullString
        ad = Trim$(CStr(g.GetCellValue(r, "VARIANT"))): tx = Trim$(CStr(g.GetCellValue(r, "TEXT")))
        MS_Teshis etiket & " düzen " & CStr(r) & ": " & ad & " | " & tx
        For Each y In adaylar
            If secim < 0 And (StrComp(ad, CStr(y), vbTextCompare) = 0 Or StrComp(tx, CStr(y), vbTextCompare) = 0) Then secim = r
        Next y
    Next r
    If secim >= 0 Then
        g.SetCurrentCell secim, "TEXT"
        g.SelectedRows = CStr(secim)
        MS_Teshis etiket & " düzen seçildi: " & CStr(secim)
    Else
        g.CurrentCellColumn = "TEXT"
        MS_Teshis etiket & " düzen: SAP varsayýlan satýrý"
    End If
    g.ClickCurrentCell
    MS_Sleep 300
    If Not ses.FindById("wnd[1]", False) Is Nothing Then ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
    On Error GoTo 0
End Sub

' SAP'nin actigi disa aktarim kitaplarini (baska Excel penceresinde olsa da) kapatir, dosyalari siler.
Private Sub MS_DisaKapat()
    Dim tur As Long, p As Variant, w As Object, uyg As Object, wb As Workbook, kalan As Long
    If mDisaDosyalar Is Nothing Then Exit Sub
    On Error Resume Next
    For tur = 1 To 8
        kalan = 0
        For Each p In mDisaDosyalar
            Set w = Nothing: Set uyg = Nothing
            For Each wb In application.Workbooks
                If StrComp(wb.FullName, CStr(p), vbTextCompare) = 0 Then wb.Close SaveChanges:=False
            Next wb
            If Len(Dir$(CStr(p))) > 0 Then
                Err.Clear
                Kill CStr(p)
                If Len(Dir$(CStr(p))) > 0 Then kalan = kalan + 1
            End If
        Next p
        If kalan = 0 And tur >= 3 Then Exit For
        DoEvents: MS_Sleep 700
    Next tur
    For Each wb In application.Workbooks
        If wb.Name Like "EXPORT_*.*" Then wb.Close SaveChanges:=False
    Next wb
    ThisWorkbook.Activate
    Set mDisaDosyalar = Nothing
    On Error GoTo 0
End Sub

' Guncelleme sirasinda SAP'nin baslattigi Excel islemleri kapatilir (AKSORA'nin kendi islemi haric).
Private Sub MS_ExcelTemizle()
    Dim wmi As Object, p As Object, benim As Long, olusma As String, adet As Long
    If Len(mBasZamanMetin) = 0 Then Exit Sub
    On Error Resume Next
    benim = MS_PID()
    Set wmi = GetObject("winmgmts:\\.\root\cimv2")
    If wmi Is Nothing Then Exit Sub
    For Each p In wmi.ExecQuery("SELECT ProcessId, CreationDate FROM Win32_Process WHERE Name = 'EXCEL.EXE'")
        olusma = Left$(CStr(p.CreationDate), 14)
        If CLng(p.ProcessId) <> benim And Len(olusma) = 14 And olusma >= mBasZamanMetin Then
            p.Terminate
            adet = adet + 1
        End If
    Next p
    If adet > 0 Then MS_Teshis "Kapatýlan SAP Excel iþlemi: " & CStr(adet)
    On Error GoTo 0
End Sub

' SAP'nin disa aktarirken kaydettigi en yeni Excel dosyasini bulur (baska Excel'de acilmis olsa da).
Private Function MS_YeniDosyaBul(ByVal bas As Date) As String
    Dim klasorler As Variant, k As Variant, uzantilar As Variant, u As Variant, f As String
    Dim yol As String, enYeni As Date, t As Date
    klasorler = Array("C:\TEMP", "C:\TMP", Environ$("TEMP"), Environ$("TMP"), _
        Environ$("APPDATA") & "\SAP\SAP GUI\tmp", Environ$("LOCALAPPDATA") & "\SAP\SAP GUI\tmp", _
        Environ$("USERPROFILE") & "\Documents\SAP\SAP GUI", Environ$("USERPROFILE") & "\Documents", _
        Environ$("USERPROFILE") & "\OneDrive\Documents\SAP\SAP GUI", Environ$("USERPROFILE") & "\Downloads")
    uzantilar = Array("*.xlsx", "*.XLSX", "*.xls", "*.MHTML", "*.mhtml", "*.xml")
    On Error Resume Next
    For Each k In klasorler
        If Len(k) > 0 Then
            For Each u In uzantilar
                f = Dir$(k & "\" & u)
                Do While Len(f) > 0
                    t = 0: t = FileDateTime(k & "\" & f)
                    If t >= bas And t > enYeni Then enYeni = t: yol = k & "\" & f
                    f = Dir$()
                Loop
            Next u
        End If
    Next k
    On Error GoTo 0
    MS_YeniDosyaBul = yol
End Function

' Listeyi "Yerel dosya" (sekmeli metin) olarak C:\TEMP'e kaydedip okur.
' SAP bu yontemde program acmaz; Excel islemi olusmaz. Basarisizsa Empty doner.
Private Function MS_AlvMetin(ByVal ses As Object, ByVal alv As Object, ByVal etiket As String) As Variant
    Dim kol As String, pen As Object, yolAlan As Object, o As Object, rb As Object, radyolar As Collection
    Dim klasor As String, dosya As String, tam As String, bas As Double, boy As Long, eskiBoy As Long
    Dim metin As String, satirlar As Variant, alanlar As Variant, r As Long, c As Long, n As Long, enCok As Long
    Dim v() As Variant, a As String, ofs As Long
    On Error GoTo Hata
    klasor = "C:\TEMP"
    If Len(Dir$(klasor, vbDirectory)) = 0 Then klasor = Environ$("TEMP")
    dosya = "AKSORA_" & etiket & "_" & Format$(Now, "yyyymmddhhnnss") & ".txt"
    tam = klasor & "\" & dosya
    kol = MS_IlkKolon(alv)
    alv.SetCurrentCell -1, kol
    alv.SelectColumn kol
    alv.ContextMenu
    alv.SelectContextMenuItem "&PC"
    Set pen = MS_PencereBekle(ses, 10)
    If pen Is Nothing Then MS_Teshis etiket & " yerel dosya penceresi açýlmadý": GoTo Temizle
    MS_Teshis etiket & " yerel dosya penceresi: " & pen.Text
    Set yolAlan = MS_AlanBul(pen, "DY_PATH")
    If yolAlan Is Nothing Then
        Set radyolar = New Collection
        MS_Topla pen, "GuiRadioButton", radyolar
        Set o = Nothing
        For Each rb In radyolar
            a = AKFIX_NORMAL(CStr(rb.Text))
            If InStr(a, "TAB") > 0 Or InStr(a, "SEKME") > 0 Then Set o = rb: Exit For
        Next rb
        If o Is Nothing And radyolar.Count >= 2 Then Set o = radyolar(2)
        If o Is Nothing Then MS_Agac pen, 0: GoTo Temizle
        MS_Teshis etiket & " biçim: " & o.Text
        o.Select
        ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
        Set pen = MS_PencereBekle(ses, 10)
        If pen Is Nothing Then MS_Teshis etiket & " dosya adý penceresi açýlmadý": GoTo Temizle
        Set yolAlan = MS_AlanBul(pen, "DY_PATH")
    End If
    If yolAlan Is Nothing Then MS_Agac pen, 0: GoTo Temizle
    yolAlan.Text = klasor & "\"
    Set o = MS_AlanBul(pen, "DY_FILENAME")
    If o Is Nothing Then GoTo Temizle
    o.Text = dosya
    Set o = MS_AlanBul(pen, "DY_FILE_ENCODING")
    If Not o Is Nothing Then o.Text = "4110"
    ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
    bas = Timer: eskiBoy = -1
    Do
        DoEvents: MS_Sleep 250
        If Len(Dir$(tam)) > 0 Then
            boy = FileLen(tam)
            If boy > 0 And boy = eskiBoy And Not ses.Busy Then Exit Do
            eskiBoy = boy
        End If
        If MS_Gecen(bas) > 180 Then MS_Teshis etiket & " metin dosyasý oluþmadý: " & tam: GoTo Temizle
    Loop
    If MS_DosyaBom(tam) = "utf-8" Or MS_DosyaBom(tam) = vbNullString Then
        metin = MS_DosyaMetni(tam, "utf-8")
        If InStr(metin, ChrW(&HFFFD)) > 0 Then metin = MS_DosyaMetni(tam, "windows-1254")
    Else
        metin = MS_DosyaMetni(tam, MS_DosyaBom(tam))
    End If
    metin = Replace(Replace(metin, vbCrLf, vbLf), vbCr, vbLf)
    If Left$(metin, 1) = ChrW(&HFEFF) Then metin = Mid$(metin, 2)
    satirlar = Split(metin, vbLf)
    For r = 0 To UBound(satirlar)
        c = UBound(Split(satirlar(r), vbTab)) + 1
        If c > enCok Then enCok = c
    Next r
    If enCok < 4 Then MS_Teshis etiket & " metin dosyasý sekmeli deðil": GoTo Temizle
    ReDim v(1 To UBound(satirlar) + 1, 1 To enCok)
    For r = 0 To UBound(satirlar)
        If Len(Trim$(Replace(satirlar(r), vbTab, ""))) > 0 Then
            n = n + 1
            alanlar = Split(satirlar(r), vbTab)
            For c = 0 To UBound(alanlar)
                v(n, c + 1) = Trim$(alanlar(c))
            Next c
        End If
    Next r
    ' Basta bos sutun varsa kaydirilir.
    ofs = 1
    For r = 1 To n
        If Len(CStr(v(r, 1))) > 0 Then ofs = 0: Exit For
    Next r
    If ofs = 1 Then
        For r = 1 To n
            For c = 1 To enCok - 1: v(r, c) = v(r, c + 1): Next c
            v(r, enCok) = Empty
        Next r
    End If
    MS_Teshis etiket & " yerel dosya okundu: " & CStr(n) & " satýr, " & AKSORA_SAYI_METNI(MS_Gecen(bas), "0.0") & " sn"
    On Error Resume Next
    Kill tam
    On Error GoTo 0
    MS_AlvMetin = MS_Kirp(v, n, enCok)
    Exit Function
Hata:
    MS_Teshis etiket & " yerel dosya hatasý: " & Err.Description
    Resume Temizle
Temizle:
    On Error Resume Next
    If Not ses.FindById("wnd[1]", False) Is Nothing Then ses.FindById("wnd[1]/tbar[0]/btn[12]").Press
    If Len(Dir$(tam)) > 0 Then Kill tam
    MS_AlvMetin = Empty
End Function

' Listeyi Tablo hesaplama (&XXL) ile Excel'e aktarir; acilan kitabin verisini dondurur.
Private Function MS_AlvExcel(ByVal ses As Object, ByVal alv As Object, ByVal etiket As String) As Variant
    Dim eski As Object, wb As Workbook, sh As Worksheet, en As Worksheet, bas As Double, kol As String
    Dim w As Object, basildi As Long, basZaman As Date, dosya As String, boy As Long, eskiBoy As Long
    Dim dis As Object, kendimiz As Boolean, kopya As String
    On Error GoTo Hata
    basZaman = Now - TimeSerial(0, 0, 5)
    Set eski = MS_AcikKitaplar()
    kol = MS_IlkKolon(alv)
    alv.SetCurrentCell -1, kol
    alv.SelectColumn kol
    alv.ContextMenu
    alv.SelectContextMenuItem "&XXL"
    bas = Timer
    Do
        DoEvents
        For Each wb In application.Workbooks
            If Not wb Is ThisWorkbook Then
                If Not eski.Exists(wb.Name) Then GoTo Bulundu
            End If
        Next wb
        Set wb = Nothing
        If basildi < 4 Then
            Set w = Nothing
            On Error Resume Next
            Set w = ses.FindById("wnd[1]", False)
            If Not w Is Nothing Then
                If Not ses.Busy Then
                    MS_Teshis etiket & " aktarým penceresi: " & w.Text
                    If basildi = 0 Then ses.FindById("wnd[1]/tbar[0]/btn[20]").Press
                    Err.Clear
                    ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
                    basildi = basildi + 1
                End If
            End If
            On Error GoTo Hata
        End If
        ' Liste baska bir Excel'de acildiysa dosyasi diskte bulunur ve buradan okunur.
        If basildi > 0 And Not ses.Busy Then
            dosya = MS_YeniDosyaBul(basZaman)
            If Len(dosya) > 0 Then
                boy = 0
                On Error Resume Next
                boy = FileLen(dosya)
                On Error GoTo Hata
                If boy > 0 And boy = eskiBoy Then
                    MS_Teshis etiket & " dosya bulundu: " & dosya & " (" & CStr(boy) & " bayt)"
                    If mDisaDosyalar Is Nothing Then Set mDisaDosyalar = New Collection
                    mDisaDosyalar.Add dosya
                    ' SAP'nin actigi Excel dosyayi kilitler; once o kapatilir, sonra salt okunur acilir.
                    MS_ExcelTemizle
                    MS_Sleep 700
                    application.DisplayAlerts = False
                    Set wb = Workbooks.Open(Filename:=dosya, ReadOnly:=True, UpdateLinks:=0, Notify:=False, AddToMru:=False)
                    application.DisplayAlerts = True
                    kendimiz = True
                    GoTo Bulundu
                End If
                eskiBoy = boy
            End If
        End If
        If MS_Gecen(bas) > 180 Then
            MS_Teshis etiket & ": Excel'e aktarýlan liste 180 sn içinde açýlmadý (dosya: " & dosya & ")"
            Exit Function
        End If
        MS_Sleep 400
    Loop
Bulundu:
    If Not kendimiz Then
        If mDisaDosyalar Is Nothing Then Set mDisaDosyalar = New Collection
        mDisaDosyalar.Add wb.FullName
    End If
    For Each sh In wb.Worksheets
        If en Is Nothing Then
            Set en = sh
        ElseIf sh.UsedRange.Rows.Count > en.UsedRange.Rows.Count Then
            Set en = sh
        End If
    Next sh
    MS_AlvExcel = en.UsedRange.Value2
    MS_Teshis etiket & " Excel'e aktarýldý: " & CStr(en.UsedRange.Rows.Count) & " satýr, " & _
        AKSORA_SAYI_METNI(MS_Gecen(bas), "0.0") & " sn"
    wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Exit Function
Hata:
    MS_Teshis etiket & " Excel aktarým hatasý: " & Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
End Function

Private Function MS_BaslikSatiri(ByVal v As Variant) As Long
    Dim r As Long, c As Long, dolu As Long
    For r = 1 To application.Min(10, UBound(v, 1))
        dolu = 0
        For c = 1 To UBound(v, 2)
            If Len(MS_Metin(v(r, c))) > 0 Then dolu = dolu + 1
        Next c
        If dolu >= 4 Then MS_BaslikSatiri = r: Exit Function
    Next r
End Function

' tur: UY, DEPO, DEPOAD, MG, PARTI, MALZ, METIN, SERI, MIKTAR, BIRIM, DEGER
Private Function MS_KolonBul(ByVal v As Variant, ByVal h As Long, ByVal tur As String) As Long
    Dim c As Long, t As String
    For c = 1 To UBound(v, 2)
        t = MS_Norm(MS_Metin(v(h, c)))
        Select Case tur
            Case "UY": If t = "URETIMYERI" Or t = "UY" Or t = "URETIMYR" Then MS_KolonBul = c
            Case "DEPO": If t = "DEPOYERI" Or t = "DPYR" Or t = "DEPO" Then MS_KolonBul = c
            Case "DEPOAD": If t = "DEPOYERITANIMI" Or t = "DEPOFIKTIFGUMRUKTANIMI" Or t = "DEPOTANIMI" Then MS_KolonBul = c
            Case "MG": If t = "MALGRUBU" Then MS_KolonBul = c
            Case "PARTI": If t = "PARTI" Or t = "DEGERLEMETURU" Or t = "DRLTURU" Or t = "DEGERLTURU" Then MS_KolonBul = c
            Case "MALZ": If t = "MALZEME" Or t = "MALZEMENO" Then MS_KolonBul = c
            Case "METIN": If Left$(t, 14) = "MALZEMEKISAMET" Or t = "MALZEMETANIMI" Or t = "TANIM" Then MS_KolonBul = c
            Case "SERI": If Left$(t, 9) = "SERINUMAR" Or t = "SERINO" Or t = "SERNR" Then MS_KolonBul = c
            Case "MIKTAR": If InStr(t, "TAHDITSIZ") > 0 And InStr(t, "DEGER") = 0 Then MS_KolonBul = c
            Case "BIRIM": If t = "TEMELOLCUBIRIMI" Or t = "TOB" Or t = "BIRIM" Or t = "TEMELOLCUBRM" Then MS_KolonBul = c
            Case "DEGER": If InStr(t, "DEGER") > 0 And InStr(t, "DEGERLEME") = 0 And InStr(t, "TURU") = 0 Then MS_KolonBul = c
        End Select
        If MS_KolonBul > 0 Then Exit Function
    Next c
End Function

Private Sub MS_BaslikTeshis(ByVal v As Variant, ByVal h As Long, ByVal etiket As String)
    Dim c As Long, t As String
    For c = 1 To UBound(v, 2)
        t = t & " | " & MS_Metin(v(h, c))
    Next c
    MS_Teshis etiket & " baþlýklar:" & t
End Sub

' IQ09: UY, Depo, Malzeme, Metin, Seri, Parti
Private Function MS_IQ09_Oku(ByVal ses As Object) As Variant
    Dim alv As Object, v As Variant, h As Long, k(1 To 6) As Long, r As Long, n As Long, c As Long
    Dim sonuc() As Variant, depo As String, adlar As Variant
    MS_Teshis "---- " & Format$(Now, "dd.mm.yyyy hh:nn:ss") & " | IQ09 baþlýyor"
    ses.StartTransaction "IQ09"
    On Error Resume Next
    ses.FindById("wnd[0]/usr/ctxtDATUV").Text = vbNullString
    On Error GoTo 0
    ses.FindById("wnd[0]/usr/ctxtWERK-LOW").Text = "7155"
    ses.FindById("wnd[0]/tbar[1]/btn[8]").Press
    Set alv = MS_GridBekle(ses, 900)
    If alv Is Nothing Then Err.Raise vbObjectError + 1960, , "IQ09 sonuç listesi açýlmadý."
    MS_Teshis "IQ09 liste: " & CStr(alv.RowCount) & " satýr"
    MS_DuzenSec ses, alv, Array("/PANEL", "PANEL", "/AKSORA", "AKSORA"), "IQ09"
    Set alv = MS_GridBekle(ses, 60)
    MS_Durum "IQ09 listesi Excel'e aktarýlýyor..."
    v = MS_AlvExcel(ses, alv, "IQ09")
    If Not IsArray(v) Then Err.Raise vbObjectError + 1961, , "IQ09 listesi Excel'e aktarýlamadý. Teþhis: " & MS_TeshisYolu()
    h = MS_BaslikSatiri(v)
    If h = 0 Then Err.Raise vbObjectError + 1962, , "IQ09 listesinde baþlýk bulunamadý."
    MS_BaslikTeshis v, h, "IQ09"
    adlar = Array("UY", "DEPO", "MALZ", "METIN", "SERI", "PARTI")
    For c = 1 To 6
        k(c) = MS_KolonBul(v, h, CStr(adlar(c - 1)))
        If k(c) = 0 And c <> 1 And c <> 4 Then Err.Raise vbObjectError + 1963, , "IQ09 listesinde sütun bulunamadý: " & adlar(c - 1)
    Next c
    ReDim sonuc(1 To UBound(v, 1), 1 To 6)
    For r = h + 1 To UBound(v, 1)
        depo = UCase$(MS_Metin(v(r, k(2))))
        If Len(depo) > 0 And Len(MS_Metin(v(r, k(5)))) > 0 And Len(MS_Metin(v(r, k(3)))) > 0 Then
            n = n + 1
            For c = 1 To 6
                If k(c) > 0 Then sonuc(n, c) = MS_Metin(v(r, k(c))) Else sonuc(n, c) = vbNullString
            Next c
            sonuc(n, 2) = depo
        End If
    Next r
    MS_Teshis "IQ09 stoktaki seri numarasý: " & CStr(n)
    If n = 0 Then Err.Raise vbObjectError + 1964, , "IQ09 listesinde depoda seri numarasý bulunamadý."
    MS_IQ09_Oku = MS_Kirp(sonuc, n, 6)
End Function

' MB52: UY, Depo, DepoAd, MG, Parti, Malzeme, Metin, Miktar, Birim, Deger
Private Function MS_MB52_Oku(ByVal ses As Object) As Variant
    Dim alv As Object, v As Variant, h As Long, k(1 To 10) As Long, r As Long, n As Long, c As Long
    Dim sonuc() As Variant, adlar As Variant, konum As Variant
    MS_Teshis "---- " & Format$(Now, "dd.mm.yyyy hh:nn:ss") & " | MB52 baþlýyor"
    ses.StartTransaction "MB52"
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = "7155"
    On Error Resume Next
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString
    On Error GoTo 0
    ses.FindById("wnd[0]").SendVKey 8
    Set alv = MS_GridBekle(ses, 600)
    If alv Is Nothing Then Err.Raise vbObjectError + 1970, , "MB52 sonuç listesi açýlmadý."
    MS_DuzenSec ses, alv, Array("PANEL", "/PANEL", "MLZ STOK PANEL"), "MB52"
    Set alv = MS_GridBekle(ses, 60)
    MS_Durum "MB52 listesi Excel'e aktarýlýyor..."
    v = MS_AlvExcel(ses, alv, "MB52")
    If Not IsArray(v) Then Err.Raise vbObjectError + 1971, , "MB52 listesi Excel'e aktarýlamadý. Teþhis: " & MS_TeshisYolu()
    h = MS_BaslikSatiri(v)
    If h = 0 Then Err.Raise vbObjectError + 1972, , "MB52 listesinde baþlýk bulunamadý."
    MS_BaslikTeshis v, h, "MB52"
    adlar = Array("UY", "DEPO", "DEPOAD", "MG", "PARTI", "MALZ", "METIN", "MIKTAR", "BIRIM", "DEGER")
    konum = Array(1, 2, 3, 4, 6, 7, 8, 9, 10, 11)
    For c = 1 To 10
        k(c) = MS_KolonBul(v, h, CStr(adlar(c - 1)))
    Next c
    ' PANEL duzeninin bilinen sirasi yedek olarak kullanilir.
    If (k(2) = 0 Or k(6) = 0 Or k(8) = 0 Or k(10) = 0) And UBound(v, 2) >= 11 Then
        For c = 1 To 10: k(c) = konum(c - 1): Next c
        MS_Teshis "MB52 sütunlarý PANEL sýrasýyla eþlendi"
    End If
    For c = 1 To 10
        If k(c) = 0 And c <> 3 And c <> 7 And c <> 9 Then Err.Raise vbObjectError + 1973, , "MB52 listesinde sütun bulunamadý: " & adlar(c - 1)
    Next c
    ReDim sonuc(1 To UBound(v, 1), 1 To 10)
    For r = h + 1 To UBound(v, 1)
        If Len(MS_Metin(v(r, k(6)))) > 0 And Len(MS_Metin(v(r, k(2)))) > 0 Then
            If k(1) = 0 Or MS_Metin(v(r, k(1))) = "7155" Then
                n = n + 1
                For c = 1 To 10
                    If k(c) > 0 Then sonuc(n, c) = MS_Metin(v(r, k(c))) Else sonuc(n, c) = vbNullString
                Next c
                sonuc(n, 2) = UCase$(CStr(sonuc(n, 2)))
                sonuc(n, 8) = MS_SayiSifir(v(r, k(8)))
                sonuc(n, 10) = MS_SayiSifir(v(r, k(10)))
            End If
        End If
    Next r
    MS_Teshis "MB52 satýr: " & CStr(n)
    If n = 0 Then Err.Raise vbObjectError + 1974, , "MB52 listesinde 7155 stok satýrý bulunamadý."
    MS_MB52_Oku = MS_Kirp(sonuc, n, 10)
End Function

Private Function MS_SayiSifir(ByVal v As Variant) As Double
    Dim d As Double
    If VarType(v) = vbDouble Or VarType(v) = vbCurrency Or VarType(v) = vbLong Or VarType(v) = vbInteger Then
        MS_SayiSifir = CDbl(v)
    ElseIf MS_SayiDene(v, d) Then
        MS_SayiSifir = d
    End If
End Function

Private Function MS_Kirp(ByRef a As Variant, ByVal n As Long, ByVal k As Long) As Variant
    Dim s() As Variant, r As Long, c As Long
    ReDim s(1 To n, 1 To k)
    For r = 1 To n
        For c = 1 To k: s(r, c) = a(r, c): Next c
    Next r
    MS_Kirp = s
End Function

Private Function MS_GrupSeri(ByVal mg As String) As Boolean
    MS_GrupSeri = (mg = "M03550000" Or mg = "M03250000")
End Function

' IQ09 + MB52 -> AKSORA satirlari (14 sutun). Seri gruplari MB52'den silinir.
Private Function MS_Birlestir(ByVal iq As Variant, ByVal mb As Variant, ByRef ozet As String) As Variant
    Dim d As Object, mgOf As Object, brOf As Object, adOf As Object, sayac As Object
    Dim r As Long, n As Long, h As Long, k As String, ks As String, t As Variant, mg As String
    Dim depo As String, malz As String, parti As String, birim As Double, ad As String, fark As Long
    Dim cikti() As Variant, nIq As Long, nMb As Long, x As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set mgOf = CreateObject("Scripting.Dictionary"): mgOf.CompareMode = vbTextCompare
    Set brOf = CreateObject("Scripting.Dictionary"): brOf.CompareMode = vbTextCompare
    Set adOf = CreateObject("Scripting.Dictionary"): adOf.CompareMode = vbTextCompare
    Set sayac = CreateObject("Scripting.Dictionary"): sayac.CompareMode = vbTextCompare
    For r = 1 To UBound(mb, 1)
        malz = MS_Malzeme(mb(r, 6)): depo = CStr(mb(r, 2)): parti = AKFIX_NORMAL(CStr(mb(r, 5)))
        For Each x In Array(depo, "*")
            k = CStr(x) & "|" & malz & "|" & parti
            If d.Exists(k) Then
                t = d(k): t(0) = t(0) + CDbl(mb(r, 8)): t(1) = t(1) + CDbl(mb(r, 10)): d(k) = t
            Else
                d.Add k, Array(CDbl(mb(r, 8)), CDbl(mb(r, 10)))
            End If
        Next x
        If Len(CStr(mb(r, 4))) > 0 Then mgOf(malz) = CStr(mb(r, 4))
        If Len(CStr(mb(r, 9))) > 0 Then brOf(malz) = CStr(mb(r, 9))
        If Len(CStr(mb(r, 3))) > 0 Then adOf(depo) = CStr(mb(r, 3))
    Next r
    ReDim cikti(1 To UBound(iq, 1) + UBound(mb, 1), 1 To 14)
    ' 1) Seri numarali stok (IQ09), her satir 1 adet
    For r = 1 To UBound(iq, 1)
        depo = CStr(iq(r, 2)): h = MS_Hedef(depo)
        If h > 0 And (Len(CStr(iq(r, 1))) = 0 Or CStr(iq(r, 1)) = "7155") Then
            malz = MS_Malzeme(iq(r, 3)): parti = AKFIX_NORMAL(CStr(iq(r, 6)))
            mg = vbNullString: If mgOf.Exists(malz) Then mg = mgOf(malz)
            If h <> 2 Or mg = vbNullString Or mg = "M03550000" Or mg = "M03550100" Or mg = "M03250000" Then
                n = n + 1: nIq = nIq + 1
                ad = MS_DepoAdi("7155", depo): If Len(ad) = 0 And adOf.Exists(depo) Then ad = adOf(depo)
                cikti(n, 1) = "7155": cikti(n, 2) = depo: cikti(n, 3) = ad: cikti(n, 4) = mg
                cikti(n, 5) = CStr(iq(r, 3)): cikti(n, 6) = CStr(iq(r, 4)): cikti(n, 7) = 1
                cikti(n, 8) = "ADT": If brOf.Exists(malz) Then cikti(n, 8) = brOf(malz)
                cikti(n, 9) = CStr(iq(r, 6)): cikti(n, 10) = CStr(iq(r, 5))
                ' Eslesme sirasi: depo+parti, il+parti, depo+bos parti, il+bos parti
                ' (bos parti = SAP'de degerleme turu ayrimi olmayan malzeme).
                k = depo & "|" & malz & "|" & parti
                If Not d.Exists(k) And Not d.Exists("*|" & malz & "|" & parti) Then
                    If d.Exists(depo & "|" & malz & "|") Or d.Exists("*|" & malz & "|") Then k = depo & "|" & malz & "|"
                End If
                sayac(k) = sayac(k) + 1
                If Not d.Exists(k) Then k = "*|" & Mid$(k, Len(depo) + 2)
                If d.Exists(k) Then
                    t = d(k)
                    If CDbl(t(0)) <> 0 Then cikti(n, 11) = Round(CDbl(t(1)) / CDbl(t(0)), 2): cikti(n, 13) = "SAP"
                End If
            End If
        End If
    Next r
    ' 2) Toplu stok (MB52); seri numarali iki grup IQ09'dan geldigi icin atlanir.
    For r = 1 To UBound(mb, 1)
        depo = CStr(mb(r, 2)): h = MS_Hedef(depo): mg = CStr(mb(r, 4))
        malz = MS_Malzeme(mb(r, 6)): parti = AKFIX_NORMAL(CStr(mb(r, 5)))
        If h > 0 And MS_GrupSeri(mg) Then
            k = depo & "|" & malz & "|" & parti
            If CLng(CDbl(mb(r, 8))) <> CLng(sayac(k)) Then fark = fark + 1
        End If
        If h > 0 And Not MS_GrupSeri(mg) And (h <> 2 Or mg = "M03550100") And CDbl(mb(r, 8)) <> 0 Then
            n = n + 1: nMb = nMb + 1
            ad = MS_DepoAdi("7155", depo): If Len(ad) = 0 Then ad = CStr(mb(r, 3))
            cikti(n, 1) = "7155": cikti(n, 2) = depo: cikti(n, 3) = ad: cikti(n, 4) = mg
            cikti(n, 5) = CStr(mb(r, 6)): cikti(n, 6) = CStr(mb(r, 7)): cikti(n, 7) = CDbl(mb(r, 8))
            cikti(n, 8) = CStr(mb(r, 9)): cikti(n, 9) = CStr(mb(r, 5)): cikti(n, 10) = vbNullString
            cikti(n, 11) = Round(CDbl(mb(r, 10)), 2): cikti(n, 13) = "SAP"
        End If
    Next r
    ozet = "IQ09 seri numaralý " & AKSORA_SAYI_METNI(nIq, "#,##0") & " satýr  |  MB52 toplu " & _
        AKSORA_SAYI_METNI(nMb, "#,##0") & " satýr" & IIf(fark > 0, "  |  IQ09/MB52 adet farký: " & fark & " kalem", "")
    MS_Teshis ozet
    If n = 0 Then Err.Raise vbObjectError + 1980, , "IQ09 ve MB52'den müþteri depolarýna ait stok bulunamadý."
    MS_Sirala cikti, n
    MS_Birlestir = MS_Kirp(cikti, n, 14)
End Function

' Satirlar sayfa, depo, mal grubu, malzeme ve seri numarasina gore siralanir.
Private Sub MS_Sirala(ByRef a As Variant, ByVal n As Long)
    Dim anahtar() As String, idx() As Long, r As Long, c As Long, gecici() As Variant
    If n < 2 Then Exit Sub
    ReDim anahtar(1 To n): ReDim idx(1 To n)
    For r = 1 To n
        idx(r) = r
        anahtar(r) = CStr(MS_Hedef(CStr(a(r, 2)))) & "|" & CStr(a(r, 2)) & "|" & CStr(a(r, 4)) & "|" & _
            Right$(String$(20, "0") & CStr(a(r, 5)), 20) & "|" & CStr(a(r, 9)) & "|" & Right$(String$(30, " ") & CStr(a(r, 10)), 30)
    Next r
    MS_HizliSirala anahtar, idx, 1, n
    ReDim gecici(1 To n, 1 To 14)
    For r = 1 To n
        For c = 1 To 14: gecici(r, c) = a(idx(r), c): Next c
    Next r
    For r = 1 To n
        For c = 1 To 14: a(r, c) = gecici(r, c): Next c
    Next r
End Sub

Private Sub MS_HizliSirala(ByRef k() As String, ByRef idx() As Long, ByVal ilk As Long, ByVal son As Long)
    Dim i As Long, j As Long, pivot As String, t As Long
    i = ilk: j = son
    pivot = k(idx((ilk + son) \ 2))
    Do While i <= j
        Do While StrComp(k(idx(i)), pivot, vbBinaryCompare) < 0: i = i + 1: Loop
        Do While StrComp(k(idx(j)), pivot, vbBinaryCompare) > 0: j = j - 1: Loop
        If i <= j Then
            t = idx(i): idx(i) = idx(j): idx(j) = t
            i = i + 1: j = j - 1
        End If
    Loop
    If ilk < j Then MS_HizliSirala k, idx, ilk, j
    If i < son Then MS_HizliSirala k, idx, i, son
End Sub

Private Function MS_Norm(ByVal v As Variant) As String
    Dim t As String
    If IsError(v) Then Exit Function
    t = Replace(Replace(Replace(Trim$(CStr(v)), ChrW(305), "I"), ChrW(160), ""), ChrW(304), "I")
    t = AKFIX_NORMAL(t)
    t = Replace(t, " ", ""): t = Replace(t, ".", ""): t = Replace(t, "/", "")
    t = Replace(t, "(", ""): t = Replace(t, ")", ""): t = Replace(t, "-", "")
    MS_Norm = t
End Function

Private Function MS_Alan(ByVal baslik As String, ByVal fiyat As Boolean) As Long
    Select Case baslik
        Case "UY", "URETIMYERI", "WERKS": MS_Alan = 1
        Case "DPYR", "DEPOYERI", "LGORT", "DEPO": MS_Alan = 2
        Case "DEPOFIKTI", "DEPOFIKTIFGUMRUKTANIMI", "DEPOYERITANIMI", "DEPOTANIMI", "LGOBE": MS_Alan = 3
        Case "MALGRUBU", "MATKL": MS_Alan = 4
        Case "MALZEME", "MATNR": MS_Alan = 5
        Case "MALZEMEKISAMETNI", "MALZEMEACIKLAMASI", "TANIM", "MAKTX": MS_Alan = 6
        Case "MIKTAR", "MIKT", "MENGE", "TAHDITSIZ", "LABST": MS_Alan = 7
        Case "TOB", "TEMELOLCUBIRIMI", "BIRIM", "MEINS": MS_Alan = 8
        Case "DEGERLEMETURU", "DEGERLTURU", "BWTAR": MS_Alan = 9
        Case "SERINUMARASI", "SERINUM", "SERINO", "SERNR": MS_Alan = 10
        Case "DEGER", "DEGERTL", "DEGERTRY", "STOKDEGERI", "STOKDEGER", "TOPLAMSTOKDEGERI", "TOPLAMDEGER", _
             "DEGERI", "TUTAR", "TOPLAMTUTAR", "STOKTUTARI", "KAYITLIDEGER", "SALK3", "WERT", "GSWRT": MS_Alan = 11
    End Select
End Function



Private Function MS_Sayi(ByVal v As Variant) As Double
    Dim t As String, re As Object
    If IsError(v) Or IsEmpty(v) Then Err.Raise vbObjectError + 1905, , "Boþ veya hatalý miktar/deðer hücresi."
    If VarType(v) <> vbString And IsNumeric(v) Then MS_Sayi = CDbl(v): Exit Function
    t = Trim$(CStr(v)): t = Replace(t, ChrW(160), ""): t = Replace(t, " ", "")
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = "^-?[0-9]{1,3}([.][0-9]{3})+$"
    If InStr(t, ",") = 0 And re.Test(t) Then t = Replace(t, ".", "")
    If Right$(t, 1) = "-" Then t = "-" & Left$(t, Len(t) - 1)
    If InStr(t, ",") > 0 Then
        If InStrRev(t, ".") > InStrRev(t, ",") Then
            t = Replace(t, ",", "")
        Else
            t = Replace(Replace(t, ".", ""), ",", ".")
        End If
    End If
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^-?[0-9]+([.][0-9]+)?$"
    If Not re.Test(t) Then Err.Raise vbObjectError + 1906, , "Sayýsal alan okunamadý: " & CStr(v)
    MS_Sayi = Val(t)
End Function



Private Sub MS_Coklu(ByVal ses As Object, ByVal alan As String, ByVal degerler As Variant, ByVal ilk As String, ByVal son As String)
    Dim sil As Object, ip As String, sp As String, tip As String
    Dim visS As Long, visI As Long, nv As Long, i As Long, satirNo As Long, deger As String
    ses.FindById("wnd[0]/usr/btn%_" & alan & "_%_APP_%-VALU_PUSH").Press
    Set sil = ses.FindById("wnd[1]/tbar[0]/btn[16]", False)
    If sil Is Nothing Then Err.Raise vbObjectError + 1913, , "SAP çoklu seçim temizleme düðmesi bulunamadý."
    tip = AKFIX_NORMAL(sil.Tooltip)
    If InStr(tip, "DELETE") = 0 And InStr(tip, "SIL") = 0 Then _
        Err.Raise vbObjectError + 1914, , "SAP çoklu seçim temizleme düðmesi doðrulanamadý: " & sil.Tooltip
    sil.Press
    sp = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/"
    ip = "wnd[1]/usr/tabsTAB_STRIP/tabpINTL/ssubSCREEN_HEADER:SAPLALDB:3020/tblSAPLALDBINTERVAL/"
    ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    visS = CLng(ses.FindById(Left$(sp, Len(sp) - 1)).VisibleRowCount)
    If IsArray(degerler) Then
        nv = UBound(degerler) - LBound(degerler) + 1
        For i = 0 To nv - 1
            If i < visS Then _
                ses.FindById(sp & "ctxtRSCSEL_255-SLOW_I[1," & CStr(i) & "]").Text = CStr(degerler(LBound(degerler) + i))
        Next i
    End If
    If Len(ilk) > 0 Or nv > visS Then
        ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpINTL").Select
        visI = CLng(ses.FindById(Left$(ip, Len(ip) - 1)).VisibleRowCount)
        satirNo = 0
        If Len(ilk) > 0 Then
            ses.FindById(ip & "ctxtRSCSEL_255-ILOW_I[1,0]").Text = ilk
            ses.FindById(ip & "ctxtRSCSEL_255-IHIGH_I[2,0]").Text = son
            satirNo = 1
        End If
        ' Tek deðer tablosuna sýðmayan deðerler "X - X" aralýðý olarak yazýlýr.
        For i = visS To nv - 1
            If satirNo >= visI Then Err.Raise vbObjectError + 1920, , "SAP çoklu seçim penceresine sýðmayan deðerler var (" & alan & ")."
            deger = CStr(degerler(LBound(degerler) + i))
            ses.FindById(ip & "ctxtRSCSEL_255-ILOW_I[1," & CStr(satirNo) & "]").Text = deger
            ses.FindById(ip & "ctxtRSCSEL_255-IHIGH_I[2," & CStr(satirNo) & "]").Text = deger
            satirNo = satirNo + 1
        Next i
    End If
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press
End Sub

Private Sub MS_Duzen(ByVal ses As Object, ByVal musteri As Boolean)
    Dim alv As Object, grid As Object, yol As String, r As Long, ad As String, secim As Long
    If musteri Then
        Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        alv.SetCurrentCell -1, "MAKTX": alv.SelectColumn "MAKTX"
        alv.ContextMenu: alv.SelectContextMenuItem "&COL0"
        ses.FindById("wnd[1]/tbar[0]/btn[36]").Press
        yol = "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R9/ssubSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell"
    Else
        ses.FindById("wnd[0]/tbar[1]/btn[33]").Press
        yol = "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell"
    End If
    Set grid = ses.FindById(yol): secim = -1
    For r = 0 To CLng(grid.RowCount) - 1
        ad = Trim$(CStr(grid.GetCellValue(r, "VARIANT")))
        If (musteri And ad = "/PANEL") Or (Not musteri And (ad = "PANEL" Or ad = "/PANEL")) Then secim = r: Exit For
    Next r
    If secim < 0 Then Err.Raise vbObjectError + 1915, , "SAP'de beklenen PANEL düzeni bulunamadý."
    grid.SetCurrentCell secim, "TEXT": grid.SelectedRows = CStr(secim): grid.ClickCurrentCell
    If musteri Then ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
End Sub

Private Function MS_Sorgu(ByVal ses As Object, ByVal ikinci As Boolean) As Variant
    Dim kutu As Variant, bos As Variant
    ses.StartTransaction "ZMMRMC001": MS_Bekle ses
    For Each kutu In Array("P_FIKTIF", "P_FIKSTT", "P_MUSSIP", "P_MUSSTK", "P_SATFAS")
        ses.FindById("wnd[0]/usr/chk" & CStr(kutu)).Selected = False
    Next kutu
    ses.FindById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = "7155"
    If ikinci Then
        ' Ýlçe depolarý: K001-K999 ve üç mal grubu
        MS_Coklu ses, "S_LGORT", bos, "K001", "K999"
        MS_Coklu ses, "S_MATKL", Array("M03550000", "M03550100", "M03250000"), "", ""
    Else
        ' CDP: E001-E999, D003, D005, S001, Z002, Z003, Z007, EM14  |  Yüklenici: M132
        MS_Coklu ses, "S_LGORT", Array("D003", "D005", "S001", "Z002", "Z003", "Z007", "EM14", "M132"), "E001", "E999"
        MS_Coklu ses, "S_MATKL", bos, "", ""
    End If
    ses.FindById("wnd[0]/tbar[1]/btn[8]").Press: MS_Bekle ses
    MS_Duzen ses, True
    MS_Sorgu = MS_ALV_Oku(ses, False)
End Function

Private Function MS_Malzeme(ByVal v As Variant) As String
    Dim t As String, re As Object
    t = Trim$(CStr(v))
    Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[0-9]+$"
    If re.Test(t) Then
        Do While Len(t) > 1 And Left$(t, 1) = "0": t = Mid$(t, 2): Loop
    End If
    MS_Malzeme = UCase$(t)
End Function

'==========================================================
' DEPO -> SAYFA ESLEMESI
' 1 = M_CDP_STOK       : E001-E999, D003, D005, S001, Z002, Z003, Z007, EM14
' 2 = M_ILCE_STOK      : K001-K999
' 3 = M_YUKLENICI_STOK : M132
' 0 = taninmayan depo
'==========================================================
Private Function MS_Hedef(ByVal depo As String) As Long
    Dim n As Long
    depo = UCase$(Trim$(depo))
    Select Case depo
        Case "D003", "D005", "S001", "Z002", "Z003", "Z007", "EM14": MS_Hedef = 1: Exit Function
        Case "M132": MS_Hedef = 3: Exit Function
    End Select
    If Len(depo) = 4 Then
        If Mid$(depo, 2) Like "###" Then
            n = CLng(Mid$(depo, 2))
            If n >= 1 Then
                If Left$(depo, 1) = "E" Then MS_Hedef = 1
                If Left$(depo, 1) = "K" Then MS_Hedef = 2
            End If
        End If
    End If
End Function

Private Sub MS_KapsamKontrol(ByRef a As Variant, ByVal izin1 As Long, ByVal izin2 As Long, ByVal grupSart As Boolean)
    Dim r As Long, h As Long, g As String
    For r = 1 To UBound(a, 1)
        h = MS_Hedef(CStr(a(r, 2)))
        If h = 0 Or (h <> izin1 And h <> izin2) Then _
            Err.Raise vbObjectError + 1916, , "SAP çýktýsýnda sorgu kapsamý dýþýnda depo var: " & CStr(a(r, 2))
        If grupSart Then
            g = UCase$(CStr(a(r, 4)))
            If g <> "M03550000" And g <> "M03550100" And g <> "M03250000" Then _
                Err.Raise vbObjectError + 1916, , "SAP çýktýsýnda sorgu kapsamý dýþýnda mal grubu var: " & CStr(a(r, 2)) & " / " & g
        End If
    Next r
End Sub

Private Sub MS_SeriKontrol(ByRef a As Variant)
    Dim seri As Object, r As Long, sk As String
    Set seri = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(a, 1)
        If Len(CStr(a(r, 10))) > 0 Then
            sk = MS_Malzeme(a(r, 5)) & "|" & CStr(a(r, 10))
            If seri.Exists(sk) Then Err.Raise vbObjectError + 1917, , "Ayný malzeme/seri numarasý birden fazla satýrda: " & sk
            seri(sk) = True
        End If
    Next r
End Sub

' Fiyat listesinden (Malzeme + Parti) birim fiyat uygular. Fiyatsiz kalan satir sayisini dondurur.
Private Function MS_FiyatUygula(ByRef a As Variant, ByVal d As Object, ByVal tarihMetni As String) As Long
    Dim r As Long, k As String, v As Variant, eksik As Long
    For r = 1 To UBound(a, 1)
      If (CStr(a(r, 13)) = "SAP" Or CStr(a(r, 13)) = "SAP deðeri") And Not IsEmpty(a(r, 11)) And VarType(a(r, 11)) <> vbString Then
        a(r, 13) = "SAP deðeri": a(r, 14) = 0
        If IsNumeric(a(r, 7)) And CDbl(a(r, 7)) <> 0 Then a(r, 12) = CDbl(a(r, 11)) / CDbl(a(r, 7)) Else a(r, 12) = vbNullString
        mSapSatir = mSapSatir + 1
      Else
        mListeSatir = mListeSatir + 1
        k = MF_Anahtar(a(r, 5), a(r, 9))
        a(r, 11) = vbNullString: a(r, 12) = vbNullString: a(r, 14) = 1
        If d.Exists(k) Then
            v = d(k)
            If CStr(v(1)) = "TEK" Then
                a(r, 12) = CDbl(v(0))
                a(r, 11) = CDec(a(r, 7)) * CDec(v(0))
                a(r, 13) = "Fiyat listesi (" & tarihMetni & ")"
                a(r, 14) = 0
            ElseIf CStr(v(1)) = "BELIRSIZ" Then
                a(r, 13) = "Fiyat belirsiz (listede çeliþkili)"
            Else
                a(r, 13) = "Listede fiyat 0"
            End If
        Else
            a(r, 13) = "Fiyat listede yok"
        End If
      End If
        If CLng(a(r, 14)) > 0 Then
            eksik = eksik + 1
            mFiyatsizAdet = mFiyatsizAdet + 1
            MS_Teshis "Fiyatsýz: " & CStr(a(r, 2)) & " | " & CStr(a(r, 5)) & " | " & CStr(a(r, 6)) & _
                " | parti '" & CStr(a(r, 9)) & "' | seri " & CStr(a(r, 10)) & " | " & CStr(a(r, 13))
            If mFiyatsizAdet <= 5 Then mFiyatsiz = mFiyatsiz & vbCrLf & "   • " & CStr(a(r, 5)) & " " & _
                Left$(CStr(a(r, 6)), 28) & "  (" & CStr(a(r, 2)) & ", parti: " & IIf(Len(CStr(a(r, 9))) > 0, CStr(a(r, 9)), "boþ") & ")"
        End If
    Next r
    MS_FiyatUygula = eksik
End Function

Private Function MS_Ayir(ByRef a As Variant, ByVal hedef As Long) As Variant
    Dim r As Long, c As Long, n As Long, k As Long, sonuc() As Variant
    For r = 1 To UBound(a, 1)
        If MS_Hedef(CStr(a(r, 2))) = hedef Then n = n + 1
    Next r
    If n = 0 Then Exit Function
    ReDim sonuc(1 To n, 1 To 14)
    For r = 1 To UBound(a, 1)
        If MS_Hedef(CStr(a(r, 2))) = hedef Then
            k = k + 1
            For c = 1 To 14: sonuc(k, c) = a(r, c): Next c
        End If
    Next r
    MS_Ayir = sonuc
End Function

Private Sub MS_SeriBirlesimKontrol(ByRef bir As Variant, ByRef iki As Variant)
    Dim d As Object, r As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(bir, 1)
        If Len(CStr(bir(r, 10))) > 0 Then d(MS_Malzeme(bir(r, 5)) & "|" & CStr(bir(r, 10))) = True
    Next r
    For r = 1 To UBound(iki, 1)
        If Len(CStr(iki(r, 10))) > 0 Then
            k = MS_Malzeme(iki(r, 5)) & "|" & CStr(iki(r, 10))
            If d.Exists(k) Then Err.Raise vbObjectError + 1919, , "Ýki sorguda ayný malzeme/seri bulundu. Ýþlem sýrasýnda depo hareketi olmuþ olabilir: " & k
        End If
    Next r
End Sub

Public Sub AKSORA_MUSTERI_SAP_BAGLA(ByVal ws As Worksheet)
    Dim s As Shape, adlar As Variant, ad As Variant
    Select Case ws.Name
        Case "M_CDP_STOK": adlar = Array("SAP_STOK_GUNCELLE_BTN")
        Case "M_ÝLÇE_STOK": adlar = Array("CDP_ILCE_STOK_BTN")
        Case "M_YÜKLENÝCÝ_STOK": adlar = Array("YUK_EDAS_GUNCELLE_BTN")
        Case "M_ANA PANEL": adlar = Array("btnCDP", "btnIlce", "btnYuk")
        Case Else: Exit Sub
    End Select
    For Each ad In adlar
        Set s = Nothing
        On Error Resume Next
        Set s = ws.Shapes(CStr(ad))
        On Error GoTo 0
        If Not s Is Nothing Then s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_7155_GUNCELLE"
    Next ad
End Sub

Private Sub MS_Yaz(ByVal ws As Worksheet, ByRef a As Variant)
    Dim son As Long, n As Long, s As Shape
    If IsEmpty(a) Then n = 0 Else n = UBound(a, 1)
    son = application.Max(7, AKSORA_M_SON(ws))
    ws.Unprotect ""
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    ws.Range("A7:N" & application.Max(son, n + 6)).ClearContents
    If n > 0 Then
        ws.Range("A7:F" & n + 6).NumberFormat = "@"
        ws.Range("H7:J" & n + 6).NumberFormat = "@"
        ws.Range("A7").Resize(n, 14).Value2 = a
    End If
    ws.Range("K6:N6").Value = Array("Deðer (TL)", "Birim stok deðeri (TL)", "Fiyat durumu", "Fiyat eksik")
    ws.Columns("K:N").Hidden = True
    If n > 0 Then ws.Range("K7:L" & n + 6).NumberFormat = "#,##0.00 ""TL"""
    ws.Range("C2").Value = Now: ws.Range("C2").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    For Each s In ws.Shapes
        If InStr(s.Name, "STOK_KPI_ETIKET_2") > 0 Then s.TextFrame2.TextRange.Text = "Toplam Stok Deðeri"
    Next s
    AKSORA_MUSTERI_SAYFA_HAZIRLA ws
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Public Sub AKSORA_MUSTERI_7155_GUNCELLE()
    Dim ses As Object, hMesaj As String, bir As Variant, iki As Variant, d As Object
    Dim mb As Object, mbMesaj As String, iq As Variant, mbv As Variant
    Dim sayfa(1 To 3) As Worksheet, yeni(1 To 3) As Variant, eski(1 To 3) As Variant
    Dim tarihEski(1 To 3) As Variant, sonEski(1 To 3) As Long
    Dim ev As Boolean, sc As Boolean, hesap As XlCalculation, yazildi As Boolean
    Dim i As Long, nn As Long, tarih As Date, adlar As Variant, eksik As Long, adet(1 To 3) As Long
    If mCalisiyor Or AKSORA_V11_MESGUL Then Exit Sub
    If CStr(ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2) <> "ARTVIN" Then
        MsgBox "Bu aktarým yalnýzca 7155 / Artvin için hazýrlanmýþtýr.", vbExclamation, "AKSORA": Exit Sub
    End If
    mCalisiyor = True: mBas = Timer: mSeritZaman = 0
    mHizliSebep = vbNullString: mHizliSatir = 0: mStandartSatir = 0: mSapSatir = 0: mListeSatir = 0
    mStandartOnay = False
    Set mDisaDosyalar = Nothing
    mFiyatsiz = vbNullString: mFiyatsizAdet = 0
    mBasZamanMetin = Format$(Now - TimeSerial(0, 0, 2), "yyyymmddhhnnss")
    MS_OleUyariKapat
    AKSORA_MUSTERI_PANEL_YENILE True
    AKSORA_MUSTERI_SERIT "SAP güncellemesi baþlatýlýyor...", 1
    ev = application.EnableEvents: sc = application.ScreenUpdating: hesap = application.Calculation
    On Error GoTo Hata
    ' Fiyat listesi yalnizca MB52'de karsiligi olmayan satirlar icin yedektir.
    Set d = MF_FIYAT_SOZLUGU(tarih, hMesaj)
    If d Is Nothing Then Set d = CreateObject("Scripting.Dictionary")
    hMesaj = vbNullString
    Set ses = SAP_SESSION_AL("CDP", hMesaj)
    If ses Is Nothing Then Err.Raise vbObjectError + 1918, , hMesaj
    adlar = Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
    For i = 1 To 3
        Set sayfa(i) = ThisWorkbook.Worksheets(CStr(adlar(i - 1)))
    Next i
    application.EnableEvents = False: application.Calculation = xlCalculationManual
    AKSORA_V11_MESGUL = True
    MS_Durum "1/2 | IQ09 seri numaralý stok (7155)..."
    iq = MS_IQ09_Oku(ses)
    MS_Durum "2/2 | MB52 toplu stok ve deðerler (7155)..."
    mbv = MS_MB52_Oku(ses)
    MS_Durum "Listeler birleþtiriliyor..."
    bir = MS_Birlestir(iq, mbv, mbMesaj)
    MS_SeriKontrol bir
    mFiyatTur = True
    MS_Durum "Deðerler uygulanýyor..."
    eksik = MS_FiyatUygula(bir, d, Format$(tarih, "dd.mm.yyyy"))
    yeni(1) = MS_Ayir(bir, 1): yeni(2) = MS_Ayir(bir, 2): yeni(3) = MS_Ayir(bir, 3)
    For i = 1 To 3
        sonEski(i) = application.Max(7, AKSORA_M_SON(sayfa(i)))
        eski(i) = sayfa(i).Range("A7:N" & sonEski(i)).Value2
        tarihEski(i) = sayfa(i).Range("C2").Value
        If IsEmpty(yeni(i)) Then adet(i) = 0 Else adet(i) = UBound(yeni(i), 1)
    Next i
    yazildi = True
    For i = 1 To 3: MS_Yaz sayfa(i), yeni(i): Next i
    yazildi = False
    MS_Teshis "Tamamlandý | CDP " & adet(1) & " | Ýlçe " & adet(2) & " | Yüklenici " & adet(3) & " | " & _
        AKSORA_SAYI_METNI(MS_Gecen(mBas), "0.0") & " sn"
    MS_Durum "Tamamlandý | CDP " & adet(1) & " | Ýlçe " & adet(2) & " | Yüklenici " & adet(3) & _
        " | Fiyatý olmayan " & eksik & " | " & AKSORA_SAYI_METNI(MS_Gecen(mBas), "0.00") & " sn"
    ThisWorkbook.names.Add Name:="_AKSORA_MUSTERI_7155_SON_SURE", RefersTo:="=" & Trim$(Str$(Round(MS_Gecen(mBas), 2))), Visible:=False
    MS_ExcelTemizle
    MS_DisaKapat
    AKSORA_MUSTERI_PANEL_YENILE
    AKSORA_MUSTERI_SERIT "Güncelleme tamamlandý  •  " & AKSORA_MUSTERI_OZET(), 2
    MsgBox "Müþteri 7155 güncellendi." & vbCrLf & vbCrLf & _
        "CDP stok satýrý: " & AKSORA_SAYI_METNI(adet(1), "#,##0") & vbCrLf & _
        "Ýlçe stok satýrý: " & AKSORA_SAYI_METNI(adet(2), "#,##0") & vbCrLf & _
        "Yüklenici stok satýrý: " & AKSORA_SAYI_METNI(adet(3), "#,##0") & vbCrLf & vbCrLf & _
        mbMesaj & vbCrLf & _
        "Deðer kaynaðý: SAP MB52 " & AKSORA_SAYI_METNI(mSapSatir, "#,##0") & " satýr  |  Fiyat listesi " & _
            AKSORA_SAYI_METNI(mListeSatir, "#,##0") & " satýr" & vbCrLf & _
        "Fiyatý bulunamayan satýr: " & AKSORA_SAYI_METNI(eksik, "#,##0") & mFiyatsiz & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(MS_Gecen(mBas), "0.00") & " sn", vbInformation, "AKSORA | Müþteri 7155"
Cikis:
    MS_OleUyariAc
    application.Calculation = hesap: application.EnableEvents = ev: application.ScreenUpdating = sc
    application.StatusBar = False: AKSORA_V11_MESGUL = False: mCalisiyor = False
    Exit Sub
Hata:
    hMesaj = mDurum & vbCrLf & "Hata " & CStr(Err.Number) & ": " & Err.Description
    If yazildi Then
        On Error Resume Next
        Err.Clear
        For i = 1 To 3
            If IsEmpty(yeni(i)) Then nn = 0 Else nn = UBound(yeni(i), 1)
            sayfa(i).Unprotect ""
            sayfa(i).Range("A7:N" & application.Max(sonEski(i), nn + 6)).ClearContents
            sayfa(i).Range("A7").Resize(UBound(eski(i), 1), 14).Value2 = eski(i)
            sayfa(i).Range("C2").Value = tarihEski(i)
            sayfa(i).Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
        Next i
        If Err.Number <> 0 Then hMesaj = hMesaj & vbCrLf & "Eski kayýtlarýn geri yüklenmesi doðrulanamadý: " & Err.Description
        On Error GoTo 0
    End If
    MS_ExcelTemizle
    MS_DisaKapat
    AKSORA_MUSTERI_PANEL_YENILE
    AKSORA_MUSTERI_SERIT "Güncelleme tamamlanamadý  •  " & Replace(hMesaj, vbCrLf, "  •  "), 3
    MsgBox "Aktarým tamamlanamadý: " & hMesaj, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Function MS_ALV_Oku(ByVal ses As Object, ByVal fiyat As Boolean) As Variant
    Dim alv As Object, kolonlar As Object, basliklar As Object, ids(1 To 11) As String
    Dim id As String, no As Long, i As Long, j As Long, k As Long, c As Long
    Dim n As Long, adet As Long, blok As Long, bas As Long, bitis As Long, r As Long
    Dim buf() As Variant, sonuc() As Variant, satir(1 To 14) As Variant
    Dim zorunlu As Variant, x As Variant, tekrar As Long, hataNo As Long, hataMetni As String
    Dim tamam As Boolean, ozet As Boolean, asama As String, ilk As Long, hizli As Variant
    Dim tumBaslik As String, nb As String, adayDeger As String
    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    On Error Resume Next
    alv.PressToolbarButton "&DELETE_FILTER"
    On Error GoTo 0
    Set kolonlar = alv.ColumnOrder
    For i = 0 To CLng(kolonlar.Count) - 1
        id = CStr(kolonlar.Item(i))
        no = MS_Alan(MS_Norm(id), fiyat)
        If no = 0 Then
            Set basliklar = alv.GetColumnTitles(id)
            For j = 0 To CLng(basliklar.Count) - 1
                no = MS_Alan(MS_Norm(CStr(basliklar.Item(j))), fiyat)
                If no > 0 Then Exit For
            Next j
        End If
        If no > 0 Then
            If Len(ids(no)) > 0 And ids(no) <> id Then _
                Err.Raise vbObjectError + 1922, , "ALV'de ayný alana uyan birden fazla sütun var: " & ids(no) & "/" & id
            ids(no) = id
        End If
        tumBaslik = vbNullString
        On Error Resume Next
        Set basliklar = Nothing
        Set basliklar = alv.GetColumnTitles(id)
        For j = 0 To CLng(basliklar.Count) - 1
            tumBaslik = tumBaslik & " / " & CStr(basliklar.Item(j))
        Next j
        On Error GoTo 0
        MS_Teshis "ALV kolon " & CStr(i) & ": " & id & " -> alan " & CStr(no) & tumBaslik
        If no = 0 And Len(ids(11)) = 0 Then
            nb = MS_Norm(id & tumBaslik)
            If (InStr(nb, "DEGER") > 0 And InStr(nb, "DEGERLEME") = 0 And InStr(nb, "DEGERLTUR") = 0) Or _
               InStr(nb, "TUTAR") > 0 Then adayDeger = id
        End If
    Next i
    If Len(ids(11)) = 0 And Len(adayDeger) > 0 Then
        ids(11) = adayDeger
        MS_Teshis "Deðer sütunu (benzer ad): " & adayDeger
    End If
    If fiyat Then zorunlu = Array(1, 2, 5, 7, 8, 11) Else zorunlu = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
    For Each x In zorunlu
        If Len(ids(CLng(x))) = 0 Then Err.Raise vbObjectError + 1923, , "ALV düzeninde gerekli alan bulunamadý: " & CStr(x)
    Next x
    If fiyat Then mFiyatTur = (Len(ids(9)) > 0)
    n = CLng(alv.RowCount)
    If n < 1 Or n > 250000 Then Err.Raise vbObjectError + 1924, , "ALV stok satýr sayýsý geçersiz: " & CStr(n)
    ' Once hizli yol: listeyi dosyaya aktarip tek seferde oku.
    asama = mDurum
    MS_Durum asama & " | Liste dosyaya aktarýlýyor..."
    hizli = Empty
    If Not fiyat Then hizli = MS_ALV_Excel(ses, alv, ids, n)
    If IsArray(hizli) Then
        mHizliSatir = mHizliSatir + UBound(hizli, 1)
        MS_Durum asama & " | " & CStr(UBound(hizli, 1)) & " satýr okundu"
        MS_ALV_Oku = hizli
        Exit Function
    End If
    MS_Teshis "Hýzlý okuma kullanýlamadý: " & mHizliSebep
    If Not mStandartOnay Then
        If MsgBox("Hýzlý okuma kullanýlamadý: " & mHizliSebep & vbCrLf & vbCrLf & _
            "Standart okuma yaklaþýk 11 dakika sürer. Devam edilsin mi?" & vbCrLf & vbCrLf & _
            "Teþhis dosyasý: " & MS_TeshisYolu(), vbQuestion + vbYesNo + vbDefaultButton2, "AKSORA") = vbNo Then
            Err.Raise vbObjectError + 1999, , "Güncelleme durduruldu. Teþhis dosyasý: " & MS_TeshisYolu()
        End If
        mStandartOnay = True
    End If
    MS_Durum asama & IIf(Len(mHizliSebep) > 0, " | standart okuma (" & mHizliSebep & ")", "")
    blok = application.Min(30, CLng(alv.VisibleRowCount))
    If blok < 1 Then Err.Raise vbObjectError + 1925, , "ALV görünür satýr sayýsý okunamadý."
    ReDim buf(1 To n, 1 To 14)
    asama = mDurum
    For bas = 0 To n - 1 Step blok
        alv.FirstVisibleRow = application.Min(bas, application.Max(0, n - CLng(alv.VisibleRowCount)))
        DoEvents: MS_Sleep 15
        bitis = application.Min(n - 1, bas + blok - 1)
        If CLng(alv.RowCount) <> n Then Err.Raise vbObjectError + 1926, , "Okuma sýrasýnda SAP satýr sayýsý deðiþti; aktarým durduruldu."
        For r = bas To bitis
            tamam = False
            For tekrar = 0 To 2
                For c = 1 To 14: satir(c) = Empty: Next c
                hataNo = 0: hataMetni = ""
                On Error Resume Next
                If tekrar > 0 Then
                    alv.FirstVisibleRow = application.Min(r, application.Max(0, n - CLng(alv.VisibleRowCount)))
                    alv.SetCurrentCell r, ids(5)
                    DoEvents: MS_Sleep 30
                End If
                For c = 1 To 11
                    If Len(ids(c)) > 0 Then
                        Err.Clear
                        satir(c) = CStr(alv.GetCellValue(r, ids(c)))
                        If Err.Number <> 0 Then hataNo = Err.Number: hataMetni = Err.Description: Exit For
                    End If
                Next c
                On Error GoTo 0
                If hataNo = 0 Then
                    ozet = (Len(Trim$(CStr(satir(1)))) = 0 And Len(Trim$(CStr(satir(5)))) = 0)
                    If ozet Then
                        tamam = (Len(Trim$(CStr(satir(7)))) > 0 Or Len(Trim$(CStr(satir(11)))) > 0)
                    Else
                        tamam = (Trim$(CStr(satir(1))) = "7155" And Len(Trim$(CStr(satir(5)))) > 0)
                    End If
                End If
                If tamam Then Exit For
            Next tekrar
            If Not tamam Then Err.Raise vbObjectError + 1927, , "ALV satýrý eksiksiz okunamadý: " & CStr(r + 1) & " / " & hataMetni
            If Not ozet Then
                adet = adet + 1
                For c = 1 To 11: buf(adet, c) = satir(c): Next c
                buf(adet, 7) = MS_Sayi(buf(adet, 7))
                If fiyat Then
                    If Len(Trim$(CStr(buf(adet, 11)))) > 0 Then buf(adet, 11) = MS_Sayi(buf(adet, 11))
                Else
                    buf(adet, 10) = CStr(satir(10))
                    MS_SapDeger buf, adet
                End If
                For Each x In Array(1, 2, 4, 5, 8, 9)
                    buf(adet, CLng(x)) = Trim$(CStr(buf(adet, CLng(x))))
                Next x
            End If
        Next r
        MS_Durum asama & " | ALV " & CStr(bitis + 1) & "/" & CStr(n)
    Next bas
    If adet = 0 Then Err.Raise vbObjectError + 1928, , "ALV'de stok kaydý bulunamadý."
    mStandartSatir = mStandartSatir + adet
    ReDim sonuc(1 To adet, 1 To 14)
    For r = 1 To adet
        For c = 1 To 14: sonuc(r, c) = buf(r, c): Next c
    Next r
    MS_ALV_Oku = sonuc
End Function

'----------------------------------------------------------
' Fiyat listesini SAP'a gitmeden mevcut musteri stoklarina uygular.
'----------------------------------------------------------
Public Sub AKSORA_MUSTERI_FIYATLARI_UYGULA(Optional ByVal sessiz As Boolean = True)
    Dim d As Object, tarih As Date, hMesaj As String, adlar As Variant, i As Long, ws As Worksheet
    Dim son As Long, a As Variant, r As Long, c As Long, cikti() As Variant, korumali As Boolean
    Dim ev As Boolean, sc As Boolean, wsF As Worksheet
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    Set d = MF_FIYAT_SOZLUGU(tarih, hMesaj)
    If d Is Nothing Then Err.Raise vbObjectError + 1940, , hMesaj
    application.EnableEvents = False: application.ScreenUpdating = False
    adlar = Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
    For i = 0 To 2
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i)))
        son = AKSORA_M_SON(ws)
        If son >= 7 Then
            a = ws.Range("A7:N" & son).Value2
            For r = 1 To UBound(a, 1)
                If IsError(a(r, 7)) Then a(r, 7) = 0
                If VarType(a(r, 7)) = vbString Or IsEmpty(a(r, 7)) Then a(r, 7) = Val(CStr(a(r, 7)))
            Next r
            MS_FiyatUygula a, d, Format$(tarih, "dd.mm.yyyy")
            ReDim cikti(1 To UBound(a, 1), 1 To 4)
            For r = 1 To UBound(a, 1)
                For c = 1 To 4: cikti(r, c) = a(r, 10 + c): Next c
            Next r
            korumali = ws.ProtectContents
            If korumali Then ws.Unprotect ""
            ws.Range("K7").Resize(UBound(a, 1), 4).Value2 = cikti
            ws.Range("K7:L" & son).NumberFormat = "#,##0.00 ""TL"""
            If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
                UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next i
    Set wsF = ThisWorkbook.Worksheets("_FIYAT_LISTESI")
    wsF.Range("I4").Value2 = "Stoklara uygulandý"
    wsF.Range("J4").Value2 = wsF.Range("J3").Value2
    wsF.Range("J4").NumberFormat = "dd.mm.yyyy hh:mm:ss"
Cikis:
    application.EnableEvents = ev: application.ScreenUpdating = sc
    Exit Sub
Hata:
    If Not sessiz Then MsgBox "Fiyatlar stoklara uygulanamadý: " & Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

' Yeni yuklenen fiyat listesi stoklara henuz uygulanmadiysa uygular (acilista).
Public Sub AKSORA_MUSTERI_FIYAT_KONTROL()
    Dim wsF As Worksheet
    On Error Resume Next
    Set wsF = ThisWorkbook.Worksheets("_FIYAT_LISTESI")
    If wsF Is Nothing Then Exit Sub
    If CStr(wsF.Range("J4").Value2) <> CStr(wsF.Range("J3").Value2) Then AKSORA_MUSTERI_FIYATLARI_UYGULA True
End Sub

'==========================================================
' HIZLI OKUMA: ALV listesi yerel dosyaya (sekmeyle ayrilmis metin)
' aktarilir ve tek seferde okunur. Herhangi bir adim dogrulanamazsa
' Empty doner; cagiran hucre hucre okumaya devam eder.
'==========================================================
Private Function MS_DosyaMetni(ByVal yol As String, ByVal karakter As String) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = karakter: st.Open
    st.LoadFromFile yol
    MS_DosyaMetni = st.ReadText(-1)
    st.Close
End Function

Private Function MS_DosyaBom(ByVal yol As String) As String
    Dim st As Object, b As Variant
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.LoadFromFile yol
    If st.Size >= 3 Then
        b = st.Read(3)
        If b(0) = &HEF And b(1) = &HBB And b(2) = &HBF Then MS_DosyaBom = "utf-8"
        If b(0) = &HFF And b(1) = &HFE Then MS_DosyaBom = "unicode"
    End If
    st.Close
End Function

Private Function MS_ALV_Dosya(ByVal ses As Object, ByVal alv As Object, ByVal kolonlar As Object, _
    ByRef ids() As String, ByVal n As Long, ByVal fiyat As Boolean) As Variant
    Dim klasor As String, dosya As String, yol As String, o As Object, bas As Double
    Dim boy As Long, eskiBoy As Long, utf As Boolean, deneme As Variant, karakter As String
    Dim metin As String, satirlar As Variant, alanlar As Variant, kc As Long, ofs As Long
    Dim baslik As Long, i As Long, j As Long, c As Long, r As Long, adet As Long, veri As Long
    Dim pos(1 To 11) As Long, sira As Object, buf() As Variant, sonuc() As Variant
    Dim satir(1 To 14) As Variant, ozet As Boolean, x As Variant, k As Long, dogru As Boolean
    Dim kontrol As Variant, a As String, b As String, y As Variant, rb As Object, sonVeri As Long
    Dim pen As Object, yolAlan As Object, radyolar As Collection
    On Error GoTo Hata
    mHizliSebep = "kolon eþleþmesi"
    kc = CLng(kolonlar.Count)
    Set sira = CreateObject("Scripting.Dictionary")
    For i = 0 To kc - 1: sira(CStr(kolonlar.Item(i))) = i: Next i
    For c = 1 To 11
        pos(c) = -1
        If Len(ids(c)) > 0 Then
            If Not sira.Exists(ids(c)) Then mHizliSebep = "kolon sýrasý (" & ids(c) & ")": Exit Function
            pos(c) = sira(ids(c))
        End If
    Next c
    klasor = Environ$("TEMP")
    If Right$(klasor, 1) <> "\" Then klasor = klasor & "\"
    dosya = "AKSORA_7155_" & Format$(Now, "yyyymmddhhnnss") & "_" & CStr(Int(Timer * 100) Mod 100000) & ".txt"
    yol = klasor & dosya
    If Len(Dir$(yol)) > 0 Then Kill yol
    ' Liste > Disa aktar > Yerel dosya > Sekmeli metin (pencereler adla bulunur)
    mHizliSebep = "dýþa aktar menüsü"
    MS_Teshis "---- " & Format$(Now, "dd.mm.yyyy hh:nn:ss") & " | Dýþa aktarma | ALV satýr " & CStr(n)
    If Not MS_DisaAktarAc(ses, alv) Then mHizliSebep = "dýþa aktarma penceresi açýlmadý": GoTo Temizle
    Set pen = MS_PencereBekle(ses, 10)
    If pen Is Nothing Then mHizliSebep = "dýþa aktarma penceresi açýlmadý": GoTo Temizle
    MS_Teshis "Pencere: " & pen.Text
    MS_Agac pen, 0
    mHizliSebep = "biçim penceresi"
    Set yolAlan = MS_AlanBul(pen, "DY_PATH")
    If yolAlan Is Nothing Then
        Set radyolar = New Collection
        MS_Topla pen, "GuiRadioButton", radyolar
        Set o = Nothing
        For Each rb In radyolar
            a = AKFIX_NORMAL(CStr(rb.Text))
            If InStr(a, "TAB") > 0 Or InStr(a, "SEKME") > 0 Then Set o = rb: Exit For
        Next rb
        If o Is Nothing And radyolar.Count >= 2 Then Set o = radyolar(2)
        If o Is Nothing Then GoTo Temizle
        MS_Teshis "Seçilen biçim: " & o.Text
        o.Select
        ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
        Set pen = MS_PencereBekle(ses, 10)
        If pen Is Nothing Then mHizliSebep = "dosya penceresi açýlmadý": GoTo Temizle
        MS_Teshis "Pencere: " & pen.Text
        MS_Agac pen, 0
        Set yolAlan = MS_AlanBul(pen, "DY_PATH")
    End If
    mHizliSebep = "dosya penceresi"
    If yolAlan Is Nothing Then GoTo Temizle
    yolAlan.Text = klasor
    Set o = MS_AlanBul(pen, "DY_FILENAME")
    If o Is Nothing Then GoTo Temizle
    o.Text = dosya
    Set o = MS_AlanBul(pen, "DY_FILE_ENCODING")
    If Not o Is Nothing Then o.Text = "4110": utf = True
    ses.FindById("wnd[1]/tbar[0]/btn[0]").Press
    mHizliSebep = "dosya oluþmadý"
    bas = Timer: eskiBoy = -1
    Do
        DoEvents: MS_Sleep 200
        If MS_Gecen(bas) > 120 Then GoTo Temizle
        If Len(Dir$(yol)) > 0 Then
            boy = FileLen(yol)
            If boy > 0 And boy = eskiBoy And Not ses.Busy Then Exit Do
            eskiBoy = boy
        End If
    Loop
    karakter = MS_DosyaBom(yol)
    If Len(karakter) > 0 Then
        deneme = Array(karakter)
    ElseIf utf Then
        deneme = Array("utf-8", "windows-1254")
    Else
        deneme = Array("windows-1254", "utf-8")
    End If
    For Each x In deneme
        metin = MS_DosyaMetni(yol, CStr(x))
        metin = Replace(metin, vbCrLf, vbLf): metin = Replace(metin, vbCr, vbLf)
        If Left$(metin, 1) = ChrW(&HFEFF) Then metin = Mid$(metin, 2)
        satirlar = Split(metin, vbLf)
        ' Baslik: ilk dolu satir; alan sayisi ALV kolon sayisiyla eslesmeli.
        baslik = -1
        For i = 0 To UBound(satirlar)
            If Len(Trim$(Replace(satirlar(i), vbTab, ""))) > 0 Then baslik = i: Exit For
        Next i
        mHizliSebep = "dosya baþlýðý"
        If baslik < 0 Then GoTo Temizle
        alanlar = Split(satirlar(baslik), vbTab)
        If UBound(alanlar) + 1 = kc Then
            ofs = 0
        ElseIf UBound(alanlar) + 1 >= kc + 1 And Len(Trim$(alanlar(0))) = 0 Then
            ofs = 1
        Else
            mHizliSebep = "kolon sayýsý (dosya " & CStr(UBound(alanlar) + 1) & " / SAP " & CStr(kc) & ")"
            GoTo Temizle
        End If
        veri = 0
        For i = baslik + 1 To UBound(satirlar)
            If Len(Trim$(Replace(satirlar(i), vbTab, ""))) > 0 Then veri = veri + 1
        Next i
        If veri > n Or veri < n - 5 Then
            mHizliSebep = "satýr sayýsý (dosya " & CStr(veri) & " / SAP " & CStr(n) & ")"
            GoTo Temizle
        End If
        ' Ornek satirlar SAP ekranindaki degerlerle birebir karsilastirilir.
        dogru = True
        kontrol = Array(0, veri \ 2, veri - 1)
        r = -1: k = 0
        For i = baslik + 1 To UBound(satirlar)
            If Len(Trim$(Replace(satirlar(i), vbTab, ""))) > 0 Then
                r = r + 1
                For Each y In kontrol
                    If CLng(y) = r Then
                        alanlar = Split(satirlar(i), vbTab)
                        alv.FirstVisibleRow = application.Min(r, application.Max(0, n - CLng(alv.VisibleRowCount)))
                        DoEvents
                        For c = 1 To 11
                            If pos(c) >= 0 Then
                                a = vbNullString
                                If pos(c) + ofs <= UBound(alanlar) Then a = Trim$(alanlar(pos(c) + ofs))
                                b = Trim$(CStr(alv.GetCellValue(r, ids(c))))
                                If c = 5 Then
                                    a = MS_Malzeme(a): b = MS_Malzeme(b)
                                ElseIf (c = 7 Or c = 11) And Len(a) > 0 And Len(b) > 0 Then
                                    a = CStr(MS_Sayi(a)): b = CStr(MS_Sayi(b))
                                End If
                                If a <> b Then
                                    dogru = False
                                    mHizliSebep = "örnek kontrol (" & CStr(x) & ", satýr " & CStr(r + 1) & ", alan " & CStr(c) & ": '" & Left$(a, 30) & "' <> '" & Left$(b, 30) & "')"
                                End If
                            End If
                        Next c
                    End If
                Next y
            End If
        Next i
        If dogru Then Exit For
    Next x
    If Not dogru Then GoTo Temizle
    mHizliSebep = "veri satýrý"
    ReDim buf(1 To n, 1 To 14)
    For i = baslik + 1 To UBound(satirlar)
        If Len(Trim$(Replace(satirlar(i), vbTab, ""))) > 0 Then
            alanlar = Split(satirlar(i), vbTab)
            For c = 1 To 14: satir(c) = Empty: Next c
            For c = 1 To 11
                If pos(c) >= 0 Then
                    If pos(c) + ofs <= UBound(alanlar) Then satir(c) = alanlar(pos(c) + ofs) Else satir(c) = vbNullString
                End If
            Next c
            ozet = (Len(Trim$(CStr(satir(1)))) = 0 And Len(Trim$(CStr(satir(5)))) = 0)
            If Not ozet Then
                If Trim$(CStr(satir(1))) <> "7155" Or Len(Trim$(CStr(satir(5)))) = 0 Then GoTo Temizle
                adet = adet + 1
                For c = 1 To 11: buf(adet, c) = satir(c): Next c
                buf(adet, 7) = MS_Sayi(buf(adet, 7))
                If fiyat Then
                    If Len(Trim$(CStr(buf(adet, 11)))) > 0 Then buf(adet, 11) = MS_Sayi(buf(adet, 11))
                Else
                    buf(adet, 10) = Trim$(CStr(satir(10)))
                    MS_SapDeger buf, adet
                End If
                For Each x In Array(1, 2, 3, 4, 5, 6, 8, 9)
                    buf(adet, CLng(x)) = Trim$(CStr(buf(adet, CLng(x))))
                Next x
            End If
        End If
    Next i
    If adet = 0 Then GoTo Temizle
    ReDim sonuc(1 To adet, 1 To 14)
    For r = 1 To adet
        For c = 1 To 14: sonuc(r, c) = buf(r, c): Next c
    Next r
    On Error Resume Next
    Kill yol
    On Error GoTo 0
    mHizliSebep = vbNullString
    MS_ALV_Dosya = sonuc
    Exit Function
Hata:
    mHizliSebep = mHizliSebep & " / " & Err.Description
    Resume Temizle
Temizle:
    On Error Resume Next
    Set o = Nothing
    Set o = ses.FindById("wnd[1]", False)
    If Not o Is Nothing Then ses.FindById("wnd[1]/tbar[0]/btn[12]").Press
    Set o = Nothing
    Set o = ses.FindById("wnd[1]", False)
    If Not o Is Nothing Then o.Close
    If Len(yol) > 0 Then
        If Len(Dir$(yol)) > 0 Then Kill yol
    End If
    MS_ALV_Dosya = Empty
End Function
