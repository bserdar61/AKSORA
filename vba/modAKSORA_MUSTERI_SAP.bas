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
#Else
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

Private Function MS_Norm(ByVal v As Variant) As String
    Dim t As String
    If IsError(v) Then Exit Function
    t = AKFIX_NORMAL(Trim$(CStr(v)))
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
        If CLng(a(r, 14)) > 0 Then eksik = eksik + 1
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
    MS_OleUyariKapat
    AKSORA_MUSTERI_PANEL_YENILE True
    AKSORA_MUSTERI_SERIT "SAP güncellemesi baþlatýlýyor...", 1
    ev = application.EnableEvents: sc = application.ScreenUpdating: hesap = application.Calculation
    On Error GoTo Hata
    ' Fiyat listesi SAP'a gitmeden önce doðrulanýr.
    Set d = MF_FIYAT_SOZLUGU(tarih, hMesaj)
    If d Is Nothing Then Err.Raise vbObjectError + 1930, , hMesaj
    Set ses = SAP_SESSION_AL("CDP", hMesaj)
    If ses Is Nothing Then Err.Raise vbObjectError + 1918, , hMesaj
    adlar = Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
    For i = 1 To 3
        Set sayfa(i) = ThisWorkbook.Worksheets(CStr(adlar(i - 1)))
    Next i
    application.EnableEvents = False: application.Calculation = xlCalculationManual
    AKSORA_V11_MESGUL = True
    MS_Durum "1/2 | CDP + Yüklenici depolarý (E001-E999, D/Z/S/EM, M132)..."
    bir = MS_Sorgu(ses, False)
    MS_Durum "2/2 | Ýlçe depolarý K001-K999 ve üç mal grubu..."
    iki = MS_Sorgu(ses, True)
    MS_KapsamKontrol bir, 1, 3, False
    MS_KapsamKontrol iki, 2, 2, True
    MS_SeriKontrol bir: MS_SeriKontrol iki
    MS_SeriBirlesimKontrol bir, iki
    mFiyatTur = True
    MS_Durum "Fiyat listesi uygulanýyor..."
    eksik = MS_FiyatUygula(bir, d, Format$(tarih, "dd.mm.yyyy"))
    eksik = eksik + MS_FiyatUygula(iki, d, Format$(tarih, "dd.mm.yyyy"))
    yeni(1) = MS_Ayir(bir, 1): yeni(2) = iki: yeni(3) = MS_Ayir(bir, 3)
    For i = 1 To 3
        sonEski(i) = application.Max(7, sayfa(i).Cells(sayfa(i).Rows.Count, "E").End(xlUp).Row)
        eski(i) = sayfa(i).Range("A7:N" & sonEski(i)).Value2
        tarihEski(i) = sayfa(i).Range("C2").Value
        If IsEmpty(yeni(i)) Then adet(i) = 0 Else adet(i) = UBound(yeni(i), 1)
    Next i
    yazildi = True
    For i = 1 To 3: MS_Yaz sayfa(i), yeni(i): Next i
    yazildi = False
    MS_Durum "Tamamlandý | CDP " & adet(1) & " | Ýlçe " & adet(2) & " | Yüklenici " & adet(3) & _
        " | Fiyatý olmayan " & eksik & " | " & AKSORA_SAYI_METNI(MS_Gecen(mBas), "0.00") & " sn"
    ThisWorkbook.names.Add Name:="_AKSORA_MUSTERI_7155_SON_SURE", RefersTo:="=" & Trim$(Str$(Round(MS_Gecen(mBas), 2))), Visible:=False
    AKSORA_MUSTERI_PANEL_YENILE
    AKSORA_MUSTERI_SERIT "Güncelleme tamamlandý  •  " & AKSORA_MUSTERI_OZET(), 2
    MsgBox "Müþteri 7155 güncellendi." & vbCrLf & vbCrLf & _
        "CDP stok satýrý: " & AKSORA_SAYI_METNI(adet(1), "#,##0") & vbCrLf & _
        "Ýlçe stok satýrý: " & AKSORA_SAYI_METNI(adet(2), "#,##0") & vbCrLf & _
        "Yüklenici stok satýrý: " & AKSORA_SAYI_METNI(adet(3), "#,##0") & vbCrLf & vbCrLf & _
        "Deðer kaynaðý: SAP raporu " & AKSORA_SAYI_METNI(mSapSatir, "#,##0") & " satýr  |  Fiyat listesi " & _
            AKSORA_SAYI_METNI(mListeSatir, "#,##0") & " satýr" & vbCrLf & _
        "Fiyatý bulunamayan/belirsiz satýr: " & AKSORA_SAYI_METNI(eksik, "#,##0") & vbCrLf & _
        "Okuma: " & IIf(mStandartSatir = 0, "Hýzlý (dosya)", IIf(mHizliSatir = 0, "Standart", "Karma")) & _
            IIf(Len(mHizliSebep) > 0, "  |  Hýzlý okuma kullanýlamadý: " & mHizliSebep, "") & vbCrLf & _
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
    hizli = MS_ALV_Dosya(ses, alv, kolonlar, ids, n, fiyat)
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
    On Error Resume Next
    alv.PressToolbarContextButton "&MB_EXPORT"
    If Err.Number = 0 Then alv.SelectContextMenuItem "&PC"
    If Err.Number <> 0 Then
        MS_Teshis "Baðlam menüsü: " & Err.Description & " -> %PC deneniyor"
        Err.Clear
        ses.FindById("wnd[0]/tbar[0]/okcd").Text = "%PC"
        ses.FindById("wnd[0]").SendVKey 0
    End If
    Err.Clear
    On Error GoTo Hata
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
