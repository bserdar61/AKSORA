Attribute VB_Name = "modAKSORA_IL_AYARLARI"
Option Explicit
' AKSORA il kodlari. Yalnizca okuma/dogrulama; mevcut makrolari ve gorselleri degistirmez.
' Kaynak: kullanicinin 30.09.2026 tarihinde teyit ettigi SAP kodlari.
' Gorsel dosyasi: AKSORA klasorunde GORSELLER\BANNER, GORSELLER\KAPAK, GORSELLER\LOGO altinda
' (eski duzende Excel'in yaninda) aranir. Ayni adli .jpg varsa o kullanilir (calisma kitabina ~5 kat kucuk gomulur).
Public Function AKSORA_GORSEL(ByVal ad As String) As String
    Dim kok As String, jpg As String, ps As String, alt As Variant, k As String
    ps = application.PathSeparator
    kok = ThisWorkbook.Path & ps
    jpg = Left$(ad, InStrRev(ad, ".")) & "jpg"
    On Error Resume Next
    For Each alt In Array("", "GORSELLER" & ps & "BANNER" & ps, "GORSELLER" & ps & "KAPAK" & ps, _
        "GORSELLER" & ps & "LOGO" & ps, "GORSELLER" & ps)
        k = kok & CStr(alt)
        If Len(Dir$(k & jpg)) > 0 Then AKSORA_GORSEL = k & jpg: Exit Function
        If Len(Dir$(k & ad)) > 0 Then AKSORA_GORSEL = k & ad: Exit Function
    Next alt
    On Error GoTo 0
    AKSORA_GORSEL = kok & ad
End Function

' Teshis ve yedek dosyalari: AKSORA klasorunde KAYITLAR varsa oraya, yoksa Excel'in yanina.
Public Function AKSORA_KAYIT_KLASORU() As String
    Dim k As String
    k = ThisWorkbook.Path
    If Len(k) = 0 Or Left$(k, 4) = "http" Then AKSORA_KAYIT_KLASORU = Environ$("TEMP"): Exit Function
    On Error Resume Next
    If Len(Dir$(k & application.PathSeparator & "KAYITLAR", vbDirectory)) > 0 Then k = k & application.PathSeparator & "KAYITLAR"
    On Error GoTo 0
    AKSORA_KAYIT_KLASORU = k
End Function

Public Function AKSORA_AKTIF_IL() As String
    Dim s As String
    s = AKFIX_NORMAL(CStr(ThisWorkbook.Worksheets("KAPAK").Range("AZ1").Value2))
    s = Replace(s, "Ä°", "I"): s = Replace(s, "Ãœ", "U")
    Select Case s
        Case "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN", "ARTVIN"
            AKSORA_AKTIF_IL = s
        Case Else
            Err.Raise vbObjectError + 510, "AKSORA_AKTIF_IL", "Gecersiz il: " & s
    End Select
End Function

Public Function AKSORA_IL_TESIS(ByVal il As String, ByVal sistem As String) As String
    Dim cdp As String, epp As String
    Select Case UCase$(Trim$(il))
        Case "RIZE": cdp = "7151": epp = "7051"
        Case "TRABZON": cdp = "7152": epp = "7052"
        Case "GUMUSHANE": cdp = "7153": epp = "7053"
        Case "GIRESUN": cdp = "7154": epp = "7054"
        Case "ARTVIN": cdp = "7155": epp = "7055"
        Case Else: Err.Raise vbObjectError + 511, "AKSORA_IL_TESIS", "Gecersiz il: " & il
    End Select
    Select Case UCase$(Trim$(sistem))
        Case "CDP": AKSORA_IL_TESIS = cdp
        Case "EPP": AKSORA_IL_TESIS = epp
        Case Else: Err.Raise vbObjectError + 512, "AKSORA_IL_TESIS", "Gecersiz sistem: " & sistem
    End Select
End Function

Public Function AKSORA_IL_YUKLENICI(ByVal il As String) As Variant
    Select Case UCase$(Trim$(il))
        Case "RIZE": AKSORA_IL_YUKLENICI = Array("BO13", "M103", "M104", "M105", "M106", "M124", "M132", "M154", "Z007")
        Case "TRABZON": AKSORA_IL_YUKLENICI = Array("BO13", "M124", "M125", "M126", "M127", "M128", "M129", "M130", "M131", "M132", "Z007")
        Case "GUMUSHANE": AKSORA_IL_YUKLENICI = Array("BO13", "M098", "M099", "M132", "Z007")
        Case "GIRESUN": AKSORA_IL_YUKLENICI = Array("BO13", "M103", "M104", "M105", "M106", "M107", "M132", "Z007")
        Case "ARTVIN": AKSORA_IL_YUKLENICI = Array("BO13", "M122", "M123", "M124")
        Case Else: Err.Raise vbObjectError + 513, "AKSORA_IL_YUKLENICI", "Gecersiz il: " & il
    End Select
End Function

Public Function AKSORA_IL_DEPO_GECERLI(ByVal depo As String, ByVal tur As String, ByVal il As String) As Boolean
    Dim kod As String, n As Long, a As Variant, x As Variant
    kod = UCase$(Trim$(depo))
    ' Ilce tarafina CDP ve EPP icin PB01-PB999 (planli bakim) depolari da dahildir.
    If UCase$(Trim$(tur)) = "ILCE" And Left$(kod, 2) = "PB" Then
        If (Len(kod) = 4 Or Len(kod) = 5) And Mid$(kod, 3) Like String$(Len(kod) - 2, "#") Then
            n = CLng(Mid$(kod, 3))
            AKSORA_IL_DEPO_GECERLI = (n >= 1 And n <= 999)
        End If
        Exit Function
    End If
    Select Case UCase$(Trim$(tur))
        Case "NORMAL", "ILCE"
            If Len(kod) <> 4 Then Exit Function
            If UCase$(tur) = "NORMAL" And Left$(kod, 1) <> "D" Then Exit Function
            If UCase$(tur) = "ILCE" And Left$(kod, 1) <> "K" Then Exit Function
            If Not IsNumeric(Mid$(kod, 2)) Then Exit Function
            n = CLng(Mid$(kod, 2))
            AKSORA_IL_DEPO_GECERLI = (n >= 1 And n <= 999)
        Case "YUKLENICI"
            a = AKSORA_IL_YUKLENICI(il)
            For Each x In a
                If kod = CStr(x) Then AKSORA_IL_DEPO_GECERLI = True: Exit Function
            Next x
    End Select
End Function
Public Function AKSORA_CANLI_STOK_IL() As String
    Dim adlar As Variant, iller As Variant, veri As Variant
    Dim ws As Worksheet, cdpMap As Object, eppMap As Object
    Dim i As Long, r As Long, son As Long, adet As Long
    Dim il As Variant, bulunan As String, satirIl As String, tesis As String
    adlar = Array("CDP_STOK", "EPP_STOK", "STOKLAR", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
    iller = Array("ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
    Set cdpMap = CreateObject("Scripting.Dictionary")
    Set eppMap = CreateObject("Scripting.Dictionary")
    For Each il In iller
        cdpMap(AKSORA_IL_TESIS(CStr(il), "CDP")) = CStr(il)
        eppMap(AKSORA_IL_TESIS(CStr(il), "EPP")) = CStr(il)
    Next il
    For i = 0 To 4
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i)))
        son = AKSORA_SON_SATIR(ws, "G")
        adet = 0
        If son > 250000 Then Err.Raise vbObjectError + 770, "AKSORA_CANLI_STOK_IL", "Beklenmeyen stok aralýðý: " & ws.Name
        If son >= 7 Then
            veri = ws.Range("A7:G" & son).Value2
            For r = 1 To UBound(veri, 1)
                If IsError(veri(r, 7)) Or IsError(veri(r, 1)) Then Err.Raise vbObjectError + 771, , "Hatalý stok hücresi: " & ws.Name & ", satýr " & r + 6
                If Len(Trim$(CStr(veri(r, 7)))) > 0 Then
                    adet = adet + 1
                    tesis = Trim$(CStr(veri(r, 1)))
                    satirIl = ""
                    If i <> 1 And cdpMap.Exists(tesis) Then satirIl = cdpMap(tesis)
                    If i <> 0 And eppMap.Exists(tesis) Then satirIl = eppMap(tesis)
                    If Len(satirIl) = 0 Then Err.Raise vbObjectError + 771, "AKSORA_CANLI_STOK_IL", "Tesis kodu geçersiz: " & ws.Name & ", satýr " & r + 6
                    If Len(bulunan) = 0 Then bulunan = satirIl
                    If bulunan <> satirIl Then Err.Raise vbObjectError + 772, "AKSORA_CANLI_STOK_IL", _
                        "Stok sayfalarýnda farklý illerin verileri var." & vbCrLf & _
                        "Sayfa: " & ws.Name & " / Satýr: " & (r + 6) & _
                        " / Tesis: " & tesis & " / Beklenen il: " & bulunan & " / Bulunan il: " & satirIl
                End If
            Next r
        End If
        If adet = 0 Then Err.Raise vbObjectError + 773, "AKSORA_CANLI_STOK_IL", "Stok sayfasý boþ: " & ws.Name
    Next i
    AKSORA_CANLI_STOK_IL = bulunan
End Function

Public Function AKSORA_SAP_GUNCELLEME_IZNI() As Boolean
    Dim secilen As String, yuklu As String
    If AKSORA_V11_MESGUL Then Exit Function
    On Error GoTo Hata

    secilen = AKSORA_AKTIF_IL()
    yuklu = AKSORA_CANLI_STOK_IL()

    If secilen <> yuklu Then
        MsgBox "Seçilen il: " & secilen & vbCrLf & _
               "Yüklü stoklarýn ili: " & yuklu & vbCrLf & vbCrLf & _
               "Önce il seçimi ile yüklü stoklarý eþleþtirin.", _
               vbExclamation, "AKSORA"
        Exit Function
    End If
    
           If yuklu <> "ARTVIN" And yuklu <> "RIZE" And _
   yuklu <> "TRABZON" And yuklu <> "GUMUSHANE" And _
   yuklu <> "GIRESUN" Then
        MsgBox "Bu ilin SAP güncelleme baðlantýsý henüz tamamlanmadý." & _
               vbCrLf & "Güncelleme baþlatýlmadý.", _
               vbExclamation, "AKSORA"
        Exit Function
    End If

    AKSORA_SAP_GUNCELLEME_IZNI = True
    Exit Function

Hata:
    MsgBox "Güncelleme baþlatýlmadý:" & vbCrLf & _
           Err.Description, vbExclamation, "AKSORA"
End Function
Public Function AKSORA_HAFIZA_ADI(ByVal tur As String) As String
    Dim il As String
    Dim ad As String

    il = AKSORA_CANLI_STOK_IL()
    tur = UCase$(Trim$(tur))

    Select Case tur
        Case "INDEX", "DELTA", "CURRENT"
        Case Else
            Err.Raise vbObjectError + 790, "AKSORA_HAFIZA_ADI", _
                      "Geçersiz stok geçmiþi türü: " & tur
    End Select

    ad = "_STOK_HAFIZA_" & tur

    If il <> "ARTVIN" Then ad = ad & "_" & il

    AKSORA_HAFIZA_ADI = ad
End Function
Public Sub AKSORA_STOK_IL_BASLIKLARI_YENILE()
    Dim adlar As Variant, ad As Variant
    Dim ws As Worksheet
    Dim il As String, metin As String
    Dim cizim As Boolean, icerik As Boolean, senaryo As Boolean
    Dim filtre As Boolean, sirala As Boolean, korumali As Boolean
    Dim hataNo As Long, hataMetni As String

    On Error GoTo Hata
    il = AKSORA_CANLI_STOK_IL()

    Select Case il
        Case "ARTVIN": metin = "Artvin Ýl Müdürlüðü"
        Case "RIZE": metin = "Rize Ýl Müdürlüðü"
        Case "TRABZON": metin = "Trabzon Ýl Müdürlüðü"
        Case "GUMUSHANE": metin = "Gümüþhane Ýl Müdürlüðü"
        Case "GIRESUN": metin = "Giresun Ýl Müdürlüðü"
    End Select

    adlar = Array("CDP_STOK", "EPP_STOK", "STOKLAR", _
        ChrW(304) & "L" & ChrW(199) & "E_STOK", _
        "Y" & ChrW(220) & "KLEN" & ChrW(304) & "C" & ChrW(304) & "_STOK")

    For Each ad In adlar
        Set ws = ThisWorkbook.Worksheets(CStr(ad))

        cizim = ws.ProtectDrawingObjects
        icerik = ws.ProtectContents
        senaryo = ws.ProtectScenarios
        filtre = ws.Protection.AllowFiltering
        sirala = ws.Protection.AllowSorting
        korumali = cizim Or icerik Or senaryo

        If korumali Then ws.Unprotect

        ws.Range("A5").MergeArea.Cells(1, 1).Value = metin

        If korumali Then
            ws.Protect Password:="", _
                DrawingObjects:=cizim, Contents:=icerik, _
                Scenarios:=senaryo, UserInterfaceOnly:=True, _
                AllowFiltering:=filtre, AllowSorting:=sirala
        End If

        korumali = False
    Next ad

    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description
    On Error Resume Next

    If korumali Then
        ws.Protect Password:="", _
            DrawingObjects:=cizim, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True, _
            AllowFiltering:=filtre, AllowSorting:=sirala
    End If

    On Error GoTo 0
    Err.Raise hataNo, "AKSORA_STOK_IL_BASLIKLARI_YENILE", hataMetni
End Sub
Public Sub AKSORA_PANEL_IL_YAZILARI_YENILE()
    Dim ws As Worksheet
    Dim il As String, mudurluk As String

    il = AKSORA_CANLI_STOK_IL()

    Select Case il
        Case "ARTVIN": mudurluk = "Artvin Ýl Müdürlüðü"
        Case "RIZE": mudurluk = "Rize Ýl Müdürlüðü"
        Case "TRABZON": mudurluk = "Trabzon Ýl Müdürlüðü"
        Case "GUMUSHANE": mudurluk = "Gümüþhane Ýl Müdürlüðü"
        Case "GIRESUN": mudurluk = "Giresun Ýl Müdürlüðü"
    End Select

    Set ws = ThisWorkbook.Worksheets("ANA PANEL")

    ws.Range("B28").Value = _
        "ÇORUH ELEKTRÝK DAÐITIM A.Þ. | " & mudurluk
    ws.Range("B29").MergeArea.ClearContents

    ws.Shapes("AKSORA15_ALT_KURUM").TextFrame2.TextRange.Text = _
        "ÇORUH ELEKTRÝK DAÐITIM A.Þ. | " & mudurluk

    ws.Shapes("AKSORA15_ALT_IMZA").TextFrame2.TextRange.Text = ""

    ws.Shapes("AKSORA15_KART_BASLIK_B").TextFrame2.TextRange.Text = _
        "CDP / " & AKSORA_IL_TESIS(il, "CDP")

    ws.Shapes("AKSORA15_KART_BASLIK_G").TextFrame2.TextRange.Text = _
        "EPP / " & AKSORA_IL_TESIS(il, "EPP")

    AKSORA_STOK_IL_BASLIKLARI_YENILE
End Sub
Public Sub AKSORA_PANEL_IL_GORSELI_YENILE()
    Dim ws As Worksheet
    Dim eski As Shape, yeni As Shape, sh As Shape
    Dim il As String, yol As String, ad As String
    Dim h0 As Double
    Dim cizim As Boolean, icerik As Boolean, senaryo As Boolean
    Dim filtre As Boolean, sirala As Boolean, korumali As Boolean
    Dim hataNo As Long, hataMetni As String

    On Error GoTo Hata

    il = AKSORA_CANLI_STOK_IL()
    Set ws = ThisWorkbook.Worksheets("ANA PANEL")
    Set eski = ws.Shapes("AKSORA15_BANNER")

    ad = "AKSORA_IL_BANNER_" & il

    On Error Resume Next
    Set yeni = ws.Shapes(ad)
    On Error GoTo Hata

    ' Il gorseli tek dosyadan ya da birlesik AKSORA_MALZEME_PANEL_BANTLARI.jpg dosyasindan gelir.

    cizim = ws.ProtectDrawingObjects
    icerik = ws.ProtectContents
    senaryo = ws.ProtectScenarios
    filtre = ws.Protection.AllowFiltering
    sirala = ws.Protection.AllowSorting
    korumali = cizim Or icerik Or senaryo

    If korumali Then ws.Unprotect

    If yeni Is Nothing Then
        Set yeni = AKSORA_MALZEME_BANT_EKLE(ws, il, eski.Left, eski.Top)
        yeni.Name = ad
    End If

    With yeni
        .LockAspectRatio = msoFalse
        .Left = eski.Left
        .Top = eski.Top
        .Width = eski.Width
        .Height = eski.Height
        .Placement = xlFreeFloating
        .Locked = True
        .OnAction = ""
        .Visible = msoTrue
        .ZOrder msoBringToFront
    End With

    For Each sh In ws.Shapes
                If Left$(sh.Name, Len("AKSORA_IL_BANNER_")) = "AKSORA_IL_BANNER_" Then
            If sh.Name <> ad Then sh.Visible = msoFalse
        End If
    Next sh

    eski.Visible = msoFalse
    AKSORA_TASARIM_SAYFA ws

    If korumali Then
        ws.Protect Password:="", _
            DrawingObjects:=cizim, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True, _
            AllowFiltering:=filtre, AllowSorting:=sirala
    End If

    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description
    On Error Resume Next

    If korumali Then
        ws.Protect Password:="", _
            DrawingObjects:=cizim, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True, _
            AllowFiltering:=filtre, AllowSorting:=sirala
    End If

    On Error GoTo 0
    Err.Raise hataNo, "AKSORA_PANEL_IL_GORSELI_YENILE", hataMetni
End Sub

' Malzeme ana panel il bandi: AKSORA_PANEL_BANNER_<IL>.png varsa o, yoksa birlesik
' AKSORA_MALZEME_PANEL_BANTLARI.jpg (5 il alt alta: Giresun, Trabzon, Rize, Gumushane, Artvin).
Public Function AKSORA_MALZEME_BANT_EKLE(ByVal ws As Worksheet, ByVal il As String, ByVal x As Double, ByVal y As Double) As Shape
    Dim yol As String, s As Shape, h As Double, i As Long
    yol = AKSORA_GORSEL("AKSORA_PANEL_BANNER_" & il & ".png")
    If Len(Dir$(yol)) > 0 Then
        Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, x, y, -1, -1)
        h = s.Height
        s.PictureFormat.CropTop = h * 225 / 684
        s.PictureFormat.CropBottom = h * 224 / 684
    Else
        yol = AKSORA_GORSEL("AKSORA_MALZEME_PANEL_BANTLARI.jpg")
        If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 850, , _
            "Ýl görseli bulunamadý: AKSORA_MALZEME_PANEL_BANTLARI.jpg Excel dosyasýnýn yanýnda olmalýdýr."
        Select Case UCase$(Trim$(il))
            Case "GIRESUN": i = 0
            Case "TRABZON": i = 1
            Case "RIZE": i = 2
            Case "GUMUSHANE": i = 3
            Case Else: i = 4
        End Select
        Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, x, y, -1, -1)
        h = s.Height
        s.PictureFormat.CropTop = h * (i + 225 / 684) / 5
        s.PictureFormat.CropBottom = h * (4 - i + 224 / 684) / 5
    End If
    Set AKSORA_MALZEME_BANT_EKLE = s
End Function
