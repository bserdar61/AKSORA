Attribute VB_Name = "modAKSORA_ANA_PANEL"
Option Explicit

' AKSORA - ANA PANEL gorunumu ve kart tasarimi.
' SAP, stok, rapor ve mevcut buton OnAction baglantilarina dokunmaz.
Private Const p As String = "AKSORA15_"
Public AKSORA_PANEL_CANLI As Boolean
Public AKSORA_PANEL_GUNCELLENEN As String

Public Sub AKSORA_ANA_PANEL_UYGULA()
    Dim ws As Worksheet, s As Shape, r As Range
    Dim i As Long, nm As String, f As String, banner As String
    Dim ev As Boolean, scr As Boolean, hataNo As Long, hataAciklama As String
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("ANA PANEL")
    f = AKSORA_GORSEL("AKSORA_PANEL_TEK_PARCA_V7.png")
    banner = AKSORA_GORSEL("AKSORA_PANEL_BANNER.png")
    If Len(Dir$(f)) = 0 Then
        MsgBox "AKSORA_PANEL_TEK_PARCA_V7.png bulunamadi. Resmi Excel ile ayni klasore koy.", vbExclamation, "AKSORA"
        Exit Sub
    End If
    ev = application.EnableEvents: scr = application.ScreenUpdating
    application.EnableEvents = False: application.ScreenUpdating = False

    ' Sadece onceki AKSORA dekorlarini sil; gercek islev butonlari korunur.
    For i = ws.Shapes.Count To 1 Step -1
        nm = ws.Shapes(i).Name
        If Left$(nm, 8) = "AKSORA3_" Or Left$(nm, 8) = "AKSORA4_" Or _
           Left$(nm, 8) = "AKSORA5_" Or Left$(nm, 8) = "AKSORA6_" Or Left$(nm, 8) = "AKSORA7_" Or Left$(nm, 8) = "AKSORA8_" Or Left$(nm, 8) = "AKSORA9_" Or Left$(nm, 9) = "AKSORA10_" Or Left$(nm, 9) = "AKSORA11_" Or Left$(nm, 9) = "AKSORA12_" Or Left$(nm, 9) = "AKSORA13_" Or Left$(nm, 9) = "AKSORA14_" Or Left$(nm, 9) = "AKSORA15_" Or _
           Left$(nm, 10) = "AKSORA_BG_" Then
            ws.Shapes(i).Delete
        End If
    Next i

    ' Onceki denemelerden kalmis ve AB disina tasmis islevsiz dekorlari temizle.
    ' Makroya bagli butonlara dokunma.
    For i = ws.Shapes.Count To 1 Step -1
        Set s = ws.Shapes(i)
        If s.Left >= ws.Range("AC1").Left - 1 Then
            If Left$(UCase$(s.Name), 6) = "AKSORA" Then
                If Len(s.OnAction) = 0 Then s.Delete
            End If
        End If
    Next i

    ' Tek bir resim: Excel arka plan tekrarini kapat, 31. satir sonrasini boya.
    On Error Resume Next
    ws.SetBackgroundPicture ""
    On Error GoTo Hata
    ' AB sutununun saginda eski tasarimdan kalan renkleri tamamen kaldir.
    ' Yalnizca ANA PANEL; diger sayfalarin bicimine dokunulmaz.
    With ws.Range("AC1:AZ80")
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
    ws.Range("A1:AB80").Interior.Color = RGB(236, 248, 254)
    ws.Range("A1:AB80").Font.Name = "Aptos"
    ws.Range("B7:AB27").Interior.Pattern = xlNone
    ws.Range("B7:T13").Interior.Pattern = xlNone
    ws.Range("B7:T13").Borders.LineStyle = xlNone
    ' Eski kart baslik hucrelerinin ustte biraktigi beyaz seridi temizle.
    With ws.Range("B7:T7")
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
    End With
    Set r = ws.Range("B7:AB27")
    Set s = ws.Shapes.AddPicture(f, msoFalse, msoTrue, _
                                 r.Left, r.Top, r.Width, r.Height)
    s.Name = p & "TEK_ARKA_PLAN"
    s.Placement = xlFreeFloating
    s.LockAspectRatio = msoFalse
    s.ZOrder msoSendToBack

    ' Ust baslik - orijinal fotografin tamami, sagda devam eden lacivert.
    Set r = ws.Range("B2:AB5")
    V6_RECT ws, p & "UST_ZEMIN", r, RGB(12, 45, 79), 0, False
    If Len(Dir$(banner)) > 0 Then
        Set r = ws.Range("B2:AB5")
        Set s = ws.Shapes.AddPicture(banner, msoFalse, msoTrue, _
                                    r.Left, r.Top, r.Width, r.Height)
        s.Name = p & "BANNER": s.Placement = xlFreeFloating
    End If
    ' Orijinal fotograflu banner AB sutununa kadar tek parca uzatildi.

    ' Durum seridi: veri mevcut hucrede kalir, sagda il secimi.
    Set r = ws.Range("B6:AB6")
    r.Interior.Color = RGB(255, 235, 210)
    r.Font.Color = RGB(150, 75, 20): r.Font.bold = True
    ws.Range("B6:T6").Font.Size = 9

    ' Kartlar: gercek hucrelerin uzerine resim/sekil koymuyoruz.
    V11_KART ws, "B", "D", "E", "CDP / 7155", True
    V11_KART ws, "G", "I", "J", "EPP / 7055", True
    V11_KART ws, "L", "N", "O", "ILCELER", False
    V11_KART ws, "Q", "S", "T", "YUKLENICI", False

    ' Mevcut butonlarin yeri ve makro baglantisi aynen korunur.
    For Each s In ws.Shapes
        nm = LCase$(s.Name)
        Select Case nm
            Case "btntumstok"
                V11_BUTON s, RGB(0, 110, 94), RGB(64, 238, 183), vbWhite, 15, True
            Case "btnbaglantikontrol", "btnarama", "btnmalzemeara", "btnmalzemearama"
                V11_BUTON s, RGB(241, 251, 255), RGB(99, 201, 235), RGB(14, 62, 100), 10, False
            Case "btncdp", "btnepp", "btnilce", "btnyuk"
                V11_BUTON s, RGB(16, 57, 94), RGB(65, 179, 220), vbWhite, 10, False
            Case Else
                If Left$(nm, 2) = "go" Then
                    V11_BUTON s, RGB(236, 250, 255), RGB(112, 195, 228), RGB(15, 68, 106), 9, False
                End If
        End Select
    Next s

    ' IL DEGISTIR: durum seridiyle ayni hizada, sagda AB sinirinda biter.
    Set r = ws.Range("W6:AB6")
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                              r.Left, r.Top, r.Width, r.Height)
    s.Name = p & "IL_DEGISTIR"
    s.TextFrame2.TextRange.Text = ChrW(8592) & "  IL DEGISTIR"
    V11_BUTON s, RGB(255, 235, 210), RGB(255, 235, 210), RGB(150, 75, 20), 10, False
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
                 "'!AKSORA_ANA_PANEL_IL_DEGISTIR"

    ' Alt bilgi: tum genislik boyunca devam eder.
    Set r = ws.Range("B28:AB29")
    r.Interior.Color = RGB(14, 48, 82): r.Font.Color = vbWhite
    Set r = ws.Range("B30:AB30")
    r.Interior.Color = RGB(217, 240, 245): r.Font.Color = RGB(71, 108, 129)
    V8_ALT_SERIT ws
    V11_PANEL_TARIHI ws
    AKSORA_ANA_PANEL_DURUM_RENK_YENILE

    ws.Shapes(p & "TEK_ARKA_PLAN").ZOrder msoSendToBack
    ' Arka plan resminden sonra tablolar ve butonlar her zaman ustte.
    V8_ONE_GETIR ws
    ' Burada Activate/Select/Zoom YOK: sekme gecisinde titreme olmasin.
    If ActiveSheet Is ws And Not AKSORA_PENCERE_KUCUK(ActiveWindow) Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayWorkbookTabs = True
    End If
Cikis:
    application.EnableEvents = ev: application.ScreenUpdating = scr
    Exit Sub
Hata:
    hataNo = Err.Number: hataAciklama = Err.Description
    MsgBox "AKSORA V15 (" & hataNo & "): " & hataAciklama, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Sub V11_KART(ByVal ws As Worksheet, ByVal sol As String, _
                     ByVal sag As String, ByVal bitis As String, _
                     ByVal baslik As String, ByVal sap As Boolean)
    Dim r As Range, x As Double, y As Double, w As Double, h As Double
    Dim a As Shape, k As Long, rowNo As Long, yy As Double, rh As Double
    Dim labels As Variant, adres As String, valColor As Long
    Set r = ws.Range(sol & "7:" & bitis & "13")
    x = r.Left: y = r.Top + 8: w = r.Width: h = r.Height - 8

    ' Tek parca govde; baslik ile govde arasinda beyaz bosluk YOK.
    Set a = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    a.Name = p & "KART_GOVDE_" & sol
    a.Fill.ForeColor.RGB = RGB(239, 250, 255)
    a.Fill.Transparency = 0
    a.line.ForeColor.RGB = RGB(77, 184, 214)
    a.line.Weight = 1.3
    a.Adjustments.Item(1) = 0.055
    a.Shadow.Visible = msoTrue
    a.Shadow.ForeColor.RGB = RGB(7, 38, 68)
    a.Shadow.Transparency = 0.84
    a.Shadow.Blur = 5
    a.Shadow.OffsetX = 0: a.Shadow.OffsetY = 2
    a.Placement = xlFreeFloating

    ' Koyu baslik govdenin ust kenarina tam oturur.
    Set a = ws.Shapes.AddShape(msoShapeRectangle, x + 1, y + 1, w - 2, 29)
    a.Name = p & "KART_BASLIK_Z_" & sol
    a.Fill.ForeColor.RGB = RGB(9, 39, 68)
    a.line.Visible = msoFalse
    V8_YAZI ws, p & "KART_BASLIK_" & sol, baslik, _
        x + 12, y + 1, w - 45, 28, 11, vbWhite, True, msoAlignLeft
    Set a = ws.Shapes.AddShape(msoShapeOval, x + w - 24, y + 10, 8, 8)
    a.Name = p & "KART_LED_" & sol
    a.Fill.ForeColor.RGB = RGB(72, 240, 164)
    a.line.Visible = msoFalse

    labels = Array("DURUM", "SON GÜNCELLEME", "STOK SATIRI", "STOK DEÐERÝ", "SAP BAÐLANTI")
    rh = (h - 34 - 12) / 5
    For k = 0 To 4
        rowNo = 8 + k
        yy = y + 34 + k * rh
        If k > 0 Then
            Set a = ws.Shapes.AddLine(x + 12, yy - 1, x + w - 12, yy - 1)
            a.Name = p & "KART_CIZGI_" & sol & CStr(k)
            a.line.ForeColor.RGB = RGB(196, 223, 235)
            a.line.Weight = 0.65
        End If
        If Not sap And k = 4 Then
            V8_YAZI ws, p & "KART_ETIKET_" & sol & CStr(k), "STOK GRUBU", _
                x + 10, yy, w - 20, rh, 8, RGB(67, 110, 138), False, msoAlignCenter
        Else
            V8_YAZI ws, p & "KART_ETIKET_" & sol & CStr(k), CStr(labels(k)), _
                x + 11, yy, w * 0.49, rh, 8, RGB(65, 106, 135), False, msoAlignLeft
            adres = sag & CStr(rowNo)
            valColor = RGB(16, 55, 86)
            If k = 0 Then valColor = RGB(16, 55, 86)
            V8_LINK ws, p & "KART_DEGER_" & sol & CStr(k), adres, _
                x + w * 0.49, yy, w * 0.49 - 9, rh, 9, valColor
        End If
    Next k
End Sub

Private Sub V8_LINK(ByVal ws As Worksheet, ByVal ad As String, ByVal adres As String, _
                    ByVal x As Double, ByVal y As Double, ByVal w As Double, _
                    ByVal h As Double, ByVal fs As Single, ByVal fc As Long)
    Dim a As Shape, linked As Boolean
    Set a = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    a.Name = ad
    a.Fill.Visible = msoFalse: a.line.Visible = msoFalse
    a.Placement = xlFreeFloating
    On Error Resume Next
    Err.Clear
    a.DrawingObject.Formula = "='" & Replace(ws.Name, "'", "''") & "'!" & adres
    linked = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
    If Not linked Then a.TextFrame2.TextRange.Text = ws.Range(adres).Text
    With a.TextFrame2
        .MarginLeft = 0: .MarginRight = 2
        .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.ParagraphFormat.Alignment = msoAlignRight
        .TextRange.Font.Name = "Aptos"
        .TextRange.Font.Size = fs
        .TextRange.Font.bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = fc
    End With
End Sub

Private Sub V8_YAZI(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
                    ByVal x As Double, ByVal y As Double, ByVal w As Double, _
                    ByVal h As Double, ByVal fs As Single, ByVal fc As Long, _
                    ByVal bold As Boolean, ByVal hiz As Long)
    Dim a As Shape
    Set a = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    a.Name = ad: a.Fill.Visible = msoFalse: a.line.Visible = msoFalse
    a.TextFrame2.MarginLeft = 2: a.TextFrame2.MarginRight = 2
    a.TextFrame2.MarginTop = 0: a.TextFrame2.MarginBottom = 0
    a.TextFrame2.TextRange.Text = metin
    a.TextFrame2.TextRange.Font.Name = "Aptos"
    a.TextFrame2.TextRange.Font.Size = fs
    a.TextFrame2.TextRange.Font.bold = IIf(bold, msoTrue, msoFalse)
    a.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = fc
    a.TextFrame2.TextRange.ParagraphFormat.Alignment = hiz
    a.TextFrame2.VerticalAnchor = msoAnchorMiddle
    a.Placement = xlFreeFloating
End Sub

Private Sub V8_ALT_SERIT(ByVal ws As Worksheet)
    Dim r As Range, a As Shape
    Set r = ws.Range("B28:AB29")
    Set a = ws.Shapes.AddShape(msoShapeRectangle, r.Left, r.Top, r.Width, r.Height)
    a.Name = p & "ALT_ZEMIN": a.Fill.ForeColor.RGB = RGB(13, 47, 79)
    a.line.Visible = msoFalse
    V8_YAZI ws, p & "ALT_KURUM", "CORUH ELEKTRIK DAGITIM A.S.  |  ARTVIN IL MUDURLUGU", _
        r.Left, r.Top + 2, r.Width, r.Height / 2 - 1, 10, vbWhite, True, msoAlignCenter
    V8_YAZI ws, p & "ALT_IMZA", "Burhan Serdar  |  Elektrik-Elektronik Muhendisi", _
        r.Left, r.Top + r.Height / 2, r.Width, r.Height / 2 - 2, 9, RGB(220, 238, 248), False, msoAlignCenter
    ' Orijinal hucreleri silme: eski footer metni korunur.
End Sub

Private Sub V11_PANEL_TARIHI(ByVal ws As Worksheet)
    Dim r As Range, c As Range, a As Shape, kaynak As String
    Dim metin As String
    Set r = ws.Range("B30:AB30")
    ' Mevcut tarih kaynagini koru. Sadece gorunen metni tek merkez kutuya tasi.
    For Each c In ws.Range("B30:AB31").Cells
        If InStr(1, CStr(c.Value2), "Panel kontrol", vbTextCompare) > 0 Then
            If Len(kaynak) = 0 Then kaynak = c.Address(False, False)
            c.Font.Color = RGB(217, 240, 245)
        End If
    Next c
    Set a = ws.Shapes.AddShape(msoShapeRectangle, r.Left, r.Top, r.Width, r.Height)
    a.Name = p & "TARIH_ZEMIN"
    a.Fill.ForeColor.RGB = RGB(217, 240, 245)
    a.line.Visible = msoFalse
    a.Placement = xlFreeFloating
    If Len(kaynak) > 0 Then
        metin = CStr(ws.Range(kaynak).Value2)
    Else
        metin = "Panel kontrolü: Güncelleme bekleniyor"
    End If
    V8_YAZI ws, p & "PANEL_TARIHI", metin, _
        r.Left, r.Top, r.Width, r.Height, 9, RGB(43, 88, 115), False, msoAlignCenter
End Sub

Private Sub V11_BUTON(ByVal s As Shape, ByVal dolgu As Long, _
                      ByVal kenar As Long, ByVal yazi As Long, _
                      ByVal punto As Single, ByVal ana As Boolean)
    On Error Resume Next
    With s
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = dolgu
        .Fill.Transparency = 0
        .line.Visible = msoTrue
        .line.ForeColor.RGB = kenar
        .line.Weight = IIf(ana, 2, 1.2)
        .Shadow.Visible = msoTrue
        .Shadow.ForeColor.RGB = RGB(9, 46, 79)
        .Shadow.Transparency = 0.79
        .Shadow.Blur = IIf(ana, 8, 5)
        .Shadow.OffsetX = 0
        .Shadow.OffsetY = 2
        .TextFrame2.TextRange.Font.Name = "Aptos Display"
        .TextFrame2.TextRange.Font.Size = punto
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = yazi
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    On Error GoTo 0
End Sub

Private Sub V8_ONE_GETIR(ByVal ws As Worksheet)
    Dim a As Shape
    For Each a In ws.Shapes
        If Left$(a.Name, Len(p)) = p Then
            If a.Name <> p & "TEK_ARKA_PLAN" And a.Name <> p & "UST_ZEMIN" Then
                a.ZOrder msoBringToFront
            End If
        End If
    Next a
    ' Gercek butonlari da arka planin onunde tut.
    For Each a In ws.Shapes
        If Len(a.OnAction) > 0 And Left$(a.Name, Len(p)) <> p Then
            a.ZOrder msoBringToFront
        End If
    Next a
End Sub

Private Sub V6_RECT(ByVal ws As Worksheet, ByVal ad As String, _
                    ByVal r As Range, ByVal renk As Long, _
                    ByVal saydamlik As Double, ByVal yuvarlak As Boolean)
    Dim s As Shape
    If yuvarlak Then
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, r.Left, r.Top, r.Width, r.Height)
    Else
        Set s = ws.Shapes.AddShape(msoShapeRectangle, r.Left, r.Top, r.Width, r.Height)
    End If
    s.Name = ad
    s.Fill.ForeColor.RGB = renk: s.Fill.Transparency = saydamlik
    s.line.Visible = msoFalse: s.Placement = xlFreeFloating
End Sub

Public Sub AKSORA_ANA_PANEL_IL_DEGISTIR()
    ' Kapak kilitlenir: giris kapanir, kapak tek gorunur sayfa olur (baska sayfaya kayma olmaz).
    On Error GoTo Hata
    If AKSORA_V11_MESGUL Then Exit Sub
    AKSORA_KAPAK_KILITLE
    If ActiveSheet.Name <> "KAPAK" Then ThisWorkbook.Worksheets("KAPAK").Activate
    AKSORA_V8_KAPAK_GORUNUM
    Exit Sub
Hata:
    MsgBox "Il secimi: " & Err.Description, vbExclamation, "AKSORA"
End Sub




' Genel durum seridini ve IL DEGISTIR butonunu tek renge getirir.
' Durum: 0-1 gun yesil, 2 gun turuncu, 3+ gun kirmizi.
Public Sub AKSORA_ANA_PANEL_DURUM_RENK_YENILE()
    Dim ws As Worksheet, kodlar As Variant, sutunlar As Variant
    Dim stoklar As Variant, i As Long, gun As Long, enKotu As Long
    Dim kaynak As Worksheet, tarih As Variant, durumRengi As Long
    Dim yaziRengi As Long, buton As Shape, kartYazisi As Shape
    On Error GoTo Cikis
    Set ws = ThisWorkbook.Worksheets("ANA PANEL")
    kodlar = Array("CDP", "EPP", "ILCE", "YUK")
    sutunlar = Array("B", "G", "L", "Q")
    stoklar = Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
    enKotu = 0
    For i = 0 To 3
        gun = 999
        Set kaynak = Nothing
        On Error Resume Next
        Set kaynak = ThisWorkbook.Worksheets(CStr(stoklar(i)))
        On Error GoTo Cikis
        If Not kaynak Is Nothing Then
            tarih = kaynak.Range("C2").Value
            If IsDate(tarih) Then
                gun = DateDiff("d", DateValue(CDate(tarih)), Date)
                If gun < 0 Then gun = 0
            End If
        End If
        If gun > enKotu Then enKotu = gun
        If gun = 0 Then
            durumRengi = RGB(0, 133, 92)
        ElseIf gun = 1 Then
            durumRengi = RGB(230, 130, 20)
        Else
            durumRengi = RGB(190, 55, 55)
        End If
        Set kartYazisi = Nothing
        On Error Resume Next
        Set kartYazisi = ws.Shapes(p & "KART_DEGER_" & CStr(sutunlar(i)) & "0")
        If Not kartYazisi Is Nothing Then
            kartYazisi.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = durumRengi
        End If
                Set kartYazisi = Nothing
        Set kartYazisi = ws.Shapes(p & "KART_LED_" & CStr(sutunlar(i)))

        If Not kartYazisi Is Nothing Then
            With kartYazisi
                .Fill.Visible = msoTrue
                .Fill.Solid
                .Fill.ForeColor.RGB = durumRengi
                .Fill.Transparency = 0
                .line.Visible = msoFalse
            End With
        End If
        If AKSORA_PANEL_KART_CANLI(CStr(kodlar(i))) Then AKSORA_PANEL_KART_GUNCELLENIYOR ws, CStr(kodlar(i))
        On Error GoTo Cikis
    Next i
    If enKotu = 0 Then
        durumRengi = RGB(224, 242, 254)
        yaziRengi = RGB(12, 45, 79)
    ElseIf enKotu = 1 Then
        durumRengi = RGB(255, 246, 219)
        yaziRengi = RGB(146, 64, 14)
    Else
        durumRengi = RGB(254, 226, 226)
        yaziRengi = RGB(153, 27, 27)
    End If
    If Not AKSORA_PANEL_CANLI Then
        With ws.Range("B6:AB6")
            .Interior.Color = durumRengi
            .Font.Color = yaziRengi
            .Font.bold = True
            .Font.Size = 10
        End With
        ws.Range("B6").HorizontalAlignment = xlLeft
        ws.Range("B6").IndentLevel = 1
    End If
    Set buton = Nothing
    On Error Resume Next
    Set buton = ws.Shapes(p & "IL_DEGISTIR")
    If Not buton Is Nothing Then
        With ws.Range("B6:AB6")
            AKSORA_IL_DEGISTIR_STIL buton, .Left, .Top, .Width, .Height
        End With
    End If
    V22_KART_SON_SATIR ws

    ws.Rows("31:" & ws.Rows.Count).Hidden = True
    ws.Columns("AC:XFD").Hidden = True
    ws.ScrollArea = "A1:AB30"
    On Error GoTo 0
Cikis:
End Sub

'==========================================================
' ORTAK PANEL OGELERI (Malzeme + Musteri ayni gorunum)
'==========================================================
' Il degistir: seridin sagina oturan, yuvarlak, gradyanli hap dugme.
Public Sub AKSORA_IL_DEGISTIR_STIL(ByVal dg As Shape, ByVal sol As Double, ByVal ust As Double, _
    ByVal genislik As Double, ByVal yukseklik As Double)
    Dim h As Double
    On Error Resume Next
    dg.AutoShapeType = msoShapeRoundedRectangle
    dg.Adjustments.Item(1) = 0.5
    h = yukseklik - 8
    If h < 16 Then h = yukseklik
    dg.Height = h: dg.Width = 196
    dg.Top = ust + (yukseklik - h) / 2
    dg.Left = sol + genislik - dg.Width - 10
    dg.Fill.Visible = msoTrue
    dg.Fill.TwoColorGradient msoGradientVertical, 1
    dg.Fill.ForeColor.RGB = RGB(12, 45, 79)
    dg.Fill.BackColor.RGB = RGB(0, 150, 136)
    dg.line.Visible = msoTrue: dg.line.ForeColor.RGB = RGB(94, 234, 212): dg.line.Weight = 1
    dg.Glow.Color.RGB = RGB(45, 212, 191): dg.Glow.Radius = 5: dg.Glow.Transparency = 0.55
    dg.Shadow.Visible = msoFalse
    With dg.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
        .TextRange.Text = ChrW(8644) & "   ÝL DEÐÝÞTÝR"
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 10
        .TextRange.Font.bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.Font.Spacing = 0.8
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    dg.ZOrder msoBringToFront
End Sub

' Kartlarin son satiri: DEPO SAYISI (stok sayfasindaki farkli depo adedi).
Private Sub V22_KART_SON_SATIR(ByVal ws As Worksheet)
    Dim kartlar As Variant, stoklar As Variant, i As Long, k As String, adet As Long
    Dim e3 As Shape, e4 As Shape, d3 As Shape, d4 As Shape, b As Shape
    kartlar = Array("B", "G", "L", "Q")
    stoklar = Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
    On Error Resume Next
    For i = 0 To 3
        k = CStr(kartlar(i))
        Set e3 = Nothing: Set e4 = Nothing: Set d3 = Nothing: Set d4 = Nothing
        Set e3 = ws.Shapes(p & "KART_ETIKET_" & k & "3")
        Set e4 = ws.Shapes(p & "KART_ETIKET_" & k & "4")
        Set d3 = ws.Shapes(p & "KART_DEGER_" & k & "3")
        Set d4 = ws.Shapes(p & "KART_DEGER_" & k & "4")
        If Not e3 Is Nothing And Not e4 Is Nothing And Not d3 Is Nothing Then
            e3.TextFrame2.TextRange.Text = "STOK DEÐERÝ"
            e4.Left = e3.Left: e4.Width = e3.Width
            With e4.TextFrame2.TextRange
                .Text = "DEPO SAYISI"
                .ParagraphFormat.Alignment = e3.TextFrame2.TextRange.ParagraphFormat.Alignment
                .Font.Size = e3.TextFrame2.TextRange.Font.Size
                .Font.bold = e3.TextFrame2.TextRange.Font.bold
                .Font.Fill.ForeColor.RGB = e3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
            End With
            If d4 Is Nothing Then
                Set d4 = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, d3.Left, e4.Top, d3.Width, e4.Height)
                d4.Name = p & "KART_DEGER_" & k & "4"
                d4.Fill.Visible = msoFalse: d4.line.Visible = msoFalse
                d4.Placement = xlFreeFloating
                With d4.TextFrame2
                    .MarginLeft = d3.TextFrame2.MarginLeft: .MarginRight = d3.TextFrame2.MarginRight
                    .MarginTop = 0: .MarginBottom = 0
                    .VerticalAnchor = msoAnchorMiddle
                    .TextRange.Text = "-"
                    .TextRange.ParagraphFormat.Alignment = msoAlignRight
                    .TextRange.Font.Name = "Aptos"
                    .TextRange.Font.Size = d3.TextFrame2.TextRange.Font.Size
                    .TextRange.Font.bold = msoTrue
                    .TextRange.Font.Fill.ForeColor.RGB = d3.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
                End With
            End If
            d4.DrawingObject.Formula = ""
            d4.Left = d3.Left: d4.Width = d3.Width
            d4.Top = e4.Top: d4.Height = e4.Height
            adet = V22_DEPO_ADET(CStr(stoklar(i)))
            d4.TextFrame2.TextRange.Text = IIf(adet > 0, AKSORA_SAYI_METNI(adet, "#,##0"), "-")
            d4.Visible = msoTrue
            d4.ZOrder msoBringToFront
        End If
    Next i
    Set b = Nothing: Set b = ws.Shapes(p & "KART_BASLIK_L")
    If Not b Is Nothing Then b.TextFrame2.TextRange.Text = "ÝLÇELER"
    Set b = Nothing: Set b = ws.Shapes(p & "KART_BASLIK_Q")
    If Not b Is Nothing Then b.TextFrame2.TextRange.Text = "YÜKLENÝCÝ"
    On Error GoTo 0
End Sub

Private Function V22_DEPO_ADET(ByVal ad As String) As Long
    Dim sh As Worksheet, son As Long, v As Variant, r As Long, d As Object, t As String
    On Error GoTo Cikis
    Set sh = ThisWorkbook.Worksheets(ad)
    son = sh.Cells(sh.Rows.Count, "G").End(xlUp).Row
    If son < 7 Then Exit Function
    v = sh.Range("B7:B" & son).Value2
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    If IsArray(v) Then
        For r = 1 To UBound(v, 1)
            If Not IsError(v(r, 1)) Then
                t = Trim$(CStr(v(r, 1)))
                If Len(t) > 0 Then d(t) = True
            End If
        Next r
    Else
        If Len(Trim$(CStr(v))) > 0 Then d(Trim$(CStr(v))) = True
    End If
    V22_DEPO_ADET = d.Count
Cikis:
End Function

' Guncelleme basladiginda: secilen kartlarin isigi mavi, durum "Guncelleniyor".
Public Sub AKSORA_PANEL_GUNCELLENIYOR(ByVal gruplar As String)
    Dim ws As Worksheet, g As Variant, k As String, s As Shape, sc As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ANA PANEL")
    If ws Is Nothing Then Exit Sub
    For Each g In Split(gruplar, ",")
        Select Case CStr(g)
            Case "CDP": k = "B"
            Case "EPP": k = "G"
            Case "ILCE": k = "L"
            Case Else: k = "Q"
        End Select
        AKSORA_PANEL_KART_GUNCELLENIYOR ws, CStr(g)
    Next g
    AKSORA_PANEL_GUNCELLENEN = UCase$(gruplar)
    AKSORA_PANEL_CANLI = True
    AKSORA_PANEL_SERIT "Güncelleniyor  •  SAP baðlantýsý hazýrlanýyor...", True
    On Error GoTo 0
End Sub

' Guncelleme surerken karttaki durum (musteri paneliyle ayni): mavi LED, "Güncelleniyor...".
Public Sub AKSORA_PANEL_KART_GUNCELLENIYOR(ByVal ws As Worksheet, ByVal kod As String)
    Dim k As String, s As Shape
    Select Case UCase$(kod)
        Case "CDP": k = "B"
        Case "EPP": k = "G"
        Case "ILCE": k = "L"
        Case Else: k = "Q"
    End Select
    On Error Resume Next
    With ws.Range(UCase$(kod) & "_DURUM").Cells(1, 1)
        .Value = "Güncelleniyor...": .Font.Color = RGB(29, 78, 216): .Font.bold = True
    End With
    Set s = ws.Shapes(p & "KART_DEGER_" & k & "0")
    If Not s Is Nothing Then
        s.TextFrame2.TextRange.Text = "Güncelleniyor..."
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(29, 78, 216)
    End If
    Set s = Nothing: Set s = ws.Shapes(p & "KART_LED_" & k)
    If Not s Is Nothing Then
        s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = RGB(59, 130, 246): s.Fill.Transparency = 0
        s.Glow.Color.RGB = RGB(59, 130, 246): s.Glow.Radius = 4: s.Glow.Transparency = 0.5
    End If
    On Error GoTo 0
End Sub

' Panel yenilemeleri guncellenen kartin durumunu ezmesin.
Public Function AKSORA_PANEL_KART_CANLI(ByVal kod As String) As Boolean
    If Not AKSORA_PANEL_CANLI Then Exit Function
    AKSORA_PANEL_KART_CANLI = (InStr("," & AKSORA_PANEL_GUNCELLENEN & ",", "," & UCase$(kod) & ",") > 0)
End Function

' Durum seridine anlik mesaj yazar (panel aciksa ekranda hemen gorunur).
Public Sub AKSORA_PANEL_SERIT(ByVal metin As String, Optional ByVal calisiyor As Boolean = False)
    Dim ws As Worksheet, sc As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ANA PANEL")
    If ws Is Nothing Then Exit Sub
    ws.Range("B6").Value = metin
    If calisiyor Or AKSORA_PANEL_CANLI Then
        ws.Range("B6:AB6").Interior.Color = RGB(255, 246, 219)
        ws.Range("B6:AB6").Font.Color = RGB(146, 64, 14)
        ws.Range("B6:AB6").Font.bold = True
        ws.Range("B6").HorizontalAlignment = xlLeft
        ws.Range("B6").IndentLevel = 1
    End If
    If ActiveSheet Is ws Then
        sc = application.ScreenUpdating
        application.ScreenUpdating = True
        DoEvents
        application.ScreenUpdating = sc
    End If
    On Error GoTo 0
End Sub

' Ana panel (malzeme ve musteri): orta siradaki uc buyuk dugme guncelleme dugmelerinin
' (dolayisiyla kartlarin) ayni sol/sag cizgisine oturur. Dort sutunda TUM STOK iki sutun kaplar.
Public Sub AKSORA_PANEL_ORTA_HIZALA(ByVal ws As Worksheet)
    Dim sap As Shape, tum As Shape, ara As Shape, c1 As Shape, c2 As Shape, c3 As Shape, c4 As Shape
    On Error Resume Next
    Set sap = ws.Shapes("btnBaglantiKontrol"): Set tum = ws.Shapes("btnTumStok")
    Set ara = ws.Shapes("btnMalzemeAraPanel")
    Set c1 = ws.Shapes("btnCDP"): Set c2 = ws.Shapes("btnEPP")
    Set c3 = ws.Shapes("btnIlce"): Set c4 = ws.Shapes("btnYuk")
    If sap Is Nothing Or tum Is Nothing Or ara Is Nothing Or c1 Is Nothing Or c3 Is Nothing Or c4 Is Nothing Then Exit Sub
    sap.Left = c1.Left: sap.Width = c1.Width
    ara.Left = c4.Left: ara.Width = c4.Width
    If c2 Is Nothing Then
        tum.Left = c3.Left: tum.Width = c3.Width
    Else
        tum.Left = c2.Left: tum.Width = c3.Left + c3.Width - c2.Left
    End If
    On Error GoTo 0
End Sub

Public Sub AKSORA_PANEL_SURE_KAYDET(ByVal bas As Double)
    Dim sn As Double
    sn = Timer - bas
    If sn < 0 Then sn = sn + 86400#
    On Error Resume Next
    ThisWorkbook.names.Add Name:="_AKSORA_MALZEME_SON_SURE", RefersTo:="=" & Trim$(Str$(Round(sn, 2))), Visible:=False
    On Error GoTo 0
End Sub

' Sure metni: 60 sn altinda "39,40 sn", ustunde "3 dk 42 sn".
Public Function AKSORA_SURE_METNI(ByVal sn As Double) As String
    If sn < 60 Then
        AKSORA_SURE_METNI = AKSORA_SAYI_METNI(sn, "0.00") & " sn"
    Else
        AKSORA_SURE_METNI = CStr(Int(sn / 60)) & " dk " & Format$(Int(sn) Mod 60, "00") & " sn"
    End If
End Function

' Bütün iller (toplu) guncellemesinin son suresi: malzeme veya musteri.
Public Function AKSORA_TOPLU_SON_SURE(ByVal ad As String) As Double
    Dim t As String
    On Error Resume Next
    t = ThisWorkbook.names(ad).RefersTo
    On Error GoTo 0
    If Len(t) > 1 Then AKSORA_TOPLU_SON_SURE = Val(Mid$(t, 2))
End Function

Public Function AKSORA_PANEL_SON_SURE() As Double
    Dim t As String
    On Error Resume Next
    t = ThisWorkbook.names("_AKSORA_MALZEME_SON_SURE").RefersTo
    On Error GoTo 0
    If Len(t) > 1 Then AKSORA_PANEL_SON_SURE = Val(Mid$(t, 2))
End Function

' Guncelleme bitince: panel yenilenir, ana panel tek ve hazir halde ekrana gelir.
Public Sub AKSORA_ANA_PANEL_GOSTER()
    Dim ws As Worksheet, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("ANA PANEL")
    If ws Is Nothing Then Exit Sub
    ev = application.EnableEvents
    application.ScreenUpdating = False
    PANEL_YENILE
    ' Yonetim merkezinden tum birimler guncellenirken panel yalniz yenilenir, ekrana getirilmez.
    If AKSORA_ADMIN_TOPLU Then Exit Sub
    application.EnableEvents = False
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    application.EnableEvents = ev
    AKSORA_V13_GORUNUM ws
    application.ScreenUpdating = True
    DoEvents
    On Error GoTo 0
End Sub
