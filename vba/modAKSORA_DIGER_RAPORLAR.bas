Attribute VB_Name = "modAKSORA_DIGER_RAPORLAR"
Option Explicit

' AKSORA - BES RAPOR / TEK MODUL
' Bu kodu eski BES RAPOR tasarim modulunun yerine yapistirin.
' CDP_RAPOR KESINLIKLE DEGISTIRILMEZ.
' Calistir: AKSORA_DIGER_DORT_RAPOR_KUR
' Grafik, KPI, banner, filigran ve alt bilgi; mevcut SAP rapor verileri korunur.
' Gorseller calisma kitabi ile ayni klasorde olmalidir.

Private Const AK_BANNER As String = "AKSORA_RAPOR_BANNER.png"
Private Const AK_LOGO As String = "AKSORA_RAPOR_FILIGRAN.png"
Private Const AK_ON As String = "AKSORA_RP_"
Private Const AK_ESKI As String = "AKSORA_CDP_V1_"

Public Sub AKSORA_DIGER_DORT_RAPOR_KUR()
    Dim adlar As Variant, i As Long, ws As Worksheet, bilgi As String
    Dim n As Long, scr As Boolean, evt As Boolean, alerts As Boolean
    Dim ilkSayfa As Worksheet
    adlar = Array("EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR")
    Set ilkSayfa = ActiveSheet
    scr = application.ScreenUpdating
    evt = application.EnableEvents
    alerts = application.DisplayAlerts
    On Error GoTo SON_HATA
    application.ScreenUpdating = False
    application.EnableEvents = False
    application.DisplayAlerts = False
    For i = LBound(adlar) To UBound(adlar)
        Set ws = AK_Sayfa(CStr(adlar(i)))
        If ws Is Nothing Then
            bilgi = bilgi & vbCrLf & CStr(adlar(i)) & ": sayfa bulunamadi"
        Else
            On Error Resume Next
            Err.Clear
            AK_RaporKur ws
            If Err.Number <> 0 Then
                bilgi = bilgi & vbCrLf & ws.Name & ": " & Err.Description
                Err.Clear
            Else
                n = n + 1
            End If
            On Error GoTo SON_HATA
        End If
    Next i
    ' Diger sayfalardaki fotograflarin konumu da sabitlenir.
    ' Calisan SAP/panel makrolarini bozmamak icin diger sayfalara koruma uygulanmaz.
    ' Diger sayfalara dokunulmaz.
    ' AK_DigerResimleriSabitle
    If Not ilkSayfa Is Nothing Then ilkSayfa.Activate
    application.DisplayAlerts = alerts
    application.EnableEvents = evt
    application.ScreenUpdating = scr
    MsgBox n & " rapor hazir." & bilgi, vbInformation, "AKSORA"
    Exit Sub
SON_HATA:
    application.DisplayAlerts = alerts
    application.EnableEvents = evt
    application.ScreenUpdating = scr
    MsgBox "Islem durdu: " & Err.Description & bilgi, vbExclamation, "AKSORA"
End Sub

Private Function AK_Norm(ByVal s As String) As String
    s = Replace(s, ChrW(304), "I")
    s = Replace(s, ChrW(305), "I")
    s = Replace(s, "i", "I")
    s = Replace(s, "ü", "U"): s = Replace(s, "Ü", "U")
    s = Replace(s, "þ", "S"): s = Replace(s, "Þ", "S")
    s = Replace(s, "ð", "G"): s = Replace(s, "Ð", "G")
    s = Replace(s, "ö", "O"): s = Replace(s, "Ö", "O")
    s = Replace(s, "ç", "C"): s = Replace(s, "Ç", "C")
    AK_Norm = UCase$(Trim$(s))
End Function

Private Function AK_Sayfa(ByVal ad As String) As Worksheet
    Dim w As Worksheet
    For Each w In ThisWorkbook.Worksheets
        If AK_Norm(w.Name) = AK_Norm(ad) Then
            Set AK_Sayfa = w
            Exit Function
        End If
    Next w
End Function

Private Function AK_Baslik(ByVal ad As String) As String
    Select Case AK_Norm(ad)
        Case "CDP_RAPOR": AK_Baslik = "CDP STOK RAPORU"
        Case "EPP_RAPOR": AK_Baslik = "EPP STOK RAPORU"
        Case "RAPORLAR": AK_Baslik = "GENEL STOK RAPORU"
        Case "ILCE_RAPORLAR": AK_Baslik = "ILCE STOK RAPORU"
        Case "YUKLENICI_RAPOR": AK_Baslik = "YUKLENICI STOK RAPORU"
        Case Else: AK_Baslik = "STOK RAPORU"
    End Select
End Function

Private Sub AK_RaporKur(ByVal ws As Worksheet)
    Dim son As Long, toplam As Long, degerKol As Long, adetKol As Long
    Dim sonKol As Long, r As Long, c As Long, s As String
    Dim eskKor As Boolean
    If ws.ProtectContents Then Err.Raise vbObjectError + 101, , "Hucre korumasi acik; atlandi."
    ' Sadece bizim onceki uyguladigimiz, hucresiz resim korumasini kaldir.
    eskKor = ws.ProtectDrawingObjects
    If eskKor Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0
        If ws.ProtectDrawingObjects Then Err.Raise vbObjectError + 102, , "Sekil korumasi sifreli; atlandi."
    End If

    ' Rapor tablosu 8. satirda baslar. Basliklara gore kolon tespiti.
    sonKol = ws.Cells(8, ws.Columns.Count).End(xlToLeft).Column
    If sonKol < 4 Or sonKol > 5 Then Err.Raise vbObjectError + 103, , "8. satirda 4-5 rapor sutunu bulunamadi; atlandi."
    For c = 1 To sonKol
        s = AK_Norm(CStr(ws.Cells(8, c).Value2))
        If InStr(s, "DEGER") > 0 Or InStr(s, "TUTAR") > 0 Then degerKol = c
        If InStr(s, "SATIR") > 0 Or InStr(s, "ADET") > 0 Then adetKol = c
    Next c
    ' Eski ilce/yuklenici raporlarinda basliklar farkli olabilir.
    ' Son iki sayisal sutunu kullanmadan once veri tiplerini kontrol et.
    If degerKol = 0 Then
        If sonKol >= 4 Then
            If IsNumeric(ws.Cells(9, sonKol).Value2) Then degerKol = sonKol
        End If
    End If
    If adetKol = 0 Then
        If degerKol > 2 Then
            If IsNumeric(ws.Cells(9, degerKol - 1).Value2) Then adetKol = degerKol - 1
        End If
    End If
    If degerKol = 0 Or adetKol = 0 Then Err.Raise vbObjectError + 104, , "Stok Satiri / Stok Degeri sutunlari tespit edilemedi; atlandi."
    For r = 9 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        s = AK_Norm(CStr(ws.Cells(r, 1).Value2))
        If s = "TOPLAM" Or s = "GENEL TOPLAM" Then
            toplam = r
            Exit For
        End If
    Next r
    If toplam < 10 Then Err.Raise vbObjectError + 105, , "TOPLAM satiri bulunamadi; atlandi."
    son = toplam - 1

    AK_EskiNesneler ws
    AK_Tablo ws, son, toplam, sonKol, adetKol, degerKol
    AK_UstBanner ws
    AK_Kart ws, "A3:E6", "SON GUNCELLEME", AK_Tarih(ws), "01", RGB(36, 111, 195)
    AK_Kart ws, "F3:J6", "TOPLAM STOK SATIRI", AK_GercekStokSatiri(ws), "02", RGB(29, 151, 91)
    AK_Kart ws, "K3:Q6", "TOPLAM STOK DEGERI", AK_Para(ws, toplam, degerKol), "03", RGB(36, 111, 195)
    AK_Grafik ws, son, degerKol, toplam
    AK_Filigran ws, toplam
    AK_AltYazi ws, toplam
    AK_AnaButonu ws

    ' Raporlarda yalnizca nesneler korunur; hucreler ve SAP yenileme serbesttir.
    ' Butonun makrosu ve filtreleme korunur.
    AK_RaporResimleriniKoru ws
    ws.Activate
    If AKSORA_PENCERE_KUCUK(ActiveWindow) Then Exit Sub
    On Error Resume Next
    With ActiveWindow
        .FreezePanes = False
        .SplitRow = 0
        .SplitColumn = 0
        .DisplayGridlines = False
    End With
End Sub

Private Sub AK_EskiNesneler(ByVal ws As Worksheet)
    Dim i As Long, ad As String
    For i = ws.ChartObjects.Count To 1 Step -1
        ad = ws.ChartObjects(i).Name
        If Left$(ad, Len(AK_ON)) = AK_ON Or Left$(ad, Len(AK_ESKI)) = AK_ESKI Then ws.ChartObjects(i).Delete
    Next i
    For i = ws.Shapes.Count To 1 Step -1
        ad = ws.Shapes(i).Name
        If Left$(ad, Len(AK_ON)) = AK_ON Or Left$(ad, Len(AK_ESKI)) = AK_ESKI Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub AK_Tablo(ByVal ws As Worksheet, ByVal son As Long, ByVal toplam As Long, _
                     ByVal sonKol As Long, ByVal adetKol As Long, ByVal degerKol As Long)
    Dim r As Long
    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 29
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 24
    If sonKol = 5 Then ws.Columns("E").ColumnWidth = 18 Else ws.Columns("E").ColumnWidth = 3
    ws.Columns("F:Q").ColumnWidth = 11
    ws.Columns("R").ColumnWidth = 3
    ws.Columns("S").ColumnWidth = 17
    ws.Rows("1:2").RowHeight = 32
    ws.Rows("3:6").RowHeight = 25
    ws.Rows(7).RowHeight = 13
    ws.Rows(8).RowHeight = 30
    ws.Rows("9:" & toplam).RowHeight = 24
    With ws.Range(ws.Cells(8, 1), ws.Cells(8, sonKol))
        .Interior.Color = RGB(24, 58, 94)
        .Font.Name = "Aptos": .Font.Size = 11
        .Font.bold = True: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
    End With
    With ws.Range(ws.Cells(9, 1), ws.Cells(son, sonKol))
        .Font.Name = "Aptos": .Font.Size = 11
        .Font.bold = False: .Font.Color = RGB(28, 49, 76)
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
    End With
    For r = 9 To son
        If r Mod 2 = 0 Then
            ws.Range(ws.Cells(r, 1), ws.Cells(r, sonKol)).Interior.Color = RGB(238, 247, 255)
        Else
            ws.Range(ws.Cells(r, 1), ws.Cells(r, sonKol)).Interior.Color = vbWhite
        End If
    Next r
    ws.Range(ws.Cells(9, 1), ws.Cells(son, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(9, adetKol), ws.Cells(toplam, adetKol)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(9, degerKol), ws.Cells(toplam, degerKol)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(9, degerKol), ws.Cells(toplam, degerKol)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(9, adetKol), ws.Cells(toplam, adetKol)).NumberFormat = "#,##0"
    With ws.Range(ws.Cells(toplam, 1), ws.Cells(toplam, sonKol))
        .Interior.Color = RGB(24, 58, 94)
        .Font.Name = "Aptos": .Font.Size = 12
        .Font.bold = True: .Font.Color = vbWhite
        .VerticalAlignment = xlCenter
        .RowHeight = 30
        .Borders.LineStyle = xlNone
    End With
    ' Toplam degerler ayni dikey hizaya gelir; bosluk icin sutun genisligi.
    ws.Range(ws.Cells(toplam, adetKol), ws.Cells(toplam, degerKol)).VerticalAlignment = xlCenter
    ws.Range("F3:Q" & application.Max(25, toplam + 4)).Interior.Color = vbWhite
End Sub

Private Sub AK_UstBanner(ByVal ws As Worksheet)
    Dim alan As Range, sh As Shape, yol As String
    Set alan = ws.Range("A1:Q2")
    yol = AK_Yol(AK_BANNER)
    If Len(yol) > 0 Then
        Set sh = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, alan.Left, alan.Top, alan.Width, alan.Height)
        sh.Name = AK_ON & "BANNER"
        AK_Kilitle sh
    Else
        alan.Interior.Color = RGB(24, 58, 94)
    End If
    Set sh = AK_Metin(ws, AK_ON & "BASLIK", "AKSORA  |  " & AK_Baslik(ws.Name), _
                      alan.Left + 20, alan.Top + 12, ws.Range("M1").Left - alan.Left - 20, _
                      45, 24, vbWhite, True)
    sh.TextFrame2.VerticalAnchor = msoAnchorMiddle
End Sub

Private Sub AK_Kart(ByVal ws As Worksheet, ByVal adres As String, ByVal baslik As String, _
                    ByVal deger As String, ByVal no As String, ByVal renk As Long)
    Dim a As Range, sh As Shape, nm As String
    Set a = ws.Range(adres)
    nm = Replace(adres, ":", "_")
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, a.Left + 3, a.Top + 5, a.Width - 6, a.Height - 10)
    With sh
        .Name = AK_ON & "KART_" & nm
        .Fill.ForeColor.RGB = RGB(243, 249, 255)
        .line.ForeColor.RGB = RGB(210, 227, 243)
        .line.Weight = 1
    End With
    AK_Kilitle sh
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, a.Left + 12, a.Top + 22, 46, 46)
    With sh
        .Name = AK_ON & "NO_" & nm
        .Fill.ForeColor.RGB = renk
        .line.Visible = msoFalse
        .TextFrame2.TextRange.Text = no
        .TextFrame2.TextRange.Font.Name = "Aptos Display"
        .TextFrame2.TextRange.Font.Size = 17
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    AK_Kilitle sh
    Set sh = AK_Metin(ws, AK_ON & "ETIKET_" & nm, baslik, _
                      a.Left + 70, a.Top + 18, a.Width - 84, 25, 12, RGB(28, 54, 83), True)
    Set sh = AK_Metin(ws, AK_ON & "DEGER_" & nm, deger, _
                      a.Left + 70, a.Top + 47, a.Width - 84, 38, 17, RGB(22, 58, 96), True)
    sh.ZOrder msoBringToFront
End Sub

Private Function AK_Metin(ByVal ws As Worksheet, ByVal ad As String, ByVal yazi As String, _
                          ByVal x As Double, ByVal y As Double, ByVal w As Double, _
                          ByVal h As Double, ByVal punto As Double, ByVal renk As Long, _
                          ByVal kalin As Boolean) As Shape
    Dim sh As Shape
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    With sh
        .Name = ad
        .Fill.Visible = msoFalse
        .line.Visible = msoFalse
        .TextFrame2.MarginLeft = 0: .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.TextRange.Text = yazi
        .TextFrame2.TextRange.Font.Name = "Aptos Display"
        .TextFrame2.TextRange.Font.Size = punto
        .TextFrame2.TextRange.Font.bold = IIf(kalin, msoTrue, msoFalse)
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = renk
    End With
    AK_Kilitle sh
    Set AK_Metin = sh
End Function

Private Function AK_Tarih(ByVal ws As Worksheet) As String
    Dim v As Variant
    v = ws.Range("B3").Value2
    If IsDate(ws.Range("B3").Value) Then
        AK_Tarih = Format$(ws.Range("B3").Value, "dd.mm.yyyy hh:mm:ss")
    Else
        AK_Tarih = CStr(ws.Range("B3").Text)
    End If
End Function

Private Function AK_Para(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As String
    If IsNumeric(ws.Cells(r, c).Value2) Then
        AK_Para = AKSORA_SAYI_METNI(CDbl(ws.Cells(r, c).Value2), "#,##0.00") & " TL"
    Else
        AK_Para = ws.Cells(r, c).Text
    End If
End Function

Private Sub AK_Grafik(ByVal ws As Worksheet, ByVal son As Long, _
                      ByVal degerKol As Long, ByVal toplam As Long)
    Dim alan As Range, co As ChartObject, CH As Chart, sr As Series
    Dim adlar() As String, tutarlar() As Double, renkler As Variant
    Dim r As Long, n As Long, i As Long, deger As Double, genel As Double
    Dim x As Double, y As Double, panelLeft As Double, panelWidth As Double
    Dim satirY As Double, satirAdimi As Double, PW As Double, ph As Double
    Dim cx As Double, cy As Double, sh As Shape

    ReDim adlar(1 To son - 8)
    ReDim tutarlar(1 To son - 8)
    For r = 9 To son
        If IsNumeric(ws.Cells(r, degerKol).Value2) Then
            deger = CDbl(ws.Cells(r, degerKol).Value2)
            If deger > 0 Then
                n = n + 1
                adlar(n) = CStr(ws.Cells(r, 2).Value2)
                If Len(Trim$(adlar(n))) = 0 Then adlar(n) = CStr(ws.Cells(r, 1).Value2)
                tutarlar(n) = deger
                genel = genel + deger
            End If
        End If
    Next r
    If n = 0 Or genel <= 0 Then Exit Sub
    ReDim Preserve adlar(1 To n)
    ReDim Preserve tutarlar(1 To n)
    renkler = Array(RGB(34, 77, 153), RGB(63, 151, 226), _
                   RGB(34, 166, 146), RGB(243, 166, 65), _
                   RGB(117, 91, 187), RGB(227, 104, 104), _
                   RGB(105, 140, 168), RGB(81, 175, 113), _
                   RGB(224, 143, 84), RGB(145, 121, 172), _
                   RGB(98, 123, 163), RGB(170, 191, 207))

    If AKSORA_TASARIM_AKTIF() Then
        ' Uzun depo listelerinde her satira en az 20 nokta yer ayir.
        If AK_Norm(ws.Name) = "ILCE_RAPORLAR" Or AK_Norm(ws.Name) = "YUKLENICI_RAPOR" Then
            ' Bes sutunlu tablo ile grafik arasinda F sutunu bosluk olarak kalir.
            Set alan = ws.Range("G8:T" & application.Max(23, n + 11))
        Else
            Set alan = ws.Range("F8:T" & application.Max(23, n + 11))
        End If
    Else
        Set alan = ws.Range("F8:Q23")
    End If
    x = alan.Left: y = alan.Top
    If AKSORA_TASARIM_AKTIF() Then
        panelWidth = alan.Width * 0.39
    Else
        panelWidth = alan.Width * 0.31
    End If
    panelLeft = x + alan.Width - panelWidth
    Set sh = ws.Shapes.AddShape(msoShapeRectangle, x, y, alan.Width, alan.Height)
    With sh
        .Name = AK_ON & "HALKA_ZEMIN"
        .Fill.ForeColor.RGB = RGB(245, 250, 255)
        .line.Visible = msoFalse
    End With
    AK_Kilitle sh
    Set sh = AK_HalkaYazi(ws, AK_ON & "HALKA_BASLIK", _
       "DEPOLARA GORE STOK DAGILIMI", x + 12, y + 7, _
       alan.Width - 24, 25, 13, RGB(26, 57, 88), True, msoAlignCenter)

    Set co = ws.ChartObjects.Add(x + 22, y + 42, _
       alan.Width - panelWidth - 34, alan.Height - 62)
    co.Name = AK_ON & "HALKA_GRAFIK"
    co.Placement = xlFreeFloating
    Set CH = co.Chart
    With CH
        .ChartType = xlDoughnut
        .HasLegend = False
        .HasTitle = False
        .ChartArea.Format.Fill.ForeColor.RGB = RGB(245, 250, 255)
        .ChartArea.Format.line.Visible = msoFalse
        .PlotArea.Format.Fill.ForeColor.RGB = RGB(245, 250, 255)
        .PlotArea.Format.line.Visible = msoFalse
        Set sr = .SeriesCollection.NewSeries
        sr.Name = "Stok dagilimi"
        sr.XValues = adlar
        sr.Values = tutarlar
        .ChartGroups(1).DoughnutHoleSize = 72
        .ChartGroups(1).FirstSliceAngle = 270
    End With
    For i = 1 To n
        With sr.Points(i)
            .Format.Fill.ForeColor.RGB = renkler((i - 1) Mod 12)
            .Format.line.ForeColor.RGB = vbWhite
            .Format.line.Weight = 1.6
        End With
    Next i
    CH.Refresh
    DoEvents
    PW = CH.PlotArea.InsideWidth: ph = CH.PlotArea.InsideHeight
    cx = co.Left + CH.PlotArea.InsideLeft + PW / 2
    cy = co.Top + CH.PlotArea.InsideTop + ph / 2
        PW = application.Min(PW, ph) * 0.72 * 0.9

    Set sh = AK_HalkaYazi(ws, AK_ON & "HALKA_MERKEZ", _
       "TOPLAM STOK" & vbCrLf & AKSORA_SAYI_METNI(genel, "#,##0.00") & " TL", _
       cx - PW / 2, cy - 20, PW, 40, _
              11, RGB(26, 57, 88), True, msoAlignCenter)

    With sh.TextFrame2
        .AutoSize = msoAutoSizeNone
        .WordWrap = msoFalse
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
    End With

    CH.HasTitle = False
    sh.ZOrder msoBringToFront

        Set sh = AK_HalkaYazi(ws, AK_ON & "HALKA_LISTE_BASLIK", _
        "DEPO", panelLeft + 22, y + 43, _
        panelWidth - 90, 20, 10, _
        RGB(26, 57, 88), True, msoAlignLeft)

    Set sh = AK_HalkaYazi(ws, AK_ON & "HALKA_PAY_BASLIK", _
        "PAY", panelLeft + panelWidth - 69, y + 43, _
        57, 20, 10, _
        RGB(26, 57, 88), True, msoAlignRight)
    satirAdimi = (alan.Height - 90) / n
    If AKSORA_TASARIM_AKTIF() And satirAdimi < 20 Then satirAdimi = 20
    If satirAdimi > 24 Then satirAdimi = 24
    For i = 1 To n
        satirY = y + 66 + (i - 1) * satirAdimi
        Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            panelLeft + 6, satirY + 4, 9, 9)
        With sh
            .Name = AK_ON & "HALKA_RENK_" & i
            .Fill.ForeColor.RGB = renkler((i - 1) Mod 12)
            .line.Visible = msoFalse
        End With
        AK_Kilitle sh
        Set sh = AK_HalkaYazi(ws, AK_ON & "HALKA_AD_" & i, _
            adlar(i), panelLeft + 22, satirY, _
            panelWidth - 90, 19, 9, RGB(35, 56, 78), False, msoAlignLeft)
        Set sh = AK_HalkaYazi(ws, AK_ON & "HALKA_PAY_" & i, _
            AKSORA_SAYI_METNI(tutarlar(i) / genel, "0.0%"), _
            panelLeft + panelWidth - 69, satirY, 57, 19, _
            10, RGB(26, 57, 88), True, msoAlignRight)
    Next i
End Sub

Private Function AK_HalkaYazi(ByVal ws As Worksheet, ByVal ad As String, _
        ByVal metin As String, ByVal sol As Double, ByVal ust As Double, _
        ByVal gen As Double, ByVal yuk As Double, ByVal punto As Single, _
        ByVal renk As Long, ByVal kalin As Boolean, _
        ByVal hiz As MsoParagraphAlignment) As Shape
    Dim sh As Shape
    Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, sol, ust, gen, yuk)
    With sh
        .Name = ad
        .Fill.Visible = msoFalse: .line.Visible = msoFalse
        .TextFrame2.MarginLeft = 1: .TextFrame2.MarginRight = 1
        .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.TextRange.Text = metin
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = punto
        .TextFrame2.TextRange.Font.bold = kalin
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = renk
        .TextFrame2.TextRange.ParagraphFormat.Alignment = hiz
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
    End With
    AK_Kilitle sh
    Set AK_HalkaYazi = sh
End Function

Private Sub AK_Filigran(ByVal ws As Worksheet, ByVal toplam As Long)
    Dim alan As Range, sh As Shape, yol As String, r As Long
    yol = AK_Yol(AK_LOGO)
    If Len(yol) = 0 Then Exit Sub
    r = application.Max(24, toplam + 4)
    Set alan = ws.Range("J" & r & ":Q" & (r + 11))
    Set sh = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, alan.Left, alan.Top, -1, -1)
    With sh
        .Name = AK_ON & "FILIGRAN"
        .LockAspectRatio = msoTrue
        If .Width > alan.Width Then .Width = alan.Width
        If .Height > alan.Height Then .Height = alan.Height
        .Left = alan.Left + (alan.Width - .Width) / 2
        .Top = alan.Top + (alan.Height - .Height) / 2
    End With
    AK_Kilitle sh
End Sub

Private Sub AK_AltYazi(ByVal ws As Worksheet, ByVal toplam As Long)
    Dim r As Long, sh As Shape
    r = toplam + 1
    ws.Rows(r).RowHeight = 23
    ws.Range(ws.Cells(r, 1), ws.Cells(r, 5)).Borders.LineStyle = xlNone
    Set sh = AK_Metin(ws, AK_ON & "ALT", _
        "AKSORA  |  " & AK_Baslik(ws.Name) & "  |  Veriler SAP raporundan alinir.", _
        ws.Range("A" & r).Left + 3, ws.Range("A" & r).Top + 3, _
        ws.Range("H" & r).Left - ws.Range("A" & r).Left, 18, _
        10, RGB(99, 124, 153), False)
End Sub

Private Sub AK_AnaButonu(ByVal ws As Worksheet)
    Dim sh As Shape, yazi As String, hedef As Range, bulundu As Boolean
    Set hedef = ws.Range("S2")
    For Each sh In ws.Shapes
        If Left$(sh.Name, Len(AK_ON)) <> AK_ON Then
            yazi = ""
            On Error Resume Next
            yazi = AK_Norm(sh.TextFrame2.TextRange.Text)
            On Error GoTo 0
            If yazi = "ANA" Or yazi = "ANA PANEL" Then
                sh.Left = hedef.Left + 2
                sh.Top = hedef.Top + 2
                sh.Width = hedef.Width - 4
                sh.Height = hedef.Height - 4
                sh.Placement = xlFreeFloating
                sh.ZOrder msoBringToFront
                ' Var olan OnAction aynen korunur.
                bulundu = True
                Exit For
            End If
        End If
    Next sh
    ' Bulunmazsa islevi bilinmeyen sahte buton eklenmez.
End Sub

Private Function AK_Yol(ByVal ad As String) As String
    Dim p As String
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    p = ThisWorkbook.Path & application.PathSeparator & ad
    If Len(Dir$(p)) > 0 Then AK_Yol = p
End Function

Private Sub AK_Kilitle(ByVal sh As Shape)
    sh.Locked = True
    sh.Placement = xlFreeFloating
    On Error Resume Next
    sh.OnAction = vbNullString
    On Error GoTo 0
End Sub

Private Sub AK_RaporResimleriniKoru(ByVal ws As Worksheet)
    Dim sh As Shape
    For Each sh In ws.Shapes
        If sh.Type = msoPicture Or sh.Type = msoLinkedPicture Then
            sh.Locked = True
            sh.Placement = xlFreeFloating
        End If
    Next sh
    ' Sekillerin kilidi ancak sayfa korumasiyla etkinlesir.
    ' Hucre korumasi kapali: SAP yenileme ve filtreleme devam edebilir.
    ws.Protect DrawingObjects:=True, Contents:=False, Scenarios:=False, _
               UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Private Sub AK_DigerResimleriSabitle()
    Dim ws As Worksheet, sh As Shape
    For Each ws In ThisWorkbook.Worksheets
        Select Case AK_Norm(ws.Name)
            Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR"
                ' Bunlar AK_RaporResimleriniKoru icinde korundu.
            Case Else
                For Each sh In ws.Shapes
                    If sh.Type = msoPicture Or sh.Type = msoLinkedPicture Then
                        sh.Locked = True
                        sh.Placement = xlFreeFloating
                    End If
                Next sh
        End Select
    Next ws
End Sub



' Mevcut modAKSORA_DIGER_RAPORLAR modulunun EN ALTINA ekleyin.
' Calistirilacak makro: AKSORA_DORT_RAPOR_INCE_DUZELTME
' Mevcut rapor tasarim kodunu silmeyin. CDP_RAPOR'a dokunmaz.


Public Sub AKSORA_DORT_RAPOR_INCE_DUZELTME()
    Dim adlar As Variant, altSatirlar As Variant
    Dim i As Long, ws As Worksheet, logo As Shape
    Dim hedef As Range, ekran As Boolean, olay As Boolean
    Dim sonuc As String

    adlar = Array("EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR")
    altSatirlar = Array(31, 31, 34, 34)
    ekran = application.ScreenUpdating
    olay = application.EnableEvents
    On Error GoTo Hata
    application.ScreenUpdating = False
    application.EnableEvents = False

    For i = LBound(adlar) To UBound(adlar)
        Set ws = AKSORA_InceSayfa(CStr(adlar(i)))
        If ws Is Nothing Then
            sonuc = sonuc & vbCrLf & CStr(adlar(i)) & ": bulunamadi"
        ElseIf ws.ProtectContents Then
            sonuc = sonuc & vbCrLf & ws.Name & ": korumali; atlandi"
        Else
            ' KPI kartlarinin altinda kalan eski hucre kenarliklarini kaldir.
            ' Hucre degerlerine, formullere ve kart sekillerine dokunmaz.
            ws.Range("A3:B5").Borders.LineStyle = xlNone
            ws.Range("A3:B5").Interior.Pattern = xlNone

            ' Yalnizca alt sagdaki AKSORA filigran resmini bul.
            Set logo = AKSORA_InceLogo(ws)
            If Not logo Is Nothing Then
                Set hedef = ws.Range("M24:P" & CLng(altSatirlar(i)))
                With logo
                    .LockAspectRatio = msoTrue
                    .Height = hedef.Height
                    If .Width > hedef.Width Then .Width = hedef.Width
                    .Left = hedef.Left + (hedef.Width - .Width) / 2
                    .Top = hedef.Top + (hedef.Height - .Height) / 2
                    .Placement = xlFreeFloating
                    .Locked = True
                End With
            Else
                sonuc = sonuc & vbCrLf & ws.Name & ": alt sag amblem bulunamadi"
            End If
        End If
        Set ws = Nothing
        Set logo = Nothing
    Next i

    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    MsgBox "Ince duzeltme tamamlandi." & sonuc, vbInformation, "AKSORA"
    Exit Sub
Hata:
    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    MsgBox "Islem durdu: " & Err.Description, vbExclamation, "AKSORA"
End Sub

Private Function AKSORA_InceSayfa(ByVal istenen As String) As Worksheet
    Dim ws As Worksheet, a As String, b As String
    b = AKSORA_InceNorm(istenen)
    For Each ws In ThisWorkbook.Worksheets
        a = AKSORA_InceNorm(ws.Name)
        If a = b Then
            Set AKSORA_InceSayfa = ws
            Exit Function
        End If
    Next ws
End Function

Private Function AKSORA_InceNorm(ByVal s As String) As String
    s = Replace(s, ChrW(304), "I")
    s = Replace(s, ChrW(305), "I")
    s = Replace(s, "i", "I")
    s = Replace(s, "ü", "U")
    s = Replace(s, "Ü", "U")
    AKSORA_InceNorm = UCase$(s)
End Function

Private Function AKSORA_InceLogo(ByVal ws As Worksheet) As Shape
    Dim sh As Shape, aday As Shape, n As String
    For Each sh In ws.Shapes
        If sh.Type = msoPicture Or sh.Type = msoLinkedPicture Then
            If sh.TopLeftCell.Row >= 20 And sh.TopLeftCell.Column >= 10 Then
                n = UCase$(sh.Name)
                If InStr(n, "AMBLEM") > 0 Or InStr(n, "FILIGRAN") > 0 Then
                    Set AKSORA_InceLogo = sh
                    Exit Function
                End If
                If aday Is Nothing Then Set aday = sh
            End If
        End If
    Next sh
    Set AKSORA_InceLogo = aday
End Function

' Mevcut modAKSORA_DIGER_RAPORLAR modulunun EN ALTINA ekleyin.
' Ikinci Option Explicit eklemeyin.
' Calistirilacak makro: AKSORA_UC_RAPOR_AMBLEM_KONUM

Public Sub AKSORA_UC_RAPOR_AMBLEM_KONUM()
    Dim adlar As Variant, baslar As Variant, bitisler As Variant
    Dim i As Long, ws As Worksheet, sh As Shape, hedef As Shape
    Dim alan As Range, rapor As String, adet As Long
    Dim eskiEkran As Boolean, eskiOlay As Boolean

    adlar = Array("EPP_RAPOR", "ÝLÇE_RAPORLAR", "YÜKLENÝCÝ_RAPOR")
    baslar = Array(24, 24, 25)
    bitisler = Array(35, 34, 40)
    eskiEkran = application.ScreenUpdating
    eskiOlay = application.EnableEvents
    On Error GoTo Hata
    application.ScreenUpdating = False
    application.EnableEvents = False

    For i = LBound(adlar) To UBound(adlar)
        Set ws = AK_UC_RAPOR_BUL(CStr(adlar(i)))
        If ws Is Nothing Then
            rapor = rapor & vbCrLf & CStr(adlar(i)) & ": sayfa bulunamadi"
        ElseIf ws.ProtectDrawingObjects Then
            rapor = rapor & vbCrLf & ws.Name & ": sekiller korumali; atlandi"
        Else
            Set hedef = Nothing
            For Each sh In ws.Shapes
                If sh.Type = msoPicture Or sh.Type = msoLinkedPicture Then
                    ' Yalnizca tablonun saginda, alt kisimdaki amblem resmini sec.
                    ' Ust banner veya buton resimleri kesinlikle secilmez.
                    If sh.Top >= ws.Rows(20).Top And sh.Left >= ws.Range("J1").Left Then
                        If hedef Is Nothing Then
                            Set hedef = sh
                        ElseIf InStr(1, sh.Name, "AMBLEM", vbTextCompare) > 0 Or _
                               InStr(1, sh.Name, "FILIGRAN", vbTextCompare) > 0 Then
                            Set hedef = sh
                        End If
                    End If
                End If
            Next sh

            If hedef Is Nothing Then
                rapor = rapor & vbCrLf & ws.Name & ": alt amblem resmi bulunamadi"
            Else
                Set alan = ws.Range("M" & baslar(i) & ":Q" & bitisler(i))
                With hedef
                    .LockAspectRatio = msoTrue
                    .Placement = xlFreeFloating
                    .Locked = True
                    ' Buyutme/kucultme sadece verilen satir ve sutun araligina sigdirir.
                    .Height = alan.Height
                    If .Width > alan.Width Then .Width = alan.Width
                    .Left = alan.Left + (alan.Width - .Width) / 2
                    .Top = alan.Top + (alan.Height - .Height) / 2
                End With
                adet = adet + 1
            End If
        End If
    Next i

    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    MsgBox adet & " rapor amblemi konumlandirildi." & rapor, vbInformation, "AKSORA"
    Exit Sub
Hata:
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    MsgBox "Islem durdu: " & Err.Description & rapor, vbExclamation, "AKSORA"
End Sub

Private Function AK_UC_RAPOR_BUL(ByVal aranan As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If AK_UC_AD_NORM(ws.Name) = AK_UC_AD_NORM(aranan) Then
            Set AK_UC_RAPOR_BUL = ws
            Exit Function
        End If
    Next ws
End Function

Private Function AK_UC_AD_NORM(ByVal s As String) As String
    s = Replace(s, ChrW(304), "I")
    s = Replace(s, ChrW(305), "I")
    s = Replace(s, "i", "I")
    s = Replace(s, "ü", "U")
    s = Replace(s, "Ü", "U")
    s = Replace(s, "þ", "S")
    s = Replace(s, "Þ", "S")
    s = Replace(s, "ð", "G")
    s = Replace(s, "Ð", "G")
    s = Replace(s, "ö", "O")
    s = Replace(s, "Ö", "O")
    s = Replace(s, "ç", "C")
    s = Replace(s, "Ç", "C")
    AK_UC_AD_NORM = UCase$(s)
End Function
Private Function AK_GercekStokSatiri(ByVal rapor As Worksheet) As String

    Dim kaynak As Worksheet
    Dim ad As String, son As Long, adet As Long

    Select Case AK_Norm(rapor.Name)
        Case "CDP_RAPOR": ad = "CDP_STOK"
        Case "EPP_RAPOR": ad = "EPP_STOK"
        Case "RAPORLAR": ad = "STOKLAR"
        Case "ILCE_RAPORLAR": ad = "ILCE_STOK"
        Case "YUKLENICI_RAPOR": ad = "YUKLENICI_STOK"
        Case Else
            Err.Raise vbObjectError + 7120, , _
                "Rapor kaynagi tanimsiz: " & rapor.Name
    End Select

    Set kaynak = AK_Sayfa(ad)

    If kaynak Is Nothing Then
        Err.Raise vbObjectError + 7121, , _
            "Stok sayfasi bulunamadi: " & ad
    End If

    son = kaynak.Cells(kaynak.Rows.Count, "G").End(xlUp).Row

    If son >= 7 Then
        adet = application.WorksheetFunction.CountA( _
            kaynak.Range("G7:G" & son))
    End If

    AK_GercekStokSatiri = AKSORA_SAYI_METNI(adet, "#,##0")

End Function

Public Sub AKSORA_RAPOR_KARTLARI_YENILE(ByVal sayfaAdi As String)

    Dim ws As Worksheet, kaynak As Worksheet
    Dim kaynakAdi As String, onEk As String
    Dim son As Long, toplam As Double
    Dim koruma As Boolean
    Dim hataNo As Long, hataMetni As String

    On Error GoTo Hata

    Set ws = AK_Sayfa(sayfaAdi)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 7130, , _
            "Rapor bulunamadi: " & sayfaAdi
    End If

    Select Case AK_Norm(ws.Name)
        Case "CDP_RAPOR": kaynakAdi = "CDP_STOK"
        Case "EPP_RAPOR": kaynakAdi = "EPP_STOK"
        Case "RAPORLAR": kaynakAdi = "STOKLAR"
        Case "ILCE_RAPORLAR": kaynakAdi = "ILCE_STOK"
        Case "YUKLENICI_RAPOR": kaynakAdi = "YUKLENICI_STOK"
        Case Else: Exit Sub
    End Select

    Set kaynak = AK_Sayfa(kaynakAdi)
    If kaynak Is Nothing Then
        Err.Raise vbObjectError + 7131, , _
            "Stok bulunamadi: " & kaynakAdi
    End If

    onEk = "AKSORA_RP_"
    If AK_Norm(ws.Name) = "CDP_RAPOR" Then
        onEk = "AKSORA_CDP_V1_"
    End If

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 7132, , _
            "Rapor hucreleri korumali: " & ws.Name
    End If

    koruma = ws.ProtectDrawingObjects
    If koruma Then ws.Unprotect Password:=""

    son = kaynak.Cells(kaynak.Rows.Count, "G").End(xlUp).Row
    If son >= 7 Then
        toplam = application.WorksheetFunction.Sum( _
            kaynak.Range("K7:K" & son))
    End If

    With ws.Shapes(onEk & "DEGER_A3_E6")
        If IsDate(kaynak.Range("C2").Value) Then
            .TextFrame2.TextRange.Text = _
                Format$(kaynak.Range("C2").Value, "dd.mm.yyyy hh:nn:ss")
        Else
            .TextFrame2.TextRange.Text = "Guncelleme bekleniyor"
        End If
    End With

    ws.Shapes(onEk & "DEGER_F3_J6").TextFrame2.TextRange.Text = _
        AK_GercekStokSatiri(ws)

    ws.Shapes(onEk & "DEGER_K3_Q6").TextFrame2.TextRange.Text = _
        AKSORA_SAYI_METNI(toplam, "#,##0.00") & " TL"

    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=False, Scenarios:=False, _
            UserInterfaceOnly:=True, AllowFiltering:=True
    End If

    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description

    On Error Resume Next
    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=False, Scenarios:=False, _
            UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    On Error GoTo 0

    Err.Raise hataNo, "AKSORA_RAPOR_KARTLARI_YENILE", hataMetni

End Sub

Public Sub AKSORA_RAPOR_GRAFIK_YENILE(ByVal sayfaAdi As String)

    Dim ws As Worksheet, sh As Shape
    Dim r As Long, toplamSatiri As Long, degerKol As Long
    Dim i As Long, ad As String
    Dim toplam As Double, deger As Double
    Dim koruma As Boolean, olay As Boolean, ekran As Boolean
    Dim hataNo As Long, hataMetni As String

    olay = application.EnableEvents
    ekran = application.ScreenUpdating
    On Error GoTo Hata

    Set ws = AK_Sayfa(sayfaAdi)
    If ws Is Nothing Then
        Err.Raise vbObjectError + 7140, , _
            "Rapor bulunamadi: " & sayfaAdi
    End If

    Select Case AK_Norm(ws.Name)
        Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR"
            degerKol = 4
        Case "ILCE_RAPORLAR", "YUKLENICI_RAPOR"
            degerKol = 5
        Case Else
            Exit Sub
    End Select

    For r = 9 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If AK_Norm(CStr(ws.Cells(r, 1).Value2)) = "TOPLAM" Then
            toplamSatiri = r
            Exit For
        End If
    Next r

    If toplamSatiri < 10 Then
        Err.Raise vbObjectError + 7141, , _
            "TOPLAM satiri bulunamadi: " & ws.Name
    End If

    For r = 9 To toplamSatiri - 1
        If Not IsNumeric(ws.Cells(r, degerKol).Value2) Then
            Err.Raise vbObjectError + 7142, , _
                "Sayisal olmayan stok degeri: " & ws.Name & ", satir " & r
        End If

        deger = CDbl(ws.Cells(r, degerKol).Value2)
        If deger < 0 Then
            Err.Raise vbObjectError + 7143, , _
                "Negatif stok degeri halka grafikte gosterilemez."
        End If
        toplam = toplam + deger
    Next r

    If Not IsNumeric(ws.Cells(toplamSatiri, degerKol).Value2) Then
        Err.Raise vbObjectError + 7144, , "TOPLAM degeri gecersiz."
    End If

    If Abs(toplam - CDbl(ws.Cells(toplamSatiri, degerKol).Value2)) > 0.02 Then
        Err.Raise vbObjectError + 7145, , _
            "Depo degerleri ile TOPLAM uyusmuyor: " & ws.Name
    End If

    If ws.ProtectContents Then
        Err.Raise vbObjectError + 7146, , _
            "Rapor hucreleri korumali: " & ws.Name
    End If

    application.EnableEvents = False
    application.ScreenUpdating = False

    koruma = ws.ProtectDrawingObjects
    If koruma Then ws.Unprotect Password:=""

    For i = ws.ChartObjects.Count To 1 Step -1
        ad = ws.ChartObjects(i).Name
        Select Case ad
            Case "AKSORA_RP_HALKA_GRAFIK", _
                 "AKSORA_CDP_HALKA_GRAFIK", _
                 "AKSORA_CDP_V1_GRAFIK"
                ws.ChartObjects(i).Delete
        End Select
    Next i

    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        ad = sh.Name

        If Left$(ad, Len("AKSORA_RP_HALKA_")) = "AKSORA_RP_HALKA_" Or _
           Left$(ad, Len("AKSORA_CDP_H3_")) = "AKSORA_CDP_H3_" Or _
           ad = "AKSORA_CDP_HALKA_MERKEZ" Or _
           ad = "AKSORA_YUK_HALKA_MERKEZ" Then
            sh.Delete
        End If
    Next i

    If toplam > 0 Then
        AK_Grafik ws, toplamSatiri - 1, degerKol, toplamSatiri
    End If

    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=False, Scenarios:=False, _
            UserInterfaceOnly:=True, AllowFiltering:=True
    End If

    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description

    On Error Resume Next
    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=False, Scenarios:=False, _
            UserInterfaceOnly:=True, AllowFiltering:=True
    End If
    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    On Error GoTo 0

    Err.Raise hataNo, "AKSORA_RAPOR_GRAFIK_YENILE", hataMetni

End Sub
Public Sub AKSORA_RAPOR_ALT_YAZI_YERLESTIR(ByVal sayfaAdi As String)

    Dim ws As Worksheet, sh As Shape
    Dim r As Long, toplamSatiri As Long
    Dim metin As String
    Dim bulundu As Boolean
    Dim koruma As Boolean
    Dim hataNo As Long, hataMetni As String

    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(sayfaAdi)

    For r = 9 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If UCase$(Trim$(CStr(ws.Cells(r, "A").Value2))) = "TOPLAM" Then
            toplamSatiri = r
            Exit For
        End If
    Next r

    If toplamSatiri = 0 Then
        Err.Raise vbObjectError + 7150, , "TOPLAM satýrý bulunamadý."
    End If

    koruma = ws.ProtectDrawingObjects
    If koruma Then ws.Unprotect Password:=""

    r = toplamSatiri + 1
    ws.Rows(r).RowHeight = 23

    For Each sh In ws.Shapes
        If sh.Type = msoTextBox Or sh.Type = msoAutoShape Then
            If sh.TextFrame2.HasText Then
                metin = sh.TextFrame2.TextRange.Text

                If InStr(1, metin, "Veriler SAP", vbTextCompare) > 0 Then
                    bulundu = True

                    With sh
                        .Placement = xlFreeFloating
                        .Left = ws.Cells(r, "A").Left + 3
                        .Top = ws.Cells(r, "A").Top + 2
                        .Width = ws.Range("A" & r & ":E" & r).Width - 6
                        .Height = 19
                        .Locked = True
                        .OnAction = vbNullString
                        .Visible = msoTrue
                        .ZOrder msoBringToFront

                        With .TextFrame2
                            .AutoSize = msoAutoSizeNone
                            .WordWrap = msoFalse
                            .MarginLeft = 0
                            .MarginRight = 0
                            .MarginTop = 0
                            .MarginBottom = 0
                            .TextRange.Font.Size = 8
                        End With
                    End With
                End If
            End If
        End If
    Next sh

    If Not bulundu Then
        Set sh = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            ws.Cells(r, "A").Left + 3, _
            ws.Cells(r, "A").Top + 2, _
            ws.Range("A" & r & ":E" & r).Width - 6, 19)

        With sh
            .Name = "AKSORA_RP_ALT"
            .Placement = xlFreeFloating
            .Fill.Visible = msoFalse
            .line.Visible = msoFalse
            .Locked = True

            With .TextFrame2
                .AutoSize = msoAutoSizeNone
                .WordWrap = msoFalse
                .MarginLeft = 0
                .MarginRight = 0
                .MarginTop = 0
                .MarginBottom = 0
                .TextRange.Text = "AKSORA | " & AK_Baslik(ws.Name) & _
                    " | Veriler SAP raporundan alinir."
                .TextRange.Font.Name = "Aptos"
                .TextRange.Font.Size = 8
                .TextRange.Font.Fill.ForeColor.RGB = RGB(99, 124, 153)
            End With
        End With
    End If

    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=False, Scenarios:=False, _
            UserInterfaceOnly:=True, AllowFiltering:=True
    End If

    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description

    On Error Resume Next
    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=False, Scenarios:=False, _
            UserInterfaceOnly:=True, AllowFiltering:=True
    End If

    On Error GoTo 0
    Err.Raise hataNo, "AKSORA_RAPOR_ALT_YAZI_YERLESTIR", hataMetni

End Sub
Public Sub AKSORA_RAPOR_AMBLEM_HIZALA(ByVal sayfaAdi As String)
    Dim ws As Worksheet, logo As Shape, eskiLogo As Shape, alan As Range
    Dim yol As String
    Dim logoAdi As String, koruma As Boolean, icerik As Boolean, senaryo As Boolean
    Dim secim As XlEnableSelection, hataNo As Long, hataMetni As String
    On Error GoTo Hata
    Set ws = AK_Sayfa(sayfaAdi)
    If ws Is Nothing Then Err.Raise vbObjectError + 7170, , "Rapor bulunamadi."
    If AKSORA_TASARIM_AKTIF() Then
        AKSORA_TASARIM_RAPOR_LOGO ws
        Exit Sub
    End If
    logoAdi = "AKSORA_RP_FILIGRAN"
    If AK_Norm(ws.Name) = "CDP_RAPOR" Then logoAdi = "AKSORA_CDP_V1_FILIGRAN"
    On Error Resume Next
    Set eskiLogo = ws.Shapes(logoAdi)
    On Error GoTo Hata
    koruma = ws.ProtectDrawingObjects
    icerik = ws.ProtectContents
    senaryo = ws.ProtectScenarios
    secim = ws.EnableSelection
    If koruma Then ws.Unprotect Password:=""
    If Not eskiLogo Is Nothing Then
        If eskiLogo.AlternativeText = "AKSORA_FILIGRAN_DUZGUN_V2" Then Set logo = eskiLogo
    End If
    If logo Is Nothing Then
        yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_RAPOR_FILIGRAN.png"
        If Len(Dir$(yol)) = 0 Then
            Err.Raise vbObjectError + 7171, , "AKSORA_RAPOR_FILIGRAN.png dosyasini Excel dosyasinin yanina kopyalayin."
        End If
        ' Duzgun kaynak resim once eklenir; basariliysa eski bozuk resim kaldirilir.
        Set logo = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
        If Not eskiLogo Is Nothing Then eskiLogo.Delete
        logo.Name = logoAdi
        logo.AlternativeText = "AKSORA_FILIGRAN_DUZGUN_V2"
    End If
    ' R:T, grafik kartinin saginda ve ustteki donus butonunun altinda.
    Set alan = ws.Range("R10:T15")
    With logo
        .Rotation = 0
        On Error Resume Next
        .LockAspectRatio = msoFalse
        .ScaleHeight 1, msoTrue: .ScaleWidth 1, msoTrue
        On Error GoTo Hata
        .LockAspectRatio = msoTrue
        .Width = 80
        If .Height > 80 Then .Height = 80
        .Left = alan.Left + (alan.Width - .Width) / 2
        .Top = alan.Top
        .Placement = xlFreeFloating
        .Locked = True
        .Visible = msoTrue
        .OnAction = vbNullString
        .ZOrder msoBringToFront
    End With
    ' Ayni orijinal soluk PNG tum raporlarda kullanilir; alpha/seffaflik korunur.
    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True, AllowFiltering:=True
        ws.EnableSelection = secim
    End If
    Exit Sub
Hata:
    hataNo = Err.Number: hataMetni = Err.Description
    On Error Resume Next
    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True, AllowFiltering:=True
        ws.EnableSelection = secim
    End If
    On Error GoTo 0
    Err.Raise hataNo, "AKSORA_RAPOR_AMBLEM_HIZALA", hataMetni
End Sub

Public Sub AKSORA_RAPOR_TABLO_BITISI(ByVal sayfaAdi As String)

    Dim ws As Worksheet
    Dim tablo As ListObject
    Dim r As Long, toplam As Long, sonKol As Long
    Dim eskiAlt As Long
    Dim hedef As Range

    Set ws = ThisWorkbook.Worksheets(sayfaAdi)

    Select Case AK_Norm(ws.Name)
        Case "ILCE_RAPORLAR", "YUKLENICI_RAPOR"
            sonKol = 5
        Case Else
            sonKol = 4
    End Select

    For r = 9 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If UCase$(Trim$(CStr(ws.Cells(r, "A").Value2))) = "TOPLAM" Then
            toplam = r
            Exit For
        End If
    Next r

    If toplam = 0 Then
        Err.Raise vbObjectError + 7180, , "TOPLAM satýrý bulunamadý."
    End If

    eskiAlt = toplam + 1
    Set hedef = ws.Range(ws.Cells(8, 1), ws.Cells(toplam, sonKol))

    For Each tablo In ws.ListObjects
        If tablo.HeaderRowRange.Row = 8 And _
           tablo.Range.Column = 1 Then

            r = tablo.Range.Row + tablo.Range.Rows.Count - 1
            If r > eskiAlt Then eskiAlt = r

            tablo.Resize hedef
        End If
    Next tablo

    ws.Range(ws.Cells(toplam + 1, 1), _
             ws.Cells(eskiAlt, sonKol)).Borders.LineStyle = xlNone

    With ws.Range(ws.Cells(toplam, 1), _
                  ws.Cells(toplam, sonKol)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .Color = RGB(31, 78, 121)
    End With

End Sub





