Attribute VB_Name = "modAKSORA_GRAFIK"
Option Explicit

' MEVCUT halka grafik modulunun TAMAMINI bununla degistirin.
' Calistir: AKSORA_CDP_HALKA_GRAFIK
' Yalnizca F8:Q23 grafik alanini yeniden olusturur.
' KPI, tablo, banner, amblem ve ANA PANEL butonuna dokunmaz.

Public Sub AKSORA_CDP_HALKA_GRAFIK()
    Dim ws As Worksheet, co As ChartObject, CH As Chart, sr As Series
    Dim alan As Range, r As Long, son As Long, toplamSatiri As Long
    Dim i As Long, n As Long, genel As Double, deger As Double
    Dim adlar() As String, tutarlar() As Double, renkler As Variant
    Dim cx As Double, cy As Double, pw As Double, ph As Double
    Dim s As Shape, x As Double, y As Double, satirY As Double
    Dim sc As Boolean, ev As Boolean, Hata As String
    Dim panelLeft As Double, panelWidth As Double
    Dim minPay As Double

    Set ws = ThisWorkbook.Worksheets("CDP_RAPOR")
    If ws.ProtectContents Then
        MsgBox "CDP_RAPOR korumali; islem yapilmadi.", vbExclamation, "AKSORA"
        Exit Sub
    End If
    For r = 9 To ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If UCase$(Trim$(CStr(ws.Cells(r, "A").Value2))) = "TOPLAM" Then
            toplamSatiri = r
            Exit For
        End If
    Next r
    If toplamSatiri < 10 Then
        MsgBox "TOPLAM satiri bulunamadi.", vbExclamation, "AKSORA"
        Exit Sub
    End If

    ReDim adlar(1 To toplamSatiri - 9)
    ReDim tutarlar(1 To toplamSatiri - 9)
    For r = 9 To toplamSatiri - 1
        If IsNumeric(ws.Cells(r, "D").Value2) Then
            deger = CDP_Halka_Sayi(ws.Cells(r, "D").Value2)
            If deger > 0 Then
                n = n + 1
                adlar(n) = CStr(ws.Cells(r, "B").Value2)
                tutarlar(n) = deger
                genel = genel + deger
            End If
        End If
    Next r
    If genel <= 0 Then
        MsgBox "Pozitif stok degeri bulunamadi.", vbExclamation, "AKSORA"
        Exit Sub
    End If
    ReDim Preserve adlar(1 To n)
    ReDim Preserve tutarlar(1 To n)

    renkler = Array(RGB(34, 77, 153), RGB(63, 151, 226), _
                   RGB(34, 166, 146), RGB(243, 166, 65), _
                   RGB(117, 91, 187), RGB(227, 104, 104), _
                   RGB(105, 140, 168), RGB(81, 175, 113), _
                   RGB(224, 143, 84), RGB(145, 121, 172), _
                   RGB(98, 123, 163), RGB(170, 191, 207))

    sc = application.ScreenUpdating
    ev = application.EnableEvents
    On Error GoTo Hata
    application.ScreenUpdating = False
    application.EnableEvents = False

    ' Yalnizca bu makronun olusturdugu grafik ve aciklamalari temizle.
    For i = ws.Shapes.Count To 1 Step -1
        Set s = ws.Shapes(i)
        If Left$(s.Name, Len("AKSORA_CDP_H3_")) = "AKSORA_CDP_H3_" Or _
           s.Name = "AKSORA_CDP_HALKA_MERKEZ" Then s.Delete
    Next i
    For i = ws.ChartObjects.Count To 1 Step -1
        If ws.ChartObjects(i).Name = "AKSORA_CDP_HALKA_GRAFIK" Or _
           ws.ChartObjects(i).Name = "AKSORA_CDP_V1_GRAFIK" Then
            ws.ChartObjects(i).Delete
        End If
    Next i

    Set alan = ws.Range("F8:Q23")
    x = alan.Left
    y = alan.Top
    panelWidth = alan.Width * 0.31
    panelLeft = x + alan.Width - panelWidth

    ' Baslik ve arka plan.
    Set s = ws.Shapes.AddShape(msoShapeRectangle, x, y, alan.Width, alan.Height)
    With s
        .Name = "AKSORA_CDP_H3_ZEMIN"
        .Fill.ForeColor.RGB = RGB(245, 250, 255)
        .line.Visible = msoFalse
        .Placement = xlFreeFloating
    End With
    Set s = CDP_H3_Yazi(ws, "AKSORA_CDP_H3_BASLIK", _
            "DEPOLARA GÖRE STOK DAÐILIMI", x + 12, y + 7, _
            alan.Width - 24, 25, 13, RGB(26, 57, 88), True, msoAlignCenter)

    ' Legend kapali: halkanin gercek merkezini hesaplamak icin
    ' grafik alani yalnizca sol bolumde tutulur.
    Set co = ws.ChartObjects.Add(x + 22, y + 42, _
             alan.Width - panelWidth - 34, alan.Height - 62)
    co.Name = "AKSORA_CDP_HALKA_GRAFIK"
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
    ' Excel yerlesimi hesaplasin; merkez sabit yuzdeyle tahmin edilmez.
    CH.Refresh
    DoEvents
    pw = CH.PlotArea.InsideWidth
    ph = CH.PlotArea.InsideHeight
    cx = co.Left + CH.PlotArea.InsideLeft + pw / 2
    cy = co.Top + CH.PlotArea.InsideTop + ph / 2

    ' Merkez etiketi dogrudan gercek grafik merkezine yerlestirilir.
    Set s = CDP_H3_Yazi(ws, "AKSORA_CDP_H3_MERKEZ", _
            AKSORA_SAYI_METNI(genel, "#,##0.00") & " TL" & vbCrLf & _
            "TOPLAM STOK DEÐERÝ", cx - 122, cy - 30, _
            244, 60, 14, RGB(26, 57, 88), True, msoAlignCenter)
    s.TextFrame2.VerticalAnchor = msoAnchorMiddle
    s.ZOrder msoBringToFront

    ' Sagda ayri, hizali yuzde listesi. Cok kucuk paylar da listelenir.
    Set s = CDP_H3_Yazi(ws, "AKSORA_CDP_H3_LISTE_BASLIK", _
            "DEPO                 PAY", panelLeft + 5, y + 43, _
            panelWidth - 16, 20, 10, RGB(26, 57, 88), True, msoAlignLeft)
    For i = 1 To n
        satirY = y + 66 + (i - 1) * ((alan.Height - 90) / n)
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                panelLeft + 6, satirY + 4, 9, 9)
        With s
            .Name = "AKSORA_CDP_H3_RENK_" & i
            .Fill.ForeColor.RGB = renkler((i - 1) Mod 12)
            .line.Visible = msoFalse
            .Placement = xlFreeFloating
        End With
        Set s = CDP_H3_Yazi(ws, "AKSORA_CDP_H3_AD_" & i, _
                adlar(i), panelLeft + 22, satirY, _
                panelWidth - 90, 19, 9, RGB(35, 56, 78), False, msoAlignLeft)
        Set s = CDP_H3_Yazi(ws, "AKSORA_CDP_H3_PAY_" & i, _
                AKSORA_SAYI_METNI(tutarlar(i) / genel, "0.0%"), _
                panelLeft + panelWidth - 69, satirY, 57, 19, _
                10, RGB(26, 57, 88), True, msoAlignRight)
    Next i

    application.EnableEvents = ev
    application.ScreenUpdating = sc
    MsgBox "Halka grafik yenilendi.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    Hata = Err.Description
    application.EnableEvents = ev
    application.ScreenUpdating = sc
    MsgBox "Grafik olusturulamadi: " & Hata, vbExclamation, "AKSORA"
End Sub

Private Function CDP_H3_Yazi(ByVal ws As Worksheet, ByVal ad As String, _
        ByVal metin As String, ByVal sol As Double, ByVal ust As Double, _
        ByVal gen As Double, ByVal yuk As Double, ByVal punto As Single, _
        ByVal renk As Long, ByVal kalin As Boolean, _
        ByVal hiz As MsoParagraphAlignment) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, sol, ust, gen, yuk)
    With s
        .Name = ad
        .Fill.Visible = msoFalse
        .line.Visible = msoFalse
        .TextFrame2.MarginLeft = 1
        .TextFrame2.MarginRight = 1
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .TextFrame2.WordWrap = msoFalse
        .TextFrame2.TextRange.Text = metin
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = punto
        .TextFrame2.TextRange.Font.bold = kalin
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = renk
        .TextFrame2.TextRange.ParagraphFormat.Alignment = hiz
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .Placement = xlFreeFloating
        .Locked = True
    End With
    Set CDP_H3_Yazi = s
End Function

Private Function CDP_Halka_Sayi(ByVal deger As Variant) As Double
    If Not IsError(deger) Then
        If IsNumeric(deger) Then CDP_Halka_Sayi = CDbl(deger)
    End If
End Function
' Yalnizca YUKLENICI_RAPOR sayfasindaki halka grafigin merkez yazisini duzeltir.
' Diger raporlara, verilere, butonlara ve mevcut grafik serilerine dokunmaz.
Public Sub AKSORA_YUKLENICI_HALKA_YAZI_DUZELT()
    Dim ws As Worksheet, co As ChartObject, hedef As ChartObject
    Dim sh As Shape, ic As Shape, yeni As Shape
    Dim cx As Double, cy As Double, kutuW As Double, kutuH As Double
    Dim toplam As Variant, toplamMetni As String, i As Long
    Dim eskiEkran As Boolean, eskiOlay As Boolean
    Dim korunuyor As Boolean

    Set ws = AK_YukleniciSayfa()
    If ws Is Nothing Then
        MsgBox "YUKLENICI_RAPOR sayfasi bulunamadi.", vbExclamation, "AKSORA"
        Exit Sub
    End If

    eskiEkran = application.ScreenUpdating
    eskiOlay = application.EnableEvents
    On Error GoTo Hata
    application.ScreenUpdating = False
    application.EnableEvents = False

    korunuyor = ws.ProtectDrawingObjects
    If korunuyor Then
        On Error Resume Next
        ws.Unprotect Password:=""
        On Error GoTo Hata
        If ws.ProtectDrawingObjects Then Err.Raise vbObjectError + 201, , "Sayfa sifreyle korundugu icin sekiller duzenlenemiyor."
    End If

    ' Grafigi turune gore bul: grafiklerin sirasi degisse bile dogru olani sec.
    For Each co In ws.ChartObjects
        If co.Chart.ChartType = xlDoughnut Or co.Chart.ChartType = xlDoughnutExploded Then
            Set hedef = co
            Exit For
        End If
    Next co
    If hedef Is Nothing Then Err.Raise vbObjectError + 202, , "Halka grafik bulunamadi."

    ' Halka ortasinin kesin konumu: grafik PlotArea'sinin ic merkezi.
    With hedef.Chart.PlotArea
        cx = hedef.Left + .InsideLeft + .InsideWidth / 2
        cy = hedef.Top + .InsideTop + .InsideHeight / 2
    End With

    ' Toplam degerini mevcut rapor tablosunun TOPLAM satirindan al.
    toplam = Empty
    For i = 9 To 1000
        If UCase$(Trim$(CStr(ws.Cells(i, 1).Value2))) = "TOPLAM" Then
            toplam = ws.Cells(i, 5).Value2
            Exit For
        End If
        If i > 12 And Len(CStr(ws.Cells(i, 1).Value2)) = 0 Then Exit For
    Next i
    If IsNumeric(toplam) Then
        toplamMetni = AKSORA_SAYI_METNI(CDbl(toplam), "#,##0.00") & " TL"
    Else
        ' Tablonun gorunen degeri degerlendirilir; veri degistirilmez.
        If i <= 1000 Then toplamMetni = CStr(ws.Cells(i, 5).Text)
        If Len(toplamMetni) = 0 Then Err.Raise vbObjectError + 203, , "TOPLAM stok degeri bulunamadi."
    End If

    ' Onceki duzeltmenin ekledigi metin kutusunu temizle.
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If sh.Name = "AKSORA_YUK_HALKA_MERKEZ" Then sh.Delete
    Next i

    ' Onceki merkez yazisi grafik icine eklenmisse yalnizca onu temizle.
    On Error Resume Next
    For i = hedef.Chart.Shapes.Count To 1 Step -1
        Set ic = hedef.Chart.Shapes(i)
        If ic.HasTextFrame Then
            If InStr(1, ic.TextFrame2.TextRange.Text, "TOPLAM STOK", vbTextCompare) > 0 Then ic.Delete
        End If
    Next i
    On Error GoTo Hata

    ' Onceki merkez yazisi sayfa ustunde bir metin kutusuysa yalnizca
    ' grafigin orta bolgesindeki ilgili metin kutusunu temizle.
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If sh.Type = msoTextBox Or sh.Type = msoAutoShape Then
            If sh.Name <> "AKSORA_YUK_HALKA_MERKEZ" Then
                If sh.Left < cx + 110 And sh.Left + sh.Width > cx - 110 _
                   And sh.Top < cy + 85 And sh.Top + sh.Height > cy - 85 Then
                    On Error Resume Next
                    If sh.TextFrame2.HasText Then
                        If InStr(1, sh.TextFrame2.TextRange.Text, "TOPLAM STOK", vbTextCompare) > 0 Then sh.Delete
                    End If
                    On Error GoTo Hata
                End If
            End If
        End If
    Next i

    ' Deligin icine tasma yapmayacak boyutta, ortalanmis yeni metin.
    kutuW = 122
    kutuH = 43
    Set yeni = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                   cx - kutuW / 2, cy - kutuH / 2, kutuW, kutuH)
    With yeni
        .Name = "AKSORA_YUK_HALKA_MERKEZ"
        .Fill.Visible = msoFalse
        .line.Visible = msoFalse
        .Placement = xlFreeFloating
        .TextFrame2.TextRange.Text = toplamMetni & vbCrLf & "TOPLAM STOK DEÐERÝ"
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(27, 53, 81)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0
        .Locked = True
        .ZOrder msoBringToFront
    End With

    If korunuyor Then ws.Protect Password:="", DrawingObjects:=True, Contents:=False, Scenarios:=False, UserInterfaceOnly:=True
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    MsgBox "YUKLENICI_RAPOR halka merkez yazisi duzeltildi.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    On Error Resume Next
    If korunuyor Then ws.Protect Password:="", DrawingObjects:=True, Contents:=False, Scenarios:=False, UserInterfaceOnly:=True
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    MsgBox "Duzeltme yapilamadi: " & Err.Description, vbExclamation, "AKSORA"
End Sub

Private Function AK_YukleniciSayfa() As Worksheet
    Dim ws As Worksheet, n As String
    For Each ws In ThisWorkbook.Worksheets
        n = UCase$(ws.Name)
        n = Replace(n, "Ü", "U")
        n = Replace(n, "Ý", "I")
        n = Replace(n, "ý", "I")
        If n = "YUKLENICI_RAPOR" Then
            Set AK_YukleniciSayfa = ws
            Exit Function
        End If
    Next ws
End Function


