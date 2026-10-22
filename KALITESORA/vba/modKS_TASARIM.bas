Attribute VB_Name = "modKS_TASARIM"
Option Explicit

' KALÝTESORA görsel dili: lacivert bant, camgöbeði vurgu, beyaz kartlar, Segoe UI.

Public Const KS_FONT As String = "Segoe UI"
Public Const KS_STIL As String = "KS_STIL"

' Üst bant (satýr 1-2), menü (satýr 3) ve sayfa baþlýðý (satýr 5).
Public Sub KS_Bant(ByVal ws As Worksheet, ByVal sonKolon As Long, ByVal baslik As String, ByVal aciklama As String)
    Dim r As Range
    If sonKolon < 18 Then sonKolon = 18
    ws.Rows(1).RowHeight = 30
    ws.Rows(2).RowHeight = 18
    ws.Rows(3).RowHeight = 28
    ws.Rows(4).RowHeight = 8
    ws.Rows(5).RowHeight = 24
    Set r = ws.Range(ws.Cells(1, 1), ws.Cells(2, sonKolon))
    r.Interior.Color = C_LAC()
    r.Font.Color = vbWhite
    With ws.Range(ws.Cells(3, 1), ws.Cells(3, sonKolon))
        .Interior.Color = C_LAC2()
        .Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    ws.Range(ws.Cells(4, 1), ws.Cells(4, sonKolon)).Interior.Pattern = xlNone
    With ws.Cells(1, 2)
        .Value = KS_AD
        .Font.Name = KS_FONT
        .Font.Size = 17
        .Font.Bold = True
        .Font.Color = vbWhite
        .VerticalAlignment = xlBottom
        .HorizontalAlignment = xlLeft
    End With
    With ws.Cells(2, 2)
        .Value = "ÇORUH BÖLGESÝ  |  MALZEME KALÝTE · GARANTÝ · ARIZA TAKÝBÝ"
        .Font.Name = KS_FONT
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = C_CAMGOBEGI()
        .VerticalAlignment = xlTop
        .HorizontalAlignment = xlLeft
    End With
    With ws.Cells(1, sonKolon)
        .Value = "v" & KS_SURUM
        .Font.Size = 8
        .Font.Color = RGB(150, 175, 200)
        .HorizontalAlignment = xlRight
    End With
    With ws.Cells(5, 2)
        .Value = baslik & IIf(Len(aciklama) > 0, "    " & aciklama, "")
        .Font.Name = KS_FONT
        .Font.Size = 9
        .Font.Bold = False
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
        .WrapText = False
        With .Characters(1, Len(baslik)).Font
            .Size = 13
            .Bold = True
            .Color = C_LAC()
        End With
    End With
End Sub

' Her sayfadaki ortak gezinme menüsü; aktif sayfa camgöbeði ile vurgulanýr.
Public Sub KS_NavKur(ByVal ws As Worksheet)
    Dim hedef As Variant, etiket As Variant, i As Long, s As Shape, sol As Double
    Dim ust As Double, gen As Double, yuk As Double, aktif As Boolean
    KS_SekilSil ws, "ksNav_"
    hedef = KS_Sayfalar()
    etiket = Array("PANEL", "KAYIT", "SEVKÝYAT", "ARIZA", "MALZEME", "KARNE", _
        "PERFORMANS", "5 YIL", "TOPLU AKTAR", "TANIMLAR", "GÜNLÜK")
    sol = ws.Cells(3, 2).Left
    ust = ws.Cells(3, 2).Top + 3
    yuk = ws.Rows(3).RowHeight - 6
    gen = 80
    For i = LBound(hedef) To UBound(hedef)
        aktif = (StrComp(CStr(hedef(i)), ws.Name, vbTextCompare) = 0)
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, sol, ust, gen, yuk)
        s.Name = "ksNav_" & i
        s.AlternativeText = CStr(hedef(i))
        KS_SekilBicim s, CStr(etiket(i)), IIf(aktif, C_CAMGOBEGI(), C_LAC2()), _
            IIf(aktif, C_LAC(), RGB(200, 220, 240)), 8.5
        If Not aktif Then
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = RGB(45, 85, 128)
            s.Line.Weight = 0.75
        End If
        s.OnAction = "KS_NAV"
        sol = sol + gen + 4
    Next i
End Sub

Public Sub KS_NAV()
    Dim s As Shape, hedef As Worksheet
    On Error GoTo Son
    Set s = ActiveSheet.Shapes(CStr(Application.Caller))
    Set hedef = KS_WS(s.AlternativeText)
    If hedef Is Nothing Then Exit Sub
    If hedef.Name = SH_PANEL Then KS_PanelYenile False
    hedef.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
Son:
End Sub

Public Function KS_Buton(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
        ByVal sol As Double, ByVal ust As Double, ByVal gen As Double, ByVal yuk As Double, _
        ByVal makro As String, Optional ByVal tur As String = "ana") As Shape
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(ad).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, sol, ust, gen, yuk)
    s.Name = ad
    Select Case tur
        Case "ana": KS_SekilBicim s, metin, C_CAMGOBEGI(), C_LAC(), 9
        Case "koyu": KS_SekilBicim s, metin, C_LAC(), vbWhite, 9
        Case "yesil": KS_SekilBicim s, metin, C_YESIL(), vbWhite, 9
        Case "cizgi"
            KS_SekilBicim s, metin, vbWhite, C_LAC(), 9
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = C_LAC()
            s.Line.Weight = 1
        Case Else: KS_SekilBicim s, metin, RGB(226, 232, 240), C_METIN(), 9
    End Select
    s.OnAction = makro
    Set KS_Buton = s
End Function

Public Sub KS_SekilBicim(ByVal s As Shape, ByVal metin As String, ByVal zemin As Long, _
        ByVal yazi As Long, ByVal boyut As Single)
    On Error Resume Next
    s.Adjustments.Item(1) = 0.22
    On Error GoTo 0
    s.Fill.Visible = msoTrue
    s.Fill.Solid
    s.Fill.ForeColor.RGB = zemin
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    With s.TextFrame2
        .MarginLeft = 2
        .MarginRight = 2
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = metin
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = KS_FONT
            .Font.Size = boyut
            .Font.Bold = msoTrue
            .Font.Fill.ForeColor.RGB = yazi
        End With
    End With
End Sub

Public Sub KS_SekilSil(ByVal ws As Worksheet, ByVal onek As String)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, Len(onek)) = onek Then ws.Shapes(i).Delete
    Next i
End Sub

' Sayfanýn tamamýný temizler ve kart zemini hazýrlar (veri tablosu olmayan sayfalar için).
Public Sub KS_SayfaSifirla(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i
    ws.Cells.Validation.Delete
    ws.Cells.FormatConditions.Delete
    ws.Cells.Clear
    ws.Cells.ClearComments
    ws.Cells.Font.Name = KS_FONT
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = C_METIN()
    ws.Cells.Interior.Color = C_ZEMIN()
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.Locked = True
    ws.Columns(1).ColumnWidth = 2
    ws.Rows.RowHeight = 18
End Sub

' Bölüm baþlýðý: lacivert metin, altýnda camgöbeði çizgi.
Public Sub KS_Bolum(ByVal r As Range, ByVal metin As String)
    With r
        .Interior.Color = C_ZEMIN()
        .Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    With r.Cells(1, 1)
        .Value = metin
        .Font.Bold = True
        .Font.Size = 10.5
        .Font.Color = C_LAC()
    End With
    r.EntireRow.RowHeight = 22
End Sub

' Rapor tablosu baþlýk satýrý.
Public Sub KS_TabloBaslik(ByVal r As Range, ByVal basliklar As Variant)
    r.Value = basliklar
    With r
        .Interior.Color = C_LAC()
        .Font.Color = vbWhite
        .Font.Bold = True
        .Font.Size = 9
        .WrapText = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    r.EntireRow.RowHeight = 30
End Sub

' Rapor tablosu gövdesi: beyaz zemin, ince çizgi, zebra.
Public Sub KS_TabloGovde(ByVal r As Range)
    Dim i As Long
    With r
        .Interior.Color = vbWhite
        .Font.Size = 9
        .Font.Color = C_METIN()
        .Borders(xlInsideHorizontal).Color = C_CIZGI()
        .Borders(xlEdgeBottom).Color = C_CIZGI()
        .Borders(xlEdgeLeft).Color = C_CIZGI()
        .Borders(xlEdgeRight).Color = C_CIZGI()
    End With
    For i = 2 To r.Rows.Count Step 2
        r.Rows(i).Interior.Color = C_SERIT()
    Next i
End Sub

' KPI kartý: üstte etiket, ortada büyük deðer, altta açýklama; solda camgöbeði þerit.
Public Sub KS_KartCiz(ByVal r As Range, ByVal etiket As String)
    With r
        .Interior.Color = C_KART()
        .Borders(xlEdgeTop).Color = C_CIZGI()
        .Borders(xlEdgeBottom).Color = C_CIZGI()
        .Borders(xlEdgeRight).Color = C_CIZGI()
        .Borders(xlEdgeLeft).Color = C_CAMGOBEGI()
        .Borders(xlEdgeLeft).Weight = xlThick
    End With
    r.Rows(1).Merge
    r.Rows(2).Merge
    r.Rows(3).Merge
    With r.Cells(1, 1)
        .Value = etiket
        .Font.Size = 8
        .Font.Bold = True
        .Font.Color = C_GRI()
        .IndentLevel = 1
        .HorizontalAlignment = xlLeft
    End With
    With r.Cells(2, 1)
        .Font.Size = 20
        .Font.Bold = True
        .Font.Color = C_LAC()
        .IndentLevel = 1
        .HorizontalAlignment = xlLeft
    End With
    With r.Cells(3, 1)
        .Font.Size = 8
        .Font.Color = C_GRI()
        .IndentLevel = 1
        .HorizontalAlignment = xlLeft
    End With
End Sub

Public Sub KS_KartYaz(ByVal r As Range, ByVal deger As Variant, ByVal bicim As String, _
        ByVal alt As String, Optional ByVal renk As Long = -1)
    With r.Cells(2, 1)
        .NumberFormat = bicim
        If IsEmpty(deger) Then .Value = "-" Else .Value = deger
        If renk = -1 Then .Font.Color = C_LAC() Else .Font.Color = renk
    End With
    r.Cells(3, 1).Value = alt
End Sub

' Pencere ayarlarý sayfa aktif edilerek yapýlýr.
Public Sub KS_Pencere(ByVal ws As Worksheet, ByVal yakinlik As Long, ByVal dondurSatir As Long, _
        ByVal basliklar As Boolean)
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .DisplayGridlines = False
        .DisplayHeadings = basliklar
        .Zoom = yakinlik
        If dondurSatir > 0 Then
            .SplitColumn = 0
            .SplitRow = dondurSatir
            .FreezePanes = True
        End If
    End With
    ws.Range("A1").Select
End Sub

Public Sub KS_TabloStiliKur()
    Dim ts As TableStyle
    On Error Resume Next
    Set ts = ThisWorkbook.TableStyles(KS_STIL)
    On Error GoTo 0
    If ts Is Nothing Then Set ts = ThisWorkbook.TableStyles.Add(KS_STIL)
    On Error Resume Next
    With ts.TableStyleElements(xlHeaderRow)
        .Clear
        .Interior.Color = C_LAC()
        .Font.Color = vbWhite
        .Font.Bold = True
    End With
    With ts.TableStyleElements(xlRowStripe1)
        .Clear
        .Interior.Color = C_SERIT()
    End With
    With ts.TableStyleElements(xlWholeTable)
        .Clear
        .Borders(xlInsideHorizontal).Color = C_CIZGI()
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = C_CIZGI()
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    On Error GoTo 0
End Sub
