Attribute VB_Name = "modKS_TASARIM"
Option Explicit

' KALÝTESORA görsel dili (AKSORA ailesi): banner görseli, açýk zemin, gölgeli beyaz kartlar,
' tek tip düðmeler ve hizalý sekme çubuðu. Görseller dosyanýn yanýndaki KALITESORA_GORSEL klasöründen
' bir kez okunur ve dosyaya gömülür.

Public Const KS_FONT As String = "Segoe UI"
Public Const KS_FONT_KALIN As String = "Segoe UI Semibold"
Public Const KS_STIL As String = "KS_STIL"
Public Const KS_GENISLIK As Double = 1180     ' sayfa içerik geniþliði (pt)
Public Const KS_KLASOR As String = "KALITESORA_GORSEL"

Private mDurumAlindi As Boolean
Private mRibbon As Boolean
Private mFormulBar As Boolean
Private mBasliklar As Boolean
Private mSigdiriliyor As Boolean

' ---------------------------------------------------------------- ek renkler

Public Function C_MAVI() As Long
    C_MAVI = RGB(21, 101, 192)
End Function

Public Function C_YESIL2() As Long
    C_YESIL2 = RGB(11, 122, 94)
End Function

Public Function C_TEAL() As Long
    C_TEAL = RGB(31, 139, 160)
End Function

Public Function C_ACIKMAVI() As Long
    C_ACIKMAVI = RGB(232, 242, 252)
End Function

' ---------------------------------------------------------------- görseller

Public Function KS_GorselYol(ByVal ad As String) As String
    Dim p As String
    If Len(ThisWorkbook.Path) = 0 Then Exit Function
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then Exit Function
    p = ThisWorkbook.Path & Application.PathSeparator & KS_KLASOR & Application.PathSeparator & ad
    If Len(Dir$(p)) > 0 Then KS_GorselYol = p
End Function

Public Sub KS_ZeminUygula(ByVal ws As Worksheet, ByVal acik As Boolean)
    Dim p As String
    On Error Resume Next
    ws.SetBackgroundPicture Filename:=""
    If acik Then
        p = KS_GorselYol("KALITESORA_ZEMIN.png")
        If Len(p) > 0 Then ws.SetBackgroundPicture Filename:=p
    End If
    On Error GoTo 0
End Sub

' ---------------------------------------------------------------- sayfa çerçevesi

' Banner + sekme çubuðu. tam = True ise büyük banner (PANEL). Çerçevenin alt kenarýný (pt) döndürür.
Public Function KS_SayfaCerceve(ByVal ws As Worksheet, ByVal tam As Boolean) As Double
    Dim sol As Double, ust As Double, yuk As Double, p As String, s As Shape, alt As Double
    sol = ws.Cells(1, 2).Left
    If tam Then
        p = KS_GorselYol("KALITESORA_BANNER.png")
        yuk = KS_GENISLIK * 250 / 2400
    Else
        p = KS_GorselYol("KALITESORA_BANNER_INCE.png")
        yuk = KS_GENISLIK * 170 / 2400
    End If
    ' Görsel klasörü yoksa daha önce gömülmüþ banner korunur.
    Set s = Nothing
    If Len(p) = 0 Then
        On Error Resume Next
        Set s = ws.Shapes("ksC_Banner")
        On Error GoTo 0
        If Not s Is Nothing Then
            If s.Type <> msoPicture Then Set s = Nothing
        End If
    End If
    If s Is Nothing Then
        KS_SekilSil ws, "ksC_"
    Else
        s.Name = "ksBanner_Gecici"
        KS_SekilSil ws, "ksC_"
        s.Name = "ksC_Banner"
        s.Left = sol
        s.Top = 0
        s.Width = KS_GENISLIK
    End If
    If Not s Is Nothing Then
        ' korunan banner kullanýlýr
    ElseIf Len(p) > 0 Then
        Set s = ws.Shapes.AddPicture(p, msoFalse, msoTrue, sol, 0, KS_GENISLIK, yuk)
    Else
        Set s = ws.Shapes.AddShape(msoShapeRectangle, sol, 0, KS_GENISLIK, yuk)
        KS_SekilBicim s, KS_AD & "   |   KALÝTE · GARANTÝ · ARIZA TAKÝP MERKEZÝ", C_LAC(), vbWhite, 18
        s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        s.TextFrame2.MarginLeft = 24
    End If
    s.Name = "ksC_Banner"
    s.Placement = xlFreeFloating
    s.Locked = True
    alt = KS_SekmeCubugu(ws, sol, yuk + 6)
    KS_SayfaCerceve = alt
End Function

' Sekme çubuðu: 8 eþit sekme; aktif sekme lacivert, diðerleri beyaz.
Private Function KS_SekmeCubugu(ByVal ws As Worksheet, ByVal sol As Double, ByVal ust As Double) As Double
    Dim et As Variant, hedef As Variant, i As Long, gen As Double, s As Shape, aktif As Boolean, yuk As Double
    et = Array("PANEL", "KAYIT", "SEVKÝYATLAR", "ARIZALAR", "MALZEMELER", "KARNE", "RAPORLAR", "YÖNETÝM")
    hedef = Array(SH_PANEL, SH_KAYIT, SH_SEVK, SH_ARIZA, SH_MALZ, SH_KARNE, SH_RAPOR, SH_TANIM)
    yuk = 26
    gen = (KS_GENISLIK - 7 * 6) / 8
    For i = 0 To 7
        aktif = (StrComp(CStr(hedef(i)), ws.Name, vbTextCompare) = 0)
        If i = 7 Then
            aktif = (ws.Name = SH_TANIM Or ws.Name = SH_AKTAR Or ws.Name = SH_LOG)
        End If
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, sol + i * (gen + 6), ust, gen, yuk)
        s.Name = "ksC_Sekme" & i
        s.AlternativeText = CStr(hedef(i))
        If aktif Then
            KS_SekilBicim s, CStr(et(i)), C_LAC(), vbWhite, 9
        Else
            KS_SekilBicim s, CStr(et(i)), vbWhite, C_LAC(), 9
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = RGB(205, 216, 230)
            s.Line.Weight = 0.75
        End If
        s.Adjustments.Item(1) = 0.3
        s.OnAction = "KS_NAV"
    Next i
    KS_SekmeCubugu = ust + yuk
End Function

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

' Sayfa baþlýðý (cerceve altýndaki satýrda, sola yaslý) ve açýklamasý.
Public Sub KS_SayfaBaslik(ByVal ws As Worksheet, ByVal ust As Double, ByVal baslik As String, ByVal aciklama As String, _
        Optional ByVal aciklamaGen As Double = 620)
    Dim sol As Double
    sol = ws.Cells(1, 2).Left
    KS_Yazi ws, "ksC_Baslik", baslik, sol, ust, aciklamaGen, 24, 15, True, C_LAC(), msoAlignLeft
    KS_Yazi ws, "ksC_Aciklama", aciklama, sol, ust + 22, aciklamaGen, 16, 8, False, C_GRI(), msoAlignLeft
End Sub

' ---------------------------------------------------------------- þekil yardýmcýlarý

Public Function KS_Yazi(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, ByVal sol As Double, _
        ByVal ust As Double, ByVal gen As Double, ByVal yuk As Double, ByVal boyut As Single, ByVal kalin As Boolean, _
        ByVal renk As Long, ByVal hiza As Long) As Shape
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(ad).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, sol, ust, gen, yuk)
    s.Name = ad
    s.Fill.Visible = msoFalse
    s.Line.Visible = msoFalse
    s.Placement = xlFreeFloating
    With s.TextFrame2
        .MarginLeft = 0
        .MarginRight = 0
        .MarginTop = 0
        .MarginBottom = 0
        .WordWrap = msoTrue
        .VerticalAnchor = msoAnchorTop
        .AutoSize = msoAutoSizeNone
        With .TextRange
            .Text = metin
            .ParagraphFormat.Alignment = hiza
            .Font.Name = IIf(kalin, KS_FONT_KALIN, KS_FONT)
            .Font.Size = boyut
            .Font.Bold = IIf(kalin, msoTrue, msoFalse)
            .Font.Fill.ForeColor.RGB = renk
        End With
    End With
    Set KS_Yazi = s
End Function

Public Sub KS_YaziGuncelle(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, Optional ByVal renk As Long = -1)
    On Error Resume Next
    With ws.Shapes(ad).TextFrame2.TextRange
        .Text = metin
        If renk <> -1 Then .Font.Fill.ForeColor.RGB = renk
    End With
    On Error GoTo 0
End Sub

' Beyaz, köþeleri yumuþak, hafif gölgeli kart.
Public Function KS_KartSekil(ByVal ws As Worksheet, ByVal ad As String, ByVal sol As Double, ByVal ust As Double, _
        ByVal gen As Double, ByVal yuk As Double) As Shape
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(ad).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, sol, ust, gen, yuk)
    s.Name = ad
    s.Adjustments.Item(1) = 0.06
    s.Fill.Solid
    s.Fill.ForeColor.RGB = vbWhite
    s.Line.Visible = msoTrue
    s.Line.ForeColor.RGB = RGB(222, 231, 241)
    s.Line.Weight = 0.75
    On Error Resume Next
    With s.Shadow
        .Visible = msoTrue
        .ForeColor.RGB = RGB(20, 50, 90)
        .Transparency = 0.88
        .Blur = 10
        .OffsetX = 0
        .OffsetY = 3
    End With
    On Error GoTo 0
    s.Placement = xlFreeFloating
    s.TextFrame2.TextRange.Text = ""
    Set KS_KartSekil = s
End Function

Public Function KS_Serit(ByVal ws As Worksheet, ByVal ad As String, ByVal sol As Double, ByVal ust As Double, _
        ByVal gen As Double, ByVal yuk As Double, ByVal renk As Long) As Shape
    Dim s As Shape
    On Error Resume Next
    ws.Shapes(ad).Delete
    On Error GoTo 0
    Set s = ws.Shapes.AddShape(msoShapeRectangle, sol, ust, gen, yuk)
    s.Name = ad
    s.Fill.Solid
    s.Fill.ForeColor.RGB = renk
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    Set KS_Serit = s
End Function

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
        Case "ana": KS_SekilBicim s, metin, C_MAVI(), vbWhite, 9
        Case "yesil": KS_SekilBicim s, metin, C_YESIL2(), vbWhite, 9
        Case "teal": KS_SekilBicim s, metin, C_TEAL(), vbWhite, 9
        Case "koyu": KS_SekilBicim s, metin, C_LAC(), vbWhite, 9
        Case "turuncu": KS_SekilBicim s, metin, RGB(214, 120, 10), vbWhite, 9
        Case "cizgi"
            KS_SekilBicim s, metin, vbWhite, C_LAC(), 9
            s.Line.Visible = msoTrue
            s.Line.ForeColor.RGB = RGB(190, 204, 222)
            s.Line.Weight = 1
        Case Else: KS_SekilBicim s, metin, RGB(232, 237, 244), C_METIN(), 9
    End Select
    s.OnAction = makro
    Set KS_Buton = s
End Function

Public Sub KS_SekilBicim(ByVal s As Shape, ByVal metin As String, ByVal zemin As Long, _
        ByVal yazi As Long, ByVal boyut As Single)
    On Error Resume Next
    s.Adjustments.Item(1) = 0.2
    On Error GoTo 0
    s.Fill.Visible = msoTrue
    s.Fill.Solid
    s.Fill.ForeColor.RGB = zemin
    s.Line.Visible = msoFalse
    s.Shadow.Visible = msoFalse
    s.Placement = xlFreeFloating
    With s.TextFrame2
        .MarginLeft = 4
        .MarginRight = 4
        .MarginTop = 0
        .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .WordWrap = msoFalse
        With .TextRange
            .Text = metin
            .ParagraphFormat.Alignment = msoAlignCenter
            .Font.Name = KS_FONT_KALIN
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

' ---------------------------------------------------------------- hücre tabanlý biçimler

' Veri tablosu olmayan sayfayý sýfýrlar. Hücre dolgusu yok: arka plan görseli görünür.
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
    ws.Rows.Hidden = False
    ws.Columns.Hidden = False
    ws.Cells.Font.Name = KS_FONT
    ws.Cells.Font.Size = 10
    ws.Cells.Font.Color = C_METIN()
    ws.Cells.Interior.Pattern = xlNone
    ws.Cells.VerticalAlignment = xlCenter
    ws.Cells.Locked = True
    ws.Columns.ColumnWidth = 8.43
    ws.Columns(1).ColumnWidth = 2
    ws.Rows.RowHeight = 20
End Sub

' Bölüm baþlýðý: lacivert yarý kalýn metin, altýnda camgöbeði çizgi.
Public Sub KS_Bolum(ByVal r As Range, ByVal metin As String)
    With r
        .Interior.Pattern = xlNone
        .Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    With r.Cells(1, 1)
        .Value = metin
        .Font.Name = KS_FONT_KALIN
        .Font.Bold = True
        .Font.Size = 11
        .Font.Color = C_LAC()
        .VerticalAlignment = xlBottom
    End With
    r.EntireRow.RowHeight = 26
End Sub

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
        .Borders(xlInsideVertical).Color = RGB(40, 75, 115)
    End With
    r.EntireRow.RowHeight = 30
End Sub

Public Sub KS_TabloGovde(ByVal r As Range)
    Dim i As Long
    With r
        .Interior.Color = vbWhite
        .Font.Size = 9.5
        .Font.Color = C_METIN()
        .Borders(xlInsideHorizontal).Color = RGB(230, 236, 243)
        .Borders(xlEdgeBottom).Color = RGB(205, 216, 230)
        .Borders(xlEdgeLeft).Color = RGB(205, 216, 230)
        .Borders(xlEdgeRight).Color = RGB(205, 216, 230)
    End With
    For i = 2 To r.Rows.Count Step 2
        r.Rows(i).Interior.Color = RGB(247, 250, 253)
    Next i
End Sub

' KPI hücre kartý (karne sayfasýnda kullanýlýr): etiket / deðer / alt yazý.
Public Sub KS_KartCiz(ByVal r As Range, ByVal etiket As String)
    With r
        .Interior.Color = vbWhite
        .Borders(xlEdgeTop).Color = RGB(222, 231, 241)
        .Borders(xlEdgeBottom).Color = RGB(222, 231, 241)
        .Borders(xlEdgeRight).Color = RGB(222, 231, 241)
        .Borders(xlEdgeLeft).Color = C_MAVI()
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
        .Font.Name = KS_FONT_KALIN
        .Font.Size = 18
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

Public Sub KS_Pencere(ByVal ws As Worksheet, ByVal yakinlik As Long, ByVal dondurSatir As Long)
    ws.Activate
    With ActiveWindow
        .FreezePanes = False
        .ScrollRow = 1
        .ScrollColumn = 1
        .DisplayGridlines = False
        .DisplayHeadings = False
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
        .Interior.Color = RGB(247, 250, 253)
    End With
    With ts.TableStyleElements(xlWholeTable)
        .Clear
        .Borders(xlInsideHorizontal).Color = RGB(230, 236, 243)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlEdgeBottom).Color = RGB(205, 216, 230)
        .Borders(xlEdgeBottom).LineStyle = xlContinuous
    End With
    On Error GoTo 0
End Sub

' ================================================================ uygulama modu ve ekran sýðdýrma (AKSORA)


Public Function KS_UygulamaModuAcik() As Boolean
    KS_UygulamaModuAcik = True
    On Error Resume Next
    KS_UygulamaModuAcik = (ThisWorkbook.Names("_KS_UYGULAMA").RefersTo <> "=0")
    On Error GoTo 0
End Function

' Excel þeridi, formül çubuðu, satýr / sütun baþlýklarý ve kýlavuz çizgileri gizlenir.
Public Sub KS_UygulamaModuUygula()
    Dim ws As Worksheet, aktif As Object
    On Error Resume Next
    If Not KS_Kurulu() Then Exit Sub
    If Not KS_UygulamaModuAcik() Then Exit Sub
    If Not mDurumAlindi Then
        mRibbon = Application.CommandBars("Ribbon").Visible
        mFormulBar = Application.DisplayFormulaBar
        mBasliklar = ThisWorkbook.Windows(1).DisplayHeadings
        mDurumAlindi = True
    End If
    Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",False)"
    Application.DisplayFormulaBar = False
    Application.DisplayStatusBar = True
    With ThisWorkbook.Windows(1)
        .DisplayHeadings = False
        .DisplayGridlines = False
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
    End With
    On Error GoTo 0
End Sub

' Excel görünümü geri getirilir (dosya kapanýrken ya da baþka kitaba geçerken).
Public Sub KS_UYGULAMA_KAPAT()
    On Error Resume Next
    If mDurumAlindi Then
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon""," & IIf(mRibbon, "True", "False") & ")"
        Application.DisplayFormulaBar = mFormulBar
    Else
        Application.ExecuteExcel4Macro "SHOW.TOOLBAR(""Ribbon"",True)"
        Application.DisplayFormulaBar = True
    End If
    Application.DisplayStatusBar = True
    mDurumAlindi = False
    On Error GoTo 0
End Sub

Public Sub KS_UYGULAMA_MODU_DEGISTIR()
    Dim acik As Boolean
    acik = KS_UygulamaModuAcik()
    KS_AdTanimla "_KS_UYGULAMA", IIf(acik, "=0", "=1")
    ThisWorkbook.Names("_KS_UYGULAMA").Visible = False
    If acik Then
        KS_UYGULAMA_KAPAT
        On Error Resume Next
        ThisWorkbook.Windows(1).DisplayHeadings = True
        On Error GoTo 0
        MsgBox "Excel görünümüne geçildi. Uygulama moduna dönmek için ayný düðmeye basýn.", vbInformation, KS_AD
    Else
        KS_UygulamaModuUygula
        KS_EkranSigdir ActiveSheet
    End If
End Sub

' Sayfanýn içerik geniþliði (banner) pencereye tam oturacak þekilde yakýnlaþtýrýr.
Public Sub KS_EkranSigdir(ByVal Sh As Object, Optional ByVal basaDon As Boolean = True)
    Dim ws As Worksheet, wn As Window, hedef As Double, c As Long, olay As Boolean, ekran As Boolean
    Dim secim As Boolean, z As Long
    If mSigdiriliyor Then Exit Sub
    If Not TypeOf Sh Is Worksheet Then Exit Sub
    Set ws = Sh
    If Not KS_BizimSayfa(ws.Name) Then Exit Sub
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    On Error GoTo Son
    If Not Application.Visible Then Exit Sub
    If Application.WindowState = xlMinimized Then Exit Sub
    Set wn = ActiveWindow
    If wn.WindowState = xlMinimized Then Exit Sub
    mSigdiriliyor = True
    olay = Application.EnableEvents
    ekran = Application.ScreenUpdating
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    hedef = ws.Cells(1, 2).Left + KS_GENISLIK + 4
    c = 2
    Do While ws.Columns(c).Left + ws.Columns(c).Width < hedef And c < 200
        c = c + 1
    Loop
    secim = (ws.EnableSelection <> xlNoRestrictions)
    If secim Then ws.EnableSelection = xlNoRestrictions
    wn.ScrollRow = 1
    wn.ScrollColumn = 1
    ws.Range(ws.Cells(1, 1), ws.Cells(1, c)).Select
    wn.Zoom = True
    z = CLng(wn.Zoom)
    If z > 12 Then wn.Zoom = z - 1
    ws.Cells(1, 1).Select
    If basaDon Then
        wn.ScrollRow = 1
        wn.ScrollColumn = 1
    End If
Son:
    On Error Resume Next
    Application.EnableEvents = olay
    Application.ScreenUpdating = ekran
    mSigdiriliyor = False
    On Error GoTo 0
End Sub

Public Sub KS_EKRANA_SIGDIR()
    KS_EkranSigdir ActiveSheet
End Sub

' Aktif sayfayý (karne / rapor) yatay, tek sayfa geniþliðinde PDF olarak kaydeder ve açar.
Public Sub KS_PDF_KAYDET()
    Dim ws As Worksheet, klasor As String, dosya As String, son As Long, sonKol As Long, korumali As Boolean, h As Range
    On Error GoTo Hata
    Set ws = ActiveSheet
    If Len(ThisWorkbook.Path) = 0 Or LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        MsgBox "PDF için dosyayý bilgisayardaki bir klasörden açýn.", vbExclamation, KS_AD
        Exit Sub
    End If
    klasor = ThisWorkbook.Path & Application.PathSeparator & "KALITESORA_RAPOR"
    If Len(Dir$(klasor, vbDirectory)) = 0 Then MkDir klasor
    dosya = klasor & Application.PathSeparator & Replace(ws.Name, " ", "_") & "_" & Format$(Now, "yyyymmdd_hhnn") & ".pdf"
    Set h = ws.Cells.Find(What:="*", LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If h Is Nothing Then son = 40 Else son = h.Row + 2
    sonKol = 2
    Do While ws.Columns(sonKol).Left + ws.Columns(sonKol).Width < ws.Cells(1, 2).Left + KS_GENISLIK And sonKol < 60
        sonKol = sonKol + 1
    Loop
    korumali = ws.ProtectContents
    KS_Ac ws
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(son, sonKol)).Address
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .LeftMargin = Application.CentimetersToPoints(0.8)
        .RightMargin = Application.CentimetersToPoints(0.8)
        .TopMargin = Application.CentimetersToPoints(0.8)
        .BottomMargin = Application.CentimetersToPoints(1)
        .CenterFooter = KS_AD & "  ·  " & ws.Name & "  ·  &D  ·  Sayfa &P / &N"
    End With
    Application.PrintCommunication = True
    If korumali Then KS_Kapat ws
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dosya, Quality:=xlQualityStandard, _
        IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=True
    KS_Log "PDF", ws.Name, dosya
    Exit Sub
Hata:
    Dim mesaj As String
    mesaj = Err.Description
    On Error Resume Next
    Application.PrintCommunication = True
    If korumali Then KS_Kapat ws
    MsgBox "PDF oluþturulamadý: " & mesaj, vbExclamation, KS_AD
End Sub
