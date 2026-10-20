Attribute VB_Name = "modAKSORA_MUSTERI_TEMA"
Option Explicit
Private Const etiket As String = "AKSORA_MUSTERI_TEMA_V1"
Private mTemaMesgul As Boolean
Private mAsama As String
Public AKSORA_MUSTERI_TOPLU_AKTIF As Boolean
Private mMusteriTopluBas As Double
Private mMusteriTopluMetin As String
Private mMusteriTopluSaniye As Long
Private mSonGorunen As String
Private mTopluMesaj As String
Private mTopluIkon As Long

Private Function TEMA_KAYNAKLAR() As Variant
    TEMA_KAYNAKLAR = Array("ANA PANEL", "MALZEME_ARA", "CDP_STOK", "CDP_RAPOR", _
        "ÝLÇE_STOK", "ÝLÇE_RAPORLAR", "YÜKLENÝCÝ_STOK", "YÜKLENÝCÝ_RAPOR")
End Function
Private Function TEMA_ADLAR() As Variant
    TEMA_ADLAR = Array("M_ANA PANEL", "M_MALZEME_ARA", "M_CDP_STOK", "M_CDP_RAPOR", _
        "M_ÝLÇE_STOK", "M_ÝLÇE_RAPOR", "M_YÜKLENÝCÝ_STOK", "M_YÜKLENÝCÝ_RAPOR")
End Function

Public Function AKSORA_MUSTERI_TEMA_SAYFA(ByVal sh As Object) As Boolean
    Dim adlar As Variant, ad As Variant
    If Not TypeOf sh Is Worksheet Then Exit Function
    adlar = TEMA_ADLAR()
    For Each ad In adlar
        If sh.Name = CStr(ad) Then
            AKSORA_MUSTERI_TEMA_SAYFA = (CStr(sh.Range("BX1").Value2) = etiket)
            Exit Function
        End If
    Next ad
End Function

Public Sub AKSORA_MUSTERI_TEMA_AC()
    Dim kaynaklar As Variant, adlar As Variant, i As Long, ws As Worksheet, kaynak As Worksheet
    Dim ev As Boolean, sc As Boolean, Hata As String, k As Worksheet, secili As Object
    If mTemaMesgul Then Exit Sub
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    mTemaMesgul = True: mAsama = "Kaynak kontrolu"
    ev = application.EnableEvents: sc = application.ScreenUpdating
    Set secili = ActiveSheet
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    kaynaklar = TEMA_KAYNAKLAR(): adlar = TEMA_ADLAR()
    ' Tum kaynaklar ve isim cakismalari yazmadan once kontrol edilir.
    For i = 0 To 7
        Set kaynak = ThisWorkbook.Worksheets(CStr(kaynaklar(i)))
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i)))
        On Error GoTo Hata
        If Not ws Is Nothing Then
            If Not AKSORA_MUSTERI_TEMA_SAYFA(ws) Then _
                Err.Raise vbObjectError + 1700, , "Sayfa adi kullaniliyor: " & ws.Name
        End If
    Next i
    For i = 0 To 7
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i)))
        On Error GoTo Hata
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(adlar(i))
            ws.Range("BX1").Value2 = etiket
        End If
        ws.Visible = xlSheetVisible
        If Val(CStr(ws.Range("BX2").Value2)) < 9 Then
            Set kaynak = ThisWorkbook.Worksheets(CStr(kaynaklar(i)))
            TEMA_KOPYALA kaynak, ws, i
        End If
        mAsama = ws.Name & " / il basligi"
        ws.Unprotect ""
        If i = 1 Then
            TEMA_MUSTERI_ARAMA_BASLIKLAR ws
            TEMA_MUSTERI_ARAMA_V15 ws
        End If
        If i = 0 Then
            TEMA_MUSTERI_ANA_DUZEN ws
            TEMA_MUSTERI_ANA_V12 ws
            TEMA_MUSTERI_ANA_V13 ws
        End If
        If i = 2 Or i = 4 Or i = 6 Then ws.Range("A5").MergeArea.Cells(1, 1).Value2 = TEMA_IL_ADI() & " Ýl Müdürlüðü"
        AKSORA_SAYI_BICIM_SAYFA ws
        AKSORA_MUSTERI_SAYFA_HAZIRLA ws
        ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, _
            AllowFiltering:=(ws.Name = "M_MALZEME_ARA" Or ws.Name = "M_CDP_STOK" Or _
                ws.Name = "M_ÝLÇE_STOK" Or ws.Name = "M_YÜKLENÝCÝ_STOK")
        If ws.Name = "M_MALZEME_ARA" Then ws.EnableSelection = xlNoRestrictions
        ws.Visible = xlSheetVisible
    Next i
    mAsama = "Ana paneli gösterme"
    ThisWorkbook.Worksheets(CStr(adlar(0))).Select Replace:=True
    For Each k In ThisWorkbook.Worksheets
        If Not AKSORA_MUSTERI_TEMA_SAYFA(k) Then
            If k.Visible = xlSheetVisible Then k.Visible = xlSheetVeryHidden
        End If
    Next k
    mAsama = "Ana panel görünümü"
    AKSORA_MUSTERI_TEMA_GORUNUM ActiveSheet
Cikis:
    application.CutCopyMode = False
    mTemaMesgul = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If Len(Hata) > 0 Then MsgBox Hata, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    Hata = "Müþteri tema kurulumu / " & mAsama & vbCrLf & _
        "Hata " & CStr(Err.Number) & ": " & Err.Description
    On Error Resume Next
    secili.Select Replace:=True
    For Each k In ThisWorkbook.Worksheets
        If AKSORA_MUSTERI_TEMA_SAYFA(k) And Not (k Is ActiveSheet) Then k.Visible = xlSheetVeryHidden
    Next k
    On Error GoTo 0
    Resume Cikis
End Sub

Private Sub TEMA_KOPYALA(ByVal kaynak As Worksheet, ByVal ws As Worksheet, ByVal ix As Long)
    Dim r As Range, c As Range, s As Shape, yeni As Shape, i As Long
    Dim sonSatir As Long, sonSutun As Long, baslikSon As Long
    mAsama = ws.Name & " / biçim hazýrlýðý"
    ws.Unprotect ""
    ws.EnableSelection = xlNoRestrictions
    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ' Bos tema icin ilk 200 satir yeterlidir; eski genis bicim alanlari temizlenir.
    ws.UsedRange.UnMerge
    ws.UsedRange.Clear
    ws.Range("BX1").Value2 = etiket
    Set r = Intersect(kaynak.UsedRange, kaynak.Rows("1:200"))
    mAsama = ws.Name & " / hücre biçimleri"
    r.Copy
    ws.Range(r.Address).PasteSpecial Paste:=xlPasteFormats
    r.Copy
    ws.Range(r.Address).PasteSpecial Paste:=xlPasteColumnWidths
    application.CutCopyMode = False
    sonSatir = r.Row + r.Rows.Count - 1
    sonSutun = r.Column + r.Columns.Count - 1
    For i = 1 To sonSatir
        ws.Rows(i).RowHeight = kaynak.Rows(i).RowHeight
        ' Veri filtrelerinden gelen gizli satirlar temaya tasinmaz.
        ws.Rows(i).Hidden = False
    Next i
    For i = 1 To sonSutun
        ws.Columns(i).Hidden = kaynak.Columns(i).Hidden
    Next i
    ' Yalniz sabit baslik metinleri; stok kayitlari, tutarlar ve formuller alinmaz.
    Select Case ix
        Case 0: baslikSon = 80
        Case 1: baslikSon = 13
        Case 2, 4, 6: baslikSon = 6
        Case Else: baslikSon = 8
    End Select
    mAsama = ws.Name & " / sabit baþlýklar"
    For Each c In kaynak.Range(kaynak.Cells(1, 1), kaynak.Cells(baslikSon, sonSutun))
        If Not c.HasFormula And VarType(c.Value2) = vbString Then
            If Len(CStr(c.Value2)) > 0 And c.Column < 26 Then
                If Not TEMA_DEGISKEN_METIN(CStr(c.Value2)) Then
                    ws.Range(c.Address).MergeArea.Cells(1, 1).Value2 = c.Value2
                End If
            End If
        End If
    Next c
    ' Stok tarihlerinin ve secili ilin kaynaktaki degerleri aktarilmaz.
    If ix = 2 Or ix = 4 Or ix = 6 Then
        TEMA_BOS_DEGER ws, "C2"
        TEMA_BOS_DEGER ws, "C3"
        TEMA_BOS_DEGER ws, "C4"
        ws.Range("A5").MergeArea.Cells(1, 1).Value2 = TEMA_IL_ADI() & " Ýl Müdürlüðü"
    ElseIf ix = 3 Or ix = 5 Or ix = 7 Then
        TEMA_BOS_DEGER ws, "B3"
        TEMA_BOS_DEGER ws, "B4"
        TEMA_BOS_DEGER ws, "B5"
    End If
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    For Each s In kaynak.Shapes
        If s.Visible <> msoFalse Then
        mAsama = ws.Name & " / görsel: " & s.Name
        Set yeni = TEMA_SEKIL_KOPYALA(kaynak, ws, s)
        mAsama = ws.Name & " / görsel adý: " & s.Name
        yeni.Name = s.Name
        ' Yerel cizilen denetimin olculeri zaten kurulumda verilmistir.
        If s.Type <> msoFormControl And LCase$(Left$(s.Name, 9)) <> "drop down" Then
            mAsama = ws.Name & " / en-boy kilidi: " & s.Name
            yeni.LockAspectRatio = msoFalse
            mAsama = ws.Name & " / konum: " & s.Name
            yeni.Left = s.Left: yeni.Top = s.Top
            mAsama = ws.Name & " / boyut: " & s.Name
            yeni.Width = s.Width: yeni.Height = s.Height
        End If
        mAsama = ws.Name & " / yeni görsel yerleþimi: " & s.Name
        ' Kaynak Form denetiminin Placement ozelligi okunmaz.
        yeni.Placement = xlFreeFloating
        yeni.Visible = s.Visible
        mAsama = ws.Name & " / görsel baðlantýsý: " & s.Name
        TEMA_EYLEM_SIL yeni
        If yeni.Type = msoChart Then
            Do While yeni.Chart.SeriesCollection.Count > 0
                yeni.Chart.SeriesCollection(1).Delete
            Loop
        End If
        End If
    Next s
    mAsama = ws.Name & " / son düzenleme"
    ws.Hyperlinks.Delete
    On Error Resume Next
    ws.Cells.Validation.Delete
    On Error GoTo 0
    ' Yalniz yeni sayfalarda birim secimine donus baglantisi eklenir.
    Set yeni = ws.Shapes.AddShape(msoShapeRoundedRectangle, 10, 740, 192, 29)
    yeni.Name = "MUSTERI_TEMA_GERI"
    yeni.Fill.ForeColor.RGB = RGB(24, 76, 104)
    yeni.line.ForeColor.RGB = RGB(74, 186, 201)
    yeni.TextFrame2.TextRange.Text = ChrW(8592) & " Birim seçimine dön"
    yeni.TextFrame2.TextRange.Font.Size = 10
    yeni.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(235, 250, 255)
    yeni.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_BIRIM_ANA_MENU"
    ws.Range("BX1").Value2 = etiket: ws.Range("BX2").Value2 = 9
    ws.Columns("BX").Hidden = True
    ws.Tab.Color = kaynak.Tab.Color
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
End Sub

Private Sub TEMA_MUSTERI_ANA_DUZEN(ByVal ws As Worksheet)
    Dim yol As String, kod As String, il As String, bolum As Long, h As Double
    Dim r As Range, s As Shape, z As Shape, c As Range, i As Long, ad As String
    Dim sag As Double, cap As Double, sinirlar As Variant
    yol = AKSORA_GORSEL("AKSORA_MUSTERI_PANEL_BANTLARI.png")
    If Not TEMA_DOSYA_VAR(yol) Then Err.Raise vbObjectError + 1720, , _
        "AKSORA_MUSTERI_PANEL_BANTLARI.png dosyasýný Excel'in yanýna koyun."
    kod = CStr(ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2)
    il = TEMA_IL_ADI()
    Select Case kod
        Case "GIRESUN": bolum = 0
        Case "TRABZON": bolum = 1
        Case "RIZE": bolum = 2
        Case "GUMUSHANE": bolum = 3
        Case "ARTVIN": bolum = 4
        Case Else: Err.Raise vbObjectError + 1721, , "Önce il seçiniz."
    End Select
    If CStr(ws.Range("BX4").Value2) = "MUSTERI_ANA_BANT_V12|" & kod Then Exit Sub
    mAsama = "M_ANA PANEL / müþteri il bannerý"
    For i = ws.Shapes.Count To 1 Step -1
        ad = UCase$(ws.Shapes(i).Name)
        If InStr(ad, "PANEL_TEK_") > 0 Or Left$(ad, 17) = "AKSORA_IL_BANNER_" Or _
           ad = "AKSORA15_BANNER" Or ad = "MUSTERI_ANA_BANNER" Then ws.Shapes(i).Delete
    Next i
    Set r = ws.Range("B2:AB5")
    Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
    h = s.Height
    ' Gercek PNG satir sinirlari: 2120x742; komsu il bandi tasmaz.
    sinirlar = Array(0, 149, 297, 448, 597, 742)
    s.PictureFormat.CropTop = h * CDbl(sinirlar(bolum)) / 742
    s.PictureFormat.CropBottom = h * (742 - CDbl(sinirlar(bolum + 1))) / 742
    s.LockAspectRatio = msoFalse
    s.Left = r.Left: s.Top = r.Top: s.Width = r.Width: s.Height = r.Height
    s.Name = "MUSTERI_ANA_BANNER": s.Placement = xlFreeFloating
    s.ZOrder msoBringToFront
    mAsama = "M_ANA PANEL / alt bant hizalama"
    ' Eski hucre footer metni sekil footer'inin ustunden tasmaz.
    For Each c In ws.Range("B28:AB31").Cells
        c.MergeArea.ClearContents
    Next c
    Set r = ws.Range("B28:AB29")
    Set z = TEMA_ANA_SEKIL(ws, "AKSORA15_ALT_ZEMIN")
    If z Is Nothing Then
        Set z = ws.Shapes.AddShape(msoShapeRectangle, r.Left, r.Top, r.Width, r.Height)
        z.Name = "AKSORA15_ALT_ZEMIN"
    End If
    z.Left = r.Left: z.Top = r.Top: z.Width = r.Width: z.Height = r.Height
    z.Fill.ForeColor.RGB = RGB(13, 47, 79): z.line.Visible = msoFalse
    z.ZOrder msoBringToFront
    TEMA_ANA_YAZI ws, "AKSORA15_ALT_KURUM", _
        "ÇORUH ELEKTRÝK DAÐITIM A.Þ. | " & il & " Ýl Müdürlüðü", _
        r.Left, r.Top + 1, r.Width, r.Height / 2 - 1, 9, RGB(255, 255, 255), True
    TEMA_ANA_YAZI ws, "AKSORA15_ALT_IMZA", _
        "Burhan Serdar | Elektrik-Elektronik Mühendisi", _
        r.Left, r.Top + r.Height / 2, r.Width, r.Height / 2 - 1, 8, RGB(220, 238, 248), False
    Set r = ws.Range("B30:AB30")
    Set z = TEMA_ANA_SEKIL(ws, "AKSORA15_TARIH_ZEMIN")
    If Not z Is Nothing Then
        z.Left = r.Left: z.Top = r.Top: z.Width = r.Width: z.Height = r.Height
        z.ZOrder msoBringToFront
    End If
    TEMA_ANA_YAZI ws, "AKSORA15_PANEL_TARIHI", _
        "Müþteri Operasyonlarý Yönetimi | Görsel tasarým aþamasý", _
        r.Left, r.Top, r.Width, r.Height, 8, RGB(43, 88, 115), False
    mAsama = "M_ANA PANEL / kart etiketleri"
    For Each s In ws.Shapes
        ad = s.Name
        If Left$(ad, Len("AKSORA15_KART_DEGER_")) = "AKSORA15_KART_DEGER_" Then
            Select Case Right$(ad, 1)
                Case "0": s.TextFrame2.TextRange.Text = "Hazýrlanýyor"
                Case "4": s.TextFrame2.TextRange.Text = "Atanmadý"
                Case Else: s.TextFrame2.TextRange.Text = "-"
            End Select
            s.TextFrame2.TextRange.Font.Size = 8
            s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(70, 100, 120)
        End If
        If Left$(ad, Len("AKSORA15_KART_LED_")) = "AKSORA15_KART_LED_" Then
            sag = s.Left + s.Width: cap = s.Height
            s.LockAspectRatio = msoFalse: s.Width = cap: s.Left = sag - cap
            s.Fill.ForeColor.RGB = RGB(145, 162, 176)
        End If
        If Left$(ad, Len("AKSORA15_KART_ETIKET_")) = "AKSORA15_KART_ETIKET_" Then
            Select Case Right$(ad, 1)
                Case "1": s.TextFrame2.TextRange.Text = "SON GÜNCELLEME"
                Case "3": s.TextFrame2.TextRange.Text = "STOK DEÐERÝ"
                Case "4"
                    If InStr(ad, "_B4") > 0 Or InStr(ad, "_G4") > 0 Then _
                        s.TextFrame2.TextRange.Text = "SAP BAÐLANTISI"
            End Select
        End If
    Next s
    Set s = TEMA_ANA_SEKIL(ws, "AKSORA15_KART_BASLIK_L")
    If Not s Is Nothing Then s.TextFrame2.TextRange.Text = "ÝLÇELER"
    Set s = TEMA_ANA_SEKIL(ws, "AKSORA15_KART_BASLIK_Q")
    If Not s Is Nothing Then s.TextFrame2.TextRange.Text = "YÜKLENÝCÝ"
    ' Il secimi ve alt kurum bilgisi stok biriminden bagimsizdir.
    Set s = TEMA_ANA_SEKIL(ws, "AKSORA15_IL_DEGISTIR")
    If Not s Is Nothing Then
        s.TextFrame2.TextRange.Text = ChrW(8592) & " ÝL DEÐÝÞTÝR"
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_BIRIM_MUSTERI"
    End If
    Set s = TEMA_ANA_SEKIL(ws, "MUSTERI_TEMA_GERI")
    If Not s Is Nothing Then
        Set r = ws.Range("B6:E6")
        s.Left = r.Left: s.Top = r.Top + 1: s.Width = r.Width: s.Height = r.Height - 2
        s.TextFrame2.TextRange.Font.Size = 8
        s.TextFrame2.MarginLeft = 4: s.TextFrame2.MarginRight = 4
        s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End If
    ws.Range("BX4").Value2 = "MUSTERI_ANA_BANT_V12|" & kod
End Sub

Private Function TEMA_ANA_SEKIL(ByVal ws As Worksheet, ByVal ad As String) As Shape
    On Error Resume Next
    Set TEMA_ANA_SEKIL = ws.Shapes(ad)
    On Error GoTo 0
End Function

Private Sub TEMA_ANA_YAZI(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
    ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
    ByVal punto As Double, ByVal renk As Long, ByVal kalin As Boolean)
    Dim s As Shape
    Set s = TEMA_ANA_SEKIL(ws, ad)
    If s Is Nothing Then
        Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
        s.Name = ad
    End If
    s.Left = x: s.Top = y: s.Width = w: s.Height = h
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
        .TextRange.Text = metin
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = punto
        .TextRange.Font.bold = kalin
        .TextRange.Font.Fill.ForeColor.RGB = renk
    End With
    s.ZOrder msoBringToFront
End Sub

Private Sub TEMA_MUSTERI_ANA_V12(ByVal ws As Worksheet)
    Dim s As Shape, govde As Shape, r As Range, c As Range
    Dim i As Long, j As Long, ad As String, etiket As String, kod As String
    Dim kolonlar As Variant, dugmeler As Variant, basliklar As Variant, hedefler As Variant
    Dim eskiX As Double, eskiW As Double, yeniX As Double, yeniW As Double, bosluk As Double
    Dim goreli As Double, cap As Double
    kod = CStr(ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2)
    If CStr(ws.Range("BX5").Value2) = "MUSTERI_ANA_V12|" & kod Then Exit Sub
    mAsama = "M_ANA PANEL / gereksiz öðeleri kaldýrma"
    For i = ws.Shapes.Count To 1 Step -1
        ad = ws.Shapes(i).Name
        etiket = Mid$(ad, InStrRev(ad, "_") + 1)
        If ad = "MUSTERI_TEMA_GERI" Or LCase$(ad) = "btnepp" Or _
           ad = "AKSORA15_ALT_IMZA" Or ad = "AKSORA15_PANEL_TARIHI" Or _
           ad = "AKSORA15_TARIH_ZEMIN" Or Left$(LCase$(ad), 2) = "go" Or _
           Left$(ad, 14) = "MUSTERI_HIZLI_" Or _
           (Left$(ad, 14) = "AKSORA15_KART_" And Left$(etiket, 1) = "G") Then
            ws.Shapes(i).Delete
        End If
    Next i
    ' B/G/L/Q eski dordlu kart duzeni yerine mevcut uc stok grubu.
    mAsama = "M_ANA PANEL / üç kartý hizalama"
    Set r = ws.Range("B7:T13")
    kolonlar = Array("B", "L", "Q")
    bosluk = 16: yeniW = (r.Width - 2 * bosluk) / 3
    For i = 0 To 2
        Set govde = TEMA_ANA_SEKIL(ws, "AKSORA15_KART_GOVDE_" & CStr(kolonlar(i)))
        If Not govde Is Nothing Then
            eskiX = govde.Left: eskiW = govde.Width
            yeniX = r.Left + i * (yeniW + bosluk)
            For Each s In ws.Shapes
                ad = s.Name: etiket = Mid$(ad, InStrRev(ad, "_") + 1)
                If Left$(ad, 14) = "AKSORA15_KART_" And Left$(etiket, 1) = CStr(kolonlar(i)) Then
                    goreli = (s.Left - eskiX) / eskiW
                    If InStr(ad, "KART_LED_") > 0 Then
                        cap = s.Height
                        s.Left = yeniX + yeniW - 24: s.Width = cap
                    Else
                        s.Left = yeniX + goreli * yeniW
                        s.Width = s.Width / eskiW * yeniW
                    End If
                End If
            Next s
        End If
    Next i
    Set r = ws.Range("B20:T21")
    dugmeler = Array("btnCDP", "btnIlce", "btnYuk")
    basliklar = Array("CDP GÜNCELLE", "ÝLÇE STOK GÜNCELLE", "YÜKLENÝCÝ STOK GÜNCELLE")
    For i = 0 To 2
        Set s = TEMA_ANA_SEKIL(ws, CStr(dugmeler(i)))
        If Not s Is Nothing Then
            s.Left = r.Left + i * (yeniW + bosluk): s.Width = yeniW
            s.Top = r.Top + 1: s.Height = r.Height - 2
            s.TextFrame2.TextRange.Text = CStr(basliklar(i))
            s.TextFrame2.TextRange.Font.Size = 10
            s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End If
    Next i
    mAsama = "M_ANA PANEL / hýzlý eriþim"
    basliklar = Array("MALZEME ARA", "CDP STOK", "CDP RAPOR", "ÝLÇE STOK", _
        "ÝLÇE RAPOR", "YÜKLENÝCÝ STOK", "YÜKLENÝCÝ RAPOR")
    hedefler = Array("M_MALZEME_ARA", "M_CDP_STOK", "M_CDP_RAPOR", "M_ÝLÇE_STOK", _
        "M_ÝLÇE_RAPOR", "M_YÜKLENÝCÝ_STOK", "M_YÜKLENÝCÝ_RAPOR")
    Set r = ws.Range("B25:T26")
    bosluk = 8: yeniW = (r.Width - 6 * bosluk) / 7
    For i = 0 To 6
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
            r.Left + i * (yeniW + bosluk), r.Top + 1, yeniW, r.Height - 2)
        s.Name = "MUSTERI_HIZLI_" & CStr(i)
        s.Adjustments.Item(1) = 0.12
        s.Fill.ForeColor.RGB = RGB(244, 251, 255)
        s.line.ForeColor.RGB = RGB(112, 195, 228): s.line.Weight = 0.7
        s.Placement = xlFreeFloating
        s.AlternativeText = CStr(hedefler(i))
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_HIZLI_AC"
        With s.TextFrame2
            .MarginLeft = 3: .MarginRight = 3: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
            .TextRange.Text = CStr(basliklar(i))
            .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 8
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(15, 68, 106)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    Next i
    mAsama = "M_ANA PANEL / alt sýnýr"
    Set r = ws.Range("B28:AB29")
    TEMA_ANA_YAZI ws, "AKSORA15_ALT_KURUM", _
        "ÇORUH ELEKTRÝK DAÐITIM A.Þ. | " & TEMA_IL_ADI() & " Ýl Müdürlüðü", _
        r.Left, r.Top, r.Width, r.Height, 9, RGB(255, 255, 255), True
    For Each c In ws.Range("B28:AB31").Cells
        c.MergeArea.ClearContents
    Next c
    ' Sadece Musteri ana panelinin lacivert bant sonrasi satirlari gizlenir.
    ws.Rows("30:" & ws.Rows.Count).Hidden = True
    ws.Columns("AC:XFD").Hidden = True
    ws.ScrollArea = "A1:AB29"
    ws.Range("BX5").Value2 = "MUSTERI_ANA_V12|" & kod
End Sub

Public Sub AKSORA_MUSTERI_HIZLI_AC()
    Dim ws As Worksheet, hedef As Worksheet, ad As String
    If mTemaMesgul Then Exit Sub
    On Error GoTo Hata
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name <> "M_ANA PANEL" Then Exit Sub
    Set ws = ThisWorkbook.Worksheets("M_ANA PANEL")
    ad = ws.Shapes(CStr(application.Caller)).AlternativeText
    Set hedef = ThisWorkbook.Worksheets(ad)
    If Not AKSORA_MUSTERI_TEMA_SAYFA(hedef) Then Err.Raise vbObjectError + 1730, , "Müþteri sayfasý bulunamadý."
    application.ScreenUpdating = False
    application.EnableEvents = False
    hedef.Visible = xlSheetVisible
    hedef.Select Replace:=True
    application.EnableEvents = True
    AKSORA_MUSTERI_TEMA_GORUNUM hedef
    application.ScreenUpdating = True
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub TEMA_MUSTERI_ANA_V13(ByVal ws As Worksheet)
    Dim banner As Shape, foto As Shape, alt As Shape, yazi As Shape, dugme As Shape
    Dim serit As Shape, s As Shape, govde As Shape, r As Range
    Dim x As Double, w As Double, sag As Double, aralik As Double, bw As Double, dx As Double
    Dim i As Long, ad As String, etiket As String, kod As String
    Dim basliklar As Variant, hedefler As Variant, arama As Variant
    kod = CStr(ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2)
    If CStr(ws.Range("BX6").Value2) = "MUSTERI_ANA_V13|" & kod Then Exit Sub
    mAsama = "M_ANA PANEL / ortak sol ve sað sýnýr"
    Set banner = TEMA_ANA_SEKIL(ws, "MUSTERI_ANA_BANNER")
    If banner Is Nothing Then Err.Raise vbObjectError + 1731, , "Müþteri bannerý bulunamadý."
    x = banner.Left: w = banner.Width: sag = x + w
    Set alt = TEMA_ANA_SEKIL(ws, "AKSORA15_ALT_ZEMIN")
    If alt Is Nothing Then Err.Raise vbObjectError + 1732, , "Alt bant bulunamadý."
    alt.Left = x: alt.Width = w
    alt.line.Visible = msoFalse
    Set foto = TEMA_ANA_SEKIL(ws, "AKSORA15_TEK_ARKA_PLAN")
    If Not foto Is Nothing Then
        foto.LockAspectRatio = msoFalse
        foto.Left = x: foto.Width = w
        foto.Height = alt.Top - foto.Top
        foto.line.Visible = msoFalse
    End If
    Set s = TEMA_ANA_SEKIL(ws, "AKSORA15_UST_ZEMIN")
    If Not s Is Nothing Then
        s.Left = x: s.Width = w
        s.Top = banner.Top: s.Height = banner.Height
    End If
    Set dugme = TEMA_ANA_SEKIL(ws, "AKSORA15_IL_DEGISTIR")
    If Not dugme Is Nothing Then
        dugme.Left = sag - dugme.Width
        dugme.Top = banner.Top + banner.Height
        If Not foto Is Nothing Then dugme.Height = foto.Top - dugme.Top
        dugme.line.Visible = msoFalse
    End If
    ' Hucre zeminine bagli farkli kenarlar yerine ayni sinirli tek durum seridi.
    Set serit = TEMA_ANA_SEKIL(ws, "MUSTERI_ANA_SECIM_SERIDI")
    If serit Is Nothing And Not foto Is Nothing Then
        Set serit = ws.Shapes.AddShape(msoShapeRectangle, x, banner.Top + banner.Height, _
            w, foto.Top - banner.Top - banner.Height)
        serit.Name = "MUSTERI_ANA_SECIM_SERIDI"
    End If
    If Not serit Is Nothing Then
        serit.Left = x: serit.Width = w
        serit.Top = banner.Top + banner.Height
        If Not foto Is Nothing Then serit.Height = foto.Top - serit.Top
        serit.Fill.ForeColor.RGB = RGB(216, 243, 234)
        serit.line.Visible = msoFalse
        serit.ZOrder msoSendToBack
    End If
    Set yazi = TEMA_ANA_SEKIL(ws, "AKSORA15_ALT_KURUM")
    If Not yazi Is Nothing Then
        yazi.Left = x: yazi.Width = w
        yazi.Top = alt.Top: yazi.Height = alt.Height
        With yazi.TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        yazi.ZOrder msoBringToFront
    End If
    ' Ilk kart ve soldaki islemler ayni ic kenarda kalir.
    Set govde = TEMA_ANA_SEKIL(ws, "AKSORA15_KART_GOVDE_B")
    If Not govde Is Nothing Then
        dx = x + 1 - govde.Left
        For Each s In ws.Shapes
            ad = s.Name: etiket = Mid$(ad, InStrRev(ad, "_") + 1)
            If Left$(ad, 14) = "AKSORA15_KART_" And Left$(etiket, 1) = "B" Then s.Left = s.Left + dx
        Next s
    End If
    Set s = TEMA_ANA_SEKIL(ws, "btnBaglantiKontrol")
    If Not s Is Nothing Then s.Left = x + 1
    Set s = TEMA_ANA_SEKIL(ws, "btnCDP")
    If Not s Is Nothing Then s.Left = x + 1
    mAsama = "M_ANA PANEL / altý hýzlý eriþim"
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 14) = "MUSTERI_HIZLI_" Then ws.Shapes(i).Delete
    Next i
    basliklar = Array("CDP STOK", "CDP RAPOR", "ÝLÇE STOK", "ÝLÇE RAPOR", _
        "YÜKLENÝCÝ STOK", "YÜKLENÝCÝ RAPOR")
    hedefler = Array("M_CDP_STOK", "M_CDP_RAPOR", "M_ÝLÇE_STOK", "M_ÝLÇE_RAPOR", _
        "M_YÜKLENÝCÝ_STOK", "M_YÜKLENÝCÝ_RAPOR")
    Set r = ws.Range("B25:T26")
    aralik = 8: bw = (r.Width - 5 * aralik) / 6
    For i = 0 To 5
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x + 1 + i * (bw + aralik), _
            r.Top + 1, bw, r.Height - 2)
        s.Name = "MUSTERI_HIZLI_" & CStr(i)
        s.Adjustments.Item(1) = 0.12
        s.Fill.ForeColor.RGB = RGB(244, 251, 255)
        s.line.ForeColor.RGB = RGB(112, 195, 228): s.line.Weight = 0.7
        s.Placement = xlFreeFloating
        s.AlternativeText = CStr(hedefler(i))
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_HIZLI_AC"
        With s.TextFrame2
            .MarginLeft = 3: .MarginRight = 3: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
            .TextRange.Text = CStr(basliklar(i))
            .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 8
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(15, 68, 106)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    Next i
    arama = Array("btnMalzemeAraPanel", "btnMalzemeAra", "btnMalzemeArama", "btnArama")
    For i = 0 To UBound(arama)
        Set s = TEMA_ANA_SEKIL(ws, CStr(arama(i)))
        If Not s Is Nothing Then
            s.AlternativeText = "M_MALZEME_ARA"
            s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_HIZLI_AC"
        End If
    Next i
    ws.Range("BX6").Value2 = "MUSTERI_ANA_V13|" & kod
End Sub

Private Sub TEMA_MUSTERI_ARAMA_BASLIKLAR(ByVal ws As Worksheet)
    Dim basliklar As Variant, agirlik As Variant, i As Long, s As Shape
    Dim tabloW As Double, toplam As Double, oldMeta As Variant, metin As String
    If CStr(ws.Range("BX7").Value2) = "MUSTERI_ARAMA_V14" Then Exit Sub
    mAsama = "M_MALZEME_ARA / baþlýk yerleþimi"
    tabloW = ws.Range("A1:K1").Width
    oldMeta = ws.Range("BX1:BX7").Value2
    ' Mevcut sag il secim alani iki kolon saga tasinir; yatay konumu korunur.
    ws.Columns("K:L").Insert Shift:=xlToRight
    ws.Range("BZ1:BZ7").ClearContents
    ws.Range("BX1:BX7").Value2 = oldMeta
    ws.Columns("BX").Hidden = True
    agirlik = Array(11, 11, 14, 13, 14, 30, 10, 8, 17, 20, 16, 12, 12)
    toplam = 0
    For i = 0 To 12
        toplam = toplam + CDbl(agirlik(i))
    Next i
    For i = 0 To 12
        ws.Columns(i + 1).ColumnWidth = CDbl(agirlik(i)) * (tabloW / 6.8) / toplam
    Next i
    ' Karakter olcusunden gelen yuvarlamayi gercek nokta genisligine gore duzelt.
    toplam = ws.Range("A1:M1").Width
    For i = 1 To 13
        ws.Columns(i).ColumnWidth = ws.Columns(i).ColumnWidth * tabloW / toplam
    Next i
    basliklar = Array("Üretim yeri", "Depo yeri", "Depo/Fikti", "Mal grubu", _
        "Malzeme", "Taným", "Miktar", "Birim", "Deðerleme türü", "Seri numarasý", _
        "Deðer", "Kaynak", "Ýl")
    For i = 0 To 12
        ws.Cells(12, i + 1).Value2 = CStr(basliklar(i))
    Next i
    With ws.Range("A12:M12")
        .Interior.Color = RGB(24, 55, 88)
        .Font.Name = "Aptos": .Font.Size = 9: .Font.bold = True
        .Font.Color = RGB(255, 255, 255)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = True
        .Borders.LineStyle = xlNone
    End With
    ws.Rows(12).RowHeight = application.Max(ws.Rows(12).RowHeight, 28)
    ws.Range("A5").MergeArea.Cells(1, 1).Value2 = "MALZEME ADI, KODU VEYA SERÝ NUMARASI"
    ws.Range("A7").MergeArea.Cells(1, 1).Value2 = _
        "Malzeme adý, kodun tamamý veya bir kýsmý ya da seri numarasý ile arama yapabilirsiniz."
    For Each s In ws.Shapes
        metin = ""
        If s.Type = msoAutoShape Or s.Type = msoTextBox Then
            metin = s.TextFrame2.TextRange.Text
            If InStr(1, metin, "MALZEME KODU VEYA MALZEME", vbTextCompare) > 0 Then _
                s.TextFrame2.TextRange.Text = "MALZEME ADI, KODU VEYA SERÝ NUMARASI"
            If InStr(1, metin, "Malzeme kodunun tamam", vbTextCompare) > 0 Then _
                s.TextFrame2.TextRange.Text = "Malzeme adý, kodun tamamý veya bir kýsmý ya da seri numarasý ile arama yapabilirsiniz."
        End If
    Next s
    ws.Range("BX7").Value2 = "MUSTERI_ARAMA_V14"
End Sub

Private Sub TEMA_MUSTERI_ARAMA_V15(ByVal ws As Worksheet)
    Dim q As String, s As Shape, ilButonu As Shape, ara As Shape, r As Range
    Dim kodlar As Variant, i As Long, metin As String, ad As String
    mAsama = "M_MALZEME_ARA / giriþ ve filtre"
    ws.Unprotect ""
    ws.EnableSelection = xlNoRestrictions
    q = CStr(ws.Range("B6").Value2)
    ws.Range("B6").MergeArea.UnMerge
    With ws.Range("B6:J6")
        .Merge
        .Cells(1, 1).Value2 = q
        .Locked = False
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(178, 195, 209)
        .Borders.Weight = xlThin
        .Font.Name = "Aptos": .Font.Size = 11
        .Font.Color = RGB(24, 55, 88)
        .VerticalAlignment = xlCenter
    End With
    For Each s In ws.Shapes
        If s.Name = "btnMalzemeAra" Then Set ara = s
        If s.Name = "AKSORA_V4_TUMU" Then Set ilButonu = s
        If s.Name = "btnMalzemeAnaPanel" Then
            s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_ARAMA_ANA_DON"
        End If
    Next s
    If Not ara Is Nothing Then
        Set r = ws.Range("K6:M6")
        ara.Left = r.Left + 4: ara.Top = ws.Rows(6).Top + 1
        ara.Width = r.Width - 4: ara.Height = ws.Rows(6).Height - 2
        ara.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_ARAMA_HAZIRLIK"
    End If
    ' Kaynak hucre basliginin kayik zemini temizlenir; diger il dugmeleri esas alinir.
    ws.Range("O6:P7").UnMerge
    ws.Range("O6:P7").ClearContents
    ws.Range("O6:P7").Interior.Color = RGB(255, 255, 255)
    ws.Range("N6:N7").Borders.LineStyle = xlNone
    Set s = TEMA_ANA_SEKIL(ws, "MUSTERI_ARAMA_IL_BASLIK")
    If Not ilButonu Is Nothing Then
        If s Is Nothing Then
            Set s = ws.Shapes.AddShape(msoShapeRectangle, ilButonu.Left, ws.Rows(6).Top, _
                ilButonu.Width, ilButonu.Top - ws.Rows(6).Top - 4)
            s.Name = "MUSTERI_ARAMA_IL_BASLIK"
        End If
        s.Left = ilButonu.Left: s.Width = ilButonu.Width
        s.Top = ws.Rows(6).Top: s.Height = ilButonu.Top - s.Top - 4
        s.Fill.ForeColor.RGB = RGB(24, 55, 88): s.line.Visible = msoFalse
        With s.TextFrame2
            .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = "ARAMA YAPILACAK ÝL"
            .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 10
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End If
    kodlar = Array("TUMU", "GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    For i = 0 To 5
        Set s = TEMA_ANA_SEKIL(ws, "AKSORA_V4_" & CStr(kodlar(i)))
        If Not s Is Nothing Then
            s.AlternativeText = CStr(kodlar(i))
            s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_ARAMA_IL_SEC"
        End If
    Next i
    TEMA_ARAMA_KPI_FORMULLER ws
    ws.ScrollArea = "A1:P60000"
    TEMA_MUSTERI_TOPLU_BUTON ws
End Sub

Private Sub TEMA_ARAMA_KPI_FORMULLER(ByVal ws As Worksheet)
    Dim son As Long, a As String, b As String, e As String, g As String, k As String, m As String
    son = application.Max(200, AKSORA_M_SON(ws))
    If Not ws.AutoFilterMode Then ws.Range("A12:M" & son).AutoFilter
    a = "$A$13:$A$" & son: b = "$B$13:$B$" & son
    e = "$E$13:$E$" & son: g = "$G$13:$G$" & son
    k = "$K$13:$K$" & son: m = "$M$13:$M$" & son
    ws.Range("A10").MergeArea.Cells(1, 1).Formula = "=SUBTOTAL(103," & e & ")"
    ws.Range("C10").MergeArea.Cells(1, 1).Value2 = TEMA_GORUNUR_DEPO_SAYISI(ws, son)
    ws.Range("F10").MergeArea.Cells(1, 1).Formula = "=SUBTOTAL(109," & g & ")"
    ws.Range("I10").MergeArea.Cells(1, 1).Formula = "=SUBTOTAL(109," & k & ")"
    ws.Range("A10").MergeArea.NumberFormat = "#,##0"
    ws.Range("C10").MergeArea.NumberFormat = "#,##0"
    ws.Range("F10").MergeArea.NumberFormat = "#,##0.###"
    ws.Range("I10").MergeArea.NumberFormat = "#,##0.00 ""TL"""
    ws.Calculate
    AKSORA_MUSTERI_ARAMA_KPI_HESAPLA ws
End Sub

Public Function MUSTERI_GORUNUR_DEPO(ByVal uretim As Range, ByVal depo As Range, _
    ByVal malzeme As Range, ByVal il As Range) As Variant
    Dim d As Object, i As Long, anahtar As String
    On Error GoTo Hata
    application.Volatile True
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For i = 1 To malzeme.Rows.Count
        If Not malzeme.Cells(i, 1).EntireRow.Hidden Then
            If Len(Trim$(CStr(malzeme.Cells(i, 1).Value2))) > 0 Then
                anahtar = Trim$(CStr(uretim.Cells(i, 1).Value2)) & "|" & _
                    Trim$(CStr(depo.Cells(i, 1).Value2)) & "|" & Trim$(CStr(il.Cells(i, 1).Value2))
                If Not d.Exists(anahtar) Then d.Add anahtar, True
            End If
        End If
    Next i
    MUSTERI_GORUNUR_DEPO = d.Count
    Exit Function
Hata:
    MUSTERI_GORUNUR_DEPO = CVErr(xlErrValue)
End Function

Public Sub AKSORA_MUSTERI_ARAMA_KPI_YENILE()
    Dim ws As Worksheet, ev As Boolean, sc As Boolean
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    Set ws = AKSORA_ARAMA_SAYFASI()
    ws.Unprotect ""
    ' Filtre yalnizca sonuc alanini kapsamiyorsa yeniden kurulur.
    If ws.AutoFilterMode Then
        If ws.AutoFilter.Range.Row <> 12 Or ws.AutoFilter.Range.Rows.Count < AKSORA_M_SON(ws) - 11 Then ws.AutoFilterMode = False
    End If
    TEMA_ARAMA_KPI_FORMULLER ws
Cikis:
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If Not ws Is Nothing Then
        ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
        ws.EnableSelection = xlNoRestrictions
    End If
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Public Sub AKSORA_MUSTERI_ARAMA_IL_SEC()
    Dim ws As Worksheet, kod As String, ev As Boolean, sc As Boolean
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    Set ws = AKSORA_ARAMA_SAYFASI()
    kod = ws.Shapes(CStr(application.Caller)).AlternativeText
    Select Case kod
        Case "TUMU", "GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN"
        Case Else: GoTo Cikis
    End Select
    AKSORA_MUSTERI_ARAMA_FILTRE kod
Cikis:
    On Error Resume Next
    If Not ws Is Nothing Then
        If Not (ActiveSheet Is ws) Then ws.Activate
        ws.Range("B6").Select
    End If
    application.EnableEvents = ev: application.ScreenUpdating = sc
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

' Arama sonucu secilen ile gore suzulur; secim BX21'de saklanir ve dugme rengi guncellenir.
Public Sub AKSORA_MUSTERI_ARAMA_FILTRE(ByVal kod As String)
    Dim ws As Worksheet, isim As String, s As Shape, ad As String, ev As Boolean
    ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    Set ws = AKSORA_ARAMA_SAYFASI()
    kod = UCase$(Trim$(kod))
    Select Case kod
        Case "GIRESUN": isim = "GÝRESUN"
        Case "TRABZON": isim = "TRABZON"
        Case "RIZE": isim = "RÝZE"
        Case "GUMUSHANE": isim = "GÜMÜÞHANE"
        Case "ARTVIN": isim = "ARTVÝN"
        Case Else: kod = "TUMU": isim = ""
    End Select
    ws.Unprotect ""
    ws.Range("BX21").Value2 = kod
    If Not ws.AutoFilterMode Then TEMA_ARAMA_KPI_FORMULLER ws
    If Len(isim) = 0 Then
        ws.AutoFilter.Range.AutoFilter Field:=13
    Else
        ws.AutoFilter.Range.AutoFilter Field:=13, Criteria1:=isim
    End If
    For Each s In ws.Shapes
        ad = s.Name
        If ad = "AKSORA_V4_TUMU" Or ad = "AKSORA_V4_GIRESUN" Or ad = "AKSORA_V4_TRABZON" Or _
           ad = "AKSORA_V4_RIZE" Or ad = "AKSORA_V4_GUMUSHANE" Or ad = "AKSORA_V4_ARTVIN" Then
            If s.AlternativeText = kod Then
                s.Fill.ForeColor.RGB = RGB(0, 121, 103)
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            Else
                s.Fill.ForeColor.RGB = RGB(230, 240, 248)
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(24, 55, 88)
            End If
        End If
    Next s
    ws.Calculate
    AKSORA_MUSTERI_ARAMA_KPI_HESAPLA ws
    AKSORA_MUSTERI_BANT ws
Cikis:
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
        ws.EnableSelection = xlNoRestrictions
    End If
    application.EnableEvents = ev
    On Error GoTo 0
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

' Arama sayfasi: etkin sayfa yatirim aramasiysa o, degilse musteri aramasi.
Public Function AKSORA_ARAMA_SAYFASI() As Worksheet
    On Error Resume Next
    If ActiveSheet.Name = "Y_MALZEME_ARA" Or ActiveSheet.Name = "I_MALZEME_ARA" Then Set AKSORA_ARAMA_SAYFASI = ActiveSheet
    On Error GoTo 0
    If AKSORA_ARAMA_SAYFASI Is Nothing Then Set AKSORA_ARAMA_SAYFASI = ThisWorkbook.Worksheets("M_MALZEME_ARA")
End Function

Public Sub AKSORA_MUSTERI_ARAMA_ANA_DON()
    With ThisWorkbook.Worksheets("M_ANA PANEL")
        .Visible = xlSheetVisible
        .Select Replace:=True
    End With
End Sub

Public Sub AKSORA_MUSTERI_ARAMA_HAZIRLIK()
    AKSORA_MUSTERI_YEREL_ARA
End Sub

Private Sub TEMA_BOS_DEGER(ByVal ws As Worksheet, ByVal adres As String)
    ws.Range(adres).MergeArea.Cells(1, 1).Value2 = "-"
End Sub

Private Function TEMA_SEKIL_KOPYALA(ByVal kaynak As Worksheet, ByVal hedef As Worksheet, _
    ByVal sekil As Shape) As Shape
    Dim gorunum As XlSheetVisibility, ekran As Boolean, deneme As Long
    Dim adet As Long, no As Long, metin As String, i As Long
    ' Form denetimleri panoyla tasinmaz; ayni yer ve boyutta yerel denetim olusturulur.
    If sekil.Type = msoFormControl Or LCase$(Left$(sekil.Name, 9)) = "drop down" Then
        Set TEMA_SEKIL_KOPYALA = TEMA_FORM_KUR(hedef, sekil)
        Exit Function
    End If
    If sekil.Type = msoAutoShape Or sekil.Type = msoTextBox Or sekil.Type = msoLine Then
        Set TEMA_SEKIL_KOPYALA = TEMA_YEREL_SEKIL(hedef, sekil)
        Exit Function
    End If
    If sekil.Type = msoPicture Or sekil.Type = msoLinkedPicture Then
        Set TEMA_SEKIL_KOPYALA = TEMA_DOSYA_RESIM(hedef, kaynak, sekil)
        Exit Function
    End If
    gorunum = kaynak.Visible: ekran = application.ScreenUpdating
    On Error GoTo Hata
    kaynak.Visible = xlSheetVisible
    hedef.Visible = xlSheetVisible
    hedef.EnableSelection = xlNoRestrictions
    application.ScreenUpdating = False
    For deneme = 1 To 3
        adet = hedef.Shapes.Count
        application.CutCopyMode = False
        kaynak.Activate
        sekil.Copy
        hedef.Activate
        hedef.Range("A1").Select
        On Error Resume Next
        Err.Clear
        hedef.Paste
        no = Err.Number: metin = Err.Description
        On Error GoTo Hata
        If no = 0 And hedef.Shapes.Count = adet + 1 Then
            Set TEMA_SEKIL_KOPYALA = hedef.Shapes(hedef.Shapes.Count)
            Exit For
        End If
        ' Basarisiz denemeden kalan sekil varsa yalniz hedefte temizlenir.
        For i = hedef.Shapes.Count To adet + 1 Step -1
            hedef.Shapes(i).Delete
        Next i
    Next deneme
    If TEMA_SEKIL_KOPYALA Is Nothing Then
        Err.Raise vbObjectError + 1701, , "Görsel kopyalanamadý: " & kaynak.Name & _
            " / " & sekil.Name & ". " & metin
    End If
    no = 0
Cikis:
    On Error Resume Next
    hedef.Activate
    kaynak.Visible = gorunum
    application.CutCopyMode = False
    application.ScreenUpdating = ekran
    On Error GoTo 0
    If no <> 0 Then Err.Raise no, "TEMA_SEKIL_KOPYALA", metin
    Exit Function
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Function

Private Function TEMA_DOSYA_RESIM(ByVal hedef As Worksheet, ByVal kaynak As Worksheet, _
    ByVal orijinal As Shape) As Shape
    Dim yol As String, s As Shape
    yol = TEMA_RESIM_YOLU(kaynak, orijinal)
    If Len(yol) = 0 Then Err.Raise vbObjectError + 1710, , _
        "Görsel dosyasý bulunamadý: " & kaynak.Name & " / " & orijinal.Name & _
        ". Orijinal AKSORA PNG dosyalarýný Excel dosyasýnýn yanýna koyun."
    Set s = hedef.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
    With s.PictureFormat
        .CropTop = orijinal.PictureFormat.CropTop
        .CropBottom = orijinal.PictureFormat.CropBottom
        .CropLeft = orijinal.PictureFormat.CropLeft
        .CropRight = orijinal.PictureFormat.CropRight
    End With
    s.LockAspectRatio = msoFalse
    s.Left = orijinal.Left: s.Top = orijinal.Top
    s.Width = orijinal.Width: s.Height = orijinal.Height
    s.Rotation = orijinal.Rotation
    s.AlternativeText = orijinal.AlternativeText
    s.OnAction = ""
    Set TEMA_DOSYA_RESIM = s
End Function

Private Function TEMA_RESIM_YOLU(ByVal kaynak As Worksheet, ByVal sekil As Shape) As String
    Dim klasor As String, aday As String, ad As String, alt As String, il As String
    klasor = ThisWorkbook.Path & application.PathSeparator
    alt = sekil.AlternativeText
    ' Orijinal resmin dosya yolu kayitliysa once onu kullan.
    If InStr(1, alt, ".png", vbTextCompare) > 0 Or InStr(1, alt, ".bmp", vbTextCompare) > 0 Or _
       InStr(1, alt, ".jpg", vbTextCompare) > 0 Then
        If TEMA_DOSYA_VAR(alt) Then TEMA_RESIM_YOLU = alt: Exit Function
        aday = Replace(alt, "/", application.PathSeparator)
        aday = Mid$(aday, InStrRev(aday, application.PathSeparator) + 1)
        If TEMA_DOSYA_VAR(AKSORA_GORSEL(aday)) Then TEMA_RESIM_YOLU = AKSORA_GORSEL(aday): Exit Function
    End If
    ad = UCase$(sekil.Name)
    If InStr(ad, "FILIGRAN") > 0 Or ad = "AKSORA_STOK_SON_SEFFAF_LOGO" Then
        aday = "AKSORA_RAPOR_FILIGRAN.png"
    ElseIf InStr(ad, "TEK_ARKA_PLAN") > 0 Then
        If alt = "AKSORA_P1_SADE_V1" Then
            aday = "AKSORA_PANEL_SADE_P1.png"
        Else
            aday = "AKSORA_PANEL_TEK_PARCA_V7.png"
        End If
    ElseIf InStr(ad, "PANEL_TEK_") > 0 Then
        ' modAKSORA_TASARIM: AKT1_PANEL_TEK_<IL>_20261007
        il = Mid$(ad, InStr(ad, "PANEL_TEK_") + Len("PANEL_TEK_"))
        If InStr(il, "_") > 0 Then il = Left$(il, InStr(il, "_") - 1)
        aday = "AKSORA_PANEL_BANNER_" & il & ".png"
    ElseIf Left$(ad, Len("AKSORA_IL_BANNER_")) = "AKSORA_IL_BANNER_" Then
        il = Mid$(ad, Len("AKSORA_IL_BANNER_") + 1)
        aday = "AKSORA_PANEL_BANNER_" & il & ".png"
    ElseIf InStr(ad, "BANNER") > 0 Then
        If kaynak.Name = "ANA PANEL" Then
            aday = "AKSORA_PANEL_BANNER.png"
        Else
            aday = "AKSORA_RAPOR_BANNER.png"
        End If
    End If
    If Len(aday) > 0 Then
        If TEMA_DOSYA_VAR(AKSORA_GORSEL(aday)) Then TEMA_RESIM_YOLU = AKSORA_GORSEL(aday)
    End If
End Function

Private Function TEMA_DOSYA_VAR(ByVal yol As String) As Boolean
    On Error GoTo Cikis
    TEMA_DOSYA_VAR = (Len(Dir$(yol)) > 0)
Cikis:
End Function

Private Function TEMA_YEREL_SEKIL(ByVal hedef As Worksheet, ByVal kaynakSekil As Shape) As Shape
    Dim s As Shape, a As Object, b As Object, i As Long, metin As String
    Select Case kaynakSekil.Type
        Case msoAutoShape
            Set s = hedef.Shapes.AddShape(kaynakSekil.AutoShapeType, kaynakSekil.Left, _
                kaynakSekil.Top, kaynakSekil.Width, kaynakSekil.Height)
        Case msoTextBox
            Set s = hedef.Shapes.AddTextbox(msoTextOrientationHorizontal, kaynakSekil.Left, _
                kaynakSekil.Top, kaynakSekil.Width, kaynakSekil.Height)
        Case msoLine
            Set s = hedef.Shapes.AddLine(kaynakSekil.Left, kaynakSekil.Top, _
                kaynakSekil.Left + kaynakSekil.Width, kaynakSekil.Top + kaynakSekil.Height)
    End Select
    ' PickUp/Apply sekil bicimini aktarir; Windows panosu ve Shape.Copy kullanilmaz.
    kaynakSekil.PickUp
    s.Apply
    s.Rotation = kaynakSekil.Rotation
    If kaynakSekil.Type <> msoLine Then
        For i = 1 To kaynakSekil.Adjustments.Count
            s.Adjustments.Item(i) = kaynakSekil.Adjustments.Item(i)
        Next i
        With s.TextFrame2
            .MarginLeft = kaynakSekil.TextFrame2.MarginLeft
            .MarginRight = kaynakSekil.TextFrame2.MarginRight
            .MarginTop = kaynakSekil.TextFrame2.MarginTop
            .MarginBottom = kaynakSekil.TextFrame2.MarginBottom
            .VerticalAnchor = kaynakSekil.TextFrame2.VerticalAnchor
            .WordWrap = kaynakSekil.TextFrame2.WordWrap
            .Orientation = kaynakSekil.TextFrame2.Orientation
            metin = kaynakSekil.TextFrame2.TextRange.Text
            .TextRange.Text = metin
        End With
        ' Farkli punto ve renkleri olan metinlerde de karakter bicimleri korunur.
        For i = 1 To Len(metin)
            Set a = kaynakSekil.TextFrame2.TextRange.Characters(i, 1)
            Set b = s.TextFrame2.TextRange.Characters(i, 1)
            With b.Font
                .Name = a.Font.Name
                .Size = a.Font.Size
                .bold = a.Font.bold
                .Italic = a.Font.Italic
                .Fill.ForeColor.RGB = a.Font.Fill.ForeColor.RGB
            End With
            b.ParagraphFormat.Alignment = a.ParagraphFormat.Alignment
        Next i
    End If
    s.OnAction = ""
    Set TEMA_YEREL_SEKIL = s
End Function

Private Function TEMA_FORM_KUR(ByVal hedef As Worksheet, ByVal kaynakSekil As Shape) As Shape
    Dim s As Shape, baslik As String
    ' Tema asamasinda form/ActiveX denetimi yerine islemsiz gorunum olusturulur.
    ' ControlFormat, AddFormControl ve Windows panosu kullanilmaz.
    Set s = hedef.Shapes.AddShape(msoShapeRectangle, kaynakSekil.Left, kaynakSekil.Top, _
        kaynakSekil.Width, kaynakSekil.Height)
    baslik = "TÜMÜ"
    If LCase$(Left$(kaynakSekil.Name, 9)) <> "drop down" Then
        On Error Resume Next
        baslik = kaynakSekil.TextFrame.Characters.Text
        On Error GoTo 0
        If Len(baslik) = 0 Then baslik = "TÜMÜ"
    End If
    s.Fill.ForeColor.RGB = RGB(255, 255, 255)
    s.line.ForeColor.RGB = RGB(150, 150, 150)
    s.line.Weight = 0.75
    With s.TextFrame2
        .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = baslik & "   " & ChrW(9662)
        .TextRange.Font.Name = "Calibri"
        .TextRange.Font.Size = 10
        .TextRange.Font.Fill.ForeColor.RGB = RGB(40, 40, 40)
    End With
    s.OnAction = ""
    Set TEMA_FORM_KUR = s
End Function

Private Sub TEMA_EYLEM_SIL(ByVal s As Shape)
    Dim i As Long
    s.OnAction = ""
    If s.Type = msoFormControl Then
        On Error Resume Next
        s.ControlFormat.LinkedCell = ""
        s.ControlFormat.ListFillRange = ""
        On Error GoTo 0
    End If
    If s.Type = msoGroup Then
        For i = 1 To s.GroupItems.Count
            TEMA_EYLEM_SIL s.GroupItems(i)
        Next i
    End If
End Sub

Private Function TEMA_DEGISKEN_METIN(ByVal metin As String) As Boolean
    Dim rx As Object
    Set rx = CreateObject("VBScript.RegExp")
    rx.Pattern = "[0-9]{2}\.[0-9]{2}\.[0-9]{4}|[0-9]+[.,][0-9]+|[0-9]+ SATIR"
    rx.IgnoreCase = True
    TEMA_DEGISKEN_METIN = rx.Test(metin) Or metin = "GÜNCEL"
End Function

Private Function TEMA_IL_ADI() As String
    Dim kod As String
    kod = CStr(ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2)
    Select Case kod
        Case "GIRESUN": TEMA_IL_ADI = "Giresun"
        Case "TRABZON": TEMA_IL_ADI = "Trabzon"
        Case "RIZE": TEMA_IL_ADI = "Rize"
        Case "GUMUSHANE": TEMA_IL_ADI = "Gümüþhane"
        Case "ARTVIN": TEMA_IL_ADI = "Artvin"
    End Select
End Function

Public Sub AKSORA_MUSTERI_TEMA_GORUNUM(ByVal ws As Worksheet)
    Dim s As Shape, genislik As Double, z As Double, altSinir As Double
    Dim sc As Boolean, no As Long, metin As String
    If Not AKSORA_MUSTERI_TEMA_SAYFA(ws) Then Exit Sub
    ' Gecis sirasinda ekran tek seferde cizilir (titreme olmaz).
    sc = application.ScreenUpdating
    application.ScreenUpdating = False
    On Error Resume Next
    application.MoveAfterReturn = (ws.Name <> "M_MALZEME_ARA")
    On Error GoTo Hata
    AKSORA_MUSTERI_FILTRE_SIFIRLA ws.Name
    ' Arama sonucu yalniz baska sayfadan gelince temizlenir (pencere yenilemesinde korunur).
    If ws.Name = "M_MALZEME_ARA" And mSonGorunen <> "M_MALZEME_ARA" Then
        If AKSORA_M_SON(ws) >= 13 Or Len(Trim$(CStr(ws.Range("B6").Value2))) > 0 Then AKSORA_MUSTERI_ARAMA_TEMIZLE
    End If
    mSonGorunen = ws.Name
    AKSORA_MUSTERI_SAYFA_HAZIRLA ws
    If AKSORA_PENCERE_KUCUK(ThisWorkbook.Windows(1)) Then GoTo Cikis
    With ThisWorkbook.Windows(1)
        If Not .DisplayWorkbookTabs Then .DisplayWorkbookTabs = True
        If .View <> xlNormalView Then .View = xlNormalView
        On Error Resume Next
        If .FreezePanes Then .FreezePanes = False
        If .SplitRow <> 0 Then .SplitRow = 0
        If .SplitColumn <> 0 Then .SplitColumn = 0
        On Error GoTo Hata
        If .DisplayGridlines Then .DisplayGridlines = False
        If .DisplayHeadings Then .DisplayHeadings = False
        If .DisplayHorizontalScrollBar Then .DisplayHorizontalScrollBar = False
        ' Kaydirma cubugu tum sayfalarda ayni: gecislerde pencere yeniden yerlesmez.
        If Not .DisplayVerticalScrollBar Then .DisplayVerticalScrollBar = True
    End With
    ' Yakinlastirma, kenar hizalama ve kart yerlesimi (pencere/veri degismedikce atlanir).
    If ws.Name = "M_ANA PANEL" Then
        AKSORA_MUSTERI_ALT_BANT ws
        ' Ana panel ozgun yerlesimiyle pencereye sigdirilir (sekiller tasinmaz).
        genislik = ws.Range("A1:L1").Width
        For Each s In ws.Shapes
            If s.Visible = msoTrue Then
                genislik = application.Max(genislik, s.Left + s.Width)
                altSinir = application.Max(altSinir, s.Top + s.Height)
            End If
        Next s
        With ThisWorkbook.Windows(1)
            z = (.UsableWidth - 8) / genislik * 100
            If altSinir > 0 Then z = application.Min(z, (.UsableHeight - 8) / altSinir * 100)
            z = application.Max(10, application.Min(400, Fix(z)))
            If .Zoom <> z Then .Zoom = z
        End With
    Else
        AKSORA_MUSTERI_DUZEN ws
    End If
    With ThisWorkbook.Windows(1)
        If .ScrollRow <> 1 Then .ScrollRow = 1
        If .ScrollColumn <> 1 Then .ScrollColumn = 1
    End With
Cikis:
    application.ScreenUpdating = sc
    If no <> 0 Then Err.Raise no, "AKSORA_MUSTERI_TEMA_GORUNUM", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Private Function TEMA_GORUNUR_DEPO_SAYISI(ByVal ws As Worksheet, ByVal son As Long) As Long
    Dim veri As Variant, gorunen As Range, parca As Range, d As Object
    Dim i As Long, ilk As Long, bitis As Long, anahtar As String, a As String, b As String, il As String
    Dim hataNo As Long, hataMetni As String
    If son < 13 Then Exit Function
    veri = ws.Range("A13:M" & son).Value2
    On Error Resume Next
    Set gorunen = ws.Range("E13:E" & son).SpecialCells(xlCellTypeVisible)
    hataNo = Err.Number: hataMetni = Err.Description
    On Error GoTo 0
    If hataNo <> 0 And hataNo <> 1004 Then Err.Raise hataNo, , hataMetni
    If gorunen Is Nothing Then Exit Function
    Set gorunen = application.Intersect(gorunen, ws.Range("E13:E" & son))
    If gorunen Is Nothing Then Exit Function
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For Each parca In gorunen.Areas
        ilk = parca.Row - 12: bitis = ilk + parca.Rows.Count - 1
        For i = ilk To bitis
            If Not IsError(veri(i, 5)) Then
                If Len(Trim$(CStr(veri(i, 5)))) > 0 Then
                    If IsError(veri(i, 1)) Or IsError(veri(i, 2)) Or IsError(veri(i, 13)) Then _
                        Err.Raise vbObjectError + 1720, , "Depo bilgisi hatalý: " & CStr(i + 12) & ". satýr"
                    a = Trim$(CStr(veri(i, 1))): b = Trim$(CStr(veri(i, 2)))
                    il = Trim$(CStr(veri(i, 13)))
                    anahtar = Len(a) & ":" & a & Len(b) & ":" & b & Len(il) & ":" & il
                    d(anahtar) = True
                End If
            End If
        Next i
    Next parca
    TEMA_GORUNUR_DEPO_SAYISI = d.Count
End Function

Public Sub AKSORA_MUSTERI_ARAMA_KPI_HESAPLA(ByVal ws As Worksheet)
    Static mesgul As Boolean
    Dim ev As Boolean, korumali As Boolean, son As Long, adet As Long
    Dim hataNo As Long, hataMetni As String
    If mesgul Or (ws.Name <> "M_MALZEME_ARA" And ws.Name <> "Y_MALZEME_ARA" And ws.Name <> "I_MALZEME_ARA") Then Exit Sub
    mesgul = True: ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    AKSORA_SAYI_AYIRACLARI
    son = application.Max(13, AKSORA_M_SON(ws))
    adet = TEMA_GORUNUR_DEPO_SAYISI(ws, son)
    korumali = ws.ProtectContents
    If korumali Then ws.Unprotect ""
    With ws.Range("C10").MergeArea.Cells(1, 1)
        If .HasFormula Or IsError(.Value2) Then
            .Value2 = adet
        ElseIf Not IsNumeric(.Value2) Then
            .Value2 = adet
        ElseIf CDbl(.Value2) <> adet Then
            .Value2 = adet
        End If
    End With
    ws.Range("A10").MergeArea.NumberFormat = "#,##0"
    ws.Range("C10").MergeArea.NumberFormat = "#,##0"
    AKSORA_V13_MIKTAR_BICIMI ws.Range("F10")
Cikis:
    On Error Resume Next
    If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev: mesgul = False
    On Error GoTo 0
    If hataNo <> 0 Then Err.Raise hataNo, "AKSORA_MUSTERI_ARAMA_KPI_HESAPLA", hataMetni
    Exit Sub
Hata:
    hataNo = Err.Number: hataMetni = Err.Description
    Resume Cikis
End Sub

Private Sub TEMA_MUSTERI_TOPLU_BUTON(ByVal ws As Worksheet)
    Dim s As Shape, vurgu As Shape, durum As Shape, sure As Double, metin As String
    ' Malzeme birimi arama sayfasindaki dugmenin birebir aynisi (renk, cizgi, golge, vurgu, durum yazisi).
    Set s = ws.Shapes("AKSORA_TOPLU_GUNCELLE")
    With s
        .Adjustments.Item(1) = 0.12
        .TextFrame2.TextRange.Text = "BÜTÜN ÝLLERÝ GÜNCELLE" & vbLf & "IQ09 + MB52  |  " & _
            CStr(UBound(AKSORA_MUSTERI_ILLER()) + 1) & " ÝL"
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(225, 255, 246)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 14: .TextFrame2.MarginRight = 10
        .TextFrame2.MarginTop = 0: .TextFrame2.MarginBottom = 0
        .Fill.ForeColor.RGB = RGB(9, 45, 62)
        .Fill.BackColor.RGB = RGB(0, 112, 84)
        .Fill.TwoColorGradient msoGradientHorizontal, 1
        .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(28, 211, 163): .line.Weight = 1.25
        .Shadow.Visible = msoTrue
        .Shadow.ForeColor.RGB = RGB(10, 39, 50)
        .Shadow.Transparency = 0.8
        .Shadow.OffsetX = 1: .Shadow.OffsetY = 2
        .Placement = xlFreeFloating
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_BUTUN_ILLERI_GUNCELLE"
    End With
    On Error Resume Next
    Set vurgu = ws.Shapes("AKSORA_TOPLU_VURGU")
    Set durum = ws.Shapes("AKSORA_TOPLU_DURUM")
    On Error GoTo 0
    If vurgu Is Nothing Then Set vurgu = ws.Shapes.AddShape(msoShapeRectangle, 0, 0, 3, 10): vurgu.Name = "AKSORA_TOPLU_VURGU"
    With vurgu
        .Left = s.Left + 6: .Top = s.Top + 10
        .Width = 3: .Height = application.Max(6, s.Height - 20)
        .Fill.ForeColor.RGB = RGB(48, 231, 179): .line.Visible = msoFalse
        .Shadow.Visible = msoFalse
        .Placement = xlFreeFloating
        .OnAction = s.OnAction
        .ZOrder msoBringToFront
    End With
    If Not durum Is Nothing Then
        With durum
            .Left = s.Left - 2: .Width = s.Width + 4
            .Top = s.Top + s.Height + 2
            .Fill.Visible = msoFalse: .line.Visible = msoFalse
            .Placement = xlFreeFloating
        End With
    End If
    If AKSORA_MUSTERI_TOPLU_AKTIF Then
        AKSORA_MUSTERI_TOPLU_TIK True
    Else
        sure = -1
        On Error Resume Next
        sure = CDbl(application.Evaluate(ThisWorkbook.names("_AKSORA_MUSTERI_SON_SURE").RefersTo))
        metin = CStr(application.Evaluate(ThisWorkbook.names("_AKSORA_MUSTERI_SON_DURUM").RefersTo))
        On Error GoTo 0
        If Len(metin) = 0 Then metin = "Arama, her ilin yukarýdaki tarihli stoklarýný kullanýr."
        TEMA_MUSTERI_SURE_YAZ metin, sure
    End If
End Sub

Public Sub AKSORA_MUSTERI_BUTON_GUNCELLE()
    Dim ws As Worksheet, korumali As Boolean, hataMetni As String
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("M_MALZEME_ARA")
    korumali = ws.ProtectContents
    ws.Unprotect ""
    TEMA_MUSTERI_TOPLU_BUTON ws
Cikis:
    If Not ws Is Nothing And korumali Then ws.Protect Password:="", DrawingObjects:=True, _
        Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    If Len(hataMetni) > 0 Then MsgBox hataMetni, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Cikis
End Sub

Private Sub TEMA_MUSTERI_SURE_YAZ(ByVal metin As String, ByVal sure As Double)
    Dim s As Shape
    Set s = ThisWorkbook.Worksheets("M_MALZEME_ARA").Shapes("AKSORA_TOPLU_DURUM")
    If sure >= 0 Then metin = metin & vbLf & "Geçen süre: " & Format$(CLng(Fix(sure)) \ 60, "00") & ":" & Format$(CLng(Fix(sure)) Mod 60, "00")
    With s.TextFrame2
        .TextRange.Text = metin
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 9: .TextRange.Font.bold = msoFalse
        .TextRange.Font.Fill.ForeColor.RGB = RGB(85, 103, 123)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoTrue
    End With
End Sub

Public Sub AKSORA_MUSTERI_TOPLU_DURUM(ByVal metin As String)
    mMusteriTopluMetin = metin
    AKSORA_MUSTERI_TOPLU_TIK True
    If AKSORA_MUSTERI_TOPLU_AKTIF Then AKSORA_MUSTERI_SERIT "Güncelleniyor  •  Bütün iller  •  " & metin, 1
    application.StatusBar = "AKSORA | " & metin
    DoEvents
End Sub

Public Sub AKSORA_MUSTERI_TOPLU_TIK(Optional ByVal zorla As Boolean = False)
    Dim sure As Double, saniye As Long
    If Not AKSORA_MUSTERI_TOPLU_AKTIF Then Exit Sub
    sure = Timer - mMusteriTopluBas
    If sure < 0 Then sure = sure + 86400#
    saniye = CLng(Fix(sure))
    If Not zorla And saniye = mMusteriTopluSaniye Then Exit Sub
    TEMA_MUSTERI_SURE_YAZ mMusteriTopluMetin, sure
    mMusteriTopluSaniye = saniye
End Sub

Public Sub AKSORA_MUSTERI_TOPLU_MESAJ()
    If Len(mTopluMesaj) = 0 Then Exit Sub
    MsgBox mTopluMesaj, mTopluIkon, "AKSORA | Müþteri Birimi"
    mTopluMesaj = vbNullString
End Sub

Public Sub AKSORA_MUSTERI_BUTUN_ILLERI_GUNCELLE()
    Dim ws As Worksheet, onceki As Object, il As Variant, secim As Variant
    Dim ev As Boolean, sc As Boolean, hesap As XlCalculation, sessiz As Boolean, korumali As Boolean
    Dim bilgi As String, sonuc As String, hataMetni As String, sure As Double, adet As Long
    If AKSORA_V11_MESGUL Or AKSORA_MUSTERI_TOPLU_AKTIF Then Exit Sub
    If Not AKSORA_ADMIN_TOPLU Then
        If MsgBox("Tanýmlý illerin (Artvin, Trabzon, Gümüþhane, Giresun, Rize) müþteri stoklarý SAP'den güncellenecek. Devam edilsin mi?", _
            vbYesNo + vbQuestion, "AKSORA | Müþteri Birimi") <> vbYes Then Exit Sub
    End If
    ' SAP kapaliysa iller tek tek arsivlenip yuklenmeden once haber verilir (uzun bekleme olmaz).
    Set onceki = SAP_SESSION_AL("CDP", hataMetni)
    If onceki Is Nothing Then
        If AKSORA_ADMIN_TOPLU Then
            AKSORA_ADMIN_SONUC = AKSORA_ADMIN_SONUC & "MÜÞTERÝ: SAP baðlantýsý yok. " & hataMetni & vbCrLf
        Else
            MsgBox "SAP baðlantýsý bulunamadý; müþteri stoklarý güncellenmedi." & vbCrLf & hataMetni, vbExclamation, "AKSORA | Müþteri Birimi"
        End If
        Exit Sub
    End If
    hataMetni = vbNullString
    ev = application.EnableEvents: sc = application.ScreenUpdating
    hesap = application.Calculation: sessiz = ANA_SESSIZ_MOD
    Set onceki = ActiveSheet
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("M_MALZEME_ARA")
    ' Ekrandaki il (girisi yapilmis il) geri yuklenir; secim ekraninda tiklanip girilmeyen il degil.
    secim = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("M_CDP_STOK").Range("BX20").Value2)))
    If Len(secim) = 0 Then secim = ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2
    korumali = ws.ProtectContents
    ws.Unprotect ""
    application.EnableEvents = False
    application.Calculation = xlCalculationManual
    ANA_SESSIZ_MOD = True: AKSORA_V11_MESGUL = True
    mMusteriTopluBas = Timer: mMusteriTopluSaniye = -1
    AKSORA_MUSTERI_TOPLU_AKTIF = True
    ws.Visible = xlSheetVisible
    ' Bütün iller guncellemesi ana panelden izlenir (serit + kartlar).
    AKSORA_MUSTERI_PANEL_ON
    application.ScreenUpdating = True
    AKSORA_MUSTERI_TOPLU_DURUM "Güncelleme hazýrlanýyor..."
    For Each il In AKSORA_MUSTERI_ILLER()
        ' Her il kendi kaydina yuklenip musteri motoruyla (IQ09 + MB52) guncellenir.
        AKSORA_MUSTERI_IL_DEGIS CStr(il)
        bilgi = ""
        AKSORA_MUSTERI_TOPLU_DURUM CStr(il) & " güncelleniyor..."
        If AKSORA_MUSTERI_IL_CALISTIR(bilgi) Then adet = adet + 1
        sonuc = sonuc & bilgi & vbCrLf
        AKSORA_MUSTERI_TOPLU_DURUM bilgi
    Next il
    mMusteriTopluMetin = "Ýþlem tamamlandý. " & adet & "/" & CStr(UBound(AKSORA_MUSTERI_ILLER()) + 1) & " il güncellendi."
Bitir:
    On Error Resume Next
    If AKSORA_MUSTERI_TOPLU_AKTIF Then
        sure = Timer - mMusteriTopluBas
        If sure < 0 Then sure = sure + 86400#
        AKSORA_MUSTERI_TOPLU_AKTIF = False
        If Len(hataMetni) > 0 Then mMusteriTopluMetin = "Ýþlem durduruldu. " & adet & " il güncellendi."
        ThisWorkbook.names.Add Name:="_AKSORA_MUSTERI_SON_SURE", _
            RefersTo:="=" & Trim$(Str$(Round(sure, 2))), Visible:=False
        ThisWorkbook.names.Add Name:="_AKSORA_MUSTERI_SON_DURUM", _
            RefersTo:="=""" & Replace(mMusteriTopluMetin, """", """""") & """", Visible:=False
        TEMA_MUSTERI_SURE_YAZ mMusteriTopluMetin, sure
        ' Baslangictaki il ekrana geri yuklenir.
        AKSORA_MUSTERI_IL_DEGIS CStr(secim)
    End If
    If Not ws Is Nothing And korumali Then ws.Protect Password:="", DrawingObjects:=True, _
        Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = sc
    ANA_SESSIZ_MOD = sessiz: AKSORA_V11_MESGUL = False
    application.StatusBar = False
    AKSORA_MUSTERI_ARAMA_TARIH_YENILE
    AKSORA_MUSTERI_PANEL_GOSTER
    On Error GoTo 0
    ' Ozet, panel ekrana cizildikten sonra gosterilir.
    If Len(hataMetni) > 0 Then
        mTopluMesaj = sonuc & vbCrLf & hataMetni: mTopluIkon = vbExclamation
    Else
        mTopluMesaj = sonuc: mTopluIkon = vbInformation
    End If
    If AKSORA_ADMIN_TOPLU Then AKSORA_ADMIN_SONUC = AKSORA_ADMIN_SONUC & "MÜÞTERÝ:" & vbCrLf & mTopluMesaj & vbCrLf: Exit Sub
    application.OnTime Now + TimeSerial(0, 0, 1), "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_TOPLU_MESAJ"
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Bitir
End Sub
