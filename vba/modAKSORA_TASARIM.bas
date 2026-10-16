Attribute VB_Name = "modAKSORA_TASARIM"
Option Explicit

Private Const etiket As String = "_AKSORA_TASARIM_P1S2R1"
Private Const AKT_PREFIX As String = "AKT1_"
Private mCalisiyor As Boolean
Private mKurulum As Boolean

Public Function AKSORA_TASARIM_AKTIF() As Boolean
    Dim n As Name, ad As String
    For Each n In ThisWorkbook.names
        ad = n.Name
        If InStrRev(ad, "!") > 0 Then ad = Mid$(ad, InStrRev(ad, "!") + 1)
        If StrComp(ad, etiket, vbTextCompare) = 0 Then
            AKSORA_TASARIM_AKTIF = (n.RefersTo = "=1")
            Exit Function
        End If
    Next n
End Function

Public Sub AKSORA_P1_S2_R1_KUR()
    Dim ws As Worksheet, eski As Worksheet, s As Shape, ad As Variant
    Dim veri As Object, bag As Object, form As Object, fil As Object
    Dim a As Variant, b As Variant, k As Variant, son As Long, r As Long, c As Long
    Dim ev As Boolean, sc As Boolean, yedek As String, Hata As String, adim As String, eNo As Long, kaynak As String
    Dim sayfalar As Variant, veriAlan As String, nn As Name, onceAktif As Boolean
    If AKSORA_V11_MESGUL Or AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Then
        MsgBox "Devam eden iþlem bittikten sonra tasarýmý kurun.", vbInformation, "AKSORA"
        Exit Sub
    End If
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    If Len(ThisWorkbook.Path) = 0 Or ThisWorkbook.ReadOnly Then
        Err.Raise vbObjectError + 9300, , "Önce dosyanýn yazýlabilir .xlsm kopyasýný kaydedin."
    End If
    If Len(Dir$(AKSORA_GORSEL("AKSORA_PANEL_SADE_P1.png"))) = 0 Then
        Err.Raise vbObjectError + 9301, , "AKSORA_PANEL_SADE_P1.png Excel dosyasýnýn yanýnda bulunmalýdýr."
    End If
    adim = "Aktif sayfa / tasarým durumu"
    Set eski = ActiveSheet
    onceAktif = AKSORA_TASARIM_AKTIF()
    sayfalar = Array("ANA PANEL", "CDP_STOK", "EPP_STOK", "STOKLAR", _
        "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK", "CDP_RAPOR", "EPP_RAPOR", _
        "RAPORLAR", "ÝLÇE_RAPORLAR", "YÜKLENÝCÝ_RAPOR")
    Set veri = CreateObject("Scripting.Dictionary")
    Set bag = CreateObject("Scripting.Dictionary")
    Set form = CreateObject("Scripting.Dictionary")
    Set fil = CreateObject("Scripting.Dictionary")
    ' Kurulumdan once butonlar ve veri bellekte birebir kaydedilir.
    For Each ad In sayfalar
        adim = "Sayfa ön kontrolü: " & CStr(ad)
        Set ws = T_Sayfa(CStr(ad))
        If ws Is Nothing Then Err.Raise vbObjectError + 9302, , "Eksik sayfa: " & CStr(ad)
        For Each s In ws.Shapes
            adim = "Buton ön kontrolü: " & ws.Name & " / " & s.Name
            If s.Type <> msoChart And s.Type <> msoFormControl And s.Type <> msoOLEControlObject Then
                If Len(s.OnAction) > 0 Then bag(ws.Name & "|" & s.Name) = s.OnAction
            End If
        Next s
        If T_Stok(ws) Then
            son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
            If son < 7 Then son = 7
            veriAlan = "A7:L" & son
            adim = "Stok verisini güvenceye alma: " & ws.Name & " / " & veriAlan
            veri(ws.Name & "|" & veriAlan) = ws.Range(veriAlan).Value2
            adim = "KPI ön kontrolü: " & ws.Name
            form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
            adim = "Filtre ön kontrolü: " & ws.Name
            fil(ws.Name) = T_FiltreImza(ws)
        ElseIf T_Rapor(ws) Then
            adim = "Rapor verisini güvenceye alma: " & ws.Name
            son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            If son < 8 Then son = 8
            veriAlan = "A8:E" & son
            veri(ws.Name & "|" & veriAlan) = ws.Range(veriAlan).Value2
        End If
    Next ad
    application.EnableEvents = False: application.ScreenUpdating = False
    yedek = AKSORA_KAYIT_KLASORU() & application.PathSeparator & "AKSORA_TASARIM_ONCESI_" & _
        Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    If Len(Dir$(yedek)) > 0 Then Err.Raise vbObjectError + 9303, , "Ayný saniyede ikinci kurulum yapýlmaz."
    adim = "Kurulum öncesi yedek dosyayý kaydetme"
    ThisWorkbook.SaveCopyAs yedek
    mKurulum = True
    adim = "Tasarým durum kaydýný oluþturma"
    ThisWorkbook.names.Add Name:=etiket, RefersTo:="=1", Visible:=False
    For Each ad In sayfalar
        adim = "Tasarýmý uygulama: " & CStr(ad)
        Set ws = T_Sayfa(CStr(ad))
        If T_Rapor(ws) Then
            AKSORA_TASARIM_SAYFA ws, True
            adim = "Rapor grafiði: " & ws.Name
            AKSORA_RAPOR_GRAFIK_YENILE ws.Name
            AKSORA_RAPOR_ALT_YAZI_YERLESTIR ws.Name
            AKSORA_RAPOR_TABLO_BITISI ws.Name
        End If
        AKSORA_TASARIM_SAYFA ws, True
    Next ad
    ' Tum stok ve rapor hucreleri tek tek karsilastirilir; SAP'a baglanilmaz.
    adim = "Kurulum sonrasý veri karþýlaþtýrmasý"
    For Each k In veri.keys
        a = veri(k): Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        b = ws.Range(Split(CStr(k), "|")(1)).Value2
        For r = 1 To UBound(a, 1)
            For c = 1 To UBound(a, 2)
                If Not T_Ayni(a(r, c), b(r, c)) Then
                    Err.Raise vbObjectError + 9304, , "Veri deðiþikliði saptandý: " & ws.Name & ", hücre " & r & "/" & c
                End If
            Next c
        Next r
    Next k
    adim = "Kurulum sonrasý buton karþýlaþtýrmasý"
    For Each k In bag.keys
        Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        Set s = ws.Shapes(Split(CStr(k), "|")(1))
        If s.OnAction <> bag(k) Then Err.Raise vbObjectError + 9305, , "Buton baðlantýsý deðiþti: " & CStr(k)
        If Not s.Visible Then Err.Raise vbObjectError + 9306, , "Ýþlev butonu gizlendi: " & CStr(k)
    Next k
    adim = "Kurulum sonrasý KPI / filtre karþýlaþtýrmasý"
    For Each k In form.keys
        Set ws = ThisWorkbook.Worksheets(CStr(k)): a = form(k)
        If CStr(ws.Range("C3").Formula) <> CStr(a(0)) Or CStr(ws.Range("C4").Formula) <> CStr(a(1)) Then
            Err.Raise vbObjectError + 9307, , "KPI formülü deðiþti: " & ws.Name
        End If
        If Not T_Ayni(ws.Range("H2").Value2, a(2)) Then Err.Raise vbObjectError + 9308, , "Depo seçimi deðiþti."
        If T_FiltreImza(ws) <> fil(k) Then Err.Raise vbObjectError + 9309, , "Filtre deðiþti: " & ws.Name
    Next k
    adim = "ANA PANEL görünümü / ekran sýðdýrma"
    Set ws = T_Sayfa("ANA PANEL")
    ws.Activate
    AKSORA_V13_GORUNUM ws
    mKurulum = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    MsgBox "P1 / S2 / R1 tasarýmý kuruldu." & vbCrLf & _
        "Stok ve rapor verileri, KPI formülleri, filtre seçimi ve mevcut makro baðlantýlarý deðiþmedi." & vbCrLf & _
        "Kaydetmeden önce buton, depo filtresi ve ekran görünümünü deneyin." & vbCrLf & _
        "Kurulum öncesi kopya: " & yedek, vbInformation, "AKSORA"
    Exit Sub
Hata:
    eNo = Err.Number: kaynak = Err.Source
    Hata = "Adým: " & adim & vbCrLf & "Hata: " & eNo & " / " & Err.Description & vbCrLf & "Kaynak: " & kaynak
    On Error Resume Next
    mKurulum = False: mCalisiyor = False
    If Not onceAktif Then
        For Each nn In ThisWorkbook.names
            If StrComp(Right$(nn.Name, Len(etiket)), etiket, vbTextCompare) = 0 Then
                nn.Delete
                Exit For
            End If
        Next nn
    End If
    If Not eski Is Nothing Then eski.Activate
    application.EnableEvents = ev: application.ScreenUpdating = sc
    On Error GoTo 0
    MsgBox "Tasarým kurulumu tamamlanamadý: " & Hata & vbCrLf & _
        "Bu denemeyi kaydetmeyin. Kurulum öncesi kopyayý açýn." & vbCrLf & yedek, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_TASARIM_SAYFA(ByVal ws As Worksheet, Optional ByVal yenile As Boolean = False)
    Dim p As Variant, koruma As Boolean, ev As Boolean, sc As Boolean
    Dim no As Long, metin As String
    If mCalisiyor Or Not AKSORA_TASARIM_AKTIF() Then Exit Sub
    If ws Is Nothing Then Exit Sub
    If Not (T_Stok(ws) Or T_Rapor(ws) Or ws.Name = "ANA PANEL") Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    mCalisiyor = True
    application.EnableEvents = False: application.ScreenUpdating = False
    T_KorumaAc ws, p, koruma
    If ws.Name = "ANA PANEL" Then
        T_Panel ws
    ElseIf T_Stok(ws) Then
        T_StokYap ws, yenile
    Else
        T_RaporYap ws
    End If
Cikis:
    On Error Resume Next
    If koruma Then
        Err.Clear
        T_KorumaKapat ws, p
        If Err.Number <> 0 And no = 0 Then no = Err.Number: metin = Err.Description
    End If
    application.EnableEvents = ev: application.ScreenUpdating = sc
    mCalisiyor = False
    On Error GoTo 0
    If no <> 0 Then Err.Raise no, "AKSORA_TASARIM_SAYFA", metin
    Exit Sub
Hata:
    no = Err.Number: metin = ws.Name & ": " & Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_TASARIM_PENCERE(ByVal ws As Worksheet)
    Dim wn As Window
    If Not AKSORA_TASARIM_AKTIF() Or Not T_Stok(ws) Then Exit Sub
    If Not (ActiveWorkbook Is ThisWorkbook) Or Not (ActiveSheet Is ws) Then Exit Sub
    Set wn = ActiveWindow
    If AKSORA_PENCERE_KUCUK(wn) Then Exit Sub
    On Error Resume Next
    ' SplitRow kullanilir; hucre secimini degistirmez.
    With wn
        If Not .FreezePanes Or .SplitRow <> 6 Or .SplitColumn <> 0 Then
            .FreezePanes = False: .SplitRow = 6: .SplitColumn = 0: .FreezePanes = True
        End If
    End With
End Sub

Private Sub T_Panel(ByVal ws As Worksheet)
    Dim s As Shape, eski As Shape, r As Range, f As String, il As String
    Dim y As Double, w As Double, fs As Single, n As String
    Set r = ws.Range("B7:AB27")
    Set eski = T_Sekil(ws, "AKSORA15_TEK_ARKA_PLAN")
    If eski Is Nothing Then Err.Raise vbObjectError + 9310, , "Panelin mevcut arka planý bulunamadý."
    If eski.AlternativeText <> "AKSORA_P1_SADE_V1" Then
        f = AKSORA_GORSEL("AKSORA_PANEL_SADE_P1.png")
        If Len(Dir$(f)) = 0 Then Err.Raise vbObjectError + 9311, , "Sade panel görseli bulunamadý."
        Set s = ws.Shapes.AddPicture(f, msoFalse, msoTrue, r.Left, r.Top, r.Width, r.Height)
        s.Name = AKT_PREFIX & "PANEL_GECICI"
        eski.Delete
        s.Name = "AKSORA15_TEK_ARKA_PLAN"
        s.AlternativeText = "AKSORA_P1_SADE_V1"
    Else
        Set s = eski
    End If
    T_Yerlestir s, r.Left, r.Top, r.Width, r.Height
    s.ZOrder msoSendToBack
    ' Bes il icin gomulu, tek parca baslik; etiket veya logo parcasi yok.
    Set r = ws.Range("B2:AB5")
    T_PanelTekBanner ws, r
    T_PanelHizala ws
    For Each s In ws.Shapes
        n = LCase$(s.Name)
        If s.Type <> msoChart And s.Type <> msoFormControl And s.Type <> msoOLEControlObject Then
        If Len(s.OnAction) > 0 Then
            If n = "btntumstok" Then
                T_Buton s, RGB(0, 110, 94), vbWhite, 14
            ElseIf Left$(n, 2) = "go" Or InStr(n, "baglanti") > 0 Or InStr(n, "malzeme") > 0 Then
                T_Buton s, RGB(244, 251, 255), RGB(16, 57, 94), 9
            ElseIf n <> "aksora15_il_degistir" Then
                T_Buton s, RGB(16, 57, 94), vbWhite, 10
            End If
            s.ZOrder msoBringToFront
        End If
        End If
    Next s
End Sub

Private Sub T_PanelParca(ByVal ws As Worksheet, ByVal ad As String, ByVal yol As String, _
    ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
    ByVal solOran As Double, ByVal sagOran As Double)
    Dim s As Shape, w0 As Double, h0 As Double
    Set s = T_Sekil(ws, ad)
    If Not s Is Nothing Then
        If s.AlternativeText <> yol Then s.Delete: Set s = Nothing
    End If
    If s Is Nothing Then
        Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
        s.Name = ad: s.AlternativeText = yol
        w0 = s.Width: h0 = s.Height
        With s.PictureFormat
            If w0 / h0 < 4 Then
                .CropTop = h0 * 225 / 684
                .CropBottom = h0 * 224 / 684
            End If
            .CropLeft = w0 * solOran
            .CropRight = w0 * sagOran
        End With
    End If
    T_Yerlestir s, x, y, w, h
    s.ZOrder msoBringToFront
End Sub

Private Sub T_FiltreOku(ByVal ws As Worksheet)
    Dim i As Long, n As String, s As Shape
    ' Yalniz ekledigimiz deneme oklarini kaldir. Yerel H2 dogrulamasi korunur.
    For i = ws.Shapes.Count To 1 Step -1
        Set s = ws.Shapes(i): n = s.Name
        If n = AKT_PREFIX & "DEPO_OK_ZEMIN" Or n = AKT_PREFIX & "DEPO_OK_UC" Then
            s.Delete
        ElseIf s.Type = msoFormControl Or s.Type = msoOLEControlObject Then
            If UCase$(Replace(n, " ", "")) = "DROPDOWN1" Then s.Visible = msoFalse
        End If
    Next i
End Sub

Private Sub T_StokYap(ByVal ws As Worksheet, ByVal yenile As Boolean)
    Dim w As Double, x As Double, y As Double, h As Double, son As Long, s As Shape
    Dim k As Long, t As String, degerler As Variant, basliklar As Variant, dugmeler As Variant
    Dim sag As Double, bw As Double, adet As Long, ad As String
    Dim depoListe As String, depoKontrol As Object, di As Long
    ' Veri baslangici 7. satirda kalir. Gizlenen 3-5 satirlardaki kaynaklar korunur.
    ' Tablo pencereye tam oturur (son sutun pencere kenarina kadar uzar).
    T_TamGenislik ws, 12
    ws.Rows(1).RowHeight = 32: ws.Rows(2).RowHeight = 46
    ws.Rows("3:5").Hidden = True: ws.Rows(6).RowHeight = 25
    ws.Range("A1:L2").Interior.Color = RGB(237, 248, 255)
    ws.Range("A2:L2").Borders.LineStyle = xlNone
    ws.Range("A2:L2").Interior.Color = RGB(235, 243, 249)
    ws.Range("A1:L1").Font.Color = RGB(12, 45, 79)
    ws.Range("C2").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Range("C3").NumberFormat = "0"
    ws.Range("C4").NumberFormat = "#,##0.00 ""TL"""
    ws.Range("C2:C4").Font.Color = RGB(26, 57, 88)
    son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    If son >= 7 Then AKSORA_V13_MIKTAR_BICIMI ws.Range("I7:I" & son)
    T_StokListeBicim ws, son
    With ws.Range("A6:L6")
        .Interior.Color = RGB(26, 75, 120): .Font.Color = vbWhite
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = False: .ShrinkToFit = True
    End With
    ' Musteri stok sayfasindaki gibi acik basliklar.
    basliklar = Array("Üretim Yeri", "Depo", "Depo Adý", "Mal Grubu", "Mal Grubu Tanýmý", "Deðerleme Türü", _
        "Malzeme", "Malzeme Tanýmý", "Miktar", "Birim", "Deðer (TL)", "Birim Fiyat (TL)")
    For k = 0 To 11
        If CStr(ws.Cells(6, k + 1).Value2) <> CStr(basliklar(k)) Then ws.Cells(6, k + 1).Value2 = basliklar(k)
    Next k
    w = ws.Range("A1:L1").Width: y = ws.Rows(1).Top
    Set s = T_Dik(ws, AKT_PREFIX & "STOK_BASLIK_ZEMIN", 0, y, w, 32, RGB(12, 45, 79))
    Set s = T_Yazi(ws, AKT_PREFIX & "STOK_MARKA", "AKSORA", 12, y + 3, 102, 26, 20, vbWhite, True, msoAlignLeft)
    Set s = T_Yazi(ws, AKT_PREFIX & "STOK_SLOGAN", "Stok ve Verimlilik", 118, y + 5, 135, 22, 10, RGB(198, 240, 97), True, msoAlignLeft)
    t = Replace(ws.Name, "_", " ") & "  |  " & T_Il() & " Ýl Müdürlüðü"
    Set s = T_Yazi(ws, AKT_PREFIX & "STOK_BASLIK", t, 260, y + 5, w - 680, 22, 11, vbWhite, True, msoAlignLeft)
    degerler = Array("C2", "C3", "C4")
    basliklar = Array("Son Güncelleme", "Toplam Stok Satýrý", "Toplam Stok Deðeri")
    h = ws.Rows(2).Height: y = ws.Rows(2).Top
    ' Kaynak hucreler ve formule bagli canli degerler korunur.
    ' Kesintisiz zemin, kart aralarinda C2 tarihinin sizmasini engeller.
    Set s = T_Dik(ws, AKT_PREFIX & "STOK_KPI_ALT_ZEMIN", 0, y, w, h, RGB(230, 240, 247))
    ' Numara bloklari bu tasarimda kullanilmaz.
    For k = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(k).Name, Len(AKT_PREFIX & "STOK_KPI_NO")) = AKT_PREFIX & "STOK_KPI_NO" Then ws.Shapes(k).Delete
    Next k
    ' Kartlar tablo genisligine esit yayilir (Stoklar: 3 kart, digerleri: 3 kart + depo karti).
    bw = w / IIf(AKFIX_NORMAL(ws.Name) = "STOKLAR", 3, 4)
    For k = 0 To 2
        x = k * bw
        Set s = T_Dik(ws, AKT_PREFIX & "STOK_KPI_ZEMIN_" & k, x + 2, y + 3, bw - 5, h - 6, RGB(232, 243, 252))
        s.AutoShapeType = msoShapeRoundedRectangle: s.Adjustments.Item(1) = 0.06
        s.Fill.BackColor.RGB = RGB(249, 252, 255)
        s.Fill.TwoColorGradient msoGradientHorizontal, 1
        s.line.Visible = msoTrue: s.line.ForeColor.RGB = RGB(103, 161, 205): s.line.Weight = 1.2
        s.Shadow.Visible = msoFalse
        Set s = T_Dik(ws, AKT_PREFIX & "STOK_KPI_SIMGE_Z_" & k, x + 9, y + 8, 33, 30, RGB(202, 225, 244))
        s.AutoShapeType = msoShapeRoundedRectangle: s.Adjustments.Item(1) = 0.12
        T_Ikon ws, CStr(k), x + 14, y + 11, RGB(25, 107, 175)
        Set s = T_Yazi(ws, AKT_PREFIX & "STOK_KPI_ETIKET_" & k, CStr(basliklar(k)), _
            x + 51, y + 5, bw - 60, 15, 8.5, RGB(39, 76, 106), True, msoAlignLeft)
        Set s = T_Yazi(ws, AKT_PREFIX & "STOK_KPI_DEGER_" & k, vbNullString, _
            x + 51, y + 23, bw - 60, 17, 10.5, RGB(18, 49, 80), True, msoAlignLeft)
        s.DrawingObject.Formula = "='" & Replace(ws.Name, "'", "''") & "'!" & CStr(degerler(k))
        s.TextFrame.Characters.Font.Color = RGB(18, 49, 80)
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(18, 49, 80)
        s.TextFrame2.TextRange.Font.bold = msoTrue
        s.TextFrame2.TextRange.Font.Size = 10.5
    Next k
    If AKFIX_NORMAL(ws.Name) <> "STOKLAR" Then
        ' Diger KPI kartlariyla ayni sekil, degrade, baslik ve deger duzeni.
        depoListe = ws.Range("H2").Validation.Formula1
        If Left$(depoListe, 1) = "=" Then depoListe = Mid$(depoListe, 2)
        ws.Range("H2").Validation.InCellDropdown = False
        ws.Range("H2").NumberFormat = ";;;"
        ws.Range("G2").Font.Color = RGB(235, 243, 249)
        bw = w / 4: x = 3 * bw
        For k = ws.Shapes.Count To 1 Step -1
            If Left$(ws.Shapes(k).Name, Len(AKT_PREFIX & "DEPO_KART_")) = AKT_PREFIX & "DEPO_KART_" Then ws.Shapes(k).Delete
        Next k
        Set s = T_Dik(ws, AKT_PREFIX & "DEPO_KART_ZEMIN", x + 2, y + 3, bw - 5, h - 6, RGB(232, 243, 252))
        s.AutoShapeType = msoShapeRoundedRectangle: s.Adjustments.Item(1) = 0.06
        s.Fill.BackColor.RGB = RGB(249, 252, 255)
        s.Fill.TwoColorGradient msoGradientHorizontal, 1
        s.line.Visible = msoTrue: s.line.ForeColor.RGB = RGB(103, 161, 205): s.line.Weight = 1.2
        s.Shadow.Visible = msoFalse
        Set s = T_Dik(ws, AKT_PREFIX & "DEPO_KART_SIMGE_Z", x + 9, y + 8, 33, 30, RGB(202, 225, 244))
        s.AutoShapeType = msoShapeRoundedRectangle: s.Adjustments.Item(1) = 0.12
        ' Depo simgesi: cati ve uc raf, yerel sekillerle cizilir.
        Set s = T_Dik(ws, AKT_PREFIX & "DEPO_KART_SIMGE_GOVDE", x + 16, y + 19, 19, 15, RGB(25, 107, 175))
        Set s = T_Dik(ws, AKT_PREFIX & "DEPO_KART_SIMGE_CATI", x + 14, y + 12, 23, 9, RGB(25, 107, 175))
        s.AutoShapeType = msoShapeIsoscelesTriangle
        For k = 0 To 2
            Set s = T_Dik(ws, AKT_PREFIX & "DEPO_KART_SIMGE_RAF_" & k, x + 19 + k * 5, y + 25, 3, 6, RGB(232, 243, 252))
        Next k
        Set s = T_Yazi(ws, AKT_PREFIX & "DEPO_KART_BASLIK", "DEPO:", x + 51, y + 5, bw - 60, 15, 8.5, RGB(39, 76, 106), True, msoAlignLeft)
        Set depoKontrol = ws.Shapes.AddFormControl(xlDropDown, x + 51, y + 23, bw - 60, 17)
        depoKontrol.Name = AKT_PREFIX & "DEPO_KART_LISTE"
        depoKontrol.ControlFormat.ListFillRange = depoListe
        depoKontrol.ControlFormat.LinkedCell = vbNullString
        depoKontrol.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_DEPO_KART_SEC"
        For di = 1 To depoKontrol.ControlFormat.ListCount
            If CStr(depoKontrol.ControlFormat.List(di)) = CStr(ws.Range("H2").Value2) Then
                depoKontrol.ControlFormat.Value = di
                Exit For
            End If
        Next di
        T_DepoListeYukseklik depoKontrol, ws.Name
        depoKontrol.Placement = xlFreeFloating
        If depoKontrol.ControlFormat.ListCount < 1 Then Err.Raise vbObjectError + 9521, , "Depo listesi bos: " & ws.Name
        If Len(CStr(ws.Range("H2").Value2)) > 0 Then
            di = depoKontrol.ControlFormat.Value
            If di < 1 Then Err.Raise vbObjectError + 9522, , "Secili depo listeyle eslesmedi: " & ws.Name
            If CStr(depoKontrol.ControlFormat.List(di)) <> CStr(ws.Range("H2").Value2) Then _
                Err.Raise vbObjectError + 9523, , "Depo secimi dogrulanamadi: " & ws.Name
        End If
        T_FiltreOku ws
    End If
    Set s = T_Sekil(ws, "AKSORA_STOK_SON_SEFFAF_LOGO")
    If Not s Is Nothing Then s.Visible = msoFalse
    Select Case AKFIX_NORMAL(ws.Name)
        Case "CDP_STOK", "EPP_STOK": dugmeler = Array("SAP_STOK_GUNCELLE_BTN")
        Case "STOKLAR": dugmeler = Array("STOKLAR_GUNCELLE_BTN")
        Case "ILCE_STOK": dugmeler = Array("CDP_ILCE_STOK_BTN", "EPP_ILCE_STOK_BTN")
        Case "YUKLENICI_STOK": dugmeler = Array("YUK_EDAS_GUNCELLE_BTN", "YUK_AKSADA_GUNCELLE_BTN")
    End Select
    sag = w - 6
    Set s = T_Sekil(ws, "btnAnaPanelDon")
    If s Is Nothing Then Err.Raise vbObjectError + 9312, , "ANA PANEL butonu eksik."
    T_Yerlestir s, sag - 88, ws.Rows(1).Top + 5, 88, 22
    T_Buton s, RGB(16, 57, 94), vbWhite, 9
    s.TextFrame2.TextRange.Text = "ANA PANEL"
    sag = sag - 94: adet = UBound(dugmeler) + 1
    bw = IIf(adet = 1, 170, 140)
    For k = UBound(dugmeler) To 0 Step -1
        ad = CStr(dugmeler(k)): Set s = T_Sekil(ws, ad)
        If s Is Nothing Then Err.Raise vbObjectError + 9313, , "Güncelleme butonu eksik: " & ad
        T_Yerlestir s, sag - bw, ws.Rows(1).Top + 5, bw, 22
        T_Buton s, RGB(26, 83, 134), vbWhite, 8
        ' Tum stok sayfasi dugmeleri ayni yoldan: hizli okuma, ayni sayfaya donus, tek ozet.
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_STOK_SAYFA_GUNCELLE"
        sag = sag - bw - 5
    Next k
    ' Sayim ciktisi dugmesi (guncelleme dugmelerinin solunda).
    Set s = T_Sekil(ws, "btnStokCikti")
    If s Is Nothing Then
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 96, 22): s.Name = "btnStokCikti"
    End If
    T_Yerlestir s, sag - 96, ws.Rows(1).Top + 5, 96, 22
    T_Buton s, RGB(15, 118, 110), vbWhite, 8
    s.TextFrame2.TextRange.Text = "ÇIKTI AL"
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_CIKTI_AL"
    sag = sag - 96 - 5
    Set s = T_Sekil(ws, AKT_PREFIX & "STOK_BASLIK")
    If sag > 280 Then s.Width = sag - 270 Else s.Visible = msoFalse
    ' Ust satirlar disinda gizli veri satirlari ve filtreler degistirilmez.
End Sub

Private Sub T_RaporYap(ByVal ws As Worksheet)
    Dim p As String, s As Shape, r As Range, ad As Variant, alanlar As Variant
    Dim i As Long, x As Double, y As Double, w As Double, h As Double
    Dim kod As String, son As Long, c As Long
    If AKFIX_NORMAL(ws.Name) = "CDP_RAPOR" Then p = "AKSORA_CDP_V1_" Else p = "AKSORA_RP_"
    T_TamGenislik ws, 20
    ws.Rows("1:2").RowHeight = 19: ws.Rows("3:6").RowHeight = 17
    ws.Rows(7).RowHeight = 8: ws.Rows(8).RowHeight = 30
    Set r = ws.Range("A1:T2")
    Set s = T_Dik(ws, AKT_PREFIX & "RAPOR_UST_ZEMIN", r.Left, r.Top, r.Width, r.Height, RGB(12, 45, 79))
    Select Case AKFIX_NORMAL(ws.Name)
        Case "CDP_RAPOR": kod = "CDP STOK RAPORU"
        Case "EPP_RAPOR": kod = "EPP STOK RAPORU"
        Case "RAPORLAR": kod = "GENEL STOK RAPORU"
        Case "ILCE_RAPORLAR": kod = "ÝLÇE STOK RAPORU"
        Case "YUKLENICI_RAPOR": kod = "YÜKLENÝCÝ STOK RAPORU"
    End Select
    Set s = T_Yazi(ws, AKT_PREFIX & "RAPOR_BASLIK", "AKSORA | " & kod, _
        r.Left + 16, r.Top + 5, r.Width - 155, r.Height - 10, 20, vbWhite, True, msoAlignLeft)
    Set s = T_Sekil(ws, p & "BANNER"): If Not s Is Nothing Then s.Visible = msoFalse
    Set s = T_Sekil(ws, p & "BASLIK"): If Not s Is Nothing Then s.Visible = msoFalse
    Set s = T_Sekil(ws, "btnAnaPanelDon")
    If s Is Nothing Then Err.Raise vbObjectError + 9314, , "Rapor ANA PANEL butonu eksik."
    T_Yerlestir s, r.Left + r.Width - 108, r.Top + 7, 96, r.Height - 14
    T_Buton s, RGB(16, 57, 94), vbWhite, 9
    s.TextFrame2.TextRange.Text = "ANA PANEL"
    alanlar = Array("A3:E6", "F3:J6", "K3:Q6")
    For i = 0 To 2
        ad = alanlar(i): Set r = ws.Range(CStr(ad))
        x = r.Left + 3: y = r.Top + 3: w = r.Width - 6: h = r.Height - 6
        If i = 2 Then w = ws.Range("T6").Left + ws.Range("T6").Width - x - 3
        Set s = T_Sekil(ws, p & "KART_" & Replace(CStr(ad), ":", "_"))
        If s Is Nothing Then Err.Raise vbObjectError + 9315, , "Mevcut rapor KPI kartý eksik."
        T_Yerlestir s, x, y, w, h
        s.Fill.ForeColor.RGB = RGB(242, 249, 255)
        s.line.ForeColor.RGB = RGB(200, 225, 245): s.line.Weight = 0.8
        Set s = T_Sekil(ws, p & "IKON_" & Replace(CStr(ad), ":", "_"))
        If s Is Nothing Then Set s = T_Sekil(ws, p & "NO_" & Replace(CStr(ad), ":", "_"))
        T_Yerlestir s, x + 8, y + 8, 40, h - 16
        Set s = T_Sekil(ws, p & "ETIKET_" & Replace(CStr(ad), ":", "_"))
        T_Yerlestir s, x + 61, y + 6, w - 69, 19
        s.TextFrame2.TextRange.Font.Size = 10
        Set s = T_Sekil(ws, p & "DEGER_" & Replace(CStr(ad), ":", "_"))
        T_Yerlestir s, x + 61, y + 29, w - 69, h - 32
        s.TextFrame2.TextRange.Font.Size = 13
    Next i
    son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    For i = 9 To son
        If AKFIX_NORMAL(CStr(ws.Cells(i, 1).Value2)) = "TOPLAM" Then son = i: Exit For
    Next i
    If son >= 9 Then
        ws.Rows("9:" & son).RowHeight = 22
        c = IIf(AKFIX_NORMAL(ws.Name) = "ILCE_RAPORLAR" Or AKFIX_NORMAL(ws.Name) = "YUKLENICI_RAPOR", 5, 4)
        With ws.Range(ws.Cells(8, 1), ws.Cells(son, c))
            .Font.Name = "Aptos": .Font.Size = 10: .VerticalAlignment = xlCenter
        End With
    End If
    T_RaporGrafikSinirla ws
    T_RaporLogo ws
    ' Musteri rapor duzeni: tablo basliklari, kartlar tablo ve grafik hizasinda, sag kenara kadar.
    For i = 1 To 5
        Select Case Trim$(CStr(ws.Cells(8, i).Value2))
            Case "Depo yeri": ws.Cells(8, i).Value2 = "Depo"
            Case "Depo yeri tanýmý": ws.Cells(8, i).Value2 = "Depo Adý"
        End Select
    Next i
    AKSORA_RAPOR_HIZALA ws, ws.Range("A1:T1").Width
End Sub

' Son sutun, pencere genisligi tam dolacak kadar genisletilir; zoom tam sayiya ayarlanir.
Private Sub T_TamGenislik(ByVal ws As Worksheet, ByVal sonKol As Long)
    Dim wn As Window, kol As Range, ad As String, taban As Double, w As Double, z As Double, ek As Double, t As String
    On Error GoTo Cikis
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    Set wn = ActiveWindow
    If AKSORA_PENCERE_KUCUK(wn) Then Exit Sub
    Set kol = ws.Columns(sonKol)
    ad = "_AKS_TABAN_" & ws.CodeName
    On Error Resume Next
    t = ThisWorkbook.names(ad).RefersTo
    On Error GoTo Cikis
    If Len(t) > 1 Then
        taban = Val(Mid$(t, 2))
    Else
        taban = kol.ColumnWidth
        ThisWorkbook.names.Add Name:=ad, RefersTo:="=" & Trim$(Str$(taban)), Visible:=False
    End If
    If taban <= 0 Then Exit Sub
    If kol.ColumnWidth <> taban Then kol.ColumnWidth = taban
    w = ws.Range(ws.Cells(1, 1), ws.Cells(1, sonKol)).Width
    If w <= 0 Then Exit Sub
    z = Fix((wn.UsableWidth - 1) / w * 100)
    If z < 10 Then z = 10
    If z > 400 Then z = 400
    If wn.Zoom <> z Then wn.Zoom = z
    ek = wn.UsableWidth * 100 / z - w - 1
    If ek > 0.5 And kol.Width > 0 Then
        If kol.ColumnWidth * (kol.Width + ek) / kol.Width < 255 Then kol.ColumnWidth = kol.ColumnWidth * (kol.Width + ek) / kol.Width
    End If
Cikis:
End Sub

Public Sub AKSORA_TASARIM_RAPOR_LOGO(ByVal ws As Worksheet)
    Dim p As Variant, koruma As Boolean, no As Long, metin As String
    On Error GoTo Hata
    T_KorumaAc ws, p, koruma
    T_RaporLogo ws
Cikis:
    On Error Resume Next
    If koruma Then
        Err.Clear
        T_KorumaKapat ws, p
        If Err.Number <> 0 And no = 0 Then no = Err.Number: metin = Err.Description
    End If
    On Error GoTo 0
    If no <> 0 Then Err.Raise no, "AKSORA_TASARIM_RAPOR_LOGO", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description: Resume Cikis
End Sub

Private Sub T_RaporLogo(ByVal ws As Worksheet)
    Dim s As Shape, z As Shape, x As Double, y As Double, ad As String
    If AKFIX_NORMAL(ws.Name) = "CDP_RAPOR" Then ad = "AKSORA_CDP_V1_FILIGRAN" Else ad = "AKSORA_RP_FILIGRAN"
    Set s = T_Sekil(ws, ad)
    If s Is Nothing Then Exit Sub
    Set z = T_Sekil(ws, "AKSORA_RP_HALKA_ZEMIN")
    If z Is Nothing Then y = ws.Range("T25").Top Else y = z.Top + z.Height + 6
    x = ws.Range("T1").Left + ws.Range("T1").Width
    With s
        On Error Resume Next
        .LockAspectRatio = msoFalse
        .ScaleHeight 1, msoTrue: .ScaleWidth 1, msoTrue
        On Error GoTo 0
        .LockAspectRatio = msoTrue: .Width = 58
        If .Height > 58 Then .Height = 58
        .Left = x - .Width - 4: .Top = y
        .Placement = xlFreeFloating: .Locked = True: .Visible = msoTrue
        .ZOrder msoBringToFront
    End With
End Sub

Private Function T_Dik(ByVal ws As Worksheet, ByVal ad As String, ByVal x As Double, _
    ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal renk As Long) As Shape
    Dim s As Shape
    Set s = T_Sekil(ws, ad)
    If s Is Nothing Then
        Set s = ws.Shapes.AddShape(msoShapeRectangle, x, y, w, h): s.Name = ad
    End If
    T_Yerlestir s, x, y, w, h
    s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = renk
    s.Fill.Transparency = 0: s.line.Visible = msoFalse
    s.ZOrder msoBringToFront
    Set T_Dik = s
End Function

Private Function T_Yazi(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
    ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
    ByVal fs As Single, ByVal renk As Long, ByVal kalin As Boolean, ByVal hiz As Long) As Shape
    Dim s As Shape
    Set s = T_Sekil(ws, ad)
    If s Is Nothing Then
        Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h): s.Name = ad
    End If
    T_Yerlestir s, x, y, w, h
    s.Fill.Visible = msoFalse: s.line.Visible = msoFalse
    With s.TextFrame2
        .AutoSize = msoAutoSizeNone: .WordWrap = msoFalse
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        If Len(metin) > 0 Then .TextRange.Text = metin
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = fs
        .TextRange.Font.bold = IIf(kalin, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = renk
        .TextRange.ParagraphFormat.Alignment = hiz
    End With
    s.ZOrder msoBringToFront
    Set T_Yazi = s
End Function

Private Sub T_Yerlestir(ByVal s As Shape, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    If s Is Nothing Then Err.Raise vbObjectError + 9316, , "Tasarým için gereken þekil bulunamadý."
    If w < 1 Or h < 1 Then Err.Raise vbObjectError + 9317, , "Tasarým alaný yetersiz."
    With s
        .LockAspectRatio = msoFalse: .Left = x: .Top = y: .Width = w: .Height = h
        .Placement = xlFreeFloating: .Locked = True: .Visible = msoTrue
    End With
End Sub

Private Sub T_Buton(ByVal s As Shape, ByVal dolgu As Long, ByVal yazi As Long, ByVal fs As Single)
    ' OnAction baglantisi asla yazilmaz veya degistirilmez.
    With s
        .Fill.Visible = msoTrue: .Fill.Solid: .Fill.ForeColor.RGB = dolgu
        .Fill.Transparency = 0: .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(116, 192, 226): .line.Weight = 1
        .Shadow.Visible = msoFalse
        With .TextFrame2
            .MarginLeft = 3: .MarginRight = 3: .MarginTop = 0: .MarginBottom = 0
            .AutoSize = msoAutoSizeNone: .WordWrap = msoTrue
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = fs
            .TextRange.Font.bold = msoTrue: .TextRange.Font.Fill.ForeColor.RGB = yazi
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        .Placement = xlFreeFloating: .Visible = msoTrue: .Locked = True
        .ZOrder msoBringToFront
    End With
End Sub

Private Sub T_Ikon(ByVal ws As Worksheet, ByVal kod As String, ByVal x As Double, ByVal y As Double, ByVal renk As Long)
    Dim s As Shape, i As Long, j As Long, ad As String
    ' Excel'in yerel sekilleri; emoji ve sembol yazi tipi kullanilmaz.
    If kod = "0" Then
        Set s = T_Dik(ws, AKT_PREFIX & "IKON_TAKVIM", x, y, 23, 23, renk)
        Set s = T_Dik(ws, AKT_PREFIX & "IKON_TAKVIM_IC", x + 3, y + 7, 17, 13, vbWhite)
        For i = 0 To 2
            For j = 0 To 1
                Set s = T_Dik(ws, AKT_PREFIX & "IKON_TAKVIM_" & i & j, x + 5 + i * 5, y + 10 + j * 5, 3, 3, renk)
            Next j
        Next i
    ElseIf kod = "1" Then
        For i = 0 To 2
            Set s = T_Dik(ws, AKT_PREFIX & "IKON_SATIR_N_" & i, x, y + 3 + i * 7, 4, 4, renk)
            Set s = T_Dik(ws, AKT_PREFIX & "IKON_SATIR_C_" & i, x + 7, y + 3 + i * 7, 17, 4, renk)
        Next i
    Else
        For i = 0 To 2
            ad = AKT_PREFIX & "IKON_PARA_" & i: Set s = T_Sekil(ws, ad)
            If s Is Nothing Then Set s = ws.Shapes.AddShape(msoShapeOval, x, y, 23, 10): s.Name = ad
            T_Yerlestir s, x, y + 11 - i * 5, 23, 10
            s.Fill.ForeColor.RGB = renk: s.line.ForeColor.RGB = vbWhite: s.line.Weight = 0.8
            s.ZOrder msoBringToFront
        Next i
    End If
End Sub

Private Function T_Sekil(ByVal ws As Worksheet, ByVal ad As String) As Shape
    On Error Resume Next
    Set T_Sekil = ws.Shapes(ad)
    On Error GoTo 0
End Function

Private Function T_Sayfa(ByVal ad As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If AKFIX_NORMAL(ws.Name) = AKFIX_NORMAL(ad) Then Set T_Sayfa = ws: Exit Function
    Next ws
End Function

Private Function T_Stok(ByVal ws As Worksheet) As Boolean
    Select Case AKFIX_NORMAL(ws.Name)
        Case "CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK": T_Stok = True
    End Select
End Function

Private Function T_Rapor(ByVal ws As Worksheet) As Boolean
    Select Case AKFIX_NORMAL(ws.Name)
        Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR": T_Rapor = True
    End Select
End Function

Private Function T_Il() As String
    Select Case AKSORA_CANLI_STOK_IL()
        Case "ARTVIN": T_Il = "Artvin"
        Case "RIZE": T_Il = "Rize"
        Case "TRABZON": T_Il = "Trabzon"
        Case "GUMUSHANE": T_Il = "Gümüþhane"
        Case "GIRESUN": T_Il = "Giresun"
    End Select
End Function

Private Sub T_KorumaAc(ByVal ws As Worksheet, ByRef p As Variant, ByRef var As Boolean)
    With ws.Protection
        p = Array(ws.ProtectDrawingObjects, ws.ProtectContents, ws.ProtectScenarios, ws.EnableSelection, _
            .AllowFormattingCells, .AllowFormattingColumns, .AllowFormattingRows, _
            .AllowInsertingColumns, .AllowInsertingRows, .AllowInsertingHyperlinks, _
            .AllowDeletingColumns, .AllowDeletingRows, .AllowSorting, .AllowFiltering, .AllowUsingPivotTables)
    End With
    var = p(0) Or p(1) Or p(2)
    If var Then ws.Unprotect Password:=""
End Sub

Private Sub T_KorumaKapat(ByVal ws As Worksheet, ByVal p As Variant)
    ws.Protect Password:="", DrawingObjects:=p(0), Contents:=p(1), Scenarios:=p(2), _
        UserInterfaceOnly:=True, AllowFormattingCells:=p(4), AllowFormattingColumns:=p(5), _
        AllowFormattingRows:=p(6), AllowInsertingColumns:=p(7), AllowInsertingRows:=p(8), _
        AllowInsertingHyperlinks:=p(9), AllowDeletingColumns:=p(10), AllowDeletingRows:=p(11), _
        AllowSorting:=p(12), AllowFiltering:=p(13), AllowUsingPivotTables:=p(14)
    ws.EnableSelection = p(3)
End Sub

Private Function T_Ayni(ByVal a As Variant, ByVal b As Variant) As Boolean
    If IsError(a) Or IsError(b) Then
        If IsError(a) And IsError(b) Then T_Ayni = (CStr(a) = CStr(b))
    ElseIf IsEmpty(a) Or IsEmpty(b) Then
        T_Ayni = IsEmpty(a) And IsEmpty(b)
    Else
        If VarType(a) = VarType(b) Then T_Ayni = (a = b)
    End If
End Function

Private Function T_FiltreImza(ByVal ws As Worksheet) As String
    Dim f As Filter, i As Long, v As Variant, k As Variant, t As String
    If Not ws.AutoFilterMode Then T_FiltreImza = "YOK": Exit Function
    t = ws.AutoFilter.Range.Address
    For i = 1 To ws.AutoFilter.Filters.Count
        Set f = ws.AutoFilter.Filters(i)
        If f.On Then
            t = t & "|" & i & ":" & f.Operator
            v = Empty
            On Error Resume Next
            v = f.Criteria1
            On Error GoTo 0
            If IsArray(v) Then
                For Each k In v: t = t & ";" & CStr(k): Next k
            Else
                t = t & ";" & CStr(v)
            End If
            v = Empty
            On Error Resume Next
            v = f.Criteria2
            On Error GoTo 0
            If IsArray(v) Then
                For Each k In v: t = t & ";" & CStr(k): Next k
            ElseIf Not IsEmpty(v) Then
                t = t & ";" & CStr(v)
            End If
        End If
    Next i
    T_FiltreImza = t
End Function



Private Sub T_PanelTekBanner(ByVal ws As Worksheet, ByVal r As Range)
    Dim s As Shape, n As String, f As String, ad As String, il As String, h0 As Double
    il = AKSORA_CANLI_STOK_IL()
    ad = AKT_PREFIX & "PANEL_TEK_" & il & "_20261007"
    For Each s In ws.Shapes
        n = s.Name
        If Left$(n, Len(AKT_PREFIX & "PANEL_")) = AKT_PREFIX & "PANEL_" Then s.Visible = msoFalse
        If Left$(n, 17) = "AKSORA_IL_BANNER_" Then s.Visible = msoFalse
    Next s
    Set s = T_Sekil(ws, ad)
    If s Is Nothing Then
        Set s = AKSORA_MALZEME_BANT_EKLE(ws, il, r.Left, r.Top)
        s.Name = ad
    End If
    T_Yerlestir s, r.Left, r.Top, r.Width, r.Height
    s.Visible = msoTrue: s.OnAction = "": s.ZOrder msoBringToFront
End Sub

Private Sub T_PanelHizala(ByVal ws As Worksheet)
    Dim s As Shape, r As Range, ad As Variant, i As Long, x As Double, bw As Double
    Dim buton As Variant, alan As Variant
    buton = Array("btnBaglantiKontrol", "btnTumStok", "btnMalzemeAraPanel", "btnCDP", "btnEPP", "btnIlce", "btnYuk", "AKSORA15_IL_DEGISTIR")
    alan = Array("B15:E18", "G15:O18", "Q15:T18", "B20:E21", "G20:J21", "L20:O21", "Q20:T21", "W6:AB6")
    For i = 0 To UBound(buton)
        Set s = T_Sekil(ws, CStr(buton(i)))
        If Not s Is Nothing Then
            Set r = ws.Range(CStr(alan(i)))
            T_Yerlestir s, r.Left, r.Top + 1, r.Width, r.Height - 2
        End If
    Next i
    Set r = ws.Range("B25:T26")
    bw = (r.Width - 6 * 8) / 7
    i = 0
    For Each ad In Array("goStok", "goRapor", "goCDP", "goEPP", "goIlce", "goYuk", "goStokKarsilastir")
        Set s = T_Sekil(ws, CStr(ad))
        If Not s Is Nothing Then T_Yerlestir s, r.Left + i * (bw + 8), r.Top + 1, bw, r.Height - 2
        i = i + 1
    Next ad
    For Each ad In Array("AKSORA15_UST_ZEMIN", "AKSORA15_ALT_ZEMIN", "AKSORA15_TARIH_ZEMIN")
        Set s = T_Sekil(ws, CStr(ad))
        If Not s Is Nothing Then
            Select Case CStr(ad)
                Case "AKSORA15_UST_ZEMIN": Set r = ws.Range("B2:AB5")
                Case "AKSORA15_ALT_ZEMIN": Set r = ws.Range("B28:AB29")
                Case Else: Set r = ws.Range("B30:AB30")
            End Select
            T_Yerlestir s, r.Left, r.Top, r.Width, r.Height
        End If
    Next ad
    Set s = T_Sekil(ws, "AKSORA15_IL_DEGISTIR")
    If Not s Is Nothing Then
        Set r = ws.Range("B6:AB6")
        AKSORA_IL_DEGISTIR_STIL s, r.Left, r.Top, r.Width, r.Height
    End If
End Sub

Public Sub AKSORA_KAPAK_GENISLIK_UYUM(ByVal ws As Worksheet, ByVal wn As Window)
    Dim s As Shape, r As Range, ad As Variant, genislik As Double, taban As Double
    Dim gap As Double, bw As Double, i As Long, x As Double, y As Double
    If ws.Name <> "KAPAK" Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    Set r = ws.Range("A1:AE39")
    taban = r.Width
    genislik = wn.VisibleRange.Width - wn.VisibleRange.Columns(wn.VisibleRange.Columns.Count).Width - 6
    If genislik < taban Then genislik = taban
    Set s = T_Sekil(ws, "AKSORA_KAPAK_YENI_V6")
    If Not s Is Nothing Then
        T_Yerlestir s, 0, 0, genislik, ws.Rows(23).Top
        s.ZOrder msoSendToBack
    End If
    Set s = T_Sekil(ws, "AKSORA_V4_PANEL")
    If Not s Is Nothing Then
        T_Yerlestir s, 0, ws.Rows(23).Top, genislik, ws.Range("A23:AE38").Height
        s.ZOrder msoSendToBack
    End If
    ' Tum alt ogeler panelin icinde belirlenen tek yerlesime baglidir.
    y = ws.Rows(23).Top + 12
    Set s = T_Sekil(ws, "AKSORA_V4_TITLE")
    If Not s Is Nothing Then T_Yerlestir s, 8, y, genislik - 16, 28
    y = y + 34
    Set s = T_Sekil(ws, "AKSORA_V4_INFO")
    If Not s Is Nothing Then T_Yerlestir s, 8, y, genislik - 16, 20
    y = y + 34
    gap = genislik * 0.015: bw = (genislik * 0.88 - gap * 4) / 5
    x = genislik * 0.06
    For i = 1 To 5
        Set s = T_Sekil(ws, "AKSORA_V4_IL_" & i)
        If Not s Is Nothing Then T_Yerlestir s, x + (i - 1) * (bw + gap), y, bw, 42
    Next i
    y = y + 54
    Set s = T_Sekil(ws, "AKSORA_V4_STATUS")
    If Not s Is Nothing Then T_Yerlestir s, 8, y, genislik - 16, 22
    y = y + 26
    Set s = T_Sekil(ws, "AKSORA_V8_TARIH")
    If Not s Is Nothing Then T_Yerlestir s, 8, y, genislik - 16, 20
    y = y + 30
    Set s = T_Sekil(ws, "AKSORA_V4_ENTER")
    If Not s Is Nothing Then T_Yerlestir s, genislik * 0.32, y, genislik * 0.36, 36
    y = y + 44
    Set s = T_Sekil(ws, "AKSORA_V4_CREDIT")
    If Not s Is Nothing Then T_Yerlestir s, 8, y, genislik - 16, 20

End Sub

Public Function AKSORA_GORSEL_PAKET_UYGULA() As String
    Dim ws As Worksheet, ad As Variant, s As Shape, k As Variant
    Dim veri As Object, bag As Object, form As Object, fil As Object
    Dim son As Long, a As Variant, b As Variant, i As Long, j As Long, il As String
    Dim ev As Boolean, sc As Boolean, hataNo As Long, metin As String, asama As String
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 9450, , "Dosya salt okunur."
    If Not AKSORA_TASARIM_AKTIF() Then Err.Raise vbObjectError + 9451, , "Bu paket güncel P1/S2/R1 dosyasý içindir."
    il = AKSORA_CANLI_STOK_IL()
    application.EnableEvents = False: application.ScreenUpdating = False
    Set veri = CreateObject("Scripting.Dictionary")
    Set bag = CreateObject("Scripting.Dictionary")
    Set form = CreateObject("Scripting.Dictionary")
    Set fil = CreateObject("Scripting.Dictionary")
    asama = "Veri ve baðlantý ön kontrolü"
    For Each ws In ThisWorkbook.Worksheets
        If T_Stok(ws) Or T_Rapor(ws) Then
            If T_Stok(ws) Then
                son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
            If son < 7 Then son = 7
            veri(ws.Name & "|A1:L" & son) = ws.Range("A1:L" & son).Value2
        End If
        If T_Stok(ws) Or T_Rapor(ws) Or ws.Name = "ANA PANEL" Or ws.Name = "KAPAK" Then
            For Each s In ws.Shapes
                If s.Type <> msoChart And s.Type <> msoFormControl And s.Type <> msoOLEControlObject Then
                    If s.Name <> "AKT1_DEPO_OK_ZEMIN" And s.Name <> "AKT1_DEPO_OK_UC" And _
                        Left$(s.Name, Len(AKT_PREFIX & "DEPO_KART_")) <> AKT_PREFIX & "DEPO_KART_" Then
                        If Len(s.OnAction) > 0 Then bag(ws.Name & "|" & s.Name) = s.OnAction
                    End If
                End If
            Next s
        End If
    Next ws
    asama = "Görsel düzenleme"
    For Each ad In Array("ANA PANEL", "CDP_STOK", "EPP_STOK", "STOKLAR", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
        AKSORA_TASARIM_SAYFA ThisWorkbook.Worksheets(CStr(ad)), True
    Next ad
    asama = "Aktif ilin arþiv bilgisini tam stok üzerinden eþitleme"
    AKSORA_CANLI_ARSIV_BILGISINI_ONAR
    asama = "Veri ve baðlantý son kontrolü"
    For Each k In veri.keys
        a = veri(k): Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        b = ws.Range(Split(CStr(k), "|")(1)).Value2
        For i = 1 To UBound(a, 1)
            For j = 1 To UBound(a, 2)
                If Not T_Ayni(a(i, j), b(i, j)) Then Err.Raise vbObjectError + 9452, , "Veri deðiþti: " & CStr(k)
            Next j
        Next i
    Next k
    For Each k In form.keys
        Set ws = ThisWorkbook.Worksheets(CStr(k)): a = form(k)
        If CStr(ws.Range("C3").Formula) <> CStr(a(0)) Or CStr(ws.Range("C4").Formula) <> CStr(a(1)) Then _
            Err.Raise vbObjectError + 9453, , "KPI formülü deðiþti: " & ws.Name
        If Not T_Ayni(ws.Range("H2").Value2, a(2)) Then Err.Raise vbObjectError + 9454, , "Depo seçimi deðiþti."
        If T_FiltreImza(ws) <> fil(k) Then Err.Raise vbObjectError + 9455, , "Filtre deðiþti: " & ws.Name
    Next k
    For Each k In bag.keys
        Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        Set s = T_Sekil(ws, Split(CStr(k), "|")(1))
        If s Is Nothing Then Err.Raise vbObjectError + 9456, , "Buton eksildi: " & CStr(k)
        If s.OnAction <> CStr(bag(k)) Then Err.Raise vbObjectError + 9457, , "Makro baðlantýsý deðiþti: " & CStr(k)
    Next k
    If AKSORA_CANLI_STOK_IL() <> il Then Err.Raise vbObjectError + 9458, , "Canlý stok ili deðiþti."
    asama = "Kapak görünümü"
    ThisWorkbook.Worksheets("KAPAK").Activate
    AKSORA_V8_KAPAK_GORUNUM
    application.EnableEvents = ev: application.ScreenUpdating = sc
    AKSORA_GORSEL_PAKET_UYGULA = "AKSORA_OK"
    Exit Function
Hata:
    hataNo = Err.Number: metin = asama & ": " & Err.Description
    application.EnableEvents = ev: application.ScreenUpdating = sc
    Err.Raise hataNo, "AKSORA_GORSEL_PAKET_UYGULA", metin
End Function

Private Sub T_RaporGrafikSinirla(ByVal ws As Worksheet)
    Dim z As Shape, s As Shape, merkez As Shape, co As ChartObject
    Dim eskiX As Double, eskiW As Double, yeniX As Double, yeniW As Double, oran As Double
    Dim cx As Double, cy As Double, pw As Double, ph As Double, ad As String
    ad = AKFIX_NORMAL(ws.Name)
    If ad <> "ILCE_RAPORLAR" And ad <> "YUKLENICI_RAPOR" Then Exit Sub
    Set z = T_Sekil(ws, "AKSORA_RP_HALKA_ZEMIN")
    If z Is Nothing Then Exit Sub
    eskiX = z.Left: eskiW = z.Width
    If eskiW <= 0 Then Exit Sub
    yeniX = ws.Range("G8").Left
    If yeniX < ws.Range("E8").MergeArea.Left + ws.Range("E8").MergeArea.Width + 12 Then _
        yeniX = ws.Range("E8").MergeArea.Left + ws.Range("E8").MergeArea.Width + 12
    yeniW = ws.Range("T8").Left + ws.Range("T8").Width - 6 - yeniX
    If yeniW < 200 Then Err.Raise vbObjectError + 9500, , "Grafik için yeterli boþ alan yok: " & ws.Name
    If Abs(eskiX - yeniX) < 0.1 And Abs(eskiW - yeniW) < 0.1 Then Exit Sub
    oran = yeniW / eskiW
    For Each s In ws.Shapes
        If Left$(s.Name, Len("AKSORA_RP_HALKA_")) = "AKSORA_RP_HALKA_" Then
            s.LockAspectRatio = msoFalse
            s.Left = yeniX + (s.Left - eskiX) * oran
            s.Width = s.Width * oran
            s.Placement = xlFreeFloating
        End If
    Next s
    Set merkez = T_Sekil(ws, "AKSORA_RP_HALKA_MERKEZ")
    If Not merkez Is Nothing Then
        For Each co In ws.ChartObjects
            If co.Name = "AKSORA_RP_HALKA_GRAFIK" Then
                co.Chart.Refresh
                pw = co.Chart.PlotArea.InsideWidth: ph = co.Chart.PlotArea.InsideHeight
                cx = co.Left + co.Chart.PlotArea.InsideLeft + pw / 2
                cy = co.Top + co.Chart.PlotArea.InsideTop + ph / 2
                pw = application.Min(pw, ph) * 0.72 * 0.9
                If pw > 0 Then T_Yerlestir merkez, cx - pw / 2, cy - 20, pw, 40
                merkez.ZOrder msoBringToFront
                Exit For
            End If
        Next co
    End If
End Sub

Public Sub AKSORA_GORSEL_SON_DUZELT()
    Dim sonuc As String
    On Error GoTo Hata
    sonuc = AKSORA_GORSEL_SON_UYGULA()
    If sonuc <> "AKSORA_OK" Then Err.Raise vbObjectError + 9501, , "Uygulama doðrulanamadý."
    MsgBox "Kapak, stok kutularý ve iki rapor grafiði düzenlendi." & vbCrLf & _
        "Veriler, filtreler, KPI formülleri ve buton baðlantýlarý doðrulandý." & vbCrLf & _
        "Görünümleri kontrol edip kaydedin.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    MsgBox "Düzeltme tamamlanamadý: " & Err.Description & vbCrLf & _
        "Bu denemeyi kaydetmeden kapatýn.", vbExclamation, "AKSORA"
End Sub

Public Function AKSORA_GORSEL_SON_UYGULA() As String
    Dim kendiGecisKilidi As Boolean
    Dim ws As Worksheet, ad As Variant, s As Shape, k As Variant
    Dim veri As Object, bag As Object, form As Object, fil As Object
    Dim son As Long, a As Variant, b As Variant, i As Long, j As Long, il As String
    Dim ev As Boolean, sc As Boolean, hataNo As Long, metin As String, asama As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then _
        Err.Raise vbObjectError + 9502, , "Devam eden iþlem bittikten sonra deneyin."
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 9450, , "Dosya salt okunur."
    If Not AKSORA_TASARIM_AKTIF() Then Err.Raise vbObjectError + 9451, , "Bu paket güncel P1/S2/R1 dosyasý içindir."
    il = AKSORA_CANLI_STOK_IL()
    AKSORA_V13_GECIS = True: kendiGecisKilidi = True
    application.EnableEvents = False: application.ScreenUpdating = False
    Set veri = CreateObject("Scripting.Dictionary")
    Set bag = CreateObject("Scripting.Dictionary")
    Set form = CreateObject("Scripting.Dictionary")
    Set fil = CreateObject("Scripting.Dictionary")
    asama = "Veri ve baðlantý ön kontrolü"
    For Each ws In ThisWorkbook.Worksheets
        If T_Stok(ws) Or T_Rapor(ws) Then
            If T_Stok(ws) Then
                son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
            If son < 7 Then son = 7
            veri(ws.Name & "|A1:L" & son) = ws.Range("A1:L" & son).Value2
        End If
        If T_Stok(ws) Or T_Rapor(ws) Or ws.Name = "ANA PANEL" Or ws.Name = "KAPAK" Then
            For Each s In ws.Shapes
                If s.Type <> msoChart And s.Type <> msoFormControl And s.Type <> msoOLEControlObject Then
                    If s.Name <> "AKT1_DEPO_OK_ZEMIN" And s.Name <> "AKT1_DEPO_OK_UC" And _
                        Left$(s.Name, Len(AKT_PREFIX & "DEPO_KART_")) <> AKT_PREFIX & "DEPO_KART_" Then
                        If Len(s.OnAction) > 0 Then bag(ws.Name & "|" & s.Name) = s.OnAction
                    End If
                End If
            Next s
        End If
    Next ws
    asama = "Görsel düzenleme"
    For Each ad In Array("ANA PANEL", "CDP_STOK", "EPP_STOK", "STOKLAR", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
        AKSORA_TASARIM_SAYFA ThisWorkbook.Worksheets(CStr(ad))
    Next ad
    asama = "Ýlçe ve yüklenici grafik yerleþimi"
    For Each ad In Array("ÝLÇE_RAPORLAR", "YÜKLENÝCÝ_RAPOR")
        AKSORA_RAPOR_GRAFIK_YENILE CStr(ad)
        AKSORA_TASARIM_SAYFA ThisWorkbook.Worksheets(CStr(ad))
    Next ad
    asama = "Veri ve baðlantý son kontrolü"
    For Each k In veri.keys
        a = veri(k): Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        b = ws.Range(Split(CStr(k), "|")(1)).Value2
        For i = 1 To UBound(a, 1)
            For j = 1 To UBound(a, 2)
                If Not T_Ayni(a(i, j), b(i, j)) Then Err.Raise vbObjectError + 9452, , "Veri deðiþti: " & CStr(k)
            Next j
        Next i
    Next k
    For Each k In form.keys
        Set ws = ThisWorkbook.Worksheets(CStr(k)): a = form(k)
        If CStr(ws.Range("C3").Formula) <> CStr(a(0)) Or CStr(ws.Range("C4").Formula) <> CStr(a(1)) Then _
            Err.Raise vbObjectError + 9453, , "KPI formülü deðiþti: " & ws.Name
        If Not T_Ayni(ws.Range("H2").Value2, a(2)) Then Err.Raise vbObjectError + 9454, , "Depo seçimi deðiþti."
        If T_FiltreImza(ws) <> fil(k) Then Err.Raise vbObjectError + 9455, , "Filtre deðiþti: " & ws.Name
    Next k
    For Each k In bag.keys
        Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        Set s = T_Sekil(ws, Split(CStr(k), "|")(1))
        If s Is Nothing Then Err.Raise vbObjectError + 9456, , "Buton eksildi: " & CStr(k)
        If s.OnAction <> CStr(bag(k)) Then Err.Raise vbObjectError + 9457, , "Makro baðlantýsý deðiþti: " & CStr(k)
    Next k
    If AKSORA_CANLI_STOK_IL() <> il Then Err.Raise vbObjectError + 9458, , "Canlý stok ili deðiþti."
    asama = "Kapak görünümü"
    ThisWorkbook.Worksheets("KAPAK").Activate
    AKSORA_V8_KAPAK_GORUNUM
    application.EnableEvents = ev: application.ScreenUpdating = sc
    AKSORA_V13_GECIS = False
    AKSORA_GORSEL_SON_UYGULA = "AKSORA_OK"
    Exit Function
Hata:
    hataNo = Err.Number: metin = asama & ": " & Err.Description
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If kendiGecisKilidi Then AKSORA_V13_GECIS = False
    Err.Raise hataNo, "AKSORA_GORSEL_SON_UYGULA", metin
End Function

Public Function AKSORA_DORT_KART_UYGULA() As String
    Dim kendiGecisKilidi As Boolean
    Dim ws As Worksheet, ad As Variant, s As Shape, k As Variant
    Dim veri As Object, bag As Object, form As Object, fil As Object
    Dim son As Long, a As Variant, b As Variant, i As Long, j As Long, il As String
    Dim ev As Boolean, sc As Boolean, hataNo As Long, metin As String, asama As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then _
        Err.Raise vbObjectError + 9502, , "Devam eden iþlem bittikten sonra deneyin."
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 9450, , "Dosya salt okunur."
    If Not AKSORA_TASARIM_AKTIF() Then Err.Raise vbObjectError + 9451, , "Bu paket güncel P1/S2/R1 dosyasý içindir."
    il = AKSORA_CANLI_STOK_IL()
    AKSORA_V13_GECIS = True: kendiGecisKilidi = True
    application.EnableEvents = False: application.ScreenUpdating = False
    Set veri = CreateObject("Scripting.Dictionary")
    Set bag = CreateObject("Scripting.Dictionary")
    Set form = CreateObject("Scripting.Dictionary")
    Set fil = CreateObject("Scripting.Dictionary")
    asama = "Veri ve baðlantý ön kontrolü"
    For Each ws In ThisWorkbook.Worksheets
        If T_Stok(ws) Or T_Rapor(ws) Then
            If T_Stok(ws) Then
                son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
            If son < 7 Then son = 7
            veri(ws.Name & "|A1:L" & son) = ws.Range("A1:L" & son).Value2
        End If
        If T_Stok(ws) Or T_Rapor(ws) Or ws.Name = "ANA PANEL" Or ws.Name = "KAPAK" Then
            For Each s In ws.Shapes
                If s.Type <> msoChart And s.Type <> msoFormControl And s.Type <> msoOLEControlObject Then
                    If s.Name <> "AKT1_DEPO_OK_ZEMIN" And s.Name <> "AKT1_DEPO_OK_UC" And _
                        Left$(s.Name, Len(AKT_PREFIX & "DEPO_KART_")) <> AKT_PREFIX & "DEPO_KART_" Then
                        If Len(s.OnAction) > 0 Then bag(ws.Name & "|" & s.Name) = s.OnAction
                    End If
                End If
            Next s
        End If
    Next ws
    asama = "Görsel düzenleme"
    For Each ad In Array("ANA PANEL", "CDP_STOK", "EPP_STOK", "STOKLAR", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
        AKSORA_TASARIM_SAYFA ThisWorkbook.Worksheets(CStr(ad))
    Next ad
    asama = "Veri ve baðlantý son kontrolü"
    For Each k In veri.keys
        a = veri(k): Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        b = ws.Range(Split(CStr(k), "|")(1)).Value2
        For i = 1 To UBound(a, 1)
            For j = 1 To UBound(a, 2)
                If Not T_Ayni(a(i, j), b(i, j)) Then Err.Raise vbObjectError + 9452, , "Veri deðiþti: " & CStr(k)
            Next j
        Next i
    Next k
    For Each k In form.keys
        Set ws = ThisWorkbook.Worksheets(CStr(k)): a = form(k)
        If CStr(ws.Range("C3").Formula) <> CStr(a(0)) Or CStr(ws.Range("C4").Formula) <> CStr(a(1)) Then _
            Err.Raise vbObjectError + 9453, , "KPI formülü deðiþti: " & ws.Name
        If Not T_Ayni(ws.Range("H2").Value2, a(2)) Then Err.Raise vbObjectError + 9454, , "Depo seçimi deðiþti."
        If T_FiltreImza(ws) <> fil(k) Then Err.Raise vbObjectError + 9455, , "Filtre deðiþti: " & ws.Name
    Next k
    For Each k In bag.keys
        Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        Set s = T_Sekil(ws, Split(CStr(k), "|")(1))
        If s Is Nothing Then Err.Raise vbObjectError + 9456, , "Buton eksildi: " & CStr(k)
        If s.OnAction <> CStr(bag(k)) Then Err.Raise vbObjectError + 9457, , "Makro baðlantýsý deðiþti: " & CStr(k)
    Next k
    If AKSORA_CANLI_STOK_IL() <> il Then Err.Raise vbObjectError + 9458, , "Canlý stok ili deðiþti."
    asama = "Kapak görünümü"
    ThisWorkbook.Worksheets("KAPAK").Activate
    AKSORA_V8_KAPAK_GORUNUM
    application.EnableEvents = ev: application.ScreenUpdating = sc
    AKSORA_V13_GECIS = False
    AKSORA_DORT_KART_UYGULA = "AKSORA_OK"
    Exit Function
Hata:
    hataNo = Err.Number: metin = asama & ": " & Err.Description
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If kendiGecisKilidi Then AKSORA_V13_GECIS = False
    Err.Raise hataNo, "AKSORA_DORT_KART_UYGULA", metin
End Function
Public Sub AKSORA_DORT_KART_DUZELT()
    Dim sonuc As String
    On Error GoTo Hata
    sonuc = AKSORA_DORT_KART_UYGULA()
    If sonuc <> "AKSORA_OK" Then Err.Raise vbObjectError + 9510, , "Doðrulama sonucu alýnamadý."
    MsgBox "Stok kartlarý ve miktar biçimleri düzenlendi." & vbCrLf & _
        "Stok deðerleri, KPI formülleri, filtreler ve buton baðlantýlarý doðrulandý." & vbCrLf & _
        "Görünümü kontrol edip kaydedin.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    MsgBox "Düzeltme tamamlanamadý: " & Err.Description & vbCrLf & _
        "Bu denemeyi kaydetmeden kapatýn.", vbExclamation, "AKSORA"
End Sub

Public Function AKSORA_STOK_TEMA_UYGULA() As String
    Dim kendiGecisKilidi As Boolean
    Dim ws As Worksheet, ad As Variant, s As Shape, k As Variant
    Dim veri As Object, bag As Object, form As Object, fil As Object
    Dim son As Long, a As Variant, b As Variant, i As Long, j As Long, il As String
    Dim ev As Boolean, sc As Boolean, hataNo As Long, metin As String, asama As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then _
        Err.Raise vbObjectError + 9502, , "Devam eden iþlem bittikten sonra deneyin."
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    If ThisWorkbook.ReadOnly Then Err.Raise vbObjectError + 9450, , "Dosya salt okunur."
    If Not AKSORA_TASARIM_AKTIF() Then Err.Raise vbObjectError + 9451, , "Bu paket güncel P1/S2/R1 dosyasý içindir."
    il = AKSORA_CANLI_STOK_IL()
    AKSORA_V13_GECIS = True: kendiGecisKilidi = True
    application.EnableEvents = False: application.ScreenUpdating = False
    Set veri = CreateObject("Scripting.Dictionary")
    Set bag = CreateObject("Scripting.Dictionary")
    Set form = CreateObject("Scripting.Dictionary")
    Set fil = CreateObject("Scripting.Dictionary")
    asama = "Veri ve baðlantý ön kontrolü"
    For Each ws In ThisWorkbook.Worksheets
        If T_Stok(ws) Or T_Rapor(ws) Then
            If T_Stok(ws) Then
                son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
            If son < 7 Then son = 7
            veri(ws.Name & "|A1:L" & son) = ws.Range("A1:L" & son).Value2
        End If
        If T_Stok(ws) Or T_Rapor(ws) Or ws.Name = "ANA PANEL" Or ws.Name = "KAPAK" Then
            For Each s In ws.Shapes
                If s.Type <> msoChart And s.Type <> msoFormControl And s.Type <> msoOLEControlObject Then
                    If s.Name <> "AKT1_DEPO_OK_ZEMIN" And s.Name <> "AKT1_DEPO_OK_UC" And _
                        Left$(s.Name, Len(AKT_PREFIX & "DEPO_KART_")) <> AKT_PREFIX & "DEPO_KART_" Then
                        If Len(s.OnAction) > 0 Then bag(ws.Name & "|" & s.Name) = s.OnAction
                    End If
                End If
            Next s
        End If
    Next ws
    asama = "Görsel düzenleme"
    For Each ad In Array("ANA PANEL", "CDP_STOK", "EPP_STOK", "STOKLAR", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
        AKSORA_TASARIM_SAYFA ThisWorkbook.Worksheets(CStr(ad))
    Next ad
    asama = "Veri ve baðlantý son kontrolü"
    For Each k In veri.keys
        a = veri(k): Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        b = ws.Range(Split(CStr(k), "|")(1)).Value2
        For i = 1 To UBound(a, 1)
            For j = 1 To UBound(a, 2)
                If Not T_Ayni(a(i, j), b(i, j)) Then Err.Raise vbObjectError + 9452, , "Veri deðiþti: " & CStr(k)
            Next j
        Next i
    Next k
    For Each k In form.keys
        Set ws = ThisWorkbook.Worksheets(CStr(k)): a = form(k)
        If CStr(ws.Range("C3").Formula) <> CStr(a(0)) Or CStr(ws.Range("C4").Formula) <> CStr(a(1)) Then _
            Err.Raise vbObjectError + 9453, , "KPI formülü deðiþti: " & ws.Name
        If Not T_Ayni(ws.Range("H2").Value2, a(2)) Then Err.Raise vbObjectError + 9454, , "Depo seçimi deðiþti."
        If T_FiltreImza(ws) <> fil(k) Then Err.Raise vbObjectError + 9455, , "Filtre deðiþti: " & ws.Name
    Next k
    For Each k In bag.keys
        Set ws = ThisWorkbook.Worksheets(Split(CStr(k), "|")(0))
        Set s = T_Sekil(ws, Split(CStr(k), "|")(1))
        If s Is Nothing Then Err.Raise vbObjectError + 9456, , "Buton eksildi: " & CStr(k)
        If s.OnAction <> CStr(bag(k)) Then Err.Raise vbObjectError + 9457, , "Makro baðlantýsý deðiþti: " & CStr(k)
    Next k
    If AKSORA_CANLI_STOK_IL() <> il Then Err.Raise vbObjectError + 9458, , "Canlý stok ili deðiþti."
    asama = "Kapak görünümü"
    ThisWorkbook.Worksheets("KAPAK").Activate
    AKSORA_V8_KAPAK_GORUNUM
    application.EnableEvents = ev: application.ScreenUpdating = sc
    AKSORA_V13_GECIS = False
    AKSORA_STOK_TEMA_UYGULA = "AKSORA_OK"
    Exit Function
Hata:
    hataNo = Err.Number: metin = asama & ": " & Err.Description
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If kendiGecisKilidi Then AKSORA_V13_GECIS = False
    Err.Raise hataNo, "AKSORA_STOK_TEMA_UYGULA", metin
End Function
Public Sub AKSORA_STOK_TEMA_DUZELT()
    Dim sonuc As String
    On Error GoTo Hata
    sonuc = AKSORA_STOK_TEMA_UYGULA()
    If sonuc <> "AKSORA_OK" Then Err.Raise vbObjectError + 9520, , "Doðrulama sonucu alýnamadý."
    MsgBox "Stok KPI tasarýmý ve depo seçimi düzenlendi." & vbCrLf & _
        "Veriler, KPI formülleri, filtreler ve buton baðlantýlarý doðrulandý." & vbCrLf & _
        "Görünümü kontrol edip kaydedin.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    MsgBox "Tasarým uygulanamadý: " & Err.Description & vbCrLf & _
        "Bu denemeyi kaydetmeden kapatýn.", vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_DEPO_KART_SEC()
    Dim ws As Worksheet, secim As Object, ev As Boolean
    Dim no As Long, metin As String
    ev = application.EnableEvents
    On Error GoTo Hata
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    Set ws = ActiveSheet
    If Not ws.parent Is ThisWorkbook Then Exit Sub
    Set secim = ws.Shapes(CStr(application.Caller)).ControlFormat
    If secim.Value < 1 Then Exit Sub
    application.EnableEvents = False
    ws.Range("H2").Value2 = CStr(secim.List(secim.Value))
    AKFIX_STOK_SECIM ws
    ws.Range("G1").Select
Cikis:
    application.EnableEvents = ev
    If no <> 0 Then MsgBox "Depo filtresi: " & metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Private Sub T_StokListeBicim(ByVal ws As Worksheet, ByVal son As Long)
    Dim bitis As Long, alan As Range, veri As Range, fc As FormatCondition, hucre As Range
    Dim mavi As Range, r As Long, n As Long, asama As String, no As Long, metin As String
    On Error GoTo Hata
    asama = "Eski serit kurallari"
    Dim i As Long, f As String
    bitis = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If bitis < son Then bitis = son
    If bitis < 7 Then Exit Sub
    Set alan = ws.Range("A7:L" & bitis)
    ' Yalniz eski satir seridi kurallari kaldirilir; diger kosullar korunur.
    For i = alan.FormatConditions.Count To 1 Step -1
        Set fc = alan.FormatConditions(i)
        If fc.Type = xlExpression Then
            f = UCase$(CStr(fc.Formula1))
            If InStr(f, "MOD(") > 0 And InStr(f, "ROW(") > 0 Then fc.Delete
        End If
    Next i
    alan.Interior.Pattern = xlNone
    If bitis > son Then ws.Range("A" & application.Max(7, son + 1) & ":L" & bitis).Borders.LineStyle = xlNone
    If son < 7 Then Exit Sub
    asama = "Satir ve yazi hizalari"
    Set veri = ws.Range("A7:L" & son)
    With veri
        .Font.Name = "Aptos": .Font.Size = 10
        .Font.bold = False: .Font.Color = RGB(18, 49, 80)
        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlLeft
        .IndentLevel = 0: .WrapText = False: .ShrinkToFit = False
        .Interior.Color = vbWhite
    End With
    ws.Rows("7:" & son).RowHeight = 20
    ws.Range("A7:B" & son).HorizontalAlignment = xlCenter
    ws.Range("D7:G" & son).HorizontalAlignment = xlCenter
    ws.Range("C7:C" & son).HorizontalAlignment = xlLeft
    ws.Range("H7:H" & son).HorizontalAlignment = xlLeft
    ws.Range("I7:I" & son).HorizontalAlignment = xlRight
    ws.Range("J7:J" & son).HorizontalAlignment = xlCenter
    asama = "TL bicimi"
    With ws.Range("K7:L" & son)
        .HorizontalAlignment = xlRight
        .NumberFormat = "#,##0.00 ""TL"""
    End With
    For Each hucre In ws.Range("K7:L" & son).Cells
        If VarType(hucre.Value2) = vbString Then
            f = Trim$(CStr(hucre.Value2))
            If Len(f) > 0 And UCase$(Right$(f, 2)) <> "TL" Then _
                hucre.NumberFormat = "#,##0.00 ""TL"";-#,##0.00 ""TL"";0.00 ""TL"";@ ""TL"""
        End If
    Next hucre
    AKSORA_V13_MIKTAR_BICIMI ws.Range("I7:I" & son)
    asama = "Mavi beyaz satirlar"
    For r = 7 To son Step 2
        If Len(CStr(ws.Cells(r, "G").Value2)) > 0 Then
            If mavi Is Nothing Then
                Set mavi = ws.Range("A" & r & ":L" & r)
            Else
                Set mavi = application.Union(mavi, ws.Range("A" & r & ":L" & r))
            End If
            n = n + 1
            If n = 64 Then
                mavi.Interior.Color = RGB(240, 247, 253)
                Set mavi = Nothing: n = 0
            End If
        End If
    Next r
    If Not mavi Is Nothing Then mavi.Interior.Color = RGB(240, 247, 253)
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Err.Raise no, "T_StokListeBicim", asama & " | " & ws.Name & " | " & metin
End Sub

Private Sub T_DepoListeYukseklik(ByVal kontrol As Object, ByVal sayfaAdi As String)
    Dim istenen As Long, no As Long
    istenen = kontrol.ControlFormat.ListCount
    If istenen < 1 Then Exit Sub
    On Error Resume Next
    kontrol.ControlFormat.DropDownLines = istenen
    no = Err.Number: Err.Clear
    On Error GoTo 0
    If no <> 0 Then
        MsgBox sayfaAdi & ": Excel bu listenin tum satirlarini ayni anda acmayi kabul etmedi." & vbCrLf & _
            "Depolarin tamami listede kalir; kaydirarak secilebilir. Kart ve filtre korunur.", vbInformation, "AKSORA"
    End If
End Sub

' ===================== CIKTI AL (tum AKSORA sayfalari) =====================
' Tur 1: stok sayfalari -> SAP sayim listesi (HURDA ve CIKMA/BEDELSIZ haric), her sayfada kase/imza kutulari.
' Tur 2: arama ve karsilastirma -> ekrandaki tablo (filtreli satirlar), renksiz, gri baslik.
' Tur 3: rapor ve diger sayfalar -> sayfanin kendisi (dugmeler basilmaz).
' Cikti PDF olarak acilir (Excel baski onizlemesi gizli seritte donar).
Private Function T_CiktiTur(ByVal ad As String) As Long
    Dim n As String
    n = AKFIX_NORMAL(ad)
    Select Case n
        Case "CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK": T_CiktiTur = 1: Exit Function
        Case "MALZEME_ARA", "STOK_KARSILASTIRMA": T_CiktiTur = 2: Exit Function
        Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR": T_CiktiTur = 3: Exit Function
    End Select
    If Left$(n, 2) = "M_" Or Left$(n, 2) = "Y_" Or Left$(n, 2) = "I_" Then
        If Right$(n, 5) = "_STOK" Then
            T_CiktiTur = 1
        ElseIf Right$(n, 11) = "MALZEME_ARA" Or InStr(n, "KARSILASTIRMA") > 0 Then
            T_CiktiTur = 2
        ElseIf InStr(n, "ANA PANEL") = 0 Then
            T_CiktiTur = 3
        End If
    End If
End Function

' Sayfa acilirken CIKTI AL dugmesi (yoksa) kurulur: ANA PANEL dugmesi ve yanindaki dugmelerin solunda.
Public Sub AKSORA_CIKTI_DUGME(ByVal ws As Worksheet)
    Dim a As Shape, s As Shape, b As Shape, x As Double, y As Double, t As String, p As Variant, koruma As Boolean
    On Error GoTo Son
    If T_CiktiTur(ws.Name) = 0 Or T_Stok(ws) Then Exit Sub
    If Not T_Sekil(ws, "btnStokCikti") Is Nothing Then Exit Sub
    Set a = T_Sekil(ws, "btnAnaPanelDon")
    If a Is Nothing Then Exit Sub
    If InStr(AKFIX_NORMAL(ws.Name), "KARSILASTIRMA") > 0 Then
        ' Karsilastirma sayfalarinda ANA PANEL sag panelde; cikti dugmesi altina.
        x = a.Left: y = a.Top + a.Height + 6
    Else
        x = a.Left
        For Each b In ws.Shapes
            If b.Visible = msoTrue And b.Name <> a.Name And b.Name <> "btnCiktiAl" Then
                t = vbNullString
                On Error Resume Next
                t = b.OnAction
                On Error GoTo Son
                If Len(t) > 0 And Abs(b.Top - a.Top) < a.Height And b.Left < a.Left And b.Left > a.Left - 800 Then
                    If b.Left < x Then x = b.Left
                End If
            End If
        Next b
        x = x - 6 - application.Max(a.Width, 90): y = a.Top
    End If
    Set s = T_Sekil(ws, "btnCiktiAl")
    If Not s Is Nothing Then
        If Abs(s.Left - x) < 1 And Abs(s.Top - y) < 1 Then Exit Sub
    End If
    T_KorumaAc ws, p, koruma
    If s Is Nothing Then
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, application.Max(a.Width, 90), a.Height)
        s.Name = "btnCiktiAl"
        T_Buton s, RGB(15, 118, 110), vbWhite, 8
        s.TextFrame2.TextRange.Text = "ÇIKTI AL"
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_CIKTI_AL"
        ' Ekrandaki dugme ve listeler kagida basilmaz.
        On Error Resume Next
        For Each b In ws.Shapes
            t = vbNullString: t = b.OnAction
            If Len(t) > 0 Or b.Type = msoFormControl Then b.DrawingObject.PrintObject = False
        Next b
        On Error GoTo Son
    End If
    s.Left = x: s.Top = y: s.Width = application.Max(a.Width, 90): s.Height = a.Height
    s.Placement = xlFreeFloating
    If koruma Then T_KorumaKapat ws, p
Son:
End Sub

Public Sub AKSORA_STOK_CIKTI()
    AKSORA_CIKTI_AL
End Sub

Public Sub AKSORA_CIKTI_AL()
    Dim ws As Worksheet, ev As Boolean, sc As Boolean, no As Long, metin As String
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Set ws = ActiveSheet
    ev = application.EnableEvents: sc = application.ScreenUpdating
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Cursor = xlWait
    On Error Resume Next
    Select Case T_CiktiTur(ws.Name)
        Case 1: T_CiktiSayim ws
        Case 2: T_CiktiListe ws
        Case 3: T_CiktiGorunum ws
    End Select
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    application.PrintCommunication = True
    application.Cursor = xlDefault
    application.EnableEvents = ev: application.ScreenUpdating = True
    ws.Activate
    On Error GoTo 0
    If no <> 0 Then MsgBox "Çýktý hazýrlanamadý." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

' Stok sayim listesi (kullanicinin SAP sayim formu).
Private Sub T_CiktiSayim(ByVal ws As Worksheet)
    Dim cM As Long, cT As Long, cP As Long, cQ As Long, cB As Long, cS As Long, son As Long, r As Long, n As Long
    Dim kaynak As Variant, veri() As Variant, cikis() As Variant, gor As Range, ar As Range, i As Long, j As Long, k As Long
    Dim parti As String, mk As Variant, seriVar As Boolean, depo As String, ust As String, tarih As Variant
    Dim adet() As Double, x As Long
    If T_Stok(ws) Then
        cM = 7: cT = 8: cP = 6: cQ = 9: cB = 10: cS = 0
    Else
        cM = 5: cT = 6: cP = 9: cQ = 7: cB = 8: cS = 10
    End If
    son = ws.Cells(ws.Rows.Count, cM).End(xlUp).Row
    If son >= 7 Then
        On Error Resume Next
        Set gor = ws.Range(ws.Cells(7, cM), ws.Cells(son, cM)).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    End If
    If Not gor Is Nothing Then
        kaynak = ws.Range(ws.Cells(7, 1), ws.Cells(son, 12)).Value2
        ' 1 UY, 2 depo yeri tanimi, 3 malzeme, 4 tanim, 5 parti, 6 birim, 7 seri, 8 depo kodu
        ReDim veri(1 To son - 6, 1 To 8): ReDim adet(1 To son - 6)
        For Each ar In gor.Areas
            For r = ar.Row To ar.Row + ar.Rows.Count - 1
                i = r - 6
                If Len(T_CiktiMetin(kaynak(i, cM))) > 0 Then
                    parti = T_CiktiMetin(kaynak(i, cP))
                    If Not T_CiktiDisi(parti) Then
                        n = n + 1
                        veri(n, 1) = T_CiktiMetin(kaynak(i, 1))
                        veri(n, 2) = T_CiktiMetin(kaynak(i, 3))
                        veri(n, 3) = T_CiktiMetin(kaynak(i, cM))
                        veri(n, 4) = T_CiktiMetin(kaynak(i, cT))
                        veri(n, 5) = parti
                        veri(n, 6) = T_CiktiMetin(kaynak(i, cB))
                        veri(n, 8) = T_CiktiMetin(kaynak(i, 2))
                        mk = kaynak(i, cQ)
                        If Not IsError(mk) Then
                            If IsNumeric(mk) And Len(CStr(mk)) > 0 Then adet(n) = CDbl(mk)
                        End If
                        If cS > 0 Then
                            veri(n, 7) = T_CiktiMetin(kaynak(i, cS))
                            If Len(veri(n, 7)) > 0 Then seriVar = True
                        End If
                    End If
                End If
            Next r
        Next ar
    End If
    If n = 0 Then Err.Raise vbObjectError + 9601, , "Çýktý alýnacak stok satýrý yok (HURDA ve ÇIKMA/BEDELSÝZ partiler çýktýya alýnmaz)."
    ' Seri numarali stoklarda (sayac vb.) seri numarasi sutunu eklenir.
    k = IIf(seriVar, 10, 9)
    ReDim cikis(1 To n, 1 To k)
    For i = 1 To n
        cikis(i, 1) = veri(i, 1): cikis(i, 2) = veri(i, 8): cikis(i, 3) = veri(i, 2)
        cikis(i, 4) = veri(i, 3): cikis(i, 5) = veri(i, 4)
        x = 6
        If seriVar Then cikis(i, 6) = veri(i, 7): x = 7
        cikis(i, x) = veri(i, 5)
        cikis(i, x + 1) = T_CiktiMiktar(adet(i), CStr(veri(i, 6)))
    Next i
    depo = Trim$(CStr(ws.Range("H2").Value2)): If Len(depo) = 0 Then depo = "TÜMÜ"
    ust = Replace(ws.Name, "_", " ") & "  |  " & Trim$(CStr(ws.Range("A5").Value2)) & "  |  Depo: " & depo
    tarih = ws.Range("C2").Value
    If Not IsError(tarih) Then
        If IsDate(tarih) Then ust = ust & "  |  Stok tarihi: " & Format$(CDate(tarih), "dd.mm.yyyy hh:nn")
    End If
    If seriVar Then
        T_CiktiYaz ws.Name, Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Seri numarasý", "Parti", _
            "Tahditsiz klnb.", "SAYILAN", "FARK"), cikis, n, Array(9.89, 7, 14.66, 11, 40.55, 16, 7.5, 13.5, 22, 20), _
            Array("@", "@", "@", "@", "@", "@", "@", "@", "@", "@"), _
            Array(xlCenter, xlCenter, xlLeft, xlCenter, xlLeft, xlCenter, xlCenter, xlRight, xlLeft, xlLeft), 24, True, ust, Array(5, 4, 2)
    Else
        T_CiktiYaz ws.Name, Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Parti", _
            "Tahditsiz klnb.", "SAYILAN", "FARK"), cikis, n, Array(9.89, 7, 14.66, 11, 40.55, 7.5, 13.5, 28, 26), _
            Array("@", "@", "@", "@", "@", "@", "@", "@", "@"), _
            Array(xlCenter, xlCenter, xlLeft, xlCenter, xlLeft, xlCenter, xlRight, xlLeft, xlLeft), 24, True, ust, Array(5, 4, 2)
    End If
End Sub

Private Function T_CiktiMiktar(ByVal x As Double, ByVal birim As String) As String
    If x = Int(x) Then
        T_CiktiMiktar = AKSORA_SAYI_METNI(x, "#,##0")
    Else
        T_CiktiMiktar = AKSORA_SAYI_METNI(x, "#,##0.###")
    End If
    T_CiktiMiktar = Trim$(T_CiktiMiktar & " " & birim)
End Function

' Arama ve karsilastirma: ekrandaki tablo (filtreli satirlar) aynen, renksiz.
Private Sub T_CiktiListe(ByVal ws As Worksheet)
    Dim hdr As Long, c1 As Long, c2 As Long, son As Long, j As Long, k As Long, n As Long, r As Long, i As Long
    Dim kol() As Long, bas() As Variant, gen() As Variant, fmt() As Variant, hiz() As Variant
    Dim kaynak As Variant, veri() As Variant, gor As Range, ar As Range, bos As Boolean, ust As String, kars As Boolean
    kars = (InStr(AKFIX_NORMAL(ws.Name), "KARSILASTIRMA") > 0)
    If ws.AutoFilterMode Then
        hdr = ws.AutoFilter.Range.Row: c1 = ws.AutoFilter.Range.Column
        c2 = c1 + ws.AutoFilter.Range.Columns.Count - 1
    Else
        hdr = IIf(kars, 15, 12): c1 = 1: c2 = 16
    End If
    ReDim kol(1 To c2 - c1 + 1)
    For j = c1 To c2
        If Not ws.Columns(j).Hidden And Len(Trim$(CStr(ws.Cells(hdr, j).Value2))) > 0 Then k = k + 1: kol(k) = j
    Next j
    If k = 0 Then Err.Raise vbObjectError + 9602, , "Tablo bulunamadý."
    For j = 1 To k
        r = ws.Cells(ws.Rows.Count, kol(j)).End(xlUp).Row
        If r > son Then son = r
    Next j
    If son > hdr Then
        On Error Resume Next
        Set gor = ws.Range(ws.Cells(hdr + 1, c1), ws.Cells(son, c1)).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
    End If
    If Not gor Is Nothing Then
        kaynak = ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(son, kol(k))).Value2
        ReDim veri(1 To son - hdr, 1 To k)
        For Each ar In gor.Areas
            For r = ar.Row To ar.Row + ar.Rows.Count - 1
                i = r - hdr: bos = True
                For j = 1 To k
                    If IsError(kaynak(i, kol(j))) Then
                        veri(n + 1, j) = "#HATA"
                    Else
                        veri(n + 1, j) = kaynak(i, kol(j))
                    End If
                    If Len(Trim$(CStr(veri(n + 1, j)))) > 0 Then bos = False
                Next j
                If Not bos Then n = n + 1
            Next r
        Next ar
    End If
    If n = 0 Then Err.Raise vbObjectError + 9603, , "Çýktý alýnacak satýr yok."
    ReDim bas(0 To k - 1): ReDim gen(0 To k - 1): ReDim fmt(0 To k - 1): ReDim hiz(0 To k - 1)
    For j = 1 To k
        bas(j - 1) = Trim$(CStr(ws.Cells(hdr, kol(j)).Value2))
        gen(j - 1) = ws.Columns(kol(j)).ColumnWidth
        fmt(j - 1) = ws.Cells(hdr + 1, kol(j)).NumberFormat
        If IsNull(fmt(j - 1)) Then fmt(j - 1) = "General"
        hiz(j - 1) = ws.Cells(hdr + 1, kol(j)).HorizontalAlignment
        If IsNull(hiz(j - 1)) Then hiz(j - 1) = xlGeneral
    Next j
    ust = Trim$(CStr(ws.Range("A1").Value2))
    If Len(Trim$(CStr(ws.Range("A3").Value2))) > 0 Then ust = ust & "  |  " & Trim$(CStr(ws.Range("A3").Value2))
    If kars Then ust = ust & "  |  " & Trim$(CStr(ws.Range("B6").Value2)) & " - " & Trim$(CStr(ws.Range("I6").Value2))
    T_CiktiYaz ws.Name, bas, veri, n, gen, fmt, hiz, 18, kars, ust
End Sub

' Rapor ve diger sayfalar: sayfanin kendisi tek sayfa genisliginde (dugmeler ve listeler basilmaz).
Private Sub T_CiktiGorunum(ByVal ws As Worksheet)
    Dim s As Shape, sagX As Double, altY As Double, sonK As Long, sonS As Long, j As Long, r As Long, dosya As String
    For Each s In ws.Shapes
        If s.Visible = msoTrue Then
            If s.Name <> "btnCiktiAl" And s.Name <> "btnAnaPanelDon" Then
                If s.Left + s.Width > sagX Then sagX = s.Left + s.Width
                If s.Top + s.Height > altY Then altY = s.Top + s.Height
            End If
        End If
    Next s
    For j = 1 To 30
        If Not ws.Columns(j).Hidden Then
            If application.WorksheetFunction.CountA(ws.Columns(j)) > 0 Then sonK = j
            If ws.Columns(j).Left < sagX - 1 Then If j > sonK Then sonK = j
            r = ws.Cells(ws.Rows.Count, j).End(xlUp).Row
            If r > sonS And r < 5000 Then sonS = r
        End If
    Next j
    r = 1
    Do While ws.Rows(r).Top + ws.Rows(r).Height < altY And r < 5000
        r = r + 1
    Loop
    If r > sonS Then sonS = r
    If sonK < 1 Then sonK = 1
    If sonS < 1 Then sonS = 1
    On Error Resume Next
    application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(sonS, sonK)).Address
        .PrintTitleRows = vbNullString
        T_CiktiSayfa ws.PageSetup, Trim$(Replace(ws.Name, "_", " ")), False
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .BlackAndWhite = False
    End With
    application.PrintCommunication = True
    On Error GoTo 0
    dosya = T_CiktiDosya(ws.Name)
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dosya, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
End Sub

' Ortak sayfa duzeni: A4 yatay, dar kenar, altta kucuk "Cikti" ve "Sayfa".
' Ortak sayfa duzeni (ilk tasarim): A4 yatay, dar kenar, ust ortada baslik.
' imza = True: alt bilgide kase/imza bandi (kullanicinin cizimi) resim olarak: dikey cizgiler arasinda
' "ILGILI DEPO SORUMLUSU/LARI" ve "KONTROL EDEN SORUMLU/LAR"; solda kucuk Cikti, sagda kucuk Sayfa.
' Alt kenar ilk tasarima yakin tutulur (sayfa basina satir sayisi azalmaz).
Private Sub T_CiktiSayfa(ByVal ps As PageSetup, ByVal ust As String, ByVal imza As Boolean)
    Dim resim As String
    On Error Resume Next
    ps.PaperSize = xlPaperA4
    ps.Orientation = xlLandscape
    ps.LeftMargin = application.CentimetersToPoints(1): ps.RightMargin = application.CentimetersToPoints(1)
    ps.TopMargin = application.CentimetersToPoints(1.6)
    ps.HeaderMargin = application.CentimetersToPoints(0.6)
    ps.CenterHorizontally = True
    ps.CenterHeader = "&""Aptos""&10&B" & Replace(ust, "&", "&&")
    ps.LeftHeader = vbNullString: ps.RightHeader = vbNullString
    ps.LeftFooter = "&""Aptos""&7Çýktý: " & Format$(Now, "dd.mm.yyyy hh:nn")
    ps.RightFooter = "&""Aptos""&7Sayfa &P / &N"
    ps.CenterFooter = vbNullString
    ps.FooterMargin = application.CentimetersToPoints(0.6)
    ps.BottomMargin = application.CentimetersToPoints(1.4)
    If imza Then
        application.PrintCommunication = True
        resim = T_CiktiImzaResmi()
        If Len(resim) > 0 Then
            ps.CenterFooterPicture.Filename = resim
            ps.CenterFooterPicture.LockAspectRatio = msoFalse
            ps.CenterFooterPicture.Width = application.CentimetersToPoints(27.7)
            ps.CenterFooterPicture.Height = application.CentimetersToPoints(1.3)
            ps.CenterFooter = "&G"
            ps.FooterMargin = application.CentimetersToPoints(0.25)
            ps.BottomMargin = application.CentimetersToPoints(1.65)
        End If
    End If
End Sub

' Kase/imza bandi (siyah-beyaz PNG, 27,7 x 1,3 cm, 300 dpi) kodun icinde; gecici klasore bir kez yazilir.
Private Function T_CiktiImzaResmi() As String
    Dim s As String, yol As String, b() As Byte, d As Object, e As Object, f As Integer
    On Error GoTo Yok
    yol = Environ$("TEMP")
    If Len(yol) = 0 Then yol = AKSORA_KAYIT_KLASORU()
    yol = yol & application.PathSeparator & "AKSORA_IMZA_BANDI_V1.png"
    If Len(Dir$(yol)) > 0 Then T_CiktiImzaResmi = yol: Exit Function
    s = s & "iVBORw0KGgoAAAANSUhEUgAADMgAAACaAQAAAADOf6gIAAAACXBIWXMAAC4jAAAuIwF4pT92AAAGL0lEQVR42u2cv47bxhaHv+ES1hTGmqWLIJ43uClTBNYE90X8CEmXwsmOfI1b5w2yj5A3CLVVSr/BpQIXKakgBWWQnBT8I0q7MZCbiDCT"
    s = s & "39eQu0tBZ+abOXNmtJCJ/CG+/YIPl7cfsRzi5/l7/54gPkxkRmaEzMiMkBmZETIjZEZmhMzIjJAZMbeZt93ll+Ot0JyRGSEzMvPX0X/GeH28FZozMvM+7hoHG7eLGXewsWyNNykHC9HAJuyN34e7QxbTfYjZwe9t4/bhlx3sXGOyxs3QBTsS"
    s = s & "DHu/25pkZwKbcPCwM25j2Zq1bRyYwNYfjDv4XWNJDj5mB998GThYTLrQORPDWDAfAGj48R28wzXhFt6cPPzmAFD2P/06z+iMyQFP0f1w24Tbk3hroOHb6N/Azyevq/trs1Az7yb3HoCQR6igpoQKoGr7B/KToNp5zLRpNXiBsh4HxhAv1NB2"
    s = s & "EQP1+O9G4ey6MDPV5L5rfx4InZTqOPDO21iUEOcxU9uScvBRVZOIi354VFB3eu61rh7H0+LMlOsh8KuyNDFflf+7KnKKlvJ5zU3XN72fuCqPAnmWz2OmyoqXFbDetnRRAcSiixegvKF61rIOAG17OuZW5UKzWeE/HW4fVwC2ffy0wJi82H5z"
    s = s & "ngnsNUAK5BW4ecyUbvufuu9nQ7H95iTeCBCK8sUPx+xG/yyULdilVs0hPX2zlCwFcOaspUNruwk2X31a+IQwptVJVMkxw2aEh8fJxeb15c2Ye7f1vUZ9XIVZWvsweQiT1T4/i9dOl0iAJ3X/aHnRqC5v5m5y/7K7uKQkBhKenT/c1ckWCPVs"
    s = s & "ZjakPB9X92lULwdVHkMWHui7Ii7XzOnWwYxp4/XkEofhd9xapgF2c+62u3eO46JhfBdvOhle3eWraUZwUPmFmmnO3qNyAK4J3XgFyMyQPt5Oauedn8dMtw7G3+uPArLoAbahm9HhJMEdsr/DnDk2JOsT+lMAXgzVQtv9jUgLj+aMq+/fTx7q"
    s = s & "IdtbeAzA9xOjOQET/h5m+rXWTtZU7BCGH9O29bPFU79/J19COl3tUzPtv+xSUc1sxsTQd8Gqb2x5UhB9OvRPSzw/H7jc6QwP1WAx7+I1wFWoJnvL4uTQIr/QUcUMZq7OiuCrac35sgJIhy3PKgwZP3M8m+nc7GZSLlflg8dL3aLfnQ4kyaS+"
    s = s & "Ti2rfKHZ7Pnk/jWwqinaDK5p2d3bUH4xlEk7N9u+xoeau74OeEQ7KQlfD1luRaQMZwuKBQ72gaVpIWbivdt0ut7Uv5v7w1wnmq4b9F2SmpbAcbIMuWPeO21csdgTzcnKOjSsrKGASA6B/I7EAlTVGijL8dzKzVSVFS2h7wnTRTWN13eZNxTg"
    s = s & "KXb9J2VV5cNk7izRjLs97vC7NiS//uz4qcWtX0Fri76Fpi7JyqQqjgdXYRYztlx/PZ7t1W796ixe2Eey288hpmU+CXay41memV2a7TzsE6DJsmg8mMZ5Yppldyl4dmBK2GSNY+/CwQPm6brMyeaZM2nlXtvhUC9kdylbC8b18TogsXaXgGMf"
    s = s & "oIKtbRywuX5Slyz10+bpXHfDVAiAs1iI6diyQMAYTA3gLXOVAEmdkQ0TtLWTiN24Bc26MF+NHRa6ki5LuNDR5uXNPJq8R9/T3gD4lAxcMvaEN5AQEi6bwO9vslp7XNK6qE7jZQgzM8fyxYyvuUwfppdvd/hhuF3110+u4TMToNjDJsTx16b4"
    s = s & "Fy++63c7xRXFPJ/S3PzX5Pt+dYtXFIeTeC18TE6IQJ7fTNoANVSnH6j/hZvyP/b8HN9usn/yf75w3m83af79p5LtEr/dxLAM/GVXvw+wxdfLEHMV/nFmhMzIjJAZmREyIzNCZoTMyIyQGZkRMiNkRmaEzMiMkBkhMzIjZEZmhMzIjJAZITMy"
    s = s & "I2RGZoTMCJmRGSEzMiNkRsiMzAiZkRkhMzIjZEbIjMwImZEZITNCZmRGyIzMCJkRMiMzQmZkRsiMkBmZETIjM0JmZEbIjJAZmREyIzNCZoTMyIyQGZkRMiNkRmaEzMiMkBmZETIjZEZmhMzIjJAZITMyI2RGZoTMCJmRGSEzMiNkRmaEzAiZ"
    s = s & "kRkhMzIjZEbIjMwImZEZITPilN8A+QvaeZaomQAAAAAASUVORK5CYII="
    Set d = CreateObject("MSXML2.DOMDocument")
    Set e = d.createElement("b")
    e.DataType = "bin.base64": e.Text = s
    b = e.nodeTypedValue
    f = FreeFile
    Open yol For Binary Access Write As #f
    Put #f, 1, b
    Close #f
    T_CiktiImzaResmi = yol
    Exit Function
Yok:
    On Error Resume Next
    Close #f
End Function

' Gecici (gorunmez) kitapta tablo kurulur ve PDF olarak acilir. Tum sutunlar tek sayfa genisliginde.
Private Sub T_CiktiYaz(ByVal kaynakAd As String, ByVal bas As Variant, ByRef veri() As Variant, ByVal n As Long, _
    ByVal gen As Variant, ByVal fmt As Variant, ByVal hiz As Variant, ByVal satirH As Double, _
    ByVal imza As Boolean, ByVal ust As String, Optional ByVal sira As Variant)
    Dim wb As Workbook, c As Worksheet, k As Long, j As Long, r As Range
    Dim dosya As String, alan As Range, no As Long, metin As String
    k = UBound(bas) - LBound(bas) + 1
    Set wb = Workbooks.Add(xlWBATWorksheet)
    On Error GoTo Hata
    wb.Windows(1).Visible = False
    Set c = wb.Worksheets(1)
    c.Name = "CIKTI"
    c.Cells.Font.Name = "Aptos": c.Cells.Font.Size = 10: c.Cells.Font.Color = RGB(0, 0, 0)
    For j = 1 To k
        c.Columns(j).ColumnWidth = gen(LBound(gen) + j - 1)
        With c.Range(c.Cells(2, j), c.Cells(n + 1, j))
            .NumberFormat = fmt(LBound(fmt) + j - 1)
            .HorizontalAlignment = hiz(LBound(hiz) + j - 1)
        End With
        c.Cells(1, j).Value2 = bas(LBound(bas) + j - 1)
    Next j
    c.Range("A2").Resize(n, k).Value2 = veri
    ' Sayim listesi malzeme adina gore A-Z (sonra malzeme kodu, depo).
    If Not IsMissing(sira) And n > 1 Then
        Set r = c.Range(c.Cells(2, 1), c.Cells(n + 1, k))
        r.Sort Key1:=r.Columns(sira(0)), Order1:=xlAscending, Key2:=r.Columns(sira(1)), Order2:=xlAscending, _
            Key3:=r.Columns(sira(2)), Order3:=xlAscending, Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom
    End If
    Set alan = c.Range(c.Cells(1, 1), c.Cells(n + 1, k))
    With alan
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(0, 0, 0)
        .VerticalAlignment = xlCenter
        .Font.Color = RGB(0, 0, 0): .Interior.Pattern = xlNone
    End With
    With c.Range(c.Cells(1, 1), c.Cells(1, k))
        .Interior.Color = RGB(217, 217, 217): .Font.bold = True
        .HorizontalAlignment = xlCenter: .WrapText = True
    End With
    c.Rows(1).RowHeight = 30
    c.Range(c.Cells(2, 1), c.Cells(n + 1, 1)).EntireRow.RowHeight = satirH
    application.PrintCommunication = False
    With c.PageSetup
        .PrintArea = alan.Address
        .PrintTitleRows = "$1:$1"
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .BlackAndWhite = True
    End With
    T_CiktiSayfa c.PageSetup, ust, imza
    application.PrintCommunication = True
    dosya = T_CiktiDosya(kaynakAd)
    c.ExportAsFixedFormat Type:=xlTypePDF, Filename:=dosya, Quality:=xlQualityStandard, _
        IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
    wb.Close SaveChanges:=False
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    application.PrintCommunication = True
    wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise no, , metin
End Sub

Private Function T_CiktiDosya(ByVal ad As String) As String
    Dim k As String, x As Variant
    k = Environ$("TEMP")
    If Len(k) = 0 Then k = AKSORA_KAYIT_KLASORU()
    ad = AKFIX_NORMAL(ad)
    For Each x In Array(" ", "/", "\", ":", "*", "?", """", "<", ">", "|")
        ad = Replace(ad, CStr(x), "_")
    Next x
    T_CiktiDosya = k & application.PathSeparator & "AKSORA_" & ad & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".pdf"
End Function

Private Function T_CiktiMetin(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    T_CiktiMetin = Trim$(CStr(v))
End Function

' HURDA ve CIKMA/BEDELSIZ (CIKMA/BDSZ) partiler sayima girmez.
Private Function T_CiktiDisi(ByVal parti As String) As Boolean
    Dim t As String, x As Variant
    t = UCase$(parti)
    For Each x In Array(" ", "_", "-", ".", "/")
        t = Replace(t, CStr(x), vbNullString)
    Next x
    t = Replace(Replace(t, "Ç", "C"), "Ý", "I")
    T_CiktiDisi = (InStr(t, "HURDA") > 0 Or InStr(t, "CIKMA") > 0 Or InStr(t, "BDSZ") > 0 Or InStr(t, "BEDELSIZ") > 0)
End Function
