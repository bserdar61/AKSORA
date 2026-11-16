Attribute VB_Name = "modAKSORA_TASARIM"
Option Explicit

Private Const etiket As String = "_AKSORA_TASARIM_P1S2R1"
Private Const AKT_PREFIX As String = "AKT1_"
' Excel ciktilarinda AKSORA baslik satirlarinin kilidi (kullanicinin VBA sifresinden bagimsiz).
Private Const T_IMZA_SIFRE As String = "AKS-imza-7c41"
Private mDonusSayfa As String
Private mCalisiyor As Boolean
Private mKurulum As Boolean
Private mExcelAktar As Boolean
Private mAzImza As Object
' Excel'e aktar: hedef dosya, toplu mod (klasor + tek depo filtresi), kaydedilen dosya.
Private mExcelYol As String
Private mExcelKapat As Boolean
Private mExcelToplu As Boolean
Private mExcelKlasor As String
Private mDepoFiltre As String
Private mExcelSonuc As String
Private mSayimExcel As Boolean
Private mSayimKilit As String
Private mTopluAdlar As Object

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
            son = AKSORA_SON_SATIR(ws, "G")
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
            son = AKSORA_SON_SATIR(ws, "A")
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
        Set s = ws.Shapes.AddPicture(f, msoFalse, msoTrue, r.Left, r.top, r.Width, r.Height)
        s.Name = AKT_PREFIX & "PANEL_GECICI"
        eski.Delete
        s.Name = "AKSORA15_TEK_ARKA_PLAN"
        s.AlternativeText = "AKSORA_P1_SADE_V1"
    Else
        Set s = eski
    End If
    T_Yerlestir s, r.Left, r.top, r.Width, r.Height
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
    ' "Uretim Yeri" basligi filtre okunun altinda kesilmesin.
    If ws.Columns(1).ColumnWidth < 15 Then ws.Columns(1).ColumnWidth = 15
    If ws.Columns(6).ColumnWidth < 19 Then ws.Columns(6).ColumnWidth = 19
    If ws.Columns(8).ColumnWidth > 46 Then ws.Columns(8).ColumnWidth = 44
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
    son = AKSORA_SON_SATIR(ws, "G")
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
    w = ws.Range("A1:L1").Width: y = ws.Rows(1).top
    Set s = T_Dik(ws, AKT_PREFIX & "STOK_BASLIK_ZEMIN", 0, y, w, 32, RGB(12, 45, 79))
    Set s = T_Yazi(ws, AKT_PREFIX & "STOK_MARKA", "AKSORA", 12, y + 3, 102, 26, 20, vbWhite, True, msoAlignLeft)
    Set s = T_Yazi(ws, AKT_PREFIX & "STOK_SLOGAN", "Stok ve Verimlilik", 118, y + 5, 135, 22, 10, RGB(198, 240, 97), True, msoAlignLeft)
    t = Replace(ws.Name, "_", " ") & "  |  " & T_Il() & " Ýl Müdürlüðü"
    Set s = T_Yazi(ws, AKT_PREFIX & "STOK_BASLIK", t, 260, y + 5, w - 680, 22, 11, vbWhite, True, msoAlignLeft)
    degerler = Array("C2", "C3", "C4")
    basliklar = Array("Son Güncelleme", "Toplam Stok Satýrý", "Toplam Stok Deðeri")
    h = ws.Rows(2).Height: y = ws.Rows(2).top
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
        AKSORA_PENCERE_GERI_AC
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
    T_Yerlestir s, sag - 88, ws.Rows(1).top + 5, 88, 22
    T_Buton s, RGB(16, 57, 94), vbWhite, 9
    s.TextFrame2.TextRange.Text = "ANA PANEL"
    sag = sag - 94: adet = UBound(dugmeler) + 1
    bw = IIf(adet = 1, 170, 140)
    For k = UBound(dugmeler) To 0 Step -1
        ad = CStr(dugmeler(k)): Set s = T_Sekil(ws, ad)
        If s Is Nothing Then Err.Raise vbObjectError + 9313, , "Güncelleme butonu eksik: " & ad
        T_Yerlestir s, sag - bw, ws.Rows(1).top + 5, bw, 22
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
    T_Yerlestir s, sag - 96, ws.Rows(1).top + 5, 96, 22
    T_Buton s, RGB(15, 118, 110), vbWhite, 8
    s.TextFrame2.TextRange.Text = "ÇIKTI AL"
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_CIKTI_AL"
    sag = sag - 96 - 5
    Set s = T_ExcelDugme(ws, sag - 110, ws.Rows(1).top + 5, 110, 22)
    sag = sag - 110 - 5
    AKSORA_STOK_AZ ws
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
    Set s = T_Dik(ws, AKT_PREFIX & "RAPOR_UST_ZEMIN", r.Left, r.top, r.Width, r.Height, RGB(12, 45, 79))
    Select Case AKFIX_NORMAL(ws.Name)
        Case "CDP_RAPOR": kod = "CDP STOK RAPORU"
        Case "EPP_RAPOR": kod = "EPP STOK RAPORU"
        Case "RAPORLAR": kod = "GENEL STOK RAPORU"
        Case "ILCE_RAPORLAR": kod = "ÝLÇE STOK RAPORU"
        Case "YUKLENICI_RAPOR": kod = "YÜKLENÝCÝ STOK RAPORU"
    End Select
    Set s = T_Yazi(ws, AKT_PREFIX & "RAPOR_BASLIK", "AKSORA | " & kod, _
        r.Left + 16, r.top + 5, r.Width - 155, r.Height - 10, 20, vbWhite, True, msoAlignLeft)
    Set s = T_Sekil(ws, p & "BANNER"): If Not s Is Nothing Then s.Visible = msoFalse
    Set s = T_Sekil(ws, p & "BASLIK"): If Not s Is Nothing Then s.Visible = msoFalse
    Set s = T_Sekil(ws, "btnAnaPanelDon")
    If s Is Nothing Then Err.Raise vbObjectError + 9314, , "Rapor ANA PANEL butonu eksik."
    T_Yerlestir s, r.Left + r.Width - 108, r.top + 7, 96, r.Height - 14
    T_Buton s, RGB(16, 57, 94), vbWhite, 9
    s.TextFrame2.TextRange.Text = "ANA PANEL"
    alanlar = Array("A3:E6", "F3:J6", "K3:Q6")
    For i = 0 To 2
        ad = alanlar(i): Set r = ws.Range(CStr(ad))
        x = r.Left + 3: y = r.top + 3: w = r.Width - 6: h = r.Height - 6
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
    son = AKSORA_SON_SATIR(ws, "A")
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
    If z Is Nothing Then y = ws.Range("T25").top Else y = z.top + z.Height + 6
    x = ws.Range("T1").Left + ws.Range("T1").Width
    With s
        On Error Resume Next
        .LockAspectRatio = msoFalse
        .ScaleHeight 1, msoTrue: .ScaleWidth 1, msoTrue
        On Error GoTo 0
        .LockAspectRatio = msoTrue: .Width = 58
        If .Height > 58 Then .Height = 58
        .Left = x - .Width - 4: .top = y
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
        .LockAspectRatio = msoFalse: .Left = x: .top = y: .Width = w: .Height = h
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
        Set s = AKSORA_MALZEME_BANT_EKLE(ws, il, r.Left, r.top)
        s.Name = ad
    End If
    T_Yerlestir s, r.Left, r.top, r.Width, r.Height
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
            T_Yerlestir s, r.Left, r.top + 1, r.Width, r.Height - 2
        End If
    Next i
    ' ATIL STOK dugmesi: stok karsilastir dugmesinin kopyasi (yoksa bir kez olusturulur).
    Set s = T_Sekil(ws, "goAtilStok")
    If s Is Nothing Then
        Set s = T_Sekil(ws, "goStokKarsilastir")
        If Not s Is Nothing Then
            Set s = s.Duplicate
            s.Name = "goAtilStok"
        End If
    End If
    If Not s Is Nothing Then
        s.TextFrame2.TextRange.Text = "ATIL STOK"
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_ATIL_AC"
    End If
    ' SATINALMA KONTROL dugmesi: ayni yolla (yoksa bir kez olusturulur).
    Set s = T_Sekil(ws, "goSatinalma")
    If s Is Nothing Then
        Set s = T_Sekil(ws, "goStokKarsilastir")
        If Not s Is Nothing Then
            Set s = s.Duplicate
            s.Name = "goSatinalma"
        End If
    End If
    If Not s Is Nothing Then
        s.TextFrame2.TextRange.Text = "SATIN ALMA"
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_SATINALMA_AC"
    End If
    ' KRITIK STOK artik SATIN ALMA sayfasinda: eski dugme kaldirilir.
    Set s = T_Sekil(ws, "goKritikStok")
    If Not s Is Nothing Then s.Delete
    ' TRANSFER LISTESI dugmesi: ayni yolla (yoksa bir kez olusturulur).
    Set s = T_Sekil(ws, "goTransfer")
    If s Is Nothing Then
        Set s = T_Sekil(ws, "goStokKarsilastir")
        If Not s Is Nothing Then
            Set s = s.Duplicate
            s.Name = "goTransfer"
        End If
    End If
    If Not s Is Nothing Then
        s.TextFrame2.TextRange.Text = "TRANSFER LÝSTESÝ"
        s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_TRANSFER_AC"
    End If
    Set r = ws.Range("B25:T26")
    bw = (r.Width - 9 * 8) / 10
    i = 0
    For Each ad In Array("goStok", "goRapor", "goCDP", "goEPP", "goIlce", "goYuk", "goStokKarsilastir", "goAtilStok", "goSatinalma", "goTransfer")
        Set s = T_Sekil(ws, CStr(ad))
        If Not s Is Nothing Then T_Yerlestir s, r.Left + i * (bw + 8), r.top + 1, bw, r.Height - 2
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
            T_Yerlestir s, r.Left, r.top, r.Width, r.Height
        End If
    Next ad
    Set s = T_Sekil(ws, "AKSORA15_IL_DEGISTIR")
    If Not s Is Nothing Then
        Set r = ws.Range("B6:AB6")
        AKSORA_IL_DEGISTIR_STIL s, r.Left, r.top, r.Width, r.Height
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
        T_Yerlestir s, 0, 0, genislik, ws.Rows(23).top
        s.ZOrder msoSendToBack
    End If
    Set s = T_Sekil(ws, "AKSORA_V4_PANEL")
    If Not s Is Nothing Then
        T_Yerlestir s, 0, ws.Rows(23).top, genislik, ws.Range("A23:AE38").Height
        s.ZOrder msoSendToBack
    End If
    ' Tum alt ogeler panelin icinde belirlenen tek yerlesime baglidir.
    y = ws.Rows(23).top + 12
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
                son = AKSORA_SON_SATIR(ws, "G")
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = AKSORA_SON_SATIR(ws, "A")
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = AKSORA_SON_SATIR(ws, "G")
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
                cy = co.top + co.Chart.PlotArea.InsideTop + ph / 2
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
                son = AKSORA_SON_SATIR(ws, "G")
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = AKSORA_SON_SATIR(ws, "A")
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = AKSORA_SON_SATIR(ws, "G")
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
                son = AKSORA_SON_SATIR(ws, "G")
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = AKSORA_SON_SATIR(ws, "A")
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = AKSORA_SON_SATIR(ws, "G")
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
                son = AKSORA_SON_SATIR(ws, "G")
                If son < 7 Then son = 7
                veri(ws.Name & "|A7:L" & son) = ws.Range("A7:L" & son).Value2
                form(ws.Name) = Array(ws.Range("C3").Formula, ws.Range("C4").Formula, ws.Range("H2").Value2)
                fil(ws.Name) = T_FiltreImza(ws)
            Else
                son = AKSORA_SON_SATIR(ws, "A")
                If son < 8 Then son = 8
                veri(ws.Name & "|A8:E" & son) = ws.Range("A8:E" & son).Value2
            End If
        ElseIf Left$(ws.Name, 8) = "_AKLIVE_" And Left$(ws.Name, Len("_AKLIVE_" & il & "_")) <> "_AKLIVE_" & il & "_" Then
            son = AKSORA_SON_SATIR(ws, "G")
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
    On Error Resume Next
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
' Stok sayfalarinda solunda EXCEL'E AKTAR dugmesi de kurulur ve liste A-Z siralanir.
Public Sub AKSORA_CIKTI_DUGME(ByVal ws As Worksheet)
    Dim a As Shape, s As Shape, b As Shape, x As Double, y As Double, t As String, p As Variant, koruma As Boolean
    Dim yerinde As Boolean
    On Error GoTo son
    If T_CiktiTur(ws.Name) = 0 Then Exit Sub
    If T_CiktiTur(ws.Name) = 1 Then AKSORA_STOK_AZ ws
    If T_Stok(ws) Then Exit Sub
    If Not T_Sekil(ws, "btnStokCikti") Is Nothing Then Exit Sub
    Set a = T_Sekil(ws, "btnAnaPanelDon")
    If a Is Nothing Then Exit Sub
    If InStr(AKFIX_NORMAL(ws.Name), "KARSILASTIRMA") > 0 Then
        ' Karsilastirma sayfalarinda ANA PANEL sag panelde; cikti dugmesi altina.
        x = a.Left: y = a.top + a.Height + 6
    Else
        x = a.Left
        For Each b In ws.Shapes
            If b.Visible = msoTrue And b.Name <> a.Name And b.Name <> "btnCiktiAl" And b.Name <> "btnExcelAktar" And _
                b.Name <> "btnTopluAktar" And b.Name <> "btnDemontajAktar" And b.Name <> "btnIsgAktar" Then
                t = vbNullString
                On Error Resume Next
                t = b.OnAction
                On Error GoTo son
                If Len(t) > 0 And Abs(b.top - a.top) < a.Height And b.Left < a.Left And b.Left > a.Left - 800 Then
                    If b.Left < x Then x = b.Left
                End If
            End If
        Next b
        x = x - 6 - application.Max(a.Width, 90): y = a.top
    End If
    Set s = T_Sekil(ws, "btnCiktiAl")
    If Not s Is Nothing Then yerinde = (Abs(s.Left - x) < 1 And Abs(s.top - y) < 1)
    If yerinde And (T_CiktiTur(ws.Name) <> 1 Or Not T_Sekil(ws, "btnExcelAktar") Is Nothing) And _
        (Not T_TopluSayfa(ws.Name) Or Not T_Sekil(ws, "btnTopluAktar") Is Nothing) And _
        (ws.Name <> "I_STOK_KARÞILAÞTIRMA" Or Not T_Sekil(ws, "btnDemontajAktar") Is Nothing) And _
        (ws.Name <> "I_ÝSG_MALZEMELERÝ" Or Not T_Sekil(ws, "btnIsgAktar") Is Nothing) Then Exit Sub
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
        On Error GoTo son
    End If
    s.Left = x: s.top = y: s.Width = application.Max(a.Width, 90): s.Height = a.Height
    s.Placement = xlFreeFloating
    If T_CiktiTur(ws.Name) = 1 Then T_ExcelDugme ws, x - 6 - 110, y, 110, a.Height
    If T_TopluSayfa(ws.Name) Then
        T_EkDugme ws, "btnTopluAktar", "TÜM DEPOLARI AKTAR", RGB(21, 87, 52), "AKSORA_EXCEL_TOPLU_AKTAR", _
            x - 12 - 110 - 130, y, 130, a.Height
    End If
    ' Ilce ISG zimmet kontrolu: iki sayfali Excel (CIKTI AL'in solunda).
    If ws.Name = "I_ÝSG_MALZEMELERÝ" Then
        T_EkDugme ws, "btnIsgAktar", "EXCEL'E AKTAR", RGB(33, 115, 70), "AKSORA_ILCE_ISG_AKTAR", x - 6 - 110, y, 110, a.Height
    End If
    ' Ilce stok karsilastirma: demontaj eksikleri her ilce icin ayri Excel (CIKTI AL'in saginda).
    If ws.Name = "I_STOK_KARÞILAÞTIRMA" Then
        T_EkDugme ws, "btnDemontajAktar", "DEMONTAJ EKSÝKLERÝ" & vbLf & "EXCEL'E AKTAR", RGB(185, 28, 28), _
            "AKSORA_ILCE_DEMONTAJ_AKTAR", x + application.Max(a.Width, 90) + 6, y, application.Max(a.Width, 90), a.Height
    End If
    If koruma Then T_KorumaKapat ws, p
son:
End Sub

' Ek dugme (yoksa olusturulur) verilen konuma yerlestirilir; kagida basilmaz.
Private Sub T_EkDugme(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, ByVal renk As Long, _
    ByVal makro As String, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double)
    Dim s As Shape
    Set s = T_Sekil(ws, ad)
    If s Is Nothing Then
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
        s.Name = ad
    End If
    T_Buton s, renk, vbWhite, IIf(InStr(metin, vbLf) > 0, 7, 8)
    s.TextFrame2.TextRange.Text = metin
    On Error Resume Next
    s.TextFrame2.MarginTop = 0: s.TextFrame2.MarginBottom = 0
    s.TextFrame2.MarginLeft = 2: s.TextFrame2.MarginRight = 2
    s.TextFrame2.WordWrap = msoTrue
    On Error GoTo 0
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & makro
    s.Left = x: s.top = y: s.Width = w: s.Height = h
    s.Placement = xlFreeFloating
    On Error Resume Next
    s.DrawingObject.PrintObject = False
    On Error GoTo 0
End Sub

' EXCEL'E AKTAR dugmesi (yoksa olusturulur) verilen konuma yerlestirilir.
Private Function T_ExcelDugme(ByVal ws As Worksheet, ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double) As Shape
    Dim s As Shape
    Set s = T_Sekil(ws, "btnExcelAktar")
    If s Is Nothing Then
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
        s.Name = "btnExcelAktar"
    End If
    T_Buton s, RGB(33, 115, 70), vbWhite, 8
    s.TextFrame2.TextRange.Text = "EXCEL'E AKTAR"
    s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_EXCEL_AKTAR"
    s.Left = x: s.top = y: s.Width = w: s.Height = h
    s.Placement = xlFreeFloating
    On Error Resume Next
    s.DrawingObject.PrintObject = False
    On Error GoTo 0
    Set T_ExcelDugme = s
End Function

' Stok dizisi malzeme adina gore A-Z (Turkce harf sirasi), sonra malzeme kodu, depo ve seri no.
' sutunlar: Array(tanim, malzeme, depo, seri) (0 = yok). Sirali ise ayni dizi doner.
Public Function AKSORA_STOK_DIZI_AZ(ByVal a As Variant, ByVal sutunlar As Variant, Optional ByRef degisti As Boolean) As Variant
    Dim n As Long, k As Long, i As Long, j As Long, anah() As String, idx() As Long, s() As Variant, sirali As Boolean
    degisti = False
    AKSORA_STOK_DIZI_AZ = a
    If Not IsArray(a) Then Exit Function
    n = UBound(a, 1): k = UBound(a, 2)
    If n < 2 Then Exit Function
    ReDim anah(1 To n): ReDim idx(1 To n)
    sirali = True
    For i = 1 To n
        anah(i) = T_AzAnahtar(a, i, sutunlar)
        idx(i) = i
        If (i And 4095) = 0 Then AKSORA_BEKLE_NABIZ
        If i > 1 Then If anah(i) < anah(i - 1) Then sirali = False
    Next i
    If sirali Then Exit Function
    T_AzSirala anah, idx, 1, n
    ReDim s(1 To n, 1 To k)
    For i = 1 To n
        For j = 1 To k
            s(i, j) = a(idx(i), j)
        Next j
    Next i
    degisti = True
    AKSORA_STOK_DIZI_AZ = s
End Function

' Sayfa imzasi: satir sayisi, guncelleme tarihi, il ve ilk / orta / son satirin malzeme ve tanimi.
Private Function T_AzImza(ByVal ws As Worksheet, ByVal son As Long, ByVal cM As Long, ByVal cT As Long) As String
    Dim r As Variant
    On Error Resume Next
    T_AzImza = son & "|" & CStr(ws.Range("C2").Value2) & "|" & CStr(ws.Range("A5").Value2)
    For Each r In Array(7, (7 + son) \ 2, son)
        T_AzImza = T_AzImza & "|" & CStr(ws.Cells(r, cM).Value2) & "|" & CStr(ws.Cells(r, cT).Value2)
    Next r
End Function

' a ByRef: ByVal her satirda tum tabloyu kopyaliyordu (18 bin satirda dakikalar).
Private Function T_AzAnahtar(ByRef a As Variant, ByVal r As Long, ByVal sutunlar As Variant) As String
    Dim i As Long, c As Long, v As Variant, t As String
    For i = 0 To UBound(sutunlar)
        c = sutunlar(i)
        t = vbNullString
        If c > 0 Then
            v = a(r, c)
            If Not IsError(v) Then t = UCase$(Trim$(CStr(v)))
        End If
        If t Like "*[ÇÐÝýÖÞÜ]*" Then
            t = Replace(Replace(Replace(Replace(t, "Ç", "C~"), "Ð", "G~"), "Ý", "I"), "ý", "I")
            t = Replace(Replace(Replace(t, "Ö", "O~"), "Þ", "S~"), "Ü", "U~")
        End If
        T_AzAnahtar = T_AzAnahtar & t & Chr$(1)
    Next i
End Function

Private Sub T_AzSirala(ByRef anah() As String, ByRef idx() As Long, ByVal ilkI As Long, ByVal sonI As Long)
    Dim i As Long, j As Long, p As String, t As String, x As Long
    Do While ilkI < sonI
        i = ilkI: j = sonI: p = anah((ilkI + sonI) \ 2)
        Do While i <= j
            Do While anah(i) < p: i = i + 1: Loop
            Do While anah(j) > p: j = j - 1: Loop
            If i <= j Then
                t = anah(i): anah(i) = anah(j): anah(j) = t
                x = idx(i): idx(i) = idx(j): idx(j) = x
                i = i + 1: j = j - 1
            End If
        Loop
        If j - ilkI < sonI - i Then
            If ilkI < j Then T_AzSirala anah, idx, ilkI, j
            ilkI = i
        Else
            If i < sonI Then T_AzSirala anah, idx, i, sonI
            sonI = j
        End If
    Loop
End Sub

' Stok sayfasindaki liste A-Z siralanir (zaten siraliysa dokunulmaz). Filtre varsa yeniden uygulanir.
Public Sub AKSORA_STOK_AZ(ByVal ws As Worksheet)
    Dim cM As Long, cT As Long, cs As Long, k As Long, son As Long, v As Variant, yeni As Variant, degisti As Boolean
    Dim p As Variant, koruma As Boolean, sc As Boolean, ev As Boolean, durum As Boolean, imza As String
    On Error GoTo son
    If T_CiktiTur(ws.Name) <> 1 Then Exit Sub
    If T_Stok(ws) Then
        cM = 7: cT = 8: cs = 0: k = 12
    Else
        cM = 5: cT = 6: cs = 10: k = 14
    End If
    son = AKSORA_SON_SATIR(ws, cM)
    If son < 8 Then Exit Sub
    ' Buyuk sayfalar (musteri seri no) yazilirken siralanir; sayfa acilisinda bekletilmez.
    If son - 6 > 5000 Then Exit Sub
    ' Ayni oturumda degismemis sayfa yeniden taranmaz (buyuk musteri sayfalarinda gecis hizli kalir).
    If mAzImza Is Nothing Then Set mAzImza = CreateObject("Scripting.Dictionary")
    imza = T_AzImza(ws, son, cM, cT)
    If mAzImza.Exists(ws.Name) Then
        If mAzImza(ws.Name) = imza Then Exit Sub
    End If
    v = ws.Range(ws.Cells(7, 1), ws.Cells(son, k)).Value2
    yeni = AKSORA_STOK_DIZI_AZ(v, Array(cT, cM, 2, cs), degisti)
    If Not degisti Then mAzImza(ws.Name) = imza: Exit Sub
    sc = application.ScreenUpdating: ev = application.EnableEvents: durum = True
    application.ScreenUpdating = False: application.EnableEvents = False
    T_KorumaAc ws, p, koruma
    ws.Range(ws.Cells(7, 1), ws.Cells(son, k)).Value2 = yeni
    If ws.FilterMode Then ws.AutoFilter.ApplyFilter
    If koruma Then T_KorumaKapat ws, p
    application.ScreenUpdating = sc: application.EnableEvents = ev
    mAzImza(ws.Name) = T_AzImza(ws, son, cM, cT)
    Exit Sub
son:
    On Error Resume Next
    If koruma Then T_KorumaKapat ws, p
    If durum Then application.ScreenUpdating = sc: application.EnableEvents = ev
End Sub

' Disaridan (yonetim merkezi vb.) sayfa gorunumu ciktisi.
Public Sub AKSORA_CIKTI_SAYFA(ByVal ws As Worksheet)
    On Error GoTo Hata
    application.ScreenUpdating = False
    T_CiktiGorunum ws
    application.ScreenUpdating = True
    Exit Sub
Hata:
    application.ScreenUpdating = True
    MsgBox "Çýktý hazýrlanamadý." & vbCrLf & vbCrLf & Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_STOK_CIKTI()
    AKSORA_CIKTI_AL
End Sub

' Secili depo(lar)in stok listesi yeni bir Excel kitabina: cikti tablosu (imzasiz), A-Z, stok tarihiyle.
Public Sub AKSORA_EXCEL_AKTAR()
    Dim ws As Worksheet, ev As Boolean, no As Long, metin As String
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Set ws = ActiveSheet
    If T_CiktiTur(ws.Name) <> 1 Then Exit Sub
    ev = application.EnableEvents
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Cursor = xlWait
    mExcelAktar = True: mExcelToplu = False: mExcelKapat = False: mDepoFiltre = vbNullString
    On Error Resume Next
    T_CiktiSayim ws
    no = Err.Number: metin = Err.Description
    mExcelAktar = False: mExcelYol = vbNullString
    application.Cursor = xlDefault
    application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If no <> 0 Then
        MsgBox "Excel'e aktarýlamadý." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
    ElseIf Len(mExcelSonuc) = 0 Then
        MsgBox "Liste açýldý ancak masaüstüne kaydedilemedi; Farklý Kaydet ile kaydediniz.", vbExclamation, "AKSORA"
    Else
        application.StatusBar = "Masaüstüne kaydedildi: " & mExcelSonuc
    End If
End Sub

' Sayfadaki her depo ayri Excel dosyasina (depo adiyla) masaustunde tek klasore kaydedilir.
Public Sub AKSORA_EXCEL_TOPLU_AKTAR()
    Dim ws As Worksheet, son As Long, v As Variant, i As Long, depolar As Object, k As Variant, kodlar() As String
    Dim ac As Boolean, ev As Boolean, klasor As String, tarihMetin As String, tarih As Variant, ok As Long, bos As Long
    Dim Hata As String, n As Long, j As Long, t As String, no As Long, metin As String, ad As String
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Set ws = ActiveSheet
    If Not T_TopluSayfa(ws.Name) Then Exit Sub
    son = AKSORA_SON_SATIR(ws, 5)
    If son < 7 Then MsgBox "Aktarýlacak stok bulunamadý.", vbInformation, "AKSORA": Exit Sub
    v = ws.Range(ws.Cells(7, 1), ws.Cells(son, 3)).Value2
    Set depolar = CreateObject("Scripting.Dictionary"): depolar.CompareMode = vbTextCompare
    For i = 1 To UBound(v, 1)
        t = UCase$(T_CiktiMetin(v(i, 2)))
        If Len(t) > 0 And Len(T_CiktiMetin(v(i, 1))) > 0 Then
            If Not depolar.Exists(t) Then depolar(t) = T_CiktiMetin(v(i, 3))
        End If
    Next i
    n = depolar.Count
    If n = 0 Then MsgBox "Aktarýlacak stok bulunamadý.", vbInformation, "AKSORA": Exit Sub
    ReDim kodlar(1 To n)
    For Each k In depolar.keys
        j = j + 1: kodlar(j) = CStr(k)
    Next k
    For i = 1 To n - 1
        For j = i + 1 To n
            If kodlar(j) < kodlar(i) Then t = kodlar(i): kodlar(i) = kodlar(j): kodlar(j) = t
        Next j
    Next i
    tarih = ws.Range("C2").Value
    tarihMetin = Format$(Date, "dd.mm.yyyy")
    If Not IsError(tarih) Then If IsDate(tarih) Then tarihMetin = Format$(CDate(tarih), "dd.mm.yyyy")
    klasor = AKSORA_MASAUSTU() & application.PathSeparator & _
        AKSORA_DOSYA_ADI("AKSORA - " & T_IlAdi(ws) & " - " & T_SayfaKisa(ws.Name) & " - " & tarihMetin)
    If MsgBox(n & " depo ayrý ayrý Excel dosyasý olarak (depo adýyla) masaüstündeki þu klasöre kaydedilecek:" & vbCrLf & vbCrLf & _
        klasor & vbCrLf & vbCrLf & "Devam edilsin mi?", vbQuestion + vbYesNo, "AKSORA  |  TÜM DEPOLARI AKTAR") <> vbYes Then Exit Sub
    If Not AKSORA_KLASOR_HAZIRLA(klasor) Then MsgBox "Klasör oluþturulamadý:" & vbCrLf & klasor, vbExclamation, "AKSORA": Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  TÜM DEPOLARI EXCEL'E AKTAR", "Hazýrlanýyor...", 0.03)
    ev = application.EnableEvents
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Cursor = xlWait
    mExcelAktar = True: mExcelToplu = True: mExcelKapat = True: mExcelKlasor = klasor
    Set mTopluAdlar = CreateObject("Scripting.Dictionary")
    On Error GoTo Hata
    For i = 1 To n
        ad = CStr(depolar(kodlar(i)))
        If Len(ad) = 0 Then ad = kodlar(i)
        AKSORA_BEKLE_ILERLE 0.03 + 0.95 * (i - 1) / n, i & " / " & n & "  •  " & ad
        mDepoFiltre = UCase$(kodlar(i))
        On Error Resume Next
        T_CiktiSayim ws
        no = Err.Number: metin = Err.Description
        On Error GoTo Hata
        If no = 0 And Len(mExcelSonuc) > 0 Then
            ok = ok + 1
        ElseIf no = vbObjectError + 9601 Then
            bos = bos + 1
        Else
            Hata = Hata & vbCrLf & ad & ": " & IIf(no = 0, "kaydedilemedi", metin)
        End If
        ThisWorkbook.Activate: ws.Activate
    Next i
Cikis:
    On Error Resume Next
    mExcelAktar = False: mExcelToplu = False: mExcelKapat = False: mExcelKlasor = vbNullString
    mDepoFiltre = vbNullString: mExcelYol = vbNullString: Set mTopluAdlar = Nothing
    ThisWorkbook.Activate: ws.Activate
    application.Cursor = xlDefault
    application.EnableEvents = ev: application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    If no <> 0 And ok = 0 And bos = 0 And Len(Hata) = 0 Then
        MsgBox "Aktarým yapýlamadý." & vbCrLf & vbCrLf & metin, vbExclamation, "AKSORA"
        Exit Sub
    End If
    MsgBox ok & " depo Excel dosyasý olarak kaydedildi." & _
        IIf(bos > 0, vbCrLf & bos & " depoda aktarýlacak satýr yok (yalnýz hurda / çýkma).", "") & _
        IIf(Len(Hata) > 0, vbCrLf & vbCrLf & "Kaydedilemeyenler:" & Hata, "") & vbCrLf & vbCrLf & "Klasör: " & klasor, _
        IIf(Len(Hata) > 0, vbExclamation, vbInformation), "AKSORA  |  TÜM DEPOLARI AKTAR"
    If ok > 0 Then AKSORA_KLASOR_AC klasor
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Hata = Hata & vbCrLf & metin
    Resume Cikis
End Sub

' Klasor Windows Gezgini'nde acilir.
Public Sub AKSORA_KLASOR_AC(ByVal klasor As String)
    On Error Resume Next
    Shell "explorer.exe """ & klasor & """", vbNormalFocus
End Sub

Public Sub AKSORA_CIKTI_AL()
    Dim ws As Worksheet, ev As Boolean, sc As Boolean, no As Long, metin As String, cvp As VbMsgBoxResult
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Set ws = ActiveSheet
    mSayimExcel = False
    If T_CiktiTur(ws.Name) = 1 Then
        cvp = MsgBox("Sayým listesi nasýl hazýrlansýn?" & vbCrLf & vbCrLf & _
            "EVET  =  EXCEL (önerilen)" & vbCrLf & _
            "SAYILAN miktarlarý bilgisayarda girin; FARK kendiliðinden hesaplanýr. Sonra yazdýrýp imzalatýn." & vbCrLf & vbCrLf & _
            "HAYIR  =  PDF" & vbCrLf & "Boþ liste; sahada elle doldurmak için.", _
            vbQuestion + vbYesNoCancel, "AKSORA  |  SAYIM LÝSTESÝ")
        If cvp = vbCancel Then Exit Sub
        mSayimExcel = (cvp = vbYes)
    End If
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
    If mSayimExcel And no = 0 Then
        ' Sayim formu acik kalir (kullanici doldurur); AKSORA'ya donuste ayni sayfa.
        mSayimExcel = False
        If Len(mExcelSonuc) > 0 Then application.StatusBar = "Sayým formu masaüstüne kaydedildi: " & mExcelSonuc
        Exit Sub
    End If
    mSayimExcel = False
    ' Gecici kitap kapaninca (olaylar kapaliyken) AKSORA'nin uygulama modu kendiliginden donmuyordu.
    ThisWorkbook.Activate
    ws.Activate
    UYGULAMA_MODU_AC
    AKSORA_BIRIM_AKTIF_GORUNUM
    On Error GoTo 0
    If no <> 0 Then MsgBox "Çýktý hazýrlanamadý." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

' Stok sayim listesi (kullanicinin SAP sayim formu).
Private Sub T_CiktiSayim(ByVal ws As Worksheet)
    Dim cM As Long, cT As Long, cP As Long, cQ As Long, cB As Long, cs As Long, son As Long, r As Long, n As Long
    Dim kaynak As Variant, veri() As Variant, Cikis() As Variant, gor As Range, ar As Range, i As Long, j As Long, k As Long
    Dim parti As String, mk As Variant, seriVar As Boolean, depo As String, ust As String, tarih As Variant
    Dim adet() As Double, x As Long, tum As Object, sec As Object, kk As Variant, gorV As Variant, gor1 As Boolean
    Dim depoAd As String, tarihMetin As String
    If T_Stok(ws) Then
        cM = 7: cT = 8: cP = 6: cQ = 9: cB = 10: cs = 0
    Else
        cM = 5: cT = 6: cP = 9: cQ = 7: cB = 8: cs = 10
    End If
    son = AKSORA_SON_SATIR(ws, cM)
    ' Gorunur satirlar tek seferde (A-Z siralamada depo filtresi binlerce parca uretir).
    ' Toplu aktarimda filtre yerine tek depo (mDepoFiltre) alinir.
    If son >= 7 And Len(mDepoFiltre) = 0 Then gorV = AKSORA_GORUNUR(ws, 7, son, IIf(T_Stok(ws), 12, 14))
    If son >= 7 And (IsArray(gorV) Or Len(mDepoFiltre) > 0) Then
        kaynak = ws.Range(ws.Cells(7, 1), ws.Cells(son, 12)).Value2
        ' 1 UY, 2 depo yeri tanimi, 3 malzeme, 4 tanim, 5 parti, 6 birim, 7 seri, 8 depo kodu
        ReDim veri(1 To son - 6, 1 To 8): ReDim adet(1 To son - 6)
        For i = 1 To son - 6
            If Len(mDepoFiltre) > 0 Then
                gor1 = (UCase$(T_CiktiMetin(kaynak(i, 2))) = mDepoFiltre)
            Else
                gor1 = (gorV(i) = 1)
            End If
            If gor1 Then
                If Len(T_CiktiMetin(kaynak(i, cM))) > 0 Then
                    parti = T_CiktiMetin(kaynak(i, cP))
                    If T_CiktiDahil(parti, kaynak(i, 4), kaynak(i, cT)) Then
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
                        If cs > 0 Then
                            veri(n, 7) = T_CiktiMetin(kaynak(i, cs))
                            If Len(veri(n, 7)) > 0 Then seriVar = True
                        End If
                    End If
                End If
            End If
        Next i
    End If
    If n = 0 Then Err.Raise vbObjectError + 9601, , "Çýktý alýnacak stok satýrý yok (HURDA ve ÇIKMA/BEDELSÝZ partiler alýnmaz; ÇIKMA/BDSZ trafo, hücre, pano ve direk alýnýr)."
    ' Baslikta depo: ciktiya giren depolarin tanimlari; sayfadaki tum depolar ise TUMU.
    Set tum = CreateObject("Scripting.Dictionary"): tum.CompareMode = vbTextCompare
    Set sec = CreateObject("Scripting.Dictionary"): sec.CompareMode = vbTextCompare
    For i = 1 To UBound(kaynak, 1)
        If Len(T_CiktiMetin(kaynak(i, cM))) > 0 Then
            If T_CiktiDahil(T_CiktiMetin(kaynak(i, cP)), kaynak(i, 4), kaynak(i, cT)) Then tum(T_CiktiMetin(kaynak(i, 2))) = True
        End If
    Next i
    For i = 1 To n
        If Not sec.Exists(veri(i, 8)) Then sec(veri(i, 8)) = veri(i, 2)
    Next i
    If sec.Count >= tum.Count And Len(mDepoFiltre) = 0 Then
        depo = "TÜMÜ"
    Else
        depo = vbNullString
        For Each kk In sec.keys
            depo = depo & IIf(Len(depo) > 0, ", ", "") & IIf(Len(CStr(sec(kk))) > 0, CStr(sec(kk)), CStr(kk))
        Next kk
        If Len(depo) > 110 Then depo = Left$(depo, 107) & "..."
    End If
    ' Seri numarali stoklarda (sayac vb.) seri numarasi sutunu eklenir.
    k = IIf(seriVar, 10, 9)
    ReDim Cikis(1 To n, 1 To k)
    For i = 1 To n
        Cikis(i, 1) = veri(i, 1): Cikis(i, 2) = veri(i, 8): Cikis(i, 3) = veri(i, 2)
        Cikis(i, 4) = veri(i, 3): Cikis(i, 5) = veri(i, 4)
        x = 6
        If seriVar Then Cikis(i, 6) = veri(i, 7): x = 7
        Cikis(i, x) = veri(i, 5)
        Cikis(i, x + 1) = T_CiktiMiktar(adet(i), CStr(veri(i, 6)))
    Next i
    ust = T_SayfaBaslik(ws.Name) & "  |  " & Trim$(CStr(ws.Range("A5").Value2)) & "  |  Depo: " & depo
    tarih = ws.Range("C2").Value
    tarihMetin = Format$(Date, "dd.mm.yyyy")
    If Not IsError(tarih) Then
        If IsDate(tarih) Then
            ust = ust & "  |  Stok tarihi: " & Format$(CDate(tarih), "dd.mm.yyyy hh:nn")
            tarihMetin = Format$(CDate(tarih), "dd.mm.yyyy")
        End If
    End If
    If mExcelAktar Then
        ' Dosya adi depo adi: tek aktarimda masaustune "<depo> - <sayfa> - <tarih>", toplu aktarimda klasore "<depo>".
        depoAd = depo
        If depoAd = "TÜMÜ" Then depoAd = T_IlAdi(ws) & " Tüm Depolar"
        If mExcelToplu Then
            depoAd = AKSORA_DOSYA_ADI(depoAd)
            If mTopluAdlar Is Nothing Then Set mTopluAdlar = CreateObject("Scripting.Dictionary")
            If mTopluAdlar.Exists(UCase$(depoAd)) Then depoAd = depoAd & " (" & mDepoFiltre & ")"
            mTopluAdlar(UCase$(depoAd)) = True
            mExcelYol = mExcelKlasor & application.PathSeparator & depoAd & ".xlsx"
        Else
            mExcelYol = AKSORA_MASAUSTU() & application.PathSeparator & _
                AKSORA_DOSYA_ADI(depoAd & " - " & T_SayfaKisa(ws.Name) & " - " & tarihMetin) & ".xlsx"
        End If
        ' Excel: miktar sayi, birim ayri sutun; FARK = SAYILAN - miktar (SAYILAN girilince).
        k = IIf(seriVar, 11, 10)
        ReDim Cikis(1 To n, 1 To k)
        For i = 1 To n
            Cikis(i, 1) = veri(i, 1): Cikis(i, 2) = veri(i, 8): Cikis(i, 3) = veri(i, 2)
            Cikis(i, 4) = veri(i, 3): Cikis(i, 5) = veri(i, 4)
            x = 6
            If seriVar Then Cikis(i, 6) = veri(i, 7): x = 7
            Cikis(i, x) = veri(i, 5)
            Cikis(i, x + 1) = adet(i)
            Cikis(i, x + 2) = veri(i, 6)
        Next i
        If seriVar Then
            T_ExcelYaz Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Seri numarasý", "Parti", _
                "Tahditsiz klnb.", "Birim", "SAYILAN", "FARK"), Cikis, n, Array(11, 8, 18, 13, 46, 18, 9, 14, 8, 14, 12), 8, ust, depo
        Else
            T_ExcelYaz Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Parti", _
                "Tahditsiz klnb.", "Birim", "SAYILAN", "FARK"), Cikis, n, Array(11, 8, 18, 13, 46, 9, 14, 8, 14, 12), 7, ust, depo
        End If
        Exit Sub
    End If
    If mSayimExcel Then
        depoAd = depo
        If depoAd = "TÜMÜ" Then depoAd = T_IlAdi(ws) & " Tüm Depolar"
        mExcelYol = AKSORA_MASAUSTU() & application.PathSeparator & _
            AKSORA_DOSYA_ADI(depoAd & " - SAYIM - " & T_SayfaKisa(ws.Name) & " - " & tarihMetin) & ".xlsx"
        k = IIf(seriVar, 12, 11)
        ReDim Cikis(1 To n, 1 To k)
        For i = 1 To n
            Cikis(i, 1) = veri(i, 1): Cikis(i, 2) = veri(i, 8): Cikis(i, 3) = veri(i, 2)
            Cikis(i, 4) = veri(i, 3): Cikis(i, 5) = veri(i, 4)
            x = 6
            If seriVar Then Cikis(i, 6) = veri(i, 7): x = 7
            Cikis(i, x) = veri(i, 5)
            Cikis(i, x + 1) = adet(i)
            Cikis(i, x + 2) = veri(i, 6)
        Next i
        mExcelKapat = False
        On Error GoTo SayimHata
        If seriVar Then
            T_SayimExcelYaz Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Seri numarasý", "Parti", _
                "Tahditsiz klnb.", "Birim", "SAYILAN", "FARK", "AÇIKLAMA"), Cikis, n, _
                Array(9.89, 7, 14.66, 11, 34, 16, 7.5, 11, 6.5, 14, 12, 14), 8, ust
        Else
            T_SayimExcelYaz Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Parti", _
                "Tahditsiz klnb.", "Birim", "SAYILAN", "FARK", "AÇIKLAMA"), Cikis, n, _
                Array(9.89, 7, 14.66, 11, 40.55, 7.5, 11, 6.5, 15, 13, 22), 7, ust
        End If
        mExcelYol = vbNullString: mSayimKilit = vbNullString
        Exit Sub
    End If
    If seriVar Then
        T_CiktiYaz ws.Name, Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Seri numarasý", "Parti", _
            "Tahditsiz klnb.", "SAYILAN", "FARK"), Cikis, n, Array(9.89, 7, 14.66, 11, 40.55, 16, 7.5, 13.5, 22, 20), _
            Array("@", "@", "@", "@", "@", "@", "@", "@", "@", "@"), _
            Array(xlCenter, xlCenter, xlLeft, xlCenter, xlLeft, xlCenter, xlCenter, xlRight, xlLeft, xlLeft), 24, True, ust, Array(5, 4, 2)
    Else
        T_CiktiYaz ws.Name, Array("Üretim yeri", "Depo", "Depo yeri tanýmý", "Malzeme", "Malzeme kýsa metni", "Parti", _
            "Tahditsiz klnb.", "SAYILAN", "FARK"), Cikis, n, Array(9.89, 7, 14.66, 11, 40.55, 7.5, 13.5, 28, 26), _
            Array("@", "@", "@", "@", "@", "@", "@", "@", "@"), _
            Array(xlCenter, xlCenter, xlLeft, xlCenter, xlLeft, xlCenter, xlRight, xlLeft, xlLeft), 24, True, ust, Array(5, 4, 2)
    End If
    Exit Sub
SayimHata:
    x = Err.Number: depoAd = Err.Description
    mExcelYol = vbNullString: mSayimKilit = vbNullString
    On Error GoTo 0
    Err.Raise x, , depoAd
End Sub

' Excel sayim formu: PDF sayim listesiyle ayni duzen (A4 yatay, baslik, kase/imza bandi, sayfa no) ama
' SAYILAN ve ACIKLAMA bilgisayarda doldurulur; FARK = SAYILAN - stok kendiliginden. Stok sutunlari kilitli.
' 1. satir baslik, 2. satir AKSORA, 3. satir ilerleme ozeti, 4. satir sutun basliklari (her sayfada), 5+ veri.
Private Sub T_SayimExcelYaz(ByVal bas As Variant, ByRef veri() As Variant, ByVal n As Long, ByVal gen As Variant, _
    ByVal mk As Long, ByVal ust As String)
    Dim wb As Workbook, c As Worksheet, k As Long, j As Long, r As Range, alan As Range, son As Long
    Dim rS As Range, rF As Range, rM As Range, fc As FormatCondition, sA As String, fA As String, mA As String
    k = UBound(bas) - LBound(bas) + 1
    son = n + 4
    AKSORA_DONUS_KAYDET
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set c = wb.Worksheets(1)
    On Error Resume Next
    c.Name = "SAYIM"
    On Error GoTo 0
    c.Cells.Font.Name = "Aptos": c.Cells.Font.Size = 10: c.Cells.Font.Color = RGB(0, 0, 0)
    c.Range("A1").Value2 = "SAYIM LÝSTESÝ  |  " & ust
    c.Range("A1").Font.bold = True: c.Range("A1").Font.Size = 12
    c.Range("A2").Value2 = "Hazýrlama: " & Format$(Now, "dd.mm.yyyy hh:nn") & "  |  " & n & " kalem  |  Sarý hücrelere sayýlan miktarý girin  |  AKSORA"
    c.Range("A2").Font.Size = 9: c.Range("A2").Font.Color = RGB(89, 89, 89)
    For j = 1 To k
        c.Columns(j).ColumnWidth = gen(LBound(gen) + j - 1)
        c.Cells(4, j).Value2 = bas(LBound(bas) + j - 1)
        If j < mk Or j = mk + 1 Or j = mk + 4 Then c.Range(c.Cells(5, j), c.Cells(son, j)).NumberFormat = "@"
    Next j
    For j = 1 To n
        If T_SadeSayi(veri(j, 1)) Then veri(j, 1) = CDbl(veri(j, 1))
        If T_SadeSayi(veri(j, 4)) Then veri(j, 4) = CDbl(veri(j, 4))
        If IsNumeric(veri(j, mk)) Then veri(j, mk) = Round(CDbl(veri(j, mk)), 3)
    Next j
    c.Range("A5").Resize(n, k).Value2 = veri
    Set r = c.Range(c.Cells(5, 1), c.Cells(son, k))
    If n > 1 Then
        r.Sort Key1:=r.Columns(5), Order1:=xlAscending, Key2:=r.Columns(4), Order2:=xlAscending, _
            Key3:=r.Columns(2), Order3:=xlAscending, Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom
    End If
    Set rM = c.Range(c.Cells(5, mk), c.Cells(son, mk))
    Set rS = c.Range(c.Cells(5, mk + 2), c.Cells(son, mk + 2))
    Set rF = c.Range(c.Cells(5, mk + 3), c.Cells(son, mk + 3))
    mA = rM.Address: sA = rS.Address: fA = rF.Address
    rS.ClearContents
    rF.FormulaR1C1 = "=IF(RC[-1]="""","""",ROUND(RC[-1]-RC[-3],3))"
    ' Miktarlar Genel bicim: tam sayida virgul yok, kusurat varsa virgulle (Excel'e aktar ile ayni).
    ' Kosullu bicim yalniz sabit sayiyla (formul kullanilmaz: Turkce / Ingilizce Excel'de ayni calisir).
    rM.NumberFormat = "General": rS.NumberFormat = "General": rF.NumberFormat = "General"
    Set fc = rF.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    fc.Font.Color = RGB(192, 0, 0): fc.Font.bold = True
    Set fc = rF.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    fc.Font.Color = RGB(0, 112, 60): fc.Font.bold = True
    ' Giris hucreleri: ekranda acik sari (baskida siyah-beyaz), yalniz 0 veya pozitif sayi.
    With c.Range(c.Cells(5, mk + 2), c.Cells(son, mk + 2))
        .Interior.Color = RGB(255, 247, 204)
        .Validation.Delete
        .Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, Formula1:="0"
        .Validation.ErrorTitle = "AKSORA": .Validation.ErrorMessage = "Sayýlan miktar 0 veya pozitif bir sayý olmalýdýr."
        .Validation.InputTitle = "SAYILAN": .Validation.InputMessage = "Depoda sayýlan miktarý girin."
    End With
    c.Range(c.Cells(5, mk + 4), c.Cells(son, mk + 4)).Interior.Color = RGB(255, 251, 230)
    c.Range("A3").Formula = "=""Girilen: ""&COUNT(" & sA & ")&"" / " & n & " kalem   |   Farklý çýkan: ""&SUMPRODUCT((" & sA & "<>"""")*(" & fA & "<>0))&"" kalem"""
    c.Range("A3").Font.Size = 9: c.Range("A3").Font.bold = True: c.Range("A3").Font.Color = RGB(31, 78, 121)
    Set alan = c.Range(c.Cells(4, 1), c.Cells(son, k))
    With alan
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(0, 0, 0)
        .VerticalAlignment = xlCenter
    End With
    With c.Range(c.Cells(4, 1), c.Cells(4, k))
        .Interior.Color = RGB(217, 217, 217): .Font.bold = True
        .HorizontalAlignment = xlCenter: .WrapText = True
    End With
    c.Rows(4).RowHeight = 30
    c.Range(c.Cells(5, 1), c.Cells(son, 1)).EntireRow.RowHeight = 22
    c.Range(c.Cells(5, 1), c.Cells(son, 4)).HorizontalAlignment = xlCenter
    c.Range(c.Cells(5, 3), c.Cells(son, 3)).HorizontalAlignment = xlLeft
    c.Range(c.Cells(5, mk - 1), c.Cells(son, mk - 1)).HorizontalAlignment = xlCenter
    If mk = 8 Then c.Range(c.Cells(5, 6), c.Cells(son, 6)).HorizontalAlignment = xlCenter
    c.Range(c.Cells(5, mk), c.Cells(son, mk)).HorizontalAlignment = xlRight
    c.Range(c.Cells(5, mk + 1), c.Cells(son, mk + 1)).HorizontalAlignment = xlCenter
    c.Range(c.Cells(5, mk + 2), c.Cells(son, mk + 3)).HorizontalAlignment = xlRight
    c.Range(c.Cells(5, mk + 2), c.Cells(son, mk + 3)).Font.bold = True
    c.Range(c.Cells(5, mk + 4), c.Cells(son, mk + 4)).HorizontalAlignment = xlLeft
    alan.AutoFilter
    ' Kilit (AKSORA imzasiyla birlikte): baslik satirlari, stok sutunlari ve FARK; serbest: SAYILAN, ACIKLAMA.
    mSayimKilit = application.Union(c.Range("A3").Resize(2, k), c.Range(c.Cells(5, 1), c.Cells(son, mk + 1)), rF).Address
    On Error Resume Next
    application.PrintCommunication = False
    With c.PageSetup
        .PrintArea = alan.Address
        .PrintTitleRows = "$4:$4"
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .BlackAndWhite = True
    End With
    T_CiktiSayfa c.PageSetup, ust, True
    application.PrintCommunication = True
    c.PageSetup.LeftFooter = "&""Aptos""&7AKSORA  •  Çýktý: " & Format$(Now, "dd.mm.yyyy hh:nn")
    On Error GoTo 0
    T_ExcelBitir wb, c, 5
    mSayimKilit = vbNullString
    On Error Resume Next
    c.Cells(5, mk + 2).Select
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
    Dim kaynak As Variant, veri() As Variant, gorV As Variant, bos As Boolean, ust As String, kars As Boolean
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
        r = AKSORA_SON_SATIR(ws, kol(j))
        If r > son Then son = r
    Next j
    If son > hdr Then gorV = AKSORA_GORUNUR(ws, hdr + 1, son, kol(k))
    If IsArray(gorV) Then
        kaynak = ws.Range(ws.Cells(hdr + 1, 1), ws.Cells(son, kol(k))).Value2
        ReDim veri(1 To son - hdr, 1 To k)
        For i = 1 To son - hdr
            If gorV(i) = 1 Then
                bos = True
                For j = 1 To k
                    If IsError(kaynak(i, kol(j))) Then
                        veri(n + 1, j) = "#HATA"
                    Else
                        veri(n + 1, j) = kaynak(i, kol(j))
                    End If
                    If Len(Trim$(CStr(veri(n + 1, j)))) > 0 Then bos = False
                Next j
                If Not bos Then n = n + 1
            End If
        Next i
    End If
    If n = 0 Then Err.Raise vbObjectError + 9603, , "Çýktý alýnacak satýr yok."
    ReDim bas(0 To k - 1): ReDim gen(0 To k - 1): ReDim fmt(0 To k - 1): ReDim hiz(0 To k - 1)
    For j = 1 To k
        bas(j - 1) = Trim$(CStr(ws.Cells(hdr, kol(j)).Value2))
        gen(j - 1) = ws.Columns(kol(j)).ColumnWidth
        fmt(j - 1) = ws.Cells(hdr + 1, kol(j)).NumberFormat
        If IsNull(fmt(j - 1)) Then fmt(j - 1) = "General"
        ' Ekrandaki secim oku (metin + dolgu + ok) ciktiya tasinmaz.
        If Left$(CStr(fmt(j - 1)), 2) = "@*" Then fmt(j - 1) = "@"
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
                If s.top + s.Height > altY Then altY = s.top + s.Height
            End If
        End If
    Next s
    For j = 1 To 30
        If Not ws.Columns(j).Hidden Then
            If application.WorksheetFunction.CountA(ws.Columns(j)) > 0 Then sonK = j
            If ws.Columns(j).Left < sagX - 1 Then If j > sonK Then sonK = j
            r = AKSORA_SON_SATIR(ws, j)
            If r > sonS And r < 5000 Then sonS = r
        End If
    Next j
    r = 1
    Do While ws.Rows(r).top + ws.Rows(r).Height < altY And r < 5000
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

' Kase/imza bandi (seffaf zeminli PNG, 27,7 x 1,3 cm, 300 dpi; yazilar Cikti/Sayfa boyutunda) kodun icinde; gecici klasore bir kez yazilir.
Private Function T_CiktiImzaResmi() As String
    Dim s As String, yol As String, b() As Byte, d As Object, e As Object, f As Integer, i As Long, p As String
    On Error GoTo Yok
    yol = Environ$("TEMP")
    If Len(yol) = 0 Then yol = AKSORA_KAYIT_KLASORU()
    yol = yol & application.PathSeparator & "AKSORA_IMZA_BANDI_V3.png"
    If Len(Dir$(yol)) > 0 Then T_CiktiImzaResmi = yol: Exit Function
    ' Resim (base64) kitabin gizli adlarinda (_AKS_IMZA_001...); kod metninde tutulmaz:
    ' buyuk rastgele metin Excel kaydinda modul kaynagini bozuyordu.
    For i = 1 To 999
        p = vbNullString
        On Error Resume Next
        p = ThisWorkbook.names("_AKS_IMZA_" & Format$(i, "000")).RefersTo
        On Error GoTo Yok
        If Len(p) < 4 Then Exit For
        s = s & Mid$(p, 3, Len(p) - 3)
    Next i
    If Len(s) = 0 Then GoTo Yok
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
    AKSORA_DONUS_KAYDET
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

' Yeni (gorunur, kaydedilmemis) kitapta tablo: 1. satir baslik, 2. satir hazirlama tarihi, 4. satir sutun basliklari.
' Satirlar malzeme adina gore A-Z (sonra malzeme kodu, depo). mk = miktar sutunu.
Private Sub T_ExcelYaz(ByVal bas As Variant, ByRef veri() As Variant, ByVal n As Long, ByVal gen As Variant, _
    ByVal mk As Long, ByVal ust As String, ByVal depo As String)
    Dim wb As Workbook, c As Worksheet, k As Long, j As Long, r As Range, alan As Range, ad As String, x As Variant
    k = UBound(bas) - LBound(bas) + 1
    AKSORA_DONUS_KAYDET
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set c = wb.Worksheets(1)
    ad = Left$(depo, 28)
    For Each x In Array("/", "\", ":", "*", "?", "[", "]")
        ad = Replace(ad, CStr(x), "-")
    Next x
    On Error Resume Next
    c.Name = IIf(Len(ad) > 0, ad, "STOK")
    On Error GoTo 0
    c.Cells.Font.Name = "Aptos": c.Cells.Font.Size = 10
    c.Range("A1").Value2 = ust
    c.Range("A1").Font.bold = True: c.Range("A1").Font.Size = 12
    c.Range("A2").Value2 = "Hazýrlama: " & Format$(Now, "dd.mm.yyyy hh:nn") & "  |  " & n & " kalem  |  AKSORA"
    c.Range("A2").Font.Size = 9: c.Range("A2").Font.Color = RGB(89, 89, 89)
    For j = 1 To k
        c.Columns(j).ColumnWidth = gen(LBound(gen) + j - 1)
        c.Cells(4, j).Value2 = bas(LBound(bas) + j - 1)
        ' Uretim yeri ve malzeme kodu: sade rakamlar sayi olarak yazilir (yesil uyari cikmaz); bastaki sifirli
        ' veya uzun kodlar metin kalir (sutun metin bicimli; sayiya donusmez, kod bozulmaz).
        If j < mk Or j = mk + 1 Then
            c.Range(c.Cells(5, j), c.Cells(n + 4, j)).NumberFormat = "@"
        End If
    Next j
    For j = 1 To n
        If T_SadeSayi(veri(j, 1)) Then veri(j, 1) = CDbl(veri(j, 1))
        If T_SadeSayi(veri(j, 4)) Then veri(j, 4) = CDbl(veri(j, 4))
        If IsNumeric(veri(j, mk)) Then veri(j, mk) = Round(CDbl(veri(j, mk)), 3)
    Next j
    c.Range("A5").Resize(n, k).Value2 = veri
    Set r = c.Range(c.Cells(5, 1), c.Cells(n + 4, k))
    If n > 1 Then
        r.Sort Key1:=r.Columns(5), Order1:=xlAscending, Key2:=r.Columns(4), Order2:=xlAscending, _
            Key3:=r.Columns(2), Order3:=xlAscending, Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom
    End If
    ' Miktar: binlik ayraci yok; kusurat varsa virgulle (Genel bicim).
    c.Range(c.Cells(5, mk), c.Cells(n + 4, mk)).NumberFormat = "General"
    c.Range(c.Cells(5, mk + 3), c.Cells(n + 4, mk + 3)).FormulaR1C1 = "=IF(RC[-1]="""","""",RC[-1]-RC[-3])"
    c.Range(c.Cells(5, mk + 2), c.Cells(n + 4, mk + 3)).NumberFormat = "General"
    c.Range(c.Cells(5, mk), c.Cells(n + 4, mk)).HorizontalAlignment = xlRight
    Set alan = c.Range(c.Cells(4, 1), c.Cells(n + 4, k))
    With alan
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(0, 0, 0)
        .VerticalAlignment = xlCenter
    End With
    With c.Range(c.Cells(4, 1), c.Cells(4, k))
        .Interior.Color = RGB(217, 217, 217): .Font.bold = True
        .HorizontalAlignment = xlCenter: .WrapText = True
    End With
    c.Rows(4).RowHeight = 30
    c.Range(c.Cells(5, 1), c.Cells(n + 4, 4)).HorizontalAlignment = xlCenter
    c.Range(c.Cells(5, 3), c.Cells(n + 4, 3)).HorizontalAlignment = xlLeft
    c.Range(c.Cells(5, mk - 1), c.Cells(n + 4, mk - 1)).HorizontalAlignment = xlCenter
    c.Range(c.Cells(5, mk + 1), c.Cells(n + 4, mk + 1)).HorizontalAlignment = xlCenter
    alan.AutoFilter
    On Error Resume Next
    application.PrintCommunication = False
    With c.PageSetup
        .PrintArea = alan.Address
        .PrintTitleRows = "$4:$4"
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    T_CiktiSayfa c.PageSetup, ust, False
    application.PrintCommunication = True
    On Error GoTo 0
    T_ExcelBitir wb, c, 5
End Sub

' Yeni kitap: basliklar dondurulur, masaustune / klasore kaydedilir (mExcelYol); toplu aktarimda kapatilir,
' tek aktarimda normal Excel gorunumunde acik kalir.
Private Sub T_ExcelBitir(ByVal wb As Workbook, ByVal c As Worksheet, ByVal ilkSatir As Long)
    T_AksoraImza wb
    wb.Windows(1).Visible = True
    wb.Activate
    c.Activate
    On Error Resume Next
    With wb.Windows(1)
        .DisplayHeadings = True: .DisplayGridlines = True: .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True: .DisplayVerticalScrollBar = True
        .ScrollRow = 1: .ScrollColumn = 1
        .FreezePanes = False
        c.Cells(ilkSatir, 1).Select
        .FreezePanes = True
        .Zoom = 100
    End With
    On Error GoTo 0
    mExcelSonuc = vbNullString
    If Len(mExcelYol) > 0 Then mExcelSonuc = T_Kaydet(wb, mExcelYol)
    If mExcelKapat Then
        wb.Close SaveChanges:=False
        Exit Sub
    End If
    ' Yeni kitap normal Excel gorunumunde (serit, formul cubugu, sistem ayraclari); AKSORA penceresine dokunulmaz.
    ' Pencere gecisleri izlenir: AKSORA'ya donulunce uygulama modu yeniden acilir.
    On Error Resume Next
    CallByName ThisWorkbook, "AKSORA_UYG_IZLE", VbMethod
    AKSORA_EXCEL_GORUNUM
    On Error GoTo 0
    application.ScreenUpdating = True
End Sub

' Kitap .xlsx olarak kaydedilir; ayni adli dosya acik / kilitliyse "(2)", "(3)" ... denenir. Donus: kaydedilen yol.
Private Function T_Kaydet(ByVal wb As Workbook, ByVal yol As String) As String
    Dim da As Boolean, i As Long, deneme As String, kok As String
    da = application.DisplayAlerts: application.DisplayAlerts = False
    kok = Left$(yol, Len(yol) - 5)
    deneme = yol
    On Error Resume Next
    For i = 1 To 20
        Err.Clear
        wb.SaveAs Filename:=deneme, FileFormat:=51
        If Err.Number = 0 Then T_Kaydet = deneme: Exit For
        deneme = kok & " (" & (i + 1) & ").xlsx"
    Next i
    On Error GoTo 0
    application.DisplayAlerts = da
End Function

' Kullanicinin masaustu klasoru (OneDrive yonlendirmesi dahil).
Public Function AKSORA_MASAUSTU() As String
    On Error Resume Next
    AKSORA_MASAUSTU = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    If Len(AKSORA_MASAUSTU) = 0 Then AKSORA_MASAUSTU = Environ$("USERPROFILE") & application.PathSeparator & "Desktop"
    If Not CreateObject("Scripting.FileSystemObject").FolderExists(AKSORA_MASAUSTU) Then AKSORA_MASAUSTU = AKSORA_KAYIT_KLASORU()
End Function

' Dosya / klasor adinda kullanilamayan karakterler cikarilir.
Public Function AKSORA_DOSYA_ADI(ByVal s As String) As String
    Dim x As Variant
    For Each x In Array("\", "/", ":", "*", "?", """", "<", ">", "|", vbTab, vbLf, vbCr)
        s = Replace(s, CStr(x), "-")
    Next x
    s = Trim$(s)
    Do While Right$(s, 1) = "." And Len(s) > 0
        s = Left$(s, Len(s) - 1)
    Loop
    If Len(s) > 90 Then s = Trim$(Left$(s, 90))
    If Len(s) = 0 Then s = "AKSORA"
    AKSORA_DOSYA_ADI = s
End Function

' Klasor yoksa olusturulur. Donus: basarili mi.
Public Function AKSORA_KLASOR_HAZIRLA(ByVal yol As String) As Boolean
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(yol) Then fso.CreateFolder yol
    AKSORA_KLASOR_HAZIRLA = fso.FolderExists(yol)
End Function

' Excel'e aktarilan tablo icin genel kitap (demontaj eksikleri vb.): 1. satir baslik, 2. satir bilgi, 4. satir sutunlar.
' bicim: sutun bicimleri ("@", "General", "0", "#,##0.00 ""TL"""); hiza: xlLeft / xlCenter / xlRight; sira: A-Z sutunlari.
' Altin kural: tutar sutunlari her zaman binlik nokta, kurus virgul ve sonda TL (TL'siz tutar bicimi verilse de eklenir).
' Donus: kaydedilen dosya yolu.
Public Function AKSORA_EXCEL_TABLO(ByVal bas As Variant, ByRef veri() As Variant, ByVal n As Long, ByVal gen As Variant, _
    ByVal bicim As Variant, ByVal hiza As Variant, ByVal ust As String, ByVal bilgi As String, ByVal sayfaAd As String, _
    ByVal yol As String, ByVal kapat As Boolean, ByVal sira As Variant) As String
    Dim wb As Workbook
    AKSORA_DONUS_KAYDET
    Set wb = Workbooks.Add(xlWBATWorksheet)
    AKSORA_EXCEL_SAYFA wb.Worksheets(1), bas, veri, n, gen, bicim, hiza, ust, bilgi, sayfaAd, sira
    mExcelYol = yol: mExcelKapat = kapat
    T_ExcelBitir wb, wb.Worksheets(1), 5
    AKSORA_EXCEL_TABLO = mExcelSonuc
    mExcelYol = vbNullString: mExcelKapat = False
End Function

' Verilen sayfaya Excel tablosu: 1. satir baslik, 2. satir bilgi, 4. satir sutunlar, 5. satirdan veri.
Public Sub AKSORA_EXCEL_SAYFA(ByVal c As Worksheet, ByVal bas As Variant, ByRef veri() As Variant, ByVal n As Long, _
    ByVal gen As Variant, ByVal bicim As Variant, ByVal hiza As Variant, ByVal ust As String, ByVal bilgi As String, _
    ByVal sayfaAd As String, ByVal sira As Variant)
    Dim k As Long, j As Long, r As Range, alan As Range, i As Long
    k = UBound(bas) - LBound(bas) + 1
    On Error Resume Next
    c.Name = Left$(AKSORA_DOSYA_ADI(Replace(Replace(sayfaAd, "[", "("), "]", ")")), 31)
    On Error GoTo 0
    c.Cells.Font.Name = "Aptos": c.Cells.Font.Size = 10
    c.Range("A1").Value2 = ust
    c.Range("A1").Font.bold = True: c.Range("A1").Font.Size = 12
    c.Range("A2").Value2 = "Hazýrlama: " & Format$(Now, "dd.mm.yyyy hh:nn") & "  |  " & n & " kalem" & _
        IIf(Len(bilgi) > 0, "  |  " & bilgi, "") & "  |  AKSORA"
    c.Range("A2").Font.Size = 9: c.Range("A2").Font.Color = RGB(89, 89, 89)
    For j = 1 To k
        c.Columns(j).ColumnWidth = gen(LBound(gen) + j - 1)
        c.Cells(4, j).Value2 = bas(LBound(bas) + j - 1)
        ' "0" (kod) sutunu metin bicimli yazilir; sade rakamlar asagida sayiya cevrilir (bastaki sifir korunur).
        c.Range(c.Cells(5, j), c.Cells(n + 4, j)).NumberFormat = T_ExcelBicim(CStr(bicim(LBound(bicim) + j - 1)))
        c.Range(c.Cells(5, j), c.Cells(n + 4, j)).HorizontalAlignment = hiza(LBound(hiza) + j - 1)
        If bicim(LBound(bicim) + j - 1) <> "@" Then
            For i = 1 To n
                If T_SadeSayi(veri(i, j)) Then veri(i, j) = CDbl(veri(i, j))
            Next i
        End If
    Next j
    c.Range("A5").Resize(n, k).Value2 = veri
    Set r = c.Range(c.Cells(5, 1), c.Cells(n + 4, k))
    If n > 1 And IsArray(sira) Then
        r.Sort Key1:=r.Columns(sira(0)), Order1:=xlAscending, Key2:=r.Columns(sira(1)), Order2:=xlAscending, _
            Key3:=r.Columns(sira(2)), Order3:=xlAscending, Header:=xlNo, MatchCase:=False, Orientation:=xlTopToBottom
    End If
    Set alan = c.Range(c.Cells(4, 1), c.Cells(n + 4, k))
    With alan
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(0, 0, 0)
        .VerticalAlignment = xlCenter
    End With
    With c.Range(c.Cells(4, 1), c.Cells(4, k))
        .Interior.Color = RGB(217, 217, 217): .Font.bold = True
        .HorizontalAlignment = xlCenter: .WrapText = True
    End With
    c.Rows(4).RowHeight = 30
    alan.AutoFilter
    On Error Resume Next
    application.PrintCommunication = False
    With c.PageSetup
        .PrintArea = alan.Address
        .PrintTitleRows = "$4:$4"
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
    End With
    T_CiktiSayfa c.PageSetup, ust, False
    application.PrintCommunication = True
    On Error GoTo 0
End Sub

' Cok sayfali kitap: her sayfada basliklar dondurulur, ilk sayfa acik kaydedilir (masaustu). Donus: kaydedilen yol.
Public Function AKSORA_EXCEL_BITIR(ByVal wb As Workbook, ByVal yol As String, ByVal kapat As Boolean) As String
    Dim i As Long
    wb.Windows(1).Visible = True
    wb.Activate
    On Error Resume Next
    For i = wb.Worksheets.Count To 2 Step -1
        wb.Worksheets(i).Activate
        With wb.Windows(1)
            .ScrollRow = 1: .ScrollColumn = 1
            .FreezePanes = False
            wb.Worksheets(i).Cells(5, 1).Select
            .FreezePanes = True
            .Zoom = 100
        End With
    Next i
    On Error GoTo 0
    mExcelYol = yol: mExcelKapat = kapat
    T_ExcelBitir wb, wb.Worksheets(1), 5
    AKSORA_EXCEL_BITIR = mExcelSonuc
    mExcelYol = vbNullString: mExcelKapat = False
End Function

' Excel sutun bicimi: kod sutunu ("0") metin yazilir (bastaki sifir korunur); tutar bicimi daima TL ile biter.
Private Function T_ExcelBicim(ByVal b As String) As String
    If b = "0" Then T_ExcelBicim = "@": Exit Function
    If InStr(b, "0.00") > 0 And InStr(b, "TL") = 0 Then b = b & " ""TL"""
    T_ExcelBicim = b
End Function

' Sayfa basligi (cikti / Excel): birim adiyla, "I_" / "M_" / "Y_" onekleri olmadan.
Private Function T_SayfaBaslik(ByVal ad As String) As String
    Select Case Left$(ad, 2)
        Case "I_": T_SayfaBaslik = "ÝLÇE YÖNETÝM  |  " & Replace(Mid$(ad, 3), "_", " ")
        Case "M_": T_SayfaBaslik = "MÜÞTERÝ  |  " & Replace(Mid$(ad, 3), "_", " ")
        Case "Y_": T_SayfaBaslik = "YATIRIM  |  " & Replace(Mid$(ad, 3), "_", " ")
        Case Else: T_SayfaBaslik = Replace(ad, "_", " ")
    End Select
End Function

' Dosya adi icin kisa sayfa adi: "Ýlçe BÝRLEÞÝK STOK", "Müþteri CDP STOK" ...
Private Function T_SayfaKisa(ByVal ad As String) As String
    Select Case Left$(ad, 2)
        Case "I_": T_SayfaKisa = "Ýlçe " & Replace(Mid$(ad, 3), "_", " ")
        Case "M_": T_SayfaKisa = "Müþteri " & Replace(Mid$(ad, 3), "_", " ")
        Case "Y_": T_SayfaKisa = "Yatýrým " & Replace(Mid$(ad, 3), "_", " ")
        Case Else: T_SayfaKisa = Replace(ad, "_", " ")
    End Select
End Function

' Il adi (sayfanin A5 hucresi: "Artvin Ýl Müdürlüðü" -> "Artvin").
Private Function T_IlAdi(ByVal ws As Worksheet) As String
    Dim t As String
    t = Trim$(CStr(ws.Range("A5").Value2))
    If InStr(t, " ") > 0 Then t = Left$(t, InStr(t, " ") - 1)
    If Len(t) = 0 Then t = "AKSORA"
    T_IlAdi = t
End Function

' Tum depolari ayri dosyaya aktarma: ilce ve musteri birimlerinin stok sayfalari.
Private Function T_TopluSayfa(ByVal ad As String) As Boolean
    If Right$(ad, 5) <> "_STOK" Then Exit Function
    T_TopluSayfa = (Left$(ad, 2) = "I_" Or Left$(ad, 2) = "M_")
End Function

' Yalniz rakamdan olusan, 0 ile baslamayan, 15 haneyi gecmeyen metin sayiya cevrilebilir.
Private Function T_SadeSayi(ByVal v As Variant) As Boolean
    Dim s As String
    If IsError(v) Then Exit Function
    s = CStr(v)
    If Len(s) = 0 Or Len(s) > 15 Or Left$(s, 1) = "0" Then Exit Function
    T_SadeSayi = Not (s Like "*[!0-9]*")
End Function

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
' Cikti / Excel'e aktar satiri: HURDA hic alinmaz; CIKMA/BDSZ yalniz trafo, hucre, pano ve direkte alinir.
Private Function T_CiktiDahil(ByVal parti As String, ByVal malGrubu As Variant, ByVal tanim As Variant) As Boolean
    Dim t As String
    If Not T_CiktiDisi(parti) Then T_CiktiDahil = True: Exit Function
    t = UCase$(parti)
    If InStr(t, "HURDA") > 0 Then Exit Function
    T_CiktiDahil = T_CikmaKritik(T_CiktiMetin(malGrubu), T_CiktiMetin(tanim))
End Function

' Trafo (M0340), hucre (M0314), pano (M0327), direk (M0308); mal grubu yoksa tanimin basi.
Private Function T_CikmaKritik(ByVal mg As String, ByVal tn As String) As Boolean
    mg = UCase$(Trim$(mg)): tn = UCase$(Trim$(tn))
    If Left$(tn, 6) = "IPTAL." Then tn = Trim$(Mid$(tn, 7))
    Select Case True
        Case Left$(mg, 5) = "M0340", Left$(mg, 5) = "M0314", Left$(mg, 5) = "M0327", Left$(mg, 5) = "M0308"
            T_CikmaKritik = True
        Case Len(mg) > 0
            T_CikmaKritik = False
        Case Left$(tn, 6) = "TRAFO ", Left$(tn, 6) = "HUCRE ", Left$(tn, 5) = "PANO ", Left$(tn, 6) = "DIREK "
            T_CikmaKritik = True
    End Select
End Function

Private Function T_CiktiDisi(ByVal parti As String) As Boolean
    Dim t As String, x As Variant
    t = UCase$(parti)
    For Each x In Array(" ", "_", "-", ".", "/")
        t = Replace(t, CStr(x), vbNullString)
    Next x
    t = Replace(Replace(t, "Ç", "C"), "Ý", "I")
    T_CiktiDisi = (InStr(t, "HURDA") > 0 Or InStr(t, "CIKMA") > 0 Or InStr(t, "BDSZ") > 0 Or InStr(t, "BEDELSIZ") > 0)
End Function

' AKSORA imzasi (butun Excel'e aktar ciktilari): baslik ve "AKSORA" satiri (1-2) kilitli, sayfa altinda AKSORA,
' dosya ozelliklerinde ve gizli imza sayfasinda uretim kaydi. Veriler serbest: filtre, siralama, bicim, satir ekle/sil calisir.
' Butun AKSORA sayfalarinin ekran goruntusu (katalog / destek icin): her sayfa ekranda gorundugu alanla JPG olarak
' masaustunde tek klasore kaydedilir. Gizli sayfalar gecici gosterilir, sonra eski haline doner; kayit sayfalari (_) alinmaz.
' Calistirma: Alt+F8 > AKSORA_EKRAN_GORUNTULERI.
Public Sub AKSORA_EKRAN_GORUNTULERI()
    Dim ws As Worksheet, bas As Worksheet, klasor As String, n As Long, top As Long, i As Long, j As Long, hata As String
    Dim rg As Range, tmp As Workbook, ch As ChartObject, yol As String, gor As XlSheetVisibility, ev As Boolean, ac As Boolean
    Dim ad() As String, an() As Double, tmpS As String, tmpD As Double, grup As String
    If MsgBox("AKSORA'nýn bütün sayfalarýnýn ekran görüntüsü alýnacak (birim birim sýrayla, JPG)." & vbCrLf & _
        "Dosyalar masaüstünde tek klasöre kaydedilir. Devam edilsin mi?", vbQuestion + vbYesNo, "AKSORA | EKRAN GÖRÜNTÜLERÝ") <> vbYes Then Exit Sub
    Set bas = ActiveSheet
    klasor = AKSORA_MASAUSTU() & application.PathSeparator & AKSORA_DOSYA_ADI("AKSORA Ekran Goruntuleri " & Format$(Now, "yyyy-mm-dd hh.nn"))
    If Not AKSORA_KLASOR_HAZIRLA(klasor) Then MsgBox "Klasör oluþturulamadý: " & klasor, vbExclamation, "AKSORA": Exit Sub
    ' Sira: birim secimi, Malzeme, Yatirim, Musteri, Ilce, Admin; her birimde il secimi, ana panel, sonra sekme sirasi.
    For Each ws In ThisWorkbook.Worksheets
        If T_EkranAlinir(ws) Then
            top = top + 1
            ReDim Preserve ad(1 To top): ReDim Preserve an(1 To top)
            ad(top) = ws.Name: an(top) = T_EkranSira(ws.Name) * 1000# + ws.Index
        End If
    Next ws
    If top = 0 Then Exit Sub
    For i = 1 To top - 1
        For j = i + 1 To top
            If an(j) < an(i) Then
                tmpD = an(i): an(i) = an(j): an(j) = tmpD
                tmpS = ad(i): ad(i) = ad(j): ad(j) = tmpS
            End If
        Next j
    Next i
    ev = application.EnableEvents
    application.EnableEvents = False
    ac = AKSORA_BEKLE_AC("AKSORA  |  EKRAN GÖRÜNTÜLERÝ", "Sayfalar hazýrlanýyor...", 0.02)
    Set tmp = Workbooks.Add(xlWBATWorksheet)
    ThisWorkbook.Activate
    For i = 1 To top
        Set ws = ThisWorkbook.Worksheets(ad(i))
        grup = Choose(Int(an(i) / 10000#) + 1, "0 GIRIS", "1 MALZEME", "2 YATIRIM", "3 MUSTERI", "4 ILCE", "5 ADMIN")
        If ac Then AKSORA_BEKLE_ILERLE 0.02 + 0.96 * i / top, i & " / " & top & "  •  " & ws.Name
        yol = vbNullString
        On Error Resume Next
        Err.Clear
        gor = ws.Visible
        If gor <> xlSheetVisible Then ws.Visible = xlSheetVisible
        ws.Activate
        application.ScreenUpdating = True
        ' Sayfa acilisindaki duzen (yakinlastirma, rapor bolumlerinin yerlesimi) olay kapali oldugu icin burada calistirilir.
        If Left$(ws.Name, 2) = "Y_" Or Left$(ws.Name, 2) = "I_" Then
            AKSORA_YATIRIM_GORUNUM ws
        ElseIf Left$(ws.Name, 2) <> "M_" Then
            AKSORA_EKRAN_SIGDIR ws
        End If
        Err.Clear
        ' Duzen yordami baska sayfaya gecirdiyse (panel yenileme vb.) sayfaya geri donulur.
        If Not (ActiveSheet Is ws) Then ws.Activate
        DoEvents
        ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
        Set rg = Nothing
        Set rg = T_EkranAlan(ws)
        If rg Is Nothing Then Set rg = ws.Range(ws.Cells(1, 1), ws.Cells(ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count - 1, ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count - 1))
        Err.Clear
        rg.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        If Err.Number <> 0 Then
            Err.Clear
            Set rg = ws.Range(ws.Cells(1, 1), ws.Cells(ActiveWindow.VisibleRange.Row + ActiveWindow.VisibleRange.Rows.Count - 1, ActiveWindow.VisibleRange.Column + ActiveWindow.VisibleRange.Columns.Count - 1))
            rg.CopyPicture Appearance:=xlScreen, Format:=xlBitmap
        End If
        If Err.Number = 0 Then
            yol = klasor & application.PathSeparator & Format$(i, "000") & " - " & grup & " - " & AKSORA_DOSYA_ADI(ws.Name) & ".jpg"
            tmp.Activate
            Set ch = tmp.Worksheets(1).ChartObjects.Add(0, 0, rg.Width, rg.Height)
            ch.Activate
            ch.Chart.Paste
            ch.Chart.Export Filename:=yol, FilterName:="JPG"
            ch.Delete
            ThisWorkbook.Activate
        End If
        If Err.Number <> 0 Or Len(yol) = 0 Then
            hata = hata & IIf(Len(hata) > 0, ", ", "") & ws.Name
        ElseIf Len(Dir$(yol)) = 0 Then
            hata = hata & IIf(Len(hata) > 0, ", ", "") & ws.Name
        Else
            n = n + 1
        End If
        If gor <> xlSheetVisible Then ws.Visible = gor
        Err.Clear
        On Error GoTo 0
    Next i
    On Error Resume Next
    application.CutCopyMode = False
    tmp.Close SaveChanges:=False
    ThisWorkbook.Activate
    bas.Activate
    On Error GoTo 0
    application.EnableEvents = ev
    If ac Then AKSORA_BEKLE_KAPAT
    MsgBox n & " sayfanýn ekran görüntüsü kaydedildi:" & vbCrLf & klasor & _
        IIf(Len(hata) > 0, vbCrLf & vbCrLf & "Alýnamayan: " & Left$(hata, 500), ""), vbInformation, "AKSORA | EKRAN GÖRÜNTÜLERÝ"
End Sub

' Ekran goruntusu sirasi: grup x 10 + birim ici sira (il secimi 0, ana panel 1, diger sayfalar 2).
Private Function T_EkranSira(ByVal ad As String) As Long
    Dim g As Long, k As Long
    k = 2
    Select Case True
        Case ad = "BIRIM SECIM": g = 0: k = 0
        Case ad = "ADMIN PANEL": g = 5
        Case ad = "YATIRIM IL SECIM": g = 2: k = 0
        Case ad = "MUSTERI IL SECIM": g = 3: k = 0
        Case ad = "ILCE IL SECIM": g = 4: k = 0
        Case Left$(ad, 2) = "Y_": g = 2
        Case Left$(ad, 2) = "M_": g = 3
        Case Left$(ad, 2) = "I_": g = 4
        Case ad = "KAPAK": g = 1: k = 0
        Case Else: g = 1
    End Select
    If k = 2 And InStr(1, ad, "ANA PANEL", vbTextCompare) > 0 Then k = 1
    T_EkranSira = g * 10 + k
End Function

' Ekran goruntusu alinacak sayfa: kayit sayfalari (_) ve "Islem ekranlari hazirlaniyor" yer tutucu birim sayfalari haric.
Private Function T_EkranAlinir(ByVal ws As Worksheet) As Boolean
    Dim s As Shape
    If Left$(ws.Name, 1) = "_" Then Exit Function
    If ws.Name = "YATIRIM PANEL" Or ws.Name = "MUSTERI PANEL" Then Exit Function
    On Error Resume Next
    Set s = ws.Shapes("PANEL_DURUM")
    On Error GoTo 0
    T_EkranAlinir = (s Is Nothing)
End Function

' Goruntu alani: A1'den ekran genisligine kadar (sagdaki gizli veri sutunlari alinmaz). Stok listeleri ve arama sayfalarinda
' ilk ekran; ADMIN PANEL sonuna kadar (en fazla 400 satir); diger panel, rapor ve analiz sayfalari en fazla 150 satir.
' Korumali sayfada Find calismadigi icin son dolu satir satir satir sayilir; bir seklin hatasi alani bozmaz.
Private Function T_EkranAlan(ByVal ws As Worksheet) As Range
    Dim vr As Range, sonC As Long, sonR As Long, ilkR As Long, s As Shape, r As Long, sinirX As Double, ust As Long, sinir As Long
    Set vr = ActiveWindow.VisibleRange
    sonC = vr.Column + vr.Columns.Count - 1
    sonR = vr.Row + vr.Rows.Count - 1
    ilkR = sonR
    If InStr(1, ws.Name, "_STOK", vbTextCompare) = 0 And ws.Name <> "STOKLAR" And InStr(1, ws.Name, "MALZEME_ARA", vbTextCompare) = 0 Then
        If ws.Name = "ADMIN PANEL" Then sinir = 400 Else sinir = 150
        On Error Resume Next
        ust = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
        If ust > sinir Then ust = sinir
        For r = ust To sonR + 1 Step -1
            If application.WorksheetFunction.CountA(ws.Range(ws.Cells(r, 1), ws.Cells(r, sonC))) > 0 Then sonR = r: Exit For
        Next r
        sinirX = ws.Cells(1, sonC + 1).Left
        For Each s In ws.Shapes
            r = 0
            If s.Visible And s.Left < sinirX Then r = s.BottomRightCell.Row
            If r > sonR Then sonR = r
        Next s
        Err.Clear
        On Error GoTo 0
        If sonR > sinir Then sonR = application.Max(sinir, ilkR)
    End If
    Set T_EkranAlan = ws.Range(ws.Cells(1, 1), ws.Cells(sonR, sonC))
End Function

' AKSORA v2 BOLGE KOMUTA MERKEZI: bes ilin Malzeme, Ilce ve Yuklenici stoklari tek bir web sayfasinda (genel bakis,
' aninda akilli arama, bolge dengesi, depolar). Excel SAP okuma motoru olarak kalir; sayfa her acilista en guncel
' veriyle masaustune yazilir ve tarayicida acilir (internet gerekmez, tek dosya).
Public Sub AKSORA_V2_AC()
    Dim ac As Boolean, il As Variant, canli As String, parca() As String, n As Long, tarihler As String, yol As String
    Dim i As Long, tur As Variant, ws As Worksheet, tarih As Variant, birim As String, html As String, hM As String
    Dim turler As Variant, canliAd As Variant, birimler As Variant
    On Error GoTo Hata
    ac = AKSORA_BEKLE_AC("AKSORA  |  v2", "Bölge verisi hazýrlanýyor...", 0.05)
    On Error Resume Next
    canli = AKSORA_CANLI_STOK_IL()
    On Error GoTo Hata
    turler = Array("TUM", "ILC", "YUK")
    canliAd = Array("STOKLAR", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
    birimler = Array("MALZEME", "ILCE", "YUKLENICI")
    ReDim parca(1 To 1000)
    For Each il In Array("ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
        If ac Then AKSORA_BEKLE_ILERLE 0.1 + 0.15 * i, "Bölge verisi  •  " & CStr(il)
        i = i + 1
        For tur = 0 To 2
            Set ws = Nothing: tarih = Empty
            On Error Resume Next
            If CStr(il) = canli Then
                Set ws = ThisWorkbook.Worksheets(CStr(canliAd(tur)))
                If Not ws Is Nothing Then tarih = ws.Range("C2").Value
            Else
                Set ws = ThisWorkbook.Worksheets("_AKLIVE_" & CStr(il) & "_" & CStr(turler(tur)))
                If Not ws Is Nothing Then tarih = ws.Range("A3").Value
            End If
            On Error GoTo Hata
            If Not ws Is Nothing Then
                birim = CStr(birimler(tur))
                V2_SATIRLAR ws, CStr(il), birim, parca, n
                If IsDate(tarih) Then
                    tarihler = tarihler & IIf(Len(tarihler) > 0, ",", "") & """" & CStr(il) & "|" & birim & """:""" & Format$(CDate(tarih), "dd.mm.yyyy hh:nn") & """"
                End If
            End If
        Next tur
    Next il
    If n = 0 Then Err.Raise vbObjectError + 9901, , "Stok verisi bulunamadý (önce stoklarý güncelleyin)."
    ReDim Preserve parca(1 To n)
    If ac Then AKSORA_BEKLE_ILERLE 0.9, "Sayfa yazýlýyor (" & AKSORA_SAYI_METNI(n, "#,##0") & " stok satýrý)..."
    html = Replace(V2_SABLON(), "/*VERI*/", "window.AKSORA_SNAPSHOT={""rows"":[" & Join(parca, ",") & "],""tarih"":{" & tarihler & "}};")
    yol = AKSORA_MASAUSTU() & application.PathSeparator & "AKSORA v2.html"
    V2_YAZ yol, html
    If ac Then AKSORA_BEKLE_KAPAT: ac = False
    ' Varsayilan tarayicida acilir (FollowHyperlink Office guvenlik uyarisi cikariyordu).
    CreateObject("Shell.Application").ShellExecute yol, "", "", "open", 1
    Exit Sub
Hata:
    hM = Err.Description
    On Error Resume Next
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    MsgBox "AKSORA v2 açýlamadý: " & hM, vbExclamation, "AKSORA v2"
End Sub

' Stok sayfasi / arsivi (A7:L): [il, birim, UY, depo, depo adi, mal grubu, degerleme, malzeme, tanim, miktar, birim, deger]
Private Sub V2_SATIRLAR(ByVal ws As Worksheet, ByVal il As String, ByVal birim As String, ByRef parca() As String, ByRef n As Long)
    Dim son As Long, v As Variant, r As Long
    son = AKSORA_SON_SATIR(ws, "G")
    If son < 7 Then Exit Sub
    v = ws.Range("A7:L" & son).Value2
    If Not IsArray(v) Then Exit Sub
    For r = 1 To UBound(v, 1)
        If Not IsError(v(r, 7)) Then
            If Len(Trim$(CStr(v(r, 7)))) > 0 Then
                n = n + 1
                If n > UBound(parca) Then ReDim Preserve parca(1 To UBound(parca) * 2)
                parca(n) = "[""" & il & """,""" & birim & """,""" & V2_J(v(r, 1)) & """,""" & V2_J(v(r, 2)) & """,""" & V2_J(v(r, 3)) & _
                    """,""" & V2_J(v(r, 4)) & """,""" & V2_J(v(r, 6)) & """,""" & V2_J(v(r, 7)) & """,""" & V2_J(v(r, 8)) & """," & _
                    V2_N(v(r, 9)) & ",""" & V2_J(v(r, 10)) & """," & V2_N(v(r, 11)) & "]"
            End If
        End If
    Next r
End Sub

Private Function V2_J(ByVal x As Variant) As String
    Dim s As String
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If VarType(x) = vbDouble Then
        If x = Int(x) And Abs(x) < 1E+15 Then s = Format$(x, "0") Else s = CStr(x)
    Else
        s = CStr(x)
    End If
    s = Replace(s, "\", "\\"): s = Replace(s, """", "\""")
    s = Replace(Replace(Replace(s, vbCr, " "), vbLf, " "), vbTab, " ")
    V2_J = Replace(s, "</", "<\/")
End Function

Private Function V2_N(ByVal x As Variant) As String
    V2_N = "0"
    If IsError(x) Or IsEmpty(x) Then Exit Function
    If IsNumeric(x) Then V2_N = Trim$(Str$(CDbl(x)))
End Function

Private Sub V2_YAZ(ByVal yol As String, ByVal metin As String)
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 2: st.Charset = "utf-8": st.Open
    st.WriteText metin
    st.SaveToFile yol, 2
    st.Close
End Sub

' AKSORA v2 sayfa sablonu (HTML + CSS + JS); /*VERI*/ yerine bolge verisi yazilir. Kaynak: v2/AKSORA_v2_sablon.html
Private Function V2_SABLON() As String
    V2_SABLON = V2_S1() & V2_S2()
End Function

Private Function V2_S1() As String
    Dim t As String
    t = t & "<!doctype html>" & vbLf
    t = t & "<html lang=""tr"">" & vbLf
    t = t & "<head>" & vbLf
    t = t & "<meta charset=""utf-8"">" & vbLf
    t = t & "<meta name=""viewport"" content=""width=device-width, initial-scale=1"">" & vbLf
    t = t & "<title>AKSORA v2 · Bölge Komuta Merkezi</title>" & vbLf
    t = t & "<style>" & vbLf
    t = t & ":root{" & vbLf
    t = t & "  --lacivert:#0c2d4f; --lacivert2:#16416b; --teal:#007d6e; --teal2:#0fa391; --mavi:#1e63a8; --acik:#eef5fb;" & vbLf
    t = t & "  --zemin:#f3f7fb; --kart:#ffffff; --cizgi:#dbe5ef; --yazi:#1c2b3a; --soluk:#64748b;" & vbLf
    t = t & "  --kirmizi:#dc2626; --turuncu:#ea8a0c; --yesil:#16a34a; --mor:#7c3aed; --dmavi:#e1effa;" & vbLf
    t = t & "  --golge:0 1px 2px rgba(12,45,79,.06),0 6px 18px rgba(12,45,79,.07);" & vbLf
    t = t & "}" & vbLf
    t = t & "@media (prefers-color-scheme: dark){" & vbLf
    t = t & "  :root:not([data-tema=""acik""]){" & vbLf
    t = t & "    --zemin:#0b1622; --kart:#122233; --cizgi:#22384d; --yazi:#e6eef6; --soluk:#94a8bb; --acik:#16293c; --dmavi:#163452;" & vbLf
    t = t & "    --golge:0 1px 2px rgba(0,0,0,.3),0 8px 22px rgba(0,0,0,.35);" & vbLf
    t = t & "  }" & vbLf
    t = t & "}" & vbLf
    t = t & ":root[data-tema=""koyu""]{" & vbLf
    t = t & "  --zemin:#0b1622; --kart:#122233; --cizgi:#22384d; --yazi:#e6eef6; --soluk:#94a8bb; --acik:#16293c; --dmavi:#163452;" & vbLf
    t = t & "  --golge:0 1px 2px rgba(0,0,0,.3),0 8px 22px rgba(0,0,0,.35);" & vbLf
    t = t & "}" & vbLf
    t = t & "*{box-sizing:border-box}" & vbLf
    t = t & "html,body{margin:0;background:var(--zemin);color:var(--yazi);font-family:""Aptos"",""Segoe UI Variable"",""Segoe UI"",system-ui,sans-serif;font-size:14px}" & vbLf
    t = t & "button{font-family:inherit}" & vbLf
    t = t & ".ust{background:linear-gradient(120deg,#0a2742 0%,#0c2d4f 45%,#0d4d63 100%);color:#fff;padding:18px 28px 0;position:relative;overflow:hidden}" & vbLf
    t = t & ".ust:after{content:"""";position:absolute;right:-120px;top:-160px;width:520px;height:520px;border-radius:50%;" & vbLf
    t = t & "  background:radial-gradient(circle,rgba(15,163,145,.35),rgba(15,163,145,0) 65%)}" & vbLf
    t = t & ".ust-ic{display:flex;align-items:center;gap:18px;position:relative;z-index:1;flex-wrap:wrap}" & vbLf
    t = t & ".logo{font-weight:800;font-size:28px;letter-spacing:3px}" & vbLf
    t = t & ".logo b{color:#3fd6c2}" & vbLf
    t = t & ".rozet{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);padding:3px 10px;border-radius:999px;font-size:12px;font-weight:700;letter-spacing:1px}" & vbLf
    t = t & ".alt-baslik{opacity:.85;font-size:13px}" & vbLf
    t = t & ".ust-sag{margin-left:auto;display:flex;gap:10px;align-items:center;font-size:12px;opacity:.95}" & vbLf
    t = t & ".tema{background:rgba(255,255,255,.12);border:1px solid rgba(255,255,255,.25);color:#fff;border-radius:8px;padding:6px 10px;cursor:pointer}" & vbLf
    t = t & ".sekmeler{display:flex;gap:4px;margin-top:16px;position:relative;z-index:1;overflow-x:auto}" & vbLf
    t = t & ".sekme{background:transparent;border:0;color:#cfe3f3;padding:11px 18px;font-size:14px;font-weight:600;cursor:pointer;border-radius:10px 10px 0 0;white-space:nowrap}" & vbLf
    t = t & ".sekme.aktif{background:var(--zemin);color:var(--yazi)}" & vbLf
    t = t & ".filtre{display:flex;gap:16px;flex-wrap:wrap;align-items:center;padding:14px 28px;border-bottom:1px solid var(--cizgi);background:var(--kart)}" & vbLf
    t = t & ".grup{display:flex;gap:6px;align-items:center;flex-wrap:wrap}" & vbLf
    t = t & ".grup span.et{font-size:11px;font-weight:700;color:var(--soluk);letter-spacing:.8px;margin-right:4px}" & vbLf
    t = t & ".cip{border:1px solid var(--cizgi);background:var(--kart);color:var(--yazi);border-radius:999px;padding:6px 13px;font-size:12.5px;font-weight:600;cursor:pointer}" & vbLf
    t = t & ".cip.aktif{background:var(--lacivert);border-color:var(--lacivert);color:#fff}" & vbLf
    t = t & ":root[data-tema=""koyu""] .cip.aktif{background:var(--teal);border-color:var(--teal)}" & vbLf
    t = t & ".icerik{padding:22px 28px 40px;max-width:1500px;margin:0 auto}" & vbLf
    t = t & ".gorunum{display:none}.gorunum.aktif{display:block}" & vbLf
    t = t & ".kpiler{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}" & vbLf
    t = t & ".kpi{background:var(--kart);border:1px solid var(--cizgi);border-radius:14px;padding:16px 18px;box-shadow:var(--golge);position:relative;overflow:hidden}" & vbLf
    t = t & ".kpi:before{content:"""";position:absolute;left:0;top:0;bottom:0;width:4px;background:var(--renk,var(--mavi))}" & vbLf
    t = t & ".kpi .ad{font-size:11px;font-weight:700;color:var(--soluk);letter-spacing:.8px}" & vbLf
    t = t & ".kpi .deger{font-size:24px;font-weight:800;margin-top:6px;font-variant-numeric:tabular-nums}" & vbLf
    t = t & ".kpi .not{font-size:12px;color:var(--soluk);margin-top:4px}" & vbLf
    t = t & ".izgara{display:grid;grid-template-columns:1.6fr 1fr;gap:14px;margin-top:14px}" & vbLf
    t = t & ".izgara2{display:grid;grid-template-columns:1fr 1fr;gap:14px;margin-top:14px}" & vbLf
    t = t & ".kart{background:var(--kart);border:1px solid var(--cizgi);border-radius:14px;padding:18px;box-shadow:var(--golge);min-width:0}" & vbLf
    t = t & ".kart h3{margin:0 0 4px;font-size:15px}" & vbLf
    t = t & ".kart .acik{font-size:12px;color:var(--soluk);margin-bottom:12px}" & vbLf
    t = t & ".cubuk-satir{display:grid;grid-template-columns:110px 1fr 150px;gap:10px;align-items:center;margin:9px 0}" & vbLf
    t = t & ".cubuk-satir .il{font-weight:700;font-size:13px}" & vbLf
    t = t & ".cubuk{display:flex;height:22px;border-radius:6px;overflow:hidden;background:var(--acik)}" & vbLf
    t = t & ".cubuk div{height:100%}" & vbLf
    t = t & ".cubuk-satir .tl{text-align:right;font-variant-numeric:tabular-nums;font-weight:700;font-size:13px}" & vbLf
    t = t & ".lejant{display:flex;gap:14px;flex-wrap:wrap;font-size:12px;color:var(--soluk);margin-top:10px}" & vbLf
    t = t & ".lejant i{display:inline-block;width:10px;height:10px;border-radius:3px;margin-right:6px;vertical-align:-1px}" & vbLf
    t = t & ".halka-kutu{display:flex;align-items:center;gap:18px;flex-wrap:wrap}" & vbLf
    t = t & ".tazelik{width:100%;border-collapse:separate;border-spacing:0 6px}" & vbLf
    t = t & ".tazelik th{font-size:11px;color:var(--soluk);text-align:left;font-weight:700;letter-spacing:.6px;padding:0 8px}" & vbLf
    t = t & ".tazelik td{background:var(--acik);padding:9px 8px;font-size:12.5px}" & vbLf
    t = t & ".tazelik td:first-child{border-radius:8px 0 0 8px;font-weight:700}" & vbLf
    t = t & ".tazelik td:last-child{border-radius:0 8px 8px 0}" & vbLf
    t = t & ".led{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:6px;vertical-align:0}" & vbLf
    t = t & ".arama-kutu{display:flex;gap:10px;align-items:center;background:var(--kart);border:2px solid var(--cizgi);border-radius:14px;padding:6px 6px 6px 16px;box-shadow:var(--golge)}" & vbLf
    t = t & ".arama-kutu:focus-within{border-color:var(--teal2)}" & vbLf
    t = t & ".arama-kutu input{flex:1;border:0;outline:0;background:transparent;color:var(--yazi);font-size:18px;padding:10px 0;font-family:inherit;min-width:0}" & vbLf
    t = t & ".ipucu{display:flex;gap:8px;flex-wrap:wrap;margin:10px 0 4px}" & vbLf
    t = t & ".ipucu button{border:1px dashed var(--cizgi);background:transparent;color:var(--soluk);border-radius:999px;padding:5px 11px;font-size:12px;cursor:pointer}" & vbLf
    t = t & ".ozet-cipler{display:flex;gap:8px;flex-wrap:wrap;margin:14px 0}" & vbLf
    t = t & ".ozet-cip{background:var(--kart);border:1px solid var(--cizgi);border-radius:10px;padding:8px 12px;font-size:12px}" & vbLf
    t = t & ".ozet-cip b{display:block;font-size:15px;font-variant-numeric:tabular-nums}" & vbLf
    t = t & ".tablo-kap{overflow:auto;border:1px solid var(--cizgi);border-radius:12px;background:var(--kart);max-height:620px}" & vbLf
    t = t & "table.t{width:100%;border-collapse:collapse;font-size:12.5px}" & vbLf
    t = t & "table.t th{position:sticky;top:0;background:var(--lacivert);color:#fff;text-align:left;padding:10px;font-size:11.5px;letter-spacing:.4px;white-space:nowrap;z-index:1}" & vbLf
    t = t & "table.t td{padding:8px 10px;border-bottom:1px solid var(--cizgi);white-space:nowrap}" & vbLf
    t = t & "table.t td.sag,table.t th.sag{text-align:right;font-variant-numeric:tabular-nums}" & vbLf
    t = t & "table.t tr.d td{background:var(--dmavi)}" & vbLf
    t = t & "table.t td.tanim{white-space:normal;min-width:260px}" & vbLf
    t = t & ".isi{display:inline-block;min-width:52px;text-align:right;padding:3px 7px;border-radius:6px;font-variant-numeric:tabular-nums}" & vbLf
    t = t & ".mini{display:inline-block;height:8px;border-radius:4px;background:linear-gradient(90deg,var(--teal),var(--teal2));vertical-align:middle}" & vbLf
    t = t & ".etiket{display:inline-block;padding:2px 8px;border-radius:999px;font-size:11px;font-weight:700}" & vbLf
    t = t & ".bos{padding:30px;text-align:center;color:var(--soluk)}" & vbLf
    t = t & ".alt{color:var(--soluk);font-size:12px;text-align:center;padding:10px 0 30px}" & vbLf
    t = t & "@media (max-width:1000px){.kpiler{grid-template-columns:repeat(2,1fr)}.izgara,.izgara2{grid-template-columns:1fr}}" & vbLf
    t = t & "@media (max-width:560px){.icerik,.filtre,.ust{padding-left:16px;padding-right:16px}.kpiler{grid-template-columns:1fr}.cubuk-satir{grid-template-columns:80px 1fr 110px}}" & vbLf
    t = t & "</style>" & vbLf
    t = t & "</head>" & vbLf
    t = t & "<body>" & vbLf
    t = t & "<div class=""ust"">" & vbLf
    t = t & "  <div class=""ust-ic"">" & vbLf
    t = t & "    <div class=""logo"">AKS<b>O</b>RA</div>" & vbLf
    t = t & "    <span class=""rozet"">v2</span>" & vbLf
    t = t & "    <div>" & vbLf
    t = t & "      <div style=""font-weight:700;font-size:16px"">Bölge Komuta Merkezi</div>" & vbLf
    t = t & "      <div class=""alt-baslik"">Çoruh Elektrik Daðýtým A.Þ. · Artvin · Rize · Trabzon · Gümüþhane · Giresun</div>" & vbLf
    t = t & "    </div>" & vbLf
    t = t & "    <div class=""ust-sag""><span id=""veriTarih""></span><button class=""tema"" id=""temaBtn"" title=""Açýk / koyu tema"">&#9680; Tema</button></div>" & vbLf
    t = t & "  </div>" & vbLf
    t = t & "  <div class=""sekmeler"" id=""sekmeler"">" & vbLf
    t = t & "    <button class=""sekme aktif"" data-g=""genel"">Genel Bakýþ</button>" & vbLf
    t = t & "    <button class=""sekme"" data-g=""arama"">Akýllý Arama</button>" & vbLf
    t = t & "    <button class=""sekme"" data-g=""denge"">Bölge Dengesi</button>" & vbLf
    t = t & "    <button class=""sekme"" data-g=""depolar"">Depolar</button>" & vbLf
    t = t & "  </div>" & vbLf
    t = t & "</div>" & vbLf
    t = t & "<div class=""filtre"">" & vbLf
    t = t & "  <div class=""grup"" id=""ilFiltre""><span class=""et"">ÝL</span></div>" & vbLf
    t = t & "  <div class=""grup"" id=""birimFiltre""><span class=""et"">BÝRÝM</span></div>" & vbLf
    t = t & "</div>" & vbLf
    t = t & "<div class=""icerik"">" & vbLf
    t = t & "  <section class=""gorunum aktif"" id=""g-genel"">" & vbLf
    t = t & "    <div class=""kpiler"" id=""kpiler""></div>" & vbLf
    t = t & "    <div class=""izgara"">" & vbLf
    t = t & "      <div class=""kart""><h3>Ýllere göre stok deðeri</h3><div class=""acik"">Birim kýrýlýmýyla, en yüksekten en düþüðe.</div><div id=""ilCubuk""></div><div class=""lejant"" id=""ilLejant""></div></div>" & vbLf
    t = t & "      <div class=""kart""><h3>Birimlere göre daðýlým</h3><div class=""acik"">Stok deðerinin birimlere göre payý.</div><div class=""halka-kutu"" id=""halka""></div></div>" & vbLf
    t = t & "    </div>" & vbLf
    t = t & "    <div class=""izgara2"">" & vbLf
    t = t & "      <div class=""kart""><h3>En deðerli 10 malzeme</h3><div class=""acik"">Seçili il ve birimlerde toplam stok deðeri.</div><div id=""enDeger""></div></div>" & vbLf
    t = t & "      <div class=""kart""><h3>Veri güncelliði</h3><div class=""acik"">Her il ve birimin son SAP güncellemesi.</div><div id=""tazelik""></div></div>" & vbLf
    t = t & "    </div>" & vbLf
    t = t & "  </section>" & vbLf
    t = t & "  <section class=""gorunum"" id=""g-arama"">" & vbLf
    t = t & "    <div class=""arama-kutu""><span style=""font-size:20px"">&#8981;</span><input id=""ara"" placeholder=""Malzeme ara: 50 kva trafo, aðaç direk 9, 1000006855 ..."" autocomplete=""off""></div>" & vbLf
    t = t & "    <div class=""ipucu"" id=""ipucu""></div>" & vbLf
    t = t & "    <div class=""ozet-cipler"" id=""araOzet""></div>" & vbLf
    t = t & "    <div id=""araSonuc""><div class=""bos"">Yazmaya baþlayýn: sonuçlar beþ ilde ve bütün birimlerde anýnda gelir. Sözcük sýrasý önemsiz, 50 = 0050.</div></div>" & vbLf
    t = t & "  </section>" & vbLf
    t = t & "  <section class=""gorunum"" id=""g-denge"">" & vbLf
    t = t & "    <div class=""kart"" style=""margin-bottom:14px""><h3>Bölge dengesi: ayný malzeme hangi ilde birikmiþ?</h3>" & vbLf
    t = t & "      <div class=""acik"">Yeni stokta birden fazla ilde bulunan malzemeler, toplam deðere göre. Koyu hücre: o il bölge stoðunun büyük kýsmýný tutuyor; satýn almadan önce oradan transfer düþünülebilir.</div>" & vbLf
    t = t & "      <div class=""arama-kutu"" style=""box-shadow:none;border-width:1px""><span>&#8981;</span><input id=""dengeAra"" placeholder=""Malzemeyi daralt (ör. kablo, armatür)"" style=""font-size:15px""></div>" & vbLf
    t = t & "    </div>" & vbLf
    t = t & "    <div id=""dengeTablo""></div>" & vbLf
    t = t & "  </section>" & vbLf
    t = t & "  <section class=""gorunum"" id=""g-depolar"">" & vbLf
    t = t & "    <div id=""depoTablo""></div>" & vbLf
    t = t & "  </section>" & vbLf
    t = t & "  <div class=""alt"">AKSORA v2 · Veri kaynaðý: AKSORA Excel (SAP MB52 / IQ09) · Excel'de <b>AKSORA v2</b> düðmesi bu sayfayý en güncel veriyle açar.</div>" & vbLf
    t = t & "</div>" & vbLf
    t = t & "<script>/*VERI*/</script>" & vbLf
    t = t & "<script>" & vbLf
    t = t & "(function(){" & vbLf
    t = t & "const D = window.AKSORA_DATA || window.AKSORA_SNAPSHOT || {rows:[],tarih:{}};" & vbLf
    t = t & "const IL = [""ARTVIN"",""RIZE"",""TRABZON"",""GUMUSHANE"",""GIRESUN""];" & vbLf
    t = t & "const ILAD = {ARTVIN:""Artvin"",RIZE:""Rize"",TRABZON:""Trabzon"",GUMUSHANE:""Gümüþhane"",GIRESUN:""Giresun""};" & vbLf
    t = t & "const BIRIM = [""MALZEME"",""ILCE"",""YUKLENICI""];" & vbLf
    t = t & "const BAD = {MALZEME:""Malzeme Stok"",ILCE:""Ýlçe"",YUKLENICI:""Yüklenici""};" & vbLf
    t = t & "const BRENK = {MALZEME:""#1e63a8"",ILCE:""#0fa391"",YUKLENICI:""#ea8a0c""};" & vbLf
    t = t & "const tl = new Intl.NumberFormat(""tr-TR"",{minimumFractionDigits:2,maximumFractionDigits:2});" & vbLf
    t = t & "const sayi = new Intl.NumberFormat(""tr-TR"",{maximumFractionDigits:3});" & vbLf
    t = t & "const TL = x => tl.format(x||0)+"" TL"";" & vbLf
    t = t & "const KISA = x => { const a=Math.abs(x); if(a>=1e9) return tl.format(x/1e9)+"" mr TL""; if(a>=1e6) return tl.format(x/1e6)+"" mn TL""; if(a>=1e3) return tl.format(x/1e3)+"" bin TL""; return TL(x); };" & vbLf
    t = t & "const MK = x => sayi.format(Math.round((x||0)*1000)/1000);" & vbLf
    t = t & "const esc = s => String(s==null?"""":s).replace(/[&<>""]/g,c=>({""&"":""&amp;"",""<"":""&lt;"","">"":""&gt;"",'""':""&quot;""}[c]));" & vbLf
    t = t & "// satir: [il, birim, uy, depo, depoAdi, mg, degerleme, malzeme, tanim, miktar, birim, deger]" & vbLf
    t = t & "const R = D.rows.map(r=>({il:r[0],b:r[1],uy:r[2],dp:r[3],da:r[4],mg:r[5],dg:r[6],mal:r[7],tn:r[8],mk:+r[9]||0,br:r[10],tl:+r[11]||0}));" & vbLf
    t = t & "let secIl = new Set(IL), secB = new Set(BIRIM);" & vbLf
    t = t & "// Tema" & vbLf
    t = t & "const kok=document.documentElement;" & vbLf
    t = t & "try{const t=localStorage.getItem(""aksora_v2_tema""); if(t) kok.dataset.tema=t;}catch(e){}" & vbLf
    t = t & "document.getElementById(""temaBtn"").onclick=()=>{const k=kok.dataset.tema===""koyu""||(!kok.dataset.tema&&matchMedia(""(prefers-color-scheme: dark)"").matches);kok.dataset.tema=k?""acik"":""koyu"";try{localStorage.setItem(""aksora_v2_tema"",kok.dataset.tema)}catch(e){}};" & vbLf
    t = t & "// Veri tarihi" & vbLf
    t = t & "const tarihler = Object.values(D.tarih||{}); document.getElementById(""veriTarih"").textContent = tarihler.length? ""Veri: ""+tarihler.sort().slice(-1)[0] : """";" & vbLf
    t = t & "// Filtreler" & vbLf
    t = t & "function cipler(kap, liste, adlar, set, tum){" & vbLf
    t = t & "  const el=document.getElementById(kap);" & vbLf
    t = t & "  const ciz=()=>{ el.querySelectorAll(""button"").forEach(b=>b.remove());" & vbLf
    t = t & "    const hepsi=document.createElement(""button""); hepsi.className=""cip""+(set.size===liste.length?"" aktif"":""""); hepsi.textContent=tum;" & vbLf
    t = t & "    hepsi.onclick=()=>{liste.forEach(x=>set.add(x)); ciz(); yenile();}; el.appendChild(hepsi);" & vbLf
    t = t & "    liste.forEach(x=>{const b=document.createElement(""button""); b.className=""cip""+(set.size!==liste.length&&set.has(x)?"" aktif"":""""); b.textContent=adlar[x];" & vbLf
    t = t & "      b.onclick=()=>{ if(set.size===liste.length){set.clear();set.add(x);} else if(set.has(x)){ if(set.size>1) set.delete(x);} else set.add(x); ciz(); yenile();}; el.appendChild(b);});" & vbLf
    t = t & "  }; ciz();" & vbLf
    t = t & "}" & vbLf
    t = t & "cipler(""ilFiltre"",IL,ILAD,secIl,""Tüm iller""); cipler(""birimFiltre"",BIRIM,BAD,secB,""Tüm birimler"");" & vbLf
    t = t & "const F = () => R.filter(r=>secIl.has(r.il)&&secB.has(r.b));" & vbLf
    t = t & "// Sekmeler" & vbLf
    t = t & "document.querySelectorAll("".sekme"").forEach(b=>b.onclick=()=>{document.querySelectorAll("".sekme"").forEach(x=>x.classList.toggle(""aktif"",x===b));" & vbLf
    t = t & "  document.querySelectorAll("".gorunum"").forEach(g=>g.classList.toggle(""aktif"",g.id===""g-""+b.dataset.g)); if(b.dataset.g===""arama"") document.getElementById(""ara"").focus();});" & vbLf
    t = t & "function degTur(dg){const t=(dg||"""").toUpperCase().replace(/[^A-Z]/g,""""); if(t.includes(""HURDA"")) return ""HURDA""; if(t.includes(""CIKMA"")||t.includes(""BDSZ"")) return ""ÇIKMA""; return ""YENÝ"";}" & vbLf
    t = t & "// GENEL" & vbLf
    t = t & "function genel(){" & vbLf
    t = t & "  const f=F(); let top=0; const mal=new Set(), dep=new Set();" & vbLf
    t = t & "  f.forEach(r=>{top+=r.tl; mal.add(r.mal); dep.add(r.il+r.uy+r.dp);});" & vbLf
    t = t & "  document.getElementById(""kpiler"").innerHTML=[" & vbLf
    t = t & "    [""TOPLAM STOK DEÐERÝ"",KISA(top),TL(top),""#1e63a8""],[""STOK SATIRI"",MK(f.length),secIl.size+"" il · ""+secB.size+"" birim"",""#0fa391""]," & vbLf
    t = t & "    [""MALZEME ÇEÞÝDÝ"",MK(mal.size),""farklý malzeme kodu"",""#7c3aed""],[""DEPO"",MK(dep.size),""stok bulunan depo"",""#ea8a0c""]]" & vbLf
    t = t & "    .map(k=>`<div class=""kpi"" style=""--renk:${k[3]}""><div class=""ad"">${k[0]}</div><div class=""deger"">${k[1]}</div><div class=""not"">${k[2]}</div></div>`).join("""");" & vbLf
    t = t & "  const ilT={}; IL.forEach(i=>{ilT[i]={t:0}; BIRIM.forEach(b=>ilT[i][b]=0);});" & vbLf
    t = t & "  f.forEach(r=>{ilT[r.il][r.b]+=r.tl; ilT[r.il].t+=r.tl;});" & vbLf
    t = t & "  const sira=IL.filter(i=>secIl.has(i)).sort((a,b)=>ilT[b].t-ilT[a].t); const mx=Math.max(1,...sira.map(i=>ilT[i].t));" & vbLf
    t = t & "  document.getElementById(""ilCubuk"").innerHTML=sira.map(i=>`<div class=""cubuk-satir""><div class=""il"">${ILAD[i]}</div><div class=""cubuk"" style=""width:${Math.max(2,ilT[i].t/mx*100)}%"">${BIRIM.filter(b=>secB.has(b)).map(b=>`<div title=""${BAD[b]}: ${TL(ilT[i][b])}"" style=""width:${ilT[i].t?ilT[i][b]/ilT[i].t*100:0}%;background:${BRENK[b]}""></div>`).join("""")}</div><div class=""tl"">${KISA(ilT[i].t)}</div></div>`).join("""");" & vbLf
    t = t & "  document.getElementById(""ilLejant"").innerHTML=BIRIM.filter(b=>secB.has(b)).map(b=>`<span><i style=""background:${BRENK[b]}""></i>${BAD[b]}</span>`).join("""");" & vbLf
    t = t & "  // halka: birim payi" & vbLf
    t = t & "  const dt={}; BIRIM.filter(b=>secB.has(b)).forEach(b=>dt[b]=0); f.forEach(r=>{if(r.b in dt) dt[r.b]+=r.tl;});" & vbLf
    t = t & "  const cik=f.filter(r=>degTur(r.dg)!==""YENÝ"").length;" & vbLf
    t = t & "  const renk=BRENK; const tt=Object.values(dt).reduce((a,b)=>a+b,0)||1;" & vbLf
    t = t & "  let ac=-Math.PI/2, yol=""""; const cx=90,cy=90,r1=80,r2=54;" & vbLf
    t = t & "  Object.keys(dt).forEach(k=>{const a=dt[k]/tt*Math.PI*2; if(a<=0) return; const b=ac+Math.min(a,Math.PI*2-1e-4);" & vbLf
    t = t & "    const p=(rr,an)=>[cx+rr*Math.cos(an),cy+rr*Math.sin(an)]; const [x1,y1]=p(r1,ac),[x2,y2]=p(r1,b),[x3,y3]=p(r2,b),[x4,y4]=p(r2,ac); const buyuk=(b-ac)>Math.PI?1:0;" & vbLf
    t = t & "    yol+=`<path d=""M${x1},${y1} A${r1},${r1} 0 ${buyuk} 1 ${x2},${y2} L${x3},${y3} A${r2},${r2} 0 ${buyuk} 0 ${x4},${y4} Z"" fill=""${renk[k]}"" stroke=""var(--kart)"" stroke-width=""2""><title>${BAD[k]}: ${TL(dt[k])}</title></path>`; ac=b;});" & vbLf
    t = t & "  const yuzde=x=>""%""+new Intl.NumberFormat(""tr-TR"",{maximumFractionDigits:1}).format(x);" & vbLf
    t = t & "  document.getElementById(""halka"").innerHTML=`<svg width=""180"" height=""180"" viewBox=""0 0 180 180"">${yol}<text x=""90"" y=""86"" text-anchor=""middle"" font-size=""11"" fill=""var(--soluk)"" font-weight=""700"">TOPLAM</text><text x=""90"" y=""104"" text-anchor=""middle"" font-size=""13"" fill=""var(--yazi)"" font-weight=""800"">${KISA(tt===1?0:tt)}</text></svg>" & vbLf
    t = t & "   <div style=""flex:1;min-width:170px"">${Object.keys(dt).map(k=>`<div style=""display:flex;justify-content:space-between;gap:10px;padding:8px 0;border-bottom:1px solid var(--cizgi)""><span><i style=""display:inline-block;width:10px;height:10px;border-radius:3px;background:${renk[k]};margin-right:8px""></i>${BAD[k]}</span><b style=""font-variant-numeric:tabular-nums"">${KISA(dt[k])} <span style=""color:var(--soluk);font-weight:600"">· ${yuzde(dt[k]/tt*100)}</span></b></div>`).join("""")}" & vbLf
    t = t & "   <div style=""color:var(--soluk);font-size:12px;margin-top:10px"">Çýkma / hurda (bedelsiz) satýr: <b>${MK(cik)}</b></div></div>`;" & vbLf
    t = t & "  // en degerli" & vbLf
    t = t & "  const m={}; f.forEach(r=>{const k=r.mal; if(!m[k]) m[k]={tn:r.tn,tl:0,mk:0,br:r.br}; m[k].tl+=r.tl; m[k].mk+=r.mk;});" & vbLf
    t = t & "  const en=Object.entries(m).sort((a,b)=>b[1].tl-a[1].tl).slice(0,10); const emx=en.length?en[0][1].tl:1;" & vbLf
    t = t & "  document.getElementById(""enDeger"").innerHTML=en.length?`<table class=""t""><tbody>${en.map(([k,v],i)=>`<tr><td style=""color:var(--soluk)"">${i+1}</td><td class=""tanim""><b>${esc(v.tn)}</b><div style=""color:var(--soluk);font-size:11.5px"">${k} · ${MK(v.mk)} ${esc(v.br)}</div></td><td class=""sag""><span class=""mini"" style=""width:${Math.max(4,v.tl/emx*90)}px""></span> ${KISA(v.tl)}</td></tr>`).join("""")}</tbody></table>`:`<div class=""bos"">Veri yok.</div>`;" & vbLf
    t = t & "  // tazelik" & vbLf
    t = t & "  const simdi=new Date();" & vbLf
    t = t & "  const yas=s=>{const m=/(\d\d)\.(\d\d)\.(\d{4})\s+(\d\d):(\d\d)/.exec(s||""""); if(!m) return null; const d=new Date(+m[3],+m[2]-1,+m[1],+m[4],+m[5]); return (simdi-d)/864e5;};" & vbLf
    t = t & "  const led=g=>g==null?""#94a3b8"":g<1?""#16a34a"":g<2?""#ea8a0c"":""#dc2626"";" & vbLf
    t = t & "  document.getElementById(""tazelik"").innerHTML=`<table class=""tazelik""><thead><tr><th>ÝL</th>${BIRIM.map(b=>`<th>${BAD[b].toUpperCase()}</th>`).join("""")}</tr></thead><tbody>${IL.map(i=>`<tr><td>${ILAD[i]}</td>${BIRIM.map(b=>{const s=(D.tarih||{})[i+""|""+b]; const g=yas(s); return `<td><span class=""led"" style=""background:${led(g)}""></span>${esc(s||""-"")}</td>`;}).join("""")}</tr>`).join("""")}</tbody></table>`;" & vbLf
    t = t & "}" & vbLf
    t = t & "// AKILLI ARAMA (AKSORA Excel ile ayni kural)" & vbLf
    t = t & "function sozcuk(t){t=String(t||"""").toLocaleUpperCase(""tr-TR"").replace(/Ý/g,""I"").replace(/Þ/g,""S"").replace(/Ð/g,""G"").replace(/Ü/g,""U"").replace(/Ö/g,""O"").replace(/Ç/g,""C"");" & vbLf
    t = t & "  const out=[]; const re=/[A-Z]+|[0-9]+/g; let m; while((m=re.exec(t))){let x=m[0]; if(/^[0-9]/.test(x)) x=x.replace(/^0+(?=.)/,""""); out.push(x);} return out;}" & vbLf
    t = t & "const IDX = R.map(r=>"" ""+sozcuk(r.mal+"" ""+r.tn).join("" "")+"" "");" & vbLf
    t = t & "function esles(i,q){const m=IDX[i]; for(const x of q){ if(/^[0-9]/.test(x)){ if(m.indexOf("" ""+x+"" "")<0) return false;} else if(m.indexOf("" ""+x)<0){ if(x===""TRAFO""&&m.indexOf("" TRF "")>=0) continue; if(x===""TRF""&&m.indexOf("" TRAFO"")>=0) continue; return false;} } return true;}" & vbLf
    t = t & "const ipucu=[""50 kva trafo"",""aðaç direk 9"",""kablo alpek 3x35"",""armatür led"",""sayaç monofaze"",""400w""];" & vbLf
    t = t & "document.getElementById(""ipucu"").innerHTML=ipucu.map(x=>`<button>${x}</button>`).join("""");" & vbLf
    t = t & "document.querySelectorAll(""#ipucu button"").forEach(b=>b.onclick=()=>{const a=document.getElementById(""ara""); a.value=b.textContent; ara();});" & vbLf
    V2_S1 = t
End Function

Private Function V2_S2() As String
    Dim t As String
    t = t & "let zaman; document.getElementById(""ara"").addEventListener(""input"",()=>{clearTimeout(zaman); zaman=setTimeout(ara,120);});" & vbLf
    t = t & "function ara(){" & vbLf
    t = t & "  const q=sozcuk(document.getElementById(""ara"").value); const ozet=document.getElementById(""araOzet""), son=document.getElementById(""araSonuc"");" & vbLf
    t = t & "  if(!q.length){ozet.innerHTML=""""; son.innerHTML=`<div class=""bos"">Yazmaya baþlayýn: sonuçlar beþ ilde ve bütün birimlerde anýnda gelir.</div>`; return;}" & vbLf
    t = t & "  const s=[]; for(let i=0;i<R.length;i++){const r=R[i]; if(secIl.has(r.il)&&secB.has(r.b)&&esles(i,q)) s.push(r);}" & vbLf
    t = t & "  s.sort((a,b)=>(IL.indexOf(a.il)-IL.indexOf(b.il))||((a.dp[0]===""D""?0:1)-(b.dp[0]===""D""?0:1))||(b.tl-a.tl));" & vbLf
    t = t & "  const ilO={}; let tmk=0,ttl=0; s.forEach(r=>{ilO[r.il]=(ilO[r.il]||0)+r.mk; tmk+=r.mk; ttl+=r.tl;});" & vbLf
    t = t & "  ozet.innerHTML=`<div class=""ozet-cip"">Bulunan satýr<b>${MK(s.length)}</b></div><div class=""ozet-cip"">Toplam deðer<b>${KISA(ttl)}</b></div>`+IL.filter(i=>ilO[i]).map(i=>`<div class=""ozet-cip"">${ILAD[i]}<b>${MK(ilO[i])}</b></div>`).join("""");" & vbLf
    t = t & "  if(!s.length){son.innerHTML=`<div class=""bos"">Eþleþen stok bulunamadý.</div>`; return;}" & vbLf
    t = t & "  const gos=s.slice(0,400);" & vbLf
    t = t & "  son.innerHTML=`<div class=""tablo-kap""><table class=""t""><thead><tr><th>Ýl</th><th>Birim</th><th>ÜY</th><th>Depo</th><th>Depo adý</th><th>Malzeme</th><th>Taným</th><th>Deðerleme</th><th class=""sag"">Miktar</th><th>Birim</th><th class=""sag"">Deðer</th></tr></thead><tbody>${gos.map(r=>`<tr class=""${r.dp[0]===""D""?""d"":""""}""><td>${ILAD[r.il]}</td><td>${BAD[r.b]}</td><td>${esc(r.uy)}</td><td><b>${esc(r.dp)}</b></td><td>${esc(r.da)}</td><td>${esc(r.mal)}</td><td class=""tanim"">${esc(r.tn)}</td><td>${esc(r.dg)}</td><td class=""sag"">${MK(r.mk)}</td><td>${esc(r.br)}</td><td class=""sag"">${TL(r.tl)}</td></tr>`"
    t = t & ").join("""")}</tbody></table></div>${s.length>400?`<div class=""alt"">Ýlk 400 satýr gösteriliyor (toplam ${MK(s.length)}). Aramayý daraltýn.</div>`:""""}`;" & vbLf
    t = t & "}" & vbLf
    t = t & "// BOLGE DENGESI" & vbLf
    t = t & "function denge(){" & vbLf
    t = t & "  const q=sozcuk(document.getElementById(""dengeAra"").value); const m={};" & vbLf
    t = t & "  R.forEach((r,i)=>{ if(!secB.has(r.b)||degTur(r.dg)!==""YENÝ""||r.mk<=0) return; if(q.length&&!esles(i,q)) return;" & vbLf
    t = t & "    if(!m[r.mal]) m[r.mal]={tn:r.tn,br:r.br,tl:0,il:{}}; const x=m[r.mal]; x.tl+=r.tl; x.il[r.il]=(x.il[r.il]||0)+r.mk;});" & vbLf
    t = t & "  const l=Object.entries(m).filter(([k,v])=>Object.keys(v.il).length>=2).sort((a,b)=>b[1].tl-a[1].tl).slice(0,150);" & vbLf
    t = t & "  const el=document.getElementById(""dengeTablo"");" & vbLf
    t = t & "  if(!l.length){el.innerHTML=`<div class=""bos"">Birden fazla ilde bulunan malzeme yok.</div>`; return;}" & vbLf
    t = t & "  el.innerHTML=`<div class=""tablo-kap""><table class=""t""><thead><tr><th>Malzeme</th><th>Taným</th>${IL.map(i=>`<th class=""sag"">${ILAD[i]}</th>`).join("""")}<th class=""sag"">Toplam deðer</th><th>Yoðunlaþma</th></tr></thead><tbody>${l.map(([k,v])=>{" & vbLf
    t = t & "    const top=IL.reduce((a,i)=>a+(v.il[i]||0),0); let enI=IL[0]; IL.forEach(i=>{if((v.il[i]||0)>(v.il[enI]||0)) enI=i;}); const pay=top?(v.il[enI]||0)/top:0;" & vbLf
    t = t & "    return `<tr><td>${esc(k)}</td><td class=""tanim"">${esc(v.tn)}</td>${IL.map(i=>{const x=v.il[i]||0; const p=top?x/top:0; return `<td class=""sag""><span class=""isi"" style=""background:rgba(15,163,145,${(p*0.85).toFixed(2)});color:${p>0.55?""#fff"":""inherit""}"">${x?MK(x):""–""}</span></td>`;}).join("""")}<td class=""sag""><b>${KISA(v.tl)}</b></td><td>${pay>=0.7?`<span class=""etiket"" style=""background:#fde7c7;color:#9a4b00"">%${Math.round(pay*100)} ${ILAD[enI]}</span>`:`<span style=""color:var(--soluk)"">dengeli</span>`}</td></tr>`;}).join("""")}</tbody></table></div>`;" & vbLf
    t = t & "}" & vbLf
    t = t & "document.getElementById(""dengeAra"").addEventListener(""input"",()=>{clearTimeout(zaman); zaman=setTimeout(denge,150);});" & vbLf
    t = t & "// DEPOLAR" & vbLf
    t = t & "function depolar(){" & vbLf
    t = t & "  const m={}; F().forEach(r=>{const k=r.il+""|""+r.b+""|""+r.uy+""|""+r.dp; if(!m[k]) m[k]={il:r.il,b:r.b,uy:r.uy,dp:r.dp,da:r.da,n:0,tl:0,mal:new Set()}; m[k].n++; m[k].tl+=r.tl; m[k].mal.add(r.mal);});" & vbLf
    t = t & "  const l=Object.values(m).sort((a,b)=>b.tl-a.tl); const mx=l.length?l[0].tl:1;" & vbLf
    t = t & "  document.getElementById(""depoTablo"").innerHTML=l.length?`<div class=""tablo-kap""><table class=""t""><thead><tr><th>#</th><th>Ýl</th><th>Birim</th><th>ÜY</th><th>Depo</th><th>Depo adý</th><th class=""sag"">Satýr</th><th class=""sag"">Çeþit</th><th class=""sag"">Stok deðeri</th><th style=""width:220px""></th></tr></thead><tbody>${l.map((d,i)=>`<tr class=""${d.dp[0]===""D""?""d"":""""}""><td style=""color:var(--soluk)"">${i+1}</td><td>${ILAD[d.il]}</td><td>${BAD[d.b]}</td><td>${esc(d.uy)}</td><td><b>${esc(d.dp)}</b></td><td>${esc(d.da)}</td><td class=""sag"">${MK(d.n)}</td><td class=""sag"">${MK(d.mal.size)}</td><td cl"
    t = t & "ass=""sag""><b>${TL(d.tl)}</b></td><td><span class=""mini"" style=""width:${Math.max(3,d.tl/mx*200)}px""></span></td></tr>`).join("""")}</tbody></table></div>`:`<div class=""bos"">Veri yok.</div>`;" & vbLf
    t = t & "}" & vbLf
    t = t & "function yenile(){genel(); depolar(); denge(); if(document.getElementById(""ara"").value) ara();}" & vbLf
    t = t & "yenile();" & vbLf
    t = t & "})();" & vbLf
    t = t & "</script>" & vbLf
    t = t & "</body>" & vbLf
    t = t & "</html>" & vbLf
    t = t & "" & vbLf
    V2_S2 = t
End Function

' Diger birimlerin Excel ciktilari icin ayni AKSORA imzasi.
Public Sub AKSORA_IMZA_UYGULA(ByVal wb As Workbook)
    T_AksoraImza wb
End Sub

Private Sub T_AksoraImza(ByVal wb As Workbook)
    Dim ws As Worksheet, im As Worksheet, kim As String, no As String, metin As String, pc As Boolean
    On Error Resume Next
    kim = Environ$("USERNAME"): If Len(kim) = 0 Then kim = application.UserName
    Randomize
    no = Format$(Now, "yyyymmddhhnnss") & "-" & Format$(Int(Rnd() * 900000) + 100000, "000000")
    metin = "Bu dosya AKSORA ile üretilmiþtir. Tarih: " & Format$(Now, "dd.mm.yyyy hh:nn") & "  •  Kullanýcý: " & kim & "  •  Ýmza: " & no
    pc = application.PrintCommunication
    application.PrintCommunication = True
    For Each ws In wb.Worksheets
        If ws.Visible = xlSheetVisible Then
            If Len(CStr(ws.Range("A1").Value2)) > 0 Then
                If InStr(1, CStr(ws.Range("A2").Value2), "AKSORA", vbTextCompare) = 0 Then
                    If Len(CStr(ws.Range("A2").Value2)) = 0 Then ws.Range("A2").Value2 = "AKSORA" Else ws.Range("A2").Value2 = CStr(ws.Range("A2").Value2) & "  |  AKSORA"
                End If
                ws.Range("A2").Value2 = Replace(CStr(ws.Range("A2").Value2), "|  AKSORA", "|  AKSORA ile hazýrlanmýþtýr")
                If InStr(1, ws.PageSetup.CenterFooter, "&G", vbTextCompare) = 0 Then
                    ws.PageSetup.CenterFooter = "&8AKSORA ile hazýrlanmýþtýr"
                End If
                ws.Unprotect T_IMZA_SIFRE
                ws.Cells.Locked = False
                ws.Rows("1:2").Locked = True
                If Len(mSayimKilit) > 0 And ws.Name = "SAYIM" Then ws.Range(mSayimKilit).Locked = True
                ws.Protect Password:=T_IMZA_SIFRE, DrawingObjects:=False, Contents:=True, Scenarios:=False, _
                    AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, AllowInsertingRows:=True, _
                    AllowInsertingColumns:=True, AllowDeletingRows:=True, AllowDeletingColumns:=False, AllowSorting:=True, _
                    AllowFiltering:=True, AllowUsingPivotTables:=True
                ws.EnableSelection = xlNoRestrictions
            End If
        End If
    Next ws
    application.PrintCommunication = pc
    wb.BuiltinDocumentProperties("Author").Value = "AKSORA"
    wb.BuiltinDocumentProperties("Company").Value = "AKSORA"
    wb.BuiltinDocumentProperties("Comments").Value = metin
    wb.CustomDocumentProperties("AKSORA_IMZA").Delete
    wb.CustomDocumentProperties.Add Name:="AKSORA_IMZA", LinkToContent:=False, Type:=msoPropertyTypeString, Value:=no & " | " & kim
    Set im = Nothing
    Set im = wb.Worksheets("AKSORA_IMZA")
    If im Is Nothing Then
        Set im = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        im.Name = "AKSORA_IMZA"
    End If
    im.Range("A1").Value2 = metin
    im.Range("A2").Value2 = "Kaynak: " & ThisWorkbook.Name
    im.Visible = xlSheetVeryHidden
    On Error GoTo 0
End Sub

' Excel'e aktar sonrasi AKSORA'ya donuste kullanici ayni sayfada kalir (aktarim sirasinda etkin sayfa degisse bile).
Public Sub AKSORA_DONUS_KAYDET(Optional ByVal zorla As Boolean = False)
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    If zorla Or Len(mDonusSayfa) = 0 Then mDonusSayfa = ActiveSheet.Name
End Sub

Public Sub AKSORA_DONUS_UYGULA()
    Dim ws As Worksheet, ev As Boolean, ad As String
    If Len(mDonusSayfa) = 0 Then Exit Sub
    ad = mDonusSayfa: mDonusSayfa = vbNullString
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then mDonusSayfa = ad: Exit Sub
    If ActiveSheet.Name = ad Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(ad)
    If ws Is Nothing Then Exit Sub
    If ws.Visible <> xlSheetVisible Then Exit Sub
    ev = application.EnableEvents
    application.EnableEvents = False
    ws.Activate
    application.EnableEvents = ev
    AKSORA_BIRIM_AKTIF_GORUNUM
End Sub
