Attribute VB_Name = "modAKSORA_PANEL"
Option Explicit

'============================================================
' AKSORA
' modAKSORA_PANEL
'
' ANA PANEL - Kurumsal kontrol merkezi
'============================================================

Private Const PANEL_SAYFA As String = "ANA PANEL"

Public Sub PANEL_KUR()

    Dim ws As Worksheet

    On Error GoTo Hata

    application.ScreenUpdating = False
    application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets(PANEL_SAYFA)

    PANEL_TEMIZLE ws
    PANEL_ZEMIN ws
    PANEL_UST_BASLIK ws
    PANEL_GENEL_DURUM ws

    PANEL_KART_OLUSTUR ws, "CDP / 7155", "B7:E13", "CDP"
    PANEL_KART_OLUSTUR ws, "EPP / 7055", "G7:J13", "EPP"
    PANEL_KART_OLUSTUR ws, "ÝLÇELER", "L7:O13", "ILCE"
    PANEL_KART_OLUSTUR ws, "YÜKLENÝCÝ", "Q7:T13", "YUK"

    PANEL_ANA_BUTON ws
    PANEL_BAGLANTI_BUTON ws
    PANEL_ALT_BUTONLAR ws
    PANEL_HIZLI_ERISIM ws
    PANEL_FOOTER ws

    PANEL_YENILE
    PANEL_ILK_SEKME_YAP ws

    ws.Activate
    PANEL_GORUNUM_AYARLA ws

Cikis:
    application.EnableEvents = True
    application.ScreenUpdating = True
    Exit Sub

Hata:
    MsgBox _
        "ANA PANEL oluþturulurken hata oluþtu." & vbCrLf & vbCrLf & _
        "Hata No: " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "AKSORA"
    Resume Cikis

End Sub

Public Sub PANEL_GORUNUM_AYARLA(ByVal ws As Worksheet)
    On Error GoTo Cikis
    ws.Activate
    ActiveWindow.DisplayGridlines = False
    AKSORA_EKRAN_SIGDIR ws
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A1").Select
Cikis:
End Sub

Public Sub PANEL_ILK_SEKME_YAP(ByVal ws As Worksheet)

    Dim kapak As Worksheet

    Set kapak = ThisWorkbook.Worksheets("KAPAK")

    If kapak.Index <> 1 Then
        kapak.Move Before:=ThisWorkbook.Sheets(1)
    End If

    If ws.Name = "ANA PANEL" Then
        If ws.Index <> kapak.Index + 1 Then
            ws.Move after:=kapak
        End If
    End If

End Sub

Public Sub PANEL_YENILE()

    Dim ws As Worksheet
    Dim ekran As Boolean
    ekran = application.ScreenUpdating

    On Error GoTo Cikis

    Set ws = ThisWorkbook.Worksheets(PANEL_SAYFA)

    application.ScreenUpdating = False

    PANEL_KART_GUNCELLE _
        ws, "CDP", "CDP_STOK", False

    PANEL_KART_GUNCELLE _
        ws, "EPP", "EPP_STOK", False

    PANEL_SAP_KONTROL_EDILMEDI ws

    PANEL_KART_GUNCELLE _
        ws, "ILCE", "ÝLÇE_STOK", False

    PANEL_KART_GUNCELLE _
        ws, "YUK", "YÜKLENÝCÝ_STOK", False

    PANEL_GENEL_DURUM_GUNCELLE ws
    AKSORA_DURUM_RENK_GUVENLI
    AKSORA_ANA_PANEL_DURUM_RENK_YENILE
    
    With ws.Range("B30:T30")
        .Merge
        .Value = "Panel kontrolü: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
        .Font.Color = RGB(120, 130, 140)
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With
    ws.Shapes("AKSORA15_PANEL_TARIHI").TextFrame2.TextRange.Text = _
        CStr(ws.Range("B30").Value2)
    ws.Range("B30").MergeArea.ClearContents
    AKSORA_PANEL_IL_YAZILARI_YENILE
    AKSORA_PANEL_IL_GORSELI_YENILE
    PANEL_BUTON_BAGLANTILARINI_YENILE
    AKSORA_PANEL_ORTA_HIZALA ws
Cikis:
    application.ScreenUpdating = ekran

End Sub

Public Sub PANEL_CANLI_DURUM(ByVal mesaj As String)

    Dim ws As Worksheet
    Dim eskiEkran As Boolean, eskiOlay As Boolean
    Dim hataNo As Long, hataMetni As String

    eskiEkran = application.ScreenUpdating
    eskiOlay = application.EnableEvents

    On Error GoTo Hata

    application.ScreenUpdating = False
    application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets(PANEL_SAYFA)

    ' Musteri panelindeki gibi: acik sari serit, solda "Guncelleniyor  •  ..." mesaji.
    With ws.Range("B6:AB6")
        .Interior.Color = RGB(255, 246, 219)
        .Font.Color = RGB(146, 64, 14)
        .Font.bold = True
        .Font.Size = 10
        .VerticalAlignment = xlCenter
    End With
    ws.Range("B6").HorizontalAlignment = xlLeft
    ws.Range("B6").IndentLevel = 1

    If InStr(1, mesaj, "TAMAMLANAMADI", vbTextCompare) > 0 Then
        ws.Range("B6").Value = mesaj
    Else
        ws.Range("B6").Value = "Güncelleniyor  •  " & mesaj
    End If

    ThisWorkbook.Activate
    ws.Activate
    UYGULAMA_MODU_AC

    application.ScreenUpdating = True
    DoEvents

Cikis:
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran

    If hataNo <> 0 Then
        Err.Raise hataNo, "PANEL_CANLI_DURUM", hataMetni
    End If
    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description
    Resume Cikis

End Sub

Public Sub PANEL_BAGLANTILARI_KONTROL_ET()

    Dim ws As Worksheet

    On Error GoTo Cikis

    Set ws = ThisWorkbook.Worksheets(PANEL_SAYFA)

    application.ScreenUpdating = False

    PANEL_SAP_DURUM_YAZ ws, "CDP"
    PANEL_SAP_DURUM_YAZ ws, "EPP"

Cikis:
    application.ScreenUpdating = True

End Sub

Public Sub PANEL_TUM_STOKLARI_GUNCELLE()
    On Error GoTo Hata
    AKHZ_TEMIZLE
    AKSORA_PANEL_GUNCELLENIYOR "CDP,EPP,ILCE,YUK"
    ANA_TUM_STOKLARI_GUNCELLE_V2
    AKSORA_PANEL_CANLI = False
    AKSORA_ANA_PANEL_GOSTER
    Exit Sub
Hata:
    AKSORA_PANEL_CANLI = False
    On Error Resume Next
    AKSORA_ANA_PANEL_GOSTER
    On Error GoTo 0
    AKSORA_HATA_GOSTER "Tüm stok güncelleme", Err.Number, Err.Description
End Sub

Public Sub PANEL_CDP_GUNCELLE()
    Dim bas As Double, sessiz As Boolean, no As Long, metin As String
    sessiz = ANA_SESSIZ_MOD
    On Error GoTo Hata
    AKHZ_TEMIZLE
    AKSORA_PANEL_GUNCELLENIYOR "CDP"
    bas = Timer
    ' Ara sonuc penceresi acilmaz; is bitince ana panel ve tek ozet gosterilir.
    ANA_SESSIZ_MOD = True
    CDP_STOK_GUNCELLE_V2
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_SURE_KAYDET bas
    AKSORA_PANEL_CANLI = False
    AKSORA_ANA_PANEL_GOSTER
    MsgBox "CDP stok güncellemesi tamamlandý." & vbCrLf & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(AKSORA_PANEL_SON_SURE(), "0.00") & " sn", vbInformation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_CANLI = False
    On Error Resume Next
    AKSORA_ANA_PANEL_GOSTER
    On Error GoTo 0
    AKSORA_HATA_GOSTER "CDP stok güncelleme", no, metin
End Sub

Public Sub PANEL_EPP_GUNCELLE()
    Dim bas As Double, sessiz As Boolean, no As Long, metin As String
    sessiz = ANA_SESSIZ_MOD
    On Error GoTo Hata
    AKHZ_TEMIZLE
    AKSORA_PANEL_GUNCELLENIYOR "EPP"
    bas = Timer
    ' Ara sonuc penceresi acilmaz; is bitince ana panel ve tek ozet gosterilir.
    ANA_SESSIZ_MOD = True
    EPP_STOK_GUNCELLE_V2
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_SURE_KAYDET bas
    AKSORA_PANEL_CANLI = False
    AKSORA_ANA_PANEL_GOSTER
    MsgBox "EPP stok güncellemesi tamamlandý." & vbCrLf & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(AKSORA_PANEL_SON_SURE(), "0.00") & " sn", vbInformation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_CANLI = False
    On Error Resume Next
    AKSORA_ANA_PANEL_GOSTER
    On Error GoTo 0
    AKSORA_HATA_GOSTER "EPP stok güncelleme", no, metin
End Sub

Public Sub PANEL_ILCE_GUNCELLE()
    Dim bas As Double, sessiz As Boolean, no As Long, metin As String
    sessiz = ANA_SESSIZ_MOD
    On Error GoTo Hata
    AKHZ_TEMIZLE
    AKSORA_PANEL_GUNCELLENIYOR "ILCE"
    bas = Timer
    ' Ara sonuc penceresi acilmaz; is bitince ana panel ve tek ozet gosterilir.
    ANA_SESSIZ_MOD = True
    ILCE_STOK_GUNCELLE_V2
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_SURE_KAYDET bas
    AKSORA_PANEL_CANLI = False
    AKSORA_ANA_PANEL_GOSTER
    MsgBox "Ýlçe stok güncellemesi tamamlandý." & vbCrLf & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(AKSORA_PANEL_SON_SURE(), "0.00") & " sn", vbInformation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_CANLI = False
    On Error Resume Next
    AKSORA_ANA_PANEL_GOSTER
    On Error GoTo 0
    AKSORA_HATA_GOSTER "Ýlçe stok güncelleme", no, metin
End Sub

Public Sub PANEL_YUKLENICI_GUNCELLE()
    Dim bas As Double, sessiz As Boolean, no As Long, metin As String
    sessiz = ANA_SESSIZ_MOD
    On Error GoTo Hata
    AKHZ_TEMIZLE
    AKSORA_PANEL_GUNCELLENIYOR "YUK"
    bas = Timer
    ' Ara sonuc penceresi acilmaz; is bitince ana panel ve tek ozet gosterilir.
    ANA_SESSIZ_MOD = True
    YUKLENICI_STOK_GUNCELLE_V2
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_SURE_KAYDET bas
    AKSORA_PANEL_CANLI = False
    AKSORA_ANA_PANEL_GOSTER
    MsgBox "Yüklenici stok güncellemesi tamamlandý." & vbCrLf & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(AKSORA_PANEL_SON_SURE(), "0.00") & " sn", vbInformation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_CANLI = False
    On Error Resume Next
    AKSORA_ANA_PANEL_GOSTER
    On Error GoTo 0
    AKSORA_HATA_GOSTER "Yüklenici stok güncelleme", no, metin
End Sub

Public Sub PANEL_GIT_STOKLAR()
    On Error GoTo Hata
    PANEL_SAYFAYA_GIT "STOKLAR"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Sayfa geçiþi", Err.Number, Err.Description
End Sub

Public Sub PANEL_GIT_RAPORLAR()
    On Error GoTo Hata
    PANEL_SAYFAYA_GIT "RAPORLAR"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Sayfa geçiþi", Err.Number, Err.Description
End Sub

Public Sub PANEL_GIT_CDP()
    On Error GoTo Hata
    PANEL_SAYFAYA_GIT "CDP_STOK"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Sayfa geçiþi", Err.Number, Err.Description
End Sub

Public Sub PANEL_GIT_EPP()
    On Error GoTo Hata
    PANEL_SAYFAYA_GIT "EPP_STOK"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Sayfa geçiþi", Err.Number, Err.Description
End Sub

Public Sub PANEL_GIT_ILCE()
    On Error GoTo Hata
    PANEL_SAYFAYA_GIT "ÝLÇE_STOK"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Sayfa geçiþi", Err.Number, Err.Description
End Sub

Public Sub PANEL_GIT_YUKLENICI()
    On Error GoTo Hata
    PANEL_SAYFAYA_GIT "YÜKLENÝCÝ_STOK"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Sayfa geçiþi", Err.Number, Err.Description
End Sub

Private Sub PANEL_TEMIZLE(ByVal ws As Worksheet)

    Dim i As Long

    ws.Cells.UnMerge
    ws.Cells.Clear

    On Error Resume Next

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    On Error GoTo 0

End Sub

Private Sub PANEL_ZEMIN(ByVal ws As Worksheet)

    With ws

        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 10

        .Rows.RowHeight = 20

        .Columns("A").ColumnWidth = 2.5
        .Columns("B:T").ColumnWidth = 8.6
        .Columns("U").ColumnWidth = 2.5

        .Rows("1:3").RowHeight = 24
        .Rows("4:5").RowHeight = 22
        .Rows("6").RowHeight = 10
        .Rows("7:13").RowHeight = 20
        .Rows("14").RowHeight = 12
        .Rows("15:18").RowHeight = 23
        .Rows("19").RowHeight = 12
        .Rows("20:23").RowHeight = 22
        .Rows("24").RowHeight = 12
        .Rows("25:27").RowHeight = 20
        .Rows("28:30").RowHeight = 20

        .Range("A1:U30").Interior.Color = RGB(245, 247, 250)

    End With

End Sub

Private Sub PANEL_UST_BASLIK(ByVal ws As Worksheet)

    With ws.Range("B2:T2")
        .Merge
        .Value = "ÇORUH ELEKTRÝK DAÐITIM A.Þ."
        .Interior.Color = RGB(20, 55, 95)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B3:T4")
        .Merge
        .Value = "AKSORA STOK YÖNETÝM PANELÝ"
        .Interior.Color = RGB(20, 55, 95)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Size = 22
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B5:T5")
        .Merge
        .Value = PANEL_AKTIF_IL_ADI() & " Ýl Müdürlüðü  •  Malzeme ve Stok Yönetimi  •  v3.0"
        .Interior.Color = RGB(20, 55, 95)
        .Font.Color = RGB(215, 225, 238)
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B30:T30")
        .Merge
        .Value = "Panel kontrolü: " & Format(Now, "dd.mm.yyyy hh:mm:ss")
        .Font.Color = RGB(120, 130, 140)
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
    End With

    ws.Range("T2").HorizontalAlignment = xlRight
    ws.Range("T2").Font.Size = 8

End Sub

Private Sub PANEL_GENEL_DURUM(ByVal ws As Worksheet)

    With ws.Range("B6:T6")
        .Merge
        .Value = "GENEL DURUM HAZIRLANIYOR..."
        .Interior.Color = RGB(225, 230, 236)
        .Font.Color = RGB(55, 65, 75)
        .Font.bold = True
        .Font.Size = 10
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

End Sub

Private Sub PANEL_KART_OLUSTUR( _
ByVal ws As Worksheet, _
ByVal baslik As String, _
ByVal alan As String, _
ByVal kod As String)

    Dim r As Range
    Dim IlkSatir As Long
    Dim ilkSutun As Long
    Dim sonSutun As Long

    Set r = ws.Range(alan)

    IlkSatir = r.Row
    ilkSutun = r.Column
    sonSutun = r.Column + r.Columns.Count - 1

    With r
        .Interior.Color = RGB(255, 255, 255)
        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(218, 223, 230)
        .Borders.Weight = xlThin
    End With

    With ws.Range( _
        ws.Cells(IlkSatir, ilkSutun), _
        ws.Cells(IlkSatir, sonSutun))

        .Merge
        .Value = baslik
        .Interior.Color = RGB(20, 55, 95)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Size = 12
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    PANEL_KART_SATIR ws, IlkSatir + 1, ilkSutun, sonSutun, _
        "DURUM", kod & "_DURUM"

    PANEL_KART_SATIR ws, IlkSatir + 2, ilkSutun, sonSutun, _
        "SON GÜNCELLEME", kod & "_TARIH"

    PANEL_KART_SATIR ws, IlkSatir + 3, ilkSutun, sonSutun, _
        "STOK SATIRI", kod & "_SATIR"

    PANEL_KART_SATIR ws, IlkSatir + 4, ilkSutun, sonSutun, _
        "STOK DEÐERÝ", kod & "_DEGER"

    If kod = "CDP" Or kod = "EPP" Then

        PANEL_KART_SATIR ws, IlkSatir + 5, ilkSutun, sonSutun, _
            "SAP BAÐLANTI", kod & "_SAP"

    Else

        With ws.Range( _
            ws.Cells(IlkSatir + 5, ilkSutun), _
            ws.Cells(IlkSatir + 5, sonSutun))

            .Merge
            .Value = "STOK GRUBU"
            .Font.Color = RGB(125, 135, 145)
            .HorizontalAlignment = xlCenter

        End With

    End If

    With ws.Range( _
        ws.Cells(IlkSatir + 6, ilkSutun), _
        ws.Cells(IlkSatir + 6, sonSutun))

        .Merge
        .Value = "Hazýr"
        .Font.Color = RGB(125, 135, 145)
        .Font.Size = 8
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub

Private Sub PANEL_KART_SATIR( _
ByVal ws As Worksheet, _
ByVal satir As Long, _
ByVal ilkSutun As Long, _
ByVal sonSutun As Long, _
ByVal etiket As String, _
ByVal ad As String)

    Dim Orta As Long

    Orta = ilkSutun + 1

    With ws.Range( _
        ws.Cells(satir, ilkSutun), _
        ws.Cells(satir, Orta))

        .Merge
        .Value = etiket
        .Font.bold = True
        .Font.Size = 8
        .Font.Color = RGB(95, 105, 115)
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    With ws.Range( _
        ws.Cells(satir, Orta + 1), _
        ws.Cells(satir, sonSutun))

        .Merge
        .Value = "-"
        .Name = ad
        .Font.bold = True
        .HorizontalAlignment = xlRight
        .VerticalAlignment = xlCenter

    End With

End Sub

Private Sub PANEL_ANA_BUTON(ByVal ws As Worksheet)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("G15").Left, _
        ws.Range("G15").Top, _
        ws.Range("N18").Left + ws.Range("N18").Width - ws.Range("G15").Left, _
        ws.Range("N18").Top + ws.Range("N18").Height - ws.Range("G15").Top)

    With shp

        .Name = "btnTumStok"
        .TextFrame2.TextRange.Text = "TÜM STOKLARI GÜNCELLE"

        .Fill.ForeColor.RGB = RGB(20, 55, 95)
        .line.ForeColor.RGB = RGB(20, 55, 95)

        .TextFrame2.TextRange.Font.Name = "Aptos Display"
        .TextFrame2.TextRange.Font.Size = 16
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        .OnAction = "'" & ThisWorkbook.Name & _
            "'!PANEL_TUM_STOKLARI_GUNCELLE"

        .Placement = xlFreeFloating

    End With

End Sub

Private Sub PANEL_BAGLANTI_BUTON(ByVal ws As Worksheet)

    Dim shp As Shape

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        ws.Range("B15").Left, _
        ws.Range("B15").Top, _
        ws.Range("E18").Left + ws.Range("E18").Width - ws.Range("B15").Left, _
        ws.Range("E18").Top + ws.Range("E18").Height - ws.Range("B15").Top)

    With shp

        .Name = "btnBaglantiKontrol"
        .TextFrame2.TextRange.Text = "SAP BAÐLANTILARINI KONTROL ET"

        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .line.ForeColor.RGB = RGB(45, 95, 145)
        .line.Weight = 1.5

        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(20, 55, 95)

        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        .OnAction = "'" & ThisWorkbook.Name & _
            "'!PANEL_BAGLANTILARI_KONTROL_ET"

        .Placement = xlFreeFloating

    End With

End Sub

Private Sub PANEL_ALT_BUTONLAR(ByVal ws As Worksheet)

    PANEL_BUTON ws, "B20:E21", _
        "CDP GÜNCELLE", _
        "PANEL_CDP_GUNCELLE", _
        "btnCDP"

    PANEL_BUTON ws, "G20:J21", _
        "EPP GÜNCELLE", _
        "PANEL_EPP_GUNCELLE", _
        "btnEPP"

    PANEL_BUTON ws, "L20:O21", _
        "ÝLÇELERÝ GÜNCELLE", _
        "PANEL_ILCE_GUNCELLE", _
        "btnIlce"

    PANEL_BUTON ws, "Q20:T21", _
        "YÜKLENÝCÝYÝ GÜNCELLE", _
        "PANEL_YUKLENICI_GUNCELLE", _
        "btnYuk"

End Sub

Private Sub PANEL_HIZLI_ERISIM(ByVal ws As Worksheet)

    With ws.Range("B23:T23")
        .Merge
        .Value = "HIZLI ERÝÞÝM"
        .Font.bold = True
        .Font.Size = 10
        .Font.Color = RGB(75, 85, 95)
        .HorizontalAlignment = xlCenter
    End With

    PANEL_BUTON_ACIK ws, "B25:D26", _
        "STOKLAR", "PANEL_GIT_STOKLAR", "goStok"

    PANEL_BUTON_ACIK ws, "E25:G26", _
        "RAPORLAR", "PANEL_GIT_RAPORLAR", "goRapor"

    PANEL_BUTON_ACIK ws, "H25:J26", _
        "CDP STOK", "PANEL_GIT_CDP", "goCDP"

    PANEL_BUTON_ACIK ws, "K25:M26", _
        "EPP STOK", "PANEL_GIT_EPP", "goEPP"

    PANEL_BUTON_ACIK ws, "N25:P26", _
        "ÝLÇE STOK", "PANEL_GIT_ILCE", "goIlce"

    PANEL_BUTON_ACIK ws, "Q25:T26", _
        "YÜKLENÝCÝ STOK", "PANEL_GIT_YUKLENICI", "goYuk"

End Sub

Private Sub PANEL_FOOTER(ByVal ws As Worksheet)

    With ws.Range("B28:T28")
        .Merge
        .Value = "ÇORUH ELEKTRÝK DAÐITIM A.Þ.  |  " & PANEL_AKTIF_IL_ADI() & " Ýl Müdürlüðü"
        .Interior.Color = RGB(20, 55, 95)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    With ws.Range("B29:T29")
        .Merge
        .Value = ""
        .Interior.Color = RGB(20, 55, 95)
        .Font.Color = RGB(215, 225, 238)
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

End Sub

Private Sub PANEL_BUTON( _
ByVal ws As Worksheet, _
ByVal alan As String, _
ByVal metin As String, _
ByVal makro As String, _
ByVal ad As String)

    Dim r As Range
    Dim shp As Shape

    Set r = ws.Range(alan)

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        r.Left, r.Top, r.Width, r.Height)

    With shp

        .Name = ad
        .TextFrame2.TextRange.Text = metin

        .Fill.ForeColor.RGB = RGB(45, 95, 145)
        .line.ForeColor.RGB = RGB(45, 95, 145)

        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        .OnAction = "'" & ThisWorkbook.Name & "'!" & makro
        .Placement = xlFreeFloating

    End With

End Sub

Private Sub PANEL_BUTON_ACIK( _
ByVal ws As Worksheet, _
ByVal alan As String, _
ByVal metin As String, _
ByVal makro As String, _
ByVal ad As String)

    Dim r As Range
    Dim shp As Shape

    Set r = ws.Range(alan)

    Set shp = ws.Shapes.AddShape( _
        msoShapeRoundedRectangle, _
        r.Left, r.Top, r.Width, r.Height)

    With shp

        .Name = ad
        .TextFrame2.TextRange.Text = metin

        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .line.ForeColor.RGB = RGB(185, 195, 205)

        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(20, 55, 95)

        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

        .OnAction = "'" & ThisWorkbook.Name & "'!" & makro
        .Placement = xlFreeFloating

    End With

End Sub

Private Sub PANEL_KART_GUNCELLE( _
ByVal wsPanel As Worksheet, _
ByVal kod As String, _
ByVal StokSayfa As String, _
ByVal SAPKontrol As Boolean)

    Dim ws As Worksheet
    Dim tarih As Variant
    Dim satir As Long
    Dim deger As Double
    Dim GunFarki As Long
    Dim durum As String

    On Error GoTo VERI_YOK

    Set ws = ThisWorkbook.Worksheets(StokSayfa)

    tarih = ws.Range("C2").Value
    satir = PANEL_LONG(ws.Range("C3").Value2)
    deger = PANEL_DOUBLE(ws.Range("C4").Value2)

    If IsDate(tarih) Then

        GunFarki = DateDiff("d", DateValue(CDate(tarih)), Date)

        If GunFarki <= 0 Then
            durum = "GÜNCEL"
            PANEL_DURUM_RENK wsPanel.Range(kod & "_DURUM"), 1
        ElseIf GunFarki = 1 Then
            durum = "1 GÜN ÖNCE"
            PANEL_DURUM_RENK wsPanel.Range(kod & "_DURUM"), 2
        Else
            durum = GunFarki & " GÜN ÖNCE"
            PANEL_DURUM_RENK wsPanel.Range(kod & "_DURUM"), 3
        End If

        wsPanel.Range(kod & "_TARIH").Value = _
            Format(CDate(tarih), "dd.mm.yyyy hh:mm")

    Else

        durum = "GÜNCEL DEÐÝL"
        PANEL_DURUM_RENK wsPanel.Range(kod & "_DURUM"), 3
        wsPanel.Range(kod & "_TARIH").Value = "-"

    End If

    wsPanel.Range(kod & "_DURUM").Value = durum
    With wsPanel.Range(kod & "_SATIR")
        .NumberFormat = "@"
        .Value = PANEL_SAYI_TR(CDbl(satir), 0)
    End With
    wsPanel.Range(kod & "_DEGER").Value = _
        PANEL_SAYI_TR(deger, 2) & " TL"

    If SAPKontrol Then
        PANEL_SAP_DURUM_YAZ wsPanel, kod
    End If

    If AKSORA_PANEL_KART_CANLI(kod) Then AKSORA_PANEL_KART_GUNCELLENIYOR wsPanel, kod

    Exit Sub

VERI_YOK:

    On Error Resume Next

    wsPanel.Range(kod & "_DURUM").Value = "VERÝ YOK"
    PANEL_DURUM_RENK wsPanel.Range(kod & "_DURUM"), 3

    wsPanel.Range(kod & "_TARIH").Value = "-"
    wsPanel.Range(kod & "_SATIR").Value = "-"
    wsPanel.Range(kod & "_DEGER").Value = "-"

    If SAPKontrol Then
        wsPanel.Range(kod & "_SAP").Value = "BAÐLI DEÐÝL"
        PANEL_DURUM_RENK wsPanel.Range(kod & "_SAP"), 3
    End If

End Sub

Private Sub PANEL_SAP_KONTROL_EDILMEDI(ByVal ws As Worksheet)

    On Error Resume Next

    ws.Range("CDP_SAP").Value = "KONTROL EDÝLMEDÝ"
    ws.Range("EPP_SAP").Value = "KONTROL EDÝLMEDÝ"

    ws.Range("CDP_SAP").Font.Color = RGB(125, 135, 145)
    ws.Range("EPP_SAP").Font.Color = RGB(125, 135, 145)

    ws.Range("CDP_SAP").Font.bold = True
    ws.Range("EPP_SAP").Font.bold = True

    On Error GoTo 0

End Sub

Private Sub PANEL_SAP_DURUM_YAZ( _
ByVal wsPanel As Worksheet, _
ByVal kod As String)

    Dim Bagli As Boolean

    Bagli = PANEL_SAP_BAGLI_MI(kod)

    If Bagli Then

        wsPanel.Range(kod & "_SAP").Value = "BAÐLI"
        PANEL_DURUM_RENK wsPanel.Range(kod & "_SAP"), 1

    Else

        wsPanel.Range(kod & "_SAP").Value = "BAÐLI DEÐÝL"
        PANEL_DURUM_RENK wsPanel.Range(kod & "_SAP"), 3

    End If

End Sub

Private Function PANEL_SAP_BAGLI_MI( _
ByVal sistem As String) As Boolean

    Dim SapGuiAuto As Object
    Dim app As Object
    Dim con As Object
    Dim ses As Object
    Dim wnd As Object

    Dim i As Long
    Dim j As Long

    Dim SysName As String
    Dim Kullanici As String

    On Error GoTo Yok

    Set SapGuiAuto = GetObject("SAPGUI")
    If SapGuiAuto Is Nothing Then GoTo Yok

    Set app = SapGuiAuto.GetScriptingEngine
    If app Is Nothing Then GoTo Yok

    For i = 0 To app.Children.Count - 1

        Set con = Nothing

        On Error Resume Next
        Set con = app.Children(CLng(i))
        On Error GoTo Yok

        If Not con Is Nothing Then

            For j = 0 To con.Children.Count - 1

                Set ses = Nothing
                Set wnd = Nothing
                SysName = ""
                Kullanici = ""

                On Error Resume Next
                Set ses = con.Children(CLng(j))

                If Not ses Is Nothing Then
                    SysName = UCase$(Trim$(CStr(ses.Info.SystemName)))
                    Kullanici = Trim$(CStr(ses.Info.User))
                    Set wnd = ses.FindById("wnd[0]", False)
                End If
                On Error GoTo Yok

                If Not ses Is Nothing Then
                    If Not wnd Is Nothing Then
                        If Len(Kullanici) > 0 Then
                            If StrComp(SysName, UCase$(Trim$(sistem)), vbTextCompare) = 0 Then
                                PANEL_SAP_BAGLI_MI = True
                                Exit Function
                            End If
                        End If
                    End If
                End If

            Next j

        End If

    Next i

Yok:

End Function

Private Sub PANEL_GENEL_DURUM_GUNCELLE( _
ByVal ws As Worksheet)

    Dim CDPDurum As String
    Dim EPPDurum As String
    Dim IlceDurum As String
    Dim YukDurum As String

    Dim ToplamSatir As Long
    Dim toplamDeger As Double
    Dim TumGuncel As Boolean

    On Error GoTo Hata

    CDPDurum = Trim$(CStr(ws.Range("CDP_DURUM").Cells(1, 1).Value))
    EPPDurum = Trim$(CStr(ws.Range("EPP_DURUM").Cells(1, 1).Value))
    IlceDurum = Trim$(CStr(ws.Range("ILCE_DURUM").Cells(1, 1).Value))
    YukDurum = Trim$(CStr(ws.Range("YUK_DURUM").Cells(1, 1).Value))

    ToplamSatir = _
        PANEL_SAYFADAN_LONG("STOKLAR", "C3") + _
        PANEL_SAYFADAN_LONG("ÝLÇE_STOK", "C3") + _
        PANEL_SAYFADAN_LONG("YÜKLENÝCÝ_STOK", "C3")

    toplamDeger = _
        PANEL_SAYFADAN_DOUBLE("STOKLAR", "C4") + _
        PANEL_SAYFADAN_DOUBLE("ÝLÇE_STOK", "C4") + _
        PANEL_SAYFADAN_DOUBLE("YÜKLENÝCÝ_STOK", "C4")

    TumGuncel = _
        (StrComp(CDPDurum, "GÜNCEL", vbTextCompare) = 0) And _
        (StrComp(EPPDurum, "GÜNCEL", vbTextCompare) = 0) And _
        (StrComp(IlceDurum, "GÜNCEL", vbTextCompare) = 0) And _
        (StrComp(YukDurum, "GÜNCEL", vbTextCompare) = 0)

    ' Musteri panelindeki gibi: tarih, sure, grup satirlari ve toplam deger.
    Dim enSon As Date, tv As Variant, sayfa As Variant, sure As Double, ozet As String
    For Each sayfa In Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
        tv = Empty
        On Error Resume Next
        tv = ThisWorkbook.Worksheets(CStr(sayfa)).Range("C2").Value
        On Error GoTo Hata
        If IsDate(tv) Then
            If CDate(tv) > enSon Then enSon = CDate(tv)
        End If
    Next sayfa
    If enSon > 0 Then ozet = "Son güncelleme " & Format$(enSon, "dd.mm.yyyy hh:nn") Else ozet = "Henüz SAP güncellemesi yapýlmadý"
    sure = AKSORA_PANEL_SON_SURE()
    ' Seritte tek sure: bütün illerin toplam guncelleme suresi (yoksa son il guncellemesi).
    If AKSORA_TOPLU_SON_SURE("_AKSORA_MALZEME_TOPLU_SURE") > 0 Then sure = AKSORA_TOPLU_SON_SURE("_AKSORA_MALZEME_TOPLU_SURE")
    If sure > 0 Then ozet = ozet & "   |   Güncelleme süresi " & AKSORA_SURE_METNI(sure)
    ozet = ozet & "   |   CDP " & PANEL_SAYI_TR(PANEL_SAYFADAN_LONG("CDP_STOK", "C3"), 0) & " satýr" & _
        "   |   EPP " & PANEL_SAYI_TR(PANEL_SAYFADAN_LONG("EPP_STOK", "C3"), 0) & " satýr" & _
        "   |   Ýlçe " & PANEL_SAYI_TR(PANEL_SAYFADAN_LONG("ÝLÇE_STOK", "C3"), 0) & " satýr" & _
        "   |   Yüklenici " & PANEL_SAYI_TR(PANEL_SAYFADAN_LONG("YÜKLENÝCÝ_STOK", "C3"), 0) & " satýr" & _
        "   |   Toplam stok deðeri " & PANEL_SAYI_TR(toplamDeger, 2) & " TL"
    ' Guncelleme surerken serit canli mesaji gosterir; ozet bitince yazilir.
    If AKSORA_PANEL_CANLI Then Exit Sub
    With ws.Range("B6")
        If TumGuncel Then
            .Value = ozet
            .Interior.Color = RGB(226, 239, 218)
            .Font.Color = RGB(40, 110, 65)
        Else
            .Value = "Kontrol gerekiyor  •  " & ozet
            .Interior.Color = RGB(255, 235, 210)
            .Font.Color = RGB(170, 90, 20)
        End If
        .Font.bold = True
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .VerticalAlignment = xlCenter
    End With

    Exit Sub

Hata:
    With ws.Range("B6")
        .Value = "PANEL DURUMU KONTROL EDÝLEMEDÝ"
        .Interior.Color = RGB(255, 235, 210)
        .Font.Color = RGB(170, 90, 20)
        .Font.bold = True
    End With

End Sub

Private Sub PANEL_DURUM_RENK( _
ByVal r As Range, _
ByVal tip As Long)

    Select Case tip

        Case 1
            r.Font.Color = RGB(40, 130, 75)

        Case 2
            r.Font.Color = RGB(190, 130, 20)

        Case Else
            r.Font.Color = RGB(190, 55, 55)

    End Select

    r.Font.bold = True

End Sub

Private Sub PANEL_SAYFAYA_GIT(ByVal sayfaAdi As String)
    On Error GoTo Hata
    AKSORA_V12_GIT sayfaAdi
    Exit Sub
Hata:
    MsgBox "Sayfa açýlamadý: " & sayfaAdi & vbCrLf & Err.Description, vbExclamation, "AKSORA"
End Sub
Private Function PANEL_LONG(ByVal deger As Variant) As Long

    If IsNumeric(deger) Then
        PANEL_LONG = CLng(deger)
    Else
        PANEL_LONG = 0
    End If

End Function

Private Function PANEL_DOUBLE(ByVal deger As Variant) As Double

    If IsNumeric(deger) Then
        PANEL_DOUBLE = CDbl(deger)
    Else
        PANEL_DOUBLE = 0
    End If

End Function

Private Function PANEL_SAYFADAN_LONG( _
ByVal sayfa As String, _
ByVal hucre As String) As Long

    On Error GoTo Cikis

    PANEL_SAYFADAN_LONG = _
        PANEL_LONG(ThisWorkbook.Worksheets(sayfa).Range(hucre).Value2)

Cikis:

End Function

Private Function PANEL_SAYFADAN_DOUBLE( _
ByVal sayfa As String, _
ByVal hucre As String) As Double

    On Error GoTo Cikis

    PANEL_SAYFADAN_DOUBLE = _
        PANEL_DOUBLE(ThisWorkbook.Worksheets(sayfa).Range(hucre).Value2)

Cikis:

End Function

'============================================================
' AKTÝF ÝL ADI
'============================================================
Private Function PANEL_AKTIF_IL_ADI() As String

    Dim kod As String

    On Error GoTo Cikis

    kod = UCase$(Trim$(CStr( _
        ThisWorkbook.Worksheets("KAPAK").Range("AZ1").Value2)))

    Select Case kod
        Case "ARTVIN"
            PANEL_AKTIF_IL_ADI = "Artvin"
        Case "RIZE"
            PANEL_AKTIF_IL_ADI = "Rize"
        Case "TRABZON"
            PANEL_AKTIF_IL_ADI = "Trabzon"
        Case "GUMUSHANE"
            PANEL_AKTIF_IL_ADI = "Gümüþhane"
        Case "GIRESUN"
            PANEL_AKTIF_IL_ADI = "Giresun"
        Case Else
            PANEL_AKTIF_IL_ADI = kod
    End Select

Cikis:

End Function


'============================================================
' PANEL SAYI GÖSTERÝMÝ: 1.234.567,89
' Excel bölgesel ayarlarýndan baðýmsýzdýr.
'============================================================
Private Function PANEL_SAYI_TR( _
    ByVal sayi As Double, _
    ByVal ondalik As Long) As String

    Dim Yuvarlanmis As Double
    Dim tam As Double
    Dim kesir As Long
    Dim Carpan As Long
    Dim Rakamlar As String
    Dim sonuc As String
    Dim i As Long

    ' Excel'in bölgesel ayarlarýna baðlý olmadan nokta/virgül üret.
    Carpan = CLng(10 ^ ondalik)
    Yuvarlanmis = application.WorksheetFunction.Round(Abs(sayi), ondalik)
    tam = Fix(Yuvarlanmis)
    kesir = CLng(application.WorksheetFunction.Round( _
        (Yuvarlanmis - tam) * Carpan, 0))

    If kesir >= Carpan Then
        tam = tam + 1
        kesir = 0
    End If

    Rakamlar = Format$(tam, "0")

    For i = 1 To Len(Rakamlar)
        If i > 1 And (Len(Rakamlar) - i + 1) Mod 3 = 0 Then
            sonuc = sonuc & "."
        End If
        sonuc = sonuc & Mid$(Rakamlar, i, 1)
    Next i

    If ondalik > 0 Then
        sonuc = sonuc & "," & Right$( _
            String$(ondalik, "0") & CStr(kesir), ondalik)
    End If

    If sayi < 0 Then sonuc = "-" & sonuc

    PANEL_SAYI_TR = sonuc

End Function

Public Sub AKSORA_DURUM_RENK_GUVENLI()

    Dim ws As Worksheet
    Dim kaynak As Worksheet
    Dim shp As Shape

    Dim sayfalar As Variant
    Dim kodlar As Variant
    Dim sutunlar As Variant

    Dim i As Long
    Dim gun As Long
    Dim tip As Long
    Dim enKotu As Long

    Dim tarih As Variant
    Dim renk As Long
    Dim bantRenk As Long
    Dim bantYazi As Long

    Dim durum As String

    Dim eskiOlay As Boolean
    Dim eskiEkran As Boolean

    Dim korumali As Boolean
    Dim icerik As Boolean
    Dim senaryo As Boolean
    Dim filtre As Boolean
    Dim sirala As Boolean
    Dim pivot As Boolean
    Dim bicim As Boolean

    Dim hataMesaji As String

    Dim sapMetni As String
    Dim sapRenk As Long
    Dim sapHucre As Range
    Dim kartAlani As Range


    On Error GoTo Hata


    Set ws = PANEL_SAYFA_BUL("ANA PANEL")

    If ws Is Nothing Then Exit Sub


    sayfalar = Array( _
        "CDP_STOK", _
        "EPP_STOK", _
        "ILCE_STOK", _
        "YUKLENICI_STOK" _
    )

    kodlar = Array( _
        "CDP", _
        "EPP", _
        "ILCE", _
        "YUK" _
    )

    sutunlar = Array( _
        "B", _
        "G", _
        "L", _
        "Q" _
    )


    eskiOlay = application.EnableEvents
    eskiEkran = application.ScreenUpdating

    application.EnableEvents = False
    application.ScreenUpdating = False


    korumali = _
        ws.ProtectDrawingObjects Or _
        ws.ProtectContents Or _
        ws.ProtectScenarios


    If korumali Then

        icerik = ws.ProtectContents
        senaryo = ws.ProtectScenarios

        filtre = ws.Protection.AllowFiltering
        sirala = ws.Protection.AllowSorting
        pivot = ws.Protection.AllowUsingPivotTables
        bicim = ws.Protection.AllowFormattingCells

        ws.Unprotect Password:=""

    End If


    enKotu = 1


    For i = 0 To 3

        Set kaynak = Nothing


        Select Case CStr(sayfalar(i))

            Case "ILCE_STOK"

                Set kaynak = PANEL_SAYFA_BUL("ILCE_STOK")

                If kaynak Is Nothing Then
                    Set kaynak = PANEL_SAYFA_BUL("ÝLÇE_STOK")
                End If


            Case "YUKLENICI_STOK"

                Set kaynak = PANEL_SAYFA_BUL("YUKLENICI_STOK")

                If kaynak Is Nothing Then
                    Set kaynak = PANEL_SAYFA_BUL("YÜKLENÝCÝ_STOK")
                End If


            Case Else

                Set kaynak = _
                    PANEL_SAYFA_BUL(CStr(sayfalar(i)))

        End Select


        If kaynak Is Nothing Then

            tip = 3
            durum = "VERÝ YOK"

        Else

            tarih = kaynak.Range("C2").Value


            If IsDate(tarih) Then

                gun = _
                    DateDiff( _
                        "d", _
                        DateValue(CDate(tarih)), _
                        Date _
                    )


                If gun < 0 Then gun = 0


                If gun = 0 Then

                    tip = 1
                    durum = "GÜNCEL"


                ElseIf gun = 1 Then

                    tip = 2
                    durum = "1 GÜN ÖNCE"


                Else

                    tip = 3
                    durum = CStr(gun) & " GÜN ÖNCE"

                End If


            Else

                tip = 3
                durum = "GÜNCEL DEÐÝL"

            End If

        End If


        If tip > enKotu Then
            enKotu = tip
        End If


        Select Case tip

            Case 1
                renk = RGB(0, 133, 92)

            Case 2
                renk = RGB(190, 120, 15)

            Case Else
                renk = RGB(190, 55, 55)

        End Select


        On Error Resume Next

        With ws.Range(CStr(kodlar(i)) & "_DURUM")

            .Value = durum
            .Font.Color = renk
            .Font.bold = True

        End With

        On Error GoTo Hata


        Set shp = Nothing

        On Error Resume Next

        Set shp = _
            ws.Shapes( _
                "AKSORA15_KART_DEGER_" & _
                CStr(sutunlar(i)) & _
                "0" _
            )

        On Error GoTo Hata


        If Not shp Is Nothing Then

            shp.TextFrame2.TextRange.Text = durum

            shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                renk

        End If

        If AKSORA_PANEL_KART_CANLI(CStr(kodlar(i))) Then AKSORA_PANEL_KART_GUNCELLENIYOR ws, CStr(kodlar(i))

    Next i


    Select Case enKotu

        Case 1

            bantRenk = RGB(213, 244, 233)
            bantYazi = RGB(8, 86, 71)


        Case 2

            bantRenk = RGB(255, 235, 210)
            bantYazi = RGB(170, 90, 20)


        Case Else

            bantRenk = RGB(255, 218, 218)
            bantYazi = RGB(165, 37, 37)

    End Select


    If Not AKSORA_PANEL_CANLI Then
        With ws.Range("B6:AB6")
            .Interior.Color = bantRenk
            .Font.Color = bantYazi
            .Font.bold = True
        End With
    End If


    '======================================================
    ' SAP DURUM RENKLERI
    '======================================================
    For i = 0 To 1

        Set sapHucre = Nothing

        On Error Resume Next
        Set sapHucre = ws.Range(CStr(kodlar(i)) & "_SAP")
        On Error GoTo Hata


        If Not sapHucre Is Nothing Then

            sapMetni = _
                UCase$( _
                    Trim$( _
                        CStr( _
                            sapHucre.Cells(1, 1).Value2 _
                        ) _
                    ) _
                )


            Select Case sapMetni

                Case "BAÐLI"

                    sapRenk = RGB(0, 145, 74)


                Case "BAÐLI DEÐÝL"

                    sapRenk = RGB(205, 35, 45)


                Case Else

                    sapRenk = RGB(125, 135, 145)

            End Select


            sapHucre.Font.Color = sapRenk
            sapHucre.Font.bold = True


            If i = 0 Then

                Set kartAlani = ws.Range("B7:E13")

            Else

                Set kartAlani = ws.Range("G7:J13")

            End If


            For Each shp In ws.Shapes

                If shp.Type = msoTextBox _
                   Or shp.Type = msoAutoShape Then


                    If shp.Left >= kartAlani.Left _
                       And shp.Left < _
                           kartAlani.Left + kartAlani.Width _
                       And shp.Top >= kartAlani.Top _
                       And shp.Top < _
                           kartAlani.Top + kartAlani.Height Then


                        If shp.TextFrame2.HasText Then

                            If UCase$( _
                                Trim$( _
                                    shp.TextFrame2.TextRange.Text _
                                ) _
                            ) = sapMetni Then


                                If sapMetni = "BAÐLI" _
                                   Or sapMetni = "BAÐLI DEÐÝL" Then

                                    shp.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                                        sapRenk

                                End If

                            End If

                        End If

                    End If

                End If

            Next shp

        End If

    Next i


    '======================================================
    ' IL DEGISTIR BUTONU
    '======================================================
    Set shp = Nothing

    On Error Resume Next
    Set shp = ws.Shapes("AKSORA15_IL_DEGISTIR")
    On Error GoTo Hata


    If Not shp Is Nothing Then

        With shp

            .Visible = msoTrue

            .Fill.Visible = msoTrue
            .Fill.Solid
            .Fill.ForeColor.RGB = RGB(0, 105, 82)
            .Fill.Transparency = 0

            .line.Visible = msoTrue
            .line.ForeColor.RGB = RGB(34, 230, 147)
            .line.Weight = 1.5

            .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
                RGB(255, 255, 255)

            .ZOrder msoBringToFront

        End With

    End If


Cikis:

    On Error Resume Next


    If korumali Then

        ws.Protect _
            Password:="", _
            DrawingObjects:=True, _
            Contents:=icerik, _
            Scenarios:=senaryo, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=filtre, _
            AllowSorting:=sirala, _
            AllowUsingPivotTables:=pivot, _
            AllowFormattingCells:=bicim

    End If


    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran


    If Len(hataMesaji) > 0 Then

        MsgBox _
            hataMesaji, _
            vbExclamation, _
            "AKSORA"

    End If


    Exit Sub


Hata:

    hataMesaji = _
        "Durum renkleri (" & _
        Err.Number & _
        "): " & _
        Err.Description

    Resume Cikis

End Sub


'==========================================================
' SAYFAYI GUVENLI BUL
'==========================================================
Private Function PANEL_SAYFA_BUL( _
    ByVal sayfaAdi As String) As Worksheet

    Dim ws As Worksheet


    For Each ws In ThisWorkbook.Worksheets

        If StrComp( _
            ws.Name, _
            sayfaAdi, _
            vbTextCompare _
        ) = 0 Then

            Set PANEL_SAYFA_BUL = ws
            Exit Function

        End If

    Next ws

End Function

'==========================================================
' ANA PANEL BUTON BAGLANTILARINI GUNCELLE
' Eski OnAction adlari dosyada kalmis olsa bile mevcut
' prosedurlere yeniden baglar.
'==========================================================
Public Sub PANEL_BUTON_BAGLANTILARINI_YENILE()

    Dim ws As Worksheet
    Dim shp As Shape
    Dim metin As String
    Dim ad As String
    Dim kitap As String

    On Error GoTo Cikis

    Set ws = ThisWorkbook.Worksheets(PANEL_SAYFA)
    kitap = Replace(ThisWorkbook.Name, "'", "''")

    For Each shp In ws.Shapes

        ad = UCase$(Trim$(shp.Name))
        metin = vbNullString

        On Error Resume Next
        If shp.TextFrame2.HasText Then
            metin = UCase$(Trim$(shp.TextFrame2.TextRange.Text))
        End If
        On Error GoTo Cikis

        Select Case ad

            Case "GOSTOK"
                shp.OnAction = "'" & kitap & "'!PANEL_GIT_STOKLAR"

            Case "GORAPOR"
                shp.OnAction = "'" & kitap & "'!PANEL_GIT_RAPORLAR"

            Case "GOCDP"
                shp.OnAction = "'" & kitap & "'!PANEL_GIT_CDP"

            Case "GOEPP"
                shp.OnAction = "'" & kitap & "'!PANEL_GIT_EPP"

            Case "GOILCE"
                shp.OnAction = "'" & kitap & "'!PANEL_GIT_ILCE"

            Case "GOYUK"
                shp.OnAction = "'" & kitap & "'!PANEL_GIT_YUKLENICI"

            Case "AKSORA15_IL_DEGISTIR"
                shp.OnAction = "'" & kitap & "'!AKSORA_ANA_PANEL_IL_DEGISTIR"

        End Select

        ' Ýsimleri sürümden sürüme deðiþebilen büyük butonlarý
        ' görünen metin üzerinden de güvenli biçimde baðla.
        If InStr(1, metin, "MALZEME ARAMA", vbTextCompare) > 0 Then

            shp.OnAction = "'" & kitap & "'!PANEL_GIT_MALZEME_ARA"

        ElseIf InStr(1, metin, "STOK KARÞILAÞTIR", vbTextCompare) > 0 _
            Or InStr(1, metin, "STOK KARSILASTIR", vbTextCompare) > 0 Then

            shp.OnAction = "'" & kitap & "'!PANEL_GIT_STOK_KARSILASTIRMA"

        ElseIf InStr(1, metin, "ÝL DEÐÝÞTÝR", vbTextCompare) > 0 _
            Or InStr(1, metin, "IL DEGISTIR", vbTextCompare) > 0 Then

            shp.OnAction = "'" & kitap & "'!AKSORA_ANA_PANEL_IL_DEGISTIR"

        End If

    Next shp

Cikis:

End Sub

' Stok sayfasindaki SAP guncelleme dugmeleri: o sayfanin stogu hizli yolla guncellenir,
' is bitince ayni stok sayfasina donulur ve tek ozet gosterilir.
Public Sub AKSORA_STOK_SAYFA_GUNCELLE()
    Dim ws As Worksheet, ad As String, bas As Double, sessiz As Boolean, no As Long, metin As String
    Dim ev As Boolean, baslik As String
    If AKSORA_V11_MESGUL Then Exit Sub
    Set ws = ActiveSheet
    ad = AKFIX_NORMAL(ws.Name)
    sessiz = ANA_SESSIZ_MOD
    On Error GoTo Hata
    AKHZ_TEMIZLE
    bas = Timer
    ANA_SESSIZ_MOD = True
    Select Case ad
        Case "CDP_STOK": baslik = "CDP stok": CDP_STOK_GUNCELLE_V2
        Case "EPP_STOK": baslik = "EPP stok": EPP_STOK_GUNCELLE_V2
        Case "STOKLAR": baslik = "Normal stok (CDP + EPP)": ANA_NORMAL_STOK_GUNCELLE_V2
        Case "ILCE_STOK": baslik = "Ýlçe stok": ILCE_STOK_GUNCELLE_V2
        Case "YUKLENICI_STOK": baslik = "Yüklenici stok": YUKLENICI_STOK_GUNCELLE_V2
        Case Else: ANA_SESSIZ_MOD = sessiz: Exit Sub
    End Select
    ANA_SESSIZ_MOD = sessiz
    AKSORA_PANEL_SURE_KAYDET bas
    On Error Resume Next
    PANEL_YENILE
    ev = application.EnableEvents
    application.EnableEvents = False
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    application.EnableEvents = ev
    AKSORA_V13_GORUNUM ws
    application.ScreenUpdating = True
    On Error GoTo 0
    MsgBox baslik & " güncellemesi tamamlandý." & vbCrLf & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(AKSORA_PANEL_SON_SURE(), "0.00") & " sn", vbInformation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    ANA_SESSIZ_MOD = sessiz
    On Error Resume Next
    application.EnableEvents = True
    ws.Select Replace:=True
    AKSORA_V13_GORUNUM ws
    application.ScreenUpdating = True
    On Error GoTo 0
    AKSORA_HATA_GOSTER "Stok güncelleme", no, metin
End Sub
