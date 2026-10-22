Attribute VB_Name = "modAKSORA_GORSEL"
Option Explicit


Public Sub AKSORA_SPLASH_GORSEL_UYGULA(ByVal frm As Object)

    Dim logoYolu As String
    Dim merkez As Double
    Dim kontrol As Variant

    On Error GoTo Hata

    With frm
        .Caption = "AKSORA | Stok ve Verimlilik"
        .Width = 440
        .Height = 280
        .BackColor = RGB(248, 252, 255)
    End With

    merkez = frm.InsideWidth / 2

    For Each kontrol In Array( _
        "lblBaslik", "lblAltBaslik", "lblDurum", "lblImza")

        With frm.Controls(CStr(kontrol))
            .AutoSize = False
            .WordWrap = False
            .BackStyle = fmBackStyleTransparent
            .TextAlign = fmTextAlignCenter
            .Font.Name = "Segoe UI"
            .Width = 300
            .Left = merkez - 150
        End With

    Next kontrol

    With frm.imgLogo
        .Width = 64
        .Height = 64
        .Left = merkez - 32
        .Top = 15
        .BackColor = RGB(248, 252, 255)
        .BorderStyle = fmBorderStyleNone
        .PictureSizeMode = fmPictureSizeModeZoom
    End With

    With frm.lblBaslik
        .Caption = "AKSORA"
        .Top = 85
        .Height = 32
        .Font.Size = 23
        .Font.bold = True
        .ForeColor = RGB(15, 49, 90)
    End With

        With frm.lblAltBaslik
        .Caption = "Stok ve Verimlilik, AKSORA!"
        .Top = 119
        .Height = 24
        .Font.Name = "Segoe UI"
        .Font.Size = 12
        .Font.bold = True
        .Font.Italic = True
        .ForeColor = RGB(15, 49, 90)
    End With

    With frm.lblDurum
        .Caption = "Sistem hazýrlanýyor..."
        .Top = 155
        .Height = 20
        .Font.Size = 9
        .Font.bold = False
        .ForeColor = RGB(15, 49, 90)
    End With

    With frm.lblBarArka
        .AutoSize = False
        .Caption = vbNullString
        .Left = merkez - 140
        .Top = 182
        .Width = 280
        .Height = 7
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(220, 233, 230)
        .BorderStyle = fmBorderStyleNone
    End With

    With frm.lblBar
        .AutoSize = False
        .Caption = vbNullString
        .Left = frm.lblBarArka.Left
        .Top = frm.lblBarArka.Top
        .Width = 1
        .Height = frm.lblBarArka.Height
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(0, 110, 70)
        .BorderStyle = fmBorderStyleNone
        .ZOrder 0
    End With

    With frm.lblImza
        .Caption = AKSORA_TELIF()
        .Top = 207
        .Height = 18
        .Font.Size = 8
        .Font.bold = False
        .ForeColor = RGB(102, 120, 138)
    End With

    logoYolu = AKSORA_GORSEL("AKSORA_SPLASH_LOGO.bmp")

    If Len(Dir$(logoYolu)) > 0 Then
        Set frm.imgLogo.Picture = LoadPicture(logoYolu)
    End If

    Exit Sub

Hata:
    ' Bekleme karti olarak acilirken gorsel hatasi sessiz gecilir (kart yine de gorunur).
    If AKSORA_BEKLE_KIPI Then Exit Sub
    MsgBox "Yükleme ekraný: " & Err.Description, _
        vbExclamation, "AKSORA"

End Sub





'Ilgili rapor yenilendikten sonra cagrilabilir.
Public Sub RAPOR_GORUNUM_TEK(ByVal sayfaAdi As String)
    Dim ws As Worksheet
    Dim baslik As Long, toplam As Long, sonKolon As Long
    Dim r As Long, c As Long, metin As String
    Dim eskiOlay As Boolean, eskiEkran As Boolean

    eskiOlay = application.EnableEvents
    eskiEkran = application.ScreenUpdating
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(sayfaAdi)
    application.EnableEvents = False
    application.ScreenUpdating = False

    'Tablo basligini ilk 25 satirda ara.
    For r = 1 To 25
        For c = 1 To 5
            metin = LCase$(Trim$(CStr(ws.Cells(r, c).Value2)))
            If metin = "depo yeri" Or metin = "depo yeri tanýmý" _
               Or metin = "depo yeri tanimi" Then
                baslik = r
                Exit For
            End If
        Next c
        If baslik > 0 Then Exit For
    Next r
    If baslik = 0 Then GoTo Cikis

    'TOPLAM satirini bul; altindaki bos satirlara dokunma.
    For r = baslik + 1 To AKSORA_SON_SATIR(ws, 1)
        For c = 1 To 5
            metin = UCase$(Trim$(CStr(ws.Cells(r, c).Value2)))
            If metin = "TOPLAM" Then
                toplam = r
                Exit For
            End If
        Next c
        If toplam > 0 Then Exit For
    Next r
    If toplam <= baslik Then GoTo Cikis

    ws.Rows(baslik & ":" & toplam).RowHeight = 20
        Select Case sayfaAdi
        Case "ÝLÇE_RAPORLAR", "YÜKLENÝCÝ_RAPOR"
            sonKolon = 5
        Case Else
            sonKolon = 4
    End Select

    ws.Rows(baslik).RowHeight = 36

    With ws.Range(ws.Cells(baslik, 1), ws.Cells(baslik, sonKolon))
        .WrapText = True
        .Font.Size = 9
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    ' Bes raporun tum veri satirlarini ortak bicime getir.
    With ws.Range(ws.Cells(baslik + 1, 1), ws.Cells(toplam, sonKolon))
        .Font.Name = "Aptos"
        .Font.Size = 10
        .ShrinkToFit = False
        .IndentLevel = 0
        .VerticalAlignment = xlCenter
        .WrapText = False
    End With
    With ws.Range(ws.Cells(baslik + 1, 1), ws.Cells(toplam - 1, 1))
        .HorizontalAlignment = xlCenter
    End With
    With ws.Range(ws.Cells(baslik + 1, 2), ws.Cells(toplam - 1, 2))
        .HorizontalAlignment = xlLeft
        .WrapText = True
    End With
    ws.Range(ws.Cells(baslik + 1, 3), ws.Cells(toplam, sonKolon)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(toplam, 1), ws.Cells(toplam, sonKolon)).Font.bold = True
    ws.Cells(toplam, 1).HorizontalAlignment = xlCenter
    ws.Rows(baslik + 1 & ":" & toplam - 1).AutoFit
    For r = baslik + 1 To toplam - 1
        If ws.Rows(r).RowHeight < 20 Then ws.Rows(r).RowHeight = 20
    Next r

    'Yalniz EPP raporunda depo koduna gore butun tabloyu sirala.
    If sayfaAdi = "EPP_RAPOR" And toplam > baslik + 2 Then
        sonKolon = ws.Cells(baslik, ws.Columns.Count).End(xlToLeft).Column
        If sonKolon < 4 Then sonKolon = 4
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Range(ws.Cells(baslik + 1, 1), _
                            ws.Cells(toplam - 1, 1)), _
                            SortOn:=xlSortOnValues, Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .SetRange ws.Range(ws.Cells(baslik + 1, 1), _
                               ws.Cells(toplam - 1, sonKolon))
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If
    AKSORA_RAPOR_KARTLARI_YENILE sayfaAdi
    AKSORA_TASARIM_SAYFA ws, True
    AKSORA_RAPOR_GRAFIK_YENILE sayfaAdi
    AKSORA_RAPOR_ALT_YAZI_YERLESTIR sayfaAdi
    AKSORA_RAPOR_TABLO_BITISI sayfaAdi
    AKSORA_RAPOR_AMBLEM_HIZALA sayfaAdi
    AKSORA_V11_BICIM_SAYFA ws
    AKSORA_TASARIM_SAYFA ws, True
    If ActiveSheet Is ws And Not AKSORA_PENCERE_KUCUK(ActiveWindow) Then
        application.CutCopyMode = False
        ws.Range("A1").MergeArea.Cells(1, 1).Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
Cikis:
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    Exit Sub
Hata:
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    Err.Raise Err.Number, "RAPOR_GORUNUM_TEK", Err.Description
End Sub



Public Sub STOK_GORUNUM_TEK(ByVal sayfaAdi As String)
    Dim ws As Worksheet
    Dim sonSatir As Long
    Dim eskiEkran As Boolean

    eskiEkran = application.ScreenUpdating
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(sayfaAdi)
    application.ScreenUpdating = False

    ' Stok tablolarinda baslik 6. satirda, veriler 7. satirdan baslar.
    ' G sutunundaki malzeme kodunu esas alarak son veri satirini bul.
    sonSatir = AKSORA_SON_SATIR(ws, "G")
    If sonSatir < 6 Then sonSatir = 6

    ws.Rows("6:" & sonSatir).RowHeight = 20

    AKSORA_V11_BICIM_SAYFA ws
    AKSORA_TASARIM_SAYFA ws, True
    If ActiveSheet Is ws And Not AKSORA_PENCERE_KUCUK(ActiveWindow) Then
        application.CutCopyMode = False
        ws.Range("A1").MergeArea.Cells(1, 1).Select
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If
Cikis:
    application.ScreenUpdating = eskiEkran
    Exit Sub

Hata:
    application.ScreenUpdating = eskiEkran
    Err.Raise Err.Number, "STOK_GORUNUM_TEK", Err.Description
End Sub


