Attribute VB_Name = "modSTOK_KARSILASTIRMA"
Option Explicit

Private Const KARS_SAYFA As String = "STOK_KARŞILAŞTIRMA"
Private Const KARS_NN As String = "NEREDEN / NEREYE"


Public Sub STOK_KARSILASTIRMA_KUR()

    Dim ws As Worksheet

    On Error GoTo Hata

    application.ScreenUpdating = False
    application.EnableEvents = False

    Set ws = STOK_KARS_SAYFA_AL()

    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    STOK_KARS_TEMIZLE ws
    STOK_KARS_ZEMIN ws
    STOK_KARS_BASLIK ws
    STOK_KARS_SECIM_ALANI ws
    STOK_KARS_OZET_ALANI ws
    STOK_KARS_HIZLI_FILTRELER ws
    STOK_KARS_TABLO_BASLIK ws
    STOK_KARS_ANA_PANEL_BUTONU ws
    STOK_KARS_TARIHLERI_YUKLE

    ws.Activate
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    ws.Range("A1").Select

Cikis:
    On Error Resume Next
    application.EnableEvents = True
    application.ScreenUpdating = True
    Exit Sub

Hata:
    MsgBox _
        "STOK KARŞILAŞTIRMA sayfası oluşturulamadı." & vbCrLf & vbCrLf & _
        "Hata No: " & Err.Number & vbCrLf & _
        "Açıklama: " & Err.Description, _
        vbCritical, _
        "AKSORA"
    Resume Cikis

End Sub

Public Sub STOK_KARSILASTIR()

    Dim ws As Worksheet
    Dim wsIndex As Worksheet
    Dim wsDelta As Worksheet

    Dim Snap1 As String
    Dim Snap2 As String

    Dim tarih1 As Date
    Dim tarih2 As Date

    Dim d1 As Object
    Dim d2 As Object
    Dim tum As Object
    Dim Kimlik1 As Object
    Dim Kimlik2 As Object

    Dim anahtar As Variant
    Dim r1 As Variant
    Dim r2 As Variant

    Dim CikisSatir As Long

    Dim azalan As Long
    Dim artan As Long
    Dim yeni As Long
    Dim sifirlanan As Long

    Dim EskiToplam As Double
    Dim YeniToplam As Double

    Dim durum As String
    Dim oncekiEkran As Boolean
    oncekiEkran = application.ScreenUpdating

    On Error GoTo Hata

    Set ws = ThisWorkbook.Worksheets(KARS_SAYFA)
    STOK_KARS_NN_HAZIRLA ws
    STOK_KARS_UST_GENISLET ws

    Snap1 = STOK_KARS_SECILEN_SNAPSHOT(ws.Range("B6").Value)
    Snap2 = STOK_KARS_SECILEN_SNAPSHOT(ws.Range("I6").Value)

    If Len(Snap1) = 0 Or Len(Snap2) = 0 Then
        MsgBox _
            "İki karşılaştırma tarihini de seçiniz.", _
            vbExclamation, _
            "STOK KARŞILAŞTIRMA"
        Exit Sub
    End If

    tarih1 = STOK_KARS_SNAPSHOT_TARIH(Snap1)
    tarih2 = STOK_KARS_SNAPSHOT_TARIH(Snap2)

    If tarih1 >= tarih2 Then
        MsgBox _
            "1. tarih, 2. tarihten eski olmalıdır.", _
            vbExclamation, _
            "STOK KARŞILAŞTIRMA"
        Exit Sub
    End If

    application.ScreenUpdating = False

    AKFIX_KARS_GECMIS_FARK = 0
    Set d1 = STOK_KARS_DURUM_OLUSTUR(Snap1)
    Set d2 = STOK_KARS_DURUM_OLUSTUR(Snap2)

    If d1 Is Nothing Or d2 Is Nothing Then
        Err.Raise _
            vbObjectError + 7301, _
            "STOK_KARSILASTIR", _
            "Seçilen stok durumları oluşturulamadı."
    End If

    Set Kimlik1 = STOK_KARS_KIMLIK_SETI(d1)
    Set Kimlik2 = STOK_KARS_KIMLIK_SETI(d2)

    Set tum = CreateObject("Scripting.Dictionary")
    tum.CompareMode = vbTextCompare

    For Each anahtar In d1.keys
        If Not tum.Exists(CStr(anahtar)) Then
            tum.Add CStr(anahtar), True
        End If
    Next anahtar

    For Each anahtar In d2.keys
        If Not tum.Exists(CStr(anahtar)) Then
            tum.Add CStr(anahtar), True
        End If
    Next anahtar

    ws.Range("A16:N" & STOK_KARS_ICERIK_SON(ws)).ClearContents

    CikisSatir = 16

    For Each anahtar In tum.keys

        r1 = Empty
        r2 = Empty
        durum = vbNullString

        If d1.Exists(CStr(anahtar)) Then
            r1 = d1(CStr(anahtar))
        End If

        If d2.Exists(CStr(anahtar)) Then
            r2 = d2(CStr(anahtar))
        End If

        If Not d1.Exists(CStr(anahtar)) And _
           d2.Exists(CStr(anahtar)) Then

            If Kimlik1.Exists(STOK_KARS_KIMLIK(r2)) Then
                durum = "ARTTI"
                artan = artan + 1
            Else
                durum = "YENİ"
                yeni = yeni + 1
            End If

        ElseIf d1.Exists(CStr(anahtar)) And _
               Not d2.Exists(CStr(anahtar)) Then

            If Kimlik2.Exists(STOK_KARS_KIMLIK(r1)) Then
                durum = "AZALDI"
                azalan = azalan + 1
            Else
                durum = "SIFIRLANDI"
                sifirlanan = sifirlanan + 1
            End If

        ElseIf STOK_KARS_DOUBLE(r2(9)) > _
               STOK_KARS_DOUBLE(r1(9)) Then

            durum = "ARTTI"
            artan = artan + 1

        ElseIf STOK_KARS_DOUBLE(r2(9)) < _
               STOK_KARS_DOUBLE(r1(9)) Then

            durum = "AZALDI"
            azalan = azalan + 1

        ElseIf CCur(STOK_KARS_DOUBLE(r2(11))) <> _
               CCur(STOK_KARS_DOUBLE(r1(11))) Then

            durum = "DEĞER DEĞİŞTİ"

        End If

        If Len(durum) > 0 Then

            STOK_KARS_SATIR_YAZ _
                ws, _
                CikisSatir, _
                durum, _
                r1, _
                r2

            CikisSatir = CikisSatir + 1

        End If

    Next anahtar

    EskiToplam = STOK_KARS_SNAPSHOT_TOPLAM(Snap1)
    YeniToplam = STOK_KARS_SNAPSHOT_TOPLAM(Snap2)

    ws.Range("A10").Value = azalan
ws.Range("D10").Value = artan
ws.Range("G10").Value = yeni
ws.Range("J10").Value = sifirlanan

ws.Range("A12").Value = EskiToplam
ws.Range("D12").Value = YeniToplam
ws.Range("G12").Value = CDbl(CDec(YeniToplam) - CDec(EskiToplam))
ws.Range("J12").Value = application.Max(0, CikisSatir - 16)

ws.Range("A12,D12,G12").NumberFormat = "#,##0.00 ""TL"""

    If CikisSatir > 16 Then
        STOK_KARS_NEREDEN ws, CikisSatir - 1, tarih1, tarih2
        STOK_KARS_TABLO_FORMAT ws, CikisSatir - 1
    End If

        ws.Range("B7").Value = _
        "Son karşılaştırılan: " & Format$(tarih1, "dd.mm.yyyy hh:nn:ss")

    ws.Range("I7").Value = _
        "Son karşılaştırılan: " & Format$(tarih2, "dd.mm.yyyy hh:nn:ss")

    With ws.Range("B7")
        .Font.Size = 8
        .WrapText = False
    End With

    With ws.Range("I7")
        .Font.Size = 8
        .WrapText = False
    End With

    STOK_KARS_FILTRE_TUMU
    AKSORA_V11_BICIM_SAYFA ws
    AKSORA_V11_KARS_UYARI ws

Cikis:
    On Error Resume Next
    application.ScreenUpdating = oncekiEkran
    Exit Sub

Hata:
    MsgBox _
        "Karşılaştırma yapılamadı." & vbCrLf & vbCrLf & _
        "Hata No: " & Err.Number & vbCrLf & _
        "Açıklama: " & Err.Description, _
        vbCritical, _
        "AKSORA - STOK KARŞILAŞTIRMA"

    Resume Cikis

End Sub

Public Sub STOK_KARS_TARIHLERI_YUKLE()
    AKFIX_KARS_TARIHLER
End Sub

Public Sub STOK_KARS_FILTRE_TUMU()
    On Error GoTo Hata
    STOK_KARS_FILTRE vbNullString
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Filtre", Err.Number, Err.Description
End Sub

Public Sub STOK_KARS_FILTRE_AZALAN()
    On Error GoTo Hata
    STOK_KARS_FILTRE "AZALDI"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Filtre", Err.Number, Err.Description
End Sub

Public Sub STOK_KARS_FILTRE_ARTAN()
    On Error GoTo Hata
    STOK_KARS_FILTRE "ARTTI"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Filtre", Err.Number, Err.Description
End Sub

Public Sub STOK_KARS_FILTRE_YENI()
    On Error GoTo Hata
    STOK_KARS_FILTRE "YENİ"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Filtre", Err.Number, Err.Description
End Sub

Public Sub STOK_KARS_FILTRE_SIFIR()
    On Error GoTo Hata
    STOK_KARS_FILTRE "SIFIRLANDI"
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Filtre", Err.Number, Err.Description
End Sub

Public Sub STOK_KARS_ANA_PANEL()
    On Error GoTo Hata
    AKSORA_V12_GIT "ANA PANEL"
    Exit Sub
Hata:
    MsgBox "Ana panel açılamadı: " & Err.Description, vbExclamation, "AKSORA"
End Sub
Public Sub STOK_KARS_KPI_GUNCELLE()

    Dim ws As Worksheet
    Dim sonSatir As Long, r As Long
    Dim azalan As Long, artan As Long
    Dim yeni As Long, sifirlanan As Long
    Dim toplamDegisim As Long
    Dim durum As String

    Set ws = ThisWorkbook.Worksheets(KARS_SAYFA)
    sonSatir = AKSORA_SON_SATIR(ws, "D")

    ' Gorunur satirlar tek seferde (satir satir okuma on binlerce satirda Excel'i kilitliyordu).
    Dim gorV As Variant, aV As Variant
    If sonSatir >= 16 Then
        gorV = AKSORA_GORUNUR(ws, 16, sonSatir, 14)
        aV = ws.Range("A16:A" & sonSatir).Value2
        If Not IsArray(aV) Then ReDim aV(1 To 1, 1 To 1): aV(1, 1) = ws.Range("A16").Value2
    End If
    For r = 16 To sonSatir
        If gorV(r - 15) = 1 Then

            durum = UCase$(Trim$(CStr(aV(r - 15, 1))))
            toplamDegisim = toplamDegisim + 1
            durum = Replace(durum, "İ", "I")

            Select Case durum
                Case "AZALDI"
                    azalan = azalan + 1
                Case "ARTTI"
                    artan = artan + 1
                Case "YENI"
                    yeni = yeni + 1
                Case "SIFIRLANDI"
                    sifirlanan = sifirlanan + 1
            End Select

        End If
    Next r

    ' Deger degisen satirlar da toplam degisim sayisina dahildir.

    ws.Range("A10").Value2 = azalan
    ws.Range("D10").Value2 = artan
    ws.Range("G10").Value2 = yeni
    ws.Range("J10").Value2 = sifirlanan
    ws.Range("J12").Value2 = toplamDegisim

    If Len(Trim$(CStr(ws.Range("B6").Value2))) > 0 And _
       Len(Trim$(CStr(ws.Range("I6").Value2))) > 0 Then
        AKSKV2_TOPLAM_GUNCELLE ws
    End If

End Sub

Private Sub STOK_KARS_FILTRE(ByVal durum As String)

    Dim ws As Worksheet
    Dim sonSatir As Long
    Dim Eslesen As Long

    Set ws = ThisWorkbook.Worksheets(KARS_SAYFA)

    On Error Resume Next

    If ws.FilterMode Then
        ws.ShowAllData
    End If

    On Error GoTo 0

    sonSatir = _
        AKSORA_SON_SATIR(ws, "D")

    If sonSatir < 16 Then
        AKFIX_EN_USTE ws
        Exit Sub
    End If

    If Len(durum) > 0 Then

        Eslesen = _
            application.WorksheetFunction.CountIf( _
                ws.Range("A16:A" & sonSatir), _
                durum)

        If Eslesen = 0 Then
            ws.Range("A15:N" & sonSatir).AutoFilter Field:=1, Criteria1:=durum
            STOK_KARS_KPI_GUNCELLE
    AKSORA_V11_BICIM_SAYFA ws
            AKFIX_EN_USTE ws
            Exit Sub
        End If

        ws.Range( _
            "A15:N" & sonSatir).AutoFilter _
            Field:=1, _
            Criteria1:=durum

    End If

    STOK_KARS_KPI_GUNCELLE
    AKFIX_EN_USTE ws

End Sub

Private Function STOK_KARS_DURUM_OLUSTUR(ByVal HedefSnapshot As String) As Object
    Set STOK_KARS_DURUM_OLUSTUR = AKFIX_KARS_DURUM(HedefSnapshot)
End Function

Private Function STOK_KARS_KIMLIK_SETI(ByVal d As Object) As Object

    Dim sonuc As Object
    Dim anahtar As Variant
    Dim kayit As Variant
    Dim Kimlik As String

    Set sonuc = CreateObject("Scripting.Dictionary")
    sonuc.CompareMode = vbTextCompare

    For Each anahtar In d.keys
        kayit = d(CStr(anahtar))
        Kimlik = STOK_KARS_KIMLIK(kayit)

        If Len(Kimlik) > 0 Then
            If Not sonuc.Exists(Kimlik) Then
                sonuc.Add Kimlik, True
            End If
        End If
    Next anahtar

    Set STOK_KARS_KIMLIK_SETI = sonuc

End Function

Private Function STOK_KARS_KIMLIK(ByVal kayit As Variant) As String

    If IsEmpty(kayit) Then Exit Function

    STOK_KARS_KIMLIK = _
        Trim$(CStr(kayit(1))) & "|" & _
        Trim$(CStr(kayit(7)))

End Function

Private Sub STOK_KARS_SATIR_YAZ( _
ByVal ws As Worksheet, _
ByVal satir As Long, _
ByVal durum As String, _
ByVal eski As Variant, _
ByVal yeni As Variant)

    Dim kayit As Variant

    Dim EskiMiktar As Double
    Dim YeniMiktar As Double

    Dim EskiDeger As Double
    Dim YeniDeger As Double

    If IsEmpty(yeni) Then
        kayit = eski
    Else
        kayit = yeni
    End If

    If Not IsEmpty(eski) Then

        EskiMiktar = _
            STOK_KARS_DOUBLE(eski(9))

        EskiDeger = _
            STOK_KARS_DOUBLE(eski(11))

    End If

    If Not IsEmpty(yeni) Then

        YeniMiktar = _
            STOK_KARS_DOUBLE(yeni(9))

        YeniDeger = _
            STOK_KARS_DOUBLE(yeni(11))

    End If

    ws.Cells(satir, 1).Value = durum
    ws.Cells(satir, 2).Value = kayit(1)
    ws.Cells(satir, 3).Value = kayit(2)
    ws.Cells(satir, 4).Value = kayit(7)
    ws.Cells(satir, 5).Value = kayit(8)
    ws.Cells(satir, 6).Value = kayit(6)

    ws.Cells(satir, 7).Value = EskiMiktar
    ws.Cells(satir, 8).Value = YeniMiktar
    ws.Cells(satir, 9).Value = YeniMiktar - EskiMiktar

    ws.Cells(satir, 10).Value = kayit(10)

    ws.Cells(satir, 11).Value = EskiDeger
    ws.Cells(satir, 12).Value = YeniDeger
    ws.Cells(satir, 13).Value = YeniDeger - EskiDeger

End Sub

Private Function STOK_KARS_SECILEN_SNAPSHOT(ByVal GorunenTarih As Variant) As String
    Dim w As Worksheet, r As Long, son As Long
    Set w = ThisWorkbook.Worksheets(KARS_SAYFA)
    If Not IsDate(GorunenTarih) Then Exit Function
    son = AKSORA_SON_SATIR(w, "X")
    For r = 2 To son
        If IsDate(w.Cells(r, "X").Value) Then
            If Abs(CDbl(CDate(w.Cells(r, "X").Value)) - CDbl(CDate(GorunenTarih))) < 0.00000001 Then
                STOK_KARS_SECILEN_SNAPSHOT = CStr(w.Cells(r, "Y").Value2)
                Exit Function
            End If
        End If
    Next r
End Function

Private Function STOK_KARS_SNAPSHOT_TARIH(ByVal SnapshotID As String) As Date
    STOK_KARS_SNAPSHOT_TARIH = AKFIX_KARS_TARIH(SnapshotID)
End Function

Private Function STOK_KARS_SNAPSHOT_TOPLAM(ByVal SnapshotID As String) As Double
    STOK_KARS_SNAPSHOT_TOPLAM = AKFIX_KARS_TOPLAM(SnapshotID)
End Function

Private Function STOK_KARS_SAYFA_AL() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(KARS_SAYFA)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                After:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = KARS_SAYFA

    End If

    Set STOK_KARS_SAYFA_AL = ws

End Function

Private Sub STOK_KARS_TEMIZLE(ByVal ws As Worksheet)

    Dim i As Long

    ws.Cells.UnMerge
    ws.Cells.Clear

    On Error Resume Next

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i

    On Error GoTo 0

End Sub

Private Sub STOK_KARS_ZEMIN(ByVal ws As Worksheet)

    With ws

        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 10

        .Range("A1:N2000").Interior.Color = _
            RGB(245, 247, 250)

        .Columns("A").ColumnWidth = 14
        .Columns("B").ColumnWidth = 10
        .Columns("C").ColumnWidth = 11
        .Columns("D").ColumnWidth = 16
        .Columns("E").ColumnWidth = 38
        .Columns("F").ColumnWidth = 14
        .Columns("G:I").ColumnWidth = 13
        .Columns("J").ColumnWidth = 10
        .Columns("K:M").ColumnWidth = 16
        .Columns("N").ColumnWidth = 44

        .Rows("1:3").RowHeight = 25
        .Rows("4:13").RowHeight = 23

        .Rows("14").RowHeight = 30
        .Rows("15").RowHeight = 30

    End With

End Sub

Private Sub STOK_KARS_BASLIK(ByVal ws As Worksheet)

    With ws.Range("A1:N2")

        .Merge
        .Value = "STOK KARŞILAŞTIRMA MERKEZİ"

        .Interior.Color = RGB(20, 55, 95)

        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .Font.Size = 20

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    With ws.Range("A3:N3")

        .Merge

        .Value = _
            "CDP 7155 + EPP 7055  |  Normal Stok Hafızası"

        .Interior.Color = RGB(20, 55, 95)

        .Font.Color = RGB(215, 225, 238)

        .HorizontalAlignment = xlCenter

    End With

End Sub

Private Sub STOK_KARS_SECIM_ALANI(ByVal ws As Worksheet)

    With ws.Range("A5:F5")

        .Merge
        .Value = "1. KARŞILAŞTIRMA TARİHİ"

        .Font.bold = True
        .Font.Color = RGB(75, 85, 95)

        .HorizontalAlignment = xlCenter

    End With

    With ws.Range("H5:N5")

        .Merge
        .Value = "2. KARŞILAŞTIRMA TARİHİ"

        .Font.bold = True
        .Font.Color = RGB(75, 85, 95)

        .HorizontalAlignment = xlCenter

    End With

    With ws.Range("B6:F6")

        .Merge
        .Name = "KARS_TARIH_1"

        .Interior.Color = RGB(255, 255, 255)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(185, 195, 205)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    With ws.Range("I6:N6")

        .Merge
        .Name = "KARS_TARIH_2"

        .Interior.Color = RGB(255, 255, 255)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(185, 195, 205)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    With ws.Range("B7:F7")

        .Merge

        .Font.Size = 8
        .Font.Color = RGB(125, 135, 145)

        .HorizontalAlignment = xlCenter

    End With

    With ws.Range("I7:N7")

        .Merge

        .Font.Size = 8
        .Font.Color = RGB(125, 135, 145)

        .HorizontalAlignment = xlCenter

    End With

End Sub

Private Sub STOK_KARS_OZET_ALANI(ByVal ws As Worksheet)

    STOK_KARS_KPI _
        ws, _
        "A9:C9", _
        "A10:C10", _
        "AZALAN", _
        "B10"

    STOK_KARS_KPI _
        ws, _
        "D9:F9", _
        "D10:F10", _
        "ARTAN", _
        "E10"

    STOK_KARS_KPI _
        ws, _
        "G9:I9", _
        "G10:I10", _
        "YENİ", _
        "H10"

    STOK_KARS_KPI _
        ws, _
        "J9:N9", _
        "J10:N10", _
        "SIFIRLANAN", _
        "K10"

    STOK_KARS_KPI _
        ws, _
        "A11:C11", _
        "A12:C12", _
        "ESKİ STOK DEĞERİ", _
        "B12"

    STOK_KARS_KPI _
        ws, _
        "D11:F11", _
        "D12:F12", _
        "YENİ STOK DEĞERİ", _
        "E12"

    STOK_KARS_KPI _
        ws, _
        "G11:I11", _
        "G12:I12", _
        "DEĞER FARKI", _
        "H12"

    STOK_KARS_KPI _
        ws, _
        "J11:N11", _
        "J12:N12", _
        "TOPLAM DEĞİŞİM", _
        "K12"

End Sub

Private Sub STOK_KARS_KPI( _
ByVal ws As Worksheet, _
ByVal BaslikAlan As String, _
ByVal DegerAlan As String, _
ByVal baslik As String, _
ByVal AdHucre As String)

    With ws.Range(BaslikAlan)

        .Merge
        .Value = baslik

        .Interior.Color = RGB(225, 230, 236)

        .Font.bold = True
        .Font.Size = 8
        .Font.Color = RGB(75, 85, 95)

        .HorizontalAlignment = xlCenter

    End With

    With ws.Range(DegerAlan)

        .Merge
        .Value = 0

        .Interior.Color = RGB(255, 255, 255)

        .Font.bold = True
        .Font.Size = 12
        .Font.Color = RGB(20, 55, 95)

        .HorizontalAlignment = xlCenter

    End With

End Sub

Private Sub STOK_KARS_HIZLI_FILTRELER( _
ByVal ws As Worksheet)

    STOK_KARS_HIZLI_BUTON _
        ws, 1, 6, _
        "TÜMÜ", _
        "STOK_KARS_FILTRE_TUMU", _
        "btnKarsTum"

    STOK_KARS_HIZLI_BUTON _
        ws, 2, 6, _
        "AZALAN", _
        "STOK_KARS_FILTRE_AZALAN", _
        "btnKarsAzalan"

    STOK_KARS_HIZLI_BUTON _
        ws, 3, 6, _
        "ARTAN", _
        "STOK_KARS_FILTRE_ARTAN", _
        "btnKarsArtan"

    STOK_KARS_HIZLI_BUTON _
        ws, 4, 6, _
        "YENİ", _
        "STOK_KARS_FILTRE_YENI", _
        "btnKarsYeni"

    STOK_KARS_HIZLI_BUTON _
        ws, 5, 6, _
        "SIFIRLANAN", _
        "STOK_KARS_FILTRE_SIFIR", _
        "btnKarsSifir"

    STOK_KARS_HIZLI_BUTON _
        ws, 6, 6, _
        "KARŞILAŞTIR", _
        "STOK_KARSILASTIR", _
        "btnKarsKarsilastir"

End Sub

Private Sub STOK_KARS_HIZLI_BUTON( _
ByVal ws As Worksheet, _
ByVal sira As Long, _
ByVal toplam As Long, _
ByVal metin As String, _
ByVal makro As String, _
ByVal ad As String)

    Dim alan As Range
    Dim shp As Shape

    Dim sol As Double
    Dim ust As Double

    Dim ToplamGenislik As Double
    Dim ButonGenislik As Double
    Dim ButonYukseklik As Double

    Set alan = ws.Range("A14:N14")

    ToplamGenislik = alan.Width
    ButonGenislik = ToplamGenislik / toplam
    ButonYukseklik = alan.Height

    sol = _
        alan.Left + _
        ((sira - 1) * ButonGenislik)

    ust = alan.top

    Set shp = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            sol, _
            ust, _
            ButonGenislik, _
            ButonYukseklik)

    With shp

        .Name = ad
        .TextFrame2.TextRange.Text = metin

        .Fill.ForeColor.RGB = RGB(45, 95, 145)

        .line.ForeColor.RGB = RGB(245, 247, 250)
        .line.Weight = 1

        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.bold = msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.MarginLeft = 1
        .TextFrame2.MarginRight = 1
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        .OnAction = _
            "'" & ThisWorkbook.Name & "'!" & makro

        .Placement = xlMoveAndSize

    End With

End Sub

Private Sub STOK_KARS_TABLO_BASLIK( _
ByVal ws As Worksheet)

    Dim basliklar As Variant
    Dim i As Long

    basliklar = Array( _
        "DURUM", _
        "ÜY", _
        "DEPO", _
        "MALZEME", _
        "MALZEME AÇIKLAMASI", _
        "PARTİ", _
        "ESKİ MİKTAR", _
        "YENİ MİKTAR", _
        "FARK", _
        "BİRİM", _
        "ESKİ DEĞER", _
        "YENİ DEĞER", _
        "DEĞER FARKI", _
        KARS_NN)

    For i = 0 To UBound(basliklar)

        ws.Cells(15, i + 1).Value = _
            basliklar(i)

    Next i

    With ws.Range("A15:N15")

        .Interior.Color = RGB(20, 55, 95)

        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .AutoFilter

    End With

End Sub

Private Sub STOK_KARS_ANA_PANEL_BUTONU( _
ByVal ws As Worksheet)

    Dim hedef As Range
    Dim shp As Shape

    Set hedef = ws.Range("O5")

    On Error Resume Next
    ws.Shapes("btnAnaPanelDon").Delete
    On Error GoTo 0

    Set shp = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            hedef.Left, _
            hedef.top, _
            hedef.Width, _
            hedef.Height)

    With shp

        .Name = "btnAnaPanelDon"
        .TextFrame2.TextRange.Text = "ANA PANEL"

        .Fill.ForeColor.RGB = RGB(20, 55, 95)
        .line.ForeColor.RGB = RGB(20, 55, 95)
        .line.Weight = 1

        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.bold = msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.MarginLeft = 1
        .TextFrame2.MarginRight = 1
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        .OnAction = _
            "'" & ThisWorkbook.Name & "'!ANA_PANELE_DON"

        .Placement = xlMoveAndSize

        .Left = hedef.Left
        .top = hedef.top
        .Width = hedef.Width
        .Height = hedef.Height

        .ZOrder msoBringToFront

    End With

End Sub

Private Sub STOK_KARS_TABLO_FORMAT( _
ByVal ws As Worksheet, _
ByVal sonSatir As Long)

    With ws.Range("A16:N" & sonSatir)

        .Interior.Color = RGB(255, 255, 255)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(225, 230, 236)
        .Borders.Weight = xlThin

        .VerticalAlignment = xlCenter

    End With

    ws.Range( _
        "G16:I" & sonSatir).NumberFormat = _
        "#,##0"

    ws.Range( _
        "K16:M" & sonSatir).NumberFormat = _
        "#,##0.00 ""TL"""

    ws.Range( _
        "A16:A" & sonSatir).Font.bold = True

    With ws.Range("N16:N" & sonSatir)
        .HorizontalAlignment = xlLeft
        .IndentLevel = 1
        .Font.Color = RGB(20, 55, 95)
        .ShrinkToFit = True
    End With

End Sub

Private Function STOK_KARS_DOUBLE( _
ByVal deger As Variant) As Double

    If IsNumeric(deger) Then

        STOK_KARS_DOUBLE = _
            CDbl(deger)

    Else

        STOK_KARS_DOUBLE = 0

    End If

End Function

Public Sub STOK_KARS_FILTRE_TAKIP_TICK()
    On Error Resume Next
    ThisWorkbook.Worksheets("STOK_KARŞILAŞTIRMA").STOK_KARS_FILTRE_TAKIP_BASLAT
    On Error GoTo 0
End Sub
Public Sub AKSORA_KARS_IL_GECIS_YENILE()
    Dim ws As Worksheet
    Dim alan As Variant, ad As Variant
    Dim son As Long
    Dim cizim As Boolean, icerik As Boolean, senaryo As Boolean
    Dim filtre As Boolean, sirala As Boolean
    Dim korumali As Boolean
    Dim hataNo As Long, hataMetni As String

    On Error GoTo Hata

    Set ws = ThisWorkbook.Worksheets("STOK_KARŞILAŞTIRMA")

    cizim = ws.ProtectDrawingObjects
    icerik = ws.ProtectContents
    senaryo = ws.ProtectScenarios
    filtre = ws.Protection.AllowFiltering
    sirala = ws.Protection.AllowSorting
    korumali = cizim Or icerik Or senaryo

    If korumali Then ws.Unprotect
        ws.Range("A3").Value = AKSORA_CANLI_STOK_IL() & _
        " | CDP " & AKSORA_IL_TESIS(AKSORA_CANLI_STOK_IL(), "CDP") & _
        " + EPP " & AKSORA_IL_TESIS(AKSORA_CANLI_STOK_IL(), "EPP") & _
        " | Normal Stok Hafızası"

    STOK_HAFIZA_ALTYAPI_KUR

    If ws.FilterMode Then ws.ShowAllData

    son = AKSORA_SON_SATIR(ws, "A")
    If son >= 16 Then ws.Range("A16:N" & son).ClearContents

    ws.Range("B6").MergeArea.ClearContents
    ws.Range("I6").MergeArea.ClearContents
    ws.Range("B7").MergeArea.ClearContents
    ws.Range("I7").MergeArea.ClearContents

    alan = Array("A10", "D10", "G10", "J10", _
                 "A12", "D12", "G12", "J12")

    For Each ad In alan
        ws.Range(CStr(ad)).Value = 0
    Next ad

    ' Ortak karsilastirmada bagimsiz il secimini koru.

    STOK_KARS_TARIHLERI_YUKLE

    If korumali Then
        ws.Protect Password:="", _
            DrawingObjects:=cizim, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True, _
            AllowFiltering:=filtre, AllowSorting:=sirala
    End If

    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description
    On Error Resume Next

    If korumali Then
        ws.Protect Password:="", _
            DrawingObjects:=cizim, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True, _
            AllowFiltering:=filtre, AllowSorting:=sirala
    End If

    On Error GoTo 0
    Err.Raise hataNo, "AKSORA_KARS_IL_GECIS_YENILE", hataMetni
End Sub



Private Function STOK_KARS_ICERIK_SON(ByVal ws As Worksheet) As Long
    Dim sonHucre As Range
    STOK_KARS_ICERIK_SON = 16
    Set sonHucre = ws.Range("A:N").Find(What:="*", _
        After:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False, SearchFormat:=False)
    If Not sonHucre Is Nothing Then
        If sonHucre.Row > 16 Then STOK_KARS_ICERIK_SON = sonHucre.Row
    End If
End Function

' Eski kurulumlu sayfaya NEREDEN / NEREYE sutunu bir kez eklenir: N genisler, sagdaki panel ayni miktar kayar,
' baslik bantlari N'ye uzar, filtre A:N olur. Sutun varsa hicbir sey yapilmaz.
Private Sub STOK_KARS_NN_HAZIRLA(ByVal ws As Worksheet)
    Dim eskiW As Double, fark As Double, sinir As Double, s As Shape, da As Boolean
    If CStr(ws.Range("N15").Value2) = KARS_NN Then Exit Sub
    On Error Resume Next
    da = application.DisplayAlerts: application.DisplayAlerts = False
    sinir = ws.Columns("O").Left - 1
    eskiW = ws.Columns("N").Width
    ws.Columns("N").ColumnWidth = 44
    fark = ws.Columns("N").Width - eskiW
    If fark > 0 Then
        For Each s In ws.Shapes
            If s.Left >= sinir Then s.Left = s.Left + fark
        Next s
    End If
    ws.Range("N1:N2000").Interior.Color = RGB(245, 247, 250)
    ws.Range("A1:N2").Merge
    ws.Range("A3:N3").Merge
    With ws.Range("N15")
        .Value = KARS_NN
        .Interior.Color = RGB(20, 55, 95)
        .Font.Color = RGB(255, 255, 255): .Font.bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A15:N15").AutoFilter
    application.DisplayAlerts = da
    On Error GoTo 0
End Sub

' NEREDEN / NEREYE: ayni malzemenin (ayni birim) iki tarih arasindaki azalislari artislarla eslestirilir.
' Ayni il icinde: once birebir ayni miktar, sonra buyukten kucuge dagitim. Iller arasi: yalniz birebir ayni miktar
' ve iki yonde tek aday (tesaduf eslesmesi olmasin). Karsi iller stok hafizasindan (kayit saati farki icin 1 saat tolerans).
' Karsiligi bulunmayan azalis "Sarf / depo dışı", artis "Alım / depo dışı".
Private Sub STOK_KARS_NEREDEN(ByVal ws As Worksheet, ByVal son As Long, ByVal t1 As Date, ByVal t2 As Date)
    Dim v As Variant, r As Long, n As Long, i As Long, j As Long, p As Long, q As Long, x As Double, c As Long
    Dim eLoc() As String, ePar() As String, eIl() As String, eQ() As Double, eRem() As Double, eRow() As Long, eTxt() As String
    Dim grup As Object, mallar As Object, tesisIl As Object, iller As Variant, sec As String, il As Variant
    Dim a As Object, b As Object, kk As Variant, q1 As Double, q2 As Double, rec As Variant, mk As String, g As Variant
    Dim dec() As Long, inc() As Long, nD As Long, ni As Long, tmp As Long, cikti() As Variant, satirlar() As String, parca() As String
    Dim metin As String, lab As String, tamami As Boolean, ok As String, ilR As String, cQ As Long
    Dim durumlar As Variant, icNakil As Boolean, sp() As String, ownT As String, ownD As String, degisti As Boolean
    If son < 16 Then Exit Sub
    v = ws.Range("A16:J" & son).Value2
    ReDim eLoc(1 To 64): ReDim ePar(1 To 64): ReDim eIl(1 To 64): ReDim eQ(1 To 64)
    ReDim eRem(1 To 64): ReDim eRow(1 To 64): ReDim eTxt(1 To 64)
    Set grup = CreateObject("Scripting.Dictionary"): grup.CompareMode = vbTextCompare
    Set mallar = CreateObject("Scripting.Dictionary"): mallar.CompareMode = vbTextCompare
    Set tesisIl = CreateObject("Scripting.Dictionary"): tesisIl.CompareMode = vbTextCompare
    iller = Array("ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
    For Each il In iller
        tesisIl(CStr(AKSORA_IL_TESIS(CStr(il), "CDP"))) = CStr(il)
        tesisIl(CStr(AKSORA_IL_TESIS(CStr(il), "EPP"))) = CStr(il)
    Next il
    ' Ekrandaki satirlar
    For r = 1 To UBound(v, 1)
        x = STOK_KARS_DOUBLE(v(r, 9))
        If Abs(x) > 0.0000001 Then
            mk = Trim$(CStr(v(r, 4))) & "|" & UCase$(Trim$(CStr(v(r, 10))))
            ilR = "?"
            If tesisIl.Exists(Trim$(CStr(v(r, 2)))) Then ilR = tesisIl(Trim$(CStr(v(r, 2))))
            STOK_KARS_E_EKLE n, eLoc, ePar, eIl, eQ, eRem, eRow, eTxt, grup, _
                Trim$(CStr(v(r, 2))) & " " & UCase$(Trim$(CStr(v(r, 3)))), UCase$(Trim$(CStr(v(r, 6)))), ilR, mk, x, r
            mallar(mk) = True
        End If
    Next r
    If n = 0 Then Exit Sub
    ' Karsi iller (secili il TUMU degilse): yalniz ekranda gecen malzemeler
    sec = AKFIX_KARS_IL()
    If sec <> "TUMU" Then
        For Each il In iller
            If CStr(il) <> sec Then
                Set a = AKFIX_KARS_IL_DURUM(CStr(il), t1 + 1# / 24#)
                Set b = Nothing
                If Not a Is Nothing Then Set b = AKFIX_KARS_IL_DURUM(CStr(il), t2 + 1# / 24#)
                If Not a Is Nothing And Not b Is Nothing Then
                    For Each kk In b.keys
                        If Not a.Exists(kk) Then a(kk) = Empty
                    Next kk
                    For Each kk In a.keys
                        rec = a(kk): q1 = 0: q2 = 0
                        If Not IsEmpty(rec) Then q1 = STOK_KARS_DOUBLE(rec(9))
                        If b.Exists(kk) Then
                            rec = b(kk): q2 = STOK_KARS_DOUBLE(rec(9))
                        End If
                        If Abs(q2 - q1) > 0.0000001 And Not IsEmpty(rec) Then
                            mk = Trim$(CStr(rec(7))) & "|" & UCase$(Trim$(CStr(rec(10))))
                            If mallar.Exists(mk) Then
                                STOK_KARS_E_EKLE n, eLoc, ePar, eIl, eQ, eRem, eRow, eTxt, grup, _
                                    Trim$(CStr(rec(1))) & " " & UCase$(Trim$(CStr(rec(2)))), UCase$(Trim$(CStr(rec(6)))), CStr(il), mk, q2 - q1, 0
                            End If
                        End If
                    Next kk
                End If
            End If
        Next il
    End If
    ' Ilce (K) ve yuklenici (M) depolari: _KH_ gunluk kaydi (yoksa ilce biriminin _IH_ gecmisi, yalniz K).
    For Each il In iller
        Set a = STOK_KARS_WS("_KH_" & CStr(il))
        c = 0: cQ = 0
        If Not a Is Nothing Then c = STOK_KARS_GUN_KOL(a, t1): cQ = STOK_KARS_GUN_KOL(a, t2)
        If c > 0 And cQ > 0 And c <> cQ Then
            STOK_KARS_GECMIS_EKLE a, c, cQ, True, CStr(il), mallar, n, eLoc, ePar, eIl, eQ, eRem, eRow, eTxt, grup
        Else
            Set a = STOK_KARS_WS("_IH_" & CStr(il))
            c = 0: cQ = 0
            If Not a Is Nothing Then c = STOK_KARS_GUN_KOL(a, t1): cQ = STOK_KARS_GUN_KOL(a, t2)
            If c > 0 And cQ > 0 And c <> cQ Then
                STOK_KARS_GECMIS_EKLE a, c, cQ, False, CStr(il), mallar, n, eLoc, ePar, eIl, eQ, eRem, eRow, eTxt, grup
            End If
        End If
    Next il
    ' Eslestirme (malzeme malzeme)
    For Each kk In grup.keys
        g = Split(Mid$(CStr(grup(kk)), 2), ",")
        ReDim dec(0 To UBound(g) + 1): ReDim inc(0 To UBound(g) + 1)
        ' 1) ayni il icinde
        For Each il In iller
            nD = 0: ni = 0
            For i = 0 To UBound(g)
                p = CLng(g(i))
                If eIl(p) = CStr(il) Then
                    If eQ(p) < 0 Then dec(nD) = p: nD = nD + 1 Else inc(ni) = p: ni = ni + 1
                End If
            Next i
            If nD > 0 And ni > 0 Then
                For i = 0 To nD - 2
                    For j = i + 1 To nD - 1
                        If eRem(dec(j)) > eRem(dec(i)) Then tmp = dec(i): dec(i) = dec(j): dec(j) = tmp
                    Next j
                Next i
                For i = 0 To ni - 2
                    For j = i + 1 To ni - 1
                        If eRem(inc(j)) > eRem(inc(i)) Then tmp = inc(i): inc(i) = inc(j): inc(j) = tmp
                    Next j
                Next i
                For i = 0 To nD - 1
                    For j = 0 To ni - 1
                        p = dec(i): q = inc(j)
                        If eRem(p) > 0.0000001 And eRem(q) > 0.0000001 And Abs(eRem(p) - eRem(q)) < 0.0000001 Then
                            STOK_KARS_E_BAGLA p, q, eRem(p), eLoc, ePar, eRem, eTxt
                        End If
                    Next j
                Next i
                For i = 0 To nD - 1
                    For j = 0 To ni - 1
                        p = dec(i): q = inc(j)
                        If eRem(p) > 0.0000001 And eRem(q) > 0.0000001 Then
                            x = eRem(p): If eRem(q) < x Then x = eRem(q)
                            STOK_KARS_E_BAGLA p, q, x, eLoc, ePar, eRem, eTxt
                        End If
                    Next j
                Next i
            End If
        Next il
        ' 2) iller arasi: birebir ayni miktar, iki yonde tek aday
        For i = 0 To UBound(g)
            p = CLng(g(i))
            If eQ(p) < 0 And eRem(p) > 0.0000001 Then
                c = 0: cQ = 0
                For j = 0 To UBound(g)
                    q = CLng(g(j))
                    If eQ(q) > 0 And eIl(q) <> eIl(p) And eRem(q) > 0.0000001 And Abs(eRem(q) - eRem(p)) < 0.0000001 Then c = c + 1: cQ = q
                Next j
                If c = 1 Then
                    c = 0
                    For j = 0 To UBound(g)
                        q = CLng(g(j))
                        If eQ(q) < 0 And eIl(q) <> eIl(cQ) And eRem(q) > 0.0000001 And Abs(eRem(q) - eRem(cQ)) < 0.0000001 Then c = c + 1
                    Next j
                    If c = 1 Then STOK_KARS_E_BAGLA p, cQ, eRem(p), eLoc, ePar, eRem, eTxt
                End If
            End If
        Next i
    Next kk
    ' Metin
    ReDim cikti(1 To UBound(v, 1), 1 To 1)
    If son > 16 Then
        durumlar = ws.Range("A16:A" & son).Value2
    Else
        ReDim durumlar(1 To 1, 1 To 1): durumlar(1, 1) = ws.Range("A16").Value2
    End If
    For p = 1 To n
        If eRow(p) > 0 Then
            ok = IIf(eQ(p) < 0, ChrW(8594), ChrW(8592))
            metin = vbNullString
            If Len(eTxt(p)) > 0 Then
                satirlar = Split(Mid$(eTxt(p), 2), vbLf)
                tamami = (UBound(satirlar) = 0 And eRem(p) <= 0.0000001)
                For i = 0 To UBound(satirlar)
                    parca = Split(satirlar(i), vbTab)
                    lab = parca(0)
                    If tamami Then
                        metin = lab
                    Else
                        metin = metin & IIf(Len(metin) > 0, ", ", "") & lab & " (" & STOK_KARS_MIK(CDbl(parca(1))) & ")"
                    End If
                Next i
                metin = ok & " " & metin
            End If
            If eRem(p) > 0.0000001 Then
                ' Karsiligi olmayan kisim: hurda / cikma partisinde demontaj, digerlerinde alim veya sarf.
                If InStr(1, ePar(p), "HURDA", vbTextCompare) > 0 Then
                    lab = IIf(eQ(p) < 0, "hurda çıkışı (satış / imha)", "hurda girişi (demontaj)")
                ElseIf InStr(1, ePar(p), "CIKMA", vbTextCompare) > 0 Or InStr(1, ePar(p), "ÇIKMA", vbTextCompare) > 0 Then
                    lab = IIf(eQ(p) < 0, "çıkma malzeme kullanıldı", "çıkma girişi (demontaj)")
                Else
                    lab = IIf(eQ(p) < 0, "sarf / depo dışı", "alım / depo dışı")
                End If
                If Len(metin) = 0 Then
                    metin = UCase$(Left$(lab, 1)) & Mid$(lab, 2)
                Else
                    metin = metin & "  •  " & STOK_KARS_MIK(eRem(p)) & " " & lab
                End If
            End If
            cikti(eRow(p), 1) = metin
            ' Ayni uretim yerinin D depolari arasinda tamami eslesen hareket: sag cepten sol cebe (IC NAKIL).
            ' Azalan / artan sayilarina ve filtrelerine girmez; toplam deger zaten degismez.
            If eRem(p) <= 0.0000001 And Len(eTxt(p)) > 0 Then
                ownT = Trim$(CStr(v(eRow(p), 2))): ownD = UCase$(Trim$(CStr(v(eRow(p), 3))))
                icNakil = (Left$(ownD, 1) = "D")
                satirlar = Split(Mid$(eTxt(p), 2), vbLf)
                For i = 0 To UBound(satirlar)
                    sp = Split(Split(satirlar(i), vbTab)(0), " ")
                    If UBound(sp) <> 1 Then
                        icNakil = False
                    ElseIf sp(0) <> ownT Or Left$(sp(1), 1) <> "D" Or sp(1) = ownD Then
                        icNakil = False
                    End If
                Next i
                If icNakil Then durumlar(eRow(p), 1) = "İÇ NAKİL": degisti = True
            End If
        End If
    Next p
    ws.Range("N16:N" & son).Value2 = cikti
    If degisti Then
        ws.Range("A16:A" & son).Value2 = durumlar
        For r = 1 To UBound(v, 1)
            If CStr(durumlar(r, 1)) = "İÇ NAKİL" Then ws.Cells(r + 15, 1).Font.Color = RGB(100, 116, 139)
        Next r
    End If
End Sub

Private Sub STOK_KARS_E_EKLE(ByRef n As Long, ByRef eLoc() As String, ByRef ePar() As String, ByRef eIl() As String, _
    ByRef eQ() As Double, ByRef eRem() As Double, ByRef eRow() As Long, ByRef eTxt() As String, ByVal grup As Object, _
    ByVal yer As String, ByVal parti As String, ByVal il As String, ByVal mk As String, ByVal fark As Double, ByVal satir As Long)
    n = n + 1
    If n > UBound(eLoc) Then
        ReDim Preserve eLoc(1 To 2 * n): ReDim Preserve ePar(1 To 2 * n): ReDim Preserve eIl(1 To 2 * n)
        ReDim Preserve eQ(1 To 2 * n): ReDim Preserve eRem(1 To 2 * n): ReDim Preserve eRow(1 To 2 * n)
        ReDim Preserve eTxt(1 To 2 * n)
    End If
    eLoc(n) = yer: ePar(n) = parti: eIl(n) = il: eQ(n) = fark: eRem(n) = Abs(fark): eRow(n) = satir: eTxt(n) = vbNullString
    If grup.Exists(mk) Then grup(mk) = grup(mk) & "," & n Else grup(mk) = "," & n
End Sub

' p (azalan) -> q (artan) x kadar: iki tarafa karsi yer yazilir. Ayni depoda parti degisiminde parti eklenir.
Private Sub STOK_KARS_E_BAGLA(ByVal p As Long, ByVal q As Long, ByVal x As Double, ByRef eLoc() As String, _
    ByRef ePar() As String, ByRef eRem() As Double, ByRef eTxt() As String)
    Dim lp As String, lq As String
    lp = eLoc(p): lq = eLoc(q)
    If lp = lq Then
        If Len(ePar(p)) > 0 Then lp = lp & " " & ePar(p)
        If Len(ePar(q)) > 0 Then lq = lq & " " & ePar(q)
    End If
    eTxt(p) = eTxt(p) & vbLf & lq & vbTab & CStr(x)
    eTxt(q) = eTxt(q) & vbLf & lp & vbTab & CStr(x)
    eRem(p) = eRem(p) - x: eRem(q) = eRem(q) - x
End Sub

Private Function STOK_KARS_MIK(ByVal x As Double) As String
    If Abs(x - Round(x, 0)) < 0.0000001 Then
        STOK_KARS_MIK = AKSORA_SAYI_METNI(Round(x, 0), "#,##0")
    Else
        STOK_KARS_MIK = AKSORA_SAYI_METNI(x, "#,##0.###")
    End If
End Function

' Eski kurulumlu sayfada ust bantlar (2. tarih, KPI'lar) ve hizli dugmeler N sutununa kadar uzatilir.
Private Sub STOK_KARS_UST_GENISLET(ByVal ws As Worksheet)
    Dim da As Boolean, ad As Variant, i As Long, s As Shape, alan As Range, w As Double
    On Error Resume Next
    da = application.DisplayAlerts: application.DisplayAlerts = False
    If ws.Range("J9").MergeArea.Columns.Count < 5 Then
        For Each ad In Array("H5:N5", "I6:N6", "I7:N7", "J9:N9", "J10:N10", "J11:N11", "J12:N12")
            ws.Range(CStr(ad)).Merge
        Next ad
    End If
    Set alan = ws.Range("A14:N14")
    w = alan.Width / 6
    i = 0
    For Each ad In Array("btnKarsTum", "btnKarsAzalan", "btnKarsArtan", "btnKarsYeni", "btnKarsSifir", "btnKarsKarsilastir")
        Set s = Nothing
        Set s = ws.Shapes(CStr(ad))
        If Not s Is Nothing Then
            s.Left = alan.Left + i * w: s.Width = w
        End If
        i = i + 1
    Next ad
    application.DisplayAlerts = da
    On Error GoTo 0
End Sub

Private Function STOK_KARS_WS(ByVal ad As String) As Worksheet
    On Error Resume Next
    Set STOK_KARS_WS = ThisWorkbook.Worksheets(ad)
End Function

' Gunluk gecmiste t gununun sutunu; o gun kayit yoksa t'den onceki son kayit. Yoksa 0.
Private Function STOK_KARS_GUN_KOL(ByVal ws As Worksheet, ByVal t As Date) As Long
    Dim c As Long, son As Long, x As Variant, enIyi As Double
    son = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To son
        x = ws.Cells(1, c).Value
        If IsDate(x) Then
            If Int(CDbl(CDate(x))) = Int(CDbl(t)) Then STOK_KARS_GUN_KOL = c: Exit Function
            If CDbl(CDate(x)) <= CDbl(t) And CDbl(CDate(x)) > enIyi Then enIyi = CDbl(CDate(x)): STOK_KARS_GUN_KOL = c
        End If
    Next c
End Function

' Gunluk gecmisin iki sutunu arasindaki farklar karsi taraf olarak eklenir (yalniz ekranda gecen malzemeler).
' kh = True: _KH_ (uretim yeri, depo, malzeme, parti, birim, miktar); False: _IH_ (depo, uretim yeri, malzeme,
' degerleme, depo adi, tanim, birim, miktar, deger; eski kayitta degerleme yok).
Private Sub STOK_KARS_GECMIS_EKLE(ByVal ws As Worksheet, ByVal c1 As Long, ByVal c2 As Long, ByVal kh As Boolean, _
    ByVal il As String, ByVal mallar As Object, ByRef n As Long, ByRef eLoc() As String, ByRef ePar() As String, _
    ByRef eIl() As String, ByRef eQ() As Double, ByRef eRem() As Double, ByRef eRow() As Long, ByRef eTxt() As String, _
    ByVal grup As Object)
    Dim d(1 To 2) As Object, dd As Object, d1 As Object, d2 As Object, i As Long, kol As Long, r As Long, son As Long, satirlar() As String, f() As String
    Dim j As Long, tesis As String, depo As String, mal As String, parti As String, birim As String, q As Double
    Dim mk As String, anahtar As String, t As Variant, kk As Variant, q1 As Double, q2 As Double
    For i = 1 To 2
        Set dd = CreateObject("Scripting.Dictionary"): dd.CompareMode = vbTextCompare
        Set d(i) = dd
        kol = IIf(i = 1, c1, c2)
        son = ws.Cells(ws.Rows.Count, kol).End(xlUp).Row
        For r = 2 To son
            satirlar = Split(CStr(ws.Cells(r, kol).Value2), vbLf)
            For j = 0 To UBound(satirlar)
                f = Split(satirlar(j), vbTab)
                mal = vbNullString
                If kh And UBound(f) >= 5 Then
                    tesis = f(0): depo = f(1): mal = f(2): parti = f(3): birim = f(4): q = Val(f(5))
                ElseIf Not kh And UBound(f) >= 8 Then
                    depo = f(0): tesis = f(1): mal = f(2): parti = f(3): birim = f(6): q = Val(f(7))
                ElseIf Not kh And UBound(f) = 7 Then
                    depo = f(0): tesis = f(1): mal = f(2): parti = vbNullString: birim = f(5): q = Val(f(6))
                End If
                If Len(mal) > 0 Then
                    mk = Trim$(mal) & "|" & UCase$(Trim$(birim))
                    If mallar.Exists(mk) Then
                        anahtar = Trim$(tesis) & " " & UCase$(Trim$(depo)) & "|" & UCase$(Trim$(parti)) & "|" & mk
                        If dd.Exists(anahtar) Then
                            t = dd(anahtar): t(3) = t(3) + q: dd(anahtar) = t
                        Else
                            dd(anahtar) = Array(Trim$(tesis) & " " & UCase$(Trim$(depo)), UCase$(Trim$(parti)), mk, q)
                        End If
                    End If
                End If
            Next j
        Next r
    Next i
    Set d1 = d(1): Set d2 = d(2)
    For Each kk In d2.keys
        If Not d1.Exists(kk) Then d1(kk) = Empty
    Next kk
    For Each kk In d1.keys
        q1 = 0: q2 = 0
        t = d1(kk)
        If Not IsEmpty(t) Then q1 = CDbl(t(3))
        If d2.Exists(kk) Then
            t = d2(kk): q2 = CDbl(t(3))
        End If
        If Abs(q2 - q1) > 0.0000001 And Not IsEmpty(t) Then
            STOK_KARS_E_EKLE n, eLoc, ePar, eIl, eQ, eRem, eRow, eTxt, grup, CStr(t(0)), CStr(t(1)), il, CStr(t(2)), q2 - q1, 0
        End If
    Next kk
End Sub

' Ilce ve yuklenici stoklarinin gunluk kaydi (_KH_<IL>): her sutun bir gun (gunun son guncellemesi), en fazla 60 gun.
' Satir: uretim yeri TAB depo TAB malzeme TAB parti TAB birim TAB miktar. Hata olursa guncelleme etkilenmez.
Public Sub STOK_KH_KAYDET_CANLI()
    Dim a As Worksheet, b As Worksheet
    On Error Resume Next
    Set a = STOK_KARS_WS("İLÇE_STOK"): Set b = STOK_KARS_WS("YÜKLENİCİ_STOK")
    STOK_KH_KAYDET AKSORA_CANLI_STOK_IL(), a, b
End Sub

Public Sub STOK_KH_KAYDET(ByVal il As String, ByVal wsIlc As Worksheet, ByVal wsYuk As Worksheet)
    Dim ws As Worksheet, d As Object, kaynak As Variant, kw As Worksheet, v As Variant, r As Long, son As Long, k As String
    Dim c As Long, x As Variant, satir As String, tampon As String, j As Long, tarih As Date, kk As Variant
    On Error GoTo son
    il = UCase$(Trim$(il))
    If Len(il) = 0 Then Exit Sub
    tarih = Now
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For Each kaynak In Array(wsIlc, wsYuk)
        If Not kaynak Is Nothing Then
            Set kw = kaynak
            son = AKSORA_SON_SATIR(kw, "G")
            If son >= 7 Then
                v = kw.Range("A7:L" & son).Value2
                For r = 1 To UBound(v, 1)
                    If Len(Trim$(CStr(v(r, 7)))) > 0 And Not IsError(v(r, 9)) Then
                        k = Trim$(CStr(v(r, 1))) & vbTab & UCase$(Trim$(CStr(v(r, 2)))) & vbTab & Trim$(CStr(v(r, 7))) & vbTab & _
                            UCase$(Trim$(CStr(v(r, 6)))) & vbTab & UCase$(Trim$(CStr(v(r, 10))))
                        d(k) = d(k) + STOK_KARS_DOUBLE(v(r, 9))
                    End If
                Next r
            End If
        End If
    Next kaynak
    If d.Count = 0 Then Exit Sub
    Set ws = STOK_KARS_WS("_KH_" & il)
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "_KH_" & il
        ws.Visible = xlSheetVeryHidden
    End If
    son = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If son = 1 And Len(CStr(ws.Cells(1, 1).Value2)) = 0 Then son = 0
    For c = son To 1 Step -1
        x = ws.Cells(1, c).Value
        If IsDate(x) Then
            If Int(CDbl(CDate(x))) = Int(CDbl(tarih)) Then ws.Columns(c).Delete
        Else
            ws.Columns(c).Delete
        End If
    Next c
    son = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If son = 1 And Len(CStr(ws.Cells(1, 1).Value2)) = 0 Then son = 0
    Do While son >= 60
        ws.Columns(1).Delete
        son = son - 1
    Loop
    c = son + 1
    ws.Columns(c).NumberFormat = "@"
    ws.Cells(1, c).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Cells(1, c).Value = tarih
    j = 1
    For Each kk In d.keys
        satir = CStr(kk) & vbTab & Trim$(Str$(d(kk)))
        If Len(tampon) + Len(satir) + 1 > 30000 And Len(tampon) > 0 Then
            j = j + 1: ws.Cells(j, c).Value2 = tampon: tampon = vbNullString
        End If
        If Len(tampon) > 0 Then tampon = tampon & vbLf
        tampon = tampon & satir
    Next kk
    If Len(tampon) > 0 Then j = j + 1: ws.Cells(j, c).Value2 = tampon
son:
End Sub
