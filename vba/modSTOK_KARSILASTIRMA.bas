Attribute VB_Name = "modSTOK_KARSILASTIRMA"
Option Explicit

Private Const KARS_SAYFA As String = "STOK_KARÞILAÞTIRMA"
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
        "STOK KARÞILAÞTIRMA sayfasý oluþturulamadý." & vbCrLf & vbCrLf & _
        "Hata No: " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
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

    Dim D1 As Object
    Dim D2 As Object
    Dim Tum As Object
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

    Snap1 = STOK_KARS_SECILEN_SNAPSHOT(ws.Range("B6").Value)
    Snap2 = STOK_KARS_SECILEN_SNAPSHOT(ws.Range("I6").Value)

    If Len(Snap1) = 0 Or Len(Snap2) = 0 Then
        MsgBox _
            "Ýki karþýlaþtýrma tarihini de seçiniz.", _
            vbExclamation, _
            "STOK KARÞILAÞTIRMA"
        Exit Sub
    End If

    tarih1 = STOK_KARS_SNAPSHOT_TARIH(Snap1)
    tarih2 = STOK_KARS_SNAPSHOT_TARIH(Snap2)

    If tarih1 >= tarih2 Then
        MsgBox _
            "1. tarih, 2. tarihten eski olmalýdýr.", _
            vbExclamation, _
            "STOK KARÞILAÞTIRMA"
        Exit Sub
    End If

    application.ScreenUpdating = False

    AKFIX_KARS_GECMIS_FARK = 0
    Set D1 = STOK_KARS_DURUM_OLUSTUR(Snap1)
    Set D2 = STOK_KARS_DURUM_OLUSTUR(Snap2)

    If D1 Is Nothing Or D2 Is Nothing Then
        Err.Raise _
            vbObjectError + 7301, _
            "STOK_KARSILASTIR", _
            "Seçilen stok durumlarý oluþturulamadý."
    End If

    Set Kimlik1 = STOK_KARS_KIMLIK_SETI(D1)
    Set Kimlik2 = STOK_KARS_KIMLIK_SETI(D2)

    Set Tum = CreateObject("Scripting.Dictionary")
    Tum.CompareMode = vbTextCompare

    For Each anahtar In D1.keys
        If Not Tum.Exists(CStr(anahtar)) Then
            Tum.Add CStr(anahtar), True
        End If
    Next anahtar

    For Each anahtar In D2.keys
        If Not Tum.Exists(CStr(anahtar)) Then
            Tum.Add CStr(anahtar), True
        End If
    Next anahtar

    ws.Range("A16:N" & STOK_KARS_ICERIK_SON(ws)).ClearContents

    CikisSatir = 16

    For Each anahtar In Tum.keys

        r1 = Empty
        r2 = Empty
        durum = vbNullString

        If D1.Exists(CStr(anahtar)) Then
            r1 = D1(CStr(anahtar))
        End If

        If D2.Exists(CStr(anahtar)) Then
            r2 = D2(CStr(anahtar))
        End If

        If Not D1.Exists(CStr(anahtar)) And _
           D2.Exists(CStr(anahtar)) Then

            If Kimlik1.Exists(STOK_KARS_KIMLIK(r2)) Then
                durum = "ARTTI"
                artan = artan + 1
            Else
                durum = "YENÝ"
                yeni = yeni + 1
            End If

        ElseIf D1.Exists(CStr(anahtar)) And _
               Not D2.Exists(CStr(anahtar)) Then

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

            durum = "DEÐER DEÐÝÞTÝ"

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
        "Son karþýlaþtýrýlan: " & Format$(tarih1, "dd.mm.yyyy hh:nn:ss")

    ws.Range("I7").Value = _
        "Son karþýlaþtýrýlan: " & Format$(tarih2, "dd.mm.yyyy hh:nn:ss")

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
        "Karþýlaþtýrma yapýlamadý." & vbCrLf & vbCrLf & _
        "Hata No: " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "AKSORA - STOK KARÞILAÞTIRMA"

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
    STOK_KARS_FILTRE "YENÝ"
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
    MsgBox "Ana panel açýlamadý: " & Err.Description, vbExclamation, "AKSORA"
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
            durum = Replace(durum, "Ý", "I")

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
                after:=ThisWorkbook.Worksheets( _
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
        .Value = "STOK KARÞILAÞTIRMA MERKEZÝ"

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
            "CDP 7155 + EPP 7055  |  Normal Stok Hafýzasý"

        .Interior.Color = RGB(20, 55, 95)

        .Font.Color = RGB(215, 225, 238)

        .HorizontalAlignment = xlCenter

    End With

End Sub

Private Sub STOK_KARS_SECIM_ALANI(ByVal ws As Worksheet)

    With ws.Range("A5:F5")

        .Merge
        .Value = "1. KARÞILAÞTIRMA TARÝHÝ"

        .Font.bold = True
        .Font.Color = RGB(75, 85, 95)

        .HorizontalAlignment = xlCenter

    End With

    With ws.Range("H5:M5")

        .Merge
        .Value = "2. KARÞILAÞTIRMA TARÝHÝ"

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

    With ws.Range("I6:M6")

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

    With ws.Range("I7:M7")

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
        "YENÝ", _
        "H10"

    STOK_KARS_KPI _
        ws, _
        "J9:M9", _
        "J10:M10", _
        "SIFIRLANAN", _
        "K10"

    STOK_KARS_KPI _
        ws, _
        "A11:C11", _
        "A12:C12", _
        "ESKÝ STOK DEÐERÝ", _
        "B12"

    STOK_KARS_KPI _
        ws, _
        "D11:F11", _
        "D12:F12", _
        "YENÝ STOK DEÐERÝ", _
        "E12"

    STOK_KARS_KPI _
        ws, _
        "G11:I11", _
        "G12:I12", _
        "DEÐER FARKI", _
        "H12"

    STOK_KARS_KPI _
        ws, _
        "J11:M11", _
        "J12:M12", _
        "TOPLAM DEÐÝÞÝM", _
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
        "YENÝ", _
        "STOK_KARS_FILTRE_YENI", _
        "btnKarsYeni"

    STOK_KARS_HIZLI_BUTON _
        ws, 5, 6, _
        "SIFIRLANAN", _
        "STOK_KARS_FILTRE_SIFIR", _
        "btnKarsSifir"

    STOK_KARS_HIZLI_BUTON _
        ws, 6, 6, _
        "KARÞILAÞTIR", _
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

    Set alan = ws.Range("A14:M14")

    ToplamGenislik = alan.Width
    ButonGenislik = ToplamGenislik / toplam
    ButonYukseklik = alan.Height

    sol = _
        alan.Left + _
        ((sira - 1) * ButonGenislik)

    ust = alan.Top

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
        "PARTÝ", _
        "ESKÝ MÝKTAR", _
        "YENÝ MÝKTAR", _
        "FARK", _
        "BÝRÝM", _
        "ESKÝ DEÐER", _
        "YENÝ DEÐER", _
        "DEÐER FARKI", _
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
            hedef.Top, _
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
        .Top = hedef.Top
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
    ThisWorkbook.Worksheets("STOK_KARÞILAÞTIRMA").STOK_KARS_FILTRE_TAKIP_BASLAT
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

    Set ws = ThisWorkbook.Worksheets("STOK_KARÞILAÞTIRMA")

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
        " | Normal Stok Hafýzasý"

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
        after:=ws.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
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
' Karsiligi bulunmayan azalis "Sarf / depo dýþý", artis "Alým / depo dýþý".
Private Sub STOK_KARS_NEREDEN(ByVal ws As Worksheet, ByVal son As Long, ByVal t1 As Date, ByVal t2 As Date)
    Dim v As Variant, r As Long, n As Long, i As Long, j As Long, p As Long, q As Long, x As Double, c As Long
    Dim eLoc() As String, ePar() As String, eIl() As String, eQ() As Double, eRem() As Double, eRow() As Long, eTxt() As String
    Dim grup As Object, mallar As Object, tesisIl As Object, iller As Variant, sec As String, il As Variant
    Dim a As Object, b As Object, kk As Variant, q1 As Double, q2 As Double, rec As Variant, mk As String, g As Variant
    Dim dec() As Long, inc() As Long, nd As Long, ni As Long, tmp As Long, cikti() As Variant, satirlar() As String, parca() As String
    Dim metin As String, lab As String, tamami As Boolean, ok As String, ilR As String, cq As Long
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
    ' Eslestirme (malzeme malzeme)
    For Each kk In grup.keys
        g = Split(Mid$(CStr(grup(kk)), 2), ",")
        ReDim dec(0 To UBound(g) + 1): ReDim inc(0 To UBound(g) + 1)
        ' 1) ayni il icinde
        For Each il In iller
            nd = 0: ni = 0
            For i = 0 To UBound(g)
                p = CLng(g(i))
                If eIl(p) = CStr(il) Then
                    If eQ(p) < 0 Then dec(nd) = p: nd = nd + 1 Else inc(ni) = p: ni = ni + 1
                End If
            Next i
            If nd > 0 And ni > 0 Then
                For i = 0 To nd - 2
                    For j = i + 1 To nd - 1
                        If eRem(dec(j)) > eRem(dec(i)) Then tmp = dec(i): dec(i) = dec(j): dec(j) = tmp
                    Next j
                Next i
                For i = 0 To ni - 2
                    For j = i + 1 To ni - 1
                        If eRem(inc(j)) > eRem(inc(i)) Then tmp = inc(i): inc(i) = inc(j): inc(j) = tmp
                    Next j
                Next i
                For i = 0 To nd - 1
                    For j = 0 To ni - 1
                        p = dec(i): q = inc(j)
                        If eRem(p) > 0.0000001 And eRem(q) > 0.0000001 And Abs(eRem(p) - eRem(q)) < 0.0000001 Then
                            STOK_KARS_E_BAGLA p, q, eRem(p), eLoc, ePar, eRem, eTxt
                        End If
                    Next j
                Next i
                For i = 0 To nd - 1
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
                c = 0: cq = 0
                For j = 0 To UBound(g)
                    q = CLng(g(j))
                    If eQ(q) > 0 And eIl(q) <> eIl(p) And eRem(q) > 0.0000001 And Abs(eRem(q) - eRem(p)) < 0.0000001 Then c = c + 1: cq = q
                Next j
                If c = 1 Then
                    c = 0
                    For j = 0 To UBound(g)
                        q = CLng(g(j))
                        If eQ(q) < 0 And eIl(q) <> eIl(cq) And eRem(q) > 0.0000001 And Abs(eRem(q) - eRem(cq)) < 0.0000001 Then c = c + 1
                    Next j
                    If c = 1 Then STOK_KARS_E_BAGLA p, cq, eRem(p), eLoc, ePar, eRem, eTxt
                End If
            End If
        Next i
    Next kk
    ' Metin
    ReDim cikti(1 To UBound(v, 1), 1 To 1)
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
                If Len(metin) = 0 Then
                    metin = IIf(eQ(p) < 0, "Sarf / depo dýþý", "Alým / depo dýþý")
                Else
                    metin = metin & "  •  " & STOK_KARS_MIK(eRem(p)) & IIf(eQ(p) < 0, " sarf / depo dýþý", " alým / depo dýþý")
                End If
            End If
            cikti(eRow(p), 1) = metin
        End If
    Next p
    ws.Range("N16:N" & son).Value2 = cikti
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
