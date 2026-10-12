Attribute VB_Name = "modAKSORA_STOK_BIRLESTIR"
Option Explicit

    '==========================================================
    ' AKSORA
    ' modBIRLESTIR_V2
    '
    ' NORMAL STOK BÝRLEÞTÝRME MOTORU
    '
    ' Kaynak:
    ' CDP_STOK = 7155
    ' EPP_STOK = 7055
    '
    ' Hedef:
    ' STOKLAR
    ' RAPORLAR
    '
    ' SAP BAÐLANTISI YOKTUR.
    '==========================================================

    Private Const BIR_CDP_STOK As String = "CDP_STOK"
    Private Const BIR_EPP_STOK As String = "EPP_STOK"
    Private Const BIR_STOKLAR As String = "STOKLAR"
    Private Const BIR_RAPORLAR As String = "RAPORLAR"

    Public Sub STOKLAR_RAPORLAR_BIRLESTIR_V2()

    Dim wsCDP As Worksheet
    Dim wsEPP As Worksheet
    Dim wsStoklar As Worksheet
    Dim wsRaporlar As Worksheet

    Dim VeriCDP As Variant
    Dim VeriEPP As Variant
    Dim BirlesikVeri As Variant

    Dim EskiStok As Variant

    Dim CDPSatir As Long
    Dim EPPSatir As Long
    Dim ToplamSatir As Long

    Dim EskiSonSatir As Long
    Dim EskiSonSutun As Long

    Dim baslangic As Double
    Dim ToplamSure As Double

    Dim YedekVar As Boolean
    Dim yazmaBasladi As Boolean

    Dim hataNo As Long
    Dim hataAciklama As String

    On Error GoTo Hata

    Set wsCDP = _
    ThisWorkbook.Worksheets(BIR_CDP_STOK)

    Set wsEPP = _
    ThisWorkbook.Worksheets(BIR_EPP_STOK)

    Set wsStoklar = _
    ThisWorkbook.Worksheets(BIR_STOKLAR)

    Set wsRaporlar = _
    ThisWorkbook.Worksheets(BIR_RAPORLAR)

    CORE_ISLEM_BASLAT _
    "Stoklar ve raporlar birleþtiriliyor"

    baslangic = Timer

    CORE_DURUM _
    "CDP ve EPP stoklarý belleðe alýnýyor..."

    CDPSatir = _
    BIR_V2_VERI_SATIRI_SAYISI(wsCDP)

    EPPSatir = _
    BIR_V2_VERI_SATIRI_SAYISI(wsEPP)

    If CDPSatir <= 0 Then

    Err.Raise _
        vbObjectError + 3301, _
        "STOKLAR_RAPORLAR_BIRLESTIR_V2", _
        "CDP_STOK sayfasýnda birleþtirilecek veri bulunamadý."

    End If

    If EPPSatir <= 0 Then

    Err.Raise _
        vbObjectError + 3302, _
        "STOKLAR_RAPORLAR_BIRLESTIR_V2", _
        "EPP_STOK sayfasýnda birleþtirilecek veri bulunamadý."

    End If

    VeriCDP = _
    wsCDP.Range("A7").Resize( _
    CDPSatir, _
    12).Value2

    VeriEPP = _
    wsEPP.Range("A7").Resize( _
    EPPSatir, _
    12).Value2

        BIR_V2_KAYNAK_DOGRULA _
        VeriCDP, _
        AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "CDP"), _
        "CDP_STOK"

    BIR_V2_KAYNAK_DOGRULA _
        VeriEPP, _
        AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "EPP"), _
        "EPP_STOK"

    CORE_DURUM _
    "CDP ve EPP stoklarý birleþtiriliyor..."

    BirlesikVeri = _
    BIR_V2_ARRAY_BIRLESTIR( _
    VeriCDP, _
    VeriEPP)

    ToplamSatir = _
    CDPSatir + EPPSatir

    EskiSonSatir = _
    STOK_SON_SATIR(wsStoklar)

    EskiSonSutun = _
    STOK_SON_SUTUN(wsStoklar)

    If EskiSonSatir >= 7 And _
    EskiSonSutun >= 1 Then

    If EskiSonSutun > 12 Then
        EskiSonSutun = 12
    End If

    EskiStok = _
        wsStoklar.Range( _
            wsStoklar.Cells(7, 1), _
            wsStoklar.Cells( _
                EskiSonSatir, _
                EskiSonSutun)).Value2

    YedekVar = True

    End If

    CORE_DURUM _
    "Birleþik stok STOKLAR sayfasýna yazýlýyor..."

    yazmaBasladi = True

    BIR_V2_STOK_BASLIKLARI wsStoklar

    wsStoklar.Range( _
    "A7:L" & AKSORA_V13_SON(wsStoklar)).ClearContents

    wsStoklar.Range("A7").Resize( _
    ToplamSatir, _
    12).Value2 = _
    BirlesikVeri

    BIR_V2_STOK_OZET _
    wsStoklar, _
    ToplamSatir

    CORE_DURUM _
    "Birleþik depo raporu hazýrlanýyor..."

    BIR_V2_RAPOR_OLUSTUR _
    wsRaporlar, _
    BirlesikVeri

    CORE_DURUM _
    "Birleþik stok görünümü düzenleniyor..."

    BIR_V2_STOK_GORSEL _
    wsStoklar, _
    ToplamSatir

    ToplamSure = _
    CORE_ISLEM_BITIR()

    application.Calculation = xlCalculationAutomatic
    application.CalculateFull

    If Not ANA_SESSIZ_MOD Then

    MsgBox _
        "STOKLAR ve RAPORLAR baþarýyla birleþtirildi." & _
        vbCrLf & vbCrLf & _
        "CDP Satýrý   : " & _
        AKSORA_SAYI_METNI(CDPSatir, "#,##0") & vbCrLf & _
        "EPP Satýrý   : " & _
        AKSORA_SAYI_METNI(EPPSatir, "#,##0") & vbCrLf & _
        "Toplam Satýr : " & _
        AKSORA_SAYI_METNI(ToplamSatir, "#,##0") & vbCrLf & _
        "Toplam Deðer : " & _
        Format( _
            BIR_V2_TOPLAM_DEGER(BirlesikVeri), _
            "#,##0.00") & " TL" & vbCrLf & _
        "Ýþlem Süresi : " & _
        AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn" & _
        vbCrLf & vbCrLf & _
        "SAP baðlantýsý kullanýlmadý.", _
        vbInformation, _
        "AKSORA - BÝRLEÞTÝR"

    End If

    Exit Sub

Hata:

    hataNo = Err.Number
    hataAciklama = Err.Description

    On Error Resume Next

    If yazmaBasladi Then

    CORE_DURUM _
        "Hata oluþtu. Eski STOKLAR verisi geri yükleniyor..."

    wsStoklar.Range( _
        "A7:L" & AKSORA_V13_SON(wsStoklar)).ClearContents

    If YedekVar Then

        If IsArray(EskiStok) Then

            wsStoklar.Range("A7").Resize( _
                UBound(EskiStok, 1), _
                UBound(EskiStok, 2)).Value2 = _
                EskiStok

        End If

    End If

    End If

    On Error GoTo 0

    ToplamSure = _
    CORE_ISLEM_BITIR()

    If Not ANA_SESSIZ_MOD Then

    MsgBox _
        "STOKLAR / RAPORLAR birleþtirme sýrasýnda hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata No : " & hataNo & vbCrLf & _
        "Açýklama: " & hataAciklama & _
        vbCrLf & vbCrLf & _
        "Mevcut STOKLAR verisi korunmaya çalýþýldý.", _
        vbCritical, _
        "AKSORA - BÝRLEÞTÝR"

    End If

    End Sub

    Private Function BIR_V2_VERI_SATIRI_SAYISI( _
    ByVal ws As Worksheet) As Long

    Dim sonSatir As Long

    sonSatir = _
    ws.Cells( _
    ws.Rows.Count, _
    "A").End(xlUp).Row

    If sonSatir < 7 Then

    BIR_V2_VERI_SATIRI_SAYISI = 0

    Else

    BIR_V2_VERI_SATIRI_SAYISI = _
        sonSatir - 6

    End If

    End Function

    Private Sub BIR_V2_KAYNAK_DOGRULA( _
    ByRef veri As Variant, _
    ByVal BeklenenUY As String, _
    ByVal kaynakAdi As String)

    Dim i As Long
    Dim uy As String

    If Not IsArray(veri) Then

    Err.Raise _
        vbObjectError + 3310, _
        "BIR_V2_KAYNAK_DOGRULA", _
        kaynakAdi & _
        " verisi dizi olarak okunamadý."

    End If

    For i = 1 To UBound(veri, 1)

    uy = _
        Trim$(CStr(veri(i, 1)))

    If Len(uy) > 0 Then

        If uy <> BeklenenUY Then

            Err.Raise _
                vbObjectError + 3311, _
                "BIR_V2_KAYNAK_DOGRULA", _
                kaynakAdi & _
                " içerisinde beklenmeyen ÜY bulundu: " & _
                uy

        End If

    End If

    Next i

    End Sub

    Private Function BIR_V2_ARRAY_BIRLESTIR( _
    ByRef Veri1 As Variant, _
    ByRef Veri2 As Variant) As Variant

    Dim sonuc() As Variant

    Dim Satir1 As Long
    Dim Satir2 As Long

    Dim i As Long
    Dim j As Long

    Satir1 = UBound(Veri1, 1)
    Satir2 = UBound(Veri2, 1)

    ReDim sonuc( _
    1 To Satir1 + Satir2, _
    1 To 12)

    For i = 1 To Satir1

    For j = 1 To 12

        sonuc(i, j) = _
            Veri1(i, j)

    Next j

    Next i

    For i = 1 To Satir2

    For j = 1 To 12

        sonuc( _
            Satir1 + i, _
            j) = _
            Veri2(i, j)

    Next j

    Next i

    BIR_V2_ARRAY_BIRLESTIR = sonuc

    End Function

    Private Sub BIR_V2_STOK_BASLIKLARI( _
    ByVal ws As Worksheet)

    ws.Range("A6:L6").Value = Array( _
    "ÜY", _
    "DpYr", _
    "Depo yeri tanýmý", _
    "Mal grubu", _
    "Mal grubu tanýmý", _
    "Parti", _
    "Malzeme", _
    "Malzeme kýsa metni", _
    "Tahditsiz", _
    "Birim", _
    "Deðer (TL)", _
    "Birim Fiyat (TL)")

    End Sub

    Private Sub BIR_V2_STOK_OZET( _
    ByVal ws As Worksheet, _
    ByVal ToplamSatir As Long)

    Dim sonSatir As Long

    sonSatir = _
    6 + ToplamSatir

    ws.Range("A2").Value = _
    "Son Güncelleme"

    ws.Range("A3").Value = _
    "Toplam Stok Satýrý"

    ws.Range("A4").Value = _
    "Toplam Stok Deðeri"

    On Error Resume Next

    ws.Range("C2:E2").UnMerge
    ws.Range("C3:E3").UnMerge
    ws.Range("C4:E4").UnMerge

    On Error GoTo 0

    ws.Range("C2:E2").Merge
    ws.Range("C3:E3").Merge
    ws.Range("C4:E4").Merge

    ws.Range("C2").Value = Now

    If ToplamSatir > 0 Then
        ws.Range("C3").Formula = _
            "=SUBTOTAL(103,A7:A" & sonSatir & ")"
        ws.Range("C4").Formula = _
            "=SUBTOTAL(109,K7:K" & sonSatir & ")"
    Else
        ws.Range("C3").Value = 0
        ws.Range("C4").Value = 0
    End If

    ws.Range("C2").NumberFormat = _
    "dd.mm.yyyy hh:mm:ss"

    ws.Range("C3").NumberFormat = _
    "#,##0"

    ws.Range("C4").NumberFormat = _
    "#,##0.00 ""TL"""

    End Sub

    Public Sub BIR_V2_RAPOR_OLUSTUR( _
    ByVal ws As Worksheet, _
    ByRef veri As Variant)

    Dim dict As Object

    Dim bilgi As Variant
    Dim anahtar As Variant

    Dim i As Long
    Dim HedefSatir As Long

    Dim depo As String
    Dim DepoTanimi As String
    Dim deger As Double

    Dim ToplamSatir As Long
    Dim toplamDeger As Double
    Dim TasarimSatir As Long

    Set dict = _
    CreateObject("Scripting.Dictionary")

    dict.CompareMode = vbTextCompare

    For i = 1 To UBound(veri, 1)

    depo = _
        Trim$(CStr(veri(i, 2)))

    DepoTanimi = _
        CStr(veri(i, 3))

    If Len(depo) > 0 Then

        deger = 0

        If IsNumeric(veri(i, 11)) Then
            deger = CDbl(veri(i, 11))
        End If

        If dict.Exists(depo) Then

            bilgi = dict(depo)

            bilgi(1) = _
                CLng(bilgi(1)) + 1

            bilgi(2) = _
                CDbl(bilgi(2)) + deger

            If Len(CStr(bilgi(0))) = 0 Then
                bilgi(0) = DepoTanimi
            End If

            dict(depo) = bilgi

        Else

            dict.Add _
                depo, _
                Array( _
                    DepoTanimi, _
                    1&, _
                    deger)

        End If

    End If

    Next i

    ws.Range( _
    "A9:D" & AKSORA_V13_SON(ws)).ClearContents

    ws.Range("A8").Value = _
    "Depo yeri"

    ws.Range("B8").Value = _
    "Depo yeri tanýmý"

    ws.Range("C8").Value = _
    "Stok Satýrý"

    ws.Range("D8").Value = _
    "Toplam Stok Deðeri"

    HedefSatir = 9
    ToplamSatir = 0
    toplamDeger = 0

    For Each anahtar In dict.keys

    bilgi = dict(anahtar)

    ws.Cells( _
        HedefSatir, _
        1).Value = _
        anahtar

    ws.Cells( _
        HedefSatir, _
        2).Value = _
        bilgi(0)

    ws.Cells( _
        HedefSatir, _
        3).Value = _
        CLng(bilgi(1))

    ws.Cells( _
        HedefSatir, _
        4).Value = _
        CDbl(bilgi(2))

    ToplamSatir = _
        ToplamSatir + CLng(bilgi(1))

    toplamDeger = _
        toplamDeger + CDbl(bilgi(2))

    HedefSatir = _
        HedefSatir + 1

    Next anahtar


    If HedefSatir > 9 Then
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add key:=ws.Range("A9:A" & HedefSatir - 1), _
                SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
            .SetRange ws.Range("A9:D" & HedefSatir - 1)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    ws.Cells( _
    HedefSatir, _
    1).Value = _
    "TOPLAM"

    ws.Cells( _
    HedefSatir, _
    3).Value = _
    ToplamSatir

    ws.Cells( _
    HedefSatir, _
    4).Value = _
    toplamDeger

    ws.Range( _
    "A" & HedefSatir & _
    ":D" & HedefSatir).Font.bold = True

    ws.Range("B3").Value = Now
    ws.Range("B4").Value = dict.Count
    ws.Range("B5").Value = toplamDeger

    ws.Range("B3").NumberFormat = _
    "dd.mm.yyyy hh:mm:ss"

    ws.Range("B4").NumberFormat = _
    "#,##0"

    ws.Range("B5").NumberFormat = _
    "#,##0.00 ""TL"""

    ws.Range( _
    "C9:C" & HedefSatir).NumberFormat = _
    "#,##0"

    ws.Range( _
    "D9:D" & HedefSatir).NumberFormat = _
    "#,##0.00 ""TL"""

    On Error Resume Next

    If ws.AutoFilterMode Then
    ws.AutoFilterMode = False
    End If

    On Error GoTo 0

    If HedefSatir > 9 Then
        ws.Range( _
            "A8:D" & HedefSatir - 1).AutoFilter
    Else
        ws.Range("A8:D8").AutoFilter
    End If

    With ws.Range("A8:D8")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
    End With

    For TasarimSatir = 9 To HedefSatir - 1

        With ws.Range("A" & TasarimSatir & ":D" & TasarimSatir)

            If (TasarimSatir - 9) Mod 2 = 0 Then
                .Interior.Color = RGB(242, 248, 252)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If

            .Font.Color = RGB(0, 0, 0)
            .Font.bold = False
            .Borders.LineStyle = xlNone

        End With

    Next TasarimSatir

    With ws.Range("A" & HedefSatir & ":D" & HedefSatir)
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(0, 0, 0)
        .Font.bold = True
        .Borders.LineStyle = xlNone
    End With

    ws.Range( _
        "A" & HedefSatir + 1 & _
        ":D" & AKSORA_V13_SON(ws)).ClearFormats

    ws.Columns("A").ColumnWidth = 14
    ws.Columns("B").ColumnWidth = 32
    ws.Columns("C").ColumnWidth = 16
    ws.Columns("D").ColumnWidth = 22

    ws.Rows("8:" & HedefSatir).RowHeight = 20
    ws.Range("A9:B" & HedefSatir).HorizontalAlignment = xlLeft
    ws.Range("C9:D" & HedefSatir).HorizontalAlignment = xlRight
    ws.Range("A8:D" & HedefSatir).VerticalAlignment = xlCenter

    End Sub

    Private Sub BIR_V2_STOK_GORSEL( _
    ByVal ws As Worksheet, _
    ByVal ToplamSatir As Long)

    Dim sonSatir As Long
    Dim TasarimSatir As Long

    sonSatir = _
    6 + ToplamSatir

    If ToplamSatir <= 0 Then Exit Sub

    ws.Range( _
    "I7:I" & sonSatir).NumberFormat = _
    "#,##0"

    ws.Range( _
    "K7:K" & sonSatir).NumberFormat = _
    "#,##0.00 ""TL"""

    ws.Range( _
    "L7:L" & sonSatir).NumberFormat = _
    "#,##0.00 ""TL"""

    On Error Resume Next

    If ws.AutoFilterMode Then
    ws.AutoFilterMode = False
    End If

    On Error GoTo 0

    ws.Range( _
    "A6:L" & sonSatir).AutoFilter

    With ws.Range("A6:L6")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ws.Rows("6:" & sonSatir).RowHeight = 20
    ws.Range("A7:H" & sonSatir).HorizontalAlignment = xlLeft
    ws.Range("I7:I" & sonSatir).HorizontalAlignment = xlRight
    ws.Range("J7:J" & sonSatir).HorizontalAlignment = xlCenter
    ws.Range("K7:K" & sonSatir).HorizontalAlignment = xlRight
    ws.Range("L7:L" & sonSatir).HorizontalAlignment = xlCenter
    ws.Range("A7:L" & sonSatir).VerticalAlignment = xlCenter

    For TasarimSatir = 7 To sonSatir

        With ws.Range("A" & TasarimSatir & ":L" & TasarimSatir)

            If (TasarimSatir - 7) Mod 2 = 0 Then
                .Interior.Color = RGB(242, 248, 252)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If

            .Font.Color = RGB(0, 0, 0)
            .Font.bold = False
            .Borders.LineStyle = xlNone

        End With

    Next TasarimSatir

    ' Eski veriden kalan koyu/tablolu biçimleri temizle.
    If sonSatir < ws.Rows.Count Then
        With ws.Range("A" & sonSatir + 1 & ":L" & ws.Rows.Count)
            .Borders.LineStyle = xlNone
            .Interior.Pattern = xlNone
            .Font.bold = False
        End With
    End If

    End Sub

    Private Function BIR_V2_TOPLAM_DEGER( _
    ByRef veri As Variant) As Double

    Dim i As Long
    Dim sonuc As Double

    sonuc = 0

    For i = 1 To UBound(veri, 1)

    If IsNumeric(veri(i, 11)) Then

        sonuc = _
            sonuc + CDbl(veri(i, 11))

    End If

    Next i

    BIR_V2_TOPLAM_DEGER = sonuc

    End Function


