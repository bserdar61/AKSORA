Attribute VB_Name = "modCDP"
Option Explicit

    '==========================================================
    ' AKSORA
    ' modCDP
    '
    ' CDP NORMAL STOK MOTORU
    '==========================================================

    Private Const CDP_SISTEM As String = "CDP"
    Private Const CDP_DEPO_ILK As String = "D001"
    Private Const CDP_DEPO_SON As String = "D999"

    Private Const CDP_STOK_SAYFA As String = "CDP_STOK"
    Private Const CDP_RAPOR_SAYFA As String = "CDP_RAPOR"

    Public Sub CDP_STOK_GUNCELLE_V2()

    Dim wsStok As Worksheet
    Dim wsRapor As Worksheet

    Dim YeniVeri As Variant
    Dim EskiVeri As Variant

    Dim YeniSatirSayisi As Long
    Dim YeniSutunSayisi As Long

    Dim EskiSonSatir As Long
    Dim EskiSonSutun As Long

    Dim ExportSuresi As Double
    Dim ExcelBaslangic As Double
    Dim ExcelSuresi As Double
    Dim ToplamSure As Double

    Dim hataMesaji As String
    Dim hataNo As Long
    Dim basarili As Boolean

    Dim yazmaBasladi As Boolean
    Dim YedekVar As Boolean
    
    If Not AKSORA_SAP_GUNCELLEME_IZNI() Then Exit Sub

    On Error GoTo Hata

    Set wsStok = _
    ThisWorkbook.Worksheets(CDP_STOK_SAYFA)

    Set wsRapor = _
    ThisWorkbook.Worksheets(CDP_RAPOR_SAYFA)

    CORE_ISLEM_BASLAT _
    "CDP stoklarý güncelleniyor"

    CORE_DURUM _
    "CDP yeni stok verisi SAP'tan alýnýyor..."

    ' 30'lu ALV blok okuma. Hata olursa mevcut guvenli export'a don.
    basarili = CDP_V2_30_BLOK_AL( _
        YeniVeri, YeniSatirSayisi, YeniSutunSayisi, _
        ExportSuresi, hataMesaji)

    If Not basarili Then
        CORE_DURUM "30'lu okuma tamamlanamadi; mevcut CDP export deneniyor..."
        basarili = STOK_EXPORT_AL( _
            CDP_SISTEM, CDP_UY, CDP_DEPO_ILK, CDP_DEPO_SON, _
            YeniVeri, YeniSatirSayisi, YeniSutunSayisi, _
            ExportSuresi, hataMesaji)
    End If

    If Not basarili Then

    ToplamSure = CORE_ISLEM_BITIR()

    If Not ANA_SESSIZ_MOD Then

        MsgBox _
            "CDP stok güncellemesi baþlatýlamadý." & _
            vbCrLf & vbCrLf & _
            hataMesaji & _
            vbCrLf & vbCrLf & _
            "Mevcut CDP stok verileri deðiþtirilmedi.", _
            vbCritical, _
            "AKSORA - CDP"

    End If

    Exit Sub

    End If

    If Not IsArray(YeniVeri) Then

    Err.Raise _
        vbObjectError + 3101, _
        "CDP_STOK_GUNCELLE_V2", _
        "SAP verisi geçerli bir dizi olarak alýnamadý."

    End If

    If YeniSatirSayisi <= 0 Then

    Err.Raise _
        vbObjectError + 3102, _
        "CDP_STOK_GUNCELLE_V2", _
        "SAP'tan stok satýrý gelmedi."

    End If

    If YeniSutunSayisi < 12 Then

    Err.Raise _
        vbObjectError + 3103, _
        "CDP_STOK_GUNCELLE_V2", _
        "CDP export yapýsý beklenen 12 sütundan daha az."

    End If

    CORE_DURUM _
    "Mevcut CDP stok verisi güvenliðe alýnýyor..."

    EskiSonSatir = _
    STOK_SON_SATIR(wsStok)

    EskiSonSutun = _
    STOK_SON_SUTUN(wsStok)

    If EskiSonSatir >= 7 And _
    EskiSonSutun >= 1 Then

    If EskiSonSutun > 12 Then
        EskiSonSutun = 12
    End If

    EskiVeri = _
        wsStok.Range( _
            wsStok.Cells(7, 1), _
            wsStok.Cells( _
                EskiSonSatir, _
                EskiSonSutun)).Value2

    YedekVar = True

    Else

    YedekVar = False

    End If

    ExcelBaslangic = Timer

    CORE_DURUM _
    "CDP stok verileri Excel'e yazýlýyor..."

    yazmaBasladi = True

    CDP_V2_BASLIKLARI_YAZ wsStok

    wsStok.Range( _
    "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

    wsStok.Range( _
    "A7").Resize( _
    YeniSatirSayisi, _
    12).Value2 = _
    CDP_V2_ILK_12_SUTUN(YeniVeri)

    CORE_DURUM _
    "CDP stok özeti hesaplanýyor..."

    CDP_V2_STOK_OZET_GUNCELLE _
    wsStok, _
    YeniSatirSayisi

    CORE_DURUM _
    "CDP depo raporu hazýrlanýyor..."

    CDP_V2_RAPOR_OLUSTUR _
    wsStok, _
    wsRapor, _
    YeniSatirSayisi

    CDP_V2_DEPO_LISTESI wsStok

    CORE_DURUM _
    "CDP görünümü düzenleniyor..."

    CDP_V2_GORSEL_DUZEN _
    wsStok, _
    YeniSatirSayisi

    STOK_GORUNUM_TEK "CDP_STOK"
    RAPOR_GORUNUM_TEK "CDP_RAPOR"

    ExcelSuresi = _
    STOK_GECEN_SURE(ExcelBaslangic)

    ToplamSure = _
    CORE_ISLEM_BITIR()

    application.Calculation = xlCalculationAutomatic
    application.CalculateFull

    If Not ANA_SESSIZ_MOD Then

    MsgBox _
        "CDP stok güncellemesi baþarýyla tamamlandý." & _
        vbCrLf & vbCrLf & _
        "Üretim Yeri : " & CDP_UY & vbCrLf & _
        "Depo Aralýðý: " & _
        CDP_DEPO_ILK & " - " & CDP_DEPO_SON & vbCrLf & _
        "Stok Satýrý : " & _
        AKSORA_SAYI_METNI(YeniSatirSayisi, "#,##0") & vbCrLf & _
        "SAP + Export : " & _
        AKSORA_SAYI_METNI(ExportSuresi, "0.00") & " sn" & vbCrLf & _
        "Excel Ýþleme: " & _
        AKSORA_SAYI_METNI(ExcelSuresi, "0.00") & " sn" & vbCrLf & _
        "Toplam Süre : " & _
        AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn", _
        vbInformation, _
        "AKSORA - CDP"

    End If

    Exit Sub

Hata:

    hataNo = Err.Number
    hataMesaji = Err.Description

    On Error Resume Next

    If yazmaBasladi Then

    CORE_DURUM _
        "Hata oluþtu. Eski CDP stok verisi geri yükleniyor..."

    wsStok.Range( _
        "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

    If YedekVar Then

        If IsArray(EskiVeri) Then

            wsStok.Range("A7").Resize( _
                UBound(EskiVeri, 1), _
                UBound(EskiVeri, 2)).Value2 = _
                EskiVeri

        End If

    End If

    End If

    On Error GoTo 0

    ToplamSure = _
    CORE_ISLEM_BITIR()

    If Not ANA_SESSIZ_MOD Then

    MsgBox _
        "CDP stok güncellemesi sýrasýnda hata oluþtu." & _
        vbCrLf & vbCrLf & _
        "Hata No : " & hataNo & vbCrLf & _
        "Açýklama: " & hataMesaji & _
        vbCrLf & vbCrLf & _
        "Yeni veri tamamlanamadý." & vbCrLf & _
        "Eski stok verisi korunmaya çalýþýldý.", _
        vbCritical, _
        "AKSORA - CDP"

    End If
    
If ANA_SESSIZ_MOD Then
    Err.Raise hataNo, "CDP_STOK_GUNCELLE_V2", hataMesaji
End If

    End Sub

    Private Function CDP_V2_ILK_12_SUTUN( _
    ByRef kaynak As Variant) As Variant

    Dim sonuc() As Variant

    Dim satir As Long
    Dim sutun As Long

    Dim SatirSayisi As Long
    Dim KaynakSutunSayisi As Long

    SatirSayisi = _
    UBound(kaynak, 1)

    KaynakSutunSayisi = _
    UBound(kaynak, 2)

    If KaynakSutunSayisi < 12 Then

    Err.Raise _
        vbObjectError + 3110, _
        "CDP_V2_ILK_12_SUTUN", _
        "Kaynak veri 12 sütundan az."

    End If

    ReDim sonuc( _
    1 To SatirSayisi, _
    1 To 12)

    For satir = 1 To SatirSayisi

    For sutun = 1 To 12

        sonuc(satir, sutun) = _
            kaynak(satir, sutun)

    Next sutun

    Next satir

    CDP_V2_ILK_12_SUTUN = sonuc

    End Function

    Private Sub CDP_V2_BASLIKLARI_YAZ( _
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

    Private Sub CDP_V2_STOK_OZET_GUNCELLE( _
    ByVal ws As Worksheet, _
    ByVal VeriSatiri As Long)

    Dim sonSatir As Long
    Dim toplamDeger As Double

    sonSatir = _
    6 + VeriSatiri

    toplamDeger = 0

    If VeriSatiri > 0 Then

    On Error Resume Next

    toplamDeger = _
        application.WorksheetFunction.Sum( _
            ws.Range( _
                "K7:K" & sonSatir))

    On Error GoTo 0

    End If

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

    ' KPI deðerleri AutoFilter sonucuna göre otomatik deðiþir.
    ws.Range("C3").Formula = _
        "=SUBTOTAL(103,A7:A" & sonSatir & ")"

    ws.Range("C4").Formula = _
        "=SUBTOTAL(109,K7:K" & sonSatir & ")"

    ws.Range("C2").NumberFormat = _
    "dd.mm.yyyy hh:mm:ss"

    ws.Range("C3").NumberFormat = _
    "#,##0"

    ws.Range("C4").NumberFormat = _
    "#,##0.00 ""TL"""

    End Sub

    Public Sub CDP_V2_RAPOR_OLUSTUR( _
    ByVal wsStok As Worksheet, _
    ByVal wsRapor As Worksheet, _
    ByVal VeriSatiri As Long)

    Dim dict As Object

    Dim veri As Variant
    Dim bilgi As Variant

    Dim i As Long
    Dim HedefSatir As Long
    Dim SonStokSatir As Long

    Dim depo As String
    Dim DepoTanimi As String
    Dim deger As Double

    Dim anahtar As Variant

    Dim ToplamSatir As Long
    Dim toplamDeger As Double
    Dim TasarimSatir As Long

    Set dict = _
    CreateObject("Scripting.Dictionary")

    dict.CompareMode = vbTextCompare

    SonStokSatir = _
    6 + VeriSatiri

    If VeriSatiri > 0 Then

    veri = _
        wsStok.Range( _
            "A7:K" & SonStokSatir).Value2

    For i = 1 To UBound(veri, 1)

        If Trim$(CStr(veri(i, 1))) = CDP_UY Then

            depo = _
                Trim$(CStr(veri(i, 2)))

            DepoTanimi = _
                CStr(veri(i, 3))

            If Len(depo) > 0 Then

                deger = 0

                If IsNumeric(veri(i, 11)) Then
                    deger = _
                        CDbl(veri(i, 11))
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

        End If

    Next i

    End If

    wsRapor.Range( _
    "A9:D" & AKSORA_V13_SON(wsRapor)).ClearContents

    wsRapor.Range("A8").Value = _
    "Depo yeri"

    wsRapor.Range("B8").Value = _
    "Depo yeri tanýmý"

    wsRapor.Range("C8").Value = _
    "Stok Satýrý"

    wsRapor.Range("D8").Value = _
    "Toplam Stok Deðeri"

    HedefSatir = 9
    ToplamSatir = 0
    toplamDeger = 0

    For Each anahtar In dict.keys

    bilgi = dict(anahtar)

    wsRapor.Cells( _
        HedefSatir, 1).Value = _
        anahtar

    wsRapor.Cells( _
        HedefSatir, 2).Value = _
        bilgi(0)

    wsRapor.Cells( _
        HedefSatir, 3).Value = _
        CLng(bilgi(1))

    wsRapor.Cells( _
        HedefSatir, 4).Value = _
        CDbl(bilgi(2))

   ToplamSatir = ToplamSatir + CLng(bilgi(1))
toplamDeger = toplamDeger + CDbl(bilgi(2))
HedefSatir = HedefSatir + 1

    Next anahtar

    If HedefSatir > 10 Then
        With wsRapor.Sort
            .SortFields.Clear
            .SortFields.Add _
                key:=wsRapor.Range("A9:A" & HedefSatir - 1), _
                SortOn:=xlSortOnValues, _
                Order:=xlAscending, _
                DataOption:=xlSortNormal
            .SetRange wsRapor.Range("A9:D" & HedefSatir - 1)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    wsRapor.Cells( _
    HedefSatir, 1).Value = _
    "TOPLAM"

    wsRapor.Cells( _
    HedefSatir, 3).Value = _
    ToplamSatir

    wsRapor.Cells( _
    HedefSatir, 4).Value = _
    toplamDeger

    wsRapor.Range( _
    "A" & HedefSatir & _
    ":D" & HedefSatir).Font.bold = True

    wsRapor.Range("B3").Value = Now
    wsRapor.Range("B4").Value = dict.Count
    wsRapor.Range("B5").Value = toplamDeger

    wsRapor.Range("B3").NumberFormat = _
    "dd.mm.yyyy hh:mm:ss"

    wsRapor.Range("B4").NumberFormat = _
    "#,##0"

    wsRapor.Range("B5").NumberFormat = _
    "#,##0.00 ""TL"""

    If HedefSatir >= 9 Then

    wsRapor.Range( _
        "C9:C" & HedefSatir).NumberFormat = _
        "#,##0"

    wsRapor.Range( _
        "D9:D" & HedefSatir).NumberFormat = _
        "#,##0.00 ""TL"""

    End If

    On Error Resume Next

    If wsRapor.AutoFilterMode Then
    wsRapor.AutoFilterMode = False
    End If

    On Error GoTo 0

    If HedefSatir > 9 Then
        wsRapor.Range( _
            "A8:D" & HedefSatir - 1).AutoFilter
    Else
        wsRapor.Range("A8:D8").AutoFilter
    End If

    '==========================================================
    ' CDP_RAPOR NÝHAÝ TABLO TASARIMI
    ' EPP satýr renkleri + ÝLÇE dinamik TOPLAM bitiþi
    '==========================================================

    With wsRapor.Range("A8:D8")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
    End With

    For TasarimSatir = 9 To HedefSatir - 1
        With wsRapor.Range("A" & TasarimSatir & ":D" & TasarimSatir)
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

    With wsRapor.Range("A" & HedefSatir & ":D" & HedefSatir)
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(0, 0, 0)
        .Font.bold = True
        .Borders.LineStyle = xlNone
    End With

    wsRapor.Range( _
        "A" & HedefSatir + 1 & _
        ":D" & AKSORA_V13_SON(wsRapor)).ClearFormats

    wsRapor.Columns("A").ColumnWidth = 14
    wsRapor.Columns("B").ColumnWidth = 32
    wsRapor.Columns("C").ColumnWidth = 16
    wsRapor.Columns("D").ColumnWidth = 22

    End Sub

    Private Sub CDP_V2_DEPO_LISTESI( _
    ByVal ws As Worksheet)

    Dim MevcutSecim As String
    Dim SonVeriSatiri As Long
    Dim SonListeSatiri As Long
    Dim i As Long

    Dim DepoTanimi As String
    Dim dict As Object
    Dim anahtar As Variant

    Dim nm As Name
    Dim BulunanAd As Name
    Dim KisaAd As String
    Dim RefersToMetni As String

    MevcutSecim = Trim$(CStr(ws.Range("H2").Value))

    If ws.Range("G2").MergeCells Then
        ws.Range("G2").MergeArea.UnMerge
    End If

    If ws.Range("H2").MergeCells Then
        ws.Range("H2").MergeArea.UnMerge
    End If

    ws.Range("G2").Value = "DEPOLAR"

    ws.Range("N2:N" & AKSORA_V13_SON(ws)).ClearContents

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    SonVeriSatiri = 6 + ANA_TUM_LONG_CDP(ws.Range("C3"))

    If SonVeriSatiri >= 7 Then

        For i = 7 To SonVeriSatiri

            DepoTanimi = Trim$(CStr(ws.Cells(i, 3).Value2))

            If Len(DepoTanimi) > 0 Then
                If Not dict.Exists(DepoTanimi) Then
                    dict.Add DepoTanimi, DepoTanimi
                End If
            End If

        Next i

    End If

    ws.Range("N2").Value = "TÜMÜ"
    SonListeSatiri = 2

    For Each anahtar In dict.keys
        SonListeSatiri = SonListeSatiri + 1
        ws.Cells(SonListeSatiri, "N").Value = CStr(anahtar)
    Next anahtar

    Set BulunanAd = Nothing

    For Each nm In ThisWorkbook.names

        KisaAd = nm.Name

        If InStrRev(KisaAd, "!") > 0 Then
            KisaAd = Mid$(KisaAd, InStrRev(KisaAd, "!") + 1)
        End If

        KisaAd = Replace(KisaAd, "'", "")

        If StrComp(KisaAd, "_CDP_DEPO_LISTESI", vbTextCompare) = 0 Then
            Set BulunanAd = nm
            Exit For
        End If

    Next nm

    RefersToMetni = _
    "='" & Replace(ws.Name, "'", "''") & _
    "'!$N$2:$N$" & SonListeSatiri

    If BulunanAd Is Nothing Then
        ThisWorkbook.names.Add _
            Name:="_CDP_DEPO_LISTESI", _
            RefersTo:=RefersToMetni
    Else
        BulunanAd.RefersTo = RefersToMetni
    End If

    On Error Resume Next
    ws.Range("H2").Validation.Delete
    On Error GoTo 0

    ws.Range("H2").Validation.Add _
    Type:=xlValidateList, _
    AlertStyle:=xlValidAlertStop, _
    Operator:=xlBetween, _
    Formula1:="=_CDP_DEPO_LISTESI"

    With ws.Range("H2").Validation
        .IgnoreBlank = True
        .InCellDropdown = True
        .ShowError = True
    End With

    If StrComp(MevcutSecim, "TÜMÜ", vbTextCompare) = 0 Then
        ws.Range("H2").Value = "TÜMÜ"
    ElseIf Len(MevcutSecim) > 0 And dict.Exists(MevcutSecim) Then
        ws.Range("H2").Value = MevcutSecim
    Else
        ws.Range("H2").Value = "TÜMÜ"
    End If

    ws.Range("G2").Font.bold = True
    ws.Range("G2").HorizontalAlignment = xlCenter
    ws.Range("G2").VerticalAlignment = xlCenter

    ws.Range("H2").HorizontalAlignment = xlCenter
    ws.Range("H2").VerticalAlignment = xlCenter

    ws.Columns("N").Hidden = True

    End Sub

    Private Function ANA_TUM_LONG_CDP(ByVal hucre As Range) As Long

    If IsNumeric(hucre.Value2) Then
        ANA_TUM_LONG_CDP = CLng(hucre.Value2)
    Else
        ANA_TUM_LONG_CDP = 0
    End If

    End Function

    Private Sub CDP_V2_GORSEL_DUZEN( _
    ByVal ws As Worksheet, _
    ByVal VeriSatiri As Long)

    Dim sonSatir As Long
    Dim TasarimSatir As Long
    Dim KullanilanSonSatir As Long

    sonSatir = _
    6 + VeriSatiri

    If VeriSatiri <= 0 Then Exit Sub

    KullanilanSonSatir = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

    If KullanilanSonSatir > sonSatir Then
        ws.Range("A" & sonSatir + 1 & ":L" & KullanilanSonSatir).ClearFormats
    End If

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
        .Borders.LineStyle = xlNone
    End With

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
    
     End Sub

'==========================================================
' CDP 30 SATIRLIK BLOK OKUMA - TEST EDILEN YONTEM
'==========================================================
    Public Function CDP_V2_30_BLOK_AL(ByRef VeriDizisi As Variant, _
    ByRef VeriSatiri As Long, ByRef SutunSayisi As Long, _
    ByRef ExportSuresi As Double, ByRef hataMesaji As String) As Boolean
    Dim uy As String
    Dim blok As Long
    Dim ses As Object, alv As Object, duzen As Object, kolon As Object
    Dim ids(1 To 12) As String, ham() As Variant
    Dim ws As Worksheet, i As Long, j As Long, b As Long, son As Long
    Dim n As Long, bulunan As Long, tekrar As Long, okumaHatasi As Boolean
    Dim t As Double, sure As Double, toplam As Double, mevcut As Double
    Dim sayi As Long, mevcutSayi As Long, ad As String, tanim As String
    Dim hataMetni As String, kaydirmaHatasi As String, ozetSatiri As Boolean
    Dim ozetAdedi As Long, veriBitti As Boolean, eskiEkran As Boolean

    On Error GoTo Hata
    uy = CDP_UY()
    Set ws = ThisWorkbook.Worksheets("CDP_STOK")
    mevcutSayi = application.WorksheetFunction.CountA(ws.Range("A7:A" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row))
    mevcut = application.WorksheetFunction.Sum(ws.Range("K7:K" & ws.Cells(ws.Rows.Count, 1).End(xlUp).Row))
    t = Timer

    Set ses = SAP_SESSION_AL("CDP", hataMetni)
    If ses Is Nothing Then Err.Raise vbObjectError + 801, , "CDP SAP oturumu bulunamadi: " & hataMetni
    ses.StartTransaction "MB52"
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = uy
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = "D001"
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = "D999"
    ses.FindById("wnd[0]").SendVKey 8
    CDP_TEST_BEKLE 0.4
    ses.FindById("wnd[0]/tbar[1]/btn[33]").Press
    CDP_TEST_BEKLE 0.3

    Set duzen = ses.FindById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    bulunan = -1
    For i = 0 To CLng(duzen.RowCount) - 1
        ad = Trim$(CStr(duzen.GetCellValue(i, "VARIANT")))
        tanim = Trim$(CStr(duzen.GetCellValue(i, "TEXT")))
        If StrComp(ad, "PANEL", vbTextCompare) = 0 Or _
           StrComp(tanim, "MLZ STOK PANEL", vbTextCompare) = 0 Then
            bulunan = i
            Exit For
        End If
    Next i
    If bulunan < 0 Then Err.Raise vbObjectError + 802, , "PANEL duzeni bulunamadi."
    duzen.SetCurrentCell bulunan, "TEXT"
    duzen.SelectedRows = CStr(bulunan)
    duzen.ClickCurrentCell
    CDP_TEST_BEKLE 0.3

    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    On Error Resume Next
    alv.PressToolbarButton "&DELETE_FILTER"
    Err.Clear
    On Error GoTo Hata
    Set kolon = alv.ColumnOrder
    If CLng(kolon.Count) < 12 Then Err.Raise vbObjectError + 803, , "PANEL duzeninde 12 sutun yok."
    For j = 1 To 12
        ids(j) = CStr(kolon.Item(j - 1))
    Next j

    n = CLng(alv.RowCount)
    If n <= 0 Then Err.Raise vbObjectError + 804, , "SAP MB52 sonucu bos."
    ReDim ham(1 To n, 1 To 12)

    blok = AKSORA_SAP_BLOK_BOYU(alv)
    For b = 0 To n - 1 Step blok
        son = b + blok - 1
        If son >= n Then son = n - 1

        On Error Resume Next
        Err.Clear
        alv.FirstVisibleRow = b
        kaydirmaHatasi = Err.Description
        okumaHatasi = Err.Number <> 0
        On Error GoTo Hata
        If okumaHatasi Then Err.Raise vbObjectError + 805, , "Blok " & b + 1 & " kaydirilamadi: " & kaydirmaHatasi
        DoEvents
        CDP_TEST_BEKLE 0.015

        AKSORA_TOPLU_SURE_TIK
        For i = b To son
            okumaHatasi = False
            On Error Resume Next
            For j = 1 To 12
                Err.Clear
                ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                If Err.Number <> 0 Then okumaHatasi = True: Exit For
            Next j
            On Error GoTo Hata

            If Not okumaHatasi Then
                If CDP_TEST_GECERLI(ham, i + 1, uy) Then GoTo SONRAKI_SATIR
                If CDP_TEST_OZET(ham, i + 1, uy) Then
                    ' Ilk ozet satirindan sonra stok verisi gelmemeli.
                    veriBitti = True
                    ozetAdedi = ozetAdedi + 1
                    GoTo SONRAKI_SATIR
                End If
            End If

            ' Gorunmeyen gercek stok satirlari icin 2 kez kaydirarak yeniden oku.
            For tekrar = 1 To 2
                okumaHatasi = False
                On Error Resume Next
                Err.Clear
                alv.FirstVisibleRow = i
                alv.SetCurrentCell i, ids(7)
                okumaHatasi = Err.Number <> 0
                On Error GoTo Hata
                If Not okumaHatasi Then
                    DoEvents
                    CDP_TEST_BEKLE 0.015
                    On Error Resume Next
                    For j = 1 To 12
                        Err.Clear
                        ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                        If Err.Number <> 0 Then okumaHatasi = True: Exit For
                    Next j
                    On Error GoTo Hata
                    If Not okumaHatasi Then
                        If CDP_TEST_GECERLI(ham, i + 1, uy) Then Exit For
                        If CDP_TEST_OZET(ham, i + 1, uy) Then Exit For
                    End If
                End If
            Next tekrar

            If Not okumaHatasi Then
                If CDP_TEST_GECERLI(ham, i + 1, uy) Then GoTo SONRAKI_SATIR
                If CDP_TEST_OZET(ham, i + 1, uy) Then
                    veriBitti = True
                    ozetAdedi = ozetAdedi + 1
                    GoTo SONRAKI_SATIR
                End If
            End If

            ' SAP ozet satirinda SetCurrentCell hata verebilir: yalnizca
            ' gercek stoklar eksiksiz okunmussa ve ALV'nin en sonundaysak gec.
            If i >= n - 10 And sayi = mevcutSayi And AKSORA_AKTIF_IL() = AKSORA_CANLI_STOK_IL() And _
               (veriBitti Or i = mevcutSayi) Then
                veriBitti = True
                ozetAdedi = ozetAdedi + 1
                GoTo SONRAKI_SATIR
            End If
            Err.Raise vbObjectError + 806, , "Satir " & (i + 1) & " dogrulanamadi; test durduruldu."

SONRAKI_SATIR:
            If CDP_TEST_GECERLI(ham, i + 1, uy) Then
                If veriBitti Then Err.Raise vbObjectError + 807, , _
                    "Ozet satirindan sonra stok satiri bulundu: " & (i + 1)
                ad = UCase$(Trim$(CStr(ham(i + 1, 2))))
                If Len(ad) = 4 And Left$(ad, 1) = "D" And ad >= "D001" And ad <= "D999" Then
                    sayi = sayi + 1
                    toplam = toplam + CDP_TEST_SAYI(ham(i + 1, 11))
                End If
            End If
        Next i
    Next b

    If sayi <= 0 Then Err.Raise vbObjectError + 810, , "CDP stok satiri bulunamadi."
    ' Onceki stokla satir ve deger birebir ayniysa ilave tutarlilik kontrolu.
    ' Stok degismisse eski degerle eslesme zorunlulugu yoktur.
    Dim sonuc() As Variant, rr As Long, cc As Long, k As Long
    ReDim sonuc(1 To sayi, 1 To 12)
    k = 0
    For rr = 1 To n
        If CDP_TEST_GECERLI(ham, rr, uy) Then
            ad = UCase$(Trim$(CStr(ham(rr, 2))))
            If Len(ad) = 4 And Left$(ad, 1) = "D" And _
               ad >= "D001" And ad <= "D999" Then
                k = k + 1
                For cc = 1 To 12
                    sonuc(k, cc) = ham(rr, cc)
                Next cc
                sonuc(k, 9) = CDP_TEST_SAYI(sonuc(k, 9))
                sonuc(k, 11) = CDP_TEST_SAYI(sonuc(k, 11))
                sonuc(k, 12) = CDP_TEST_SAYI(sonuc(k, 12))
            End If
        End If
    Next rr
    If k <> sayi Then Err.Raise vbObjectError + 811, , "Okunan satir sayisi tutarsiz."
    VeriDizisi = sonuc
    VeriSatiri = sayi
    SutunSayisi = 12
    ExportSuresi = Timer - t
    If ExportSuresi < 0 Then ExportSuresi = ExportSuresi + 86400
    CDP_V2_30_BLOK_AL = True
    Exit Function
Hata:
    hataMesaji = "CDP 30'lu blok: " & Err.Number & " - " & Err.Description
    CDP_V2_30_BLOK_AL = False
End Function


Private Function CDP_TEST_GECERLI(ByRef ham() As Variant, ByVal satir As Long, ByVal uy As String) As Boolean
    CDP_TEST_GECERLI = (Trim$(CStr(ham(satir, 1))) = uy And _
                       Len(Trim$(CStr(ham(satir, 7)))) > 0)
End Function

Private Function CDP_TEST_OZET(ByRef ham() As Variant, ByVal satir As Long, ByVal uy As String) As Boolean
    Dim ilk As String, malzeme As String, c As Long, hucre As String
    ilk = Trim$(CStr(ham(satir, 1)))
    malzeme = Trim$(CStr(ham(satir, 7)))
    If ilk = uy And Len(malzeme) > 0 Then Exit Function
    ' MB52 ozetlerinde UY ve malzeme numarasi bos; birim veya toplam gorunur.
    If Len(ilk) = 0 And Len(malzeme) = 0 Then
        For c = 8 To 12
            hucre = UCase$(Trim$(CStr(ham(satir, c))))
            If Len(hucre) > 0 Then
                CDP_TEST_OZET = True
                Exit Function
            End If
        Next c
    End If
End Function

Private Sub CDP_TEST_BEKLE(ByVal saniye As Double)
    Dim bas As Double, gecen As Double
    bas = Timer
    Do
        DoEvents
        gecen = Timer - bas
        If gecen < 0 Then gecen = gecen + 86400
    Loop While gecen < saniye
End Sub

Private Function CDP_TEST_SAYI(ByVal deger As Variant) As Double
    Dim s As String, negatif As Boolean
    If IsNumeric(deger) Then
        CDP_TEST_SAYI = CDbl(deger)
        Exit Function
    End If
    s = Trim$(CStr(deger))
    If Len(s) = 0 Or s = "-" Then Exit Function
    If Left$(s, 1) = "-" Then
        negatif = True
        s = Mid$(s, 2)
    ElseIf Right$(s, 1) = "-" Then
        negatif = True
        s = Left$(s, Len(s) - 1)
    End If
    s = Replace(s, " ", vbNullString)
    If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
        If InStrRev(s, ",") > InStrRev(s, ".") Then
            s = Replace(s, ".", vbNullString)
            s = Replace(s, ",", application.DecimalSeparator)
        Else
            s = Replace(s, ",", vbNullString)
            s = Replace(s, ".", application.DecimalSeparator)
        End If
    ElseIf InStr(s, ",") > 0 Then
        s = Replace(s, ",", application.DecimalSeparator)
    ElseIf InStr(s, ".") > 0 Then
        s = Replace(s, ".", application.DecimalSeparator)
    End If
    If IsNumeric(s) Then
        CDP_TEST_SAYI = CDbl(s)
        If negatif Then CDP_TEST_SAYI = -CDP_TEST_SAYI
    End If
End Function
Private Function CDP_UY() As String
    CDP_UY = AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "CDP")
End Function

