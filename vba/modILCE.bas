Attribute VB_Name = "modILCE"
Option Explicit

    '==========================================================
    ' AKSORA
    ' modILCE
    '
    ' ÝLÇE STOK MOTORU
    '
    ' CDP : 7155 / K001 - K999
    ' EPP : 7055 / K001 - K999
    '
    ' AKIÞ:
    ' 1. CDP verisini RAM'e al
    ' 2. EPP verisini RAM'e al
    ' 3. Her iki kaynaðý doðrula
    ' 4. Mevcut ÝLÇE_STOK verisini RAM'de yedekle
    ' 5. CDP + EPP verilerini RAM'de birleþtir
    ' 6. ÝLÇE_STOK'a tek blokta yaz
    ' 7. ÝLÇE_RAPORLAR'ý A:E olarak yeniden oluþtur
    ' 8. H2 ilçe/depo listesini yenile
    ' 9. Hata halinde eski stok verisini geri yükle
    '==========================================================

    Private Const ILCE_CDP_SISTEM As String = "CDP"
    

    Private Const ILCE_EPP_SISTEM As String = "EPP"
    

    Private Const ILCE_DEPO_ILK As String = "K001"
    Private Const ILCE_DEPO_SON As String = "K999"

    Private Const ILCE_STOK_SAYFA As String = "ÝLÇE_STOK"
    Private Const ILCE_RAPOR_SAYFA As String = "ÝLÇE_RAPORLAR"

    '==========================================================
    ' ANA ÝLÇE V2 GÜNCELLEME
    '==========================================================
    Public Sub ILCE_STOK_GUNCELLE_V2()

    Dim islemIl As String
    Dim beklenenCDP As String, beklenenEPP As String

    Dim wsStok As Worksheet
    Dim wsRapor As Worksheet

    Dim CDPVeri As Variant
    Dim EPPVeri As Variant
    Dim BirlesikVeri As Variant
    Dim EskiVeri As Variant

    Dim CDPSatir As Long
    Dim CDPSutun As Long

    Dim EPPSatir As Long
    Dim EPPSutun As Long

    Dim ToplamSatir As Long

    Dim EskiSonSatir As Long
    Dim EskiSonSutun As Long

    Dim CDPExportSuresi As Double
    Dim EPPExportSuresi As Double

    Dim ExcelBaslangic As Double
    Dim ExcelSuresi As Double
    Dim ToplamSure As Double

    Dim hataMesaji As String
    Dim basarili As Boolean

    Dim yazmaBasladi As Boolean
    Dim YedekVar As Boolean

    Dim hataNo As Long
    Dim hataAciklama As String
    
    If Not AKSORA_SAP_GUNCELLEME_IZNI() Then Exit Sub

    On Error GoTo Hata

    islemIl = AKSORA_AKTIF_IL()
    beklenenCDP = AKSORA_IL_TESIS(islemIl, "CDP")
    beklenenEPP = AKSORA_IL_TESIS(islemIl, "EPP")

    Set wsStok = ThisWorkbook.Worksheets(ILCE_STOK_SAYFA)
    Set wsRapor = ThisWorkbook.Worksheets(ILCE_RAPOR_SAYFA)

    CORE_ISLEM_BASLAT _
    "Ýlçe stoklarý güncelleniyor"

    '==========================================================
    ' 1. CDP / 7155 / K001-K999
    '==========================================================

    CORE_DURUM _
    "[1/4] CDP ilçe stoklarý SAP'tan alýnýyor..."

    basarili = ILCE_CDP_BLOK_AL(CDPVeri, CDPSatir, CDPSutun, _
        CDPExportSuresi, hataMesaji)
    If Not basarili Then
        hataMesaji = vbNullString
        basarili = STOK_EXPORT_AL(ILCE_CDP_SISTEM, ILCE_CDP_UY, _
            ILCE_DEPO_ILK, ILCE_DEPO_SON, CDPVeri, CDPSatir, _
            CDPSutun, CDPExportSuresi, hataMesaji)
    End If

    If Not basarili Then

    Err.Raise _
        vbObjectError + 4101, _
        "ILCE_STOK_GUNCELLE_V2", _
        "CDP ilçe stoklarý alýnamadý." & _
        vbCrLf & hataMesaji

    End If

    If Not IsArray(CDPVeri) Then

    Err.Raise _
        vbObjectError + 4102, _
        "ILCE_STOK_GUNCELLE_V2", _
        "CDP ilçe verisi geçerli bir dizi deðil."

    End If

    If CDPSatir <= 0 Then

    Err.Raise _
        vbObjectError + 4103, _
        "ILCE_STOK_GUNCELLE_V2", _
        "CDP sisteminden K001-K999 aralýðýnda stok satýrý gelmedi."

    End If

    If CDPSutun < 12 Then

    Err.Raise _
        vbObjectError + 4104, _
        "ILCE_STOK_GUNCELLE_V2", _
        "CDP ilçe export yapýsý beklenen 12 sütundan daha az."

    End If

    '==========================================================
    ' 2. EPP / 7055 / K001-K999
    '==========================================================

    CORE_DURUM _
    "[2/4] EPP ilçe stoklarý SAP'tan alýnýyor..."

    hataMesaji = vbNullString

    basarili = ILCE_EPP_BLOK_AL(EPPVeri, EPPSatir, EPPSutun, _
        EPPExportSuresi, hataMesaji)
    If Not basarili Then
        hataMesaji = vbNullString
        basarili = ILCE_V2_EPP_ALV_AL(EPPVeri, EPPSatir, EPPSutun, _
            EPPExportSuresi, hataMesaji)
    End If

    If Not basarili Then

    Err.Raise _
        vbObjectError + 4111, _
        "ILCE_STOK_GUNCELLE_V2", _
        "EPP ilçe stoklarý alýnamadý." & _
        vbCrLf & hataMesaji

    End If

    If Not IsArray(EPPVeri) Then

    Err.Raise _
        vbObjectError + 4112, _
        "ILCE_STOK_GUNCELLE_V2", _
        "EPP ilçe verisi geçerli bir dizi deðil."

    End If

    If EPPSatir <= 0 Then

    Err.Raise _
        vbObjectError + 4113, _
        "ILCE_STOK_GUNCELLE_V2", _
        "EPP sisteminden K001-K999 aralýðýnda stok satýrý gelmedi."

    End If

    If EPPSutun < 12 Then

    Err.Raise _
        vbObjectError + 4114, _
        "ILCE_STOK_GUNCELLE_V2", _
        "EPP ilçe export yapýsý beklenen 12 sütundan daha az."

    End If

    '==========================================================
    ' 3. ÜRETÝM YERLERÝNÝ DOÐRULA
    '==========================================================

    ILCE_V2_UY_DOGRULA _
    CDPVeri, _
    CDPSatir, _
    beklenenCDP, _
    "CDP"

    ILCE_V2_UY_DOGRULA _
    EPPVeri, _
    EPPSatir, _
    beklenenEPP, _
    "EPP"

    '==========================================================
    ' 4. ESKÝ ÝLÇE_STOK VERÝSÝNÝ YEDEKLE
    '==========================================================

    CORE_DURUM _
    "[3/4] Mevcut ilçe stoklarý güvenliðe alýnýyor..."

    EskiSonSatir = STOK_SON_SATIR(wsStok)
    EskiSonSutun = STOK_SON_SUTUN(wsStok)

    If EskiSonSatir >= 7 And EskiSonSutun >= 1 Then

    If EskiSonSutun > 12 Then
        EskiSonSutun = 12
    End If

    EskiVeri = _
        wsStok.Range( _
            wsStok.Cells(7, 1), _
            wsStok.Cells(EskiSonSatir, EskiSonSutun)).Value2

    YedekVar = True

    Else

    YedekVar = False

    End If

    '==========================================================
    ' 5. RAM'DE BÝRLEÞTÝR
    '==========================================================

    BirlesikVeri = _
    ILCE_V2_VERILERI_BIRLESTIR( _
    CDPVeri, _
    CDPSatir, _
    EPPVeri, _
    EPPSatir)

    ToplamSatir = CDPSatir + EPPSatir

    If ToplamSatir <= 0 Then

    Err.Raise _
        vbObjectError + 4120, _
        "ILCE_STOK_GUNCELLE_V2", _
        "Birleþtirilecek ilçe stok verisi bulunamadý."

    End If

    '==========================================================
    ' 6. EXCEL ÝÞLEMLERÝ
    '==========================================================

    ExcelBaslangic = Timer

    CORE_DURUM _
    "[4/4] Ýlçe stoklarý ve raporlarý hazýrlanýyor..."

    ' SAP okumasinda il degisse bile farkli ilin verisini yazma.
    If AKSORA_AKTIF_IL() <> islemIl Then
        Err.Raise vbObjectError + 4131, "ILCE_STOK_GUNCELLE_V2", _
            "Islem sirasinda il secimi degisti. Mevcut stoklar korunuyor."
    End If
    If AKSORA_CANLI_STOK_IL() <> islemIl Then
        Err.Raise vbObjectError + 4132, "ILCE_STOK_GUNCELLE_V2", _
            "Yuklu stoklar islem iliyle uyusmuyor. Mevcut stoklar korunuyor."
    End If
    ILCE_V2_UY_DOGRULA CDPVeri, CDPSatir, beklenenCDP, "CDP"
    ILCE_V2_UY_DOGRULA EPPVeri, EPPSatir, beklenenEPP, "EPP"
    If UBound(BirlesikVeri, 1) <> ToplamSatir Or _
       UBound(BirlesikVeri, 2) <> 12 Then
        Err.Raise vbObjectError + 4133, "ILCE_STOK_GUNCELLE_V2", _
            "Birlesik ilce stok dizisi gecersiz. Mevcut stoklar korunuyor."
    End If

    yazmaBasladi = True

    ILCE_V2_BASLIKLARI_YAZ wsStok

    wsStok.Range( _
    "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

    wsStok.Range("A7").Resize( _
    ToplamSatir, _
    12).Value2 = BirlesikVeri

    ILCE_V2_STOK_OZET_GUNCELLE _
    wsStok, _
    ToplamSatir

    ILCE_V2_RAPOR_OLUSTUR _
    wsStok, _
    wsRapor, _
    ToplamSatir

    ILCE_V2_LISTE_OLUSTUR wsStok

    ILCE_V2_STOK_GORSEL_DUZEN _
    wsStok, _
    ToplamSatir

    STOK_GORUNUM_TEK "ÝLÇE_STOK"
    RAPOR_GORUNUM_TEK "ÝLÇE_RAPORLAR"

    ExcelSuresi = _
    STOK_GECEN_SURE(ExcelBaslangic)

    ToplamSure = _
    CORE_ISLEM_BITIR()

    application.Calculation = xlCalculationAutomatic
    application.CalculateFull

    If Not ANA_SESSIZ_MOD Then

    MsgBox _
    "Ýlçe stok güncellemesi baþarýyla tamamlandý." & _
    vbCrLf & vbCrLf & _
    "CDP / 7155 : " & _
    AKSORA_SAYI_METNI(CDPSatir, "#,##0") & " satýr" & vbCrLf & _
    "EPP / 7055 : " & _
    AKSORA_SAYI_METNI(EPPSatir, "#,##0") & " satýr" & vbCrLf & _
    "Toplam : " & _
    AKSORA_SAYI_METNI(ToplamSatir, "#,##0") & " satýr" & vbCrLf & vbCrLf & _
    "CDP SAP + Export : " & _
    AKSORA_SAYI_METNI(CDPExportSuresi, "0.00") & " sn" & vbCrLf & _
    "EPP SAP + Export : " & _
    AKSORA_SAYI_METNI(EPPExportSuresi, "0.00") & " sn" & vbCrLf & _
    "Excel Ýþleme : " & _
    AKSORA_SAYI_METNI(ExcelSuresi, "0.00") & " sn" & vbCrLf & _
    "Toplam Süre : " & _
    AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn", _
    vbInformation, _
    "AKSORA - ÝLÇE"

    End If

    Exit Sub

Hata:

    hataNo = Err.Number
    hataAciklama = Err.Description

    On Error Resume Next

    If yazmaBasladi Then

    CORE_DURUM _
        "Hata oluþtu. Eski ilçe stoklarý geri yükleniyor..."

    wsStok.Range( _
        "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

    If YedekVar Then

        If IsArray(EskiVeri) Then

            wsStok.Range("A7").Resize( _
                UBound(EskiVeri, 1), _
                UBound(EskiVeri, 2)).Value2 = EskiVeri

        End If

    End If

    ILCE_V2_STOK_OZET_GUNCELLE _
        wsStok, _
        ILCE_V2_VERI_SATIRI_SAYISI(wsStok)

    ILCE_V2_RAPOR_OLUSTUR _
        wsStok, _
        wsRapor, _
        ILCE_V2_VERI_SATIRI_SAYISI(wsStok)

    ILCE_V2_LISTE_OLUSTUR wsStok

    End If

    ToplamSure = CORE_ISLEM_BITIR()

    On Error GoTo 0

    If Not ANA_SESSIZ_MOD Then

    MsgBox _
    "Ýlçe stok güncellemesi sýrasýnda hata oluþtu." & _
    vbCrLf & vbCrLf & _
    "Hata No : " & hataNo & vbCrLf & _
    "Açýklama: " & hataAciklama & _
    vbCrLf & vbCrLf & _
    "Yeni veri tamamlanamadý." & vbCrLf & _
    "Mevcut ilçe stok verileri korunmaya çalýþýldý.", _
    vbCritical, _
    "AKSORA - ÝLÇE"

    End If
    
If ANA_SESSIZ_MOD Then
    Err.Raise hataNo, "ILCE_STOK_GUNCELLE_V2", hataAciklama
End If

    End Sub

    '==========================================================
    ' ÜRETÝM YERÝ DOÐRULAMA
    '==========================================================
    Private Sub ILCE_V2_UY_DOGRULA( _
    ByRef veri As Variant, _
    ByVal SatirSayisi As Long, _
    ByVal BeklenenUY As String, _
    ByVal kaynakAdi As String)

    Dim i As Long, c As Variant, depo As String, neden As String
    On Error GoTo Hata
    If Not IsArray(veri) Then neden = "Veri dizisi yok.": GoTo gecersiz
    If SatirSayisi <= 0 Or SatirSayisi > 250000 Then
        neden = "Satir sayisi gecersiz.": GoTo gecersiz
    End If
    If LBound(veri, 1) <> 1 Or UBound(veri, 1) <> SatirSayisi Or _
       LBound(veri, 2) <> 1 Or UBound(veri, 2) < 12 Then
        neden = "Veri dizisinin boyutlari gecersiz.": GoTo gecersiz
    End If
    For i = 1 To SatirSayisi
        For Each c In Array(1, 2, 7, 9, 10, 11)
            If IsError(veri(i, c)) Or IsNull(veri(i, c)) Then
                neden = "Hatali hucre; sutun " & c: GoTo gecersiz
            End If
        Next c
        If Trim$(CStr(veri(i, 1))) <> BeklenenUY Then
            neden = "Tesis: " & CStr(veri(i, 1)) & " / Beklenen: " & BeklenenUY
            GoTo gecersiz
        End If
        depo = UCase$(Trim$(CStr(veri(i, 2))))
        If Len(depo) <> 4 Then neden = "Gecersiz ilce deposu: " & depo: GoTo gecersiz
        If Left$(depo, 1) <> "K" Or Not (Mid$(depo, 2) Like "###") Then
            neden = "Gecersiz ilce deposu: " & depo: GoTo gecersiz
        End If
        If CLng(Mid$(depo, 2)) < 1 Then neden = "Gecersiz ilce deposu: " & depo: GoTo gecersiz
        If Len(Trim$(CStr(veri(i, 7)))) = 0 Or _
           Len(Trim$(CStr(veri(i, 10)))) = 0 Then
            neden = "Malzeme veya birim bos.": GoTo gecersiz
        End If
        If Not IsNumeric(veri(i, 9)) Or Not IsNumeric(veri(i, 11)) Then
            neden = "Miktar veya deger sayisal degil.": GoTo gecersiz
        End If
    Next i
    Exit Sub
gecersiz:
    On Error GoTo 0
    Err.Raise vbObjectError + 4130, "ILCE_V2_UY_DOGRULA", _
        kaynakAdi & " ilce verisi reddedildi. Satir: " & i & vbCrLf & _
        neden & vbCrLf & "Mevcut stoklar degistirilmedi."
Hata:
    neden = "Veri yapisi okunamadi: " & Err.Description
    Resume gecersiz
    End Sub

    '==========================================================
    ' CDP + EPP VERÝLERÝNÝ 12 SÜTUNDA BÝRLEÞTÝR
    '==========================================================
    Private Function ILCE_V2_VERILERI_BIRLESTIR( _
    ByRef CDPVeri As Variant, _
    ByVal CDPSatir As Long, _
    ByRef EPPVeri As Variant, _
    ByVal EPPSatir As Long) As Variant

    Dim sonuc() As Variant
    Dim i As Long
    Dim j As Long
    Dim HedefSatir As Long

    ReDim sonuc( _
    1 To CDPSatir + EPPSatir, _
    1 To 12)

    HedefSatir = 1

    For i = 1 To CDPSatir

    For j = 1 To 12
        sonuc(HedefSatir, j) = CDPVeri(i, j)
    Next j

    HedefSatir = HedefSatir + 1

    Next i

    For i = 1 To EPPSatir

    For j = 1 To 12
        sonuc(HedefSatir, j) = EPPVeri(i, j)
    Next j

    HedefSatir = HedefSatir + 1

    Next i

    ILCE_V2_VERILERI_BIRLESTIR = sonuc

    End Function

    '==========================================================
    ' ÝLÇE_STOK BAÞLIKLARI
    '==========================================================
    Private Sub ILCE_V2_BASLIKLARI_YAZ( _
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

    '==========================================================
    ' ÝLÇE_STOK ÜST ÖZET
    '==========================================================
    Private Sub ILCE_V2_STOK_OZET_GUNCELLE( _
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

    '==========================================================
    ' ÝLÇE_RAPORLAR
    '
    ' A = Depo yeri
    ' B = Depo yeri tanýmý
    ' C = 7155 toplam
    ' D = 7055 toplam
    ' E = Genel toplam
    '==========================================================
    Public Sub ILCE_V2_RAPOR_OLUSTUR( _
    ByVal wsStok As Worksheet, _
    ByVal wsRapor As Worksheet, _
    ByVal VeriSatiri As Long)

    Dim dict As Object

    Dim veri As Variant
    Dim bilgi As Variant

    Dim i As Long
    Dim HedefSatir As Long
    Dim SonStokSatir As Long
    Dim TasarimSatir As Long

    Dim uy As String
    Dim depo As String
    Dim DepoTanimi As String
    Dim deger As Double

    Dim anahtar As Variant

    Dim Toplam7155 As Double
    Dim Toplam7055 As Double
    Dim GenelToplam As Double

    Set dict = _
    CreateObject("Scripting.Dictionary")

    dict.CompareMode = vbTextCompare

    SonStokSatir = 6 + VeriSatiri

    If VeriSatiri > 0 Then

    veri = _
        wsStok.Range( _
            "A7:K" & SonStokSatir).Value2

    For i = 1 To UBound(veri, 1)

        uy = Trim$(CStr(veri(i, 1)))
        depo = Trim$(CStr(veri(i, 2)))
        DepoTanimi = CStr(veri(i, 3))

        deger = 0

        If IsNumeric(veri(i, 11)) Then
            deger = CDbl(veri(i, 11))
        End If

        If Len(depo) > 0 Then

            If Not dict.Exists(depo) Then

                dict.Add _
                    depo, _
                    Array( _
                        DepoTanimi, _
                        0#, _
                        0#)

            End If

            bilgi = dict(depo)

            If Len(Trim$(CStr(bilgi(0)))) = 0 Then
                bilgi(0) = DepoTanimi
            End If

            If uy = ILCE_CDP_UY Then

                bilgi(1) = _
                    CDbl(bilgi(1)) + deger

                Toplam7155 = _
                    Toplam7155 + deger

            ElseIf uy = ILCE_EPP_UY Then

                bilgi(2) = _
                    CDbl(bilgi(2)) + deger

                Toplam7055 = _
                    Toplam7055 + deger

            End If

            dict(depo) = bilgi

        End If

    Next i

    End If

    '==========================================================
    ' ESKÝ RAPOR VERÝLERÝNÝ TEMÝZLE
    '
    ' A:E temizlenir.
    ' Böylece eski E sütunu ve eski toplam kalýntýlarý silinir.
    '==========================================================

    On Error Resume Next

    If wsRapor.AutoFilterMode Then
    wsRapor.AutoFilterMode = False
    End If

    On Error GoTo 0

    wsRapor.Range( _
    "A8:E" & AKSORA_V13_SON(wsRapor)).ClearContents

    '==========================================================
    ' BAÞLIKLAR
    '==========================================================

    wsRapor.Range("A8").Value = _
    "Depo yeri"

    wsRapor.Range("B8").Value = _
    "Depo yeri tanýmý"

    wsRapor.Range("C8").Value = _
    ILCE_CDP_UY() & " Toplam Stok Deðeri"

    wsRapor.Range("D8").Value = _
    ILCE_EPP_UY() & " Toplam Stok Deðeri"

    wsRapor.Range("E8").Value = _
    "Toplam Stok Deðeri"

    '==========================================================
    ' DEPO SATIRLARI
    '==========================================================

    HedefSatir = 9

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
        CDbl(bilgi(1))

    wsRapor.Cells( _
        HedefSatir, 4).Value = _
        CDbl(bilgi(2))

    wsRapor.Cells( _
        HedefSatir, 5).Value = _
        CDbl(bilgi(1)) + CDbl(bilgi(2))

    HedefSatir = HedefSatir + 1

    Next anahtar

    '==========================================================
    ' GENEL TOPLAM
    '==========================================================

    GenelToplam = _
    Toplam7155 + Toplam7055

    wsRapor.Cells( _
    HedefSatir, 1).Value = _
    "TOPLAM"

    wsRapor.Cells( _
    HedefSatir, 3).Value = _
    Toplam7155

    wsRapor.Cells( _
    HedefSatir, 4).Value = _
    Toplam7055

    wsRapor.Cells( _
    HedefSatir, 5).Value = _
    GenelToplam

    wsRapor.Range( _
    "A" & HedefSatir & _
    ":E" & HedefSatir).Font.bold = True

    '==========================================================
    ' ÜST ÖZET
    '==========================================================

    wsRapor.Range("B3").Value = Now
    wsRapor.Range("B4").Value = dict.Count
    wsRapor.Range("B5").Value = GenelToplam

    wsRapor.Range("B3").NumberFormat = _
    "dd.mm.yyyy hh:mm:ss"

    wsRapor.Range("B4").NumberFormat = _
    "#,##0"

    wsRapor.Range("B5").NumberFormat = _
    "#,##0.00 ""TL"""

    '==========================================================
    ' SAYI FORMATLARI
    '==========================================================

    If HedefSatir >= 9 Then

    wsRapor.Range( _
        "C9:E" & HedefSatir).NumberFormat = _
        "#,##0.00 ""TL"""

    End If

    '==========================================================
    ' FÝLTRE
    '==========================================================

    If HedefSatir > 9 Then

    wsRapor.Range( _
        "A8:E" & HedefSatir - 1).AutoFilter

    Else

    wsRapor.Range("A8:E8").AutoFilter

    End If

    '==========================================================
    ' ÝLÇE_RAPORLAR NÝHAÝ TABLO TASARIMI
    ' CDP_RAPOR / EPP_RAPOR ile tamamen ayný standart
    '==========================================================

    With wsRapor.Range("A8:E8")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With

    For TasarimSatir = 9 To HedefSatir - 1
        With wsRapor.Range("A" & TasarimSatir & ":E" & TasarimSatir)
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

    With wsRapor.Range("A" & HedefSatir & ":E" & HedefSatir)
        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(0, 0, 0)
        .Font.bold = True
        .Borders.LineStyle = xlNone
    End With

    If HedefSatir > 10 Then
        With wsRapor.Sort
            .SortFields.Clear
            .SortFields.Add key:=wsRapor.Range("A9:A" & HedefSatir - 1), _
                            SortOn:=xlSortOnValues, _
                            Order:=xlAscending, _
                            DataOption:=xlSortNormal
            .SetRange wsRapor.Range("A8:E" & HedefSatir - 1)
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .Apply
        End With
    End If

    wsRapor.Range("A8:E" & HedefSatir).Borders.LineStyle = xlNone

    wsRapor.Range( _
        "A" & HedefSatir + 1 & _
        ":E" & AKSORA_V13_SON(wsRapor)).ClearFormats

    wsRapor.Columns("A").ColumnWidth = 12
    wsRapor.Columns("B").ColumnWidth = 28
    wsRapor.Columns("C").ColumnWidth = 20
    wsRapor.Columns("D").ColumnWidth = 20
    wsRapor.Columns("E").ColumnWidth = 20

    wsRapor.Rows("8:" & HedefSatir).RowHeight = 20
    wsRapor.Range("A9:B" & HedefSatir).HorizontalAlignment = xlLeft
    wsRapor.Range("C9:E" & HedefSatir).HorizontalAlignment = xlRight
    wsRapor.Range("A8:E" & HedefSatir).VerticalAlignment = xlCenter

    End Sub

    '==========================================================
    ' H2 ÝLÇE / DEPO SEÇÝM LÝSTESÝ
    '==========================================================
Private Sub ILCE_V2_LISTE_OLUSTUR(ByVal ws As Worksheet)
    AKFIX_STOK_LISTE ws
End Sub

    '==========================================================
    ' ÝLÇE_STOK TEMEL GÖRSEL DÜZEN
    '==========================================================
    Private Sub ILCE_V2_STOK_GORSEL_DUZEN( _
    ByVal ws As Worksheet, _
    ByVal VeriSatiri As Long)

    Dim sonSatir As Long
    Dim TasarimSatir As Long

    sonSatir = 6 + VeriSatiri

    If VeriSatiri <= 0 Then Exit Sub

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
    ws.Range("A7:G" & sonSatir).HorizontalAlignment = xlLeft
    ws.Range("H7:H" & sonSatir).HorizontalAlignment = xlLeft
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

    ws.Range("A6:L" & sonSatir).Borders.LineStyle = xlNone
    ws.Range("A" & sonSatir + 1 & ":L" & AKSORA_V13_SON(ws)).ClearFormats

    End Sub

    '==========================================================
    ' ÝLÇE_STOK MEVCUT VERÝ SATIRI SAYISI
    '==========================================================
    Private Function ILCE_V2_VERI_SATIRI_SAYISI( _
    ByVal ws As Worksheet) As Long

    Dim sonSatir As Long

    sonSatir = STOK_SON_SATIR(ws)

    If sonSatir < 7 Then

    ILCE_V2_VERI_SATIRI_SAYISI = 0

    Else

    ILCE_V2_VERI_SATIRI_SAYISI = _
        sonSatir - 6

    End If
    End Function

Private Function ILCE_V2_EPP_ALV_AL( _
ByRef VeriDizisi As Variant, _
ByRef VeriSatiri As Long, _
ByRef SutunSayisi As Long, _
ByRef ExportSuresi As Double, _
ByRef hataMesaji As String) As Boolean

Dim SAPSession As Object
Dim alv As Object
Dim KolonSirasi As Object
Dim KolonID(1 To 12) As String
Dim ham() As Variant
Dim sonuc() As Variant
Dim temiz() As Variant

Dim baslangic As Double
Dim SatirSayisi As Long
Dim r As Long, c As Long, n As Long
Dim uy As String, depo As String, malzeme As String

On Error GoTo Hata

ILCE_V2_EPP_ALV_AL = False
VeriSatiri = 0
SutunSayisi = 0
ExportSuresi = 0
hataMesaji = vbNullString
baslangic = Timer

Set SAPSession = SAP_SESSION_AL(ILCE_EPP_SISTEM)
If SAPSession Is Nothing Then
    hataMesaji = "EPP SAP oturumu bulunamadý."
    Exit Function
End If

SAPSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nMB52"
SAPSession.FindById("wnd[0]").SendVKey 0
ILCE_V2_EPP_BEKLE 0.3

SAPSession.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = ILCE_EPP_UY
SAPSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = ILCE_DEPO_ILK
SAPSession.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = ILCE_DEPO_SON
SAPSession.FindById("wnd[0]").SendVKey 8
ILCE_V2_EPP_BEKLE 0.5

If Not ILCE_V2_EPP_PANEL_SEC(SAPSession, hataMesaji) Then Exit Function

Set alv = SAPSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")

On Error Resume Next
alv.PressToolbarButton "&DELETE_FILTER"
Err.Clear
On Error GoTo Hata
ILCE_V2_EPP_BEKLE 0.3

Set KolonSirasi = alv.ColumnOrder
If KolonSirasi Is Nothing Then
    hataMesaji = "EPP MB52 sütun sýrasý alýnamadý."
    Exit Function
End If
If CLng(KolonSirasi.Count) < 12 Then
    hataMesaji = "EPP PANEL düzeninde 12 sütun bulunamadý."
    Exit Function
End If

For c = 1 To 12
    KolonID(c) = CStr(KolonSirasi.Item(c - 1))
Next c

SatirSayisi = CLng(alv.RowCount)
If SatirSayisi <= 0 Then
    hataMesaji = "EPP K001-K999 stok satýrý bulunamadý."
    Exit Function
End If

ReDim ham(1 To SatirSayisi, 1 To 12)

For r = 0 To SatirSayisi - 1
    On Error Resume Next
    alv.FirstVisibleRow = r
    alv.SetCurrentCell r, KolonID(7)
    On Error GoTo Hata
    DoEvents
    ILCE_V2_EPP_BEKLE 0.015

    For c = 1 To 12
        ham(r + 1, c) = alv.GetCellValue(r, KolonID(c))
    Next c
Next r

ReDim sonuc(1 To SatirSayisi, 1 To 12)

For r = 1 To SatirSayisi
    uy = Trim$(CStr(ham(r, 1)))
    depo = UCase$(Trim$(CStr(ham(r, 2))))
    malzeme = Trim$(CStr(ham(r, 7)))

    If uy = ILCE_EPP_UY And _
       Len(depo) = 4 And Left$(depo, 1) = "K" And _
       depo >= ILCE_DEPO_ILK And depo <= ILCE_DEPO_SON And _
       Len(malzeme) > 0 Then

        n = n + 1
        For c = 1 To 12
            sonuc(n, c) = ham(r, c)
        Next c

        sonuc(n, 9) = ILCE_V2_EPP_SAYI(sonuc(n, 9))
        sonuc(n, 11) = ILCE_V2_EPP_SAYI(sonuc(n, 11))
        sonuc(n, 12) = ILCE_V2_EPP_SAYI(sonuc(n, 12))
    End If
Next r

If n <= 0 Then
    hataMesaji = "EPP geçerli K depo stok satýrý bulunamadý."
    Exit Function
End If

ReDim temiz(1 To n, 1 To 12)
For r = 1 To n
    For c = 1 To 12
        temiz(r, c) = sonuc(r, c)
    Next c
Next r

VeriDizisi = temiz
VeriSatiri = n
SutunSayisi = 12
ExportSuresi = ILCE_V2_EPP_GECEN(baslangic)
ILCE_V2_EPP_ALV_AL = True
Exit Function

Hata:
hataMesaji = "EPP ilçe stok verisi SAP tablosundan alýnamadý." & vbCrLf & Err.Description
ExportSuresi = ILCE_V2_EPP_GECEN(baslangic)
ILCE_V2_EPP_ALV_AL = False
End Function

Private Function ILCE_V2_EPP_PANEL_SEC(ByVal SAPSession As Object, ByRef hataMesaji As String) As Boolean
Dim g As Object
Dim r As Long, n As Long, bulunan As Long
Dim v As String, t As String

On Error GoTo Hata
bulunan = -1

SAPSession.FindById("wnd[0]/tbar[1]/btn[33]").Press
ILCE_V2_EPP_BEKLE 0.2
Set g = SAPSession.FindById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")

n = g.RowCount
For r = 0 To n - 1
    v = vbNullString: t = vbNullString
    On Error Resume Next
    v = Trim$(CStr(g.GetCellValue(r, "VARIANT")))
    t = Trim$(CStr(g.GetCellValue(r, "TEXT")))
    On Error GoTo Hata
    If StrComp(v, "PANEL", vbTextCompare) = 0 Or StrComp(t, "MLZ STOK PANEL", vbTextCompare) = 0 Then
        bulunan = r
        Exit For
    End If
Next r

If bulunan < 0 Then
    hataMesaji = "EPP PANEL düzeni bulunamadý."
    Exit Function
End If

g.SetCurrentCell bulunan, "TEXT"
g.SelectedRows = CStr(bulunan)
g.ClickCurrentCell
ILCE_V2_EPP_BEKLE 0.3
ILCE_V2_EPP_PANEL_SEC = True
Exit Function

Hata:
hataMesaji = "EPP PANEL düzeni seçilemedi." & vbCrLf & Err.Description
ILCE_V2_EPP_PANEL_SEC = False
End Function

Private Function ILCE_V2_EPP_SAYI(ByVal deger As Variant) As Double
Dim s As String, neg As Boolean
If IsNumeric(deger) Then ILCE_V2_EPP_SAYI = CDbl(deger): Exit Function
s = Trim$(CStr(deger))
If Len(s) = 0 Or s = "-" Then Exit Function
If Left$(s, 1) = "-" Then neg = True: s = Mid$(s, 2)
If Right$(s, 1) = "-" Then neg = True: s = Left$(s, Len(s) - 1)
s = Replace(s, " ", "")
If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
    If InStrRev(s, ",") > InStrRev(s, ".") Then
        s = Replace(s, ".", "")
        s = Replace(s, ",", application.DecimalSeparator)
    Else
        s = Replace(s, ",", "")
        s = Replace(s, ".", application.DecimalSeparator)
    End If
ElseIf InStr(s, ",") > 0 Then
    s = Replace(s, ",", application.DecimalSeparator)
ElseIf InStr(s, ".") > 0 And application.DecimalSeparator <> "." Then
    s = Replace(s, ".", application.DecimalSeparator)
End If
If IsNumeric(s) Then
    ILCE_V2_EPP_SAYI = CDbl(s)
    If neg Then ILCE_V2_EPP_SAYI = -ILCE_V2_EPP_SAYI
End If
End Function

Private Sub ILCE_V2_EPP_BEKLE(ByVal saniye As Double)
Dim t As Double
t = Timer
Do While ILCE_V2_EPP_GECEN(t) < saniye
    DoEvents
Loop
End Sub

Private Function ILCE_V2_EPP_GECEN(ByVal baslangic As Double) As Double
If Timer >= baslangic Then
    ILCE_V2_EPP_GECEN = Timer - baslangic
Else
    ILCE_V2_EPP_GECEN = (86400# - baslangic) + Timer
End If
End Function





' 30'lu SAP ALV okuma: yalnizca gecerli stok satirlari aktarilir.
Public Function ILCE_CDP_BLOK_AL(ByRef veri As Variant, _
    ByRef satir As Long, ByRef sutun As Long, _
    ByRef sure As Double, ByRef Hata As String) As Boolean
    Dim blok As Long
    Dim ses As Object, alv As Object, duzen As Object, kolon As Object
    Dim ids(1 To 12) As String, ham() As Variant, temiz() As Variant
    Dim i As Long, j As Long, b As Long, son As Long, n As Long
    Dim secili As Long, tekrar As Long, stok As Long
    Dim ad As String, tanim As String, depo As String
    Dim sorun As Boolean, ozet As Boolean, bitti As Boolean
    Dim t As Double, neden As String, uy As String
    On Error GoTo Hata
    satir = 0: sutun = 0: sure = 0: Hata = vbNullString
        uy = ILCE_CDP_UY()
    t = Timer
    ' Hizli yol: ilin MB52 listesi tek sorguda metin olarak alinir (olmazsa asagidaki eski yontem).
    If AKHZ_TEK_AL(3, veri, satir, sutun, Hata) Then ILCE_CDP_BLOK_AL = True: Exit Function
    Set ses = SAP_SESSION_AL("CDP", Hata)
    If ses Is Nothing Then Err.Raise vbObjectError + 870, , "SAP oturumu bulunamadi: " & Hata
    ses.StartTransaction "MB52"
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = uy
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = "K001"
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = "K999"
    ses.FindById("wnd[0]").SendVKey 8
    ILCE_CDP_BLOK_BEKLE 0.4
    ses.FindById("wnd[0]/tbar[1]/btn[33]").Press
    ILCE_CDP_BLOK_BEKLE 0.3
    Set duzen = ses.FindById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    secili = -1
    For i = 0 To CLng(duzen.RowCount) - 1
        ad = Trim$(CStr(duzen.GetCellValue(i, "VARIANT")))
        tanim = Trim$(CStr(duzen.GetCellValue(i, "TEXT")))
        If StrComp(ad, "PANEL", vbTextCompare) = 0 Or _
           StrComp(tanim, "MLZ STOK PANEL", vbTextCompare) = 0 Then
            secili = i: Exit For
        End If
    Next i
    If secili < 0 Then Err.Raise vbObjectError + 871, , "PANEL duzeni bulunamadi."
    duzen.SetCurrentCell secili, "TEXT"
    duzen.SelectedRows = CStr(secili)
    duzen.ClickCurrentCell
    ILCE_CDP_BLOK_BEKLE 0.3
    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    On Error Resume Next
    alv.PressToolbarButton "&DELETE_FILTER"
    Err.Clear
    On Error GoTo Hata
    Set kolon = alv.ColumnOrder
    If CLng(kolon.Count) < 12 Then Err.Raise vbObjectError + 872, , "PANEL 12 sutun icermiyor."
    For j = 1 To 12
        ids(j) = CStr(kolon.Item(j - 1))
    Next j
    n = CLng(alv.RowCount)
    If n <= 0 Then Err.Raise vbObjectError + 873, , "SAP sonucu bos."
    ReDim ham(1 To n, 1 To 12)
    ReDim temiz(1 To n, 1 To 12)
    blok = AKSORA_SAP_BLOK_BOYU(alv)
    For b = 0 To n - 1 Step blok
        If bitti Then Exit For
        son = b + blok - 1
        If son >= n Then son = n - 1
        On Error Resume Next
        Err.Clear
        alv.FirstVisibleRow = b
        sorun = (Err.Number <> 0)
        On Error GoTo Hata
        If sorun Then Err.Raise vbObjectError + 874, , "Blok kaydirilamadi: " & (b + 1)
        DoEvents
        ILCE_CDP_BLOK_BEKLE 0.015
        AKSORA_TOPLU_SURE_TIK
        For i = b To son
            sorun = False: ozet = False
            On Error Resume Next
            For j = 1 To 12
                Err.Clear
                If j = 12 Then
                    ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                    ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                    If Trim$(CStr(ham(i + 1, 1))) = uy And _
                       Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                        ham(i + 1, j) = 0#
                    Else
                        ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                    End If
                Else
                    ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                End If
                If Err.Number <> 0 Then sorun = True: Exit For
            Next j
            On Error GoTo Hata
            If Not sorun Then
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
                ozet = ILCE_CDP_OZET_MI(ham, i + 1, uy)
                If ozet Then bitti = True: Exit For
            End If
            For tekrar = 1 To 2
                sorun = False
                On Error Resume Next
                Err.Clear
                alv.FirstVisibleRow = i
                alv.SetCurrentCell i, ids(7)
                sorun = (Err.Number <> 0)
                On Error GoTo Hata
                If Not sorun Then
                    DoEvents
                    ILCE_CDP_BLOK_BEKLE 0.015
                    On Error Resume Next
                    For j = 1 To 12
                        Err.Clear
                        If j = 12 Then
                            ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                            ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                            If Trim$(CStr(ham(i + 1, 1))) = uy And _
                               Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                                ham(i + 1, j) = 0#
                            Else
                                ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                            End If
                        Else
                            ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                        End If
                        If Err.Number <> 0 Then sorun = True: Exit For
                    Next j
                    On Error GoTo Hata
                    If Not sorun Then
                        If Trim$(CStr(ham(i + 1, 1))) = uy And _
                           Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then Exit For
                        If ILCE_CDP_OZET_MI(ham, i + 1, uy) Then Exit For
                    End If
                End If
            Next tekrar
            If Not sorun Then
                If ILCE_CDP_OZET_MI(ham, i + 1, uy) Then bitti = True: Exit For
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
            End If
            ' Ozet satiri SAP tarafindan secilemiyorsa son kisimda dur.
            If i >= n - 10 And stok > 0 And _
               Len(Trim$(CStr(ham(i + 1, 1)))) = 0 Then
                bitti = True: Exit For
            End If
            Err.Raise vbObjectError + 875, , "Satir dogrulanamadi: " & (i + 1)
KAYDET:
            depo = UCase$(Trim$(CStr(ham(i + 1, 2))))
            If Not (Len(depo) = 4 And Left$(depo, 1) = "K" And depo >= "K001" And depo <= "K999") Then
                Err.Raise vbObjectError + 876, , "Beklenmeyen depo: " & depo
            End If
            stok = stok + 1
            For j = 1 To 12
                temiz(stok, j) = ham(i + 1, j)
            Next j
            temiz(stok, 9) = ILCE_V2_EPP_SAYI(temiz(stok, 9))
            temiz(stok, 11) = ILCE_V2_EPP_SAYI(temiz(stok, 11))
            If IsNumeric(temiz(stok, 9)) And CDbl(temiz(stok, 9)) <> 0 Then
                temiz(stok, 12) = CDbl(temiz(stok, 11)) / CDbl(temiz(stok, 9))
            Else
                temiz(stok, 12) = 0#
            End If
        Next i
    Next b
    If stok <= 0 Then Err.Raise vbObjectError + 877, , "Gecerli stok satiri yok."
    ReDim ham(1 To stok, 1 To 12)
    For i = 1 To stok
        For j = 1 To 12
            ham(i, j) = temiz(i, j)
        Next j
    Next i
    veri = ham: satir = stok: sutun = 12
    sure = ILCE_CDP_BLOK_SURE(t)
    ILCE_CDP_BLOK_AL = True
    Exit Function
Hata:
    neden = Err.Description
    Hata = "30'lu blok okuma: " & neden
    sure = ILCE_CDP_BLOK_SURE(t)
    ILCE_CDP_BLOK_AL = False
End Function

Private Function ILCE_CDP_OZET_MI(ByRef ham() As Variant, _
    ByVal r As Long, ByVal uy As String) As Boolean
    Dim c As Long
    If Trim$(CStr(ham(r, 1))) = uy And _
       Len(Trim$(CStr(ham(r, 7)))) > 0 Then Exit Function
    If Len(Trim$(CStr(ham(r, 1)))) = 0 And _
       Len(Trim$(CStr(ham(r, 7)))) = 0 Then
        For c = 8 To 12
            If Len(Trim$(CStr(ham(r, c)))) > 0 Then
                ILCE_CDP_OZET_MI = True: Exit Function
            End If
        Next c
    End If
End Function

Private Sub ILCE_CDP_BLOK_BEKLE(ByVal saniye As Double)
    Dim t As Double
    t = Timer
    Do While ILCE_CDP_BLOK_SURE(t) < saniye
        DoEvents
    Loop
End Sub

Private Function ILCE_CDP_BLOK_SURE(ByVal t As Double) As Double
    If Timer >= t Then
        ILCE_CDP_BLOK_SURE = Timer - t
    Else
        ILCE_CDP_BLOK_SURE = 86400# - t + Timer
    End If
End Function


' 30'lu SAP ALV okuma: yalnizca gecerli stok satirlari aktarilir.
Public Function ILCE_EPP_BLOK_AL(ByRef veri As Variant, _
    ByRef satir As Long, ByRef sutun As Long, _
    ByRef sure As Double, ByRef Hata As String) As Boolean
    Dim blok As Long
    Dim ses As Object, alv As Object, duzen As Object, kolon As Object
    Dim ids(1 To 12) As String, ham() As Variant, temiz() As Variant
    Dim i As Long, j As Long, b As Long, son As Long, n As Long
    Dim secili As Long, tekrar As Long, stok As Long
    Dim ad As String, tanim As String, depo As String
    Dim sorun As Boolean, ozet As Boolean, bitti As Boolean
    Dim t As Double, neden As String, uy As String
    On Error GoTo Hata
    satir = 0: sutun = 0: sure = 0: Hata = vbNullString
        uy = ILCE_EPP_UY()
    t = Timer
    ' Hizli yol: ilin MB52 listesi tek sorguda metin olarak alinir (olmazsa asagidaki eski yontem).
    If AKHZ_TEK_AL(4, veri, satir, sutun, Hata) Then ILCE_EPP_BLOK_AL = True: Exit Function
    Set ses = SAP_SESSION_AL("EPP", Hata)
    If ses Is Nothing Then Err.Raise vbObjectError + 870, , "SAP oturumu bulunamadi: " & Hata
    ses.StartTransaction "MB52"
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = uy
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = "K001"
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = "K999"
    ses.FindById("wnd[0]").SendVKey 8
    ILCE_EPP_BLOK_BEKLE 0.4
    ses.FindById("wnd[0]/tbar[1]/btn[33]").Press
    ILCE_EPP_BLOK_BEKLE 0.3
    Set duzen = ses.FindById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    secili = -1
    For i = 0 To CLng(duzen.RowCount) - 1
        ad = Trim$(CStr(duzen.GetCellValue(i, "VARIANT")))
        tanim = Trim$(CStr(duzen.GetCellValue(i, "TEXT")))
        If StrComp(ad, "PANEL", vbTextCompare) = 0 Or _
           StrComp(tanim, "MLZ STOK PANEL", vbTextCompare) = 0 Then
            secili = i: Exit For
        End If
    Next i
    If secili < 0 Then Err.Raise vbObjectError + 871, , "PANEL duzeni bulunamadi."
    duzen.SetCurrentCell secili, "TEXT"
    duzen.SelectedRows = CStr(secili)
    duzen.ClickCurrentCell
    ILCE_EPP_BLOK_BEKLE 0.3
    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    On Error Resume Next
    alv.PressToolbarButton "&DELETE_FILTER"
    Err.Clear
    On Error GoTo Hata
    Set kolon = alv.ColumnOrder
    If CLng(kolon.Count) < 12 Then Err.Raise vbObjectError + 872, , "PANEL 12 sutun icermiyor."
    For j = 1 To 12
        ids(j) = CStr(kolon.Item(j - 1))
    Next j
    n = CLng(alv.RowCount)
    If n <= 0 Then Err.Raise vbObjectError + 873, , "SAP sonucu bos."
    ReDim ham(1 To n, 1 To 12)
    ReDim temiz(1 To n, 1 To 12)
    blok = AKSORA_SAP_BLOK_BOYU(alv)
    For b = 0 To n - 1 Step blok
        If bitti Then Exit For
        son = b + blok - 1
        If son >= n Then son = n - 1
        On Error Resume Next
        Err.Clear
        alv.FirstVisibleRow = b
        sorun = (Err.Number <> 0)
        On Error GoTo Hata
        If sorun Then Err.Raise vbObjectError + 874, , "Blok kaydirilamadi: " & (b + 1)
        DoEvents
        ILCE_EPP_BLOK_BEKLE 0.015
        AKSORA_TOPLU_SURE_TIK
        For i = b To son
            sorun = False: ozet = False
            On Error Resume Next
            For j = 1 To 12
                Err.Clear
                If j = 12 Then
                    ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                    ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                    If Trim$(CStr(ham(i + 1, 1))) = uy And _
                       Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                        ham(i + 1, j) = 0#
                    Else
                        ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                    End If
                Else
                    ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                End If
                If Err.Number <> 0 Then sorun = True: Exit For
            Next j
            On Error GoTo Hata
            If Not sorun Then
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
                ozet = ILCE_EPP_OZET_MI(ham, i + 1, uy)
                If ozet Then bitti = True: Exit For
            End If
            For tekrar = 1 To 2
                sorun = False
                On Error Resume Next
                Err.Clear
                alv.FirstVisibleRow = i
                alv.SetCurrentCell i, ids(7)
                sorun = (Err.Number <> 0)
                On Error GoTo Hata
                If Not sorun Then
                    DoEvents
                    ILCE_EPP_BLOK_BEKLE 0.015
                    On Error Resume Next
                    For j = 1 To 12
                        Err.Clear
                        If j = 12 Then
                            ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                            ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                            If Trim$(CStr(ham(i + 1, 1))) = uy And _
                               Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                                ham(i + 1, j) = 0#
                            Else
                                ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                            End If
                        Else
                            ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                        End If
                        If Err.Number <> 0 Then sorun = True: Exit For
                    Next j
                    On Error GoTo Hata
                    If Not sorun Then
                        If Trim$(CStr(ham(i + 1, 1))) = uy And _
                           Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then Exit For
                        If ILCE_EPP_OZET_MI(ham, i + 1, uy) Then Exit For
                    End If
                End If
            Next tekrar
            If Not sorun Then
                If ILCE_EPP_OZET_MI(ham, i + 1, uy) Then bitti = True: Exit For
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
            End If
            ' Ozet satiri SAP tarafindan secilemiyorsa son kisimda dur.
            If i >= n - 10 And stok > 0 And _
               Len(Trim$(CStr(ham(i + 1, 1)))) = 0 Then
                bitti = True: Exit For
            End If
            Err.Raise vbObjectError + 875, , "Satir dogrulanamadi: " & (i + 1)
KAYDET:
            depo = UCase$(Trim$(CStr(ham(i + 1, 2))))
            If Not (Len(depo) = 4 And Left$(depo, 1) = "K" And depo >= "K001" And depo <= "K999") Then
                Err.Raise vbObjectError + 876, , "Beklenmeyen depo: " & depo
            End If
            stok = stok + 1
            For j = 1 To 12
                temiz(stok, j) = ham(i + 1, j)
            Next j
            temiz(stok, 9) = ILCE_V2_EPP_SAYI(temiz(stok, 9))
            temiz(stok, 11) = ILCE_V2_EPP_SAYI(temiz(stok, 11))
            If IsNumeric(temiz(stok, 9)) And CDbl(temiz(stok, 9)) <> 0 Then
                temiz(stok, 12) = CDbl(temiz(stok, 11)) / CDbl(temiz(stok, 9))
            Else
                temiz(stok, 12) = 0#
            End If
        Next i
    Next b
    If stok <= 0 Then Err.Raise vbObjectError + 877, , "Gecerli stok satiri yok."
    ReDim ham(1 To stok, 1 To 12)
    For i = 1 To stok
        For j = 1 To 12
            ham(i, j) = temiz(i, j)
        Next j
    Next i
    veri = ham: satir = stok: sutun = 12
    sure = ILCE_EPP_BLOK_SURE(t)
    ILCE_EPP_BLOK_AL = True
    Exit Function
Hata:
    neden = Err.Description
    Hata = "30'lu blok okuma: " & neden
    sure = ILCE_EPP_BLOK_SURE(t)
    ILCE_EPP_BLOK_AL = False
End Function

Private Function ILCE_EPP_OZET_MI(ByRef ham() As Variant, _
    ByVal r As Long, ByVal uy As String) As Boolean
    Dim c As Long
    If Trim$(CStr(ham(r, 1))) = uy And _
       Len(Trim$(CStr(ham(r, 7)))) > 0 Then Exit Function
    If Len(Trim$(CStr(ham(r, 1)))) = 0 And _
       Len(Trim$(CStr(ham(r, 7)))) = 0 Then
        For c = 8 To 12
            If Len(Trim$(CStr(ham(r, c)))) > 0 Then
                ILCE_EPP_OZET_MI = True: Exit Function
            End If
        Next c
    End If
End Function

Private Sub ILCE_EPP_BLOK_BEKLE(ByVal saniye As Double)
    Dim t As Double
    t = Timer
    Do While ILCE_EPP_BLOK_SURE(t) < saniye
        DoEvents
    Loop
End Sub

Private Function ILCE_EPP_BLOK_SURE(ByVal t As Double) As Double
    If Timer >= t Then
        ILCE_EPP_BLOK_SURE = Timer - t
    Else
        ILCE_EPP_BLOK_SURE = 86400# - t + Timer
    End If
End Function

Private Function ILCE_CDP_UY() As String
    ILCE_CDP_UY = AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "CDP")
End Function

Private Function ILCE_EPP_UY() As String
    ILCE_EPP_UY = AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "EPP")
End Function

