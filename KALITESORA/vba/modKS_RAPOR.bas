Attribute VB_Name = "modKS_RAPOR"
Option Explicit

' RAPORLAR: 5 yýllýk karþýlaþtýrma, aylýk daðýlým ve grafikler.
' KARNE  : tek arama kutusu; malzeme (kod / ad) ya da üretici (kod / ad) karnesi, tedarikçi sýralamasý.

Public Const R_YIL_SATIR As Long = 6        ' 5 yýllýk tablo baþlýk satýrý
Public Const R_GOSTERGE_ILK As Long = 7     ' ilk gösterge satýrý
Public Const R_GARIC As Long = 15           ' Garanti içi arýza satýrý (grafik)
Public Const R_GARDIS As Long = 16          ' Garanti dýþý arýza satýrý (grafik)
Public Const R_AY_BASLIK As Long = 25       ' aylýk tablo baþlýk satýrý
Private Const K_ILK As Long = 5             ' karne çýktýsýnýn ilk satýrý

' ================================================================ ortak toplama

Public Function KS_SevkYil(ByRef arS As Variant, ByVal i As Long) As Long
    Dim d As Date
    If KS_Tarih(arS(i, S_GELIS), d) Then
        KS_SevkYil = Year(d)
    ElseIf KS_Tarih(arS(i, S_PLAN), d) Then
        KS_SevkYil = Year(d)
    End If
End Function

Public Function KS_SevkAy(ByRef arS As Variant, ByVal i As Long) As Long
    Dim d As Date
    If KS_Tarih(arS(i, S_GELIS), d) Then
        KS_SevkAy = Month(d)
    ElseIf KS_Tarih(arS(i, S_PLAN), d) Then
        KS_SevkAy = Month(d)
    End If
End Function

Private Function KS_Bol(ByVal pay As Double, ByVal payda As Double) As Variant
    If payda <> 0 Then KS_Bol = pay / payda Else KS_Bol = Empty
End Function

' y1..y2 yýllarý arasýndaki göstergeler (0..16). Sevkiyat yýlý: geliþ, yoksa planlanan tarih.
Public Function KS_Topla(ByRef arS As Variant, ByRef aA As Variant, ByVal y1 As Long, ByVal y2 As Long) As Variant
    Dim r(0 To 16) As Variant, i As Long, y As Long, x As Double, d As Date, k As String
    Dim cnt As Long, mik As Double, tut As Double, zam As Long, tes As Long, gecTop As Double
    Dim aC As Long, aM As Double, gi As Long, gd As Long, donTop As Double, donN As Long, donMax As Double
    Dim cozTop As Double, cozN As Long, mal As Double, mU As Double, mK As Double, ku As String, kk As String
    ku = KS_Anahtar("Üretici")
    kk = KS_Anahtar("Kurum")
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            y = KS_SevkYil(arS, i)
            If y >= y1 And y <= y2 Then
                cnt = cnt + 1
                If KS_Sayi(arS(i, S_ADET), x) Then mik = mik + x
                If KS_Sayi(arS(i, S_TUTAR), x) Then tut = tut + x
                Select Case KS_Metin(arS(i, S_TDURUM))
                    Case "ZAMANINDA"
                        zam = zam + 1
                        tes = tes + 1
                    Case "GECÝKMELÝ"
                        tes = tes + 1
                        If KS_Sayi(arS(i, S_GECIKME), x) Then gecTop = gecTop + x
                End Select
            End If
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Tarih(aA(i, A_TARIH), d) Then
                y = Year(d)
                If y >= y1 And y <= y2 Then
                    aC = aC + 1
                    If KS_Sayi(aA(i, A_ADET), x) Then aM = aM + x
                    Select Case KS_Metin(aA(i, A_KAPSAM))
                        Case "EVET": gi = gi + 1
                        Case "HAYIR": gd = gd + 1
                    End Select
                    If KS_Sayi(aA(i, A_DONUSGUN), x) Then
                        donTop = donTop + x
                        donN = donN + 1
                        If x > donMax Then donMax = x
                    End If
                    If KS_Sayi(aA(i, A_COZUMGUN), x) Then
                        cozTop = cozTop + x
                        cozN = cozN + 1
                    End If
                    If KS_Sayi(aA(i, A_MALIYET), x) Then
                        mal = mal + x
                        k = KS_Anahtar(aA(i, A_KARS))
                        If k = ku Then mU = mU + x
                        If k = kk Then mK = mK + x
                    End If
                End If
            End If
        Next i
    End If
    r(0) = cnt: r(1) = mik: r(2) = tut
    r(3) = KS_Bol(zam, tes)
    r(4) = KS_Bol(gecTop, tes)
    r(5) = aC: r(6) = aM
    r(7) = KS_Bol(aM, mik)
    r(8) = gi: r(9) = gd
    r(10) = KS_Bol(donTop, donN)
    r(11) = KS_Bol(cozTop, cozN)
    r(12) = mal: r(13) = mU: r(14) = mK
    r(15) = donMax: r(16) = tes
    KS_Topla = r
End Function

' ================================================================ RAPORLAR

Private Function KS_Gostergeler() As Variant
    KS_Gostergeler = Array( _
        Array("Sevk Kalemi", "n"), Array("Sevk Edilen Miktar", "n"), Array("Sevk Tutarý (TL)", "tl"), _
        Array("Zamanýnda Teslim Oraný", "oran"), Array("Ortalama Gecikme (gün)", "gun"), _
        Array("Arýza Kaydý", "n"), Array("Arýzalý Miktar", "n"), Array("Arýza Oraný (arýzalý / sevk)", "oran"), _
        Array("Garanti Ýçi Arýza", "n"), Array("Garanti Dýþý Arýza", "n"), _
        Array("Ort. Üretici Dönüþ Süresi (gün)", "gun"), Array("Ort. Çözüm Süresi (gün)", "gun"), _
        Array("Tamir Maliyeti (TL)", "tl"), Array("Üreticinin Karþýladýðý (TL)", "tl"), Array("Kurumun Karþýladýðý (TL)", "tl"))
End Function

Private Function KS_TurBicim(ByVal tur As String) As String
    Select Case tur
        Case "tl": KS_TurBicim = "#,##0"
        Case "oran": KS_TurBicim = "0.0%"
        Case "gun": KS_TurBicim = "0.0"
        Case Else: KS_TurBicim = "#,##0"
    End Select
End Function

Public Function KS_AyAdi(ByVal ay As Long) As String
    KS_AyAdi = Choose(ay, "Ocak", "Þubat", "Mart", "Nisan", "Mayýs", "Haziran", "Temmuz", "Aðustos", _
        "Eylül", "Ekim", "Kasým", "Aralýk")
End Function

Public Sub KS_RaporKur()
    Dim ws As Worksheet, g As Variant, i As Long, alt As Double
    Set ws = KS_WS(SH_RAPOR)
    KS_SayfaSifirla ws
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:G").ColumnWidth = 13
    ws.Columns("H").ColumnWidth = 14
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 3
    ws.Columns("K:R").ColumnWidth = 9
    alt = KS_SayfaCerceve(ws, False)
    ws.Rows(1).RowHeight = alt + 6
    ws.Rows(2).RowHeight = 44
    KS_SayfaBaslik ws, ws.Rows(2).Top + 4, "RAPORLAR", "Seçilen yýl ve önceki 4 yýl. Sevkiyat yýlý: geliþ tarihi, yoksa planlanan tarih."

    ws.Rows(3).RowHeight = 28
    With ws.Range("B3")
        .Value = "RAPOR YILI (son yýl)"
        .Font.Bold = True
        .Font.Size = 8.5
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("C3")
        .Value = Year(Date)
        .Interior.Color = vbWhite
        .Font.Bold = True
        .Font.Color = C_LAC()
        .HorizontalAlignment = xlCenter
        .Locked = False
        .BorderAround xlContinuous, xlThin, , RGB(190, 204, 222)
    End With
    KS_AdHucre "rYil", ws.Range("C3")
    KS_Liste ws.Range("C3"), "lstYilSayi", True
    KS_Buton ws, "ksR_Hesapla", "HESAPLA", ws.Range("D3").Left + 8, ws.Rows(3).Top + 2, 100, 24, "KS_RAPOR_HESAPLA", "ana"
    KS_Buton ws, "ksR_Pdf", "PDF OLARAK KAYDET", ws.Range("D3").Left + 116, ws.Rows(3).Top + 2, 140, 24, "KS_PDF_KAYDET", "cizgi"

    KS_Bolum ws.Range("B5:I5"), "5 YILLIK KARÞILAÞTIRMA"
    KS_TabloBaslik ws.Range("B" & R_YIL_SATIR & ":I" & R_YIL_SATIR), Array("GÖSTERGE", "", "", "", "", "", "5 YIL", "DEÐÝÞÝM")
    g = KS_Gostergeler()
    For i = 0 To UBound(g)
        ws.Cells(R_GOSTERGE_ILK + i, 2).Value = g(i)(0)
        ws.Range(ws.Cells(R_GOSTERGE_ILK + i, 3), ws.Cells(R_GOSTERGE_ILK + i, 8)).NumberFormat = KS_TurBicim(CStr(g(i)(1)))
    Next i
    KS_TabloGovde ws.Range(ws.Cells(R_GOSTERGE_ILK, 2), ws.Cells(R_GOSTERGE_ILK + UBound(g), 9))
    ws.Range(ws.Cells(R_GOSTERGE_ILK, 2), ws.Cells(R_GOSTERGE_ILK + UBound(g), 2)).IndentLevel = 1
    ws.Range(ws.Cells(R_GOSTERGE_ILK, 3), ws.Cells(R_GOSTERGE_ILK + UBound(g), 9)).HorizontalAlignment = xlRight
    With ws.Range(ws.Cells(R_GOSTERGE_ILK, 8), ws.Cells(R_GOSTERGE_ILK + UBound(g), 8))
        .Interior.Color = C_ACIKMAVI()
        .Font.Bold = True
    End With

    KS_Bolum ws.Range("B" & R_AY_BASLIK - 1 & ":I" & R_AY_BASLIK - 1), "AYLIK DAÐILIM"
    KS_TabloBaslik ws.Range("B" & R_AY_BASLIK & ":H" & R_AY_BASLIK), Array("AY", "Sevk Kalemi", "Sevk Tutarý (TL)", _
        "Arýza Kaydý", "Garanti Ýçi", "Tamir Maliyeti", "Ort. Dönüþ (gün)")
    For i = 1 To 12
        ws.Cells(R_AY_BASLIK + i, 2).Value = KS_AyAdi(i)
    Next i
    ws.Cells(R_AY_BASLIK + 13, 2).Value = "TOPLAM"
    KS_TabloGovde ws.Range(ws.Cells(R_AY_BASLIK + 1, 2), ws.Cells(R_AY_BASLIK + 13, 8))
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 2), ws.Cells(R_AY_BASLIK + 13, 2)).IndentLevel = 1
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 3), ws.Cells(R_AY_BASLIK + 13, 7)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 8), ws.Cells(R_AY_BASLIK + 13, 8)).NumberFormat = "0.0"
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 3), ws.Cells(R_AY_BASLIK + 13, 8)).HorizontalAlignment = xlRight
    With ws.Range(ws.Cells(R_AY_BASLIK + 13, 2), ws.Cells(R_AY_BASLIK + 13, 8))
        .Font.Bold = True
        .Interior.Color = C_ACIKMAVI()
        .Borders(xlEdgeTop).Color = C_LAC()
    End With
    With ws.Range("B" & R_AY_BASLIK + 15)
        .Value = "Deðiþim: adet ve tutarda önceki yýla göre % deðiþim; oranlarda puan, sürelerde gün farký."
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = C_GRI()
    End With
    KS_Bolum ws.Range("K5:R5"), "5 YILLIK ARIZA EÐÝLÝMÝ"
    KS_Bolum ws.Range("K" & R_AY_BASLIK - 1 & ":R" & R_AY_BASLIK - 1), "AYLIK ARIZA VE TAMÝR MALÝYETÝ"
    KS_GrafikKur ws
    KS_ZeminUygula ws, True
    KS_Pencere ws, 90, 0
End Sub

Private Sub KS_GrafikKur(ByVal ws As Worksheet)
    Dim co As ChartObject, s As Series, sol As Double, gen As Double
    sol = ws.Columns("K").Left
    gen = ws.Columns("S").Left - sol
    Set co = ws.ChartObjects.Add(sol, ws.Rows(R_YIL_SATIR).Top, gen, _
        ws.Rows(R_GOSTERGE_ILK + 15).Top - ws.Rows(R_YIL_SATIR).Top)
    co.Name = "ksGrafikYil"
    With co.Chart
        .ChartType = xlColumnStacked
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = "Garanti içi"
        s.Values = ws.Range(ws.Cells(R_GARIC, 3), ws.Cells(R_GARIC, 7))
        s.XValues = ws.Range(ws.Cells(R_YIL_SATIR, 3), ws.Cells(R_YIL_SATIR, 7))
        s.Format.Fill.ForeColor.RGB = C_LAC()
        Set s = .SeriesCollection.NewSeries
        s.Name = "Garanti dýþý"
        s.Values = ws.Range(ws.Cells(R_GARDIS, 3), ws.Cells(R_GARDIS, 7))
        s.XValues = ws.Range(ws.Cells(R_YIL_SATIR, 3), ws.Cells(R_YIL_SATIR, 7))
        s.Format.Fill.ForeColor.RGB = C_CAMGOBEGI()
        .ChartGroups(1).GapWidth = 70
    End With
    KS_GrafikBicim co.Chart

    Set co = ws.ChartObjects.Add(sol, ws.Rows(R_AY_BASLIK).Top, gen, _
        ws.Rows(R_AY_BASLIK + 14).Top - ws.Rows(R_AY_BASLIK).Top)
    co.Name = "ksGrafikAy"
    With co.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = "Arýza kaydý"
        s.Values = ws.Range(ws.Cells(R_AY_BASLIK + 1, 5), ws.Cells(R_AY_BASLIK + 12, 5))
        s.XValues = ws.Range(ws.Cells(R_AY_BASLIK + 1, 2), ws.Cells(R_AY_BASLIK + 12, 2))
        s.Format.Fill.ForeColor.RGB = C_LAC()
        Set s = .SeriesCollection.NewSeries
        s.Name = "Tamir maliyeti (TL)"
        s.Values = ws.Range(ws.Cells(R_AY_BASLIK + 1, 7), ws.Cells(R_AY_BASLIK + 12, 7))
        s.XValues = ws.Range(ws.Cells(R_AY_BASLIK + 1, 2), ws.Cells(R_AY_BASLIK + 12, 2))
        s.ChartType = xlLineMarkers
        s.AxisGroup = xlSecondary
        s.Format.Line.ForeColor.RGB = C_CAMGOBEGI()
        s.Format.Line.Weight = 2.25
        s.MarkerStyle = xlMarkerStyleCircle
        s.MarkerSize = 6
        s.MarkerBackgroundColor = vbWhite
        s.MarkerForegroundColor = C_CAMGOBEGI()
        .ChartGroups(1).GapWidth = 80
    End With
    KS_GrafikBicim co.Chart
    On Error Resume Next
    co.Chart.Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "#,##0"
    co.Chart.Axes(xlValue, xlSecondary).Format.Line.Visible = msoFalse
    On Error GoTo 0
End Sub

Private Sub KS_GrafikBicim(ByVal ch As Chart)
    On Error Resume Next
    ch.HasTitle = False
    ch.HasLegend = True
    ch.Legend.Position = xlLegendPositionTop
    ch.Legend.Font.Size = 8
    ch.Legend.Font.Color = C_GRI()
    ch.ChartArea.Format.Fill.ForeColor.RGB = vbWhite
    ch.ChartArea.Format.Line.ForeColor.RGB = RGB(222, 231, 241)
    ch.ChartArea.Font.Name = KS_FONT
    ch.ChartArea.Font.Size = 8
    ch.ChartArea.Font.Color = C_GRI()
    ch.PlotArea.Format.Fill.Visible = msoFalse
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(232, 237, 243)
    ch.Axes(xlValue).Format.Line.Visible = msoFalse
    ch.Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = RGB(205, 216, 230)
    ch.Axes(xlCategory).TickLabels.Font.Color = C_METIN()
    On Error GoTo 0
End Sub

Public Sub KS_RAPOR_HESAPLA()
    Dim y As Double
    On Error GoTo Hata
    If Not KS_Sayi(KS_FV("rYil"), y) Then y = Year(Date)
    KS_Basla
    KS_HesaplaTumu
    KS_RaporHesapla CLng(y)
    KS_Bitir
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Rapor hesaplanamadý: " & Err.Description, vbExclamation, KS_AD
End Sub

Public Sub KS_RaporHesapla(ByVal yil As Long)
    Dim ws As Worksheet, arS As Variant, aA As Variant, g As Variant, sut(1 To 6) As Variant
    Dim i As Long, j As Long, ay As Long, d As Date, x As Double, korumali As Boolean
    Dim mS(1 To 13) As Double, mT(1 To 13) As Double, mA(1 To 13) As Double, mG(1 To 13) As Double
    Dim mM(1 To 13) As Double, mD(1 To 13) As Double, mDn(1 To 13) As Long, v As Variant, onc As Variant

    Set ws = KS_WS(SH_RAPOR)
    korumali = ws.ProtectContents
    KS_Ac ws
    arS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    g = KS_Gostergeler()
    ws.Range("C3").Value = yil
    For j = 1 To 5
        ws.Cells(R_YIL_SATIR, 2 + j).Value = CStr(yil - 5 + j)
        sut(j) = KS_Topla(arS, aA, yil - 5 + j, yil - 5 + j)
    Next j
    sut(6) = KS_Topla(arS, aA, yil - 4, yil)
    ws.Cells(R_YIL_SATIR, 8).Value = "5 YIL"
    For i = 0 To UBound(g)
        For j = 1 To 6
            ws.Cells(R_GOSTERGE_ILK + i, 2 + j).Value = sut(j)(i)
        Next j
        v = sut(5)(i): onc = sut(4)(i)
        With ws.Cells(R_GOSTERGE_ILK + i, 9)
            .Value = Empty
            .Font.Color = C_GRI()
            If Not IsEmpty(v) And Not IsEmpty(onc) Then
                Select Case g(i)(1)
                    Case "oran"
                        .Value = (v - onc) * 100
                        .NumberFormat = "+0.0"" puan"";-0.0"" puan"";0"
                    Case "gun"
                        .Value = v - onc
                        .NumberFormat = "+0.0"" gün"";-0.0"" gün"";0"
                    Case Else
                        If onc <> 0 Then
                            .Value = (v - onc) / onc
                            .NumberFormat = "+0%;-0%;0%"
                        End If
                End Select
            End If
        End With
    Next i

    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            If KS_SevkYil(arS, i) = yil Then
                ay = KS_SevkAy(arS, i)
                mS(ay) = mS(ay) + 1
                If KS_Sayi(arS(i, S_TUTAR), x) Then mT(ay) = mT(ay) + x
            End If
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Tarih(aA(i, A_TARIH), d) Then
                If Year(d) = yil Then
                    ay = Month(d)
                    mA(ay) = mA(ay) + 1
                    If KS_Metin(aA(i, A_KAPSAM)) = "EVET" Then mG(ay) = mG(ay) + 1
                    If KS_Sayi(aA(i, A_MALIYET), x) Then mM(ay) = mM(ay) + x
                    If KS_Sayi(aA(i, A_DONUSGUN), x) Then
                        mD(ay) = mD(ay) + x
                        mDn(ay) = mDn(ay) + 1
                    End If
                End If
            End If
        Next i
    End If
    For ay = 1 To 12
        mS(13) = mS(13) + mS(ay): mT(13) = mT(13) + mT(ay): mA(13) = mA(13) + mA(ay)
        mG(13) = mG(13) + mG(ay): mM(13) = mM(13) + mM(ay): mD(13) = mD(13) + mD(ay): mDn(13) = mDn(13) + mDn(ay)
    Next ay
    For ay = 1 To 13
        ws.Cells(R_AY_BASLIK + ay, 3).Value = mS(ay)
        ws.Cells(R_AY_BASLIK + ay, 4).Value = mT(ay)
        ws.Cells(R_AY_BASLIK + ay, 5).Value = mA(ay)
        ws.Cells(R_AY_BASLIK + ay, 6).Value = mG(ay)
        ws.Cells(R_AY_BASLIK + ay, 7).Value = mM(ay)
        If mDn(ay) > 0 Then ws.Cells(R_AY_BASLIK + ay, 8).Value = mD(ay) / mDn(ay) Else ws.Cells(R_AY_BASLIK + ay, 8).Value = Empty
    Next ay
    ws.Range("B" & R_AY_BASLIK - 1).Value = "AYLIK DAÐILIM  -  " & yil
    If korumali Then KS_Kapat ws
End Sub

' ================================================================ KARNE

Public Sub KS_KarneKur()
    Dim ws As Worksheet, alt As Double, et As Variant, mk As Variant, i As Long, gen As Double, x As Double
    Set ws = KS_WS(SH_KARNE)
    KS_SayfaSifirla ws
    ws.Columns("B:N").ColumnWidth = 16
    ws.Columns("O").ColumnWidth = 2
    alt = KS_SayfaCerceve(ws, False)
    ws.Rows(1).RowHeight = alt + 6
    ws.Rows(2).RowHeight = 44
    KS_SayfaBaslik ws, ws.Rows(2).Top + 4, "KARNE", _
        "Malzeme kodu / adý ya da üretici kodu / adý yazýn; uygun karne kendiliðinden açýlýr."
    et = Array("MALZEME KARNESÝ", "ÜRETÝCÝ KARNESÝ", "TEDARÝKÇÝ SIRALAMASI")
    mk = Array("KS_KARNE_MALZEME", "KS_KARNE_URETICI", "KS_KARNE_SIRALAMA")
    gen = 150
    x = ws.Cells(1, 2).Left + KS_GENISLIK - 3 * gen - 2 * 6
    For i = 0 To 2
        KS_Buton ws, "ksK_" & i, CStr(et(i)), x + i * (gen + 6), ws.Rows(2).Top + 9, gen, 28, CStr(mk(i)), "cizgi"
    Next i

    ws.Rows(3).RowHeight = 32
    With ws.Range("B3")
        .Value = "ARA  (kod ya da ad)"
        .Font.Bold = True
        .Font.Size = 8.5
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
    End With
    ws.Range("C3:F3").Merge
    With ws.Range("C3:F3")
        .Interior.Color = vbWhite
        .Font.Size = 11
        .Font.Bold = True
        .Font.Color = C_LAC()
        .Locked = False
        .NumberFormat = "@"
        .IndentLevel = 1
        .BorderAround xlContinuous, xlMedium, , C_MAVI()
    End With
    KS_AdHucre "kAra", ws.Range("C3")
    KS_Buton ws, "ksK_Goster", "GÖSTER", ws.Range("G3").Left + 8, ws.Rows(3).Top + 4, 100, 24, "KS_KARNE_GOSTER", "ana"
    With ws.Range("I3")
        .Value = "YIL (sýralama)"
        .Font.Bold = True
        .Font.Size = 8.5
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlRight
    End With
    With ws.Range("J3")
        .Value = "TÜMÜ"
        .Interior.Color = vbWhite
        .Font.Bold = True
        .Font.Color = C_LAC()
        .HorizontalAlignment = xlCenter
        .Locked = False
        .BorderAround xlContinuous, xlThin, , RGB(190, 204, 222)
    End With
    KS_AdHucre "kYil", ws.Range("J3")
    KS_Liste ws.Range("J3"), "lstYil", True
    KS_Buton ws, "ksK_Pdf", "PDF OLARAK KAYDET", ws.Range("L3").Left + 8, ws.Rows(3).Top + 4, 140, 24, "KS_PDF_KAYDET", "cizgi"
    ws.Rows(4).RowHeight = 10
    With ws.Range("B" & K_ILK)
        .Value = "Arama kutusuna malzeme ya da üretici yazýn. Örnek: 1000006162, REDRESOR, 1300001416, LUNA ELEKTRÝK."
        .Font.Italic = True
        .Font.Color = C_GRI()
    End With
    KS_ZeminUygula ws, True
    KS_Pencere ws, 90, 3
End Sub

' KARNE sayfasýnda arama kutusu deðiþince (Workbook_SheetChange) çaðrýlýr.
Public Sub KS_KarneDegisti(ByVal Target As Range)
    If Intersect(Target, KS_F("kAra").MergeArea) Is Nothing Then Exit Sub
    Application.EnableEvents = False
    KS_KarneCalistir ""
    Application.EnableEvents = True
End Sub

Public Sub KS_KARNE_GOSTER()
    KS_KarneCalistir ""
End Sub

Public Sub KS_KARNE_MALZEME()
    KS_KarneCalistir "MALZEME"
End Sub

Public Sub KS_KARNE_URETICI()
    KS_KarneCalistir "URETICI"
End Sub

Public Sub KS_KARNE_SIRALAMA()
    KS_KarneCalistir "SIRALAMA"
End Sub

Private Sub KS_KarneSekme(ByVal ws As Worksheet, ByVal aktif As Long)
    Dim i As Long, s As Shape
    For i = 0 To 2
        Set s = ws.Shapes("ksK_" & i)
        If i = aktif Then
            s.Fill.ForeColor.RGB = C_LAC()
            s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = vbWhite
            s.Line.Visible = msoFalse
        Else
            s.Fill.ForeColor.RGB = vbWhite
            s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = C_LAC()
            s.Line.Visible = msoTrue
        End If
    Next i
End Sub

' Üretici / sevkçi adaylarýný toplar: TANIMLAR, sevkiyat sevkçileri, malzeme üreticileri.
Private Function KS_UreticiAdaylari(ByVal metin As String) As Collection
    Dim c As New Collection, d As Object, a As Variant, i As Long, k As String, v As String, ara As String
    Set d = CreateObject("Scripting.Dictionary")
    ara = KS_Anahtar(metin)
    a = KS_Veri(KS_LOZ(T_SEVKCI))
    If IsArray(a) Then
        For i = 1 To UBound(a, 1)
            v = KS_Metin(a(i, 1))
            k = KS_FirmaAnahtar(v)
            If Len(k) > 0 And Not d.Exists(k) Then
                If InStr(1, KS_Anahtar(v), ara, vbBinaryCompare) > 0 Or KS_Anahtar(a(i, 2)) = ara Then
                    d(k) = v
                    c.Add v
                End If
            End If
        Next i
    End If
    a = KS_Veri(KS_LOZ(T_SEVK))
    If IsArray(a) Then
        For i = 1 To UBound(a, 1)
            v = KS_Metin(a(i, S_SEVKCI))
            k = KS_FirmaAnahtar(v)
            If Len(k) > 0 And Not d.Exists(k) And InStr(1, KS_Anahtar(v), ara, vbBinaryCompare) > 0 Then
                d(k) = v
                c.Add v
            End If
        Next i
    End If
    a = KS_Veri(KS_LOZ(T_MALZ))
    If IsArray(a) Then
        For i = 1 To UBound(a, 1)
            v = KS_Metin(a(i, M_URETICI))
            k = KS_FirmaAnahtar(v)
            If Len(k) > 0 And Not d.Exists(k) And InStr(1, KS_Anahtar(v), ara, vbBinaryCompare) > 0 Then
                d(k) = v
                c.Add v
            End If
        Next i
    End If
    Set KS_UreticiAdaylari = c
End Function

Private Sub KS_KarneCalistir(ByVal tur As String)
    Dim ws As Worksheet, ara As String, mi As Long, aday As Collection, korumali As Boolean, i As Long
    On Error GoTo Hata
    Set ws = KS_WS(SH_KARNE)
    ara = KS_Metin(KS_FV("kAra"))
    KS_Basla
    KS_HesaplaTumu
    KS_BaglamYukle
    korumali = ws.ProtectContents
    KS_Ac ws
    KS_AlanTemizle ws, K_ILK
    If tur = "SIRALAMA" Then
        KS_KarneSekme ws, 2
        KS_SiralamaYaz ws
    ElseIf Len(ara) = 0 Then
        ws.Range("B" & K_ILK).Value = "Arama kutusuna malzeme ya da üretici yazýn."
        ws.Range("B" & K_ILK).Font.Color = C_GRI()
    Else
        If tur <> "URETICI" Then mi = KS_MalzemeAra(ara)
        If mi > 0 Then
            KS_KarneSekme ws, 0
            KS_MalzemeKarnesi ws, mi
        Else
            Set aday = KS_UreticiAdaylari(ara)
            If aday.Count = 1 And tur <> "MALZEME" Then
                KS_KarneSekme ws, 1
                KS_UreticiKarnesi ws, aday(1)
            ElseIf aday.Count > 1 And tur <> "MALZEME" Then
                KS_KarneSekme ws, 1
                KS_Bolum ws.Range(ws.Cells(K_ILK, 2), ws.Cells(K_ILK, 14)), aday.Count & " ÜRETÝCÝ / SEVKÇÝ EÞLEÞTÝ  -  birini kopyalayýp arama kutusuna yapýþtýrýn"
                For i = 1 To aday.Count
                    If i > 40 Then Exit For
                    ws.Cells(K_ILK + i, 2).Value = aday(i)
                Next i
            ElseIf mi < 0 Then
                ws.Range("B" & K_ILK).Value = -mi & " malzeme bu ifadeyi içeriyor. Daha ayrýntýlý yazýn ya da kodu girin."
            Else
                ws.Range("B" & K_ILK).Value = "Eþleþen malzeme ya da üretici bulunamadý: " & ara
                ws.Range("B" & K_ILK).Font.Color = C_KIRMIZI()
            End If
        End If
    End If
    If korumali Then KS_Kapat ws
    KS_Bitir
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Karne oluþturulamadý: " & Err.Description, vbExclamation, KS_AD
End Sub

Private Sub KS_KarneBaslik(ByVal ws As Worksheet, ByVal baslik As String, ByVal bilgi As String)
    ws.Range(ws.Cells(K_ILK, 2), ws.Cells(K_ILK, 14)).Merge
    With ws.Cells(K_ILK, 2)
        .Value = baslik
        .Font.Name = KS_FONT_KALIN
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = C_LAC()
        .Font.Italic = False
    End With
    ws.Rows(K_ILK).RowHeight = 30
    ws.Range(ws.Cells(K_ILK + 1, 2), ws.Cells(K_ILK + 1, 14)).Merge
    With ws.Cells(K_ILK + 1, 2)
        .Value = bilgi
        .Font.Size = 9
        .Font.Color = C_GRI()
    End With
    With ws.Range(ws.Cells(K_ILK + 1, 2), ws.Cells(K_ILK + 1, 14)).Borders(xlEdgeBottom)
        .Color = C_CAMGOBEGI()
        .Weight = xlMedium
    End With
    ws.Rows(K_ILK + 2).RowHeight = 10
    ws.Rows(K_ILK + 3).RowHeight = 18
    ws.Rows(K_ILK + 4).RowHeight = 30
    ws.Rows(K_ILK + 5).RowHeight = 16
End Sub

Private Sub KS_Kart2(ByVal ws As Worksheet, ByVal sira As Long, ByVal etiket As String, ByVal deger As Variant, _
        ByVal bicim As String, ByVal alt As String, Optional ByVal renk As Long = -1)
    Dim r As Range
    Set r = ws.Range(ws.Cells(K_ILK + 3, 2 + sira * 2), ws.Cells(K_ILK + 5, 3 + sira * 2))
    KS_KartCiz r, etiket
    KS_KartYaz r, deger, bicim, alt, renk
End Sub

Private Sub KS_ListeTablo(ByVal ws As Worksheet, ByRef r As Long, ByVal baslik As String, ByVal basliklar As Variant, _
        ByRef liste() As Variant, ByVal n As Long, ByVal siraKol As Long, ByVal azalan As Boolean, ByVal bicimler As Variant, _
        ByVal bos As String)
    Dim i As Long, k As Long, m As Long, idx As Variant
    m = UBound(basliklar) + 1
    KS_Bolum ws.Range(ws.Cells(r, 2), ws.Cells(r, 14)), baslik
    KS_TabloBaslik ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 1 + m)), basliklar
    r = r + 2
    If n = 0 Then
        ws.Cells(r, 2).Value = bos
        ws.Cells(r, 2).Font.Italic = True
        ws.Cells(r, 2).Font.Color = C_GRI()
        r = r + 2
        Exit Sub
    End If
    idx = KS_SiraIndeks(liste, siraKol, azalan, n)
    For i = 1 To n
        For k = 1 To m
            ws.Cells(r + i - 1, 1 + k).Value = liste(idx(i), k)
        Next k
    Next i
    KS_TabloGovde ws.Range(ws.Cells(r, 2), ws.Cells(r + n - 1, 1 + m))
    For k = 1 To m
        If Len(bicimler(k - 1)) > 0 Then ws.Range(ws.Cells(r, 1 + k), ws.Cells(r + n - 1, 1 + k)).NumberFormat = bicimler(k - 1)
    Next k
    ws.Range(ws.Cells(r, 2), ws.Cells(r + n - 1, 1 + m)).ShrinkToFit = True
    r = r + n + 1
End Sub

Private Sub KS_MalzemeKarnesi(ByVal ws As Worksheet, ByVal mi As Long)
    Dim aM As Variant, arS As Variant, aA As Variant, kod As String, i As Long, n As Long, r As Long, x As Double
    Dim liste() As Variant, sevkMik As Double, sevkTut As Double, arzN As Long, arzMik As Double, donTop As Double
    Dim donN As Long, mal As Double, garIc As Long, acik As Long, d As Date, tr As String
    tr = KS_FmtTarih()
    aM = KS_Veri(KS_LOZ(T_MALZ))
    arS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    kod = KS_Anahtar(aM(mi, M_KOD))
    KS_KarneBaslik ws, "MALZEME   " & KS_Metin(aM(mi, M_KOD)) & "   " & KS_Metin(aM(mi, M_AD)), _
        "Kategori: " & KS_Bilgi(aM(mi, M_KAT)) & "     |     Birim: " & KS_Bilgi(aM(mi, M_BIRIM)) & _
        "     |     Üretici: " & KS_Bilgi(aM(mi, M_URETICI)) & "     |     Liste fiyatý: " & _
        IIf(KS_Bos(aM(mi, M_FIYAT)), "-", Format$(aM(mi, M_FIYAT), "#,##0.00") & " " & KS_Metin(aM(mi, M_PB))) & _
        "     |     Garanti: " & IIf(KS_Bos(aM(mi, M_GARANTI)), "GÝRÝLMEMÝÞ", KS_Metin(aM(mi, M_GARANTI)) & " ay")

    r = K_ILK + 7
    n = 0
    If IsArray(arS) Then
        ReDim liste(1 To UBound(arS, 1), 1 To 12)
        For i = 1 To UBound(arS, 1)
            If KS_Anahtar(arS(i, S_KOD)) = kod Then
                n = n + 1
                liste(n, 1) = arS(i, S_NO): liste(n, 2) = arS(i, S_SEVKCI): liste(n, 3) = arS(i, S_YER)
                liste(n, 4) = arS(i, S_ADET): liste(n, 5) = arS(i, S_TUTAR): liste(n, 6) = arS(i, S_PLAN)
                liste(n, 7) = arS(i, S_GELIS): liste(n, 8) = arS(i, S_TDURUM): liste(n, 9) = arS(i, S_GARBIT)
                liste(n, 10) = arS(i, S_GDURUM): liste(n, 11) = arS(i, S_ARIZA)
                If KS_Tarih(arS(i, S_GELIS), d) Then liste(n, 12) = d Else liste(n, 12) = arS(i, S_PLAN)
                If KS_Sayi(arS(i, S_ADET), x) Then sevkMik = sevkMik + x
                If KS_Sayi(arS(i, S_TUTAR), x) Then sevkTut = sevkTut + x
            End If
        Next i
    Else
        ReDim liste(1 To 1, 1 To 12)
    End If
    KS_ListeTablo ws, r, "SEVKÝYAT GEÇMÝÞÝ  (yeniden eskiye)", Array("Sevk No", "Sevkçi", "Sevk Yeri", "Adet", _
        "Tutar (TL)", "Planlanan", "Geliþ", "Teslim", "Garanti Bitiþ", "Garanti", "Arýzalý"), liste, n, 12, True, _
        Array("", "", "", "#,##0", "#,##0", tr, tr, "", tr, "", "#,##0"), "Bu malzeme için sevkiyat kaydý yok."

    n = 0
    If IsArray(aA) Then
        ReDim liste(1 To UBound(aA, 1), 1 To 11)
        For i = 1 To UBound(aA, 1)
            If KS_Anahtar(aA(i, A_KOD)) = kod Then
                n = n + 1
                liste(n, 1) = aA(i, A_NO): liste(n, 2) = aA(i, A_TARIH): liste(n, 3) = aA(i, A_SERI)
                liste(n, 4) = aA(i, A_ONCEKI): liste(n, 5) = aA(i, A_ADET): liste(n, 6) = aA(i, A_KAPSAM)
                liste(n, 7) = aA(i, A_ACIK): liste(n, 8) = aA(i, A_BILDIRIM): liste(n, 9) = aA(i, A_DONUSGUN)
                liste(n, 10) = aA(i, A_MALIYET): liste(n, 11) = aA(i, A_DURUM)
                arzN = arzN + 1
                If KS_Sayi(aA(i, A_ADET), x) Then arzMik = arzMik + x
                If KS_Sayi(aA(i, A_DONUSGUN), x) Then donTop = donTop + x: donN = donN + 1
                If KS_Sayi(aA(i, A_MALIYET), x) Then mal = mal + x
                If KS_Metin(aA(i, A_KAPSAM)) = "EVET" Then garIc = garIc + 1
                If KS_ArizaAcik(aA(i, A_DURUM)) Then acik = acik + 1
            End If
        Next i
    Else
        ReDim liste(1 To 1, 1 To 11)
    End If
    KS_ListeTablo ws, r, "ARIZA GEÇMÝÞÝ  (eskiden yeniye)", Array("Arýza No", "Arýza Tarihi", "Seri No", _
        "Önceki Arýza", "Adet", "Garanti", "Sebep / Açýklama", "Bildirim", "Dönüþ (gün)", "Maliyet (TL)", "Durum"), _
        liste, n, 2, False, Array("", tr, "", tr, "#,##0", "", "", tr, "0", "#,##0", ""), _
        "Bu malzeme için arýza kaydý yok."

    KS_Kart2 ws, 0, "SEVK EDÝLEN", sevkMik, "#,##0", KS_Bilgi(aM(mi, M_BIRIM))
    KS_Kart2 ws, 1, "SEVK TUTARI (TL)", sevkTut, "#,##0", "tüm sevkiyatlar"
    KS_Kart2 ws, 2, "ARIZA KAYDI", arzN, "#,##0", garIc & " garanti içi  |  " & acik & " açýk", _
        IIf(arzN > 0, C_KIRMIZI(), C_YESIL())
    KS_Kart2 ws, 3, "ARIZA ORANI", KS_Bol(arzMik, sevkMik), "0.0%", "arýzalý / sevk edilen"
    KS_Kart2 ws, 4, "ORT. ÜRETÝCÝ DÖNÜÞÜ", KS_Bol(donTop, donN), "0.0"" gün""", donN & " dönüþ kaydý"
    KS_Kart2 ws, 5, "TAMÝR MALÝYETÝ (TL)", mal, "#,##0", "toplam"
End Sub

Private Sub KS_UreticiKarnesi(ByVal ws As Worksheet, ByVal ad As String)
    Dim aM As Variant, arS As Variant, aA As Variant, k As String, i As Long, n As Long, r As Long, x As Double
    Dim dM As Object, dSec As Object, liste() As Variant, j As Long, kalem As Long, tut As Double, zam As Long, tes As Long
    Dim gecTop As Double, arzN As Long, arzMik As Double, sevkMik As Double, donTop As Double, donN As Long, mal As Double
    Dim acik As Long, skor As Double, oran As Double, gec As Double, don As Double, tr As String, kk As String
    tr = KS_FmtTarih()
    k = KS_FirmaAnahtar(ad)
    aM = KS_Veri(KS_LOZ(T_MALZ))
    arS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    Set dM = CreateObject("Scripting.Dictionary")
    If IsArray(aM) Then
        For i = 1 To UBound(aM, 1)
            dM(KS_Anahtar(aM(i, M_KOD))) = i
        Next i
    End If
    ' Malzeme bazýnda topla
    Set dSec = CreateObject("Scripting.Dictionary")
    n = 0
    ReDim liste(1 To KS_Satir(arS) + KS_Satir(aA) + 1, 1 To 8)
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            kk = KS_Anahtar(arS(i, S_KOD))
            If KS_FirmaAnahtar(arS(i, S_SEVKCI)) = k Or KS_UreticiMi(aM, dM, kk, k) Then
                If Not dSec.Exists(kk) Then
                    n = n + 1
                    dSec(kk) = n
                    liste(n, 1) = arS(i, S_KOD): liste(n, 2) = arS(i, S_AD)
                    For j = 3 To 7
                        liste(n, j) = 0
                    Next j
                End If
                j = dSec(kk)
                liste(j, 3) = liste(j, 3) + 1
                If KS_Sayi(arS(i, S_ADET), x) Then liste(j, 4) = liste(j, 4) + x: sevkMik = sevkMik + x
                If KS_Sayi(arS(i, S_TUTAR), x) Then liste(j, 5) = liste(j, 5) + x: tut = tut + x
                kalem = kalem + 1
                Select Case KS_Metin(arS(i, S_TDURUM))
                    Case "ZAMANINDA": zam = zam + 1: tes = tes + 1
                    Case "GECÝKMELÝ"
                        tes = tes + 1
                        If KS_Sayi(arS(i, S_GECIKME), x) Then gecTop = gecTop + x
                End Select
            End If
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            kk = KS_Anahtar(aA(i, A_KOD))
            If KS_FirmaAnahtar(aA(i, A_SEVKCI)) = k Or KS_FirmaAnahtar(aA(i, A_URETICI)) = k Then
                If Not dSec.Exists(kk) Then
                    n = n + 1
                    dSec(kk) = n
                    liste(n, 1) = aA(i, A_KOD): liste(n, 2) = aA(i, A_AD)
                    For j = 3 To 7
                        liste(n, j) = 0
                    Next j
                End If
                j = dSec(kk)
                liste(j, 6) = liste(j, 6) + 1
                arzN = arzN + 1
                If KS_Sayi(aA(i, A_ADET), x) Then liste(j, 7) = liste(j, 7) + x: arzMik = arzMik + x
                If KS_Sayi(aA(i, A_DONUSGUN), x) Then donTop = donTop + x: donN = donN + 1
                If KS_Sayi(aA(i, A_MALIYET), x) Then mal = mal + x
                If KS_ArizaAcik(aA(i, A_DURUM)) Then acik = acik + 1
            End If
        Next i
    End If
    For j = 1 To n
        If liste(j, 4) > 0 Then liste(j, 8) = liste(j, 7) / liste(j, 4) Else liste(j, 8) = Empty
    Next j

    KS_KarneBaslik ws, "ÜRETÝCÝ / SEVKÇÝ   " & KS_FirmaAd(ad), "Bu üreticinin sevk ettiði ya da ürettiði " & n & " malzeme, " & kalem & _
        " sevkiyat kalemi, " & arzN & " arýza kaydý."
    r = K_ILK + 7
    KS_ListeTablo ws, r, "MALZEMELERÝ  (arýza sayýsýna göre)", Array("Malzeme Kodu", "Malzeme Adý", "Sevk Kalemi", _
        "Sevk Miktarý", "Tutar (TL)", "Arýza", "Arýzalý Miktar", "Arýza Oraný"), liste, n, 6, True, _
        Array("", "", "#,##0", "#,##0", "#,##0", "#,##0", "#,##0", "0.0%"), "Kayýt yok."

    n = 0
    If IsArray(aA) Then
        ReDim liste(1 To UBound(aA, 1), 1 To 10)
        For i = 1 To UBound(aA, 1)
            If KS_FirmaAnahtar(aA(i, A_SEVKCI)) = k Or KS_FirmaAnahtar(aA(i, A_URETICI)) = k Then
                n = n + 1
                liste(n, 1) = aA(i, A_NO): liste(n, 2) = aA(i, A_TARIH): liste(n, 3) = aA(i, A_KOD)
                liste(n, 4) = aA(i, A_AD): liste(n, 5) = aA(i, A_SERI): liste(n, 6) = aA(i, A_KAPSAM)
                liste(n, 7) = aA(i, A_ACIK): liste(n, 8) = aA(i, A_BILDIRIM): liste(n, 9) = aA(i, A_DONUSGUN)
                liste(n, 10) = aA(i, A_DURUM)
            End If
        Next i
    Else
        ReDim liste(1 To 1, 1 To 10)
    End If
    KS_ListeTablo ws, r, "ARIZA GEÇMÝÞÝ  (yeniden eskiye)", Array("Arýza No", "Tarih", "Malzeme Kodu", "Malzeme Adý", _
        "Seri No", "Garanti", "Sebep / Açýklama", "Bildirim", "Dönüþ (gün)", "Durum"), liste, n, 2, True, _
        Array("", tr, "", "", "", "", "", tr, "0", ""), "Bu üretici için arýza kaydý yok."

    If sevkMik > 0 Then oran = arzMik / sevkMik * 100
    If tes > 0 Then gec = gecTop / tes
    If donN > 0 Then don = donTop / donN
    skor = 100 - KS_Min(50, oran * KS_ParamSayi("pKatAriza", 5)) - KS_Min(25, gec * KS_ParamSayi("pKatGecikme", 1)) - _
        KS_Min(25, don * KS_ParamSayi("pKatDonus", 0.5))
    If skor < 0 Then skor = 0
    KS_Kart2 ws, 0, "SEVK KALEMÝ", kalem, "#,##0", Format$(sevkMik, "#,##0") & " birim"
    KS_Kart2 ws, 1, "SEVK TUTARI (TL)", tut, "#,##0", "tüm sevkiyatlar"
    KS_Kart2 ws, 2, "ZAMANINDA TESLÝM", KS_Bol(zam, tes), "0%", tes & " teslim  |  ort. gecikme " & Format$(gec, "0.0") & " gün"
    KS_Kart2 ws, 3, "ARIZA KAYDI", arzN, "#,##0", acik & " açýk  |  oran " & Format$(oran / 100, "0.0%"), _
        IIf(arzN > 0, C_KIRMIZI(), C_YESIL())
    KS_Kart2 ws, 4, "ORT. DÖNÜÞ SÜRESÝ", KS_Bol(donTop, donN), "0.0"" gün""", donN & " dönüþ kaydý"
    KS_Kart2 ws, 5, "KALÝTE SKORU", Round(skor, 1), "0.0", "Not: " & KS_Harf(skor) & "   (TANIMLAR katsayýlarý)", _
        IIf(skor >= 70, C_YESIL(), IIf(skor >= 50, C_TURUNCU(), C_KIRMIZI()))
End Sub

Private Function KS_UreticiMi(ByRef aM As Variant, ByVal dM As Object, ByVal kod As String, ByVal k As String) As Boolean
    Dim mi As Long
    If dM.Exists(kod) Then
        mi = dM(kod)
        KS_UreticiMi = (KS_FirmaAnahtar(aM(mi, M_URETICI)) = k)
    End If
End Function

Private Function KS_Bilgi(ByVal v As Variant) As String
    KS_Bilgi = KS_Metin(v)
    If Len(KS_Bilgi) = 0 Then KS_Bilgi = "-"
End Function

' Çýktý alanýný (satýr ilk..) biçim dahil temizler.
Public Sub KS_AlanTemizle(ByVal ws As Worksheet, ByVal ilk As Long)
    Dim r As Range
    Set r = ws.Range(ws.Rows(ilk), ws.Rows(ilk + 3000))
    r.FormatConditions.Delete
    r.UnMerge
    r.Clear
    r.Interior.Pattern = xlNone
    r.Font.Name = KS_FONT
    r.Font.Size = 10
    r.Font.Color = C_METIN()
    r.RowHeight = 20
    r.VerticalAlignment = xlCenter
End Sub

' ---------------------------------------------------------------- tedarikçi sýralamasý

Private Sub KS_SiralamaYaz(ByVal ws As Worksheet)
    Dim aM As Variant, arS As Variant, aA As Variant, dG As Object
    Dim yilS As String, y1 As Long, y2 As Long, i As Long, j As Long, n As Long, g As Long
    Dim k As String, x As Double, d As Date, t() As Variant, idx As Variant, r As Long
    Dim k1 As Double, k2 As Double, k3 As Double, skor As Double, oran As Double, gec As Double, don As Double
    yilS = KS_Metin(KS_FV("kYil"))
    If IsNumeric(yilS) And Len(yilS) = 4 Then
        y1 = CLng(yilS): y2 = y1
    Else
        y1 = 0: y2 = 9999
    End If
    k1 = KS_ParamSayi("pKatAriza", 5)
    k2 = KS_ParamSayi("pKatGecikme", 1)
    k3 = KS_ParamSayi("pKatDonus", 0.5)
    aM = KS_Veri(KS_LOZ(T_MALZ))
    arS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    Set dG = CreateObject("Scripting.Dictionary")
    ReDim t(1 To KS_Satir(arS) + KS_Satir(aA) + 1, 1 To 16)
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            j = KS_SevkYil(arS, i)
            If j >= y1 And j <= y2 Then
                g = KS_GrupNo(dG, t, n, KS_Metin(arS(i, S_SEVKCI)))
                t(g, 2) = t(g, 2) + 1
                If KS_Sayi(arS(i, S_ADET), x) Then t(g, 3) = t(g, 3) + x
                If KS_Sayi(arS(i, S_TUTAR), x) Then t(g, 4) = t(g, 4) + x
                Select Case KS_Metin(arS(i, S_TDURUM))
                    Case "ZAMANINDA"
                        t(g, 5) = t(g, 5) + 1
                        t(g, 6) = t(g, 6) + 1
                    Case "GECÝKMELÝ"
                        t(g, 6) = t(g, 6) + 1
                        If KS_Sayi(arS(i, S_GECIKME), x) Then t(g, 7) = t(g, 7) + x
                End Select
            End If
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Tarih(aA(i, A_TARIH), d) Then
                If Year(d) >= y1 And Year(d) <= y2 Then
                    k = KS_Metin(aA(i, A_SEVKCI))
                    If Len(k) = 0 Then k = KS_Metin(aA(i, A_URETICI))
                    g = KS_GrupNo(dG, t, n, k)
                    t(g, 8) = t(g, 8) + 1
                    If KS_Sayi(aA(i, A_ADET), x) Then t(g, 9) = t(g, 9) + x
                    If KS_Metin(aA(i, A_KAPSAM)) = "EVET" Then t(g, 10) = t(g, 10) + 1
                    If KS_Sayi(aA(i, A_DONUSGUN), x) Then
                        t(g, 11) = t(g, 11) + x
                        t(g, 12) = t(g, 12) + 1
                        If x > t(g, 13) Then t(g, 13) = x
                    End If
                    If KS_ArizaAcik(aA(i, A_DURUM)) Then t(g, 14) = t(g, 14) + 1
                    If KS_Sayi(aA(i, A_MALIYET), x) Then t(g, 15) = t(g, 15) + x
                End If
            End If
        Next i
    End If
    For g = 1 To n
        oran = 0: gec = 0: don = 0
        If t(g, 3) > 0 Then oran = t(g, 9) / t(g, 3) * 100
        If t(g, 6) > 0 Then gec = t(g, 7) / t(g, 6)
        If t(g, 12) > 0 Then don = t(g, 11) / t(g, 12)
        skor = 100 - KS_Min(50, oran * k1) - KS_Min(25, gec * k2) - KS_Min(25, don * k3)
        If skor < 0 Then skor = 0
        t(g, 16) = Round(skor, 1)
    Next g

    KS_Bolum ws.Range(ws.Cells(K_ILK, 2), ws.Cells(K_ILK, 14)), "TEDARÝKÇÝ KALÝTE SIRALAMASI  -  " & _
        IIf(y1 = 0, "TÜM YILLAR", CStr(y1))
    KS_TabloBaslik ws.Range(ws.Cells(K_ILK + 1, 2), ws.Cells(K_ILK + 1, 14)), Array("Sevkçi / Üretici", "Sevk Kalemi", _
        "Sevk Tutarý (TL)", "Zamanýnda", "Ort. Gecikme", "Arýza", "Arýza Oraný", "Garanti Ýçi", "Ort. Dönüþ", _
        "Açýk Arýza", "Tamir (TL)", "Skor", "Not")
    r = K_ILK + 2
    If n = 0 Then
        ws.Cells(r, 2).Value = "Seçilen dönemde kayýt yok."
        ws.Cells(r, 2).Font.Color = C_GRI()
        Exit Sub
    End If
    idx = KS_SiraIndeks(t, 16, True, n)
    For i = 1 To n
        g = idx(i)
        ws.Cells(r, 2).Value = t(g, 1)
        ws.Cells(r, 3).Value = t(g, 2)
        ws.Cells(r, 4).Value = t(g, 4)
        ws.Cells(r, 5).Value = KS_Bol(t(g, 5), t(g, 6))
        ws.Cells(r, 6).Value = KS_Bol(t(g, 7), t(g, 6))
        ws.Cells(r, 7).Value = t(g, 8)
        ws.Cells(r, 8).Value = KS_Bol(t(g, 9), t(g, 3))
        ws.Cells(r, 9).Value = t(g, 10)
        ws.Cells(r, 10).Value = KS_Bol(t(g, 11), t(g, 12))
        ws.Cells(r, 11).Value = t(g, 14)
        ws.Cells(r, 12).Value = t(g, 15)
        ws.Cells(r, 13).Value = t(g, 16)
        ws.Cells(r, 14).Value = KS_Harf(t(g, 16))
        r = r + 1
    Next i
    KS_TabloGovde ws.Range(ws.Cells(K_ILK + 2, 2), ws.Cells(r - 1, 14))
    ws.Range(ws.Cells(K_ILK + 2, 2), ws.Cells(r - 1, 2)).ShrinkToFit = True
    ws.Range(ws.Cells(K_ILK + 2, 3), ws.Cells(r - 1, 14)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(K_ILK + 2, 3), ws.Cells(r - 1, 4)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(K_ILK + 2, 5), ws.Cells(r - 1, 5)).NumberFormat = "0%"
    ws.Range(ws.Cells(K_ILK + 2, 6), ws.Cells(r - 1, 6)).NumberFormat = "0.0"
    ws.Range(ws.Cells(K_ILK + 2, 7), ws.Cells(r - 1, 7)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(K_ILK + 2, 8), ws.Cells(r - 1, 8)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(K_ILK + 2, 10), ws.Cells(r - 1, 10)).NumberFormat = "0.0"
    ws.Range(ws.Cells(K_ILK + 2, 12), ws.Cells(r - 1, 12)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(K_ILK + 2, 13), ws.Cells(r - 1, 13)).NumberFormat = "0.0"
    ws.Range(ws.Cells(K_ILK + 2, 13), ws.Cells(r - 1, 14)).Font.Bold = True
    KS_SiraRenk ws.Range(ws.Cells(K_ILK + 2, 8), ws.Cells(r - 1, 8)), ws.Range(ws.Cells(K_ILK + 2, 13), ws.Cells(r - 1, 13)), _
        ws.Range(ws.Cells(K_ILK + 2, 14), ws.Cells(r - 1, 14))
    r = r + 1
    With ws.Cells(r, 2)
        .Value = "Skor = 100 - en çok 50 (arýza oraný % x " & k1 & ") - en çok 25 (ort. gecikme gün x " & k2 & _
            ") - en çok 25 (ort. üretici dönüþ gün x " & k3 & ").   Not: A >= 85, B >= 70, C >= 50, D < 50."
        .Font.Size = 8
        .Font.Italic = True
        .Font.Color = C_GRI()
    End With
End Sub

Private Function KS_GrupNo(ByVal d As Object, ByRef t() As Variant, ByRef n As Long, ByVal ad As String) As Long
    Dim k As String, j As Long
    If Len(ad) = 0 Then ad = "(Belirtilmemiþ)"
    k = KS_FirmaAnahtar(ad)
    If Not d.Exists(k) Then
        n = n + 1
        d(k) = n
        t(n, 1) = KS_FirmaAd(ad)
        For j = 2 To 16
            t(n, j) = 0
        Next j
    End If
    KS_GrupNo = d(k)
End Function

Private Function KS_Min(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then KS_Min = a Else KS_Min = b
End Function

Public Function KS_Harf(ByVal skor As Double) As String
    If skor >= 85 Then
        KS_Harf = "A"
    ElseIf skor >= 70 Then
        KS_Harf = "B"
    ElseIf skor >= 50 Then
        KS_Harf = "C"
    Else
        KS_Harf = "D"
    End If
End Function

Private Sub KS_SiraRenk(ByVal oran As Range, ByVal skor As Range, ByVal harf As Range)
    Dim cs As Object, db As Object, fc As Object
    On Error Resume Next
    Set db = oran.FormatConditions.AddDatabar
    db.BarColor.Color = RGB(255, 140, 150)
    Set cs = skor.FormatConditions.AddColorScale(3)
    cs.ColorScaleCriteria(1).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(1).Value = 40
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(253, 210, 214)
    cs.ColorScaleCriteria(2).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(2).Value = 70
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(255, 243, 205)
    cs.ColorScaleCriteria(3).Type = xlConditionValueNumber
    cs.ColorScaleCriteria(3).Value = 100
    cs.ColorScaleCriteria(3).FormatColor.Color = RGB(209, 245, 228)
    Set fc = harf.FormatConditions.Add(Type:=xlTextString, String:="A", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_YESIL()
    Set fc = harf.FormatConditions.Add(Type:=xlTextString, String:="B", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_TEAL()
    Set fc = harf.FormatConditions.Add(Type:=xlTextString, String:="C", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_TURUNCU()
    Set fc = harf.FormatConditions.Add(Type:=xlTextString, String:="D", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_KIRMIZI()
    On Error GoTo 0
End Sub
