Attribute VB_Name = "modKS_PANEL"
Option Explicit

' PANEL: 10 KPI kartý, 4 canlý liste ve 2 grafik. Grafik verisi 5 YILLIK RAPOR sayfasýndan gelir.

Private Const P_KART1 As Long = 9
Private Const P_KART2 As Long = 13
Private Const P_LISTE1 As Long = 17
Private Const P_LISTE2 As Long = 28
Private Const P_GRAFIK As Long = 39
Private Const P_SATIR As Long = 8   ' her listede gösterilen satýr

Private Function KS_KartAlan(ByVal ws As Worksheet, ByVal ust As Long, ByVal sira As Long) As Range
    Set KS_KartAlan = ws.Range(ws.Cells(ust, 2 + sira * 3), ws.Cells(ust + 2, 4 + sira * 3))
End Function

Public Sub KS_PanelKur()
    Dim ws As Worksheet, i As Long, et As Variant
    Set ws = KS_WS(SH_PANEL)
    KS_SayfaSifirla ws
    ' I sütunu iki liste arasýndaki boþluktur; H ve J geniþletilerek 5 kartýn eþit geniþlikte kalmasý saðlanýr.
    ws.Columns("B:P").ColumnWidth = 10.6
    ws.Columns("H").ColumnWidth = 14.4
    ws.Columns("I").ColumnWidth = 3
    ws.Columns("J").ColumnWidth = 14.4
    ws.Columns("Q").ColumnWidth = 3
    KS_Bant ws, 17, "KONTROL PANELÝ", "Dönem seçin, YENÝLE'ye basýn. Güncel durum kartlarý dönemden baðýmsýzdýr."

    ' Filtre çubuðu
    ws.Rows(7).RowHeight = 26
    With ws.Range("B7")
        .Value = "DÖNEM"
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlRight
    End With
    ws.Range("C7:D7").Merge
    With ws.Range("C7:D7")
        .Value = "TÜMÜ"
        .Interior.Color = C_GIRIS()
        .Font.Bold = True
        .Font.Color = C_LAC()
        .HorizontalAlignment = xlCenter
        .Locked = False
        .Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
        .Borders(xlEdgeBottom).Weight = xlMedium
    End With
    KS_AdHucre "pYil", ws.Range("C7")
    KS_Liste ws.Range("C7:D7"), "lstYil", True
    KS_Buton ws, "ksBtnYenile", "YENÝLE", ws.Range("E7").Left + 4, ws.Range("E7").Top + 2, 92, 22, "KS_PANEL_YENILE", "ana"
    KS_Buton ws, "ksBtnYedek", "YEDEK AL", ws.Range("F7").Left + 40, ws.Range("F7").Top + 2, 92, 22, "KS_YEDEK_AL", "cizgi"
    ws.Range("J7:P7").Merge
    With ws.Range("J7")
        .HorizontalAlignment = xlRight
        .Font.Size = 8
        .Font.Color = C_GRI()
    End With
    KS_AdHucre "pGuncel", ws.Range("J7")
    ws.Rows(8).RowHeight = 8

    ' KPI kartlarý: üst sýra akýþ (döneme göre), alt sýra kalite ve güncel durum.
    et = Array("SEVK KALEMÝ", "SEVK TUTARI", "ZAMANINDA TESLÝM", "BEKLEYEN TESLÝMAT  ·  GÜNCEL", _
        "GARANTÝSÝ BÝTÝYOR  ·  GÜNCEL", "ARIZA KAYDI", "GARANTÝ ÝÇÝ ARIZA", "AÇIK ARIZA  ·  GÜNCEL", _
        "ORT. ÜRETÝCÝ DÖNÜÞÜ", "TAMÝR MALÝYETÝ")
    For i = 0 To 4
        KS_KartCiz KS_KartAlan(ws, P_KART1, i), CStr(et(i))
        KS_KartCiz KS_KartAlan(ws, P_KART2, i), CStr(et(i + 5))
    Next i
    For i = 0 To 1
        ws.Rows(P_KART1 + i * 4).RowHeight = 20
        ws.Rows(P_KART1 + i * 4 + 1).RowHeight = 34
        ws.Rows(P_KART1 + i * 4 + 2).RowHeight = 18
    Next i
    ws.Rows(P_KART1 + 3).RowHeight = 8
    ws.Rows(P_KART2 + 3).RowHeight = 12

    ' Listeler
    KS_PanelListe ws, P_LISTE1, 2, "EN ÇOK ARIZA VEREN MALZEMELER", _
        Array("Malzeme Kodu", "Malzeme Adý", "", "", "Arýza", "Arýzalý Adet", "Maliyet")
    KS_PanelListe ws, P_LISTE1, 10, "AÇIK / ÜRETÝCÝDE BEKLEYEN ARIZALAR", _
        Array("Arýza No", "Malzeme", "", "", "Durum", "Bekleme (gün)", "Sevk Yeri")
    KS_PanelListe ws, P_LISTE2, 2, "GARANTÝSÝ YAKINDA BÝTECEK SEVKÝYATLAR", _
        Array("Sevk No", "Malzeme", "", "", "Sevk Yeri", "Garanti Bitiþ", "Kalan (gün)")
    KS_PanelListe ws, P_LISTE2, 10, "TESLÝMATI GECÝKEN SEVKÝYATLAR", _
        Array("Sevk No", "Malzeme", "", "", "Sevkçi", "Planlanan", "Gecikme (gün)")

    ' Grafikler
    KS_Bolum ws.Range(ws.Cells(P_GRAFIK, 2), ws.Cells(P_GRAFIK, 8)), "5 YILLIK ARIZA EÐÝLÝMÝ"
    KS_Bolum ws.Range(ws.Cells(P_GRAFIK, 10), ws.Cells(P_GRAFIK, 16)), "AYLIK ARIZA VE TAMÝR MALÝYETÝ"
    KS_GrafikKur ws
    KS_Pencere ws, 85, 0, False
End Sub

Private Sub KS_PanelListe(ByVal ws As Worksheet, ByVal ust As Long, ByVal kol As Long, ByVal baslik As String, _
        ByVal basliklar As Variant)
    Dim r As Long
    KS_Bolum ws.Range(ws.Cells(ust, kol), ws.Cells(ust, kol + 6)), baslik
    KS_TabloBaslik ws.Range(ws.Cells(ust + 1, kol), ws.Cells(ust + 1, kol + 6)), basliklar
    ws.Range(ws.Cells(ust + 1, kol + 1), ws.Cells(ust + 1, kol + 3)).Merge
    For r = ust + 2 To ust + 1 + P_SATIR
        ws.Range(ws.Cells(r, kol + 1), ws.Cells(r, kol + 3)).Merge
        ws.Rows(r).RowHeight = 19
    Next r
    KS_TabloGovde ws.Range(ws.Cells(ust + 2, kol), ws.Cells(ust + 1 + P_SATIR, kol + 6))
    ws.Range(ws.Cells(ust + 2, kol + 1), ws.Cells(ust + 1 + P_SATIR, kol + 3)).ShrinkToFit = True
    ws.Range(ws.Cells(ust + 2, kol + 4), ws.Cells(ust + 1 + P_SATIR, kol + 4)).ShrinkToFit = True
    ws.Range(ws.Cells(ust + 2, kol), ws.Cells(ust + 1 + P_SATIR, kol)).Font.Bold = True
    ws.Range(ws.Cells(ust + 2, kol), ws.Cells(ust + 1 + P_SATIR, kol)).Font.Color = C_LAC()
    ws.Range(ws.Cells(ust + 2, kol + 5), ws.Cells(ust + 1 + P_SATIR, kol + 6)).HorizontalAlignment = xlCenter
End Sub

Private Sub KS_GrafikKur(ByVal ws As Worksheet)
    Dim rp As Worksheet, co As ChartObject, s As Series, ust As Double, yuk As Double
    Set rp = KS_WS(SH_RAPOR)
    ust = ws.Cells(P_GRAFIK + 1, 2).Top + 4
    yuk = 250

    Set co = ws.ChartObjects.Add(ws.Cells(P_GRAFIK + 1, 2).Left, ust, _
        ws.Cells(P_GRAFIK + 1, 9).Left - ws.Cells(P_GRAFIK + 1, 2).Left, yuk)
    co.Name = "ksGrafikYil"
    With co.Chart
        .ChartType = xlColumnStacked
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = "Garanti içi"
        s.Values = rp.Range(rp.Cells(R_GARIC, 3), rp.Cells(R_GARIC, 7))
        s.XValues = rp.Range(rp.Cells(R_YIL_SATIR, 3), rp.Cells(R_YIL_SATIR, 7))
        s.Format.Fill.ForeColor.RGB = C_LAC()
        Set s = .SeriesCollection.NewSeries
        s.Name = "Garanti dýþý"
        s.Values = rp.Range(rp.Cells(R_GARDIS, 3), rp.Cells(R_GARDIS, 7))
        s.XValues = rp.Range(rp.Cells(R_YIL_SATIR, 3), rp.Cells(R_YIL_SATIR, 7))
        s.Format.Fill.ForeColor.RGB = C_CAMGOBEGI()
        .ChartGroups(1).GapWidth = 70
    End With
    KS_GrafikBicim co.Chart

    Set co = ws.ChartObjects.Add(ws.Cells(P_GRAFIK + 1, 10).Left, ust, _
        ws.Cells(P_GRAFIK + 1, 17).Left - ws.Cells(P_GRAFIK + 1, 10).Left, yuk)
    co.Name = "ksGrafikAy"
    With co.Chart
        .ChartType = xlColumnClustered
        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop
        Set s = .SeriesCollection.NewSeries
        s.Name = "Arýza kaydý"
        s.Values = rp.Range(rp.Cells(R_AY_BASLIK + 1, 5), rp.Cells(R_AY_BASLIK + 12, 5))
        s.XValues = rp.Range(rp.Cells(R_AY_BASLIK + 1, 2), rp.Cells(R_AY_BASLIK + 12, 2))
        s.Format.Fill.ForeColor.RGB = C_LAC()
        Set s = .SeriesCollection.NewSeries
        s.Name = "Tamir maliyeti (TL)"
        s.Values = rp.Range(rp.Cells(R_AY_BASLIK + 1, 7), rp.Cells(R_AY_BASLIK + 12, 7))
        s.XValues = rp.Range(rp.Cells(R_AY_BASLIK + 1, 2), rp.Cells(R_AY_BASLIK + 12, 2))
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
    ch.ChartArea.Format.Line.ForeColor.RGB = C_CIZGI()
    ch.ChartArea.Font.Name = KS_FONT
    ch.ChartArea.Font.Size = 8
    ch.ChartArea.Font.Color = C_GRI()
    ch.PlotArea.Format.Fill.Visible = msoFalse
    ch.Axes(xlValue).MajorGridlines.Format.Line.ForeColor.RGB = RGB(232, 237, 243)
    ch.Axes(xlValue).Format.Line.Visible = msoFalse
    ch.Axes(xlValue).TickLabels.NumberFormat = "#,##0"
    ch.Axes(xlCategory).Format.Line.ForeColor.RGB = C_CIZGI()
    ch.Axes(xlCategory).TickLabels.Font.Color = C_METIN()
    On Error GoTo 0
End Sub

' ================================================================ yenileme

Public Sub KS_PANEL_YENILE()
    KS_PanelYenile True
End Sub

Public Sub KS_PanelYenile(ByVal mesaj As Boolean)
    Dim ws As Worksheet, aS As Variant, aA As Variant, aM As Variant, t As Variant, korumali As Boolean
    Dim yilS As String, y1 As Long, y2 As Long, i As Long, n As Long, x As Double, d As Date, d2 As Date
    Dim bekleyen As Long, geciken As Long, garBit As Long, acik As Long, ureticide As Long, uyari As Long
    Dim liste() As Variant, idx As Variant, dK As Object, k As String, j As Long, renk As Long
    On Error GoTo Hata
    Set ws = KS_WS(SH_PANEL)
    If ws Is Nothing Then Exit Sub
    KS_Basla
    KS_Durum "Panel yenileniyor..."
    KS_HesaplaTumu
    KS_YilListesiGuncelle
    yilS = KS_Metin(KS_FV("pYil"))
    If IsNumeric(yilS) And Len(yilS) = 4 Then
        y1 = CLng(yilS): y2 = y1
    Else
        y1 = 0: y2 = 9999
    End If
    KS_RaporHesapla IIf(y1 = 0, Year(Date), y1)
    aS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    aM = KS_Veri(KS_LOZ(T_MALZ))
    uyari = CLng(KS_ParamSayi("pGarUyari", 60))
    t = KS_Topla(aS, aA, y1, y2)

    ' Güncel durum sayýlarý
    If IsArray(aS) Then
        For i = 1 To UBound(aS, 1)
            Select Case KS_Metin(aS(i, S_TDURUM))
                Case "BEKLENÝYOR": bekleyen = bekleyen + 1
                Case "GECÝKÝYOR": bekleyen = bekleyen + 1: geciken = geciken + 1
            End Select
            If Left$(KS_Metin(aS(i, S_GDURUM)), 7) = "BÝTÝYOR" Then garBit = garBit + 1
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Metin(aA(i, A_DURUM)) <> "KAPANDI" Then acik = acik + 1
            If KS_Metin(aA(i, A_DURUM)) = "ÜRETÝCÝDE" Then ureticide = ureticide + 1
        Next i
    End If

    korumali = ws.ProtectContents
    KS_Ac ws

    KS_KartYaz KS_KartAlan(ws, P_KART1, 0), t(0), "#,##0", Format$(t(1), "#,##0") & " birim sevk edildi"
    KS_KartYaz KS_KartAlan(ws, P_KART1, 1), t(2), KS_FmtTLKisa(), IIf(y1 = 0, "tüm kayýtlar", y1 & " yýlý")
    If IsEmpty(t(3)) Then
        renk = C_LAC()
    ElseIf t(3) >= 0.9 Then
        renk = C_YESIL()
    ElseIf t(3) >= 0.75 Then
        renk = C_TURUNCU()
    Else
        renk = C_KIRMIZI()
    End If
    KS_KartYaz KS_KartAlan(ws, P_KART1, 2), t(3), "0%", "ort. gecikme " & Format$(t(4), "0.0") & " gün  |  " & _
        t(16) & " teslim", renk
    KS_KartYaz KS_KartAlan(ws, P_KART1, 3), bekleyen, "#,##0", geciken & " tanesi planlanan tarihi geçti", _
        IIf(geciken > 0, C_KIRMIZI(), C_LAC())
    KS_KartYaz KS_KartAlan(ws, P_KART1, 4), garBit, "#,##0", "önümüzdeki " & uyari & " gün içinde", _
        IIf(garBit > 0, C_TURUNCU(), C_LAC())
    KS_KartYaz KS_KartAlan(ws, P_KART2, 0), t(5), "#,##0", Format$(t(6), "#,##0") & " adet arýzalý  |  oran " & _
        Format$(IIf(IsEmpty(t(7)), 0, t(7)), "0.0%")
    KS_KartYaz KS_KartAlan(ws, P_KART2, 1), KS_Oran(t(8), t(8) + t(9)), "0%", t(8) & " kayýt garanti kapsamýnda"
    KS_KartYaz KS_KartAlan(ws, P_KART2, 2), acik, "#,##0", ureticide & " tanesi üreticide", _
        IIf(acik > 0, C_KIRMIZI(), C_YESIL())
    KS_KartYaz KS_KartAlan(ws, P_KART2, 3), t(10), "0.0"" gün""", "en uzun " & Format$(t(15), "0") & " gün"
    KS_KartYaz KS_KartAlan(ws, P_KART2, 4), t(12), KS_FmtTLKisa(), "kurum payý " & Format$(t(14), "#,##0") & _
        " " & ChrW(8378)

    ' Liste 1: en çok arýza veren malzemeler (döneme göre)
    Set dK = CreateObject("Scripting.Dictionary")
    n = 0
    ReDim liste(1 To KS_Satir(aA) + 1, 1 To 7)
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Tarih(aA(i, A_TARIH), d) Then
                If Year(d) >= y1 And Year(d) <= y2 Then
                    k = KS_Anahtar(aA(i, A_KOD))
                    If Not dK.Exists(k) Then
                        n = n + 1
                        dK(k) = n
                        liste(n, 1) = aA(i, A_KOD): liste(n, 2) = aA(i, A_AD)
                        liste(n, 5) = 0: liste(n, 6) = 0: liste(n, 7) = 0
                    End If
                    j = dK(k)
                    liste(j, 5) = liste(j, 5) + 1
                    If KS_Sayi(aA(i, A_ADET), x) Then liste(j, 6) = liste(j, 6) + x
                    If KS_Sayi(aA(i, A_MALIYET), x) Then liste(j, 7) = liste(j, 7) + x
                End If
            End If
        Next i
    End If
    KS_ListeYaz ws, P_LISTE1, 2, liste, n, 5, True, Array("@", "@", "#,##0", "#,##0", KS_FmtTLKisa()), _
        "Bu dönemde arýza kaydý yok."

    ' Liste 2: açýk / üreticide arýzalar (bekleme süresine göre)
    n = 0
    ReDim liste(1 To KS_Satir(aA) + 1, 1 To 7)
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Metin(aA(i, A_DURUM)) <> "KAPANDI" Then
                n = n + 1
                liste(n, 1) = aA(i, A_NO)
                liste(n, 2) = KS_Metin(aA(i, A_KOD)) & "  " & KS_Metin(aA(i, A_AD))
                liste(n, 5) = aA(i, A_DURUM)
                If KS_Tarih(aA(i, A_BILDIRIM), d) Then
                    liste(n, 6) = CLng(Date - d)
                ElseIf KS_Tarih(aA(i, A_TARIH), d) Then
                    liste(n, 6) = CLng(Date - d)
                End If
                liste(n, 7) = aA(i, A_YER)
            End If
        Next i
    End If
    KS_ListeYaz ws, P_LISTE1, 10, liste, n, 6, True, Array("@", "@", "@", "#,##0", "@"), _
        "Açýk arýza yok."

    ' Liste 3: garantisi yakýnda bitecekler
    n = 0
    ReDim liste(1 To KS_Satir(aS) + 1, 1 To 7)
    If IsArray(aS) Then
        For i = 1 To UBound(aS, 1)
            If KS_Tarih(aS(i, S_GARBIT), d) Then
                If d >= Date And d - Date <= uyari Then
                    n = n + 1
                    liste(n, 1) = aS(i, S_NO)
                    liste(n, 2) = KS_Metin(aS(i, S_KOD)) & "  " & KS_Metin(aS(i, S_AD))
                    liste(n, 5) = aS(i, S_YER)
                    liste(n, 6) = d
                    liste(n, 7) = CLng(d - Date)
                End If
            End If
        Next i
    End If
    KS_ListeYaz ws, P_LISTE2, 2, liste, n, 7, False, Array("@", "@", "@", KS_FmtTarih(), "#,##0"), _
        "Önümüzdeki " & uyari & " günde garantisi biten sevkiyat yok."

    ' Liste 4: geciken teslimatlar
    n = 0
    ReDim liste(1 To KS_Satir(aS) + 1, 1 To 7)
    If IsArray(aS) Then
        For i = 1 To UBound(aS, 1)
            If KS_Metin(aS(i, S_TDURUM)) = "GECÝKÝYOR" Then
                n = n + 1
                liste(n, 1) = aS(i, S_NO)
                liste(n, 2) = KS_Metin(aS(i, S_KOD)) & "  " & KS_Metin(aS(i, S_AD))
                liste(n, 5) = aS(i, S_SEVKCI)
                liste(n, 6) = aS(i, S_PLAN)
                liste(n, 7) = aS(i, S_GECIKME)
            End If
        Next i
    End If
    KS_ListeYaz ws, P_LISTE2, 10, liste, n, 7, True, Array("@", "@", "@", KS_FmtTarih(), "#,##0"), _
        "Geciken teslimat yok."

    ' Durum satýrý ve yedek hatýrlatmasý
    ws.Range("J7").Value = "Son güncelleme: " & Format$(Now, "dd.mm.yyyy hh:nn") & KS_YedekNotu()
    ws.Range("J7").Font.Color = IIf(InStr(ws.Range("J7").Value, "YEDEK") > 0, C_TURUNCU(), C_GRI())
    ws.Range("J7").Font.Bold = (InStr(ws.Range("J7").Value, "YEDEK") > 0)

    If korumali Then KS_Kapat ws
    KS_Bitir
    If mesaj Then KS_WS(SH_PANEL).Activate
    Exit Sub
Hata:
    KS_Bitir True
    If mesaj Then MsgBox "Panel yenilenemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Private Function KS_Oran(ByVal pay As Variant, ByVal payda As Variant) As Variant
    If payda <> 0 Then KS_Oran = pay / payda Else KS_Oran = Empty
End Function

' liste sütunlarý: 1 ilk sütun, 2 birleþik ad, 5..7 son üç sütun. bicim: 1, 2, 5, 6, 7 sütunlarýnýn biçimi.
Private Sub KS_ListeYaz(ByVal ws As Worksheet, ByVal ust As Long, ByVal kol As Long, ByRef liste() As Variant, _
        ByVal n As Long, ByVal siraKol As Long, ByVal azalan As Boolean, ByVal bicim As Variant, ByVal bos As String)
    Dim r As Long, i As Long, idx As Variant, satir As Long, c As Variant, j As Long
    c = Array(1, 2, 5, 6, 7)
    For r = ust + 2 To ust + 1 + P_SATIR
        ws.Cells(r, kol).Value = Empty
        ws.Cells(r, kol + 1).Value = Empty
        ws.Range(ws.Cells(r, kol + 4), ws.Cells(r, kol + 6)).Value = Empty
        ws.Cells(r, kol).Font.Italic = False
    Next r
    If n = 0 Then
        ws.Cells(ust + 2, kol + 1).Value = bos
        ws.Cells(ust + 2, kol + 1).Font.Italic = True
        ws.Cells(ust + 2, kol + 1).Font.Color = C_GRI()
        Exit Sub
    End If
    ws.Cells(ust + 2, kol + 1).Font.Italic = False
    ws.Range(ws.Cells(ust + 2, kol + 1), ws.Cells(ust + 1 + P_SATIR, kol + 1)).Font.Color = C_METIN()
    idx = KS_SiraIndeks(liste, siraKol, azalan, n)
    For i = 1 To n
        If i > P_SATIR Then Exit For
        satir = ust + 1 + i
        For j = 0 To 4
            With ws.Cells(satir, kol + c(j) - 1)
                .NumberFormat = bicim(j)
                .Value = liste(idx(i), c(j))
            End With
        Next j
    Next i
    If n > P_SATIR Then
        ws.Cells(ust, kol + 6).Value = "ilk " & P_SATIR & " / " & n
    Else
        ws.Cells(ust, kol + 6).Value = n & " kayýt"
    End If
    ws.Cells(ust, kol + 6).HorizontalAlignment = xlRight
    ws.Cells(ust, kol + 6).Font.Size = 8
    ws.Cells(ust, kol + 6).Font.Bold = False
    ws.Cells(ust, kol + 6).Font.Color = C_GRI()
End Sub
