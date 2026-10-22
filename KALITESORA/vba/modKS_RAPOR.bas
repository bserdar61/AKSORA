Attribute VB_Name = "modKS_RAPOR"
Option Explicit

' Raporlar: 5 yýllýk karþýlaþtýrma, malzeme karnesi, sevkçi / üretici performans karnesi.

Public Const R_YIL_SATIR As Long = 10       ' 5 yýllýk tablo baþlýk satýrý
Public Const R_GOSTERGE_ILK As Long = 11    ' ilk gösterge satýrý
Public Const R_AY_BASLIK As Long = 28       ' aylýk tablo baþlýk satýrý
Public Const R_GARIC As Long = 19           ' Garanti içi arýza satýrý (grafik)
Public Const R_GARDIS As Long = 20          ' Garanti dýþý arýza satýrý (grafik)

' ================================================================ ortak toplama

Public Function KS_SevkYil(ByRef aS As Variant, ByVal i As Long) As Long
    Dim d As Date
    If KS_Tarih(aS(i, S_GELIS), d) Then
        KS_SevkYil = Year(d)
    ElseIf KS_Tarih(aS(i, S_PLAN), d) Then
        KS_SevkYil = Year(d)
    End If
End Function

Public Function KS_SevkAy(ByRef aS As Variant, ByVal i As Long) As Long
    Dim d As Date
    If KS_Tarih(aS(i, S_GELIS), d) Then
        KS_SevkAy = Month(d)
    ElseIf KS_Tarih(aS(i, S_PLAN), d) Then
        KS_SevkAy = Month(d)
    End If
End Function

Private Function KS_Bol(ByVal pay As Double, ByVal payda As Double) As Variant
    If payda <> 0 Then KS_Bol = pay / payda Else KS_Bol = Empty
End Function

' y1..y2 yýllarý arasýndaki göstergeler (0..16). Sevkiyat yýlý: geliþ, yoksa planlanan tarih.
Public Function KS_Topla(ByRef aS As Variant, ByRef aA As Variant, ByVal y1 As Long, ByVal y2 As Long) As Variant
    Dim r(0 To 16) As Variant, i As Long, y As Long, x As Double, d As Date, k As String
    Dim cnt As Long, mik As Double, tut As Double, zam As Long, tes As Long, gecTop As Double
    Dim aC As Long, aM As Double, gi As Long, gd As Long, donTop As Double, donN As Long, donMax As Double
    Dim cozTop As Double, cozN As Long, mal As Double, mU As Double, mK As Double, ku As String, kk As String
    ku = KS_Anahtar("Üretici")
    kk = KS_Anahtar("Kurum")
    If IsArray(aS) Then
        For i = 1 To UBound(aS, 1)
            y = KS_SevkYil(aS, i)
            If y >= y1 And y <= y2 Then
                cnt = cnt + 1
                If KS_Sayi(aS(i, S_ADET), x) Then mik = mik + x
                If KS_Sayi(aS(i, S_TUTAR), x) Then tut = tut + x
                Select Case KS_Metin(aS(i, S_TDURUM))
                    Case "ZAMANINDA"
                        zam = zam + 1
                        tes = tes + 1
                    Case "GECÝKMELÝ"
                        tes = tes + 1
                        If KS_Sayi(aS(i, S_GECIKME), x) Then gecTop = gecTop + x
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

' ================================================================ 5 YILLIK RAPOR

Private Function KS_Gostergeler() As Variant
    ' etiket, tür (n = adet, tl, oran, gun)
    KS_Gostergeler = Array( _
        Array("Sevk Kalemi", "n"), Array("Sevk Edilen Miktar", "n"), Array("Sevk Tutarý", "tl"), _
        Array("Zamanýnda Teslim Oraný", "oran"), Array("Ortalama Gecikme (gün)", "gun"), _
        Array("Arýza Kaydý", "n"), Array("Arýzalý Miktar", "n"), Array("Arýza Oraný (arýzalý / sevk)", "oran"), _
        Array("Garanti Ýçi Arýza", "n"), Array("Garanti Dýþý Arýza", "n"), _
        Array("Ort. Üretici Dönüþ Süresi (gün)", "gun"), Array("Ort. Çözüm Süresi (gün)", "gun"), _
        Array("Tamir Maliyeti", "tl"), Array("Üreticinin Karþýladýðý", "tl"), Array("Kurumun Karþýladýðý", "tl"))
End Function

Private Function KS_TurBicim(ByVal tur As String) As String
    Select Case tur
        Case "tl": KS_TurBicim = KS_FmtTLKisa()
        Case "oran": KS_TurBicim = "0.0%"
        Case "gun": KS_TurBicim = "0.0"
        Case Else: KS_TurBicim = "#,##0"
    End Select
End Function

Public Sub KS_RaporKur()
    Dim ws As Worksheet, g As Variant, i As Long
    Set ws = KS_WS(SH_RAPOR)
    KS_SayfaSifirla ws
    ws.Columns("B").ColumnWidth = 34
    ws.Columns("C:G").ColumnWidth = 15
    ws.Columns("H").ColumnWidth = 16
    ws.Columns("I").ColumnWidth = 14
    ws.Columns("J").ColumnWidth = 3
    KS_Bant ws, 10, "5 YILLIK KALÝTE RAPORU", "Seçilen yýl ve önceki 4 yýl. Sevkiyat yýlý: geliþ tarihi, yoksa planlanan tarih."

    ws.Range("B7").Value = "RAPOR YILI (son yýl)"
    ws.Range("B7").Font.Bold = True
    ws.Range("B7").Font.Size = 9
    ws.Range("B7").Font.Color = C_GRI()
    With ws.Range("C7")
        .Value = Year(Date)
        .Interior.Color = C_GIRIS()
        .Font.Bold = True
        .Font.Color = C_LAC()
        .HorizontalAlignment = xlCenter
        .Locked = False
        .Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
    End With
    KS_AdHucre "rYil", ws.Range("C7")
    KS_Liste ws.Range("C7"), "lstYilSayi", True
    ws.Rows(7).RowHeight = 24
    KS_Buton ws, "ksBtnRapor", "HESAPLA", ws.Range("E7").Left, ws.Range("E7").Top + 1, 110, 22, "KS_RAPOR_HESAPLA", "ana"

    KS_Bolum ws.Range("B9:I9"), "5 YILLIK KARÞILAÞTIRMA"
    KS_TabloBaslik ws.Range("B10:I10"), Array("GÖSTERGE", "", "", "", "", "", "5 YIL", "DEÐÝÞÝM (son yýl)")
    g = KS_Gostergeler()
    For i = 0 To UBound(g)
        ws.Cells(R_GOSTERGE_ILK + i, 2).Value = g(i)(0)
        ws.Range(ws.Cells(R_GOSTERGE_ILK + i, 3), ws.Cells(R_GOSTERGE_ILK + i, 8)).NumberFormat = KS_TurBicim(CStr(g(i)(1)))
    Next i
    With ws.Range(ws.Cells(R_GOSTERGE_ILK, 2), ws.Cells(R_GOSTERGE_ILK + UBound(g), 9))
        KS_TabloGovde .Cells
        .RowHeight = 20
    End With
    ws.Range(ws.Cells(R_GOSTERGE_ILK, 2), ws.Cells(R_GOSTERGE_ILK + UBound(g), 2)).Font.Bold = True
    ws.Range(ws.Cells(R_GOSTERGE_ILK, 2), ws.Cells(R_GOSTERGE_ILK + UBound(g), 2)).IndentLevel = 1
    ws.Range(ws.Cells(R_GOSTERGE_ILK, 3), ws.Cells(R_GOSTERGE_ILK + UBound(g), 9)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(R_GOSTERGE_ILK, 8), ws.Cells(R_GOSTERGE_ILK + UBound(g), 8)).Interior.Color = C_OTO()
    ws.Range(ws.Cells(R_GOSTERGE_ILK, 8), ws.Cells(R_GOSTERGE_ILK + UBound(g), 8)).Font.Bold = True

    KS_Bolum ws.Range("B27:I27"), "AYLIK DAÐILIM"
    KS_TabloBaslik ws.Range("B28:H28"), Array("AY", "Sevk Kalemi", "Sevk Tutarý", "Arýza Kaydý", _
        "Garanti Ýçi Arýza", "Tamir Maliyeti", "Ort. Dönüþ (gün)")
    For i = 1 To 12
        ws.Cells(R_AY_BASLIK + i, 2).Value = KS_AyAdi(i)
    Next i
    ws.Cells(R_AY_BASLIK + 13, 2).Value = "TOPLAM"
    KS_TabloGovde ws.Range(ws.Cells(R_AY_BASLIK + 1, 2), ws.Cells(R_AY_BASLIK + 13, 8))
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 2), ws.Cells(R_AY_BASLIK + 13, 2)).IndentLevel = 1
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 3), ws.Cells(R_AY_BASLIK + 13, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 4), ws.Cells(R_AY_BASLIK + 13, 4)).NumberFormat = KS_FmtTLKisa()
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 5), ws.Cells(R_AY_BASLIK + 13, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 7), ws.Cells(R_AY_BASLIK + 13, 7)).NumberFormat = KS_FmtTLKisa()
    ws.Range(ws.Cells(R_AY_BASLIK + 1, 8), ws.Cells(R_AY_BASLIK + 13, 8)).NumberFormat = "0.0"
    With ws.Range(ws.Cells(R_AY_BASLIK + 13, 2), ws.Cells(R_AY_BASLIK + 13, 8))
        .Font.Bold = True
        .Interior.Color = C_OTO()
        .Borders(xlEdgeTop).Color = C_LAC()
    End With
    ws.Range("B44").Value = "Deðiþim sütunu: adet ve tutarda önceki yýla göre % deðiþim; oranlarda puan, sürelerde gün farký."
    ws.Range("B44").Font.Size = 8
    ws.Range("B44").Font.Italic = True
    ws.Range("B44").Font.Color = C_GRI()
    KS_Pencere ws, 90, 0, False
End Sub

Public Function KS_AyAdi(ByVal ay As Long) As String
    KS_AyAdi = Choose(ay, "Ocak", "Þubat", "Mart", "Nisan", "Mayýs", "Haziran", "Temmuz", "Aðustos", _
        "Eylül", "Ekim", "Kasým", "Aralýk")
End Function

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
    Dim ws As Worksheet, aS As Variant, aA As Variant, g As Variant, t As Variant, sut(1 To 6) As Variant
    Dim i As Long, j As Long, y As Long, ay As Long, d As Date, x As Double, korumali As Boolean
    Dim mS(1 To 13) As Double, mT(1 To 13) As Double, mA(1 To 13) As Double, mG(1 To 13) As Double
    Dim mM(1 To 13) As Double, mD(1 To 13) As Double, mDn(1 To 13) As Long, v As Variant, onc As Variant

    Set ws = KS_WS(SH_RAPOR)
    korumali = ws.ProtectContents
    KS_Ac ws
    aS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    g = KS_Gostergeler()
    ws.Range("C7").Value = yil
    For j = 1 To 5
        ws.Cells(R_YIL_SATIR, 2 + j).Value = CStr(yil - 5 + j)
        sut(j) = KS_Topla(aS, aA, yil - 5 + j, yil - 5 + j)
    Next j
    sut(6) = KS_Topla(aS, aA, yil - 4, yil)
    ws.Cells(R_YIL_SATIR, 8).Value = "5 YIL"
    For i = 0 To UBound(g)
        For j = 1 To 6
            ws.Cells(R_GOSTERGE_ILK + i, 2 + j).Value = sut(j)(i)
        Next j
        ' Deðiþim
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

    ' Aylýk daðýlým (seçili yýl)
    If IsArray(aS) Then
        For i = 1 To UBound(aS, 1)
            If KS_SevkYil(aS, i) = yil Then
                ay = KS_SevkAy(aS, i)
                mS(ay) = mS(ay) + 1
                If KS_Sayi(aS(i, S_TUTAR), x) Then mT(ay) = mT(ay) + x
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
    ws.Range("B27").Value = "AYLIK DAÐILIM  -  " & yil
    If korumali Then KS_Kapat ws
End Sub

' ================================================================ MALZEME KARNESÝ

Private Const K_ILK As Long = 9

Public Sub KS_KarneKur()
    Dim ws As Worksheet
    Set ws = KS_WS(SH_KARNE)
    KS_SayfaSifirla ws
    ws.Columns("B:N").ColumnWidth = 12.5
    ws.Columns("O").ColumnWidth = 3
    KS_Bant ws, 15, "MALZEME KARNESÝ", "Bir malzemenin tüm sevkiyat ve arýza geçmiþi tek sayfada."
    With ws.Range("B7")
        .Value = "MALZEME KODU"
        .Font.Bold = True
        .Font.Size = 9
        .Font.Color = C_GRI()
    End With
    ws.Range("C7:D7").Merge
    With ws.Range("C7:D7")
        .Interior.Color = C_GIRIS()
        .Font.Bold = True
        .Font.Color = C_LAC()
        .Locked = False
        .NumberFormat = "@"
        .Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
        .HorizontalAlignment = xlCenter
    End With
    ws.Rows(7).RowHeight = 24
    KS_AdHucre "kKod", ws.Range("C7")
    KS_Liste ws.Range("C7:D7"), "lstMalzeme", True
    KS_Buton ws, "ksBtnKarne", "GÖSTER", ws.Range("F7").Left, ws.Range("F7").Top + 1, 110, 22, "KS_KARNE_GOSTER", "ana"
    ws.Range("B" & K_ILK).Value = "Malzeme kodunu seçip GÖSTER'e basýn."
    ws.Range("B" & K_ILK).Font.Color = C_GRI()
    ws.Range("B" & K_ILK).Font.Italic = True
    KS_Pencere ws, 90, 0, False
End Sub

Public Sub KS_KARNE_GOSTER()
    Dim ws As Worksheet, aM As Variant, aS As Variant, aA As Variant, kod As String, mi As Long
    Dim i As Long, n As Long, r As Long, idx As Variant, liste() As Variant, x As Double, korumali As Boolean
    Dim sevkMik As Double, sevkTut As Double, arzN As Long, arzMik As Double, donTop As Double, donN As Long
    Dim mal As Double, garIc As Long, ilk As Date, son As Date, d As Date, k As Long
    On Error GoTo Hata
    Set ws = KS_WS(SH_KARNE)
    kod = KS_Anahtar(KS_FV("kKod"))
    If Len(kod) = 0 Then
        MsgBox "Malzeme kodu seçin.", vbExclamation, KS_AD
        Exit Sub
    End If
    KS_Basla
    KS_HesaplaTumu
    aM = KS_Veri(KS_LOZ(T_MALZ))
    aS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    If IsArray(aM) Then
        For i = 1 To UBound(aM, 1)
            If KS_Anahtar(aM(i, M_KOD)) = kod Then mi = i: Exit For
        Next i
    End If
    If mi = 0 Then
        KS_Bitir
        MsgBox "Malzeme bulunamadý: " & KS_FV("kKod"), vbExclamation, KS_AD
        Exit Sub
    End If

    korumali = ws.ProtectContents
    KS_Ac ws
    KS_AlanTemizle ws, K_ILK

    ' Baþlýk ve bilgi satýrý
    ws.Range(ws.Cells(K_ILK, 2), ws.Cells(K_ILK, 14)).Merge
    With ws.Cells(K_ILK, 2)
        .Value = KS_Metin(aM(mi, M_KOD)) & "   " & KS_Metin(aM(mi, M_AD))
        .Font.Size = 16
        .Font.Bold = True
        .Font.Color = C_LAC()
        .Font.Italic = False
    End With
    ws.Rows(K_ILK).RowHeight = 30
    ws.Range(ws.Cells(K_ILK + 1, 2), ws.Cells(K_ILK + 1, 14)).Merge
    With ws.Cells(K_ILK + 1, 2)
        .Value = "Kategori: " & KS_Bilgi(aM(mi, M_KAT)) & "     |     Birim: " & KS_Bilgi(aM(mi, M_BIRIM)) & _
            "     |     Üretici: " & KS_Bilgi(aM(mi, M_URETICI)) & "     |     Liste fiyatý: " & _
            Format$(aM(mi, M_FIYAT), "#,##0.00") & " " & ChrW(8378) & "     |     Garanti: " & aM(mi, M_GARANTI) & " ay"
        .Font.Size = 9
        .Font.Color = C_GRI()
    End With
    ws.Range(ws.Cells(K_ILK + 1, 2), ws.Cells(K_ILK + 1, 14)).Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
    ws.Range(ws.Cells(K_ILK + 1, 2), ws.Cells(K_ILK + 1, 14)).Borders(xlEdgeBottom).Weight = xlMedium

    ' Sevkiyat geçmiþi
    r = K_ILK + 7
    KS_Bolum ws.Range(ws.Cells(r, 2), ws.Cells(r, 14)), "SEVKÝYAT GEÇMÝÞÝ"
    KS_TabloBaslik ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 13)), Array("Sevk No", "Sevkçi", "Sevk Yeri", _
        "Adet", "Birim Fiyat", "Tutar", "Planlanan", "Geliþ", "Gecikme (gün)", "Garanti Bitiþ", "Garanti Durumu", _
        "Arýzalý Adet")
    n = 0
    If IsArray(aS) Then
        ReDim liste(1 To UBound(aS, 1), 1 To 13)
        For i = 1 To UBound(aS, 1)
            If KS_Anahtar(aS(i, S_KOD)) = kod Then
                n = n + 1
                liste(n, 1) = aS(i, S_NO): liste(n, 2) = aS(i, S_SEVKCI): liste(n, 3) = aS(i, S_YER)
                liste(n, 4) = aS(i, S_ADET): liste(n, 5) = aS(i, S_FIYAT): liste(n, 6) = aS(i, S_TUTAR)
                liste(n, 7) = aS(i, S_PLAN): liste(n, 8) = aS(i, S_GELIS): liste(n, 9) = aS(i, S_GECIKME)
                liste(n, 10) = aS(i, S_GARBIT): liste(n, 11) = aS(i, S_GDURUM): liste(n, 12) = aS(i, S_ARIZA)
                If KS_Tarih(aS(i, S_GELIS), d) Then liste(n, 13) = d Else liste(n, 13) = aS(i, S_PLAN)
                If KS_Sayi(aS(i, S_ADET), x) Then sevkMik = sevkMik + x
                If KS_Sayi(aS(i, S_TUTAR), x) Then sevkTut = sevkTut + x
            End If
        Next i
    End If
    r = r + 2
    If n = 0 Then
        ws.Cells(r, 2).Value = "Bu malzeme için sevkiyat kaydý yok."
        ws.Cells(r, 2).Font.Color = C_GRI()
        r = r + 1
    Else
        idx = KS_SiraIndeks(liste, 13, True, n)
        For i = 1 To n
            For k = 1 To 12
                ws.Cells(r + i - 1, 1 + k).Value = liste(idx(i), k)
            Next k
        Next i
        KS_TabloGovde ws.Range(ws.Cells(r, 2), ws.Cells(r + n - 1, 13))
        ws.Range(ws.Cells(r, 5), ws.Cells(r + n - 1, 5)).NumberFormat = "General"
        ws.Range(ws.Cells(r, 6), ws.Cells(r + n - 1, 7)).NumberFormat = KS_FmtTL()
        ws.Range(ws.Cells(r, 8), ws.Cells(r + n - 1, 9)).NumberFormat = KS_FmtTarih()
        ws.Range(ws.Cells(r, 11), ws.Cells(r + n - 1, 11)).NumberFormat = KS_FmtTarih()
        ws.Range(ws.Cells(r, 3), ws.Cells(r + n - 1, 4)).ShrinkToFit = True
        ws.Range(ws.Cells(r, 12), ws.Cells(r + n - 1, 12)).ShrinkToFit = True
        r = r + n
    End If

    ' Arýza geçmiþi (kronolojik)
    r = r + 1
    KS_Bolum ws.Range(ws.Cells(r, 2), ws.Cells(r, 14)), "ARIZA GEÇMÝÞÝ  (eskiden yeniye)"
    KS_TabloBaslik ws.Range(ws.Cells(r + 1, 2), ws.Cells(r + 1, 14)), Array("Arýza No", "Sevk No", "Seri No", _
        "Arýza Tarihi", "Önceki Arýza", "Adet", "Arýza Tipi", "Garanti", "Bildirim", "Üretici Dönüþ", _
        "Dönüþ (gün)", "Tamir Maliyeti", "Durum")
    n = 0
    If IsArray(aA) Then
        ReDim liste(1 To UBound(aA, 1), 1 To 13)
        For i = 1 To UBound(aA, 1)
            If KS_Anahtar(aA(i, A_KOD)) = kod Then
                n = n + 1
                liste(n, 1) = aA(i, A_NO): liste(n, 2) = aA(i, A_SEVK): liste(n, 3) = aA(i, A_SERI)
                liste(n, 4) = aA(i, A_TARIH): liste(n, 5) = aA(i, A_ONCEKI): liste(n, 6) = aA(i, A_ADET)
                liste(n, 7) = aA(i, A_TIP): liste(n, 8) = aA(i, A_KAPSAM): liste(n, 9) = aA(i, A_BILDIRIM)
                liste(n, 10) = aA(i, A_DONUS): liste(n, 11) = aA(i, A_DONUSGUN): liste(n, 12) = aA(i, A_MALIYET)
                liste(n, 13) = aA(i, A_DURUM)
                arzN = arzN + 1
                If KS_Sayi(aA(i, A_ADET), x) Then arzMik = arzMik + x
                If KS_Sayi(aA(i, A_DONUSGUN), x) Then donTop = donTop + x: donN = donN + 1
                If KS_Sayi(aA(i, A_MALIYET), x) Then mal = mal + x
                If KS_Metin(aA(i, A_KAPSAM)) = "EVET" Then garIc = garIc + 1
            End If
        Next i
    End If
    r = r + 2
    If n = 0 Then
        ws.Cells(r, 2).Value = "Bu malzeme için arýza kaydý yok."
        ws.Cells(r, 2).Font.Color = C_YESIL()
        ws.Cells(r, 2).Font.Bold = True
    Else
        idx = KS_SiraIndeks(liste, 4, False, n)
        For i = 1 To n
            For k = 1 To 13
                ws.Cells(r + i - 1, 1 + k).Value = liste(idx(i), k)
            Next k
        Next i
        KS_TabloGovde ws.Range(ws.Cells(r, 2), ws.Cells(r + n - 1, 14))
        ws.Range(ws.Cells(r, 5), ws.Cells(r + n - 1, 6)).NumberFormat = KS_FmtTarih()
        ws.Range(ws.Cells(r, 10), ws.Cells(r + n - 1, 11)).NumberFormat = KS_FmtTarih()
        ws.Range(ws.Cells(r, 13), ws.Cells(r + n - 1, 13)).NumberFormat = KS_FmtTL()
        KS_DurumYazRenk ws.Range(ws.Cells(r, 9), ws.Cells(r + n - 1, 9)), "EVET", "HAYIR"
        KS_DurumYazRenk ws.Range(ws.Cells(r, 14), ws.Cells(r + n - 1, 14)), "KAPANDI", "AÇIK"
    End If

    ' KPI kartlarý
    KS_KartCiz ws.Range(ws.Cells(K_ILK + 3, 2), ws.Cells(K_ILK + 5, 3)), "SEVK EDÝLEN MÝKTAR"
    KS_KartYaz ws.Range(ws.Cells(K_ILK + 3, 2), ws.Cells(K_ILK + 5, 3)), sevkMik, "#,##0", _
        KS_Bilgi(aM(mi, M_BIRIM))
    KS_KartCiz ws.Range(ws.Cells(K_ILK + 3, 4), ws.Cells(K_ILK + 5, 5)), "SEVK TUTARI"
    KS_KartYaz ws.Range(ws.Cells(K_ILK + 3, 4), ws.Cells(K_ILK + 5, 5)), sevkTut, KS_FmtTLKisa(), "tüm sevkiyatlar"
    KS_KartCiz ws.Range(ws.Cells(K_ILK + 3, 6), ws.Cells(K_ILK + 5, 7)), "ARIZA KAYDI"
    KS_KartYaz ws.Range(ws.Cells(K_ILK + 3, 6), ws.Cells(K_ILK + 5, 7)), arzN, "#,##0", _
        garIc & " garanti içi  |  " & arzMik & " adet", IIf(arzN > 0, C_KIRMIZI(), C_YESIL())
    KS_KartCiz ws.Range(ws.Cells(K_ILK + 3, 8), ws.Cells(K_ILK + 5, 9)), "ARIZA ORANI"
    KS_KartYaz ws.Range(ws.Cells(K_ILK + 3, 8), ws.Cells(K_ILK + 5, 9)), KS_Bol(arzMik, sevkMik), "0.0%", _
        "arýzalý / sevk edilen"
    KS_KartCiz ws.Range(ws.Cells(K_ILK + 3, 10), ws.Cells(K_ILK + 5, 11)), "ORT. ÜRETÝCÝ DÖNÜÞÜ"
    KS_KartYaz ws.Range(ws.Cells(K_ILK + 3, 10), ws.Cells(K_ILK + 5, 11)), KS_Bol(donTop, donN), "0.0"" gün""", _
        donN & " dönüþ kaydý"
    KS_KartCiz ws.Range(ws.Cells(K_ILK + 3, 12), ws.Cells(K_ILK + 5, 13)), "TAMÝR MALÝYETÝ"
    KS_KartYaz ws.Range(ws.Cells(K_ILK + 3, 12), ws.Cells(K_ILK + 5, 13)), mal, KS_FmtTLKisa(), "toplam"
    ws.Rows(K_ILK + 2).RowHeight = 8
    ws.Rows(K_ILK + 3).RowHeight = 18
    ws.Rows(K_ILK + 4).RowHeight = 30
    ws.Rows(K_ILK + 5).RowHeight = 16

    If korumali Then KS_Kapat ws
    KS_Bitir
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Karne oluþturulamadý: " & Err.Description, vbExclamation, KS_AD
End Sub

Private Function KS_Bilgi(ByVal v As Variant) As String
    KS_Bilgi = KS_Metin(v)
    If Len(KS_Bilgi) = 0 Then KS_Bilgi = "-"
End Function

' Rapor alanýný (satýr ilk..son) biçim dahil temizler.
Public Sub KS_AlanTemizle(ByVal ws As Worksheet, ByVal ilk As Long)
    Dim r As Range
    Set r = ws.Range(ws.Rows(ilk), ws.Rows(ilk + 3000))
    r.FormatConditions.Delete
    r.UnMerge
    r.Clear
    r.Interior.Color = C_ZEMIN()
    r.Font.Name = KS_FONT
    r.Font.Size = 10
    r.Font.Color = C_METIN()
    r.RowHeight = 18
    r.VerticalAlignment = xlCenter
End Sub

Public Sub KS_DurumYazRenk(ByVal r As Range, ByVal iyi As String, ByVal kotu As String)
    Dim fc As Object
    Set fc = r.FormatConditions.Add(Type:=xlTextString, String:=iyi, TextOperator:=xlBeginsWith)
    fc.Font.Color = C_YESIL(): fc.Font.Bold = True
    Set fc = r.FormatConditions.Add(Type:=xlTextString, String:=kotu, TextOperator:=xlBeginsWith)
    fc.Font.Color = C_KIRMIZI(): fc.Font.Bold = True
End Sub

' ================================================================ PERFORMANS

Private Const P_BASLIK As Long = 10

Public Sub KS_PerfKur()
    Dim ws As Worksheet
    Set ws = KS_WS(SH_PERF)
    KS_SayfaSifirla ws
    ws.Columns("B").ColumnWidth = 5
    ws.Columns("C").ColumnWidth = 42
    ws.Columns("D:R").ColumnWidth = 11
    ws.Columns("S").ColumnWidth = 3
    KS_Bant ws, 19, "TEDARÝKÇÝ / ÜRETÝCÝ PERFORMANSI", _
        "Teslim, arýza ve dönüþ verilerinden hesaplanan kalite skoru. Katsayýlar TANIMLAR sayfasýnda."
    ws.Range("C7").Value = "GRUPLAMA"
    ws.Range("F7").Value = "YIL"
    ws.Range("C7,F7").Font.Bold = True
    ws.Range("C7,F7").Font.Size = 9
    ws.Range("C7,F7").Font.Color = C_GRI()
    ws.Range("C7,F7").HorizontalAlignment = xlRight
    With ws.Range("D7:E7")
        .Merge
        .Value = "SEVKÇÝ"
    End With
    With ws.Range("G7")
        .Value = "TÜMÜ"
    End With
    With ws.Range("D7:E7,G7")
        .Interior.Color = C_GIRIS()
        .Font.Bold = True
        .Font.Color = C_LAC()
        .Locked = False
        .HorizontalAlignment = xlCenter
        .Borders(xlEdgeBottom).Color = C_CAMGOBEGI()
    End With
    KS_AdHucre "pfGrup", ws.Range("D7")
    KS_AdHucre "pfYil", ws.Range("G7")
    KS_Liste ws.Range("D7:E7"), "lstGrup", True
    KS_Liste ws.Range("G7"), "lstYil", True
    ws.Rows(7).RowHeight = 24
    KS_Buton ws, "ksBtnPerf", "HESAPLA", ws.Range("I7").Left, ws.Range("I7").Top + 1, 110, 22, "KS_PERF_HESAPLA", "ana"
    KS_Bolum ws.Range("B9:R9"), "KALÝTE KARNESÝ"
    KS_TabloBaslik ws.Range(ws.Cells(P_BASLIK, 2), ws.Cells(P_BASLIK, 18)), KS_PerfBasliklar("SEVKÇÝ")
    KS_Pencere ws, 85, 0, False
End Sub

Private Function KS_PerfBasliklar(ByVal grup As String) As Variant
    KS_PerfBasliklar = Array("#", IIf(grup = "ÜRETÝCÝ", "Üretici", "Sevkçi"), "Sevk Kalemi", "Sevk Miktarý", _
        "Sevk Tutarý", "Zamanýnda Teslim", "Ort. Gecikme (gün)", "Arýza Kaydý", "Arýzalý Miktar", "Arýza Oraný", _
        "Garanti Ýçi Arýza", "Ort. Dönüþ (gün)", "En Uzun Dönüþ", "Açýk Arýza", "Tamir Maliyeti", "Kalite Skoru", "Not")
End Function

Public Sub KS_PERF_HESAPLA()
    Dim ws As Worksheet, aM As Variant, aS As Variant, aA As Variant, dM As Object, dG As Object
    Dim grup As String, yilS As String, y1 As Long, y2 As Long, i As Long, j As Long, n As Long, g As Long
    Dim k As String, x As Double, d As Date, t() As Variant, idx As Variant, r As Long, korumali As Boolean
    Dim k1 As Double, k2 As Double, k3 As Double, skor As Double, oran As Double, gec As Double, don As Double
    ' t sütunlarý: 1 ad, 2 kalem, 3 miktar, 4 tutar, 5 zamanýnda, 6 teslim, 7 gecikme top, 8 arýza, 9 arýzalý mik,
    ' 10 garanti içi, 11 dönüþ top, 12 dönüþ n, 13 en uzun, 14 açýk, 15 maliyet, 16 skor
    On Error GoTo Hata
    Set ws = KS_WS(SH_PERF)
    grup = KS_Metin(KS_FV("pfGrup"))
    If grup <> "ÜRETÝCÝ" Then grup = "SEVKÇÝ"
    yilS = KS_Metin(KS_FV("pfYil"))
    If IsNumeric(yilS) And Len(yilS) = 4 Then
        y1 = CLng(yilS): y2 = y1
    Else
        y1 = 0: y2 = 9999
    End If
    k1 = KS_ParamSayi("pKatAriza", 5)
    k2 = KS_ParamSayi("pKatGecikme", 1)
    k3 = KS_ParamSayi("pKatDonus", 0.5)

    KS_Basla
    KS_HesaplaTumu
    aM = KS_Veri(KS_LOZ(T_MALZ))
    aS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    Set dM = CreateObject("Scripting.Dictionary")
    If IsArray(aM) Then
        For i = 1 To UBound(aM, 1)
            dM(KS_Anahtar(aM(i, M_KOD))) = KS_Metin(aM(i, M_URETICI))
        Next i
    End If
    Set dG = CreateObject("Scripting.Dictionary")
    ReDim t(1 To KS_Satir(aS) + KS_Satir(aA) + 1, 1 To 16)

    If IsArray(aS) Then
        For i = 1 To UBound(aS, 1)
            j = KS_SevkYil(aS, i)
            If j >= y1 And j <= y2 Then
                If grup = "SEVKÇÝ" Then
                    k = KS_Metin(aS(i, S_SEVKCI))
                ElseIf dM.Exists(KS_Anahtar(aS(i, S_KOD))) Then
                    k = dM(KS_Anahtar(aS(i, S_KOD)))
                Else
                    k = ""
                End If
                g = KS_GrupNo(dG, t, n, k)
                t(g, 2) = t(g, 2) + 1
                If KS_Sayi(aS(i, S_ADET), x) Then t(g, 3) = t(g, 3) + x
                If KS_Sayi(aS(i, S_TUTAR), x) Then t(g, 4) = t(g, 4) + x
                Select Case KS_Metin(aS(i, S_TDURUM))
                    Case "ZAMANINDA"
                        t(g, 5) = t(g, 5) + 1
                        t(g, 6) = t(g, 6) + 1
                    Case "GECÝKMELÝ"
                        t(g, 6) = t(g, 6) + 1
                        If KS_Sayi(aS(i, S_GECIKME), x) Then t(g, 7) = t(g, 7) + x
                End Select
            End If
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Tarih(aA(i, A_TARIH), d) Then
                If Year(d) >= y1 And Year(d) <= y2 Then
                    If grup = "SEVKÇÝ" Then k = KS_Metin(aA(i, A_SEVKCI)) Else k = KS_Metin(aA(i, A_URETICI))
                    g = KS_GrupNo(dG, t, n, k)
                    t(g, 8) = t(g, 8) + 1
                    If KS_Sayi(aA(i, A_ADET), x) Then t(g, 9) = t(g, 9) + x
                    If KS_Metin(aA(i, A_KAPSAM)) = "EVET" Then t(g, 10) = t(g, 10) + 1
                    If KS_Sayi(aA(i, A_DONUSGUN), x) Then
                        t(g, 11) = t(g, 11) + x
                        t(g, 12) = t(g, 12) + 1
                        If x > t(g, 13) Then t(g, 13) = x
                    End If
                    If KS_Metin(aA(i, A_DURUM)) <> "KAPANDI" Then t(g, 14) = t(g, 14) + 1
                    If KS_Sayi(aA(i, A_MALIYET), x) Then t(g, 15) = t(g, 15) + x
                End If
            End If
        Next i
    End If

    ' Skor = 100 - min(50, arýza% x k1) - min(25, ort. gecikme x k2) - min(25, ort. dönüþ x k3)
    For g = 1 To n
        oran = 0: gec = 0: don = 0
        If t(g, 3) > 0 Then oran = t(g, 9) / t(g, 3) * 100
        If t(g, 6) > 0 Then gec = t(g, 7) / t(g, 6)
        If t(g, 12) > 0 Then don = t(g, 11) / t(g, 12)
        skor = 100 - KS_Min(50, oran * k1) - KS_Min(25, gec * k2) - KS_Min(25, don * k3)
        If skor < 0 Then skor = 0
        t(g, 16) = Round(skor, 1)
    Next g

    korumali = ws.ProtectContents
    KS_Ac ws
    KS_AlanTemizle ws, P_BASLIK + 1
    KS_TabloBaslik ws.Range(ws.Cells(P_BASLIK, 2), ws.Cells(P_BASLIK, 18)), KS_PerfBasliklar(grup)
    ws.Range("B9").Value = "KALÝTE KARNESÝ  -  " & grup & "  -  " & IIf(y1 = 0, "TÜM YILLAR", CStr(y1))
    r = P_BASLIK + 1
    If n = 0 Then
        ws.Cells(r, 3).Value = "Seçilen dönemde kayýt yok."
        ws.Cells(r, 3).Font.Color = C_GRI()
    Else
        idx = KS_SiraIndeks(t, 16, True, n)
        For i = 1 To n
            g = idx(i)
            ws.Cells(r, 2).Value = i
            ws.Cells(r, 3).Value = t(g, 1)
            ws.Cells(r, 4).Value = t(g, 2)
            ws.Cells(r, 5).Value = t(g, 3)
            ws.Cells(r, 6).Value = t(g, 4)
            ws.Cells(r, 7).Value = KS_Bol(t(g, 5), t(g, 6))
            ws.Cells(r, 8).Value = KS_Bol(t(g, 7), t(g, 6))
            ws.Cells(r, 9).Value = t(g, 8)
            ws.Cells(r, 10).Value = t(g, 9)
            ws.Cells(r, 11).Value = KS_Bol(t(g, 9), t(g, 3))
            ws.Cells(r, 12).Value = t(g, 10)
            ws.Cells(r, 13).Value = KS_Bol(t(g, 11), t(g, 12))
            If t(g, 12) > 0 Then ws.Cells(r, 14).Value = t(g, 13)
            ws.Cells(r, 15).Value = t(g, 14)
            ws.Cells(r, 16).Value = t(g, 15)
            ws.Cells(r, 17).Value = t(g, 16)
            ws.Cells(r, 18).Value = KS_Harf(t(g, 16))
            r = r + 1
        Next i
        With ws.Range(ws.Cells(P_BASLIK + 1, 2), ws.Cells(r - 1, 18))
            KS_TabloGovde .Cells
            .RowHeight = 20
        End With
        ws.Range(ws.Cells(P_BASLIK + 1, 2), ws.Cells(r - 1, 2)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(P_BASLIK + 1, 2), ws.Cells(r - 1, 2)).Font.Color = C_GRI()
        ws.Range(ws.Cells(P_BASLIK + 1, 3), ws.Cells(r - 1, 3)).Font.Bold = True
        ws.Range(ws.Cells(P_BASLIK + 1, 4), ws.Cells(r - 1, 5)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(P_BASLIK + 1, 6), ws.Cells(r - 1, 6)).NumberFormat = KS_FmtTLKisa()
        ws.Range(ws.Cells(P_BASLIK + 1, 7), ws.Cells(r - 1, 7)).NumberFormat = "0%"
        ws.Range(ws.Cells(P_BASLIK + 1, 8), ws.Cells(r - 1, 8)).NumberFormat = "0.0"
        ws.Range(ws.Cells(P_BASLIK + 1, 9), ws.Cells(r - 1, 10)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(P_BASLIK + 1, 11), ws.Cells(r - 1, 11)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(P_BASLIK + 1, 12), ws.Cells(r - 1, 12)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(P_BASLIK + 1, 13), ws.Cells(r - 1, 14)).NumberFormat = "0.0"
        ws.Range(ws.Cells(P_BASLIK + 1, 15), ws.Cells(r - 1, 15)).NumberFormat = "#,##0"
        ws.Range(ws.Cells(P_BASLIK + 1, 16), ws.Cells(r - 1, 16)).NumberFormat = KS_FmtTLKisa()
        ws.Range(ws.Cells(P_BASLIK + 1, 17), ws.Cells(r - 1, 17)).NumberFormat = "0.0"
        ws.Range(ws.Cells(P_BASLIK + 1, 17), ws.Cells(r - 1, 18)).Font.Bold = True
        ws.Range(ws.Cells(P_BASLIK + 1, 17), ws.Cells(r - 1, 18)).HorizontalAlignment = xlCenter
        KS_PerfRenk ws.Range(ws.Cells(P_BASLIK + 1, 11), ws.Cells(r - 1, 11)), _
            ws.Range(ws.Cells(P_BASLIK + 1, 17), ws.Cells(r - 1, 17)), ws.Range(ws.Cells(P_BASLIK + 1, 18), ws.Cells(r - 1, 18))
    End If
    r = r + 1
    ws.Cells(r, 3).Value = "Kalite skoru = 100 - en çok 50 (arýza oraný % x " & k1 & ") - en çok 25 (ort. gecikme gün x " & _
        k2 & ") - en çok 25 (ort. üretici dönüþ gün x " & k3 & ").   Not: A >= 85, B >= 70, C >= 50, D < 50."
    ws.Cells(r + 1, 3).Value = "Arýza oraný = arýzalý miktar / sevk miktarý (ayný dönem). Az sevkiyatý olan firmalarda tek arýza oraný belirgin yükseltir; skor birlikte okunmalýdýr."
    ws.Range(ws.Cells(r, 3), ws.Cells(r + 1, 3)).Font.Size = 8
    ws.Range(ws.Cells(r, 3), ws.Cells(r + 1, 3)).Font.Italic = True
    ws.Range(ws.Cells(r, 3), ws.Cells(r + 1, 3)).Font.Color = C_GRI()
    If korumali Then KS_Kapat ws
    KS_Bitir
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Performans hesaplanamadý: " & Err.Description, vbExclamation, KS_AD
End Sub

Private Function KS_GrupNo(ByVal d As Object, ByRef t() As Variant, ByRef n As Long, ByVal ad As String) As Long
    Dim k As String, j As Long
    If Len(ad) = 0 Then ad = "(Belirtilmemiþ)"
    k = KS_Anahtar(ad)
    If Not d.Exists(k) Then
        n = n + 1
        d(k) = n
        t(n, 1) = ad
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

Private Sub KS_PerfRenk(ByVal oran As Range, ByVal skor As Range, ByVal harf As Range)
    Dim cs As Object, db As Object, fc As Object
    On Error Resume Next
    Set db = oran.FormatConditions.AddDatabar
    db.BarColor.Color = RGB(255, 140, 150)
    db.BarFillType = 0
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
    fc.Font.Color = C_TURKUAZ()
    Set fc = harf.FormatConditions.Add(Type:=xlTextString, String:="C", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_TURUNCU()
    Set fc = harf.FormatConditions.Add(Type:=xlTextString, String:="D", TextOperator:=xlBeginsWith)
    fc.Font.Color = C_KIRMIZI()
    On Error GoTo 0
End Sub
