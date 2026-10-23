Attribute VB_Name = "modKS_PANEL"
Option Explicit

' PANEL: özet kartlarý, týklanýnca ilgili listeyi süzerek açan uyarý kartlarý ve adet / tutar özet tablolarý.
' Ayrýca sayfalar arasý geçiþ ve hazýr süzgeç makrolarý bu modüldedir.

Private Const P_DONEM As Long = 3      ' dönem seçimi satýrý
Private Const P_BOLUM As Long = 5      ' tablo bölüm baþlýklarý
Private Const P_TABLO As Long = 6      ' tablo baþlýk satýrý

' ================================================================ kurulum

Public Sub KS_PanelKur()
    Dim ws As Worksheet, alt As Double, sol As Double, ust As Double, gen As Double, i As Long
    Dim kpi As Variant, uy As Variant, w As Double, fark As Double
    Set ws = KS_WS(SH_PANEL)
    KS_SayfaSifirla ws
    ' Sütun ýzgarasý: sol tablo B:G, boþluk H, sað tablo I:M
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C").ColumnWidth = 13
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 11
    ws.Columns("F").ColumnWidth = 11
    ws.Columns("G").ColumnWidth = 14
    ws.Columns("H").ColumnWidth = 3
    ws.Columns("I").ColumnWidth = 34
    ws.Columns("J").ColumnWidth = 12
    ws.Columns("K").ColumnWidth = 19
    ws.Columns("L").ColumnWidth = 10
    ws.Columns("M").ColumnWidth = 20
    ws.Columns("N").ColumnWidth = 2
    ' Sað kenarý banner geniþliðine eþitle.
    For i = 1 To 6
        fark = KS_GENISLIK - (ws.Columns("N").Left - ws.Columns("B").Left)
        If Abs(fark) < 3 Then Exit For
        w = ws.Columns("M").ColumnWidth + fark / 5.6
        If w < 6 Then w = 6
        ws.Columns("M").ColumnWidth = w
    Next i

    alt = KS_SayfaCerceve(ws, True)
    ws.Rows(1).RowHeight = alt + 4
    sol = ws.Cells(1, 2).Left
    ws.Rows(2).RowHeight = 330
    ust = ws.Rows(2).Top

    ' Karþýlama
    KS_Yazi ws, "ksP_Hos", "HOÞ GELDÝNÝZ", sol, ust + 10, 300, 14, 8.5, True, C_TEAL(), msoAlignLeft
    KS_Yazi ws, "ksP_Baslik", "Kalite ve garanti durumu bir bakýþta", sol, ust + 24, 700, 30, 19, True, C_LAC(), msoAlignLeft
    KS_Yazi ws, "ksP_Alt", "Kartlara týklayarak ilgili listeyi süzülmüþ olarak açabilirsiniz.", _
        sol, ust + 56, 700, 16, 9, False, C_GRI(), msoAlignLeft

    ' Özet kartlarý
    gen = (KS_GENISLIK - 3 * 14) / 4
    kpi = Array( _
        Array("SEVKÝYAT", C_MAVI(), "KS_GIT_SEVK"), _
        Array("ARIZA", C_KIRMIZI(), "KS_GIT_ARIZA"), _
        Array("GARANTÝ", C_YESIL2(), "KS_F_SEVK_GARDEVAM"), _
        Array("ÜRETÝCÝ DÖNÜÞÜ", C_TEAL(), "KS_GIT_KARNE"))
    For i = 0 To 3
        KS_PanelKart ws, "ksP_K" & i, sol + i * (gen + 14), ust + 84, gen, 100, CStr(kpi(i)(0)), CLng(kpi(i)(1)), _
            CStr(kpi(i)(2)), False
    Next i

    ' Uyarý kartlarý
    KS_Yazi ws, "ksP_UyBaslik", "UYARI MERKEZÝ", sol, ust + 200, 300, 16, 10, True, C_LAC(), msoAlignLeft
    uy = Array( _
        Array("GARANTÝSÝ BÝTECEK", RGB(214, 120, 10), "KS_F_SEVK_GARBITIYOR"), _
        Array("TESLÝMATI GECÝKEN", C_KIRMIZI(), "KS_F_SEVK_GECIKEN"), _
        Array("ÜRETÝCÝDE BEKLEYEN ARIZA", C_TEAL(), "KS_F_ARIZA_URETICIDE"), _
        Array("GARANTÝSÝ GÝRÝLMEMÝÞ", C_MAVI(), "KS_F_SEVK_GARYOK"))
    For i = 0 To 3
        KS_PanelKart ws, "ksP_U" & i, sol + i * (gen + 14), ust + 222, gen, 96, CStr(uy(i)(0)), CLng(uy(i)(1)), _
            CStr(uy(i)(2)), True
    Next i

    ' Dönem seçimi
    ws.Rows(P_DONEM).RowHeight = 30
    With ws.Range("B" & P_DONEM)
        .Value = "DÖNEM (tablolar ve akýþ kartlarý)"
        .Font.Size = 8.5
        .Font.Bold = True
        .Font.Color = C_GRI()
        .HorizontalAlignment = xlRight
        .IndentLevel = 1
    End With
    With ws.Range("C" & P_DONEM)
        .Value = "TÜMÜ"
        .Interior.Color = vbWhite
        .Font.Bold = True
        .Font.Color = C_LAC()
        .HorizontalAlignment = xlCenter
        .Locked = False
        .BorderAround xlContinuous, xlThin, , RGB(190, 204, 222)
    End With
    KS_AdHucre "pYil", ws.Range("C" & P_DONEM)
    KS_Liste ws.Range("C" & P_DONEM), "lstYil", True
    KS_Buton ws, "ksP_Yenile", "YENÝLE", ws.Range("D" & P_DONEM).Left + 8, ws.Rows(P_DONEM).Top + 3, 90, 24, _
        "KS_PANEL_YENILE", "ana"
    KS_Buton ws, "ksP_Yedek", "YEDEK AL", ws.Range("D" & P_DONEM).Left + 104, ws.Rows(P_DONEM).Top + 3, 90, 24, _
        "KS_YEDEK_AL", "cizgi"
    KS_Buton ws, "ksP_Mod", "EXCEL / UYGULAMA", ws.Range("D" & P_DONEM).Left + 200, ws.Rows(P_DONEM).Top + 3, 128, 24, _
        "KS_UYGULAMA_MODU_DEGISTIR", "cizgi"
    ws.Range("I" & P_DONEM & ":M" & P_DONEM).Merge
    With ws.Range("I" & P_DONEM)
        .HorizontalAlignment = xlRight
        .Font.Size = 8.5
        .Font.Color = C_GRI()
    End With
    KS_AdHucre "pGuncel", ws.Range("I" & P_DONEM)
    ws.Rows(P_DONEM + 1).RowHeight = 8

    ' Tablo iskeletleri
    KS_Bolum ws.Range("B" & P_BOLUM & ":G" & P_BOLUM), "ÝL BAZINDA ÖZET  (sevk yerleri ile birleþik)"
    KS_Bolum ws.Range("I" & P_BOLUM & ":M" & P_BOLUM), "DURUM ÖZETÝ  (kalem / tutar)"
    KS_TabloBaslik ws.Range("B" & P_TABLO & ":G" & P_TABLO), _
        Array("Ýl / Sevk Yeri", "Sevk Kalemi", "Sevk Tutarý (TL)", "Arýza", "Açýk Arýza", "Garantisi Bitiyor")
    KS_TabloBaslik ws.Range("I" & P_TABLO & ":M" & P_TABLO), _
        Array("Durum", "Kalem", "Tutar / Maliyet (TL)", "Pay", "Açýklama")
    KS_ZeminUygula ws, True
    KS_Pencere ws, 90, 0
End Sub

' Kart: beyaz gövde, solda renkli þerit, etiket / büyük deðer / alt satýr; týklanabilir.
Private Sub KS_PanelKart(ByVal ws As Worksheet, ByVal ad As String, ByVal sol As Double, ByVal ust As Double, _
        ByVal gen As Double, ByVal yuk As Double, ByVal etiket As String, ByVal renk As Long, _
        ByVal makro As String, ByVal uyari As Boolean)
    Dim s As Shape
    Set s = KS_KartSekil(ws, ad, sol, ust, gen, yuk)
    s.OnAction = makro
    Set s = KS_Serit(ws, ad & "_S", sol, ust + 12, 4, yuk - 24, renk)
    s.OnAction = makro
    Set s = KS_Yazi(ws, ad & "_E", etiket, sol + 18, ust + 12, gen - 30, 14, 8.5, True, renk, msoAlignLeft)
    s.OnAction = makro
    Set s = KS_Yazi(ws, ad & "_D", "-", sol + 18, ust + 28, gen - 30, 34, 22, True, C_LAC(), msoAlignLeft)
    s.OnAction = makro
    Set s = KS_Yazi(ws, ad & "_A", "", sol + 18, ust + yuk - 30, gen - 30, 28, 8.5, False, C_GRI(), msoAlignLeft)
    s.OnAction = makro
    If uyari Then
        Set s = KS_Yazi(ws, ad & "_G", "Listeyi aç  " & ChrW(8594), sol + gen - 110, ust + 14, 96, 14, 8.5, True, _
            C_MAVI(), msoAlignRight)
        s.OnAction = makro
    End If
End Sub

' ================================================================ yenileme

Public Sub KS_PANEL_YENILE()
    KS_PanelYenile True
End Sub

Public Sub KS_PanelYenile(ByVal mesaj As Boolean)
    Dim ws As Worksheet, arS As Variant, aA As Variant, t As Variant, korumali As Boolean
    Dim yilS As String, y1 As Long, y2 As Long, i As Long, x As Double, d As Date, uyari As Long
    Dim garDevam As Long, garBit As Long, garBitTut As Double, gecik As Long, gecikTut As Double
    Dim ureticide As Long, ureticideMal As Double, garYok As Long, garYokTut As Double, acik As Long, garIc As Long
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
    arS = KS_Veri(KS_LOZ(T_SEVK))
    aA = KS_Veri(KS_LOZ(T_ARIZA))
    uyari = CLng(KS_ParamSayi("pGarUyari", 60))
    t = KS_Topla(arS, aA, y1, y2)

    ' Güncel durum (dönemden baðýmsýz)
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            x = 0
            If KS_Sayi(arS(i, S_TUTAR), x) Then x = x
            Select Case True
                Case Left$(KS_Metin(arS(i, S_GDURUM)), 7) = "BÝTÝYOR"
                    garBit = garBit + 1: garBitTut = garBitTut + x: garDevam = garDevam + 1
                Case KS_Metin(arS(i, S_GDURUM)) = "DEVAM EDÝYOR"
                    garDevam = garDevam + 1
                Case KS_Metin(arS(i, S_GDURUM)) = "GARANTÝ GÝRÝLMEDÝ"
                    garYok = garYok + 1: garYokTut = garYokTut + x
            End Select
            If KS_Metin(arS(i, S_TDURUM)) = "GECÝKÝYOR" Then gecik = gecik + 1: gecikTut = gecikTut + x
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_ArizaAcik(aA(i, A_DURUM)) Then acik = acik + 1
            If KS_Metin(aA(i, A_DURUM)) = "ÜRETÝCÝDE" Then
                ureticide = ureticide + 1
                If KS_Sayi(aA(i, A_MALIYET), x) Then ureticideMal = ureticideMal + x
            End If
        Next i
    End If

    korumali = ws.ProtectContents
    KS_Ac ws
    KS_KartDeger ws, "ksP_K0", Format$(t(0), "#,##0") & " kalem", _
        KS_TL(t(2)) & "   ·   " & Format$(t(1), "#,##0") & " birim" & vbLf & IIf(y1 = 0, "Tüm kayýtlar", y1 & " yýlý")
    KS_KartDeger ws, "ksP_K1", Format$(t(5), "#,##0") & " kayýt", _
        acik & " açýk   ·   " & t(8) & " garanti kapsamýnda" & vbLf & "Tamir maliyeti " & KS_TL(t(12))
    KS_KartDeger ws, "ksP_K2", Format$(garDevam, "#,##0") & " kalem", _
        "garantisi devam ediyor" & vbLf & garBit & " kalem " & uyari & " gün içinde bitiyor"
    KS_KartDeger ws, "ksP_K3", IIf(IsEmpty(t(10)), "-", Format$(t(10), "0.0") & " gün"), _
        "ortalama üretici dönüþ süresi" & vbLf & ureticide & " arýza üreticide bekliyor"
    KS_KartDeger ws, "ksP_U0", Format$(garBit, "#,##0") & " kalem", KS_TL(garBitTut) & vbLf & "önümüzdeki " & uyari & " gün"
    KS_KartDeger ws, "ksP_U1", Format$(gecik, "#,##0") & " kalem", KS_TL(gecikTut) & vbLf & "planlanan tarihi geçti"
    KS_KartDeger ws, "ksP_U2", Format$(ureticide, "#,##0") & " arýza", "tamir maliyeti " & KS_TL(ureticideMal) & vbLf & _
        "dönüþ bekleniyor"
    KS_KartDeger ws, "ksP_U3", Format$(garYok, "#,##0") & " kalem", KS_TL(garYokTut) & vbLf & "garanti süresi elle girilmeli"

    KS_YerTablosu ws, arS, aA, y1, y2
    KS_DurumTablosu ws, arS, aA, y1, y2, uyari

    ws.Range("I" & P_DONEM).Value = "Son güncelleme: " & Format$(Now, "dd.mm.yyyy hh:nn") & KS_YedekNotu()
    ws.Range("I" & P_DONEM).Font.Color = IIf(InStr(ws.Range("I" & P_DONEM).Value, "YEDEK") > 0, C_TURUNCU(), C_GRI())
    ws.Range("I" & P_DONEM).Font.Bold = (InStr(ws.Range("I" & P_DONEM).Value, "YEDEK") > 0)
    KS_PanelAlt ws

    If korumali Then KS_Kapat ws
    KS_Bitir
    If mesaj Then ws.Activate
    Exit Sub
Hata:
    KS_Bitir True
    If mesaj Then MsgBox "Panel yenilenemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Public Function KS_TL(ByVal v As Variant) As String
    Dim x As Double
    If KS_Sayi(v, x) Then KS_TL = Format$(x, "#,##0") & " " & ChrW(8378) Else KS_TL = "0 " & ChrW(8378)
End Function

Private Sub KS_KartDeger(ByVal ws As Worksheet, ByVal ad As String, ByVal deger As String, ByVal alt As String)
    KS_YaziGuncelle ws, ad & "_D", deger
    KS_YaziGuncelle ws, ad & "_A", alt
End Sub

Private Sub KS_SatirTemizle(ByVal ws As Worksheet, ByVal ilk As Long, ByVal sonKol As String, ByVal ilkKol As String)
    With ws.Range(ilkKol & ilk & ":" & sonKol & ilk + 40)
        .ClearContents
        .Interior.Pattern = xlNone
        .Borders.LineStyle = xlNone
        .Font.Bold = False
        .Font.Italic = False
        .Font.Color = C_METIN()
        .IndentLevel = 0
    End With
End Sub

' Ýl bazýnda özet + toplam satýrý. Ayný ildeki sevk yerleri (ör. ÇEDAÞ TRABZON ve Çoruh Aksa- Trabzon ÜY) birleþir;
' ili girilmemiþ sevk yeri kendi adýyla ayrý satýrdýr. Çoruh illeri her zaman, diðerleri kaydý varsa görünür.
Private Sub KS_YerTablosu(ByVal ws As Worksheet, ByRef arS As Variant, ByRef aA As Variant, ByVal y1 As Long, ByVal y2 As Long)
    Dim yer As Variant, d As Object, dYer As Object, i As Long, j As Long, n As Long, k As String, r As Long
    Dim x As Double, dt As Date, v() As Variant, ad() As String, grup As String, il As Variant
    Const SABIT As Long = 5
    KS_SatirTemizle ws, P_TABLO + 1, "G", "B"
    Set d = CreateObject("Scripting.Dictionary")
    Set dYer = CreateObject("Scripting.Dictionary")
    yer = KS_Veri(KS_LOZ(T_YER))
    ReDim ad(1 To KS_Satir(yer) + SABIT + 2)
    For Each il In Array("Artvin", "Giresun", "Gümüþhane", "Rize", "Trabzon")
        n = n + 1
        d(KS_Anahtar(il)) = n
        ad(n) = CStr(il)
    Next il
    If IsArray(yer) Then
        For i = 1 To UBound(yer, 1)
            k = KS_Anahtar(yer(i, 1))
            If Len(k) > 0 Then
                grup = Trim$(KS_Metin(yer(i, 3)))
                If Len(grup) = 0 Then grup = KS_Metin(yer(i, 1))
                If Not d.Exists(KS_Anahtar(grup)) Then
                    n = n + 1
                    d(KS_Anahtar(grup)) = n
                    ad(n) = grup
                End If
                dYer(k) = d(KS_Anahtar(grup))
            End If
        Next i
    End If
    Set d = dYer
    ReDim v(1 To n + 2, 1 To 6)
    For i = 1 To n
        v(i, 1) = ad(i)
    Next i
    v(n + 1, 1) = "Diðer / tanýmsýz"
    For i = 1 To n + 1
        For j = 2 To 6
            v(i, j) = 0
        Next j
    Next i
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            j = KS_SevkYil(arS, i)
            If j >= y1 And j <= y2 Then
                k = KS_Anahtar(arS(i, S_YER))
                If d.Exists(k) Then r = d(k) Else r = n + 1
                v(r, 2) = v(r, 2) + 1
                If KS_Sayi(arS(i, S_TUTAR), x) Then v(r, 3) = v(r, 3) + x
            End If
            If Left$(KS_Metin(arS(i, S_GDURUM)), 7) = "BÝTÝYOR" Then
                k = KS_Anahtar(arS(i, S_YER))
                If d.Exists(k) Then r = d(k) Else r = n + 1
                v(r, 6) = v(r, 6) + 1
            End If
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            k = KS_Anahtar(aA(i, A_YER))
            If d.Exists(k) Then r = d(k) Else r = n + 1
            If KS_Tarih(aA(i, A_TARIH), dt) Then
                If Year(dt) >= y1 And Year(dt) <= y2 Then v(r, 4) = v(r, 4) + 1
            End If
            If KS_ArizaAcik(aA(i, A_DURUM)) Then v(r, 5) = v(r, 5) + 1
        Next i
    End If
    v(n + 2, 1) = "TOPLAM"
    For j = 2 To 6
        v(n + 2, j) = 0
        For i = 1 To n + 1
            v(n + 2, j) = v(n + 2, j) + v(i, j)
        Next i
    Next j
    r = P_TABLO + 1
    For i = 1 To n + 2
        If i <= SABIT Or i = n + 2 Or v(i, 2) + v(i, 4) + v(i, 5) + v(i, 6) > 0 Then
            For j = 1 To 6
                ws.Cells(r, 1 + j).Value = v(i, j)
            Next j
            r = r + 1
        End If
    Next i
    KS_TabloGovde ws.Range("B" & P_TABLO + 1 & ":G" & r - 1)
    ws.Range("B" & P_TABLO + 1 & ":B" & r - 1).IndentLevel = 1
    ws.Range("C" & P_TABLO + 1 & ":G" & r - 1).HorizontalAlignment = xlCenter
    ws.Range("C" & P_TABLO + 1 & ":C" & r - 1).NumberFormat = "#,##0"
    ws.Range("D" & P_TABLO + 1 & ":D" & r - 1).NumberFormat = "#,##0"
    ws.Range("E" & P_TABLO + 1 & ":G" & r - 1).NumberFormat = "#,##0"
    With ws.Range("B" & r - 1 & ":G" & r - 1)
        .Font.Bold = True
        .Interior.Color = C_ACIKMAVI()
        .Borders(xlEdgeTop).Color = C_LAC()
    End With
    ws.Range("F" & P_TABLO + 1 & ":F" & r - 2).Font.Color = C_KIRMIZI()
    ws.Range("G" & P_TABLO + 1 & ":G" & r - 2).Font.Color = C_TURUNCU()
End Sub

' Teslimat / garanti / arýza durum daðýlýmý.
Private Sub KS_DurumTablosu(ByVal ws As Worksheet, ByRef arS As Variant, ByRef aA As Variant, ByVal y1 As Long, _
        ByVal y2 As Long, ByVal uyari As Long)
    Dim grup As Variant, i As Long, j As Long, g As Long, r As Long, x As Double, dt As Date, top As Double
    Dim adet(1 To 3, 1 To 5) As Double, tut(1 To 3, 1 To 5) As Double, s As String, idx As Long
    KS_SatirTemizle ws, P_TABLO + 1, "M", "I"
    grup = Array( _
        Array("TESLÝMAT", Array("Zamanýnda", "Gecikmeli", "Gecikiyor (gelmedi)", "Bekleniyor", "Tarih yok"), _
            Array("planlanan tarihte / öncesinde", "geç teslim edildi", "planlanan tarih geçti", "tarihi henüz gelmedi", "tarih girilmemiþ")), _
        Array("GARANTÝ", Array("Devam ediyor", "Bitiyor (" & uyari & " gün)", "Bitti", "Garanti girilmemiþ", "Teslim bekliyor"), _
            Array("garanti süresi içinde", "yakýnda bitecek", "süresi doldu", "MALZEMELER'den girin", "geliþ tarihi yok")), _
        Array("ARIZA", Array("Devam ediyor", "Üreticide", "Kýsmen çözüldü", "Çözüldü", "Çözülemedi"), _
            Array("iþlem sürüyor", "üretici dönüþü bekleniyor", "kýsmi çözüm", "kapandý", "çözümsüz kapandý")))
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            j = KS_SevkYil(arS, i)
            If j >= y1 And j <= y2 Then
                x = 0
                If KS_Sayi(arS(i, S_TUTAR), x) Then x = x
                Select Case KS_Metin(arS(i, S_TDURUM))
                    Case "ZAMANINDA": idx = 1
                    Case "GECÝKMELÝ": idx = 2
                    Case "GECÝKÝYOR": idx = 3
                    Case "BEKLENÝYOR": idx = 4
                    Case Else: idx = 5
                End Select
                adet(1, idx) = adet(1, idx) + 1: tut(1, idx) = tut(1, idx) + x
                s = KS_Metin(arS(i, S_GDURUM))
                If s = "DEVAM EDÝYOR" Then
                    idx = 1
                ElseIf Left$(s, 7) = "BÝTÝYOR" Then
                    idx = 2
                ElseIf s = "BÝTTÝ" Then
                    idx = 3
                ElseIf s = "GARANTÝ GÝRÝLMEDÝ" Then
                    idx = 4
                Else
                    idx = 5
                End If
                adet(2, idx) = adet(2, idx) + 1: tut(2, idx) = tut(2, idx) + x
            End If
        Next i
    End If
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            If KS_Tarih(aA(i, A_TARIH), dt) Then
                If Year(dt) >= y1 And Year(dt) <= y2 Then
                    x = 0
                    If KS_Sayi(aA(i, A_MALIYET), x) Then x = x
                    Select Case KS_Metin(aA(i, A_DURUM))
                        Case "DEVAM EDÝYOR": idx = 1
                        Case "ÜRETÝCÝDE": idx = 2
                        Case "KISMEN ÇÖZÜLDÜ": idx = 3
                        Case "ÇÖZÜLDÜ": idx = 4
                        Case Else: idx = 5
                    End Select
                    adet(3, idx) = adet(3, idx) + 1: tut(3, idx) = tut(3, idx) + x
                End If
            End If
        Next i
    End If
    r = P_TABLO + 1
    For g = 0 To 2
        With ws.Range("I" & r & ":M" & r)
            .Interior.Color = C_ACIKMAVI()
            .Font.Bold = True
            .Font.Color = C_LAC()
            .Borders(xlEdgeBottom).Color = RGB(200, 214, 230)
        End With
        ws.Cells(r, "I").Value = grup(g)(0)
        ws.Cells(r, "I").IndentLevel = 1
        top = 0
        For j = 1 To 5
            top = top + adet(g + 1, j)
        Next j
        ws.Cells(r, "J").Value = top
        ws.Cells(r, "K").Value = tut(g + 1, 1) + tut(g + 1, 2) + tut(g + 1, 3) + tut(g + 1, 4) + tut(g + 1, 5)
        ws.Range("J" & r & ":K" & r).NumberFormat = "#,##0"
        ws.Range("J" & r & ":L" & r).HorizontalAlignment = xlCenter
        r = r + 1
        For j = 1 To 5
            ws.Cells(r, "I").Value = grup(g)(1)(j - 1)
            ws.Cells(r, "I").IndentLevel = 2
            ws.Cells(r, "J").Value = adet(g + 1, j)
            ws.Cells(r, "K").Value = tut(g + 1, j)
            If top > 0 Then ws.Cells(r, "L").Value = adet(g + 1, j) / top Else ws.Cells(r, "L").Value = Empty
            ws.Cells(r, "M").Value = grup(g)(2)(j - 1)
            ws.Range("I" & r & ":M" & r).Interior.Color = vbWhite
            ws.Range("I" & r & ":M" & r).Borders(xlEdgeBottom).Color = RGB(232, 238, 245)
            ws.Range("J" & r & ":K" & r).NumberFormat = "#,##0"
            ws.Cells(r, "L").NumberFormat = "0%"
            ws.Range("J" & r & ":L" & r).HorizontalAlignment = xlCenter
            ws.Cells(r, "M").Font.Size = 8.5
            ws.Cells(r, "M").Font.Color = C_GRI()
            r = r + 1
        Next j
    Next g
    ws.Range("I" & P_TABLO + 1 & ":M" & r - 1).BorderAround xlContinuous, xlThin, , RGB(205, 216, 230)
    ws.Range("J" & P_TABLO + 3 & ":J" & P_TABLO + 4).Font.Color = C_KIRMIZI()
    ws.Range("J" & P_TABLO + 8 & ":J" & P_TABLO + 8).Font.Color = C_TURUNCU()
    ws.Range("J" & P_TABLO + 10 & ":J" & P_TABLO + 10).Font.Color = C_MAVI()
End Sub

' Sayfa sonu bandý (AKSORA tarzý lacivert alt bilgi).
Private Sub KS_PanelAlt(ByVal ws As Worksheet)
    Dim s As Shape, ust As Double, son As Long
    son = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    If ws.Cells(ws.Rows.Count, "B").End(xlUp).Row > son Then son = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    ust = ws.Rows(son + 2).Top
    Set s = KS_Serit(ws, "ksP_Alt1", ws.Cells(1, 2).Left, ust, KS_GENISLIK, 34, C_LAC())
    Set s = KS_Yazi(ws, "ksP_Alt2", KS_AD & "   /   Kalite, garanti, güven.", ws.Cells(1, 2).Left + 18, ust + 10, 500, 16, _
        8.5, True, vbWhite, msoAlignLeft)
    Set s = KS_Yazi(ws, "ksP_Alt3", "ÇORUH ELEKTRÝK DAÐITIM A.Þ.   ·   v" & KS_SURUM, ws.Cells(1, 2).Left + KS_GENISLIK - 418, _
        ust + 10, 400, 16, 8.5, True, RGB(110, 215, 240), msoAlignRight)
End Sub

' ================================================================ gezinme

Private Sub KS_Git(ByVal ad As String)
    Dim ws As Worksheet
    Set ws = KS_WS(ad)
    If ws Is Nothing Then Exit Sub
    ws.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Public Sub KS_GIT_MALZ()
    KS_Git SH_MALZ
End Sub

Public Sub KS_GIT_SEVK()
    KS_Git SH_SEVK
End Sub

Public Sub KS_GIT_ARIZA()
    KS_Git SH_ARIZA
End Sub

Public Sub KS_GIT_KARNE()
    KS_Git SH_KARNE
End Sub

Public Sub KS_GIT_RAPOR()
    KS_Git SH_RAPOR
End Sub

Public Sub KS_GIT_TANIM()
    KS_Git SH_TANIM
End Sub

Public Sub KS_GIT_AKTAR()
    KS_Git SH_AKTAR
End Sub

Public Sub KS_GIT_LOG()
    KS_Git SH_LOG
End Sub

' ================================================================ hazýr süzgeçler

' Tabloyu açar, önceki süzgeçleri kaldýrýr, verilen sütuna ölçüt uygular.
Private Sub KS_Suz(ByVal tablo As String, ByVal kolon As Long, ByVal olcut As Variant, Optional ByVal coklu As Boolean = False)
    Dim lo As ListObject, ws As Worksheet, korumali As Boolean
    On Error GoTo Hata
    Set lo = KS_LOZ(tablo)
    Set ws = lo.Parent
    korumali = ws.ProtectContents
    KS_Ac ws
    If lo.ShowAutoFilter Then
        If ws.FilterMode Then lo.AutoFilter.ShowAllData
    Else
        lo.ShowAutoFilter = True
    End If
    If Not lo.DataBodyRange Is Nothing And kolon > 0 Then
        If coklu Then
            lo.Range.AutoFilter Field:=kolon, Criteria1:=olcut, Operator:=xlFilterValues
        Else
            lo.Range.AutoFilter Field:=kolon, Criteria1:=olcut
        End If
    End If
    If korumali Then KS_Kapat ws
    ws.Activate
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
    Exit Sub
Hata:
    If Not ws Is Nothing Then
        If korumali Then KS_Kapat ws
    End If
    MsgBox "Liste süzülemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Public Sub KS_F_TEMIZLE()
    Select Case ActiveSheet.Name
        Case SH_SEVK: KS_Suz T_SEVK, 0, Empty
        Case SH_ARIZA: KS_Suz T_ARIZA, 0, Empty
        Case SH_MALZ: KS_Suz T_MALZ, 0, Empty
    End Select
End Sub

Public Sub KS_F_SEVK_GARBITIYOR()
    KS_Suz T_SEVK, S_GDURUM, "BÝTÝYOR*"
End Sub

Public Sub KS_F_SEVK_GARDEVAM()
    KS_Suz T_SEVK, S_GDURUM, Array("DEVAM EDÝYOR"), True
End Sub

Public Sub KS_F_SEVK_GECIKEN()
    KS_Suz T_SEVK, S_TDURUM, "GECÝKÝYOR"
End Sub

Public Sub KS_F_SEVK_GARYOK()
    KS_Suz T_SEVK, S_GDURUM, "GARANTÝ GÝRÝLMEDÝ"
End Sub

Public Sub KS_F_ARIZA_ACIK()
    KS_Suz T_ARIZA, A_DURUM, Array("DEVAM EDÝYOR", "ÜRETÝCÝDE", "KISMEN ÇÖZÜLDÜ"), True
End Sub

Public Sub KS_F_ARIZA_URETICIDE()
    KS_Suz T_ARIZA, A_DURUM, "ÜRETÝCÝDE"
End Sub

Public Sub KS_F_ARIZA_COZULEN()
    KS_Suz T_ARIZA, A_DURUM, "ÇÖZÜLDÜ"
End Sub

Public Sub KS_F_ARIZA_GARANTI()
    KS_Suz T_ARIZA, A_KAPSAM, "EVET"
End Sub

Public Sub KS_F_MALZ_GARYOK()
    KS_Suz T_MALZ, M_GARANTI, "="
End Sub

Public Sub KS_F_MALZ_ARIZALI()
    KS_Suz T_MALZ, M_ARIZAMIK, ">0"
End Sub

' ================================================================ arama

' Aktif listede metni sýrayla aday sütunlarda arar; ilk eþleþen sütuna "*metin*" süzgeci uygular.
Public Sub KS_ARA()
    Dim lo As ListObject, aday As Variant, k As Variant, a As Variant, i As Long, ara As String, bulunan As Long
    Select Case ActiveSheet.Name
        Case SH_SEVK
            Set lo = KS_LOZ(T_SEVK)
            aday = Array(S_KOD, S_AD, S_NO, S_SEVKCI, S_YER, S_SERI, S_IRS)
        Case SH_ARIZA
            Set lo = KS_LOZ(T_ARIZA)
            aday = Array(A_KOD, A_AD, A_NO, A_SERI, A_URETICI, A_YER, A_ACIK)
        Case SH_MALZ
            Set lo = KS_LOZ(T_MALZ)
            aday = Array(M_KOD, M_AD, M_URETICI, M_KAT)
        Case Else
            Exit Sub
    End Select
    ara = Trim$(InputBox("Aranacak metin (malzeme kodu / adý, numara, sevkçi, seri no ...):", KS_AD))
    If Len(ara) = 0 Then Exit Sub
    a = KS_Veri(lo)
    If IsArray(a) Then
        For Each k In aday
            For i = 1 To UBound(a, 1)
                If InStr(1, KS_Anahtar(a(i, k)), KS_Anahtar(ara), vbBinaryCompare) > 0 Then
                    bulunan = k
                    Exit For
                End If
            Next i
            If bulunan > 0 Then Exit For
        Next k
    End If
    If bulunan = 0 Then
        MsgBox "Eþleþen kayýt bulunamadý: " & ara, vbInformation, KS_AD
        Exit Sub
    End If
    KS_Suz lo.Name, bulunan, "*" & ara & "*"
End Sub
