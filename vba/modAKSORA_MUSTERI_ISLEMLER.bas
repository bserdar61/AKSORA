Attribute VB_Name = "modAKSORA_MUSTERI_ISLEMLER"
Option Explicit
Private mHazir As Boolean

Private Function M_SAYFA(ByVal ad As String) As Boolean
    Select Case ad
        Case "M_ANA PANEL", "M_MALZEME_ARA", "M_CDP_STOK", "M_CDP_RAPOR", _
             "M_ÝLÇE_STOK", "M_ÝLÇE_RAPOR", "M_YÜKLENÝCÝ_STOK", "M_YÜKLENÝCÝ_RAPOR"
            M_SAYFA = True
    End Select
End Function

Private Function M_STOK(ByVal ad As String) As Boolean
    M_STOK = (ad = "M_CDP_STOK" Or ad = "M_ÝLÇE_STOK" Or ad = "M_YÜKLENÝCÝ_STOK")
End Function

Private Function M_METIN(ByVal s As Shape) As String
    On Error Resume Next
    M_METIN = Trim$(s.TextFrame2.TextRange.Text)
    On Error GoTo 0
End Function

Private Sub M_BAGLA(ByVal ws As Worksheet, ByVal s As Shape)
    Dim t As String, hedef As String, makro As String, i As Long
    If s.Type = msoGroup Then
        For i = 1 To s.GroupItems.Count
            M_BAGLA ws, s.GroupItems(i)
        Next i
        Exit Sub
    End If
    t = UCase$(M_METIN(s))
    hedef = s.AlternativeText
    If Not M_SAYFA(hedef) Then hedef = ""
    Select Case t
        Case "ANA PANEL": hedef = "M_ANA PANEL"
        Case "MALZEME ARAMA", "MALZEME ARA": hedef = "M_MALZEME_ARA"
        Case "CDP STOK": hedef = "M_CDP_STOK"
        Case "CDP RAPOR": hedef = "M_CDP_RAPOR"
        Case "ÝLÇE STOK": hedef = "M_ÝLÇE_STOK"
        Case "ÝLÇE RAPOR": hedef = "M_ÝLÇE_RAPOR"
        Case "YÜKLENÝCÝ STOK": hedef = "M_YÜKLENÝCÝ_STOK"
        Case "YÜKLENÝCÝ RAPOR": hedef = "M_YÜKLENÝCÝ_RAPOR"
    End Select
    If Len(hedef) > 0 Then
        s.AlternativeText = hedef
        makro = "AKSORA_MUSTERI_SAYFA_AC"
    ElseIf s.Name = "MUSTERI_TEMA_GERI" Then
        makro = "AKSORA_BIRIM_ANA_MENU"
    ElseIf ws.Name = "M_MALZEME_ARA" And s.Name = "btnMalzemeAra" Then
        makro = "AKSORA_MUSTERI_YEREL_ARA"
    End If
    If Len(makro) > 0 Then s.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & makro
End Sub

Public Sub AKSORA_MUSTERI_SAYFA_HAZIRLA(ByVal ws As Worksheet)
    Dim s As Shape, ev As Boolean, koruma As Boolean, hataNo As Long, hataMetni As String
    If mHazir Or Not M_SAYFA(ws.Name) Then Exit Sub
    mHazir = True: ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    koruma = ws.ProtectContents
    ws.Unprotect ""
    If ws.Name = "M_ÝLÇE_RAPOR" Or ws.Name = "M_YÜKLENÝCÝ_RAPOR" Then
        On Error Resume Next
        ws.Shapes("MUSTERI_TEMA_GERI").Delete
        On Error GoTo Hata
    End If
    If M_STOK(ws.Name) Then M_STOK_KUR ws
    If ws.Name = "M_MALZEME_ARA" Then M_ARAMA_TARIHLERI ws
    If ws.Name = "M_CDP_RAPOR" Or ws.Name = "M_ÝLÇE_RAPOR" Or ws.Name = "M_YÜKLENÝCÝ_RAPOR" Then M_BOS_RAPOR ws
    For Each s In ws.Shapes
        M_BAGLA ws, s
    Next s
    AKSORA_MUSTERI_SAP_BAGLA ws
Cikis:
    On Error Resume Next
    If koruma Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev: mHazir = False
    On Error GoTo 0
    If hataNo <> 0 Then Err.Raise hataNo, "AKSORA_MUSTERI_SAYFA_HAZIRLA", hataMetni
    Exit Sub
Hata:
    hataNo = Err.Number: hataMetni = Err.Description
    Resume Cikis
End Sub

Private Sub M_ARAMA_TARIHLERI(ByVal ws As Worksheet)
    Dim s As Shape, stok As Worksheet, ad As Variant, secili As String, kod As String, isim As String
    Dim tarih As Date, v As Variant, yazi As String
    secili = CStr(ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2)
    For Each ad In Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
        Set stok = ThisWorkbook.Worksheets(CStr(ad))
        v = stok.Range("C2").Value
        If Not IsError(v) Then
            If IsDate(v) Then
                If CDate(v) > tarih Then tarih = CDate(v)
            End If
        End If
    Next ad
    For Each s In ws.Shapes
        kod = s.AlternativeText: isim = ""
        Select Case kod
            Case "GIRESUN": isim = "GÝRESUN"
            Case "TRABZON": isim = "TRABZON"
            Case "RIZE": isim = "RÝZE"
            Case "GUMUSHANE": isim = "GÜMÜÞHANE"
            Case "ARTVIN": isim = "ARTVÝN"
        End Select
        If Len(isim) > 0 Then
            yazi = "-"
            If kod = secili And tarih > 0 Then yazi = Format$(tarih, "dd.mm.yyyy hh:nn")
            s.TextFrame2.TextRange.Text = isim & vbLf & yazi
        End If
    Next s
End Sub

Public Sub AKSORA_MUSTERI_BUTONLARI_AKTIF_ET()
    Dim ws As Worksheet
    On Error GoTo Hata
    For Each ws In ThisWorkbook.Worksheets
        If M_SAYFA(ws.Name) Then AKSORA_MUSTERI_SAYFA_HAZIRLA ws
    Next ws
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_MUSTERI_SAYFA_AC()
    Dim ws As Worksheet, hedef As Worksheet, ad As String
    On Error GoTo Hata
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Set ws = ActiveSheet
    If Not M_SAYFA(ws.Name) Then Exit Sub
    ad = ws.Shapes(CStr(application.Caller)).AlternativeText
    If Not M_SAYFA(ad) Then Err.Raise vbObjectError + 1800, , "Müþteri sayfa baðlantýsý geçersiz."
    Set hedef = ThisWorkbook.Worksheets(ad)
    hedef.Visible = xlSheetVisible
    hedef.Select Replace:=True
    AKSORA_MUSTERI_TEMA_GORUNUM hedef
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub M_STOK_KUR(ByVal ws As Worksheet)
    Dim baslik As Variant, agirlik As Variant, i As Long, son As Long, toplam As Double, w As Double
    Dim s As Shape, ad As String, eski As Shape, liste As Shape, x As Double, y As Double, h As Double, bw As Double
    Dim d As Object, k As Variant, r As Long, depo As String, secim As String, p As Long
    baslik = Array("ÜY", "DpYr", "Depo/Fikti", "Mal grubu", "Malzeme", _
        "Malzeme kýsa metni", "Miktar", "TÖB", "Deðerleme türü", "Seri numarasý")
    For i = ws.Shapes.Count To 1 Step -1
        If LCase$(Left$(ws.Shapes(i).Name, 9)) = "drop down" Then ws.Shapes(i).Delete
    Next i
    If CStr(ws.Range("BX8").Value2) <> "MUSTERI_STOK_V1" Then
        w = ws.Range("A1:L1").Width
        agirlik = Array(8, 9, 16, 14, 16, 46, 13, 8, 18, 23)
        For i = 0 To 9: toplam = toplam + CDbl(agirlik(i)): Next i
        ws.Columns("K:L").Hidden = False
        For i = 0 To 9: ws.Columns(i + 1).ColumnWidth = w / 7 * CDbl(agirlik(i)) / toplam: Next i
        bw = ws.Range("A1:J1").Width
        For i = 1 To 10: ws.Columns(i).ColumnWidth = ws.Columns(i).ColumnWidth * w / bw: Next i
        ws.Rows("3:5").Hidden = True
        ws.Columns("K:L").Hidden = True
        ws.Range("BX8").Value2 = "MUSTERI_STOK_V1"
    End If
    For i = 0 To 9: ws.Cells(6, i + 1).Value2 = baslik(i): Next i
    ws.Range("K6").Value2 = "Deðer (TL)"
    ws.Columns("I:J").NumberFormat = "@"
    ws.EnableSelection = xlNoRestrictions
    With ws.Range("A6:J6")
        .Interior.Color = RGB(26, 75, 120): .Font.Color = vbWhite
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = False: .ShrinkToFit = True
    End With
    son = application.Max(7, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    M_FILTRE_ALANI ws, son
    ws.Range("C3").Formula = "=SUBTOTAL(103,E7:E" & son & ")"
    ws.Range("C4").Formula = "=IF(SUBTOTAL(109,N7:N" & son & ")>0,""Fiyat eksik (""&SUBTOTAL(109,N7:N" & son & ")&"")"",SUBTOTAL(109,K7:K" & son & "))"
    ws.Range("C3").NumberFormat = "#,##0"
    ws.Range("C4").NumberFormat = "#,##0.00 ""TL"""
    AKSORA_V13_MIKTAR_BICIMI ws.Range("G7:G" & son)
    For Each s In ws.Shapes
        ad = s.Name
        If InStr(ad, "STOK_KPI_DEGER_") > 0 Then
            i = Val(Mid$(ad, InStrRev(ad, "_") + 1))
            If i >= 0 And i <= 2 Then s.DrawingObject.Formula = _
                "='" & Replace(ws.Name, "'", "''") & "'!C" & CStr(i + 2)
        End If
        If InStr(ad, "DEPO_KART_LISTE") > 0 Then Set eski = s
    Next s
    If eski Is Nothing Then Exit Sub
    ad = eski.Name
    If eski.Type <> msoFormControl Then
        x = eski.Left: y = eski.Top: bw = eski.Width: h = eski.Height
        eski.Delete
        Set liste = ws.Shapes.AddFormControl(xlDropDown, x, y, bw, h)
        liste.Name = ad: liste.Placement = xlFreeFloating
    Else
        Set liste = eski
        On Error Resume Next
        secim = CStr(liste.ControlFormat.List(liste.ControlFormat.Value))
        On Error GoTo 0
    End If
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = 7 To son
        If Not IsError(ws.Cells(r, "B").Value2) Then
            depo = Trim$(CStr(ws.Cells(r, "B").Value2))
            If Len(depo) > 0 Then d(depo) = True
        End If
    Next r
    With liste.ControlFormat
        .ListFillRange = "": .LinkedCell = "": .RemoveAllItems
        .AddItem "TÜMÜ"
        For Each k In d.keys: .AddItem CStr(k): Next k
        p = 1
        For i = 1 To .ListCount
            If CStr(.List(i)) = secim Then p = i
        Next i
        .Value = p: .DropDownLines = application.Min(15, .ListCount)
    End With
    liste.OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_MUSTERI_DEPO_SEC"
    M_DEPO_KONUM ws, liste
    ws.Calculate
End Sub

Private Sub M_FILTRE_ALANI(ByVal ws As Worksheet, ByVal son As Long)
    Dim filtre As AutoFilter, f As Filter, c1(1 To 11) As Variant, c2(1 To 11) As Variant
    Dim op(1 To 11) As Long, etkin(1 To 11) As Boolean, ikinci(1 To 11) As Boolean, i As Long
    If ws.AutoFilterMode Then
        Set filtre = ws.AutoFilter
        If filtre.Range.Row = 6 And filtre.Range.Rows.Count = son - 5 And filtre.Range.Columns.Count = 11 Then Exit Sub
        For i = 1 To application.Min(11, filtre.Filters.Count)
            Set f = filtre.Filters(i)
            If f.On Then
                etkin(i) = True: c1(i) = f.Criteria1: op(i) = f.Operator
                On Error Resume Next
                Err.Clear
                c2(i) = f.Criteria2
                ikinci(i) = (Err.Number = 0)
                On Error GoTo 0
            End If
        Next i
        ws.AutoFilterMode = False
    End If
    ws.Range("A6:K" & son).AutoFilter
    For i = 1 To 11
        If etkin(i) Then
            If ikinci(i) Then
                ws.AutoFilter.Range.AutoFilter Field:=i, Criteria1:=c1(i), Operator:=op(i), Criteria2:=c2(i)
            ElseIf op(i) <> 0 Then
                ws.AutoFilter.Range.AutoFilter Field:=i, Criteria1:=c1(i), Operator:=op(i)
            Else
                ws.AutoFilter.Range.AutoFilter Field:=i, Criteria1:=c1(i)
            End If
        End If
    Next i
End Sub

Private Sub M_DEPO_KONUM(ByVal ws As Worksheet, ByVal liste As Shape)
    Dim kart As Shape, baslik As Shape
    Set kart = ws.Shapes("AKT1_DEPO_KART_ZEMIN")
    Set baslik = ws.Shapes("AKT1_DEPO_KART_BASLIK")
    With liste
        .Left = baslik.Left
        .Top = kart.Top + kart.Height * 0.52
        .Width = kart.Left + kart.Width - .Left - 8
        .Height = application.Min(17, kart.Height * 0.34)
        .Placement = xlFreeFloating
        .ZOrder msoBringToFront
    End With
End Sub

Private Sub M_BOS_RAPOR(ByVal ws As Worksheet)
    Dim kaynak As Worksheet, s As Shape, i As Long, ad As String, metin As String
    Dim d As Object, veri As Variant, v As Variant, keys As Variant, sonuc() As Variant
    Dim son As Long, r As Long, kayit As Long, eksik As Long, dep As String, toplam As Variant
    Dim fiyatMetni As String, tarihMetni As String, names() As Variant, vals() As Variant, seri As Series, negatif As Boolean, fiyatBaslik As String
    Select Case ws.Name
        Case "M_CDP_RAPOR": Set kaynak = ThisWorkbook.Worksheets("M_CDP_STOK")
        Case "M_ÝLÇE_RAPOR": Set kaynak = ThisWorkbook.Worksheets("M_ÝLÇE_STOK")
        Case "M_YÜKLENÝCÝ_RAPOR": Set kaynak = ThisWorkbook.Worksheets("M_YÜKLENÝCÝ_STOK")
    End Select
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    son = kaynak.Cells(kaynak.Rows.Count, "E").End(xlUp).Row
    toplam = CDec(0)
    If son >= 7 Then
        veri = kaynak.Range("A7:N" & son).Value2
        For r = 1 To UBound(veri, 1)
            If Not IsError(veri(r, 5)) Then
                If Len(Trim$(CStr(veri(r, 5)))) > 0 Then
                    kayit = kayit + 1: dep = Trim$(CStr(veri(r, 2)))
                    If Not d.Exists(dep) Then d(dep) = Array(CStr(veri(r, 3)), 0&, CDec(0), CDec(0), 0&)
                    v = d(dep): v(1) = v(1) + 1
                    If IsNumeric(veri(r, 14)) Then
                        If CDbl(veri(r, 14)) > 0 Then v(4) = v(4) + 1: eksik = eksik + 1
                    End If
                    If Not IsError(veri(r, 11)) Then
                        If IsNumeric(veri(r, 11)) And Len(CStr(veri(r, 11))) > 0 Then
                            If Left$(CStr(veri(r, 1)), 2) = "70" Then
                                v(3) = v(3) + CDec(veri(r, 11))
                            Else
                                v(2) = v(2) + CDec(veri(r, 11))
                            End If
                            toplam = toplam + CDec(veri(r, 11))
                        End If
                    End If
                    d(dep) = v
                End If
            End If
        Next r
    End If
    tarihMetni = "-"
    fiyatBaslik = kaynak.Shapes("AKT1_STOK_KPI_ETIKET_2").TextFrame2.TextRange.Text
    If IsDate(kaynak.Range("C2").Value) Then tarihMetni = Format$(kaynak.Range("C2").Value, "dd.mm.yyyy hh:nn:ss")
    fiyatMetni = AKSORA_SAYI_METNI(toplam, "#,##0.00") & " TL"
    If eksik > 0 Then fiyatMetni = "Fiyat eksik (" & CStr(eksik) & ")"
    ws.Range("A9:E" & application.Max(9, ws.Cells(ws.Rows.Count, "A").End(xlUp).Row)).ClearContents
    If d.Count > 0 Then
        keys = d.keys
        ReDim sonuc(1 To d.Count, 1 To 5): ReDim names(1 To d.Count): ReDim vals(1 To d.Count)
        For i = 1 To d.Count
            v = d(keys(i - 1))
            sonuc(i, 1) = CStr(keys(i - 1)): sonuc(i, 2) = v(0)
            If ws.Name = "M_CDP_RAPOR" Then
                sonuc(i, 3) = v(1): sonuc(i, 4) = v(2) + v(3)
                If v(4) > 0 Then sonuc(i, 4) = "Fiyat eksik"
            Else
                sonuc(i, 3) = v(2): sonuc(i, 4) = v(3): sonuc(i, 5) = v(2) + v(3)
                If v(4) > 0 Then sonuc(i, 3) = "Fiyat eksik": sonuc(i, 4) = "Fiyat eksik": sonuc(i, 5) = "Fiyat eksik"
            End If
            names(i) = CStr(v(0)): vals(i) = CDbl(v(2) + v(3))
            If CDbl(vals(i)) < 0 Then negatif = True
        Next i
        ws.Range("A9").Resize(d.Count, 5).Value2 = sonuc
        ws.Range("C9:E" & d.Count + 8).NumberFormat = "#,##0.00"
        If ws.Name = "M_CDP_RAPOR" Then ws.Range("C9:C" & d.Count + 8).NumberFormat = "#,##0"
    End If
    For Each s In ws.Shapes
        ad = s.Name
        If InStr(ad, "_DEGER_A3_E6") > 0 Then s.TextFrame2.TextRange.Text = tarihMetni
        If InStr(ad, "_DEGER_F3_J6") > 0 Then s.TextFrame2.TextRange.Text = AKSORA_SAYI_METNI(kayit, "#,##0")
        If InStr(ad, "_DEGER_K3_Q6") > 0 Then s.TextFrame2.TextRange.Text = fiyatMetni
        If InStr(ad, "_ETIKET_K3_Q6") > 0 Then s.TextFrame2.TextRange.Text = UCase$(fiyatBaslik)
        If ad = "AKSORA_RP_HALKA_MERKEZ" Then s.TextFrame2.TextRange.Text = "TOPLAM STOK" & vbLf & fiyatMetni
        If Left$(ad, 21) = "AKSORA_RP_HALKA_RENK_" Or Left$(ad, 19) = "AKSORA_RP_HALKA_AD_" Then s.Visible = msoFalse
        If Left$(ad, 20) = "AKSORA_RP_HALKA_PAY_" And ad <> "AKSORA_RP_HALKA_PAY_BASLIK" Then s.Visible = msoFalse
    Next s
    For i = 1 To ws.ChartObjects.Count
        With ws.ChartObjects(i).Chart
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            If d.Count > 0 And eksik = 0 And toplam > 0 And Not negatif Then
                .ChartType = xlDoughnut
                Set seri = .SeriesCollection.NewSeries
                seri.XValues = names: seri.Values = vals
                .HasLegend = False
            End If
        End With
    Next i
    If d.Count > 0 And eksik = 0 And toplam > 0 And Not negatif Then
        For Each s In ws.Shapes
            ad = s.Name: i = Val(Mid$(ad, InStrRev(ad, "_") + 1))
            If i >= 1 And i <= d.Count Then
                If Left$(ad, 19) = "AKSORA_RP_HALKA_AD_" Then s.TextFrame2.TextRange.Text = CStr(names(i)): s.Visible = msoTrue
                If Left$(ad, 20) = "AKSORA_RP_HALKA_PAY_" Then s.TextFrame2.TextRange.Text = AKSORA_SAYI_METNI(CDbl(vals(i)) / CDbl(toplam), "0.0%"): s.Visible = msoTrue
                If Left$(ad, 21) = "AKSORA_RP_HALKA_RENK_" Then
                    s.Visible = msoTrue
                    If Not seri Is Nothing Then seri.Points(i).Format.Fill.ForeColor.RGB = s.Fill.ForeColor.RGB
                End If
            End If
        Next s
    End If
End Sub

Public Sub AKSORA_MUSTERI_DEPO_SEC()
    Dim ws As Worksheet, s As Shape, depo As String
    On Error GoTo Hata
    Set ws = ActiveSheet
    If Not M_STOK(ws.Name) Then Exit Sub
    Set s = ws.Shapes(CStr(application.Caller))
    depo = CStr(s.ControlFormat.List(s.ControlFormat.Value))
    ws.Unprotect ""
    If depo = "TÜMÜ" Then
        ws.AutoFilter.Range.AutoFilter Field:=2
    Else
        depo = Replace(Replace(Replace(depo, "~", "~~"), "*", "~*"), "?", "~?")
        ws.AutoFilter.Range.AutoFilter Field:=2, Criteria1:="=" & depo
    End If
    ws.Calculate
Cikis:
    If Not ws Is Nothing Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Public Sub AKSORA_MUSTERI_YEREL_ARA()
    Dim ws As Worksheet, kaynak As Worksheet, ad As Variant, veri As Variant, buffer() As Variant, sonuc() As Variant
    Dim q As String, r As Long, c As Long, n As Long, son As Long, eski As Long, satir As Long, kod As String
    Dim ev As Boolean, sc As Boolean, hataMetni As String
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("M_MALZEME_ARA")
    q = Trim$(CStr(ws.Range("B6").Value2))
    If Len(q) < 2 Then MsgBox "En az iki karakter giriniz.", vbInformation, "AKSORA": Exit Sub
    application.EnableEvents = False: application.ScreenUpdating = False
    ReDim buffer(1 To 14, 1 To 500)
    kod = CStr(ThisWorkbook.Worksheets("MUSTERI IL SECIM").Range("AZ1").Value2)
    For Each ad In Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
        Set kaynak = ThisWorkbook.Worksheets(CStr(ad))
        son = kaynak.Cells(kaynak.Rows.Count, "E").End(xlUp).Row
        If son >= 7 Then
            veri = kaynak.Range("A7:N" & son).Value2
            For r = 1 To UBound(veri, 1)
                If M_ESLESIR(veri(r, 5), q) Or M_ESLESIR(veri(r, 6), q) Or M_ESLESIR(veri(r, 10), q) Then
                    n = n + 1
                    If n > UBound(buffer, 2) Then ReDim Preserve buffer(1 To 14, 1 To UBound(buffer, 2) + 500)
                    For c = 1 To 11: buffer(c, n) = veri(r, c): Next c
                    buffer(12, n) = Mid$(CStr(ad), 3)
                    buffer(14, n) = veri(r, 14)
                    Select Case kod
                        Case "GIRESUN": buffer(13, n) = "GÝRESUN"
                        Case "TRABZON": buffer(13, n) = "TRABZON"
                        Case "RIZE": buffer(13, n) = "RÝZE"
                        Case "GUMUSHANE": buffer(13, n) = "GÜMÜÞHANE"
                        Case "ARTVIN": buffer(13, n) = "ARTVÝN"
                    End Select
                End If
            Next r
        End If
    Next ad
    ws.Unprotect ""
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    eski = application.Max(13, ws.Cells(ws.Rows.Count, "E").End(xlUp).Row)
    ws.Range("A13:N" & eski).ClearContents
    If n > 0 Then
        ReDim sonuc(1 To n, 1 To 14)
        For r = 1 To n
            For c = 1 To 14: sonuc(r, c) = buffer(c, r): Next c
        Next r
        ws.Range("A13").Resize(n, 14).Value2 = sonuc
        AKSORA_V13_MIKTAR_BICIMI ws.Range("G13:G" & CStr(n + 12))
        ws.Range("K13:K" & CStr(n + 12)).NumberFormat = "#,##0.00"
    End If
    ws.Columns("N").Hidden = True
    AKSORA_MUSTERI_ARAMA_KPI_YENILE
    If n = 0 Then MsgBox "Müþteri stoklarýnda eþleþen kayýt bulunamadý.", vbInformation, "AKSORA"
Cikis:
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev: application.ScreenUpdating = sc
    On Error GoTo 0
    If Len(hataMetni) > 0 Then MsgBox hataMetni, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Cikis
End Sub

Private Function M_ESLESIR(ByVal v As Variant, ByVal q As String) As Boolean
    If IsError(v) Or IsEmpty(v) Then Exit Function
    M_ESLESIR = (InStr(1, CStr(v), q, vbTextCompare) > 0)
End Function
