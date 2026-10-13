Attribute VB_Name = "modAKSORA_MUSTERI_ISLEMLER"
Option Explicit
Private mHazir As Boolean
#If VBA7 Then
Private Declare PtrSafe Function M_PencereKilit Lib "user32" Alias "LockWindowUpdate" (ByVal hwnd As LongPtr) As Long
#Else
Private Declare Function M_PencereKilit Lib "user32" Alias "LockWindowUpdate" (ByVal hwnd As Long) As Long
#End If

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
    Dim s As Shape, ev As Boolean, koruma As Boolean, hataNo As Long, hataMetni As String, imza As String
    If mHazir Or Not M_SAYFA(ws.Name) Then Exit Sub
    ' Veri degismediyse sayfa yeniden kurulmaz (gecisler aninda olur).
    imza = M_IMZA(ws)
    If CStr(ws.Range("BX11").Value2) = imza Then Exit Sub
    mHazir = True: ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    koruma = ws.ProtectContents
    ws.Unprotect ""
    On Error Resume Next
    ws.Shapes("MUSTERI_TEMA_GERI").Delete
    On Error GoTo Hata
    If M_STOK(ws.Name) Then M_STOK_KUR ws
    If ws.Name = "M_MALZEME_ARA" Then M_ARAMA_TARIHLERI ws
    If ws.Name = "M_CDP_RAPOR" Or ws.Name = "M_ÝLÇE_RAPOR" Or ws.Name = "M_YÜKLENÝCÝ_RAPOR" Then M_BOS_RAPOR ws
    For Each s In ws.Shapes
        M_BAGLA ws, s
    Next s
    AKSORA_MUSTERI_SAP_BAGLA ws
    If ws.Name = "M_ANA PANEL" Then M_PANEL_KUR ws, False
    ws.Range("BX11").Value2 = imza
Cikis:
    On Error Resume Next
    If koruma Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev: mHazir = False
    On Error GoTo 0
    If hataNo <> 0 Then Err.Raise hataNo, "AKSORA_MUSTERI_SAYFA_HAZIRLA", hataMetni
    If ws.Name = "M_ANA PANEL" Then AKSORA_MUSTERI_SERIT AKSORA_MUSTERI_OZET(), 0
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
    Dim ws As Worksheet, hedef As Worksheet, ad As String, kilit As Boolean
    On Error GoTo Hata
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    Set ws = ActiveSheet
    If Not M_SAYFA(ws.Name) Then Exit Sub
    ad = ws.Shapes(CStr(application.Caller)).AlternativeText
    If Not M_SAYFA(ad) Then Err.Raise vbObjectError + 1800, , "Müþteri sayfa baðlantýsý geçersiz."
    Set hedef = ThisWorkbook.Worksheets(ad)
    ' Gecis boyunca pencere cizimi kilitlenir; yeni sayfa tek seferde ve hazir halde gorunur.
    kilit = (M_PencereKilit(application.hwnd) <> 0)
    application.ScreenUpdating = False
    application.EnableEvents = False
    hedef.Visible = xlSheetVisible
    hedef.Select Replace:=True
    application.EnableEvents = True
    AKSORA_MUSTERI_TEMA_GORUNUM hedef
    AKSORA_MUSTERI_SECIM_BASA hedef
    application.ScreenUpdating = True
    If kilit Then M_PencereKilit 0
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    If kilit Then M_PencereKilit 0
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub M_STOK_KUR(ByVal ws As Worksheet)
    Dim baslik As Variant, agirlik As Variant, i As Long, son As Long, toplam As Double, w As Double
    Dim s As Shape, ad As String, eski As Shape, liste As Shape, x As Double, y As Double, h As Double, bw As Double
    Dim d As Object, k As Variant, r As Long, depo As String, secim As String, p As Long
    Dim imza As String, eksik As Long, veri As Variant, adlar() As String
    baslik = Array("Üretim Yeri", "Depo", "Depo Adý", "Mal Grubu", "Malzeme", _
        "Malzeme Tanýmý", "Miktar", "Birim", "Deðerleme Türü", "Seri Numarasý")
    For i = ws.Shapes.Count To 1 Step -1
        If LCase$(Left$(ws.Shapes(i).Name, 9)) = "drop down" Then ws.Shapes(i).Delete
    Next i
    If CStr(ws.Range("BX8").Value2) <> "MUSTERI_STOK_V2" Then
        If Len(CStr(ws.Range("BX12").Value2)) > 0 And IsNumeric(ws.Range("BX12").Value2) Then _
            ws.Columns(10).ColumnWidth = CDbl(ws.Range("BX12").Value2)
        w = ws.Range("A1:L1").Width
        agirlik = Array(10, 8, 16, 12, 13, 44, 11, 8, 15, 21)
        For i = 0 To 9: toplam = toplam + CDbl(agirlik(i)): Next i
        ws.Columns("K:L").Hidden = False
        For i = 0 To 9: ws.Columns(i + 1).ColumnWidth = w / 7 * CDbl(agirlik(i)) / toplam: Next i
        bw = ws.Range("A1:J1").Width
        For i = 1 To 10: ws.Columns(i).ColumnWidth = ws.Columns(i).ColumnWidth * w / bw: Next i
        ws.Rows("3:5").Hidden = True
        ws.Columns("K:L").Hidden = True
        ws.Range("BX8").Value2 = "MUSTERI_STOK_V2"
        ws.Range("BX12:BX13").ClearContents
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
    son = application.Max(7, AKSORA_M_SON(ws))
    M_FILTRE_ALANI ws, son
    ws.Range("C3").Formula = "=SUBTOTAL(103,E7:E" & son & ")"
    ws.Range("C4").Formula = "=SUBTOTAL(109,K7:K" & son & ")"
    ws.Range("C3").NumberFormat = "#,##0"
    ws.Range("C4").NumberFormat = "#,##0.00 ""TL"""
    imza = CStr(son) & "|" & CStr(ws.Range("C2").Value2) & "|TABLO_V3"
    If CStr(ws.Range("BX9").Value2) <> imza Then
        AKSORA_V13_MIKTAR_BICIMI ws.Range("G7:G" & son)
        M_TABLO_BICIM ws, 7, son, 10, "3,6", "7"
        ws.Range("K7:L" & son).NumberFormat = "#,##0.00 ""TL"""
        ws.Range("BX9").Value2 = imza
    End If
    AKSORA_MUSTERI_BANT ws
    eksik = 0
    veri = ws.Range("N7:N" & son).Value2
    If IsArray(veri) Then
        For r = 1 To UBound(veri, 1)
            If Not IsError(veri(r, 1)) Then
                If IsNumeric(veri(r, 1)) And Len(CStr(veri(r, 1))) > 0 Then
                    If CDbl(veri(r, 1)) > 0 Then eksik = eksik + 1
                End If
            End If
        Next r
    End If
    For Each s In ws.Shapes
        ad = s.Name
        If InStr(ad, "STOK_KPI_DEGER_") > 0 Then
            i = Val(Mid$(ad, InStrRev(ad, "_") + 1))
            If i >= 0 And i <= 2 Then s.DrawingObject.Formula = _
                "='" & Replace(ws.Name, "'", "''") & "'!C" & CStr(i + 2)
        End If
        If InStr(ad, "STOK_KPI_ETIKET_2") > 0 Then
            If eksik > 0 Then
                s.TextFrame2.TextRange.Text = "Toplam Stok Deðeri  •  " & AKSORA_SAYI_METNI(eksik, "#,##0") & " fiyatsýz"
            Else
                s.TextFrame2.TextRange.Text = "Toplam Stok Deðeri"
            End If
        End If
        If InStr(ad, "DEPO_KART_LISTE") > 0 Then Set eski = s
    Next s
    M_BUTON_DUZEN ws
    If eski Is Nothing Then Exit Sub
    ad = eski.Name
    If eski.Type <> msoFormControl Then
        x = eski.Left: y = eski.Top: bw = eski.Width: h = eski.Height
        eski.Delete
        Set liste = ws.Shapes.AddFormControl(xlDropDown, x, y, bw, h)
        liste.Name = ad: liste.Placement = xlFreeFloating
    Else
        Set liste = eski
    End If
    ' Depo listesi depo adlariyla gosterilir; secim Depo/Fikti sutununa filtre uygular.
    secim = vbNullString
    On Error Resume Next
    If ws.AutoFilterMode Then
        If ws.AutoFilter.Filters(2).On Then ws.AutoFilter.Range.AutoFilter Field:=2
        If ws.AutoFilter.Filters(3).On Then secim = Mid$(CStr(ws.AutoFilter.Filters(3).Criteria1), 2)
    End If
    On Error GoTo 0
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    veri = ws.Range("B7:C" & son).Value2
    If IsArray(veri) Then
        For r = 1 To UBound(veri, 1)
            If Not IsError(veri(r, 2)) Then
                depo = Trim$(CStr(veri(r, 2)))
                If Len(depo) > 0 Then d(depo) = True
            End If
        Next r
    End If
    If d.Count > 0 Then
        ReDim adlar(0 To d.Count - 1)
        i = 0
        For Each k In d.keys: adlar(i) = CStr(k): i = i + 1: Next k
        M_SIRALA adlar
    End If
    With liste.ControlFormat
        .ListFillRange = "": .LinkedCell = "": .RemoveAllItems
        .AddItem "TÜMÜ"
        If d.Count > 0 Then
            For i = 0 To UBound(adlar): .AddItem adlar(i): Next i
        End If
        p = 1
        For i = 1 To .ListCount
            If StrComp(CStr(.List(i)), secim, vbTextCompare) = 0 Then p = i
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
    Dim kaynak As Worksheet, s As Shape, a As Shape, i As Long, j As Long, ad As String
    Dim d As Object, veri As Variant, v As Variant, keys As Variant, sonuc() As Variant
    Dim son As Long, r As Long, kayit As Long, eksik As Long, seriAdet As Long, dep As String
    Dim toplam As Double, tarihMetni As String, degerMetni As String, baslik As String
    Dim n As Long, m As Long, slot As Long, kol As Long, degerli As Boolean, negatif As Boolean
    Dim sira() As Long, pay() As Double, adlar As Variant, oranlar As Variant, toplamPay As Double
    Dim t As Long, seri As Series, ilk As Long, eskiSon As Long, tablo As Range, x As Double
    Dim renkler As Variant
    Select Case ws.Name
        Case "M_CDP_RAPOR": Set kaynak = ThisWorkbook.Worksheets("M_CDP_STOK")
        Case "M_ÝLÇE_RAPOR": Set kaynak = ThisWorkbook.Worksheets("M_ÝLÇE_STOK")
        Case "M_YÜKLENÝCÝ_RAPOR": Set kaynak = ThisWorkbook.Worksheets("M_YÜKLENÝCÝ_STOK")
        Case Else: Exit Sub
    End Select
    kol = 5
    ' Depo bazli ozet: ad, satir, deger, fiyatsiz satir, seri numarali satir
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    son = AKSORA_M_SON(kaynak)
    If son >= 7 Then
        veri = kaynak.Range("A7:N" & son).Value2
        For r = 1 To UBound(veri, 1)
            If Not IsError(veri(r, 5)) Then
                If Len(Trim$(CStr(veri(r, 5)))) > 0 Then
                    kayit = kayit + 1
                    dep = vbNullString: If Not IsError(veri(r, 2)) Then dep = Trim$(CStr(veri(r, 2)))
                    If Not d.Exists(dep) Then
                        If IsError(veri(r, 3)) Then
                            d(dep) = Array(vbNullString, 0&, 0#, 0&, 0&)
                        Else
                            d(dep) = Array(Trim$(CStr(veri(r, 3))), 0&, 0#, 0&, 0&)
                        End If
                    End If
                    v = d(dep): v(1) = v(1) + 1
                    If Not IsError(veri(r, 14)) Then
                        If IsNumeric(veri(r, 14)) And Len(CStr(veri(r, 14))) > 0 Then
                            If CDbl(veri(r, 14)) > 0 Then v(3) = v(3) + 1: eksik = eksik + 1
                        End If
                    End If
                    If Not IsError(veri(r, 11)) Then
                        If VarType(veri(r, 11)) <> vbString And IsNumeric(veri(r, 11)) Then
                            v(2) = v(2) + CDbl(veri(r, 11)): toplam = toplam + CDbl(veri(r, 11))
                        End If
                    End If
                    If Not IsError(veri(r, 10)) Then
                        If Len(Trim$(CStr(veri(r, 10)))) > 0 Then v(4) = v(4) + 1: seriAdet = seriAdet + 1
                    End If
                    d(dep) = v
                End If
            End If
        Next r
    End If
    n = d.Count
    tarihMetni = "-"
    If IsDate(kaynak.Range("C2").Value) Then tarihMetni = Format$(kaynak.Range("C2").Value, "dd.mm.yyyy hh:nn:ss")
    degerMetni = M_TL(toplam)
    ' Siralama: degere (fiyat yoksa satir sayisina) gore buyukten kucuge
    degerli = (toplam > 0)
    If n > 0 Then
        keys = d.keys
        ReDim sira(1 To n): ReDim pay(1 To n)
        For i = 1 To n
            v = d(keys(i - 1)): sira(i) = i
            If CDbl(v(2)) < 0 Then negatif = True
        Next i
        If negatif Then degerli = False
        For i = 1 To n
            v = d(keys(i - 1))
            If degerli Then pay(i) = CDbl(v(2)) Else pay(i) = CDbl(v(1))
        Next i
        For i = 2 To n
            t = sira(i): j = i - 1
            Do While j >= 1
                If pay(sira(j)) >= pay(t) Then Exit Do
                sira(j + 1) = sira(j): j = j - 1
            Loop
            sira(j + 1) = t
        Next i
    End If
    ' Tablo: veri nerede bitiyorsa tablo orada biter (TOPLAM satiri dahil)
    ilk = 9
    eskiSon = application.Max(60, ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 5)
    With ws.Range(ws.Cells(ilk, 1), ws.Cells(eskiSon, 5))
        .ClearContents: .ClearFormats
    End With
    ws.Range(ws.Rows(ilk), ws.Rows(eskiSon)).RowHeight = ws.StandardHeight
    ' Uc raporda ayni tablo: ayni kolonlar, ayni genislik, ayni baslik bicimi.
    ws.Columns(1).ColumnWidth = 12.78: ws.Columns(2).ColumnWidth = 28.78
    ws.Range("C1:E1").EntireColumn.ColumnWidth = 20.78
    ws.Range("A8:E8").Value2 = Array("Depo", "Depo Adý", "Stok Satýrý", "Seri Numaralý", "Toplam Stok Deðeri")
    With ws.Range("A8:E8")
        .Interior.Color = RGB(26, 75, 120): .Font.Color = vbWhite
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = False: .Borders.LineStyle = xlNone
    End With
    ReDim sonuc(1 To n + 1, 1 To kol)
    For i = 1 To n
        v = d(keys(sira(i) - 1))
        sonuc(i, 1) = CStr(keys(sira(i) - 1)): sonuc(i, 2) = v(0): sonuc(i, 3) = v(1)
        If kol = 4 Then
            sonuc(i, 4) = v(2)
        Else
            sonuc(i, 4) = v(4): sonuc(i, 5) = v(2)
        End If
    Next i
    sonuc(n + 1, 1) = "TOPLAM": sonuc(n + 1, 3) = kayit
    If kol = 4 Then
        sonuc(n + 1, 4) = toplam
    Else
        sonuc(n + 1, 4) = seriAdet: sonuc(n + 1, 5) = toplam
    End If
    Set tablo = ws.Range(ws.Cells(ilk, 1), ws.Cells(ilk + n, kol))
    tablo.Columns(1).NumberFormat = "@": tablo.Columns(2).NumberFormat = "@"
    tablo.Value2 = sonuc
    With tablo
        .Font.Name = "Aptos": .Font.Size = 10: .Font.Color = RGB(30, 41, 59)
        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlCenter
        .RowHeight = 20.25
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(214, 224, 236)
        End With
    End With
    For i = 1 To n
        If i Mod 2 = 0 Then tablo.Rows(i).Interior.Color = RGB(244, 248, 252)
    Next i
    tablo.Columns(2).HorizontalAlignment = xlLeft: tablo.Columns(2).IndentLevel = 1
    For j = 3 To kol
        tablo.Columns(j).HorizontalAlignment = xlRight: tablo.Columns(j).IndentLevel = 1
        tablo.Columns(j).NumberFormat = "#,##0"
    Next j
    tablo.Columns(kol).NumberFormat = "#,##0.00 ""TL"""
    With tablo.Rows(n + 1)
        .Font.bold = True: .Interior.Color = RGB(226, 236, 247)
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(26, 75, 120)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(26, 75, 120)
        End With
    End With
    ' Kartlar ve basliklar
    baslik = M_METIN(M_SEKIL(ws, "AKT1_RAPOR_BASLIK"))
    If Left$(baslik, 9) = "AKSORA | " Then baslik = Mid$(baslik, 10)
    For Each s In ws.Shapes
        ad = s.Name
        If InStr(ad, "_ETIKET_A3_E6") > 0 Then s.TextFrame2.TextRange.Text = "SON GÜNCELLEME"
        If InStr(ad, "_DEGER_A3_E6") > 0 Then s.TextFrame2.TextRange.Text = tarihMetni
        If InStr(ad, "_ETIKET_F3_J6") > 0 Then s.TextFrame2.TextRange.Text = "TOPLAM STOK SATIRI"
        If InStr(ad, "_DEGER_F3_J6") > 0 Then s.TextFrame2.TextRange.Text = M_ADET(kayit)
        If InStr(ad, "_ETIKET_K3_Q6") > 0 Then
            If eksik > 0 Then
                s.TextFrame2.TextRange.Text = "TOPLAM STOK DEÐERÝ  •  " & M_ADET(eksik) & " FÝYATSIZ SATIR"
            Else
                s.TextFrame2.TextRange.Text = "TOPLAM STOK DEÐERÝ"
            End If
        End If
        If InStr(ad, "_DEGER_K3_Q6") > 0 Then s.TextFrame2.TextRange.Text = degerMetni
        If ad = "AKSORA_RP_HALKA_BASLIK" Then s.TextFrame2.TextRange.Text = "DEPOLARA GÖRE STOK DAÐILIMI"
        If ad = "AKSORA_RP_HALKA_PAY_BASLIK" Then s.TextFrame2.TextRange.Text = IIf(degerli, "PAY", "PAY (SATIR)")
        If ad = "AKSORA_RP_HALKA_MERKEZ" Then
            If degerli Then
                s.TextFrame2.TextRange.Text = "TOPLAM STOK" & vbLf & degerMetni
            Else
                s.TextFrame2.TextRange.Text = "TOPLAM STOK" & vbLf & M_ADET(kayit) & " satýr"
            End If
        End If
        If Left$(ad, 19) = "AKSORA_RP_HALKA_AD_" Then slot = slot + 1
        If Left$(ad, 21) = "AKSORA_RP_HALKA_RENK_" Or Left$(ad, 19) = "AKSORA_RP_HALKA_AD_" Then s.Visible = msoFalse
        If Left$(ad, 20) = "AKSORA_RP_HALKA_PAY_" And ad <> "AKSORA_RP_HALKA_PAY_BASLIK" Then s.Visible = msoFalse
    Next s
    Set s = M_SEKIL(ws, "AKSORA_RP_ALT")
    If Not s Is Nothing Then
        s.TextFrame2.TextRange.Text = "AKSORA  |  " & baslik & "  |  Veriler SAP'den alýnýr."
        s.Left = tablo.Left: s.Top = tablo.Top + tablo.Height + 4
        s.Width = tablo.Width
        On Error Resume Next
        s.TextFrame2.WordWrap = msoFalse: s.TextFrame2.MarginLeft = 2
        s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft
        On Error GoTo 0
    End If
    ' Halka grafik: her durumda cizilir (fiyat eksikse satir sayisina gore)
    If slot = 0 Then slot = 10
    m = n: If m > slot Then m = slot
    If m > 0 Then
        ReDim adlar(1 To m): ReDim oranlar(1 To m)
        For i = 1 To m
            v = d(keys(sira(i) - 1))
            adlar(i) = IIf(Len(CStr(v(0))) > 0, CStr(v(0)), CStr(keys(sira(i) - 1)))
            oranlar(i) = pay(sira(i))
        Next i
        If n > m Then
            adlar(m) = "Diðer": oranlar(m) = 0
            For i = m To n: oranlar(m) = oranlar(m) + pay(sira(i)): Next i
        End If
        For i = 1 To m: toplamPay = toplamPay + oranlar(i): Next i
    End If
    renkler = Array(RGB(30, 99, 168), RGB(0, 150, 136), RGB(245, 158, 11), RGB(124, 58, 237), RGB(239, 68, 68), _
        RGB(14, 165, 233), RGB(34, 197, 94), RGB(236, 72, 153), RGB(100, 116, 139), RGB(202, 138, 4))
    For i = 1 To ws.ChartObjects.Count
        With ws.ChartObjects(i).Chart
            Do While .SeriesCollection.Count > 0
                .SeriesCollection(1).Delete
            Loop
            If m > 0 And toplamPay > 0 Then
                .ChartType = xlDoughnut
                Set seri = .SeriesCollection.NewSeries
                seri.XValues = adlar: seri.Values = oranlar
                .HasLegend = False: .HasTitle = False
                On Error Resume Next
                .ChartGroups(1).DoughnutHoleSize = 62
                seri.Format.line.Visible = msoTrue
                seri.Format.line.ForeColor.RGB = RGB(255, 255, 255)
                For j = 1 To m
                    Set a = Nothing
                    Set a = ws.Shapes("AKSORA_RP_HALKA_RENK_" & j)
                    If a Is Nothing Then
                        seri.Points(j).Format.Fill.ForeColor.RGB = renkler((j - 1) Mod 10)
                    Else
                        seri.Points(j).Format.Fill.ForeColor.RGB = a.Fill.ForeColor.RGB
                    End If
                Next j
                On Error GoTo 0
            End If
        End With
    Next i
    If m > 0 And toplamPay > 0 Then
        On Error Resume Next
        Dim payBas As Shape
        Set payBas = ws.Shapes("AKSORA_RP_HALKA_PAY_BASLIK")
        For j = 1 To m
            Set a = Nothing: Set s = Nothing
            Set a = ws.Shapes("AKSORA_RP_HALKA_AD_" & j)
            If Not a Is Nothing Then a.TextFrame2.TextRange.Text = CStr(adlar(j)): a.Visible = msoTrue
            Set s = ws.Shapes("AKSORA_RP_HALKA_PAY_" & j)
            If Not s Is Nothing Then
                s.TextFrame2.TextRange.Text = AKSORA_SAYI_METNI(CDbl(oranlar(j)) / toplamPay, "0.0%")
                s.Visible = msoTrue
                If Not a Is Nothing Then s.Top = a.Top
                If Not payBas Is Nothing Then
                    s.Left = payBas.Left: s.Width = payBas.Width
                    s.TextFrame2.MarginLeft = payBas.TextFrame2.MarginLeft
                    s.TextFrame2.MarginRight = payBas.TextFrame2.MarginRight
                    s.TextFrame2.TextRange.ParagraphFormat.Alignment = payBas.TextFrame2.TextRange.ParagraphFormat.Alignment
                End If
            End If
            Set s = Nothing
            Set s = ws.Shapes("AKSORA_RP_HALKA_RENK_" & j)
            If Not s Is Nothing Then
                s.Visible = msoTrue
                If Not a Is Nothing Then s.Top = a.Top + (a.Height - s.Height) / 2
            End If
        Next j
        On Error GoTo 0
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
    If Not ws.AutoFilterMode Then M_FILTRE_ALANI ws, application.Max(7, AKSORA_M_SON(ws))
    If ws.AutoFilter.Filters(2).On Then ws.AutoFilter.Range.AutoFilter Field:=2
    If depo = "TÜMÜ" Then
        ws.AutoFilter.Range.AutoFilter Field:=3
    Else
        depo = Replace(Replace(Replace(depo, "~", "~~"), "*", "~*"), "?", "~?")
        ws.AutoFilter.Range.AutoFilter Field:=3, Criteria1:="=" & depo
    End If
    ws.Calculate
    AKFIX_EN_USTE ws
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
        son = AKSORA_M_SON(kaynak)
        If son >= 7 Then
            veri = kaynak.Range("A7:N" & son).Value2
            For r = 1 To UBound(veri, 1)
                If M_ESLESIR(veri(r, 5), q) Or M_ESLESIR(veri(r, 6), q) Or M_ESLESIR(veri(r, 10), q) Then
                    n = n + 1
                    If n > UBound(buffer, 2) Then ReDim Preserve buffer(1 To 14, 1 To UBound(buffer, 2) + 500)
                    For c = 1 To 11: buffer(c, n) = veri(r, c): Next c
                    Select Case CStr(ad)
                        Case "M_CDP_STOK": buffer(12, n) = "CDP"
                        Case "M_ÝLÇE_STOK": buffer(12, n) = "ÝLÇE"
                        Case Else: buffer(12, n) = "YÜKLENÝCÝ"
                    End Select
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
    eski = application.Max(13, AKSORA_M_SON(ws))
    ws.Range("A13:N" & eski).ClearContents
    If n > 0 Then
        ReDim sonuc(1 To n, 1 To 14)
        For r = 1 To n
            For c = 1 To 14: sonuc(r, c) = buffer(c, r): Next c
        Next r
        ws.Range("A13").Resize(n, 14).Value2 = sonuc
    End If
    M_TABLO_BICIM ws, 13, n + 12, 13, "3,6", "7,11"
    If n > 0 Then
        AKSORA_V13_MIKTAR_BICIMI ws.Range("G13:G" & CStr(n + 12))
        ws.Range("K13:K" & CStr(n + 12)).NumberFormat = "#,##0.00 ""TL"""
    End If
    ws.Columns("N").Hidden = True
    AKSORA_MUSTERI_BANT ws
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

'==========================================================
' AKSORA - MUSTERI ORTAK TABLO / PANEL / RAPOR YARDIMCILARI
'==========================================================

Private Function M_SEKIL(ByVal ws As Worksheet, ByVal ad As String) As Shape
    On Error Resume Next
    Set M_SEKIL = ws.Shapes(ad)
    On Error GoTo 0
End Function

Private Sub M_YAZ(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String)
    Dim s As Shape
    Set s = M_SEKIL(ws, ad)
    If s Is Nothing Then Exit Sub
    On Error Resume Next
    s.TextFrame2.TextRange.Text = metin
    On Error GoTo 0
End Sub

Private Function M_TL(ByVal deger As Double) As String
    M_TL = AKSORA_SAYI_METNI(deger, "#,##0.00") & " TL"
End Function

Private Function M_ADET(ByVal deger As Double) As String
    M_ADET = AKSORA_SAYI_METNI(deger, "#,##0")
End Function

' Tablo satir nerede bitiyorsa bicim de orada biter; tum satirlar ayni yukseklik ve yazi tipindedir.
Private Sub M_TABLO_BICIM(ByVal ws As Worksheet, ByVal ilk As Long, ByVal son As Long, ByVal sonKol As Long, _
    ByVal metinKol As String, ByVal sayiKol As String)
    Dim eskiSon As Long, c As Variant, alan As Range
    eskiSon = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If eskiSon < ilk + 200 Then eskiSon = ilk + 200
    If eskiSon > son Then
        ws.Range(ws.Cells(application.Max(ilk, son + 1), 1), ws.Cells(eskiSon, sonKol)).ClearFormats
        If ws.Name = "M_MALZEME_ARA" Then
            ws.Range(ws.Rows(ilk), ws.Rows(eskiSon)).RowHeight = 22.05
        Else
            ws.Range(ws.Rows(application.Max(ilk, son + 1)), ws.Rows(eskiSon)).RowHeight = ws.StandardHeight
        End If
    End If
    If son < ilk Then Exit Sub
    Set alan = ws.Range(ws.Cells(ilk, 1), ws.Cells(son, sonKol))
    With alan
        .Interior.Pattern = xlNone
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = False: .Font.Italic = False
        .Font.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .WrapText = False: .ShrinkToFit = False: .IndentLevel = 0
        .Borders.LineStyle = xlNone
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(214, 224, 236)
        End With
        If son > ilk Then
            With .Borders(xlInsideHorizontal)
                .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(214, 224, 236)
            End With
        End If
    End With
    ws.Range(ws.Rows(ilk), ws.Rows(son)).RowHeight = IIf(ws.Name = "M_MALZEME_ARA", 22.05, 20.25)
    M_BANT_KUR ws, alan
    If Len(metinKol) > 0 Then
        For Each c In Split(metinKol, ",")
            With ws.Range(ws.Cells(ilk, CLng(c)), ws.Cells(son, CLng(c)))
                .HorizontalAlignment = xlLeft: .IndentLevel = 1
            End With
        Next c
    End If
    If Len(sayiKol) > 0 Then
        For Each c In Split(sayiKol, ",")
            With ws.Range(ws.Cells(ilk, CLng(c)), ws.Cells(son, CLng(c)))
                .HorizontalAlignment = xlRight: .IndentLevel = 1
            End With
        Next c
    End If
End Sub

Private Sub M_SIRALA(ByRef a() As String)
    Dim i As Long, j As Long, t As String
    For i = LBound(a) + 1 To UBound(a)
        t = a(i): j = i - 1
        Do While j >= LBound(a)
            If StrComp(a(j), t, vbTextCompare) <= 0 Then Exit Do
            a(j + 1) = a(j): j = j - 1
        Loop
        a(j + 1) = t
    Next i
End Sub

' Musteri stok sayfalarinda tek SAP guncelleme dugmesi kalir.
Private Sub M_BUTON_DUZEN(ByVal ws As Worksheet)
    Dim kalan As Shape, giden As Shape
    Select Case ws.Name
        Case "M_ÝLÇE_STOK"
            Set kalan = M_SEKIL(ws, "CDP_ILCE_STOK_BTN"): Set giden = M_SEKIL(ws, "EPP_ILCE_STOK_BTN")
        Case "M_YÜKLENÝCÝ_STOK"
            Set kalan = M_SEKIL(ws, "YUK_EDAS_GUNCELLE_BTN"): Set giden = M_SEKIL(ws, "YUK_AKSADA_GUNCELLE_BTN")
        Case "M_CDP_STOK"
            Set kalan = M_SEKIL(ws, "SAP_STOK_GUNCELLE_BTN")
    End Select
    If kalan Is Nothing Then Exit Sub
    If Not giden Is Nothing Then
        kalan.Left = giden.Left: kalan.Width = giden.Width
        giden.Delete
    End If
    On Error Resume Next
    kalan.TextFrame2.TextRange.Text = "SAP STOKLARINI GÜNCELLE"
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' DURUM SERIDI (M_ANA PANEL)  durum: 0 bilgi, 1 calisiyor, 2 tamam, 3 hata
'----------------------------------------------------------
Public Sub AKSORA_MUSTERI_SERIT(ByVal metin As String, ByVal durum As Long)
    Dim ws As Worksheet, s As Shape, korumali As Boolean, zemin As Long, yazi As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("M_ANA PANEL")
    If ws Is Nothing Then Exit Sub
    Set s = ws.Shapes("MUSTERI_ANA_SECIM_SERIDI")
    If s Is Nothing Then Exit Sub
    Select Case durum
        Case 1: zemin = RGB(255, 246, 219): yazi = RGB(146, 64, 14)
        Case 2: zemin = RGB(209, 250, 229): yazi = RGB(6, 95, 70)
        Case 3: zemin = RGB(254, 226, 226): yazi = RGB(153, 27, 27)
        Case Else: zemin = RGB(224, 242, 254): yazi = RGB(12, 45, 79)
    End Select
    korumali = ws.ProtectDrawingObjects
    If korumali Then ws.Unprotect ""
    s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = zemin
    s.line.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 16: .MarginRight = 220: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
        .TextRange.Text = metin
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 10
        .TextRange.Font.bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = yazi
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With
    If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    DoEvents
    On Error GoTo 0
End Sub

Private Sub M_KAYNAK_OZET(ByVal ad As String, ByRef satir As Long, ByRef deger As Double, _
    ByRef eksik As Long, ByRef depo As Long, ByRef tarih As Date)
    Dim sh As Worksheet, son As Long, v As Variant, r As Long, d As Object
    satir = 0: deger = 0: eksik = 0: depo = 0: tarih = 0
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub
    If IsDate(sh.Range("C2").Value) Then tarih = CDate(sh.Range("C2").Value)
    son = AKSORA_M_SON(sh)
    If son < 7 Then Exit Sub
    v = sh.Range("A7:N" & son).Value2
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = 1 To UBound(v, 1)
        If Not IsError(v(r, 5)) Then
            If Len(Trim$(CStr(v(r, 5)))) > 0 Then
                satir = satir + 1
                If Not IsError(v(r, 2)) Then d(Trim$(CStr(v(r, 2)))) = True
                If Not IsError(v(r, 11)) Then
                    If VarType(v(r, 11)) <> vbString And IsNumeric(v(r, 11)) Then deger = deger + CDbl(v(r, 11))
                End If
                If Not IsError(v(r, 14)) Then
                    If IsNumeric(v(r, 14)) And Len(CStr(v(r, 14))) > 0 Then
                        If CDbl(v(r, 14)) > 0 Then eksik = eksik + 1
                    End If
                End If
            End If
        End If
    Next r
    depo = d.Count
End Sub

' Son tam guncellemenin suresi (saniye); kayit yoksa 0.
Private Function M_SON_SURE() As Double
    Dim t As String
    On Error Resume Next
    t = ThisWorkbook.names("_AKSORA_MUSTERI_7155_SON_SURE").RefersTo
    On Error GoTo 0
    If Len(t) > 1 Then M_SON_SURE = Val(Mid$(t, 2))
End Function

Public Function AKSORA_MUSTERI_OZET() As String
    Dim adlar As Variant, kisa As Variant, i As Long, satir As Long, deger As Double, eksik As Long
    Dim depo As Long, tarih As Date, enSon As Date, toplam As Double, topEksik As Long, metin As String
    Dim sure As Double
    adlar = Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
    kisa = Array("CDP", "Ýlçe", "Yüklenici")
    For i = 0 To 2
        M_KAYNAK_OZET CStr(adlar(i)), satir, deger, eksik, depo, tarih
        If tarih > enSon Then enSon = tarih
        toplam = toplam + deger: topEksik = topEksik + eksik
        metin = metin & "   |   " & kisa(i) & " " & M_ADET(satir) & " satýr"
    Next i
    If enSon > 0 Then
        sure = M_SON_SURE()
        If sure > 0 Then metin = "   |   Güncelleme süresi " & AKSORA_SAYI_METNI(sure, "0.00") & " sn" & metin
        metin = "Son güncelleme " & Format$(enSon, "dd.mm.yyyy hh:nn") & metin
    Else
        metin = "Henüz SAP güncellemesi yapýlmadý" & metin
    End If
    metin = metin & "   |   Toplam stok deðeri " & M_TL(toplam)
    If topEksik > 0 Then metin = metin & "   |   Fiyatsýz " & M_ADET(topEksik) & " satýr"
    AKSORA_MUSTERI_OZET = metin
End Function

Private Sub M_KART_DURUM(ByVal ws As Worksheet, ByVal k As String, ByVal metin As String, ByVal renk As Long)
    Dim s As Shape, koyu As Long
    Select Case renk
        Case RGB(34, 197, 94): koyu = RGB(21, 128, 61)
        Case RGB(245, 158, 11): koyu = RGB(180, 83, 9)
        Case RGB(239, 68, 68): koyu = RGB(185, 28, 28)
        Case Else: koyu = RGB(29, 78, 216)
    End Select
    M_YAZ ws, "AKSORA15_KART_DEGER_" & k & "0", metin
    On Error Resume Next
    Set s = ws.Shapes("AKSORA15_KART_DEGER_" & k & "0")
    If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = koyu
    Set s = Nothing
    Set s = ws.Shapes("AKSORA15_KART_LED_" & k)
    If Not s Is Nothing Then
        s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = renk
        s.Glow.Color.RGB = renk: s.Glow.Radius = 4: s.Glow.Transparency = 0.5
    End If
    On Error GoTo 0
End Sub

Private Sub M_PANEL_KUR(ByVal ws As Worksheet, ByVal calisiyor As Boolean)
    Dim kartlar As Variant, adlar As Variant, i As Long, k As String
    Dim satir As Long, deger As Double, eksik As Long, depo As Long, tarih As Date
    Dim e3 As Shape, e4 As Shape, d3 As Shape, d4 As Shape, dg As Shape, sr As Shape, s As Shape
    Dim sap As String, h As Double, makro As String
    kartlar = Array("B", "L", "Q")
    adlar = Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
    sap = CStr(ws.Range("BX10").Value2)
    For i = 0 To 2
        k = CStr(kartlar(i))
        M_KAYNAK_OZET CStr(adlar(i)), satir, deger, eksik, depo, tarih
        ' Son satir: CDP'de SAP baglantisi, diger kartlarda depo sayisi.
        Set e3 = M_SEKIL(ws, "AKSORA15_KART_ETIKET_" & k & "3")
        Set e4 = M_SEKIL(ws, "AKSORA15_KART_ETIKET_" & k & "4")
        Set d3 = M_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "3")
        Set d4 = M_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "4")
        If Not e3 Is Nothing And Not e4 Is Nothing And Not d3 Is Nothing Then
            On Error Resume Next
            If d4 Is Nothing Then
                Set d4 = d3.Duplicate
                d4.Name = "AKSORA15_KART_DEGER_" & k & "4"
            End If
            e4.Left = e3.Left: e4.Width = e3.Width
            e4.TextFrame2.TextRange.ParagraphFormat.Alignment = e3.TextFrame2.TextRange.ParagraphFormat.Alignment
            d4.Left = d3.Left: d4.Top = e4.Top + (d3.Top - e3.Top)
            d4.Visible = msoTrue
            On Error GoTo 0
        End If
        M_YAZ ws, "AKSORA15_KART_DEGER_" & k & "1", IIf(tarih > 0, Format$(tarih, "dd.mm.yyyy hh:nn"), "-")
        M_YAZ ws, "AKSORA15_KART_DEGER_" & k & "2", IIf(satir > 0, M_ADET(satir), "-")
        M_YAZ ws, "AKSORA15_KART_DEGER_" & k & "3", IIf(satir > 0, M_TL(deger), "-")
        M_YAZ ws, "AKSORA15_KART_ETIKET_" & k & "4", "DEPO SAYISI"
        M_YAZ ws, "AKSORA15_KART_DEGER_" & k & "4", IIf(depo > 0, M_ADET(depo), "-")
        If calisiyor Then
            M_KART_DURUM ws, k, "Güncelleniyor...", RGB(59, 130, 246)
        ElseIf satir = 0 Or tarih = 0 Then
            M_KART_DURUM ws, k, "Veri yok", RGB(239, 68, 68)
        ElseIf Int(tarih) >= Date Then
            M_KART_DURUM ws, k, "GÜNCEL", RGB(34, 197, 94)
        ElseIf Date - Int(tarih) = 1 Then
            M_KART_DURUM ws, k, "1 gündür güncellenmedi", RGB(245, 158, 11)
        Else
            M_KART_DURUM ws, k, CStr(Date - Int(tarih)) & " gündür güncellenmedi", RGB(239, 68, 68)
        End If
    Next i
    ' Islem dugmeleri
    makro = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    Set s = M_SEKIL(ws, "btnTumStok")
    If Not s Is Nothing Then s.OnAction = makro & "AKSORA_MUSTERI_7155_GUNCELLE"
    Set s = M_SEKIL(ws, "btnBaglantiKontrol")
    If Not s Is Nothing Then s.OnAction = makro & "AKSORA_MUSTERI_SAP_KONTROL"
    ' Il degistir: seridin sagina oturan teknolojik hap dugme
    Set dg = M_SEKIL(ws, "AKSORA15_IL_DEGISTIR")
    Set sr = M_SEKIL(ws, "MUSTERI_ANA_SECIM_SERIDI")
    If Not dg Is Nothing And Not sr Is Nothing Then
        On Error Resume Next
        dg.AutoShapeType = msoShapeRoundedRectangle
        dg.Adjustments.Item(1) = 0.5
        h = sr.Height - 8
        If h < 16 Then h = sr.Height
        dg.Height = h: dg.Width = 196
        dg.Top = sr.Top + (sr.Height - h) / 2
        dg.Left = sr.Left + sr.Width - dg.Width - 10
        dg.Fill.Visible = msoTrue
        dg.Fill.TwoColorGradient msoGradientVertical, 1
        dg.Fill.ForeColor.RGB = RGB(12, 45, 79)
        dg.Fill.BackColor.RGB = RGB(0, 150, 136)
        dg.line.Visible = msoTrue: dg.line.ForeColor.RGB = RGB(94, 234, 212): dg.line.Weight = 1
        dg.Glow.Color.RGB = RGB(45, 212, 191): dg.Glow.Radius = 5: dg.Glow.Transparency = 0.55
        dg.Shadow.Visible = msoFalse
        With dg.TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoFalse
            .TextRange.Text = ChrW(8644) & "   ÝL DEÐÝÞTÝR"
            .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 10
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
            .TextRange.Font.Spacing = 0.8
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
        dg.ZOrder msoBringToFront
        On Error GoTo 0
    End If
End Sub

Public Sub AKSORA_MUSTERI_PANEL_YENILE(Optional ByVal calisiyor As Boolean = False)
    Dim ws As Worksheet, korumali As Boolean, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("M_ANA PANEL")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    korumali = ws.ProtectContents Or ws.ProtectDrawingObjects
    If korumali Then ws.Unprotect ""
    M_PANEL_KUR ws, calisiyor
Cikis:
    On Error Resume Next
    If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    application.EnableEvents = ev
    On Error GoTo 0
    If Not calisiyor Then AKSORA_MUSTERI_SERIT AKSORA_MUSTERI_OZET(), 0
    Exit Sub
Hata:
    Resume Cikis
End Sub

Public Sub AKSORA_MUSTERI_SAP_KONTROL()
    Dim ses As Object, mesaj As String, ws As Worksheet, bilgi As String, korumali As Boolean
    On Error GoTo Hata
    AKSORA_MUSTERI_SERIT "SAP baðlantýsý kontrol ediliyor...", 1
    Set ses = SAP_SESSION_AL("CDP", mesaj)
    Set ws = ThisWorkbook.Worksheets("M_ANA PANEL")
    korumali = ws.ProtectContents
    If korumali Then ws.Unprotect ""
    If ses Is Nothing Then
        ws.Range("BX10").Value2 = "YOK|" & Format$(Now, "dd.mm.yyyy hh:nn")
    Else
        ws.Range("BX10").Value2 = "BAGLI|" & Format$(Now, "dd.mm.yyyy hh:nn")
        On Error Resume Next
        bilgi = vbCrLf & vbCrLf & "Sistem: " & ses.Info.SystemName & "  |  Ýstemci: " & ses.Info.Client & _
            "  |  Kullanýcý: " & ses.Info.User
        On Error GoTo Hata
    End If
    If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True
    AKSORA_MUSTERI_PANEL_YENILE
    If ses Is Nothing Then
        AKSORA_MUSTERI_SERIT "SAP baðlantýsý bulunamadý. SAP Logon'dan CDP sistemine giriþ yapýn.", 3
        MsgBox "SAP baðlantýsý bulunamadý." & vbCrLf & vbCrLf & mesaj, vbExclamation, "AKSORA"
    Else
        AKSORA_MUSTERI_SERIT "SAP baðlantýsý hazýr" & Replace(Replace(bilgi, vbCrLf & vbCrLf, "  •  "), vbCrLf, ""), 2
        MsgBox "SAP baðlantýsý hazýr." & bilgi, vbInformation, "AKSORA"
    End If
    Exit Sub
Hata:
    AKSORA_MUSTERI_SERIT "SAP baðlantý kontrolü tamamlanamadý.", 3
    AKSORA_HATA_GOSTER "SAP baðlantý kontrolü", Err.Number, Err.Description
End Sub

Public Sub AKSORA_MUSTERI_ARAMA_TEMIZLE()
    Dim ws As Worksheet, eski As Long, ev As Boolean, sc As Boolean
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("M_MALZEME_ARA")
    application.EnableEvents = False: application.ScreenUpdating = False
    ws.Unprotect ""
    If ws.FilterMode Then ws.ShowAllData
    eski = application.Max(13, AKSORA_M_SON(ws))
    ws.Range("A13:N" & eski).ClearContents
    ws.Range("B6").MergeArea.ClearContents
    M_TABLO_BICIM ws, 13, 12, 13, "", ""
    ws.Range(ws.Cells(13, 75), ws.Cells(application.Max(13, eski), 75)).ClearContents
    AKSORA_MUSTERI_ARAMA_KPI_YENILE
    application.EnableEvents = ev: application.ScreenUpdating = sc
    Exit Sub
Hata:
    application.EnableEvents = ev: application.ScreenUpdating = sc
    AKSORA_HATA_GOSTER "Arama temizleme", Err.Number, Err.Description
End Sub

'----------------------------------------------------------
' BEYAZ / MAVI SATIR DESENI (filtreye uyumlu)
' Gorunen satirlar BW yardimci sutununda 0/1 ile isaretlenir;
' kosullu bicim =INDEX($BW:$BW;SATIR())=1 mavi zemini uygular.
'----------------------------------------------------------
Private Function M_YEREL_FORMUL(ByVal ws As Worksheet, ByVal formul As String) As String
    Dim c As Range
    Set c = ws.Range("BW1")
    On Error Resume Next
    c.Formula = formul
    M_YEREL_FORMUL = c.FormulaLocal
    c.ClearContents
    On Error GoTo 0
    If Len(M_YEREL_FORMUL) = 0 Then M_YEREL_FORMUL = formul
End Function

Private Sub M_BANT_KUR(ByVal ws As Worksheet, ByVal alan As Range)
    Dim fc As Object, yerel As String
    alan.FormatConditions.Delete
    yerel = M_YEREL_FORMUL(ws, "=INDEX($BW:$BW,ROW())=1")
    On Error Resume Next
    Set fc = alan.FormatConditions.Add(Type:=xlExpression, Formula1:=yerel)
    If fc Is Nothing Then
        Err.Clear
        Set fc = alan.FormatConditions.Add(Type:=xlExpression, Formula1:="=INDEX($BW:$BW,ROW())=1")
    End If
    If Not fc Is Nothing Then
        fc.Interior.Color = RGB(232, 241, 251)
        fc.StopIfTrue = False
    End If
    On Error GoTo 0
    ws.Columns("BW").Hidden = True
End Sub

Public Sub AKSORA_MUSTERI_BANT(ByVal ws As Worksheet)
    Static mesgul As Boolean
    Dim ilk As Long, son As Long, v() As Variant, eski As Variant, alan As Range, ar As Range
    Dim r As Long, n As Long, korumali As Boolean, ev As Boolean, farkli As Boolean
    If mesgul Then Exit Sub
    If AKSORA_V11_MESGUL Then Exit Sub
    Select Case ws.Name
        Case "M_MALZEME_ARA": ilk = 13
        Case "M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK": ilk = 7
        Case Else: Exit Sub
    End Select
    son = AKSORA_M_SON(ws)
    If son < ilk Then Exit Sub
    mesgul = True: ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    korumali = ws.ProtectContents
    If korumali Then ws.Unprotect ""
    ReDim v(1 To son - ilk + 1, 1 To 1)
    For r = 1 To son - ilk + 1: v(r, 1) = 0: Next r
    On Error Resume Next
    Set alan = ws.Range(ws.Cells(ilk, 5), ws.Cells(son, 5)).SpecialCells(xlCellTypeVisible)
    On Error GoTo Hata
    If Not alan Is Nothing Then
        For Each ar In alan.Areas
            For r = ar.Row To ar.Row + ar.Rows.Count - 1
                n = n + 1
                If n Mod 2 = 0 Then v(r - ilk + 1, 1) = 1
            Next r
        Next ar
    End If
    eski = ws.Range(ws.Cells(ilk, 75), ws.Cells(son, 75)).Value2
    If Not IsArray(eski) Then
        farkli = True
    Else
        For r = 1 To son - ilk + 1
            If CStr(eski(r, 1)) <> CStr(v(r, 1)) Then farkli = True: Exit For
        Next r
    End If
    If farkli Then ws.Range(ws.Cells(ilk, 75), ws.Cells(son, 75)).Value2 = v
Cikis:
    On Error Resume Next
    If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev: mesgul = False
    Exit Sub
Hata:
    Resume Cikis
End Sub

'----------------------------------------------------------
' Son veri satiri: filtreyle gizlenen satirlar da sayilir.
'----------------------------------------------------------
Public Function AKSORA_M_SON(ByVal ws As Worksheet) As Long
    Dim c As Range
    On Error Resume Next
    Set c = ws.Columns(5).Find(What:="*", LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0
    If c Is Nothing Then AKSORA_M_SON = 1 Else AKSORA_M_SON = c.Row
End Function

Private Function M_KAYNAK_IMZA(ByVal ad As String) As String
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(ad)
    If sh Is Nothing Then Exit Function
    M_KAYNAK_IMZA = "|" & CStr(AKSORA_M_SON(sh)) & "|" & CStr(sh.Range("C2").Value2)
    On Error GoTo 0
End Function

Private Function M_IMZA(ByVal ws As Worksheet) As String
    Dim t As String, wsF As Worksheet
    t = "IMZA_V2|" & ws.Name
    Select Case ws.Name
        Case "M_CDP_STOK", "M_CDP_RAPOR": t = t & M_KAYNAK_IMZA("M_CDP_STOK")
        Case "M_ÝLÇE_STOK", "M_ÝLÇE_RAPOR": t = t & M_KAYNAK_IMZA("M_ÝLÇE_STOK")
        Case "M_YÜKLENÝCÝ_STOK", "M_YÜKLENÝCÝ_RAPOR": t = t & M_KAYNAK_IMZA("M_YÜKLENÝCÝ_STOK")
        Case Else
            t = t & M_KAYNAK_IMZA("M_CDP_STOK") & M_KAYNAK_IMZA("M_ÝLÇE_STOK") & M_KAYNAK_IMZA("M_YÜKLENÝCÝ_STOK")
            t = t & "|" & CStr(CLng(Date)) & "|" & CStr(ws.Range("BX10").Value2)
    End Select
    On Error Resume Next
    Set wsF = ThisWorkbook.Worksheets("_FIYAT_LISTESI")
    If Not wsF Is Nothing Then t = t & "|" & CStr(wsF.Range("J4").Value2)
    On Error GoTo 0
    M_IMZA = t
End Function

' Stok sayfasindan cikinca filtreler ve depo secimi ilk haline doner.
Public Sub AKSORA_MUSTERI_FILTRE_SIFIRLA(ByVal haric As String)
    Dim ad As Variant, sh As Worksheet, s As Shape, korumali As Boolean
    On Error Resume Next
    For Each ad In Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK")
        If CStr(ad) <> haric Then
            Set sh = Nothing
            Set sh = ThisWorkbook.Worksheets(CStr(ad))
            If Not sh Is Nothing Then
                If sh.FilterMode Then
                    korumali = sh.ProtectContents
                    If korumali Then sh.Unprotect ""
                    sh.ShowAllData
                    For Each s In sh.Shapes
                        If InStr(s.Name, "DEPO_KART_LISTE") > 0 Then s.ControlFormat.Value = 1
                    Next s
                    If korumali Then sh.Protect Password:="", DrawingObjects:=True, Contents:=True, _
                        UserInterfaceOnly:=True, AllowFiltering:=True
                    AKSORA_MUSTERI_BANT sh
                End If
            End If
        End If
    Next ad
    On Error GoTo 0
End Sub

'----------------------------------------------------------
' EKRAN DUZENI: icerik pencere genisligine tam oturur; serit, kart,
' tablo ve grafikler ayni sol/sag cizgide biter. Pencere boyutu veya
' veri degismedikce tekrar calismaz (sayfa gecisi aninda olur).
'----------------------------------------------------------
Private Function M_SON_KOLON(ByVal ad As String) As Long
    Select Case ad
        Case "M_MALZEME_ARA": M_SON_KOLON = 16
        Case "M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK": M_SON_KOLON = 10
        Case "M_CDP_RAPOR", "M_ÝLÇE_RAPOR", "M_YÜKLENÝCÝ_RAPOR": M_SON_KOLON = 21
    End Select
End Function

Private Function M_DUZEN_IMZA(ByVal ws As Worksheet, ByVal pn As Window) As String
    M_DUZEN_IMZA = "DUZEN_V2|" & Format$(pn.UsableWidth, "0.0") & "|" & Format$(pn.UsableHeight, "0.0") & "|" & _
        CStr(pn.Zoom) & "|" & CStr(ws.Range("BX11").Value2)
End Function

Public Sub AKSORA_MUSTERI_DUZEN(ByVal ws As Worksheet)
    Dim pn As Window, sonKol As Long, kol As Range, w As Double, z As Double, ek As Double
    Dim alt As Double, s As Shape, korumali As Boolean, ev As Boolean
    sonKol = M_SON_KOLON(ws.Name)
    If sonKol = 0 Then Exit Sub
    Set pn = ThisWorkbook.Windows(1)
    If Not (pn.ActiveSheet Is ws) Then Exit Sub
    If CStr(ws.Range("BX13").Value2) = M_DUZEN_IMZA(ws, pn) Then Exit Sub
    ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    korumali = ws.ProtectContents Or ws.ProtectDrawingObjects
    If korumali Then ws.Unprotect ""
    Set kol = ws.Columns(sonKol)
    ' Son kolonun asil genisligi saklanir; pencere farki her seferinde bu tabandan hesaplanir.
    If Len(CStr(ws.Range("BX12").Value2)) = 0 Or Not IsNumeric(ws.Range("BX12").Value2) Then _
        ws.Range("BX12").Value2 = kol.ColumnWidth
    If kol.ColumnWidth <> CDbl(ws.Range("BX12").Value2) Then kol.ColumnWidth = CDbl(ws.Range("BX12").Value2)
    w = ws.Range(ws.Cells(1, 1), ws.Cells(1, sonKol)).Width
    If w <= 0 Then GoTo Cikis
    z = (pn.UsableWidth - 1) / w * 100
    z = application.Max(10, application.Min(400, Fix(z)))
    If pn.Zoom <> z Then pn.Zoom = z
    ' Tam sayi yakinlastirmadan kalan bosluk son kolona eklenir: sag kenar pencereyle biter.
    ek = pn.UsableWidth * 100 / z - w - 1
    If ek > 0.5 And kol.Width > 0 Then _
        kol.ColumnWidth = application.Min(255, kol.ColumnWidth * (kol.Width + ek) / kol.Width)
    w = ws.Range(ws.Cells(1, 1), ws.Cells(1, sonKol)).Width
    Select Case ws.Name
        Case "M_MALZEME_ARA": M_DUZEN_ARAMA ws, w
        Case "M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK": M_DUZEN_STOK ws, w
        Case Else: M_DUZEN_RAPOR ws, w
    End Select
    ws.Range("BX13").Value2 = M_DUZEN_IMZA(ws, pn)
Cikis:
    On Error Resume Next
    If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
        UserInterfaceOnly:=True, AllowFiltering:=True
    If ws.Name = "M_MALZEME_ARA" Or M_STOK(ws.Name) Then ws.EnableSelection = xlNoRestrictions
    application.EnableEvents = ev
    On Error GoTo 0
    Exit Sub
Hata:
    Resume Cikis
End Sub

' Sekli yatayda verilen konum ve genislige oturtur (resimlerde oran kilidi acilir).
Private Sub M_YATAY(ByVal ws As Worksheet, ByVal ad As String, ByVal x As Double, ByVal gen As Double)
    Dim s As Shape
    Set s = M_SEKIL(ws, ad)
    If s Is Nothing Then Exit Sub
    On Error Resume Next
    s.LockAspectRatio = msoFalse
    On Error GoTo 0
    s.Left = x: s.Width = gen
End Sub

' Kartlari ic sekilleriyle (ikon, etiket, deger, liste) birlikte yeni yerlerine tasir.
Private Sub M_KART_TASI(ByVal ws As Worksheet, ByVal kartlar As Variant, ByVal sol As Variant, ByVal gen As Variant)
    Dim i As Long, j As Long, s As Shape, k As Shape, cx As Double, cy As Double, ust As Long, ad As String
    Dim uye() As Long, dx() As Double, dw() As Double, adlar() As String
    ust = UBound(kartlar)
    ReDim dx(0 To ust): ReDim dw(0 To ust)
    For i = 0 To ust
        Set k = kartlar(i)
        dx(i) = CDbl(sol(i)) - k.Left: dw(i) = CDbl(gen(i)) - k.Width
    Next i
    ReDim uye(1 To ws.Shapes.Count)
    For Each s In ws.Shapes
        j = j + 1: uye(j) = -1
        cx = s.Left + s.Width / 2: cy = s.Top + s.Height / 2
        For i = 0 To ust
            Set k = kartlar(i)
            If s.Name <> k.Name And s.Width < k.Width And s.Height <= k.Height Then
                If cx >= k.Left And cx <= k.Left + k.Width And cy >= k.Top And cy <= k.Top + k.Height Then
                    uye(j) = i: Exit For
                End If
            End If
        Next i
    Next s
    ' Esnek ogeler (etiket/deger/liste) genisler; ikon parcalari tek grup halinde kayar
    ' (tek tek kaydirilinca piksel yuvarlamasi simgeyi bozar).
    ReDim adlar(0 To ust)
    j = 0
    For Each s In ws.Shapes
        j = j + 1
        If uye(j) >= 0 Then
            i = uye(j): ad = UCase$(s.Name)
            If InStr(ad, "ETIKET") > 0 Or InStr(ad, "DEGER") > 0 Or InStr(ad, "BASLIK") > 0 Or InStr(ad, "LISTE") > 0 Then
                s.Left = s.Left + dx(i): s.Width = application.Max(10, s.Width + dw(i))
            ElseIf s.Type = msoFormControl Then
                s.Left = s.Left + dx(i)
            Else
                adlar(i) = adlar(i) & vbTab & s.Name
            End If
        End If
    Next s
    For i = 0 To ust
        If Len(adlar(i)) > 0 And dx(i) <> 0 Then M_GRUP_KAYDIR ws, Split(Mid$(adlar(i), 2), vbTab), dx(i)
    Next i
    For i = 0 To ust
        Set k = kartlar(i)
        k.Left = CDbl(sol(i)): k.Width = CDbl(gen(i))
    Next i
End Sub

Private Sub M_GRUP_KAYDIR(ByVal ws As Worksheet, ByVal adlar As Variant, ByVal dx As Double)
    Dim g As Shape, i As Long
    If UBound(adlar) = 0 Then
        ws.Shapes(CStr(adlar(0))).Left = ws.Shapes(CStr(adlar(0))).Left + dx
        Exit Sub
    End If
    On Error Resume Next
    Set g = ws.Shapes.Range(adlar).Group
    On Error GoTo 0
    If g Is Nothing Then
        For i = 0 To UBound(adlar)
            ws.Shapes(CStr(adlar(i))).Left = ws.Shapes(CStr(adlar(i))).Left + dx
        Next i
    Else
        g.Left = g.Left + dx
        g.Ungroup
    End If
End Sub

Private Sub M_DUZEN_STOK(ByVal ws As Worksheet, ByVal w As Double)
    Dim a As Shape, b As Shape, ad As Variant, kartlar As Variant, i As Long, cw As Double
    Dim sol(0 To 3) As Double, gen(0 To 3) As Double
    M_YATAY ws, "AKT1_STOK_BASLIK_ZEMIN", 0, w
    M_YATAY ws, "AKT1_STOK_KPI_ALT_ZEMIN", 0, w
    Set a = M_SEKIL(ws, "btnAnaPanelDon")
    If Not a Is Nothing Then
        a.Left = w - 10 - a.Width
        For Each ad In Array("SAP_STOK_GUNCELLE_BTN", "CDP_ILCE_STOK_BTN", "YUK_EDAS_GUNCELLE_BTN")
            Set b = M_SEKIL(ws, CStr(ad))
            If Not b Is Nothing Then b.Left = a.Left - 8 - b.Width
        Next ad
    End If
    kartlar = Array(M_SEKIL(ws, "AKT1_STOK_KPI_ZEMIN_0"), M_SEKIL(ws, "AKT1_STOK_KPI_ZEMIN_1"), _
        M_SEKIL(ws, "AKT1_STOK_KPI_ZEMIN_2"), M_SEKIL(ws, "AKT1_DEPO_KART_ZEMIN"))
    For i = 0 To 3
        If kartlar(i) Is Nothing Then Exit Sub
    Next i
    cw = (w - 16 - 3 * 10) / 4
    For i = 0 To 3
        sol(i) = 8 + i * (cw + 10): gen(i) = cw
    Next i
    M_KART_TASI ws, kartlar, sol, gen
    M_EV_SIMGE ws
End Sub

' Depo kartindaki ev simgesi zemin karesine gore simetrik cizilir (cati govdeye ortali).
Private Sub M_EV_SIMGE(ByVal ws As Worksheet)
    Dim z As Shape, cati As Shape, govde As Shape, r0 As Shape, r1 As Shape, r2 As Shape
    Dim cx As Double, w As Double, h As Double, t As Double, gw As Double, cw As Double, pw As Double
    Set z = M_SEKIL(ws, "AKT1_DEPO_KART_SIMGE_Z")
    Set cati = M_SEKIL(ws, "AKT1_DEPO_KART_SIMGE_CATI")
    Set govde = M_SEKIL(ws, "AKT1_DEPO_KART_SIMGE_GOVDE")
    If z Is Nothing Or cati Is Nothing Or govde Is Nothing Then Exit Sub
    Set r0 = M_SEKIL(ws, "AKT1_DEPO_KART_SIMGE_RAF_0")
    Set r1 = M_SEKIL(ws, "AKT1_DEPO_KART_SIMGE_RAF_1")
    Set r2 = M_SEKIL(ws, "AKT1_DEPO_KART_SIMGE_RAF_2")
    On Error Resume Next
    cati.LockAspectRatio = msoFalse: govde.LockAspectRatio = msoFalse
    cati.Rotation = 0: govde.Rotation = 0
    w = z.Width: h = z.Height: t = z.Top
    cx = Round(z.Left + w / 2, 1)
    cw = Round(w * 0.78 / 1.5, 0) * 1.5: gw = Round(w * 0.48 / 1.5, 0) * 1.5
    cati.Width = cw: cati.Height = Round(h * 0.3, 1)
    cati.Left = cx - cw / 2: cati.Top = t + Round(h * 0.18, 1)
    govde.Width = gw: govde.Height = Round(h * 0.42, 1)
    govde.Left = cx - gw / 2: govde.Top = cati.Top + cati.Height - 1
    pw = Round(w * 0.1 / 0.75, 0) * 0.75
    If Not r0 Is Nothing Then
        r0.Width = pw: r0.Height = Round(h * 0.2, 1)
        r0.Left = cx - pw - Round(w * 0.05, 1): r0.Top = govde.Top + (govde.Height - r0.Height) / 2
        r0.Visible = msoTrue
    End If
    If Not r1 Is Nothing Then
        r1.Width = pw: r1.Height = Round(h * 0.2, 1)
        r1.Left = cx + Round(w * 0.05, 1): r1.Top = govde.Top + (govde.Height - r1.Height) / 2
        r1.Visible = msoTrue
    End If
    If Not r2 Is Nothing Then r2.Visible = msoFalse
    govde.ZOrder msoBringToFront
    cati.ZOrder msoBringToFront
    If Not r0 Is Nothing Then r0.ZOrder msoBringToFront
    If Not r1 Is Nothing Then r1.ZOrder msoBringToFront
    On Error GoTo 0
End Sub

Private Sub M_DUZEN_RAPOR(ByVal ws As Worksheet, ByVal w As Double)
    Dim s As Shape, z As Shape, a As Shape, kartlar(0 To 2) As Object, i As Long
    Dim tw As Double, x2 As Double, sag As Double, g As Double, cw As Double, k As Double
    Dim oL As Double, nL As Double, nW As Double, cx As Double
    Dim sol(0 To 2) As Double, gen(0 To 2) As Double
    tw = ws.Range("A1:E1").Width
    g = 12: sag = w - 2: x2 = tw + g
    M_YATAY ws, "AKT1_RAPOR_UST_ZEMIN", 0, w
    Set a = M_SEKIL(ws, "btnAnaPanelDon")
    If Not a Is Nothing Then a.Left = w - 12 - a.Width
    For Each s In ws.Shapes
        If InStr(s.Name, "_KART_A3_E6") > 0 Then Set kartlar(0) = s
        If InStr(s.Name, "_KART_F3_J6") > 0 Then Set kartlar(1) = s
        If InStr(s.Name, "_KART_K3_Q6") > 0 Then Set kartlar(2) = s
    Next s
    If Not (kartlar(0) Is Nothing Or kartlar(1) Is Nothing Or kartlar(2) Is Nothing) And sag - x2 > 120 Then
        cw = (sag - x2 - g) / 2
        sol(0) = 0.5: gen(0) = tw - 0.5
        sol(1) = x2: gen(1) = cw
        sol(2) = x2 + cw + g: gen(2) = cw
        M_KART_TASI ws, Array(kartlar(0), kartlar(1), kartlar(2)), sol, gen
    End If
    ' Grafik alani tablonun sagindan sayfa kenarina kadar; ic sekiller orantili yerlesir.
    Set z = M_SEKIL(ws, "AKSORA_RP_HALKA_ZEMIN")
    If Not z Is Nothing And sag - x2 > 120 Then
        oL = z.Left: nL = x2: nW = sag - x2
        If z.Width > 0 Then
            k = nW / z.Width
            For Each s In ws.Shapes
                If s.Name <> z.Name And (Left$(s.Name, 16) = "AKSORA_RP_HALKA_" Or s.HasChart = msoTrue) Then
                    If s.HasChart = msoTrue Or InStr(s.Name, "_RENK_") > 0 Or InStr(s.Name, "_MERKEZ") > 0 Or _
                       InStr(s.Name, "_GRAFIK") > 0 Then
                        cx = nL + (s.Left + s.Width / 2 - oL) * k
                        s.Left = cx - s.Width / 2
                    Else
                        s.Left = nL + (s.Left - oL) * k: s.Width = s.Width * k
                    End If
                End If
            Next s
            z.Left = nL: z.Width = nW
        End If
    End If
    For Each s In ws.Shapes
        If InStr(s.Name, "FILIGRAN") > 0 Then
            s.Left = sag - s.Width
            If Not z Is Nothing Then s.Top = z.Top + z.Height + 8
        End If
    Next s
    Set s = M_SEKIL(ws, "AKSORA_RP_ALT")
    If Not s Is Nothing Then s.Left = 0: s.Width = tw
End Sub

Private Sub M_DUZEN_ARAMA(ByVal ws As Worksheet, ByVal w As Double)
    Dim s As Shape, ref As Shape, oL As Double, nL As Double, nW As Double, k As Double
    Dim baslik As Variant, i As Long, ad As String, da As Boolean
    ' Ust baslik tum sayfa boyunca (sag il paneli dahil) uzanir.
    If ws.Range("A1").MergeArea.Columns.Count <> 16 Then
        da = application.DisplayAlerts: application.DisplayAlerts = False
        ws.Range("A1:P3").UnMerge
        ws.Range("N1:P2").Interior.Color = ws.Range("A1").Interior.Color
        ws.Range("N3:P3").Interior.Color = ws.Range("A3").Interior.Color
        ws.Range("A1:P2").Merge: ws.Range("A3:P3").Merge
        application.DisplayAlerts = da
    End If
    ws.Range("A3").Value2 = "CDP + Ýlçe + Yüklenici Stoklarý"
    baslik = Array("Üretim Yeri", "Depo", "Depo Adý", "Mal Grubu", "Malzeme", "Malzeme Tanýmý", "Miktar", _
        "Birim", "Deðerleme Türü", "Seri Numarasý", "Deðer", "Kaynak", "Ýl")
    For i = 0 To 12
        If CStr(ws.Cells(12, i + 1).Value2) <> CStr(baslik(i)) Then ws.Cells(12, i + 1).Value2 = baslik(i)
    Next i
    ws.Range("A12:M12").Interior.Color = RGB(26, 75, 120)
    Set s = M_SEKIL(ws, "btnMalzemeAnaPanel")
    If Not s Is Nothing Then
        s.Left = w - 12 - s.Width
        s.Top = ws.Rows(1).Top + (ws.Rows(3).Top - ws.Rows(1).Top - s.Height) / 2
        s.line.Visible = msoTrue: s.line.ForeColor.RGB = RGB(255, 255, 255): s.line.Weight = 0.75
    End If
    ' Il paneli O:P alanina tam oturur.
    Set ref = M_SEKIL(ws, "AKSORA_V4_TUMU")
    If Not ref Is Nothing Then
        oL = ref.Left
        nL = ws.Columns(15).Left + 10: nW = w - 10 - nL
        If ref.Width > 0 And nW > 40 Then
            k = nW / ref.Width
            For Each s In ws.Shapes
                ad = s.Name
                If Left$(ad, 10) = "AKSORA_V4_" Or ad = "MUSTERI_ARAMA_IL_BASLIK" Or ad = "AKSORA_TOPLU_GUNCELLE" Or _
                   ad = "AKSORA_TOPLU_DURUM" Or ad = "AKSORA_ARAMA_SABIT_ACIKLAMA" Then
                    s.Left = nL + (s.Left - oL) * k: s.Width = s.Width * k
                ElseIf ad = "AKSORA_TOPLU_VURGU" Then
                    s.Left = nL + (s.Left - oL) * k
                End If
            Next s
        End If
    End If
    ws.ScrollArea = "A1:P60000"
End Sub

' Sayfaya girildiginde secim gorunmeyen bir hucreye (veya arama kutusuna) alinir.
Public Sub AKSORA_MUSTERI_SECIM_BASA(ByVal ws As Worksheet)
    Dim ev As Boolean
    If Not (ActiveSheet Is ws) Then Exit Sub
    ev = application.EnableEvents
    On Error Resume Next
    application.EnableEvents = False
    Select Case ws.Name
        Case "M_MALZEME_ARA": ws.Range("B6").Select
        Case "M_ANA PANEL": ws.Range("B3").Select
        Case Else: ws.Range("A1").Select
    End Select
    application.EnableEvents = ev
    On Error GoTo 0
End Sub
