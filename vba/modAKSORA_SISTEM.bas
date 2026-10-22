Attribute VB_Name = "modAKSORA_SISTEM"
Option Explicit

Public AKSORA_V13_ACILIS As Boolean
Public AKSORA_V13_GECIS As Boolean
Private mSureler As Object

Public Function AKSORA_V13_SON(ByVal ws As Worksheet) As Long
    ' Limit cleanup to the occupied/formatted area, including an empty final row.
    AKSORA_V13_SON = application.Min(ws.Rows.Count, _
        application.Max(100, ws.UsedRange.Row + ws.UsedRange.Rows.Count))
End Function

Public Sub AKSORA_V13_ARAMA_OZET(ByVal ws As Worksheet)
    Dim ev As Boolean, son As Long, veri As Variant, gorunen As Range, parca As Range
    Dim i As Long, ilk As Long, bitis As Long, n As Long, depo As Object, birim As Object
    Dim anahtar As String, b As String, k As Variant, miktar As Double, deger As Variant
    Dim aciklama As String, hataNo As Long, hataMetni As String
    ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    Set depo = CreateObject("Scripting.Dictionary")
    Set birim = CreateObject("Scripting.Dictionary")
    deger = CDec(0)
    son = AKSORA_SON_SATIR(ws, "D")
    If son >= 13 Then
        veri = ws.Range("A13:K" & son).Value2
        Set gorunen = AKSORA_V13_GORUNEN_SATIRLAR(ws, son)
        If Not gorunen Is Nothing Then
            For Each parca In gorunen.Areas
                ilk = parca.Row - 12: bitis = ilk + parca.Rows.Count - 1
                If ilk < 1 Or bitis > UBound(veri, 1) Then _
                    Err.Raise vbObjectError + 1314, , "Arama özeti sonuç alanýnýn dýþýna çýktý."
                For i = ilk To bitis
                    If Not IsError(veri(i, 4)) Then
                        If Len(Trim$(CStr(veri(i, 4)))) > 0 Then
                            n = n + 1
                            anahtar = CStr(veri(i, 11)) & "|" & CStr(veri(i, 1)) & "|" & CStr(veri(i, 2))
                            depo(anahtar) = True
                            b = UCase$(Trim$(CStr(veri(i, 8))))
                            If Len(b) = 0 Then b = "BELÝRSÝZ"
                            If IsNumeric(veri(i, 7)) Then
                                miktar = CDbl(veri(i, 7))
                                If Not birim.Exists(b) Then birim(b) = 0#
                                birim(b) = birim(b) + miktar
                            End If
                            If IsNumeric(veri(i, 9)) Then deger = deger + CDec(veri(i, 9))
                        End If
                    End If
                Next i
            Next parca
        End If
    End If
    ws.Range("A10").Value2 = n
    ws.Range("C10").Value2 = depo.Count
    ws.Range("I10").Value2 = deger
    ws.Range("I10").NumberFormat = "#,##0.00 ""TL"""
    On Error Resume Next
    ws.Range("F10").Comment.Delete
    On Error GoTo Hata
    If birim.Count > 1 Then
        ws.Range("F10").Value2 = "Birimler farklý"
        For Each k In birim.keys
            aciklama = aciklama & CStr(k) & ": " & AKSORA_SAYI_METNI(birim(k), "#,##0.##########") & vbCrLf
        Next k
        ws.Range("F10").AddComment "Farklý birimler tek miktar olarak toplanmaz." & vbCrLf & aciklama
    Else
        miktar = 0
        For Each k In birim.keys: miktar = birim(k): Next k
        ws.Range("F10").Value2 = miktar
        AKSORA_V13_MIKTAR_BICIMI ws.Range("F10")
    End If
    ws.Range("A10").MergeArea.HorizontalAlignment = xlCenter
    ws.Range("C10").MergeArea.HorizontalAlignment = xlCenter
    ws.Range("F10").MergeArea.HorizontalAlignment = xlCenter
    ws.Range("I10").MergeArea.HorizontalAlignment = xlCenter
Cikis:
    application.EnableEvents = ev
    If hataNo <> 0 Then Err.Raise hataNo, "AKSORA_V13_ARAMA_OZET", hataMetni
    Exit Sub
Hata:
    hataNo = Err.Number: hataMetni = Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_V13_SURE_KAYDET(ByVal islem As String, ByVal baslangic As Double)
    Dim sure As Double
    sure = Timer - baslangic
    If sure < 0 Then sure = sure + 86400#
    If mSureler Is Nothing Then Set mSureler = CreateObject("Scripting.Dictionary")
    mSureler(islem) = sure
    Debug.Print "AKSORA V13 | " & islem & " | " & AKSORA_SAYI_METNI(sure, "0.000") & " sn"
End Sub

Public Sub AKSORA_V13_GIT(ByVal sayfaAdi As String)
    Dim ws As Worksheet, ev As Boolean, hesap As XlCalculation
    Dim imlec As XlMousePointer, hataNo As Long, hataMetni As String, bas As Double
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    bas = Timer
    ev = application.EnableEvents: hesap = application.Calculation
    imlec = application.Cursor
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(sayfaAdi)
    If Left$(ws.Name, 1) = "_" Then Err.Raise vbObjectError + 1301, , "Ýç veri sayfasý hýzlý eriþimle açýlamaz."
    If Not AKSORA_GIRIS_ACIK And ws.Name <> "KAPAK" Then Err.Raise vbObjectError + 1302, , "Önce kapaktan panele giriþ yapýn."
    AKSORA_V13_GECIS = True
    application.EnableEvents = False
    application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    application.Cursor = xlWait
    Sayfa12.STOK_KARS_FILTRE_TAKIP_DURDUR
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    Select Case AKFIX_NORMAL(ws.Name)
        Case "CDP_STOK", "EPP_STOK", "ILCE_STOK", "YUKLENICI_STOK"
            AKFIX_STOK_LISTE ws
            AKFIX_STOK_SECIM ws
        Case "STOKLAR"
            AKFIX_NORMAL_STOK_HAZIRLA ws
        Case "MALZEME_ARA"
            AKSORA_V11_ARAMA_TARIH
        Case "KAPAK"
            AKSORA_KAPAK_KORU
            AKSORA_V8_KAPAK_GORUNUM
    End Select
    AKSORA_V13_GORUNUM ws
    If AKFIX_NORMAL(ws.Name) = "STOK_KARSILASTIRMA" Then AKSORA_V13_KARS_TAKIP_BASLAT
Cikis:
    On Error Resume Next
    application.Calculation = hesap
    application.Cursor = imlec
    AKSORA_V13_GECIS = False
    application.EnableEvents = ev
    ' This is a user-facing entry point; always release the screen on completion.
    application.ScreenUpdating = True
    If hataNo = 0 Then AKSORA_V13_SURE_KAYDET "Geçiþ: " & ws.Name, bas
    On Error GoTo 0
    If hataNo <> 0 Then Err.Raise hataNo, "AKSORA_V13_GIT", hataMetni
    Exit Sub
Hata:
    hataNo = Err.Number: hataMetni = Err.Description
    Resume Cikis
End Sub

Private Sub AKSORA_V13_KARS_TAKIP_BASLAT()
    ' The worksheet activation event is intentionally skipped by controlled navigation.
    Sayfa12.STOK_KARS_V13_TAKIP_AC
End Sub

Public Sub AKSORA_V13_GORUNUM(ByVal ws As Worksheet)
    Dim hucre As String, ad As String
    If Not (ActiveSheet Is ws) Then Exit Sub
    If Left$(ws.Name, 1) = "_" Then Exit Sub
    ad = AKFIX_NORMAL(ws.Name)
    Select Case ad
        Case "KAPAK": Exit Sub
        Case "ANA PANEL": hucre = "A1"
        Case "MALZEME_ARA", "STOK_KARSILASTIRMA": hucre = "A4"
        Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR": hucre = "A7"
        Case "CDP_STOK", "EPP_STOK", "ILCE_STOK", "YUKLENICI_STOK": hucre = "H2"
Case Else: hucre = "M1"
    End Select
    application.CutCopyMode = False
    If ws.EnableSelection <> xlNoSelection Then ws.Range(hucre).Select
    If AKSORA_PENCERE_KUCUK(ThisWorkbook.Windows(1)) Then Exit Sub
    With ThisWorkbook.Windows(1)
        .DisplayGridlines = False
        .DisplayHeadings = False
        .DisplayWorkbookTabs = True
        .DisplayHorizontalScrollBar = True
        .DisplayVerticalScrollBar = True
        .ScrollRow = 1
        .ScrollColumn = 1
    End With
    AKSORA_EKRAN_SIGDIR ws
End Sub

Public Sub AKSORA_V13_MIKTAR_BICIMI(ByVal alan As Range)
    Dim veri As Variant, i As Long, ilk As Long, son As Long
    Dim v As Variant, kesir As Boolean, parcali As Range, parca As Range, adet As Long
    AKSORA_SAYI_AYIRACLARI
    alan.NumberFormat = "#,##0"
    son = alan.Rows.Count
    veri = alan.Value2
    For i = 1 To son + 1
        kesir = False
        If i <= son Then
            If son = 1 Then v = veri Else v = veri(i, 1)
            If Not IsError(v) And Not IsEmpty(v) Then
                If IsNumeric(v) And Len(CStr(v)) > 0 Then kesir = (Round(CDbl(v), 10) <> Fix(Round(CDbl(v), 10)))
            End If
        End If
        If kesir Then
            If ilk = 0 Then ilk = i
        ElseIf ilk > 0 Then
            Set parca = alan.Cells(ilk, 1).Resize(i - ilk, 1)
            If parcali Is Nothing Then Set parcali = parca Else Set parcali = Union(parcali, parca)
            ilk = 0: adet = adet + 1
            If adet = 25 Then
                parcali.NumberFormat = "#,##0.##########"
                Set parcali = Nothing: adet = 0
            End If
        End If
    Next i
    If Not parcali Is Nothing Then parcali.NumberFormat = "#,##0.##########"
End Sub

Public Sub AKSORA_V13_BUTONLARI_BAGLA()
    Dim ws As Worksheet, s As Shape, makro As String, tamAd As String, nokta As Long
    Dim ilNo As Long
    tamAd = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!"
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 1) <> "_" Then
            For Each s In ws.Shapes
                makro = ""
                On Error Resume Next
                makro = s.OnAction
                On Error GoTo 0
                If Len(makro) > 0 Then
                    nokta = InStrRev(makro, "!")
                    If nokta > 0 Then makro = Mid$(makro, nokta + 1)
                    s.OnAction = tamAd & makro
                    s.Placement = xlFreeFloating
                End If
            Next s
        End If
    Next ws
    Set ws = ThisWorkbook.Worksheets("KAPAK")
    ws.Unprotect Password:=""
    For ilNo = 1 To 5
        With ws.Shapes("AKSORA_V4_IL_" & ilNo)
            .OnAction = tamAd & "AKSORA_V8_IL_SEC"
            .Visible = msoTrue
            .ZOrder msoBringToFront
        End With
    Next ilNo
    With ws.Shapes("AKSORA_V4_ENTER")
        .OnAction = tamAd & "AKSORA_V8_GIRIS"
        .Visible = msoTrue
        .ZOrder msoBringToFront
    End With
    AKSORA_KAPAK_KORU
End Sub

Public Sub AKSORA_V13_KORUMALARI_YENILE()
    Dim ws As Worksheet, izin(1 To 11) As Boolean
    Dim cizim As Boolean, icerik As Boolean, senaryo As Boolean, secim As Long
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 1) <> "_" Then
            cizim = ws.ProtectDrawingObjects: icerik = ws.ProtectContents
            senaryo = ws.ProtectScenarios: secim = ws.EnableSelection
            If cizim Or icerik Or senaryo Then
                With ws.Protection
                    izin(1) = .AllowFiltering: izin(2) = .AllowSorting
                    izin(3) = .AllowFormattingCells: izin(4) = .AllowFormattingColumns
                    izin(5) = .AllowFormattingRows: izin(6) = .AllowInsertingColumns
                    izin(7) = .AllowInsertingRows: izin(8) = .AllowInsertingHyperlinks
                    izin(9) = .AllowDeletingColumns: izin(10) = .AllowDeletingRows
                    izin(11) = .AllowUsingPivotTables
                End With
                ws.Unprotect Password:=""
                ws.Protect Password:="", DrawingObjects:=cizim, Contents:=icerik, _
                    Scenarios:=senaryo, UserInterfaceOnly:=True, _
                    AllowFiltering:=izin(1), AllowSorting:=izin(2), _
                    AllowFormattingCells:=izin(3), AllowFormattingColumns:=izin(4), _
                    AllowFormattingRows:=izin(5), AllowInsertingColumns:=izin(6), _
                    AllowInsertingRows:=izin(7), AllowInsertingHyperlinks:=izin(8), _
                    AllowDeletingColumns:=izin(9), AllowDeletingRows:=izin(10), _
                    AllowUsingPivotTables:=izin(11)
                ws.EnableSelection = secim
            End If
        End If
    Next ws
End Sub

Public Sub AKSORA_V13_KUR()
    Dim ev As Boolean, sc As Boolean, hesap As XlCalculation
    Dim Hata As String, ws As Worksheet
    ev = application.EnableEvents: sc = application.ScreenUpdating
    hesap = application.Calculation
    On Error GoTo Hata
    Set ws = ActiveSheet
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    AKSORA_V13_KORUMALARI_YENILE
    AKSORA_V12_GORSEL_ONAR
    AKSORA_V11_BICIMLER
    AKSORA_V13_BUTONLARI_BAGLA
    AKSORA_V13_ARAMA_ZEMIN
    If ws.Name <> "KAPAK" Then AKSORA_V13_GORUNUM ws
Cikis:
    On Error Resume Next
    application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If Len(Hata) > 0 Then
        MsgBox "Kurulum tamamlanamadý: " & Hata, vbExclamation, "AKSORA"
    Else
        MsgBox "V13 kuruldu. Stoklar ve karþýlaþtýrma geçmiþi korunmuþtur.", vbInformation, "AKSORA"
    End If
    Exit Sub
Hata:
    Hata = Err.Description
    Resume Cikis
End Sub

Private Sub AKSORA_V13_ARAMA_ZEMIN()
    Dim ws As Worksheet, s As Shape, alan As Range
    Set ws = ThisWorkbook.Worksheets("MALZEME_ARA")
    With ws.Range("A1").MergeArea.Font
        .Color = vbWhite
        .bold = True
        .Size = 20
    End With
    ws.Range("A1").MergeArea.HorizontalAlignment = xlCenter
    ws.Range("A1").MergeArea.VerticalAlignment = xlCenter
    ws.Range("M20").MergeArea.ClearContents
    On Error Resume Next
    Set s = ws.Shapes("AKSORA_ARAMA_SABIT_ACIKLAMA")
    On Error GoTo 0
    Set alan = ws.Range("M20:N20")
    If s Is Nothing Then Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, alan.Left, alan.Top, alan.Width, alan.Height)
    With s
        .Name = "AKSORA_ARAMA_SABIT_ACIKLAMA"
        .TextFrame.Characters.Text = "Seçili ildeki mevcut stoklar aranýr."
        .TextFrame.Characters.Font.Name = "Aptos"
        .TextFrame.Characters.Font.Size = 8
        .TextFrame.Characters.Font.Color = RGB(85, 103, 123)
        .TextFrame.HorizontalAlignment = xlHAlignCenter
        .TextFrame.VerticalAlignment = xlVAlignCenter
        .TextFrame.MarginTop = 0: .TextFrame.MarginBottom = 0
        .Fill.Visible = msoFalse: .line.Visible = msoFalse
        .Placement = xlFreeFloating
        .Locked = True
    End With
End Sub

Public Function AKSORA_V13_GORUNEN_SATIRLAR(ByVal ws As Worksheet, ByVal son As Long) As Range
    Dim aday As Range, alan As Range
    If son < 13 Then Exit Function
    Set alan = ws.Range("D13:D" & son)
    If son = 13 Then
        If Not ws.Rows(13).Hidden Then Set AKSORA_V13_GORUNEN_SATIRLAR = alan
        Exit Function
    End If
    If Not ws.FilterMode Then
        Set AKSORA_V13_GORUNEN_SATIRLAR = alan
        Exit Function
    End If
    On Error Resume Next
    Set aday = ws.Range("D12:D" & son).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0
    If Not aday Is Nothing Then Set AKSORA_V13_GORUNEN_SATIRLAR = Intersect(aday, alan)
End Function






Public Sub AKSORA_SAYI_AYIRACLARI()
    application.UseSystemSeparators = False
    application.DecimalSeparator = ","
    application.ThousandsSeparator = "."
End Sub

Public Function AKSORA_SAYI_METNI(ByVal deger As Variant, ByVal bicim As String) As String
    Dim metin As String, ondalik As String, binlik As String, ornek As String
    metin = Format$(deger, bicim)
    ondalik = Mid$(Format$(1.5, "0.0"), 2, 1)
    ornek = Format$(1000, "#,##0")
    If Len(ornek) > 4 Then binlik = Mid$(ornek, 2, Len(ornek) - 4)
    If Len(binlik) > 0 Then metin = Replace(metin, binlik, ChrW(1))
    metin = Replace(metin, ondalik, ",")
    metin = Replace(metin, ChrW(1), ".")
    AKSORA_SAYI_METNI = metin
End Function

Public Sub AKSORA_SAYI_BICIM_SAYFA(ByVal ws As Worksheet)
    Dim ad As String, son As Long
    ad = AKFIX_NORMAL(ws.Name)
    If Left$(ad, 2) = "M_" Then ad = Mid$(ad, 3)
    Select Case ad
        Case "CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK"
            son = AKSORA_SON_SATIR(ws, "G")
            If son >= 7 Then AKSORA_V13_MIKTAR_BICIMI ws.Range("I7:I" & son)
        Case "MALZEME_ARA"
            son = AKSORA_SON_SATIR(ws, "G")
            If son >= 13 Then AKSORA_V13_MIKTAR_BICIMI ws.Range("G13:G" & son)
            AKSORA_V13_MIKTAR_BICIMI ws.Range("F10")
            ws.Range("A10").MergeArea.NumberFormat = "#,##0"
            ws.Range("C10").MergeArea.NumberFormat = "#,##0"
    End Select
End Sub

Public Sub AKSORA_SAYI_BICIMLERINI_YENILE()
    Dim ws As Worksheet
    AKSORA_SAYI_AYIRACLARI
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 1) <> "_" And Left$(ws.Name, 2) <> "M_" Then AKSORA_SAYI_BICIM_SAYFA ws
    Next ws
End Sub

' Arama deseni: Turkce harf esdegerleri (i/I/Ý/ý, s/þ, g/ð, u/ü, o/ö, c/ç) ve buyuk/kucuk harf ayni sayilir.
' Kullanim: metin Like AKSORA_ARAMA_DESEN(aranan)  (aranan metnin herhangi bir yerde gecmesi).
Public Function AKSORA_ARAMA_DESEN(ByVal q As String) As String
    Dim i As Long, ch As String, d As String
    For i = 1 To Len(q)
        ch = Mid$(q, i, 1)
        Select Case ch
            Case "i", "I", "Ý", "ý": d = d & "[iIÝý]"
            Case "s", "S", "þ", "Þ": d = d & "[sSþÞ]"
            Case "g", "G", "ð", "Ð": d = d & "[gGðÐ]"
            Case "u", "U", "ü", "Ü": d = d & "[uUüÜ]"
            Case "o", "O", "ö", "Ö": d = d & "[oOöÖ]"
            Case "c", "C", "ç", "Ç": d = d & "[cCçÇ]"
            Case "[", "?", "*", "#": d = d & "[" & ch & "]"
            Case Else
                If LCase$(ch) <> UCase$(ch) Then d = d & "[" & LCase$(ch) & UCase$(ch) & "]" Else d = d & ch
        End Select
    Next i
    AKSORA_ARAMA_DESEN = "*" & d & "*"
End Function
