Attribute VB_Name = "modAKSORA_IL_ARAMA"
Option Explicit

Private mAramaCalisiyor As Boolean
' MODUL: modAKSORA_IL_ARAMA  |  V5: K dahil ust basliklar ve tablo, M20 aciklama, M3 ANA
Private Const sayfa As String = "MALZEME_ARA"
Private Const ON_EK As String = "AKSORA_V4_"

Public Sub AKSORA_IL_ARAMA_KUR()
    Dim ws As Worksheet, sh As Shape, nm As String, secim As String
    Dim i As Long, il As Variant, scr As Boolean, ev As Boolean
    Dim adres As Variant, eskiAlan As Range, yeniAlan As Range, metin As Variant
    scr = application.ScreenUpdating: ev = application.EnableEvents
    On Error GoTo Hata
    application.ScreenUpdating = False: application.EnableEvents = False
    Set ws = ThisWorkbook.Worksheets(sayfa)
    ' V2/V3 eski alanlarini temizle; arama ve sonuc alanlarini koru.
    On Error Resume Next
    ws.Range("L4:N22").UnMerge
    ws.Range("M20:N20").UnMerge
    ws.Range("M21:N22").UnMerge
    ws.Range("L4:N22").Validation.Delete
    ws.Range("P2").Validation.Delete
    On Error GoTo Hata
    ws.Range("L4:N22").ClearContents
    ws.Range("L4:N22").Interior.Pattern = xlNone
    ws.Range("P2").ClearContents
    ws.Range("M21:N22").ClearContents
    ws.Columns("K").Hidden = False
    ws.Columns("K").ColumnWidth = 17
    ws.Columns("L").ColumnWidth = 4
    ws.Columns("M").ColumnWidth = 18
    ws.Columns("N").ColumnWidth = 18
    ws.Rows(12).RowHeight = 29
    ws.Rows("13:1000").RowHeight = 20
    ws.Rows("6:22").RowHeight = 22
    ws.Rows(12).RowHeight = 29

    ' Ust alanin tamami K sutununa kadar uzansin.
    ' Var olan basliklarin metinleri A1, A3, A5 ve A7'de korunur.
    ws.Range("J1:J11").Copy
    ws.Range("K1:K11").PasteSpecial Paste:=xlPasteFormats
    application.CutCopyMode = False
    For Each adres In Array("A1:J2|A1:K2", "A3:J3|A3:K3", _
                            "A5:J5|A5:K5", "A7:J7|A7:K7", _
                            "I9:J9|I9:K9", "I10:J10|I10:K10")
        Set eskiAlan = ws.Range(Split(CStr(adres), "|")(0))
        Set yeniAlan = ws.Range(Split(CStr(adres), "|")(1))
        metin = eskiAlan.Cells(1, 1).Value
        eskiAlan.UnMerge
        yeniAlan.Merge
        yeniAlan.Cells(1, 1).Value = metin
        yeniAlan.HorizontalAlignment = xlCenter
        yeniAlan.VerticalAlignment = xlCenter
    Next adres
    ws.Range("K4:K11").Interior.Color = ws.Range("J4").Interior.Color
    ws.Range("K9").Interior.Color = ws.Range("J9").Interior.Color
    ws.Range("K10").Interior.Color = ws.Range("J10").Interior.Color

    With ws.Range("M6:N7")
        .Merge
        .Value = "ARAMA YAPILACAK ÝL"
        .Interior.Color = RGB(24, 55, 88)
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = True
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With

    ' Yalnizca bu modulun ekledigi sekilleri sil.
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If Left$(sh.Name, Len(ON_EK)) = ON_EK Then sh.Delete
    Next i
    ' Eski surumlerin dugmelerini de kaldir; ANA ve ARA korunur.
    For i = ws.Shapes.Count To 1 Step -1
        Set sh = ws.Shapes(i)
        If Left$(sh.Name, 10) = "AKSORA_V3_" Or _
           Left$(sh.Name, 10) = "AKSORA_V2_" Then sh.Delete
    Next i
    For Each sh In ws.Shapes
        nm = ""
        On Error Resume Next
        nm = UCase$(Trim$(sh.TextFrame.Characters.Text))
        On Error GoTo Hata
        If nm = "ANA" Or nm = "ANA PANEL" Then
            sh.Left = ws.Range("M3").Left + 1
            sh.Top = ws.Range("M3").Top + 1
            sh.Width = ws.Range("M3").Width - 2
            sh.Height = ws.Range("M3").Height - 2
            sh.Placement = xlFreeFloating
        ElseIf nm = "ARA" Then
            sh.OnAction = "'" & ThisWorkbook.Name & "'!AKSORA_IL_ARAMA_YAP"
        End If
    Next sh

    i = 0
    For Each il In Array("TÜMÜ", "GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN")
        AKSORA_V4_DUGME ws, CStr(il), 8 + i * 2
        i = i + 1
    Next il
    With ws.Range("M20:N20")
        .Merge
        .Value = "Seçili ildeki mevcut stoklar aranýr."
        .WrapText = False
        .Font.Name = "Aptos": .Font.Size = 8
        .Font.Color = RGB(85, 103, 123)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    secim = AKSORA_V4_NORM(CStr(ws.Range("AA2").Value2))
    If Not AKSORA_V4_GECERLI(secim) Then secim = "TUMU"
    ws.Range("AA2").Value = secim
    ws.Columns("AA").Hidden = True
    AKSORA_V4_DUGMELERI_RENKLE ws, secim

    ' Baslik + filtre + veri: A:K tek tablo. K sutunu da tum satirlarda yer alir.
    With ws.Range("A12:K12")
        .Interior.Color = RGB(24, 55, 88)
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = True
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlNone
    End With
    ws.Range("K12").Value = "ÝL"
    ws.Range("K12").Interior.Color = ws.Range("J12").Interior.Color
    ' J sutunundaki ayni tablo bicimini IL sutununa tasir.
    ws.Range("J13:J1000").Copy
    ws.Range("K13:K1000").PasteSpecial Paste:=xlPasteFormats
    application.CutCopyMode = False
    With ws.Range("K13:K1000")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(215, 228, 242)
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).Color = RGB(215, 228, 242)
        .Borders(xlInsideHorizontal).LineStyle = xlContinuous
        .Borders(xlInsideHorizontal).Color = RGB(215, 228, 242)
    End With
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A12:K" & application.Max(13, AKSORA_SON_SATIR(ws, "D"))).AutoFilter
    AKSORA_V11_ARAMA_DUGME
    AKSORA_V11_ARAMA_TARIH
    ws.Activate
    ActiveWindow.DisplayGridlines = False
Cikis:
    application.EnableEvents = ev: application.ScreenUpdating = scr
    Exit Sub
Hata:
    MsgBox "Kurulum hatasý: " & Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Sub AKSORA_V4_DUGME(ByVal ws As Worksheet, ByVal etiket As String, ByVal satir As Long)
    Dim sh As Shape, alan As Range, kod As String
    Set alan = ws.Range("M" & satir & ":N" & (satir + 1))
    kod = AKSORA_V4_NORM(etiket)
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, alan.Left + 2, alan.Top + 1, alan.Width - 4, alan.Height - 2)
    sh.Name = ON_EK & kod
    sh.TextFrame.Characters.Text = etiket
    sh.TextFrame.HorizontalAlignment = xlHAlignCenter
    sh.TextFrame.VerticalAlignment = xlVAlignCenter
    With sh.TextFrame.Characters.Font
        .Name = "Aptos": .Size = 10: .bold = True
    End With
    sh.line.Visible = msoFalse
    sh.Placement = xlFreeFloating
    sh.OnAction = "'" & ThisWorkbook.Name & "'!AKSORA_V4_SEC_" & kod
End Sub

Private Sub AKSORA_V4_DUGMELERI_RENKLE(ByVal ws As Worksheet, ByVal secim As String)
    Dim sh As Shape, kod As String
    For Each sh In ws.Shapes
        If Left$(sh.Name, Len(ON_EK)) = ON_EK Then
            kod = Mid$(sh.Name, Len(ON_EK) + 1)
            If kod = secim Then
                sh.Fill.ForeColor.RGB = RGB(0, 121, 103)
                sh.TextFrame.Characters.Font.Color = vbWhite
            Else
                sh.Fill.ForeColor.RGB = RGB(230, 240, 248)
                sh.TextFrame.Characters.Font.Color = RGB(24, 55, 88)
            End If
        End If
    Next sh
End Sub

Private Sub AKSORA_V4_SEC(ByVal secim As String)
    Dim ws As Worksheet, ev As Boolean, sc As Boolean, Hata As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or mAramaCalisiyor Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(sayfa)
    ws.Range("AA2").Value2 = secim
    AKSORA_V4_DUGMELERI_RENKLE ws, secim
    If Len(Trim$(CStr(ws.Range("B6").Value2))) >= 2 Then
        AKSORA_IL_ARAMA_YAP
    Else
        AKSORA_ARAMA_SONUCLARI_TEMIZLE
    End If
    AKSORA_V13_GORUNUM ws
Cikis:
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If Len(Hata) > 0 Then MsgBox Hata, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    Hata = "Ýl seçimi tamamlanamadý: " & Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_V4_SEC_TUMU(): AKSORA_V4_SEC "TUMU": End Sub
Public Sub AKSORA_V4_SEC_GIRESUN(): AKSORA_V4_SEC "GIRESUN": End Sub
Public Sub AKSORA_V4_SEC_TRABZON(): AKSORA_V4_SEC "TRABZON": End Sub
Public Sub AKSORA_V4_SEC_RIZE(): AKSORA_V4_SEC "RIZE": End Sub
Public Sub AKSORA_V4_SEC_GUMUSHANE(): AKSORA_V4_SEC "GUMUSHANE": End Sub
Public Sub AKSORA_V4_SEC_ARTVIN(): AKSORA_V4_SEC "ARTVIN": End Sub

Public Sub AKSORA_IL_ARAMA_YAP()
    Dim hedef As Worksheet, kaynak As Worksheet, il As Variant, grup As Variant
    Dim secim As String, terim As String, kod As String, dogrulanmisCanli As String
    Dim veri As Variant, buffer() As Variant, cikti() As Variant
    Dim i As Long, j As Long, n As Long, kapasite As Long, son As Long, eskiSon As Long
    Dim oncelik As Long, ciktiSatir As Long, dDeposu As Boolean
    Dim toplamMiktar As Double, toplamDeger As Double
    Dim hataMetni As String, adim As String, bas As Double
    Dim iptal As XlEnableCancelKey, imlec As XlMousePointer
    Dim depolar As Object, mevcutIller As Object
    Dim scr As Boolean, ev As Boolean, hesap As XlCalculation, desen As String, ac As Boolean
    If mAramaCalisiyor Then Exit Sub
    Set hedef = ThisWorkbook.Worksheets(sayfa)
    terim = Trim$(CStr(hedef.Range("B6").Value2))
    secim = AKSORA_V4_NORM(CStr(hedef.Range("AA2").Value2))
    If Not AKSORA_V4_GECERLI(secim) Then secim = "TUMU"
    If Len(terim) < 2 Then
        MsgBox "En az iki karakterli malzeme kodu veya açýklamasý girin.", vbExclamation, "AKSORA"
        Exit Sub
    End If
    scr = application.ScreenUpdating: ev = application.EnableEvents
    hesap = application.Calculation
    iptal = application.EnableCancelKey: imlec = application.Cursor
    ' Turkce harfler (i/I/Ý/ý, þ/s ...) ve buyuk/kucuk harf ayni sayilir.
    desen = AKSORA_ARAMA_DESEN(terim)
    ac = AKSORA_BEKLE_AC("AKSORA  |  MALZEME ARAMA", "Aranýyor: " & terim, 0.05)
    bas = Timer
    On Error GoTo Hata
    mAramaCalisiyor = True
    application.EnableCancelKey = xlErrorHandler
    application.Cursor = xlWait
    adim = "Kaynak doðrulama"
    application.ScreenUpdating = False: application.EnableEvents = False
    application.Calculation = xlCalculationManual
    dogrulanmisCanli = AKSORA_CANLI_STOK_IL()
    Set depolar = CreateObject("Scripting.Dictionary")
    Set mevcutIller = CreateObject("Scripting.Dictionary")
    depolar.CompareMode = vbTextCompare: mevcutIller.CompareMode = vbTextCompare
    kapasite = 500
    ReDim buffer(1 To 11, 1 To kapasite)
    For Each il In Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
        If secim = "TUMU" Or secim = CStr(il) Then
            For Each grup In Array("CDP_STOK", "EPP_STOK", "ILCE_STOK", "YUKLENICI_STOK")
                AKSORA_BEKLE_ILERLE 0.1, "Aranýyor: " & terim & "  •  " & il & " " & grup
                Set kaynak = AKSORA_V4_KAYNAK(CStr(il), CStr(grup), dogrulanmisCanli, veri)
                If Not kaynak Is Nothing Then
                    mevcutIller(CStr(il)) = True
                    son = AKSORA_SON_SATIR(kaynak, "G")
                    If son >= 7 Then
                        ' Kaynak doðrulamasýnda okunan ayný dizi; ikinci Excel okumasý yok.
                        For i = 1 To UBound(veri, 1)
                            If (i And 4095) = 0 Then AKSORA_BEKLE_NABIZ
                            kod = CStr(veri(i, 7))
                            If kod Like desen Or CStr(veri(i, 8)) Like desen Then
                                n = n + 1
                                If n > kapasite Then
                                    kapasite = kapasite * 2
                                    ReDim Preserve buffer(1 To 11, 1 To kapasite)
                                End If
                                buffer(1, n) = veri(i, 1)
                                buffer(2, n) = veri(i, 2)
                                buffer(3, n) = veri(i, 3)
                                buffer(4, n) = veri(i, 7)
                                buffer(5, n) = veri(i, 8)
                                buffer(6, n) = veri(i, 6)
                                buffer(7, n) = veri(i, 9)
                                buffer(8, n) = veri(i, 10)
                                buffer(9, n) = veri(i, 11)
                                buffer(10, n) = CStr(grup)
                                buffer(11, n) = CStr(il)
                                If IsNumeric(veri(i, 9)) Then toplamMiktar = toplamMiktar + CDbl(veri(i, 9))
                                If IsNumeric(veri(i, 11)) Then toplamDeger = toplamDeger + CDbl(veri(i, 11))
                                depolar(CStr(il) & "|" & CStr(grup) & "|" & CStr(veri(i, 1)) & "|" & CStr(veri(i, 2))) = True
                            End If
                        Next i
                    End If
                End If
                Set kaynak = Nothing
            Next grup
        End If
    Next il
    adim = "Sonuçlarý yazma"
    If hedef.AutoFilterMode Then hedef.AutoFilterMode = False
    eskiSon = AKSORA_SON_SATIR(hedef, "D")
    If eskiSon >= 13 Then
        With hedef.Range("A13:K" & eskiSon)
            .ClearContents
            .Interior.Color = vbWhite
        End With
    End If
    hedef.Range("A10").Value = n
    hedef.Range("C10").Value = depolar.Count
    hedef.Range("F10").Value = toplamMiktar
    hedef.Range("I10").Value = toplamDeger
    hedef.Range("A10").NumberFormat = "#,##0"
    hedef.Range("C10").NumberFormat = "#,##0"
    hedef.Range("F10").NumberFormat = "#,##0.##"
    hedef.Range("I10").NumberFormat = "#,##0.00 \T\L"
    If n > 0 Then
        ReDim cikti(1 To n, 1 To 11)
        ' Mavi D depolarý önce; her grubun kendi sýrasý korunur.
        For oncelik = 1 To 2
            For i = 1 To n
                dDeposu = (Left$(UCase$(Trim$(CStr(buffer(2, i)))), 1) = "D")
                If (oncelik = 1 And dDeposu) Or (oncelik = 2 And Not dDeposu) Then
                    ciktiSatir = ciktiSatir + 1
                    For j = 1 To 11
                        cikti(ciktiSatir, j) = buffer(j, i)
                    Next j
                End If
            Next i
        Next oncelik
        hedef.Range("A13").Resize(n, 11).Value = cikti
        hedef.Range("G13:G" & 12 + n).NumberFormat = "#,##0.##"
        hedef.Range("I13:I" & 12 + n).NumberFormat = "#,##0.00 ""TL"""
        hedef.Range("K13:K" & 12 + n).HorizontalAlignment = xlCenter
        ' Sonuç biçimleri doðrudan uygulanýr; pano kullanýlmaz.
    ElseIf mevcutIller.Count = 0 Then
        hedef.Range("A13").Value = "Seçilen il için dosyada stok verisi bulunmuyor."
    Else
        hedef.Range("A13").Value = "Aranan malzeme bulunamadý."
    End If
    If n > 0 Then hedef.Range("A12:K" & 12 + n).AutoFilter
    adim = "Sonuç biçimleri"
    AKSORA_V11_BICIM_SAYFA hedef
    If n > 0 Then
        With hedef.Range("A13:K" & 12 + n)
            .IndentLevel = 0
            .VerticalAlignment = xlCenter
        End With
        hedef.Range("D13:F" & 12 + n).HorizontalAlignment = xlLeft
        hedef.Range("G13:G" & 12 + n).HorizontalAlignment = xlRight
        hedef.Range("I13:I" & 12 + n).HorizontalAlignment = xlRight
        hedef.Range("I13:I" & 12 + n).NumberFormat = "#,##0.00 ""TL"""
    End If
    AKSORA_ARAMA_DEPO_RENKLERI
    hedef.Range("AZ1").Formula = "=SUBTOTAL(103,D13:D" & application.Max(13, 12 + n) & ")"
    hedef.Columns("AZ").Hidden = True
    adim = "Arama özeti"
    AKSORA_V13_ARAMA_OZET hedef
    adim = "Ekran görünümü"
    AKSORA_V13_GORUNUM hedef
Cikis:
    On Error Resume Next
    application.Calculation = hesap
    If Err.Number <> 0 And Len(hataMetni) = 0 Then hataMetni = "Hesaplama ayarý: " & Err.Description
    Err.Clear
    application.EnableEvents = ev
    application.ScreenUpdating = scr
    application.Cursor = imlec
    application.EnableCancelKey = iptal
    mAramaCalisiyor = False
    AKSORA_V13_SURE_KAYDET "Malzeme arama", bas
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    If Len(hataMetni) > 0 Then MsgBox hataMetni, vbExclamation, "AKSORA | Arama"
    Exit Sub
Hata:
    If Err.Number = 18 Then
        hataMetni = "Arama kullanýcý tarafýndan durduruldu."
    Else
        hataMetni = adim & ": " & Err.Description
    End If
    Resume Cikis
End Sub

Private Function AKSORA_V4_KAYNAK( _
    ByVal il As String, _
    ByVal grup As String, _
    ByVal yuklu As String, _
    ByRef kontrol As Variant) As Worksheet

    Dim ws As Worksheet
    Dim ad As String, tip As String
    Dim son As Long, r As Long
    Dim tesis As String, cdp As String, epp As String
    Dim arsiv As Boolean

    kontrol = Empty
    cdp = AKSORA_IL_TESIS(il, "CDP")
    epp = AKSORA_IL_TESIS(il, "EPP")

    Select Case grup
        Case "CDP_STOK"
            ad = "CDP_STOK": tip = "CDP"
        Case "EPP_STOK"
            ad = "EPP_STOK": tip = "EPP"
        Case "ILCE_STOK"
            ad = ChrW(304) & "L" & ChrW(199) & "E_STOK"
            tip = "ILC"
        Case "YUKLENICI_STOK"
            ad = "Y" & ChrW(220) & "KLEN" & ChrW(304) & _
                 "C" & ChrW(304) & "_STOK"
            tip = "YUK"
        Case Else
            Err.Raise vbObjectError + 860, , "Geçersiz arama grubu."
    End Select

    arsiv = (il <> yuklu)
    If arsiv Then ad = "_AKLIVE_" & il & "_" & tip

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0

    If ws Is Nothing Then Exit Function

    If arsiv Then
        If CStr(ws.Range("A1").Value2) <> "AKSORA_CANLI_ARSIV_V1" Or _
           CStr(ws.Range("B1").Value2) <> il Or _
           CStr(ws.Range("C1").Value2) <> tip Then

            Err.Raise vbObjectError + 861, , _
                      "Arama arþivinin kimliði uyuþmuyor: " & ad
        End If
    End If

    son = AKSORA_SON_SATIR(ws, "G")
    If son < 7 Then Exit Function

    If son > 250000 Then
        Err.Raise vbObjectError + 862, , _
                  "Beklenmeyen arama kaynaðý: " & ad
    End If

    kontrol = ws.Range("A7:K" & son).Value2
    For r = 1 To UBound(kontrol, 1)
        If IsError(kontrol(r, 7)) Or IsError(kontrol(r, 1)) Then GoTo YANLIS_IL
        If Len(Trim$(CStr(kontrol(r, 7)))) > 0 Then
            tesis = Trim$(CStr(kontrol(r, 1)))

            Select Case tip
                Case "CDP"
                    If tesis <> cdp Then GoTo YANLIS_IL
                Case "EPP"
                    If tesis <> epp Then GoTo YANLIS_IL
                Case Else
                    If tesis <> cdp And tesis <> epp Then GoTo YANLIS_IL
            End Select
        End If
    Next r

    Set AKSORA_V4_KAYNAK = ws
    Exit Function

YANLIS_IL:
    Err.Raise vbObjectError + 863, , _
              "Arama kaynaðýnda baþka ilin verisi var: " & ad
End Function

Private Function AKSORA_V4_GECERLI(ByVal secim As String) As Boolean
    Select Case secim
        Case "TUMU", "GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN": AKSORA_V4_GECERLI = True
    End Select
End Function
Private Function AKSORA_V4_NORM(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, "Ý", "I")
    s = Replace(s, "Ü", "U")
    s = Replace(s, "Þ", "S")
    s = Replace(s, "Ð", "G")
    s = Replace(s, "Ç", "C")
    s = Replace(s, "Ö", "O")
    AKSORA_V4_NORM = s
End Function
Public Sub AKSORA_ARAMA_MIKTAR_BICIMI()
    Dim ws As Worksheet, son As Long
    Set ws = ThisWorkbook.Worksheets("MALZEME_ARA")
    son = AKSORA_SON_SATIR(ws, "D")
    AKSORA_V13_MIKTAR_BICIMI ws.Range("F10")
    If son >= 13 Then AKSORA_V13_MIKTAR_BICIMI ws.Range("G13:G" & son)
End Sub

Public Sub AKSORA_ARAMA_DEPO_RENKLERI()
    Dim ws As Worksheet, son As Long, eskiSon As Long, r As Long
    Dim veri As Variant, ilk As Long, depo As String
    Set ws = ThisWorkbook.Worksheets("MALZEME_ARA")
    son = AKSORA_SON_SATIR(ws, "D")
    eskiSon = application.Max(13, son)
    ws.Range("A13:K" & eskiSon).Interior.Color = vbWhite
    If son < 13 Then Exit Sub
    veri = ws.Range("B13:B" & son).Value2
    If son = 13 Then
        If Left$(UCase$(Trim$(CStr(veri))), 1) = "D" Then ws.Range("A13:K13").Interior.Color = RGB(225, 239, 249)
        Exit Sub
    End If
    For r = 1 To UBound(veri, 1) + 1
        depo = ""
        If r <= UBound(veri, 1) Then depo = UCase$(Trim$(CStr(veri(r, 1))))
        If Left$(depo, 1) = "D" Then
            If ilk = 0 Then ilk = r + 12
        ElseIf ilk > 0 Then
            ws.Range("A" & ilk & ":K" & r + 11).Interior.Color = RGB(225, 239, 249)
            ilk = 0
        End If
    Next r
End Sub

Public Sub AKSORA_ARAMA_SONUCLARI_TEMIZLE()
    Dim ws As Worksheet
    Dim son As Long, r As Long, c As Long
    Dim olay As Boolean
    Dim hataNo As Long, hataMetni As String

    olay = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False

    Set ws = ThisWorkbook.Worksheets("MALZEME_ARA")
    son = 13

    For c = 1 To 11
        r = AKSORA_SON_SATIR(ws, c)
        If r > son Then son = r
    Next c

    If ws.FilterMode Then ws.ShowAllData

    With ws.Range("A13:K" & son)
        .ClearContents
        .Interior.Color = RGB(255, 255, 255)
    End With

    ws.Range("A10").MergeArea.ClearContents
    ws.Range("C10").MergeArea.ClearContents
    ws.Range("F10").MergeArea.ClearContents
    ws.Range("I10").MergeArea.ClearContents
    On Error Resume Next
    ws.Range("F10").Comment.Delete
    On Error GoTo Hata

    application.EnableEvents = olay
    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description
    application.EnableEvents = olay
    Err.Raise hataNo, "AKSORA_ARAMA_SONUCLARI_TEMIZLE", hataMetni
End Sub


