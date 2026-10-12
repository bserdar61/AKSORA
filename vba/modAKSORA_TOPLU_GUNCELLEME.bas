Attribute VB_Name = "modAKSORA_TOPLU_GUNCELLEME"
Option Explicit

Private mTopluSaatAktif As Boolean
Private mTopluSaatBas As Double
Private mTopluSonSaniye As Long
Private mTopluDurumMetni As String
Public AKSORA_V11_MESGUL As Boolean
Private mAK14Sureler As Object
Private mAK14Toplam As Double
Private mAK14Arayuz As Double
Private Const PARA As String = "#,##0.00 ""TL"""
Private Const miktar As String = "#,##0"

Public Sub AKSORA_V11_ESKI_KONTROL_SIL()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        AK12_KONTROL_SIL ws
    Next ws
    application.DisplayStatusBar = True
    application.StatusBar = False
End Sub

Public Sub AKSORA_V11_BICIM_SAYFA(ByVal ws As Worksheet)
    Dim h As Long, son As Long, kolon As Long, c As Long, ilk As Long
    Dim baslik As String, alan As Range, metin As String, r As Long, hucre As Range
    Select Case AKFIX_NORMAL(ws.Name)
        Case "CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK"
            h = 6: kolon = 12: son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
            ws.Range("C4").NumberFormat = PARA
        Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR"
            h = 8: kolon = ws.Cells(h, ws.Columns.Count).End(xlToLeft).Column
            If kolon > 5 Then kolon = 5
            son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
            ' End at the last real depot/total row; leave the note outside the table.
            Do While son > h
                metin = AKFIX_NORMAL(CStr(ws.Cells(son, 1).Value2))
                If metin = "TOPLAM" Or metin = "GENEL TOPLAM" Then Exit Do
                If Len(metin) = 4 And Len(CStr(ws.Cells(son, 2).Value2)) > 0 Then Exit Do
                son = son - 1
            Loop
            ws.Range("B5").NumberFormat = PARA
        Case "MALZEME_ARA"
            h = 12: kolon = 11: son = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
            ws.Range("I10").NumberFormat = PARA
        Case "STOK_KARSILASTIRMA"
            h = 15: kolon = 13: son = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        Case "ANA PANEL"
            ' Numeric dashboard money cells retain their values and merges.
            For r = 1 To 40
                For c = 1 To 16
                    If ws.Cells(r, c).MergeArea.Cells(1, 1).Address = ws.Cells(r, c).Address Then
                        If IsNumeric(ws.Cells(r, c).Value2) And Len(CStr(ws.Cells(r, c).Value2)) > 0 Then
                            If InStr(1, ws.Cells(r, c).NumberFormat, "TL", vbTextCompare) > 0 Then ws.Cells(r, c).NumberFormat = PARA
                        End If
                    End If
                Next c
            Next r
            Exit Sub
        Case Else: Exit Sub
    End Select
    With ws.Range(ws.Cells(h, 1), ws.Cells(h, kolon))
        .Font.Name = "Aptos"
        .Font.Size = 10
        .Font.bold = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .IndentLevel = 0
        .ShrinkToFit = True
    End With
    If son <= h Then Exit Sub
    ilk = h + 1
    With ws.Range(ws.Cells(ilk, 1), ws.Cells(son, kolon))
        .IndentLevel = 0
        .VerticalAlignment = xlCenter
        .Font.Name = "Aptos"
        .Font.Size = 10
    End With
    For c = 1 To kolon
        baslik = AKFIX_NORMAL(CStr(ws.Cells(h, c).Value2))
        Set alan = ws.Range(ws.Cells(ilk, c), ws.Cells(son, c))
        If InStr(baslik, "DEGER") > 0 Or InStr(baslik, "FIYAT") > 0 Or InStr(baslik, "TUTAR") > 0 Then
            alan.HorizontalAlignment = xlRight
            alan.NumberFormat = PARA
        ElseIf InStr(baslik, "MIKTAR") > 0 Or InStr(baslik, "TAHDITSIZ") > 0 Or InStr(baslik, "STOK SATIRI") > 0 Or baslik = "FARK" Then
            alan.HorizontalAlignment = xlRight
            AKSORA_V13_MIKTAR_BICIMI alan
        ElseIf InStr(baslik, "TANIM") > 0 Or InStr(baslik, "METNI") > 0 Or InStr(baslik, "ACIKLAMA") > 0 Or baslik = "MALZEME" Then
            alan.HorizontalAlignment = xlLeft
        Else
            alan.HorizontalAlignment = xlCenter
        End If
    Next c
    If AKFIX_NORMAL(ws.Name) = "MALZEME_ARA" Then
        ws.Range("C" & ilk & ":F" & son).HorizontalAlignment = xlLeft
    End If
End Sub

Public Sub AKSORA_V11_BICIMLER()
    Dim ws As Worksheet
    application.UseSystemSeparators = False
    application.DecimalSeparator = ","
    application.ThousandsSeparator = "."
    For Each ws In ThisWorkbook.Worksheets
        AKSORA_V11_BICIM_SAYFA ws
    Next ws
End Sub

Public Function AKSORA_V11_IL_TARIH(ByVal il As String, Optional ByVal dogrulanmisCanli As String = "") As Date
    Dim tip As Variant, kaynak As Variant, i As Long, t As Date, tarih As Date
    Dim ar As Worksheet, ws As Worksheet, canli As String
    il = AKFIX_NORMAL(il)
    canli = dogrulanmisCanli
    If Len(canli) = 0 Then canli = AKSORA_CANLI_STOK_IL()
    tip = Array("CDP", "EPP", "TUM", "ILC", "YUK")
    kaynak = Array("CDP_STOK", "EPP_STOK", "STOKLAR", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
    For i = 0 To 4
        tarih = 0
        Set ar = Nothing
        On Error Resume Next
        Set ar = ThisWorkbook.Worksheets("_AKLIVE_" & il & "_" & tip(i))
        On Error GoTo 0
        If il = canli Then
            Set ws = ThisWorkbook.Worksheets(CStr(kaynak(i)))
            tarih = AK12_TARIH(ws.Range("C2").Value2)
        End If
        If tarih = 0 And Not ar Is Nothing Then
            If CStr(ar.Range("A1").Value2) = "AKSORA_CANLI_ARSIV_V1" And _
               CStr(ar.Range("B1").Value2) = il And _
               CStr(ar.Range("C1").Value2) = CStr(tip(i)) Then
                If Val(CStr(ar.Range("A4").Value2)) > 0 And _
                   ar.Cells(ar.Rows.Count, "G").End(xlUp).Row >= 7 Then
                    tarih = AK12_TARIH(ar.Range("A3").Value2)
                End If
            End If
        End If
        If tarih = 0 Then Exit Function
        If t = 0 Or tarih < t Then t = tarih
    Next i
    AKSORA_V11_IL_TARIH = t
End Function

Private Function AK12_TARIH(ByVal v As Variant) As Date
    Dim n As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then
        n = CDbl(v)
        If ThisWorkbook.Date1904 Then n = n + 1462
        If n >= 1 And n < 2958466 Then AK12_TARIH = CDate(n)
    ElseIf IsDate(v) Then
        AK12_TARIH = CDate(v)
    End If
End Function
Public Sub AKSORA_V11_ARAMA_TARIH()
    Dim ws As Worksheet, il As Variant, s As Shape, tarih As Date, etiket As String, canli As String
    Set ws = ThisWorkbook.Worksheets("MALZEME_ARA")
    canli = AKSORA_CANLI_STOK_IL()
    For Each il In Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
        Set s = Nothing
        On Error Resume Next
        Set s = ws.Shapes("AKSORA_V4_" & il)
        On Error GoTo 0
        If Not s Is Nothing Then
            etiket = CStr(il)
            Select Case il
                Case "GIRESUN": etiket = "GÝRESUN"
                Case "RIZE": etiket = "RÝZE"
                Case "GUMUSHANE": etiket = "GÜMÜÞHANE"
                Case "ARTVIN": etiket = "ARTVÝN"
            End Select
            tarih = AKSORA_V11_IL_TARIH(CStr(il), canli)
            If tarih > 0 Then
                etiket = etiket & vbLf & Format$(tarih, "dd.mm.yyyy hh:nn")
            Else
                etiket = etiket & vbLf & "Henüz güncellenmedi"
            End If
            s.Placement = xlFreeFloating
            With s.TextFrame2
                .TextRange.Text = etiket
                .TextRange.Font.Size = 9
                .MarginTop = 2: .MarginBottom = 2
                .VerticalAnchor = msoAnchorMiddle
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End With
        End If
    Next il
End Sub

Public Sub AKSORA_V11_ARAMA_DUGME()
    Dim ws As Worksheet, s As Shape, a As Range, onceki As Shape
    Dim durum As Shape, vurgu As Shape
    Set ws = ThisWorkbook.Worksheets("MALZEME_ARA")
    On Error Resume Next
    Set onceki = ws.Shapes("AKSORA_TOPLU_GUNCELLE")
    Set durum = ws.Shapes("AKSORA_TOPLU_DURUM")
    Set vurgu = ws.Shapes("AKSORA_TOPLU_VURGU")
    On Error GoTo 0
    If Not onceki Is Nothing Then
        If onceki.AlternativeText = "AKSORA_V12_M26_N28" Then
            If Not durum Is Nothing And Not vurgu Is Nothing Then
                AKSORA_V12_BUTON_RENK
                Exit Sub
            End If
        End If
    End If
    On Error Resume Next
    ws.Shapes("AKSORA_TOPLU_GUNCELLE").Delete
    ws.Shapes("AKSORA_TOPLU_DURUM").Delete
    ws.Shapes("AKSORA_TOPLU_VURGU").Delete
    On Error GoTo 0
    ws.Rows("26:31").RowHeight = 22
    Set a = ws.Range("M26:N28")
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, a.Left + 2, a.Top + 2, a.Width - 4, a.Height - 4)
    With s
        .Name = "AKSORA_TOPLU_GUNCELLE"
        .AlternativeText = "AKSORA_V12_M26_N28"
        .Adjustments.Item(1) = 0.12
        .TextFrame2.TextRange.Text = "BÜTÜN ÝLLERÝ GÜNCELLE" & vbLf & "CDP + EPP  |  5 ÝL"
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(225, 255, 246)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.MarginLeft = 14: .TextFrame2.MarginRight = 10
        .Fill.ForeColor.RGB = RGB(9, 45, 62)
        .Fill.BackColor.RGB = RGB(0, 112, 84)
        .Fill.TwoColorGradient msoGradientHorizontal, 1
        .line.Visible = msoTrue
        .line.ForeColor.RGB = RGB(28, 211, 163): .line.Weight = 1.25
        .Shadow.Visible = msoTrue
        .Shadow.ForeColor.RGB = RGB(10, 39, 50)
        .Shadow.Transparency = 0.8
        .Shadow.OffsetX = 1: .Shadow.OffsetY = 2
        .Placement = xlFreeFloating
        .Locked = True
        .OnAction = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!AKSORA_BUTUN_ILLERI_GUNCELLE"
    End With
    Set onceki = s
    Set s = ws.Shapes.AddShape(msoShapeRectangle, a.Left + 8, a.Top + 12, 3, a.Height - 24)
    With s
        .Name = "AKSORA_TOPLU_VURGU"
        .Fill.ForeColor.RGB = RGB(48, 231, 179): .line.Visible = msoFalse
        .Placement = xlFreeFloating: .Locked = True
        .OnAction = onceki.OnAction
    End With
    Set a = ws.Range("M29:N31")
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, a.Left, a.Top, a.Width, a.Height)
    With s
        .Name = "AKSORA_TOPLU_DURUM"
        .TextFrame2.TextRange.Text = "Arama, her ilin yukarýdaki tarihli stoklarýný kullanýr."
        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 9
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(85, 103, 123)
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .Fill.Visible = msoFalse: .line.Visible = msoFalse
        .Placement = xlFreeFloating: .Locked = True
    End With
End Sub
Private Sub AK11_DURUM(ByVal metin As String)
    Dim ws As Worksheet
    If AKSORA_MUSTERI_TOPLU_AKTIF Then
        AKSORA_MUSTERI_TOPLU_DURUM metin
        Exit Sub
    End If
    Set ws = ThisWorkbook.Worksheets("MALZEME_ARA")
    mTopluDurumMetni = metin
    If mTopluSaatAktif Then
        AKSORA_TOPLU_SURE_TIK True
    Else
        ws.Shapes("AKSORA_TOPLU_DURUM").TextFrame2.TextRange.Text = metin
    End If
    application.StatusBar = "AKSORA | " & metin
    DoEvents
End Sub

Private Function AK11_SON(ByVal ws As Worksheet, ByVal kolon As Long) As Long
    Dim c As Long, n As Long
    For c = 1 To kolon
        n = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
        If n > AK11_SON Then AK11_SON = n
    Next c
    If AK11_SON < 1 Then AK11_SON = 1
End Function

Private Function AK11_IL_GUNCELLE(ByVal il As String, ByRef aciklama As String) As Boolean
    Dim ham(1 To 6) As Variant, veri(1 To 5) As Variant
    Dim ar(1 To 8) As Worksheet, eski(1 To 8) As Variant, son(1 To 8) As Long
    Dim tarihler(1 To 6) As Date, tip As Variant, i As Long, j As Long, n As Long
    Dim toplam As Variant, yaziliyor As Boolean, geri As String, Hata As String
    Dim ilBas As Double, sapBas As Double, sapSure As Double, excelSure As Double
    ilBas = Timer
    On Error GoTo Hata
    tip = Array("CDP", "EPP", "TUM", "ILC", "YUK")
    For i = 1 To 5
        Set ar(i) = ThisWorkbook.Worksheets("_AKLIVE_" & il & "_" & tip(i - 1))
        If ar(i).ProtectContents Then Err.Raise vbObjectError + 1102, , "Arþiv korumalý: " & ar(i).Name
        If CStr(ar(i).Range("A1").Value2) <> "AKSORA_CANLI_ARSIV_V1" Or _
           CStr(ar(i).Range("B1").Value2) <> il Or _
           CStr(ar(i).Range("C1").Value2) <> tip(i - 1) Then
            Err.Raise vbObjectError + 1103, , "Arþiv kimliði uyuþmuyor: " & ar(i).Name
        End If
        son(i) = AK11_SON(ar(i), 12)
        eski(i) = ar(i).Range("A1").Resize(son(i), 12).Value2
    Next i
    tip = Array("INDEX", "DELTA", "CURRENT")
    For i = 6 To 8
        Set ar(i) = ThisWorkbook.Worksheets(AKFIX_HAFIZA_ADI(il, CStr(tip(i - 6))))
        If ar(i).ProtectContents Then Err.Raise vbObjectError + 1104, , "Geçmiþ korumalý: " & ar(i).Name
        son(i) = AK11_SON(ar(i), 15)
        eski(i) = ar(i).Range("A1").Resize(son(i), 15).Value2
    Next i
    For i = 1 To 6
        AK11_DURUM il & " | SAP " & i & "/6"
        sapBas = Timer
        ham(i) = AKIL_SAP_VERI_AL(i, il)
        tarihler(i) = Now
        sapSure = sapSure + AK14_GECEN_SURE(sapBas)
    Next i
    veri(1) = ham(1): veri(2) = ham(2)
    veri(3) = AKIL_VERI_BIRLESTIR(ham(1), ham(2))
    veri(4) = AKIL_VERI_BIRLESTIR(ham(3), ham(4))
    veri(5) = AKIL_VERI_BIRLESTIR(ham(5), ham(6))
    application.ScreenUpdating = False
    yaziliyor = True
    For i = 1 To 5
        n = UBound(veri(i), 1)
        toplam = CDec(0)
        For j = 1 To n: toplam = toplam + CDec(veri(i)(j, 11)): Next j
        ar(i).Range("A7:L" & application.Max(7, son(i), n + 6)).ClearContents
        For j = 1 To n
            If CDbl(veri(i)(j, 9)) <> 0 Then
                veri(i)(j, 12) = CDbl(veri(i)(j, 11)) / CDbl(veri(i)(j, 9))
            Else
                veri(i)(j, 12) = 0
            End If
        Next j
        ar(i).Range("A7").Resize(n, 12).Value2 = veri(i)
        Select Case i
            Case 1: ar(i).Range("A3").Value = tarihler(1)
            Case 2, 3: ar(i).Range("A3").Value = tarihler(2)
            Case 4: ar(i).Range("A3").Value = tarihler(4)
            Case 5: ar(i).Range("A3").Value = tarihler(6)
        End Select
        ar(i).Range("A4").Value2 = n
        ar(i).Range("A5").Value2 = CDbl(toplam)
        ar(i).Range("D1").Value = Now
        ar(i).Visible = xlSheetVeryHidden
    Next i
    STOK_HAFIZA_KAYDET il
    yaziliyor = False
    application.ScreenUpdating = True
    AK11_IL_GUNCELLE = True
    aciklama = il & ": güncellendi"
    excelSure = AK14_GECEN_SURE(ilBas) - sapSure
    AK14_SURE_KAYDET il, sapSure, excelSure, True
    Exit Function
Hata:
    Hata = Err.Description
    If yaziliyor Then
        On Error Resume Next
        For i = 1 To 8
            Err.Clear
            If i <= 5 Then n = 12 Else n = 15
            ar(i).Range("A1").Resize(application.Max(son(i), AK11_SON(ar(i), n)), n).ClearContents
            ar(i).Range("A1").Resize(son(i), n).Value2 = eski(i)
            If Err.Number <> 0 Then geri = geri & " Geri alma baþarýsýz: " & ar(i).Name
        Next i
        On Error GoTo 0
    End If
    application.ScreenUpdating = True
    aciklama = il & ": güncellenemedi; " & Hata & geri
    excelSure = AK14_GECEN_SURE(ilBas) - sapSure
    AK14_SURE_KAYDET il, sapSure, excelSure, False
    ' A rollback error must stop the batch, never be presented as preserved data.
    If Len(geri) > 0 Then Err.Raise vbObjectError + 1105, , aciklama
End Function

Public Sub AKSORA_BUTUN_ILLERI_GUNCELLE()
    Dim il As Variant, canli As String, secim As String, sonuc As String, bilgi As String
    Dim ev As Boolean, sc As Boolean, hesap As XlCalculation, sessiz As Boolean
    Dim basarili As Boolean, canliGuncel As Boolean, Hata As String, sayfa As Worksheet
    Dim adres As String, satir As Long, sutun As Long, basladi As Boolean
    Dim toplamBas As Double, arayuzBas As Double
    If AKSORA_V11_MESGUL Then Exit Sub
    If MsgBox("Beþ ilin stoklarý SAP'den sýrayla alýnacak." & vbCrLf & _
              "Baþarýsýz ilin önceki kayýtlarý korunur. Devam edilsin mi?", _
              vbYesNo + vbQuestion, "AKSORA | Bütün Ýller") <> vbYes Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    hesap = application.Calculation: sessiz = ANA_SESSIZ_MOD
    On Error GoTo Hata
    toplamBas = Timer
    mTopluSaatBas = toplamBas
    mTopluSonSaniye = -1
    mTopluSaatAktif = True
    mTopluDurumMetni = "Güncelleme hazýrlanýyor..."
    Set mAK14Sureler = Nothing
    mAK14Toplam = 0: mAK14Arayuz = 0
    canli = AKSORA_CANLI_STOK_IL(): secim = AKSORA_AKTIF_IL()
    Set sayfa = ActiveSheet: adres = ActiveCell.Address
    satir = ActiveWindow.ScrollRow: sutun = ActiveWindow.ScrollColumn
    AKSORA_V11_MESGUL = True: ANA_SESSIZ_MOD = True
    application.EnableEvents = False
    application.Calculation = xlCalculationManual
    application.ScreenUpdating = False
    AKSORA_V11_ARAMA_DUGME
    ThisWorkbook.Worksheets("MALZEME_ARA").Visible = xlSheetVisible
    ThisWorkbook.Worksheets("MALZEME_ARA").Activate
    ThisWorkbook.Worksheets("MALZEME_ARA").Shapes("AKSORA_TOPLU_GUNCELLE").Fill.ForeColor.RGB = RGB(125, 140, 135)
    application.ScreenUpdating = True
    basladi = True
    For Each il In Array("ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
        ThisWorkbook.Worksheets("KAPAK").Range("AZ1").Value2 = CStr(il)
        bilgi = ""
        basarili = AK11_IL_GUNCELLE(CStr(il), bilgi)
        If basarili And CStr(il) = canli Then canliGuncel = True
        sonuc = sonuc & bilgi & vbCrLf
        AK11_DURUM bilgi
    Next il
Bitir:
    On Error GoTo SonHata
    application.ScreenUpdating = False
    ThisWorkbook.Worksheets("KAPAK").Range("AZ1").Value2 = secim
    arayuzBas = Timer
    If canliGuncel Then
        ' This routine already formats the refreshed live sheets and reports.
        AKSORA_V11_CANLI_YENILE canli
    Else
        AKSORA_V11_BICIMLER
    End If
    AKSORA_V11_ARAMA_TARIH
    AKSORA_V10_TARIH_GOSTER
    If Len(Trim$(CStr(ThisWorkbook.Worksheets("MALZEME_ARA").Range("B6").Value2))) >= 2 Then
        AKSORA_IL_ARAMA_YAP
    Else
        AKSORA_ARAMA_SONUCLARI_TEMIZLE
    End If
    If Len(Hata) = 0 Then AK11_DURUM "Ýþlem tamamlandý. Baþarýlý illerin tarihleri güncellendi."
    mAK14Arayuz = AK14_GECEN_SURE(arayuzBas)
Temizle:
    On Error Resume Next
    If basladi Then
        ThisWorkbook.Worksheets("KAPAK").Range("AZ1").Value2 = secim
        AKSORA_V12_BUTON_RENK
        sayfa.Activate
        AKSORA_V13_GORUNUM sayfa
        ActiveWindow.ScrollRow = satir: ActiveWindow.ScrollColumn = sutun
    End If
    application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = sc
    application.StatusBar = False: application.DisplayStatusBar = True
    ANA_SESSIZ_MOD = sessiz: AKSORA_V11_MESGUL = False
    If basladi Then mAK14Toplam = AK14_GECEN_SURE(toplamBas)
    If mTopluSaatAktif Then AKSORA_TOPLU_SURE_TIK True
    mTopluSaatAktif = False
    On Error GoTo 0
    If Len(Hata) > 0 Then
        MsgBox sonuc & vbCrLf & Hata, vbExclamation, "AKSORA | Güncelleme Sonucu"
    Else
        MsgBox sonuc, vbInformation, "AKSORA | Güncelleme Sonucu"
    End If
    Exit Sub
Hata:
    Hata = "Ýþlem durduruldu: " & Err.Description
    If basladi Then Resume Bitir
    GoTo Temizle
SonHata:
    Hata = Hata & vbCrLf & "Görünüm yenilenemedi: " & Err.Description
    Resume Temizle
End Sub

Public Sub AKSORA_V11_KUR()
    Dim ev As Boolean, sc As Boolean, ws As Worksheet, Hata As String
    ev = application.EnableEvents: sc = application.ScreenUpdating
    Set ws = ActiveSheet
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    AKSORA_V11_ESKI_KONTROL_SIL
    AKSORA_V11_BICIMLER
    AKSORA_ARAMA_MIKTAR_BICIMI
    AKSORA_V11_ARAMA_DUGME
    AKSORA_V11_ARAMA_TARIH
    AKSORA_V10_TARIH_GOSTER
Cikis:
    On Error Resume Next
    ws.Activate
    application.EnableEvents = True: application.ScreenUpdating = True
    On Error GoTo 0
    If Len(Hata) > 0 Then
        MsgBox Hata, vbExclamation, "AKSORA | Kurulum"
    Else
        MsgBox "Düzenlemeler uygulandý. Stoklar ve geçmiþ kayýtlarý korunmuþtur.", vbInformation, "AKSORA"
    End If
    Exit Sub
Hata:
    Hata = "Kurulum tamamlanamadý: " & Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_V11_KARS_UYARI(ByVal ws As Worksheet)
    Dim s As Shape, eski As Shape
    Dim buton As Shape, alt As Shape
    Dim x As Double, y As Double, genislik As Double
    Dim metin As String, adim As String
    Dim renk As Long
    Dim icerik As Boolean, senaryo As Boolean
    Dim hataNo As Long, hataMetni As String

    icerik = ws.ProtectContents
    senaryo = ws.ProtectScenarios

    On Error GoTo Hata
    adim = "Sayfa korumasýný geçici kaldýrma"
    ws.Unprotect Password:=""

    On Error Resume Next
    Set eski = ws.Shapes("AKSORA_GECMIS_KAYIT_NOTU")
    Set buton = ws.Shapes("AKSKV1_ARTVIN")
    Set alt = ws.Shapes("AKSORA_KARS_SABIT_ACIKLAMA")
    On Error GoTo Hata

    adim = "Konum belirleme"
    x = ws.Columns("O").Left
    genislik = ws.Range("O1:P1").Width
    y = 480

    If Not buton Is Nothing Then
        x = buton.Left
        y = buton.Top + buton.Height + 35
    End If

    If Not alt Is Nothing Then
        y = alt.Top + alt.Height + 8
    End If

    If genislik < 100 Then genislik = 240
    If x < 0 Then x = 0
    If y < 0 Then y = 480

    adim = "Eski açýklama kutusunu kaldýrma"
    If Not eski Is Nothing Then Set s = eski

    If AKFIX_KARS_GECMIS_FARK >= 0.005 Then
        metin = "Eski deðer geçmiþinde " & _
                AKSORA_SAYI_METNI(AKFIX_KARS_GECMIS_FARK, "#,##0.00") & _
                " TL kayýt farký var. Küçük TL deðiþiklikleri " & _
                "bu tarihler için kesin deðildir."
        renk = RGB(160, 90, 25)
    Else
        metin = "Karþýlaþtýrma, seçili ilin kayýtlý " & _
                "stok geçmiþini kullanýr."
        renk = RGB(85, 103, 123)
    End If

    adim = "Yeni açýklama kutusunu oluþturma"
    If s Is Nothing Then Set s = ws.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, x, y, genislik, 64)

    With s
        .Name = "AKSORA_GECMIS_KAYIT_NOTU"
        .Left = x: .Top = y: .Width = genislik: .Height = 64
        .Placement = xlFreeFloating
        .Locked = True
        .Fill.Visible = msoFalse
        .line.Visible = msoFalse
    End With

    adim = "Açýklama yazýsýný biçimlendirme"
    With s.TextFrame
        .AutoSize = False
        .MarginLeft = 4
        .MarginRight = 4
        .MarginTop = 2
        .MarginBottom = 2
        .HorizontalAlignment = xlHAlignCenter
        .VerticalAlignment = xlVAlignCenter
        .Characters.Text = metin
        .Characters.Font.Name = "Aptos"
        .Characters.Font.Size = 9
        .Characters.Font.Color = renk
    End With

Bitir:
    On Error GoTo KorumaHatasi
    ws.Protect Password:="", DrawingObjects:=True, _
        Contents:=icerik, Scenarios:=senaryo, _
        UserInterfaceOnly:=True, _
        AllowFiltering:=True, AllowSorting:=True

    On Error GoTo 0
    If hataNo <> 0 Then
        Err.Raise hataNo, "AKSORA", hataMetni
    End If
    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = adim & vbCrLf & Err.Description
    Resume Bitir

KorumaHatasi:
    hataMetni = hataMetni & vbCrLf & _
                "Sayfa korumasý geri açýlamadý: " & Err.Description
    On Error GoTo 0
    Err.Raise vbObjectError + 1210, "AKSORA", hataMetni
End Sub
Public Sub AKSORA_V12_GIT(ByVal sayfaAdi As String)
    AKSORA_V13_GIT sayfaAdi
End Sub

Public Sub AKSORA_V12_YUK_ALT_SABITLE()
    Dim ws As Worksheet, r As Long, c As Long, son As Long
    Dim v As Variant
    Set ws = ThisWorkbook.Worksheets("YÜKLENÝCÝ_RAPOR")
    son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    ' Some older versions used cell text. Remove only matching footer text.
    For r = 9 To son
        For c = 1 To 5
            v = ws.Cells(r, c).Value2
            If VarType(v) = vbString Then
                If InStr(1, CStr(v), "Veriler SAP", vbTextCompare) > 0 Then ws.Cells(r, c).ClearContents
            End If
        Next c
    Next r
    AKSORA_RAPOR_TABLO_BITISI ws.Name
    AKSORA_RAPOR_ALT_YAZI_YERLESTIR ws.Name
End Sub

Public Sub AKSORA_V12_BUTON_RENK()
    With ThisWorkbook.Worksheets("MALZEME_ARA").Shapes("AKSORA_TOPLU_GUNCELLE").Fill
        .ForeColor.RGB = RGB(9, 45, 62)
        .BackColor.RGB = RGB(0, 112, 84)
        .TwoColorGradient msoGradientHorizontal, 1
    End With
End Sub

Private Function AK12_ESKI_SEKIL(ByVal s As Shape) As Boolean
    Dim ad As String, makro As String
    ad = UCase$(s.Name)
    On Error Resume Next
    makro = UCase$(s.OnAction)
    On Error GoTo 0
    AK12_ESKI_SEKIL = (Left$(ad, 13) = "AKSORA_SECIM_" Or _
        Left$(ad, 12) = "AKSORA_ZOOM_" Or _
        InStr(makro, "AKSORA_ZOOM_ARTIR") > 0 Or _
        InStr(makro, "AKSORA_ZOOM_AZALT") > 0 Or _
        InStr(makro, "AKSORA_ZOOM_SIFIRLA") > 0)
End Function

Private Sub AK12_KONTROL_SIL(ByVal ws As Worksheet)
    Dim i As Long, bulundu As Boolean, acildi As Boolean
    Dim cizim As Boolean, icerik As Boolean, senaryo As Boolean, secim As Long
    Dim izin(1 To 11) As Boolean, no As Long, metin As String
    For i = 1 To ws.Shapes.Count
        If AK12_ESKI_SEKIL(ws.Shapes(i)) Then bulundu = True: Exit For
    Next i
    If Not bulundu Then Exit Sub
    On Error GoTo Hata
    cizim = ws.ProtectDrawingObjects: icerik = ws.ProtectContents
    senaryo = ws.ProtectScenarios: secim = ws.EnableSelection
    If cizim And Not ws.ProtectionMode Then
        With ws.Protection
            izin(1) = .AllowFiltering: izin(2) = .AllowSorting
            izin(3) = .AllowFormattingCells: izin(4) = .AllowFormattingColumns
            izin(5) = .AllowFormattingRows: izin(6) = .AllowInsertingColumns
            izin(7) = .AllowInsertingRows: izin(8) = .AllowInsertingHyperlinks
            izin(9) = .AllowDeletingColumns: izin(10) = .AllowDeletingRows
            izin(11) = .AllowUsingPivotTables
        End With
        ws.Unprotect Password:=""
        acildi = True
    End If
    For i = ws.Shapes.Count To 1 Step -1
        If AK12_ESKI_SEKIL(ws.Shapes(i)) Then ws.Shapes(i).Delete
    Next i
Cikis:
    If acildi Then
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
    If no <> 0 Then Err.Raise no, "AK12_KONTROL_SIL", ws.Name & ": " & metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    ' The restore path must still run, but a restore failure must not loop.
    On Error GoTo 0
    GoTo Cikis
End Sub

Public Sub AKSORA_V12_GORSEL_ONAR()
    Dim adim As String
    On Error GoTo Hata

    adim = "1 - Eski zoom kutularýný kaldýrma"
    AKSORA_V11_ESKI_KONTROL_SIL

    adim = "2 - Bütün illeri güncelle butonu"
    AKSORA_V11_ARAMA_DUGME

    adim = "3 - Ýl tarihlerini gösterme"
    AKSORA_V11_ARAMA_TARIH

    adim = "4 - Karþýlaþtýrma açýklamasý"
    AKSORA_V11_KARS_UYARI _
        ThisWorkbook.Worksheets("STOK_KARÞILAÞTIRMA")

    adim = "5 - Yüklenici raporu alt yazýsý"
    AKSORA_V12_YUK_ALT_SABITLE

    adim = "6 - Kapak tarihi"
    AKSORA_V10_TARIH_GOSTER
    Exit Sub

Hata:
    Err.Raise Err.Number, "AKSORA", _
              adim & vbCrLf & Err.Description
End Sub

Public Sub AKSORA_V12_KUR()
    Dim sc As Boolean, ev As Boolean, ws As Worksheet, adres As String
    Dim r As Long, c As Long, Hata As String
    sc = application.ScreenUpdating: ev = application.EnableEvents
    On Error GoTo Hata
    Set ws = ActiveSheet: adres = ActiveCell.Address
    r = ActiveWindow.ScrollRow: c = ActiveWindow.ScrollColumn
    application.ScreenUpdating = False: application.EnableEvents = False
    AKSORA_V12_GORSEL_ONAR
Cikis:
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Activate
        If ws.EnableSelection <> xlNoSelection Then ws.Range(adres).Select
        ActiveWindow.ScrollRow = r: ActiveWindow.ScrollColumn = c
    End If
    application.EnableEvents = ev: application.ScreenUpdating = sc
    On Error GoTo 0
    If Len(Hata) > 0 Then
        MsgBox "Görsel düzenleme tamamlanamadý: " & Hata, vbExclamation, "AKSORA"
    Else
        MsgBox "Altý düzenleme uygulandý. Stoklar ve geçmiþ kayýtlarý deðiþtirilmedi.", vbInformation, "AKSORA"
    End If
    Exit Sub
Hata:
    Hata = Err.Description
    Resume Cikis
End Sub


Private Function AK14_GECEN_SURE(ByVal bas As Double) As Double
    Dim sure As Double
    sure = Timer - bas
    If sure < 0 Then sure = sure + 86400#
    AK14_GECEN_SURE = sure
End Function

Private Sub AK14_SURE_KAYDET(ByVal il As String, ByVal sap As Double, _
                           ByVal excel As Double, ByVal basarili As Boolean)
    ' Diagnostics must never interrupt the stock transaction or its rollback.
    On Error Resume Next
    If mAK14Sureler Is Nothing Then Set mAK14Sureler = CreateObject("Scripting.Dictionary")
    If excel < 0 Then excel = 0
    mAK14Sureler(il) = Array(sap, excel, basarili)
    Debug.Print "AKSORA | " & il & " | SAP: " & AKSORA_SAYI_METNI(sap, "0.0") & _
                " sn | Excel: " & AKSORA_SAYI_METNI(excel, "0.0") & " sn | Baþarýlý: " & basarili
    On Error GoTo 0
End Sub

Public Sub AKSORA_GUNCELLEME_SURELERI()
    Dim metin As String, il As Variant, veri As Variant
    If AKSORA_V11_MESGUL Then Exit Sub
    If mAK14Sureler Is Nothing Then
        MsgBox "Bu açýlýþta henüz bütün iller güncellenmedi.", vbInformation, "AKSORA | Süreler"
        Exit Sub
    End If
    metin = "Son toplu güncelleme: " & AKSORA_SAYI_METNI(mAK14Toplam, "0.0") & " sn" & vbCrLf
    For Each il In Array("ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
        If mAK14Sureler.Exists(CStr(il)) Then
            veri = mAK14Sureler(CStr(il))
            metin = metin & vbCrLf & CStr(il) & ": SAP " & AKSORA_SAYI_METNI(veri(0), "0.0") & _
                    " sn / Excel " & AKSORA_SAYI_METNI(veri(1), "0.0") & " sn"
            If Not CBool(veri(2)) Then metin = metin & " (tamamlanmadý)"
        End If
    Next il
    metin = metin & vbCrLf & vbCrLf & "Son arayüz yenileme: " & AKSORA_SAYI_METNI(mAK14Arayuz, "0.0") & " sn" & _
            vbCrLf & "Hata olan ilde süreler kýsmi olabilir."
    MsgBox metin, vbInformation, "AKSORA | Güncelleme Süreleri"
End Sub

Public Sub AKSORA_TOPLU_SURE_TIK(Optional ByVal zorla As Boolean = False)
    Dim saniye As Long, gecen As Double, s As Shape
    If AKSORA_MUSTERI_TOPLU_AKTIF Then
        AKSORA_MUSTERI_TOPLU_TIK zorla
        Exit Sub
    End If
    If Not mTopluSaatAktif Then Exit Sub
    gecen = Timer - mTopluSaatBas
    If gecen < 0 Then gecen = gecen + 86400#
    saniye = CLng(Fix(gecen))
    If Not zorla And saniye = mTopluSonSaniye Then Exit Sub
    ' UI-only update: no event, calculation, stock or SAP state changes.
    On Error Resume Next
    Set s = ThisWorkbook.Worksheets("MALZEME_ARA").Shapes("AKSORA_TOPLU_DURUM")
    If s Is Nothing Then Exit Sub
    s.TextFrame2.TextRange.Text = mTopluDurumMetni & vbLf & _
        "Geçen süre: " & Format$(saniye \ 60, "00") & ":" & Format$(saniye Mod 60, "00")
    mTopluSonSaniye = saniye
    On Error GoTo 0
End Sub


Public Function AKSORA_MUSTERI_IL_VERI_GUNCELLE(ByVal il As String, ByRef bilgi As String) As Boolean
    AKSORA_MUSTERI_IL_VERI_GUNCELLE = AK11_IL_GUNCELLE(il, bilgi)
End Function
