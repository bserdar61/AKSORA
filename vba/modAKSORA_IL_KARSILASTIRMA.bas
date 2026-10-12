Attribute VB_Name = "modAKSORA_IL_KARSILASTIRMA"
Option Explicit
' AKSORA | STOK KARSILASTIRMA - GUVENLI IL SECIMI V8
' Mevcut tarihsel hafiza / tarih secimi / hesaplama koduna dokunmaz.
' TUMU mevcut dosyada kayitli il verilerini toplar; eksik illeri varmis gibi gostermez.
Private Const s As String = "STOK_KARÞILAÞTIRMA"
Private Const PREFIX As String = "AKSKV1_"

Public Sub AKSORA_KARS_IL_KUR()
    Dim w As Worksheet, sh As Shape, i As Long, il As Variant
    Dim eski As Boolean, olay As Boolean, sec As String, ac As String, anaBulundu As Boolean
    Set w = ThisWorkbook.Worksheets(s)
    eski = application.ScreenUpdating: olay = application.EnableEvents
    On Error GoTo Hata
    application.ScreenUpdating = False: application.EnableEvents = False
    ' Sadece bu modulun onceki cizimlerini sil.
    For i = w.Shapes.Count To 1 Step -1
        If Left$(w.Shapes(i).Name, Len(PREFIX)) = PREFIX Then w.Shapes(i).Delete
    Next i
    w.Columns("N").ColumnWidth = 4
    w.Columns("O").ColumnWidth = 18
    w.Columns("P").ColumnWidth = 18
    w.Columns("Q").ColumnWidth = 4
    w.Range("O3").ClearContents
    w.Range("O6:P7").UnMerge
    w.Range("O6:P7").Merge
    With w.Range("O6:P7")
        .Value = "KARÞILAÞTIRILACAK ÝL"
        .Interior.Color = RGB(24, 55, 88)
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = True: .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    i = 0
    For Each il In Array("TÜMÜ", "GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN")
        AKSKV1_BUTON w, CStr(il), i
        i = i + 1
    Next il
    ' Eski aciklama alanini temizle; metin artik O22:P22'de sabittir.
    w.Range("O20:P21").UnMerge
    w.Range("O20:P21").ClearContents
    w.Range("O22:P22").UnMerge
    w.Range("O22:P22").Merge
    With w.Range("O22:P22")
        .Value = "Seçili ilin kayýtlý stok geçmiþi karþýlaþtýrýlýr."
        .WrapText = False
        .Font.Name = "Aptos": .Font.Size = 9: .Font.Color = RGB(85, 103, 123)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ' ANA butonunu yeni basliktan uzakta, O3:P3'e tasir; eski makrosu korunur.
    ' ANA dugmesini metinden veya mevcut O5 konumundan bul.
    ' Sekil adi ya da TextFrame2 kullanilsa da tasima calissin.
    anaBulundu = False
    For Each sh In w.Shapes
        If Left$(sh.Name, Len(PREFIX)) <> PREFIX Then
            ac = AKSKV5_YAZI(sh)
            If ac = "ANA" Or ac = "ANA PANEL" Or (Not anaBulundu And _
                sh.Left >= w.Range("N4").Left And _
                sh.Left < w.Range("Q4").Left And _
                sh.Top >= w.Range("O4").Top And _
                sh.Top < w.Range("O6").Top And _
                sh.Width < w.Range("O6:P7").Width / 2) Then
                If Not anaBulundu Then
                    With sh
                        .Placement = xlFreeFloating
                        .LockAspectRatio = msoFalse
                        .Left = w.Range("O3").Left + 2
                        .Top = w.Range("O3").Top + 1
                        .Width = w.Range("O3").Width - 4
                        .Height = w.Range("O3").Height - 2
                        .ZOrder msoBringToFront
                    End With
                    anaBulundu = True
                End If
            ElseIf ac = "KARSILASTIR" Then
                If Len(CStr(w.Range("AA3").Value2)) = 0 Then
                    If Len(sh.OnAction) > 0 And InStr(1, sh.OnAction, "AKSKV1_", vbTextCompare) = 0 Then
                        w.Range("AA3").Value = sh.OnAction
                    End If
                End If
                If Len(CStr(w.Range("AA3").Value2)) > 0 Then
                    sh.OnAction = "'" & ThisWorkbook.Name & "'!AKSKV1_KARSILASTIR"
                End If
            End If
        End If
    Next sh
    sec = AKSKV1_NORM(CStr(w.Range("AA2").Value2))
    If Not AKSKV1_GECERLI(sec) Then sec = AKSKV1_NORM(CStr(ThisWorkbook.Worksheets("KAPAK").Range("AZ1").Value2))
    If Not AKSKV1_GECERLI(sec) Then sec = "ARTVIN"
    w.Range("AA2").Value = sec
    w.Columns("AA").Hidden = True
    AKFIX_KARS_TARIHLER
    AKSKV3_OZET w, sec
    AKSKV8_FILTRELERI_GERI_YUKLE w
    AKSKV6_FARK_RENGI w
    w.Activate
    AKSKV5_SABITLE w, sec
    ActiveWindow.DisplayGridlines = False
Cikis:
    application.EnableEvents = olay: application.ScreenUpdating = eski
    Exit Sub
Hata:
    MsgBox "Karþýlaþtýrma il paneli: " & Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Sub AKSKV1_BUTON(ByVal w As Worksheet, ByVal etiket As String, ByVal sira As Long)
    Dim r As Range, sh As Shape, kod As String
    Set r = w.Range("O8:P9")
    kod = AKSKV1_NORM(etiket)
    Set sh = w.Shapes.AddShape(msoShapeRoundedRectangle, r.Left + 2, r.Top + sira * 46, r.Width - 4, 40)
    sh.Name = PREFIX & kod
    sh.TextFrame.Characters.Text = etiket
    sh.TextFrame.HorizontalAlignment = xlHAlignCenter
    sh.TextFrame.VerticalAlignment = xlVAlignCenter
    sh.TextFrame.Characters.Font.Name = "Aptos"
    sh.TextFrame.Characters.Font.Size = 10
    sh.TextFrame.Characters.Font.bold = True
    sh.line.Visible = msoFalse
    sh.Placement = xlFreeFloating
    sh.LockAspectRatio = msoFalse
    sh.OnAction = "'" & ThisWorkbook.Name & "'!AKSKV1_SEC_" & kod
End Sub

Private Sub AKSKV1_RENKLE(ByVal w As Worksheet, ByVal sec As String)
    Dim sh As Shape, kod As String
    For Each sh In w.Shapes
        If Left$(sh.Name, Len(PREFIX)) = PREFIX Then
            sh.Placement = xlFreeFloating
            kod = Mid$(sh.Name, Len(PREFIX) + 1)
            If kod = sec Then
                sh.Fill.ForeColor.RGB = RGB(0, 121, 103)
                sh.TextFrame.Characters.Font.Color = vbWhite
            Else
                sh.Fill.ForeColor.RGB = RGB(230, 240, 248)
                sh.TextFrame.Characters.Font.Color = RGB(24, 55, 88)
            End If
        End If
    Next sh
End Sub
Private Sub AKSKV1_SEC(ByVal sec As String)
    Dim w As Worksheet
    On Error GoTo Hata
    Set w = ThisWorkbook.Worksheets(s)
    w.Range("AA2").Value = sec
    AKFIX_KARS_TARIHLER
    AKSKV5_SABITLE w, sec
    AKSKV6_FARK_RENGI w
    Exit Sub
Hata:
    MsgBox "Ýl seçimi: " & Err.Description, vbExclamation, "AKSORA"
End Sub

' Butona basýnca yalnýzca renklendirme deðiþir; hiçbir butonun boyutu deðiþmez.
' TÜMÜ, dosyada bulunan ilin iki seçili tarihteki gerçek toplamlarýný gösterir.
Private Sub AKSKV3_OZET(ByVal w As Worksheet, ByVal sec As String)
    Dim aktif As String
    aktif = AKSKV1_NORM(CStr(ThisWorkbook.Worksheets("KAPAK").Range("AZ1").Value2))
    ' Aciklama secime gore yer degistirmez.
    w.Range("O22").Value = "Seçili ilin kayýtlý stok geçmiþi karþýlaþtýrýlýr."
    AKSKV2_TOPLAM_GUNCELLE w
End Sub

Public Sub AKSKV1_SEC_TUMU(): AKSKV1_SEC "TUMU": End Sub
Public Sub AKSKV1_SEC_GIRESUN(): AKSKV1_SEC "GIRESUN": End Sub
Public Sub AKSKV1_SEC_TRABZON(): AKSKV1_SEC "TRABZON": End Sub
Public Sub AKSKV1_SEC_RIZE(): AKSKV1_SEC "RIZE": End Sub
Public Sub AKSKV1_SEC_GUMUSHANE(): AKSKV1_SEC "GUMUSHANE": End Sub
Public Sub AKSKV1_SEC_ARTVIN(): AKSKV1_SEC "ARTVIN": End Sub

Public Sub AKSKV1_KARSILASTIR()
    Dim w As Worksheet, sec As String
    On Error GoTo Hata
    Set w = ThisWorkbook.Worksheets(s)
    sec = AKFIX_KARS_IL()
    STOK_KARSILASTIR
    AKSKV5_SABITLE w, sec
    AKSKV6_FARK_RENGI w
    AKFIX_EN_USTE w
    Exit Sub
Hata:
    MsgBox "Karsilastirma: " & Err.Description, vbExclamation, "AKSORA"
End Sub

' Secim sonrasinda Excel sekil secimini kaldirir ve buton renklerini yeniden uygular.
Private Sub AKSKV5_SABITLE(ByVal w As Worksheet, ByVal sec As String)
    Dim sh As Shape
    AKSKV1_RENKLE w, sec
    For Each sh In w.Shapes
        If Left$(sh.Name, Len(PREFIX)) = PREFIX Then
            sh.Placement = xlFreeFloating
            sh.LockAspectRatio = msoFalse
            sh.line.Visible = msoFalse
        End If
    Next sh
    AKFIX_EN_USTE w
    AKSKV1_RENKLE w, sec
    AKSORA_KARS_ALT_YAZI_SABITLE
End Sub

Private Function AKSKV5_YAZI(ByVal sh As Shape) As String
    Dim t As String
    On Error Resume Next
    t = Trim$(sh.TextFrame2.TextRange.Text)
    If Len(t) = 0 Then t = Trim$(sh.TextFrame.Characters.Text)
    On Error GoTo 0
    AKSKV5_YAZI = AKSKV1_NORM(t)
End Function

Private Function AKSKV1_GECERLI(ByVal s As String) As Boolean
    Select Case s
        Case "TUMU", "GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN": AKSKV1_GECERLI = True
    End Select
End Function
Private Function AKSKV1_NORM(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, "Ý", "I"): s = Replace(s, "Ü", "U")
    s = Replace(s, "Þ", "S"): s = Replace(s, "Ð", "G")
    s = Replace(s, "Ç", "C"): s = Replace(s, "Ö", "O")
    AKSKV1_NORM = s
End Function

' V2: Secilen iki snapshot'in BUTUN normal stok degerleri.
' Degisen satirlarin toplamini degil, hafiza indeksindeki E sutununu kullanir.
Public Sub AKSKV2_TOPLAM_GUNCELLE(ByVal w As Worksheet)
    ' Son karsilastirma toplamlarini il butonuna tiklayinca yeniden etiketleme.
    ' Tarihler ve toplamlar ilgili ilin kendi gecmisinden yuklenir.
End Sub
Private Sub AKSKV8_FILTRELERI_GERI_YUKLE(ByVal w As Worksheet)
    Dim adlar As Variant, makrolar As Variant, i As Long, sh As Shape
    adlar = Array("btnKarsTum", "btnKarsAzalan", "btnKarsArtan", "btnKarsYeni", "btnKarsSifir")
    makrolar = Array("STOK_KARS_FILTRE_TUMU", "STOK_KARS_FILTRE_AZALAN", _
                     "STOK_KARS_FILTRE_ARTAN", "STOK_KARS_FILTRE_YENI", "STOK_KARS_FILTRE_SIFIR")
    For i = LBound(adlar) To UBound(adlar)
        Set sh = Nothing
        On Error Resume Next
        Set sh = w.Shapes(CStr(adlar(i)))
        On Error GoTo 0
        If Not sh Is Nothing Then
            sh.OnAction = "'" & ThisWorkbook.Name & "'!" & CStr(makrolar(i))
        End If
    Next i
    w.Range("AB2:AC6").ClearContents
End Sub

' Formul kullanmayan iki sayisal kosul. Eski filtre makrolari
' font rengini degistirse bile kosullu bicimlendirme gorunur kalir.
Private Sub AKSKV6_FARK_RENGI(ByVal w As Worksheet)
    Dim c As Range, k As FormatCondition
    Set c = w.Range("G12")
    On Error GoTo Hata
    c.FormatConditions.Delete
    c.Font.bold = True
    c.Font.Color = RGB(24, 55, 88)
    Set k = c.FormatConditions.Add(Type:=xlCellValue, Operator:=xlLess, Formula1:="0")
    k.Font.Color = RGB(192, 0, 0)
    Set k = c.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="0")
    k.Font.Color = RGB(0, 140, 75)
    Exit Sub
Hata:
    ' Bicimlendirme Excel tarafindan reddedilirse filtreleri bozma.
    On Error Resume Next
    If IsNumeric(c.Value2) Then
        If CDbl(c.Value2) < 0 Then
            c.Font.Color = RGB(192, 0, 0)
        ElseIf CDbl(c.Value2) > 0 Then
            c.Font.Color = RGB(0, 140, 75)
        End If
    End If
End Sub

Public Sub AKSORA_KARS_ALT_YAZI_SABITLE()

    Dim ws As Worksheet
    Dim buton As Shape, yazi As Shape
    Dim hataNo As Long, hataMetni As String
    Dim icerik As Boolean, senaryo As Boolean

    On Error GoTo Hata

    Set ws = ThisWorkbook.Worksheets("STOK_KARÞILAÞTIRMA")
    Set buton = ws.Shapes("AKSKV1_ARTVIN")

    icerik = ws.ProtectContents
    senaryo = ws.ProtectScenarios

    ws.Unprotect Password:=""

    ws.Range("O22:P22").ClearContents

    On Error Resume Next
    Set yazi = ws.Shapes("AKSORA_KARS_SABIT_ACIKLAMA")
    On Error GoTo Hata

    If yazi Is Nothing Then
        Set yazi = ws.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, _
            buton.Left, buton.Top + buton.Height + 4, _
            buton.Width, 22)
    End If

    With yazi
        .Name = "AKSORA_KARS_SABIT_ACIKLAMA"
        .Placement = xlFreeFloating
        .Left = buton.Left
        .Top = buton.Top + buton.Height + 4
        .Width = buton.Width
        .Height = 22
        .Visible = msoTrue
        .Locked = True
        .Fill.Visible = msoFalse
        .line.Visible = msoFalse
        .OnAction = vbNullString

        With .TextFrame2
            .AutoSize = msoAutoSizeNone
            .WordWrap = msoFalse
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 0
            .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = "Seçili ilin kayýtlý stok geçmiþi karþýlaþtýrýlýr."
            .TextRange.Font.Name = "Aptos"
            .TextRange.Font.Size = 8
            .TextRange.Font.Fill.ForeColor.RGB = RGB(85, 103, 123)
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With

        .ZOrder msoBringToFront
    End With

    ws.Protect Password:="", DrawingObjects:=True, _
        Contents:=icerik, Scenarios:=senaryo, _
        UserInterfaceOnly:=True, AllowFiltering:=True, _
        AllowSorting:=True

    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description

    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=icerik, Scenarios:=senaryo, _
            UserInterfaceOnly:=True, AllowFiltering:=True, _
            AllowSorting:=True
    End If
    On Error GoTo 0

    Err.Raise hataNo, "AKSORA_KARS_ALT_YAZI_SABITLE", hataMetni

End Sub

