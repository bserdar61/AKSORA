Attribute VB_Name = "modAKSORA_BIRIM"
Option Explicit

Private Const HUB As String = "BIRIM SECIM"
Private Const YAT As String = "YATIRIM PANEL"
Private Const MUS As String = "MUSTERI PANEL"
Private Const MUK As String = "MUSTERI IL SECIM"
Private Const YAK As String = "YATIRIM IL SECIM"
Private Const tag As String = "AKSORA_BIRIM_V1"
Private Const pw As String = ""
Private Const cw As Double = 1080
Private Const CH As Double = 660
Private mYonlendirme As Boolean
Private mGorunum As Boolean
Private Const Y_ETIKET As String = "AKSORA_YATIRIM_V1"

' Yalniz yeni sayfalari ve ThisWorkbook acilis yonlendirmesini kurar.
Public Sub AKSORA_BIRIM_KUR()
    Dim cM As Object, eskiKod As String, yeniKod As String
    Dim ev As Boolean, sc As Boolean, degisti As Boolean
    Dim yedek As String, Hata As String, geri As String
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then
        Err.Raise vbObjectError + 1600, , "Devam eden islemin bitmesini bekleyin."
    End If
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 1601, , "Once dosyayi kaydedin."
    Set cM = ThisWorkbook.VBProject.VBComponents(ThisWorkbook.CodeName).CodeModule
    eskiKod = cM.Lines(1, cM.CountOfLines)
    yeniKod = BIRIM_ENTEGRE_KOD(eskiKod)
    BIRIM_SAYFA_KONTROL HUB
    BIRIM_SAYFA_KONTROL YAT
    BIRIM_SAYFA_KONTROL MUS
    BIRIM_SAYFA_KONTROL MUK
    BIRIM_SAYFA_KONTROL YAK
    yedek = ThisWorkbook.Path & application.PathSeparator & _
        "AKSORA_BIRIM_ONCESI_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    If Len(Dir$(yedek)) > 0 Then Err.Raise vbObjectError + 1602, , "Yedek dosya zaten var. Tekrar deneyin."
    ThisWorkbook.SaveCopyAs yedek
    application.EnableEvents = False: application.ScreenUpdating = False
    BIRIM_HAZIRLA
    AKSORA_BIRIM_KAPAK_GERI
    If yeniKod <> eskiKod Then
        degisti = True
        cM.DeleteLines 1, cM.CountOfLines
        cM.AddFromString yeniKod
    End If
    AKSORA_BIRIM_ACILIS
    ThisWorkbook.Save
Cikis:
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If Len(Hata) > 0 Then
        MsgBox "Kurulum tamamlanamadi: " & Hata & geri & vbCrLf & _
            "Olusturulan yedek: " & yedek, vbExclamation, "AKSORA"
    Else
        MsgBox "Birim acilis ekrani hazir. Dosyayi kapatip yeniden acin.", vbInformation, "AKSORA"
    End If
    Exit Sub
Hata:
    Hata = Err.Description
    On Error Resume Next
    If degisti Then
        Err.Clear
        cM.DeleteLines 1, cM.CountOfLines
        cM.AddFromString eskiKod
        If Err.Number <> 0 Then geri = vbCrLf & "Acilis kodu geri alinamadi; yedegi kullanin."
    End If
    mYonlendirme = False
    On Error GoTo 0
    GoTo Cikis
End Sub

Private Function BIRIM_ENTEGRE_KOD(ByVal kod As String) As String
    If InStr(1, kod, "' AKSORA_BIRIM_HOOK_OPEN", vbBinaryCompare) > 0 Then
        If InStr(kod, "' AKSORA_BIRIM_HOOK_ACTIVATE") = 0 Or _
           InStr(kod, "' AKSORA_BIRIM_HOOK_SHEET") = 0 Or _
           InStr(kod, "' AKSORA_BIRIM_HOOK_RESIZE") = 0 Or _
           InStr(kod, "' AKSORA_BIRIM_HOOK_WINDOW") = 0 Then
            Err.Raise vbObjectError + 1603, , "Birim entegrasyonu eksik. Yedekten geri donun."
        End If
        BIRIM_ENTEGRE_KOD = kod
        Exit Function
    End If
    kod = BIRIM_KOD_EKLE(kod, "Workbook_Open", _
        "    SPLASH_DURUM 1, ""AKSORA hazýr""", _
        "    ' AKSORA_BIRIM_HOOK_OPEN" & vbCrLf & "    AKSORA_BIRIM_ACILIS", True)
    kod = BIRIM_KOD_EKLE(kod, "Workbook_Activate", _
        "    UYGULAMA_MODU_AC", _
        "    ' AKSORA_BIRIM_HOOK_ACTIVATE" & vbCrLf & _
        "    If AKSORA_BIRIM_AKTIF_GORUNUM() Then Exit Sub", False)
    kod = BIRIM_KOD_EKLE(kod, "Workbook_SheetActivate", _
        "    If AKSORA_V13_ACILIS _", _
        "    ' AKSORA_BIRIM_HOOK_SHEET" & vbCrLf & _
        "    If AKSORA_BIRIM_OLAY(Sh) Then Exit Sub", True)
    kod = BIRIM_KOD_EKLE(kod, "Workbook_WindowResize", _
        "    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub", _
        "    ' AKSORA_BIRIM_HOOK_RESIZE" & vbCrLf & _
        "    If AKSORA_BIRIM_AKTIF_GORUNUM() Then Exit Sub", False)
    kod = BIRIM_KOD_EKLE(kod, "Workbook_WindowActivate", _
        "    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub", _
        "    ' AKSORA_BIRIM_HOOK_WINDOW" & vbCrLf & _
        "    If AKSORA_BIRIM_AKTIF_GORUNUM() Then Exit Sub", False)
    BIRIM_ENTEGRE_KOD = kod
End Function

Private Function BIRIM_KOD_EKLE(ByVal kod As String, ByVal olay As String, _
    ByVal hedef As String, ByVal ek As String, ByVal once As Boolean) As String
    Dim bas As Long, son As Long, parca As String, yer As Long
    bas = InStr(1, kod, "Private Sub " & olay & "(", vbBinaryCompare)
    If bas = 0 Then Err.Raise vbObjectError + 1604, , "Olay bulunamadi: " & olay
    son = InStr(bas, kod, "End Sub", vbBinaryCompare)
    If son = 0 Then Err.Raise vbObjectError + 1605, , "Olay sonu bulunamadi: " & olay
    parca = Mid$(kod, bas, son + Len("End Sub") - bas)
    yer = InStr(1, parca, hedef, vbBinaryCompare)
    If yer = 0 Or InStr(yer + Len(hedef), parca, hedef, vbBinaryCompare) > 0 Then
        Err.Raise vbObjectError + 1606, , "Acilis kodu beklenen surumle uyusmuyor: " & olay
    End If
    If once Then
        parca = Replace(parca, hedef, ek & vbCrLf & hedef, 1, 1, vbBinaryCompare)
    Else
        parca = Replace(parca, hedef, hedef & vbCrLf & ek, 1, 1, vbBinaryCompare)
    End If
    BIRIM_KOD_EKLE = Left$(kod, bas - 1) & parca & Mid$(kod, son + Len("End Sub"))
End Function

Private Sub BIRIM_SAYFA_KONTROL(ByVal ad As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
    If Not ws Is Nothing Then
        If CStr(ws.Range("BX1").Value2) <> tag Then
            Err.Raise vbObjectError + 1607, , "Ayni isimde baska bir sayfa var: " & ad
        End If
    End If
End Sub

Private Function BIRIM_SAYFA(ByVal ad As String) As Worksheet
    Dim ws As Worksheet
    BIRIM_SAYFA_KONTROL ad
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ad
        ws.Range("BX1").Value2 = tag
    End If
    Set BIRIM_SAYFA = ws
End Function

Private Sub BIRIM_HAZIRLA()
    Dim ws As Worksheet, ad As Variant
    Set ws = BIRIM_SAYFA(HUB)
    BIRIM_CIZ ws, ""
    For Each ad In Array(YAT, MUS)
        Set ws = BIRIM_SAYFA(CStr(ad))
        If CStr(ad) = YAT Then BIRIM_CIZ ws, "Yatýrým Birimi" Else BIRIM_CIZ ws, "Müþteri Birimi"
        ws.Visible = xlSheetVeryHidden
    Next ad
End Sub

Public Sub AKSORA_BIRIM_ACILIS()
    Dim ws As Worksheet, diger As Worksheet, ev As Boolean, sc As Boolean
    Dim no As Long, metin As String
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    Sayfa12.STOK_KARS_FILTRE_TAKIP_DURDUR
    If AKSORA_GIRIS_ACIK Then AKSORA_KAPAK_KILITLE
    Set ws = BIRIM_SAYFA(HUB)
    If ws.Shapes.Count = 0 Or Val(CStr(ws.Range("BX2").Value2)) < 8 Then BIRIM_CIZ ws, ""
    ws.Visible = xlSheetVisible
    ws.Move Before:=ThisWorkbook.Worksheets(1)
    ws.Select Replace:=True
    ' Birim secilmeden yalniz ana secim sayfasi gorunur.
    For Each diger In ThisWorkbook.Worksheets
        If diger.Name <> HUB And diger.Visible = xlSheetVisible Then diger.Visible = xlSheetVeryHidden
    Next diger
    BIRIM_GORUNUM ws
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then Err.Raise no, "AKSORA_BIRIM_ACILIS", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Public Function AKSORA_BIRIM_OLAY(ByVal sh As Object) As Boolean
    If mYonlendirme Then AKSORA_BIRIM_OLAY = True: Exit Function
    If Not TypeOf sh Is Worksheet Then Exit Function
    If AKSORA_MUSTERI_GECIS_KORU(sh) Then AKSORA_BIRIM_OLAY = True: Exit Function
    On Error Resume Next
    application.MoveAfterReturn = (sh.Name <> "M_MALZEME_ARA")
    On Error GoTo 0
    If AKSORA_MUSTERI_TEMA_SAYFA(sh) Then
        AKSORA_MUSTERI_TEMA_GORUNUM sh
        AKSORA_MUSTERI_SECIM_BASA sh
        AKSORA_BIRIM_OLAY = True
        Exit Function
    End If
    If AKSORA_YATIRIM_SAYFA(sh) Then
        AKSORA_YATIRIM_GORUNUM sh
        AKSORA_BIRIM_OLAY = True
        Exit Function
    End If
    If Not BIRIM_YENI_SAYFA(CStr(sh.Name)) Then Exit Function
    If sh.Name = MUK Or sh.Name = YAK Then MUSTERI_KAPAK_GORUNUM sh Else BIRIM_GORUNUM sh
    AKSORA_BIRIM_OLAY = True
End Function

Public Function AKSORA_BIRIM_AKTIF_GORUNUM() As Boolean
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Function
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Function
    If AKSORA_MUSTERI_TEMA_SAYFA(ActiveSheet) Then
        AKSORA_MUSTERI_TEMA_GORUNUM ActiveSheet
        AKSORA_BIRIM_AKTIF_GORUNUM = True
        Exit Function
    End If
    If AKSORA_YATIRIM_SAYFA(ActiveSheet) Then
        AKSORA_YATIRIM_GORUNUM ActiveSheet
        AKSORA_BIRIM_AKTIF_GORUNUM = True
        Exit Function
    End If
    If Not BIRIM_YENI_SAYFA(CStr(ActiveSheet.Name)) Then Exit Function
    If ActiveSheet.Name = MUK Or ActiveSheet.Name = YAK Then MUSTERI_KAPAK_GORUNUM ActiveSheet Else BIRIM_GORUNUM ActiveSheet
    AKSORA_BIRIM_AKTIF_GORUNUM = True
End Function

Private Function BIRIM_YENI_SAYFA(ByVal ad As String) As Boolean
    BIRIM_YENI_SAYFA = (ad = HUB Or ad = YAT Or ad = MUS Or ad = MUK Or ad = YAK)
End Function

Public Sub AKSORA_BIRIM_MALZEME()
    Dim ev As Boolean, sc As Boolean, no As Long, metin As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    With ThisWorkbook.Worksheets("KAPAK")
        .Visible = xlSheetVisible
        .Select Replace:=True
    End With
    AKSORA_KAPAK_KORU
    AKSORA_V8_KAPAK_GORUNUM
    AKSORA_BIRIM_KAPAK_GERI
    BIRIM_DIGERLERINI_GIZLE "KAPAK"
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_BIRIM_YATIRIM()
    On Error GoTo Hata
    YATIRIM_KAPAK_AC
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Yatýrým birimi", Err.Number, Err.Description
End Sub
Public Sub AKSORA_BIRIM_MUSTERI()
    On Error GoTo Hata
    MUSTERI_KAPAK_AC
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Müþteri birimi", Err.Number, Err.Description
End Sub
Private Sub MUSTERI_KAPAK_AC()
    Dim ws As Worksheet, kaynak As Worksheet, s As Shape, foto As Shape, bant As Shape
    Dim yol As String, ev As Boolean, sc As Boolean, no As Long, metin As String
    Dim i As Long, h As Double, kaynakGorunum As XlSheetVisibility
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_MUSTERI_BANT.png"
    If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1610, , _
        "AKSORA_MUSTERI_BANT.png Excel dosyasinin yaninda bulunmalidir."
    BIRIM_SAYFA_KONTROL MUK
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(MUK)
    On Error GoTo Hata
    If ws Is Nothing Then
        Set kaynak = ThisWorkbook.Worksheets("KAPAK")
        kaynakGorunum = kaynak.Visible
        kaynak.Visible = xlSheetVisible
        kaynak.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set ws = ActiveSheet
        ws.Name = MUK
        kaynak.Visible = kaynakGorunum
        ws.Unprotect pw
        ws.Range("BX1").Value2 = tag
        ws.Range("AZ1").ClearContents
        ' Kopyadaki eski stok yonlendirmeleri kesinlikle kullanilmaz.
        For Each s In ws.Shapes
            s.OnAction = ""
        Next s
        For i = 1 To 5
            With ws.Shapes("AKSORA_V4_IL_" & i)
                .AlternativeText = CStr(i)
                .OnAction = BIRIM_MAKRO("AKSORA_MUSTERI_IL_SEC")
            End With
        Next i
        ws.Shapes("AKSORA_V4_ENTER").OnAction = BIRIM_MAKRO("AKSORA_MUSTERI_GIRIS")
        AKSORA_BIRIM_KAPAK_GERI MUK
        ws.Range("BX2").Value2 = 0
    End If
    ws.Unprotect pw
    If Val(CStr(ws.Range("BX2").Value2)) < 7 Then
        On Error Resume Next
        ws.Shapes("BIRIM_MUSTERI_BANT").Delete
        On Error GoTo Hata
        Set foto = ws.Shapes("AKSORA_KAPAK_YENI_V6")
        ' Ust fotograf korunur; sadece iki alt bilgi seridi kaplanir.
        Set bant = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
        h = bant.Height
        bant.PictureFormat.CropTop = h * 0.7
        bant.LockAspectRatio = msoFalse
        bant.Left = foto.Left: bant.Top = foto.Top + foto.Height * 0.7
        bant.Width = foto.Width: bant.Height = foto.Height * 0.3
        bant.Name = "BIRIM_MUSTERI_BANT"
        bant.Placement = xlFreeFloating: bant.Locked = True
        ws.Range("BX2").Value2 = 7
    End If
    MUSTERI_SECIM_CIZ ws
    ws.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
    ws.Visible = xlSheetVisible: ws.Select Replace:=True
    BIRIM_DIGERLERINI_GIZLE MUK
    ThisWorkbook.Worksheets("KAPAK").Visible = xlSheetVeryHidden
    ' Il secim ekrani tek gorunur sayfa: arkada acik kalan musteri sayfalarina kayma olmaz.
    Dim k As Worksheet
    For Each k In ThisWorkbook.Worksheets
        If k.Name <> MUK And k.Visible = xlSheetVisible Then k.Visible = xlSheetVeryHidden
    Next k
    MUSTERI_KAPAK_GORUNUM ws
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=pw, UserInterfaceOnly:=True
    If Not kaynak Is Nothing Then kaynak.Visible = kaynakGorunum
    On Error GoTo 0
    Resume Cikis
End Sub

Private Sub MUSTERI_KAPAK_GORUNUM(ByVal ws As Worksheet)
    Dim panel As Shape, z As Double
    Set panel = ws.Shapes("AKSORA_V4_PANEL")
    With ThisWorkbook.Windows(1)
        .DisplayGridlines = False: .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False: .DisplayVerticalScrollBar = False
        .ScrollRow = 1: .ScrollColumn = 1
        z = application.Min((.UsableWidth - 4) / (panel.Left + panel.Width), _
            (.UsableHeight - 4) / (panel.Top + panel.Height)) * 100
        .Zoom = application.Max(10, application.Min(200, Fix(z)))
    End With
End Sub

Public Sub AKSORA_MUSTERI_IL_SEC()
    Dim ws As Worksheet, ix As Long, keys As Variant, ev As Boolean, sc As Boolean, metin As String
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name <> MUK Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(MUK)
    ix = CLng(ws.Shapes(CStr(application.Caller)).AlternativeText)
    If ix < 1 Or ix > 5 Then Exit Sub
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    ' Musteri depo tanimi olmayan il secilemez (aktif degil).
    If Not AKSORA_MUSTERI_IL_DESTEK(CStr(keys(ix - 1))) Then
        MsgBox Choose(ix, "GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN") & _
            " müþteri birimi þu an aktif deðil.", vbInformation, "AKSORA"
        Exit Sub
    End If
    ' Secim tek seferde cizilir; hesaplama olaylari baska sayfaya gecis yaptirmaz.
    application.ScreenUpdating = False: application.EnableEvents = False
    ws.Unprotect pw
    ws.Range("AZ1").Value2 = keys(ix - 1)
    MUSTERI_SECIM_CIZ ws
Cikis:
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect Password:=pw, UserInterfaceOnly:=True
        If Not (ActiveSheet Is ws) Then ws.Activate
    End If
    application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If Len(metin) > 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    metin = Err.Description
    Resume Cikis
End Sub

Private Sub MUSTERI_SECIM_CIZ(ByVal ws As Worksheet)
    Dim keys As Variant, names As Variant, i As Long, key As String
    Dim t As Date, yazi As String
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    names = Array("GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN")
    key = CStr(ws.Range("AZ1").Value2)
    ws.Shapes("AKSORA_V4_STATUS").TextFrame2.TextRange.Text = "SEÇÝLÝ ÝL: -"
    ' Secili ilin musteri stoklarinin son guncelleme tarihi.
    On Error Resume Next
    If ws.Name = MUK Then
        If Len(key) = 0 Then
            yazi = "Son stok güncellemesi: -"
        ElseIf Not AKSORA_MUSTERI_IL_DESTEK(key) Then
            yazi = "Bu il için müþteri birimi þu an aktif deðil"
        Else
            t = AKSORA_MUSTERI_IL_SON_TARIH(key)
            If t > 0 Then yazi = "Son stok güncellemesi: " & Format$(t, "dd.mm.yyyy hh:nn") Else yazi = "Son stok güncellemesi: Henüz güncellenmedi"
        End If
        ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = yazi
    End If
    On Error GoTo 0
    For i = 1 To 5
        With ws.Shapes("AKSORA_V4_IL_" & i)
            If keys(i - 1) = key Then
                .Fill.ForeColor.RGB = RGB(32, 78, 76)
                .line.ForeColor.RGB = RGB(115, 223, 157): .line.Weight = 2.5
                ws.Shapes("AKSORA_V4_STATUS").TextFrame2.TextRange.Text = "SEÇÝLÝ ÝL: " & names(i - 1)
            Else
                .Fill.ForeColor.RGB = RGB(40, 79, 111)
                .line.ForeColor.RGB = RGB(83, 133, 168): .line.Weight = 1
            End If
        End With
    Next i
End Sub

Public Sub AKSORA_MUSTERI_GIRIS()
    Dim ws As Worksheet
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(MUK)
    If Len(CStr(ws.Range("AZ1").Value2)) = 0 Then
        MsgBox "Lütfen bir il seçiniz.", vbInformation, "AKSORA"
    ElseIf Not AKSORA_MUSTERI_IL_DESTEK(CStr(ws.Range("AZ1").Value2)) Then
        MsgBox "Seçilen ilin müþteri birimi þu an aktif deðil. Aktif iller: Artvin, Trabzon, Gümüþhane, Giresun, Rize.", vbInformation, "AKSORA"
    Else
        ' Secilen ilin stoklari ekrana yuklenir (onceki ilin verisi kendi kaydina saklanir).
        AKSORA_MUSTERI_IL_DEGIS CStr(ws.Range("AZ1").Value2)
        AKSORA_MUSTERI_TEMA_AC
    End If
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub YATIRIM_KAPAK_AC()
    Dim ws As Worksheet, kaynak As Worksheet, s As Shape, foto As Shape, bant As Shape
    Dim yol As String, ev As Boolean, sc As Boolean, no As Long, metin As String
    Dim i As Long, h As Double, kaynakGorunum As XlSheetVisibility
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_YATIRIM_BANT.png"
    If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1610, , _
        "AKSORA_YATIRIM_BANT.png Excel dosyasinin yaninda bulunmalidir."
    BIRIM_SAYFA_KONTROL YAK
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(YAK)
    On Error GoTo Hata
    If ws Is Nothing Then
        Set kaynak = ThisWorkbook.Worksheets("KAPAK")
        kaynakGorunum = kaynak.Visible
        kaynak.Visible = xlSheetVisible
        kaynak.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set ws = ActiveSheet
        ws.Name = YAK
        kaynak.Visible = kaynakGorunum
        ws.Unprotect pw
        ws.Range("BX1").Value2 = tag
        ws.Range("AZ1").ClearContents
        ' Kopyadaki eski stok yonlendirmeleri kesinlikle kullanilmaz.
        For Each s In ws.Shapes
            s.OnAction = ""
        Next s
        For i = 1 To 5
            With ws.Shapes("AKSORA_V4_IL_" & i)
                .AlternativeText = CStr(i)
                .OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_IL_SEC")
            End With
        Next i
        ws.Shapes("AKSORA_V4_ENTER").OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_GIRIS")
        AKSORA_BIRIM_KAPAK_GERI YAK
        ws.Range("BX2").Value2 = 0
    End If
    ws.Unprotect pw
    If Val(CStr(ws.Range("BX2").Value2)) < 7 Then
        On Error Resume Next
        ws.Shapes("BIRIM_YATIRIM_BANT").Delete
        On Error GoTo Hata
        Set foto = ws.Shapes("AKSORA_KAPAK_YENI_V6")
        ' Ust fotograf korunur; sadece iki alt bilgi seridi kaplanir.
        Set bant = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
        h = bant.Height
        bant.PictureFormat.CropTop = h * 0.7
        bant.LockAspectRatio = msoFalse
        bant.Left = foto.Left: bant.Top = foto.Top + foto.Height * 0.7
        bant.Width = foto.Width: bant.Height = foto.Height * 0.3
        bant.Name = "BIRIM_YATIRIM_BANT"
        bant.Placement = xlFreeFloating: bant.Locked = True
        ws.Range("BX2").Value2 = 7
    End If
    YATIRIM_SECIM_CIZ ws
    ws.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
    ws.Visible = xlSheetVisible: ws.Select Replace:=True
    BIRIM_DIGERLERINI_GIZLE YAK
    ThisWorkbook.Worksheets("KAPAK").Visible = xlSheetVeryHidden
    ' Il secim ekrani tek gorunur sayfa: arkada acik kalan sayfalara kayma olmaz.
    Dim yk As Worksheet
    For Each yk In ThisWorkbook.Worksheets
        If yk.Name <> YAK And yk.Visible = xlSheetVisible Then yk.Visible = xlSheetVeryHidden
    Next yk
    YATIRIM_KAPAK_GORUNUM ws
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=pw, UserInterfaceOnly:=True
    If Not kaynak Is Nothing Then kaynak.Visible = kaynakGorunum
    On Error GoTo 0
    Resume Cikis
End Sub

Private Sub YATIRIM_KAPAK_GORUNUM(ByVal ws As Worksheet)
    Dim panel As Shape, z As Double
    Set panel = ws.Shapes("AKSORA_V4_PANEL")
    With ThisWorkbook.Windows(1)
        .DisplayGridlines = False: .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False: .DisplayVerticalScrollBar = False
        .ScrollRow = 1: .ScrollColumn = 1
        z = application.Min((.UsableWidth - 4) / (panel.Left + panel.Width), _
            (.UsableHeight - 4) / (panel.Top + panel.Height)) * 100
        .Zoom = application.Max(10, application.Min(200, Fix(z)))
    End With
End Sub

Public Sub AKSORA_YATIRIM_IL_SEC()
    Dim ws As Worksheet, ix As Long, keys As Variant, ev As Boolean, metin As String
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name <> YAK Then Exit Sub
    ev = application.EnableEvents
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(YAK)
    ix = CLng(ws.Shapes(CStr(application.Caller)).AlternativeText)
    If ix < 1 Or ix > 5 Then Exit Sub
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    ' Secim tek seferde cizilir; olaylar baska sayfaya gecis yaptirmaz.
    application.ScreenUpdating = False: application.EnableEvents = False
    ws.Unprotect pw
    ws.Range("AZ1").Value2 = keys(ix - 1)
    YATIRIM_SECIM_CIZ ws
Cikis:
    On Error Resume Next
    If Not ws Is Nothing Then
        ws.Protect Password:=pw, UserInterfaceOnly:=True
        If Not (ActiveSheet Is ws) Then ws.Activate
    End If
    application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If Len(metin) > 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    metin = Err.Description
    Resume Cikis
End Sub

Private Sub YATIRIM_SECIM_CIZ(ByVal ws As Worksheet)
    Dim keys As Variant, names As Variant, i As Long, key As String
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    names = Array("GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN")
    key = CStr(ws.Range("AZ1").Value2)
    ws.Shapes("AKSORA_V4_STATUS").TextFrame2.TextRange.Text = "SEÇÝLÝ ÝL: -"
    On Error Resume Next
    ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = "Yatýrým birimi: depo tanýmlarý bekleniyor"
    On Error GoTo 0
    For i = 1 To 5
        With ws.Shapes("AKSORA_V4_IL_" & i)
            If keys(i - 1) = key Then
                .Fill.ForeColor.RGB = RGB(32, 78, 76)
                .line.ForeColor.RGB = RGB(115, 223, 157): .line.Weight = 2.5
                ws.Shapes("AKSORA_V4_STATUS").TextFrame2.TextRange.Text = "SEÇÝLÝ ÝL: " & names(i - 1)
            Else
                .Fill.ForeColor.RGB = RGB(40, 79, 111)
                .line.ForeColor.RGB = RGB(83, 133, 168): .line.Weight = 1
            End If
        End With
    Next i
End Sub

Public Sub AKSORA_YATIRIM_GIRIS()
    Dim ws As Worksheet
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(YAK)
    If Len(CStr(ws.Range("AZ1").Value2)) = 0 Then
        MsgBox "Lütfen bir il seçiniz.", vbInformation, "AKSORA"
    Else
        Y_GIRIS CStr(ws.Range("AZ1").Value2)
    End If
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_BIRIM_ANA_MENU()
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    On Error GoTo Hata
    AKSORA_BIRIM_ACILIS
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub BIRIM_PANEL_AC(ByVal ad As String, ByVal baslik As String)
    Dim ws As Worksheet, ev As Boolean, sc As Boolean, no As Long, metin As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    Set ws = BIRIM_SAYFA(ad)
    If ws.Shapes.Count = 0 Or Val(CStr(ws.Range("BX2").Value2)) < 8 Then BIRIM_CIZ ws, baslik
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    BIRIM_DIGERLERINI_GIZLE ad
    BIRIM_GORUNUM ws
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Private Sub BIRIM_DIGERLERINI_GIZLE(ByVal secilen As String)
    Dim ad As Variant, ws As Worksheet
    For Each ad In Array(HUB, YAT, MUS, MUK, YAK)
        If CStr(ad) <> secilen Then
            Set ws = Nothing
            On Error Resume Next
            Set ws = ThisWorkbook.Worksheets(CStr(ad))
            On Error GoTo 0
            If Not ws Is Nothing Then ws.Visible = xlSheetVeryHidden
        End If
    Next ad
End Sub

Public Sub AKSORA_BIRIM_KAPAK_GERI(Optional ByVal hedef As String = "KAPAK")
    Dim ws As Worksheet, panel As Shape, btn As Shape
    Dim cizim As Boolean, icerik As Boolean, senaryo As Boolean, secim As Long
    Dim no As Long, metin As String
    Set ws = ThisWorkbook.Worksheets(hedef)
    cizim = ws.ProtectDrawingObjects: icerik = ws.ProtectContents
    senaryo = ws.ProtectScenarios: secim = ws.EnableSelection
    On Error GoTo Hata
    Set panel = ws.Shapes("AKSORA_V4_PANEL")
    ws.Unprotect Password:=""
    On Error Resume Next
    Set btn = ws.Shapes("BIRIM_KAPAK_GERI")
    On Error GoTo Hata
    If btn Is Nothing Then
        Set btn = ws.Shapes.AddShape(msoShapeRoundedRectangle, 0, 0, 170, 29)
        btn.Name = "BIRIM_KAPAK_GERI"
    End If
    With btn
        .Left = panel.Left + 18
        .Top = panel.Top + panel.Height - 43
        .Width = 192: .Height = 29
        .Adjustments.Item(1) = 0.12
        .Fill.ForeColor.RGB = RGB(24, 76, 104)
        .line.ForeColor.RGB = RGB(74, 186, 201)
        .line.Weight = 0.8
        .Placement = xlFreeFloating: .Locked = True
        .OnAction = BIRIM_MAKRO("AKSORA_BIRIM_ANA_MENU")
        .Visible = msoTrue: .ZOrder msoBringToFront
        With .TextFrame2
            .MarginLeft = 12: .MarginRight = 8: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = ChrW(8592) & "  Birim seçimine dön"
            .TextRange.Font.Name = "Aptos"
            .TextRange.Font.Size = 10
            .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = RGB(235, 250, 255)
        End With
    End With
Cikis:
    If cizim Or icerik Or senaryo Then
        ws.Protect Password:="", DrawingObjects:=cizim, Contents:=icerik, _
            Scenarios:=senaryo, UserInterfaceOnly:=True
    End If
    ws.EnableSelection = secim
    If no <> 0 Then Err.Raise no, "AKSORA_BIRIM_KAPAK_GERI", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Private Sub BIRIM_GORUNUM(ByVal ws As Worksheet)
    Dim z As Double, oran As Double, genislik As Double, parca As Variant, s As Shape
    Dim no As Long, metin As String
    If mGorunum Or Not (ActiveSheet Is ws) Then Exit Sub
    mGorunum = True
    On Error GoTo Hata
    With ThisWorkbook.Windows(1)
        .DisplayGridlines = False: .DisplayHeadings = False
        .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False: .DisplayVerticalScrollBar = False
        .ScrollRow = 1: .ScrollColumn = 1
        z = .UsableHeight / CH * 100
        .Zoom = application.Max(10, application.Min(200, Fix(z)))
        genislik = (.UsableWidth - 4) * 100 / CDbl(.Zoom)
    End With
    oran = genislik / cw
    ws.Unprotect Password:=pw
    For Each s In ws.Shapes
        If Left$(s.AlternativeText, 17) = "AKSORA_LAYOUT_V2|" Then
            parca = Split(s.AlternativeText, "|")
            s.LockAspectRatio = msoFalse
            s.Left = Val(parca(1)) * oran
            s.Top = Val(parca(2))
            s.Width = Val(parca(3)) * oran
            s.Height = Val(parca(4))
            If s.Name = "BIRIM_LOGO" Then
                s.Width = 98
                s.Left = 30 * oran
            End If
        End If
    Next s
Cikis:
    On Error Resume Next
    ws.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
    mGorunum = False
    On Error GoTo 0
    If no <> 0 Then Err.Raise no, "BIRIM_GORUNUM", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Private Sub BIRIM_CIZ(ByVal ws As Worksheet, ByVal panel As String)
    Dim i As Long, s As Shape, yol As String, resimYukseklik As Double
    yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_BIRIM_BANNER_03.png"
    If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1608, , _
        "AKSORA_BIRIM_BANNER_03.png Excel dosyasinin yaninda bulunmalidir."
    ws.Unprotect pw
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "BIRIM_" Then ws.Shapes(i).Delete
    Next i
    ws.Range("A1:BH44").Interior.Color = RGB(236, 248, 254)
    ws.Range("A1:BH44").ClearContents
    ws.Columns("A:BH").ColumnWidth = 2.5
    Do While ws.Range("A1:BH1").Width < cw
        ws.Columns("A:BH").ColumnWidth = ws.Columns("A").ColumnWidth + 0.25
    Loop
    ws.Rows("1:44").RowHeight = 15
    ws.Range("BX1").Value2 = tag
    ws.Range("BX2").Value2 = 8
    ws.Columns("BX").Hidden = True
    ws.ScrollArea = "A1:BH44"
    ws.Tab.Color = RGB(12, 45, 79)
    BIRIM_RECT ws, "ZEMIN", 0, 0, cw, CH, RGB(236, 248, 254), -1, 1
    yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_PANEL_TEK_PARCA_V7.png"
    If Len(Dir$(yol)) > 0 Then
        Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 126, cw, 483)
        s.Name = "BIRIM_FOTOGRAF": s.Placement = xlFreeFloating: s.Locked = True
    End If
    Set s = BIRIM_RECT(ws, "FOTO_SIS", 0, 126, cw, 483, RGB(235, 248, 255), -1, 1)
    s.Fill.Transparency = 0.45
    yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_BIRIM_BANNER_03.png"
    If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1608, , _
        "AKSORA_BIRIM_BANNER_03.png Excel dosyasinin yaninda bulunmalidir."
    Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
    s.Name = "BIRIM_BANNER_03"
    resimYukseklik = s.Height
    With s.PictureFormat
        .CropTop = resimYukseklik * 195 / 683
        .CropBottom = resimYukseklik * 231 / 683
    End With
    With s
        .LockAspectRatio = msoFalse
        .Left = 0: .Top = 0: .Width = cw: .Height = 126
        .Placement = xlFreeFloating: .Locked = True
    End With
    BIRIM_RECT ws, "UST_CIZGI", 0, 124, cw, 3, RGB(34, 185, 191), -1, 1
    BIRIM_RECT ws, "ALT", 0, 609, cw, 51, RGB(12, 45, 79), -1, 1
    BIRIM_TEXT ws, "ALT_SOL", "AKSORA  /  Birlikte daha güçlü.", 34, 622, 550, 24, 10, RGB(224, 241, 250), False
    BIRIM_TEXT ws, "ALT_SAG", "BÝRÝM ODAKLI  ·  TEK MERKEZ", 797, 622, 253, 24, 9, RGB(137, 210, 229), True
    If Len(panel) = 0 Then
        BIRIM_TEXT ws, "EYEBROW", "HOÞ GELDÝNÝZ", 34, 151, 850, 22, 10, RGB(0, 126, 113), True
        BIRIM_TEXT ws, "BASLIK", "Hangi birim için iþlem yapacaksýnýz?", 34, 183, 1010, 45, 27, RGB(12, 45, 79), True
        BIRIM_TEXT ws, "ALT_BASLIK", "Devam etmek için biriminizi seçin.", 35, 236, 1000, 26, 12, RGB(76, 111, 134), False
        BIRIM_KART ws, "STOK", 34, "01", "Malzeme Stok Birimi", _
            "", "Stok takibi" & vbLf & "Ýller arasý malzeme arama" & vbLf & _
            "Gün bazlý stok karþýlaþtýrma" & vbLf & "Ýl depolarý raporlarý", _
            "Ýþlemlere geç", "AKSORA_BIRIM_MALZEME", RGB(0, 125, 110), True
        BIRIM_KART ws, "YAT", 378, "02", "Yatýrým Birimi", _
            "", "Hakediþ" & vbLf & "Teminat takibi" & vbLf & _
            "Yüklenici stok takibi" & vbLf & "Yatýrým depo takibi", _
            "Ýþlemlere geç", "AKSORA_BIRIM_YATIRIM", RGB(20, 103, 173), False
        BIRIM_KART ws, "MUS", 722, "03", "Müþteri Birimi", _
            "", "Seri numaralý malzeme takibi" & vbLf & "Ýlçe stok kontrolleri" & vbLf & _
            "Yüklenici stok kontrolleri" & vbLf & "Müþteri depo takibi", _
            "Ýþlemlere geç", "AKSORA_BIRIM_MUSTERI", RGB(32, 133, 158), True
        BIRIM_TEXT ws, "SECIM_NOT", "Ýhtiyacýnýz olan iþlem alanýna tek seçimle ulaþýn.", 35, 565, 1000, 24, 10, RGB(87, 121, 142), False
    Else
        BIRIM_TEXT ws, "PANEL_EYEBROW", "BÝRÝM ÇALIÞMA ALANI", 35, 157, 900, 24, 11, RGB(0, 126, 113), True
        BIRIM_TEXT ws, "PANEL_BASLIK", panel, 35, 196, 1000, 52, 32, RGB(12, 45, 79), True
        BIRIM_RECT ws, "PANEL_KART", 35, 282, 1010, 213, RGB(255, 255, 255), RGB(201, 226, 238), msoShapeRoundedRectangle
        BIRIM_TEXT ws, "PANEL_DURUM", "Ýþlem ekranlarý hazýrlanýyor", 65, 328, 915, 39, 22, RGB(12, 45, 79), True
        BIRIM_TEXT ws, "PANEL_NOT", "Bu birime ait iþlemler sonraki adýmda burada yer alacak.", 66, 386, 920, 32, 13, RGB(77, 112, 137), False
        BIRIM_BUTON ws, "GERI", "Birim seçimine dön", 35, 523, 260, 48, RGB(12, 45, 79), "AKSORA_BIRIM_ANA_MENU"
    End If
    For Each s In ws.Shapes
        If Left$(s.Name, 6) = "BIRIM_" Then
            s.AlternativeText = "AKSORA_LAYOUT_V2|" & Trim$(Str$(s.Left)) & "|" & _
                Trim$(Str$(s.Top)) & "|" & Trim$(Str$(s.Width)) & "|" & Trim$(Str$(s.Height))
        End If
    Next s
    ws.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
End Sub

Private Sub BIRIM_KART(ByVal ws As Worksheet, ByVal ad As String, ByVal x As Double, _
    ByVal numara As String, ByVal baslik As String, ByVal aciklama As String, _
    ByVal detay As String, ByVal buton As String, ByVal makro As String, _
    ByVal renk As Long, ByVal aktif As Boolean)
    Dim s As Shape
    BIRIM_RECT ws, ad & "_GOLGE", x + 2, 299, 324, 244, RGB(221, 237, 246), -1, msoShapeRoundedRectangle
    Set s = BIRIM_RECT(ws, ad & "_KART", x, 295, 324, 244, RGB(255, 255, 255), RGB(193, 224, 235), msoShapeRoundedRectangle)
    s.OnAction = BIRIM_MAKRO(makro)
    BIRIM_RECT ws, ad & "_VURGU", x + 18, 315, 4, 33, renk, -1, 1
    BIRIM_TEXT ws, ad & "_NO", numara, x + 35, 315, 220, 31, 19, renk, True
    BIRIM_TEXT ws, ad & "_BASLIK", baslik, x + 21, 367, 286, 31, 18, RGB(12, 45, 79), True
    BIRIM_TEXT ws, ad & "_DETAY", detay, x + 21, 407, 286, 66, 10, RGB(55, 99, 125), False
    With ws.Shapes("BIRIM_" & ad & "_DETAY").TextFrame2
        .VerticalAnchor = msoAnchorTop
        .TextRange.ParagraphFormat.SpaceAfter = 3
    End With
    BIRIM_BUTON ws, ad & "_GIT", buton & "  " & ChrW(8594), x + 21, 484, 282, 36, renk, makro
    If aktif Then
        BIRIM_RECT ws, ad & "_ROZET", x + 224, 318, 78, 25, RGB(222, 248, 235), -1, msoShapeRoundedRectangle
        BIRIM_TEXT ws, ad & "_ROZET_YAZI", "AKTÝF", x + 237, 321, 54, 19, 8, RGB(0, 111, 80), True
    Else
        BIRIM_RECT ws, ad & "_ROZET", x + 165, 318, 137, 25, RGB(184, 45, 55), -1, msoShapeRoundedRectangle
        BIRIM_TEXT ws, ad & "_ROZET_YAZI", "Yapým Aþamasýnda", x + 176, 321, 117, 19, 8, RGB(255, 255, 255), True
    End If
    ' Metinler de ayni hedefe gider; kartin her noktasi secilebilir.
    For Each s In ws.Shapes
        If Left$(s.Name, Len("BIRIM_" & ad & "_")) = "BIRIM_" & ad & "_" Then s.OnAction = BIRIM_MAKRO(makro)
    Next s
End Sub

Private Function BIRIM_RECT(ByVal ws As Worksheet, ByVal ad As String, _
    ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
    ByVal dolgu As Long, ByVal kenar As Long, ByVal tip As Long) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(tip, x, y, w, h)
    With s
        .Name = "BIRIM_" & ad
        .Fill.ForeColor.RGB = dolgu
        If kenar < 0 Then .line.Visible = msoFalse Else .line.ForeColor.RGB = kenar: .line.Weight = 0.8
        .Placement = xlFreeFloating: .Locked = True
        If tip = msoShapeRoundedRectangle Then .Adjustments.Item(1) = 0.09
    End With
    Set BIRIM_RECT = s
End Function

Private Sub BIRIM_TEXT(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
    ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
    ByVal boy As Single, ByVal renk As Long, ByVal kalin As Boolean)
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    With s
        .Name = "BIRIM_" & ad
        .Fill.Visible = msoFalse: .line.Visible = msoFalse
        .Placement = xlFreeFloating: .Locked = True
        With .TextFrame2
            .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
            .WordWrap = msoTrue: .VerticalAnchor = msoAnchorMiddle
            .TextRange.Text = metin
            .TextRange.Font.Name = "Aptos"
            .TextRange.Font.Size = boy
            .TextRange.Font.Fill.ForeColor.RGB = renk
            .TextRange.Font.bold = IIf(kalin, msoTrue, msoFalse)
        End With
    End With
End Sub

Private Sub BIRIM_BUTON(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
    ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, _
    ByVal renk As Long, ByVal makro As String)
    Dim s As Shape
    Set s = BIRIM_RECT(ws, ad, x, y, w, h, renk, -1, msoShapeRoundedRectangle)
    With s.TextFrame2
        .MarginLeft = 15: .MarginRight = 15: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle
        .TextRange.Text = metin
        .TextRange.Font.Name = "Aptos"
        .TextRange.Font.Size = 11
        .TextRange.Font.bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
    End With
    s.OnAction = BIRIM_MAKRO(makro)
End Sub

Private Function BIRIM_MAKRO(ByVal ad As String) As String
    BIRIM_MAKRO = "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & ad
End Function

'==========================================================
' YATIRIM BIRIMI (Y_ sayfalari)
' Ana panel, stok ve rapor sayfalari musteri biriminin kilitli tasarimindan
' kopyalanir; Hakedis ve Teminat sayfalari ayni basliklarla hazirlanir.
'==========================================================

Private Function Y_ADLAR() As Variant
    Y_ADLAR = Array("Y_ANA PANEL", "Y_HAKEDÝÞ", "Y_TEMÝNAT", _
        "Y_CDP_YATIRIM_STOK", "Y_CDP_YATIRIM_RAPOR", "Y_EPP_YATIRIM_STOK", "Y_EPP_YATIRIM_RAPOR", _
        "Y_CDP_YÜKLENÝCÝ_STOK", "Y_CDP_YÜKLENÝCÝ_RAPOR", "Y_EPP_YÜKLENÝCÝ_STOK", "Y_EPP_YÜKLENÝCÝ_RAPOR")
End Function

Private Function Y_BASLIK(ByVal ad As String) As String
    Select Case ad
        Case "Y_ANA PANEL": Y_BASLIK = "ANA PANEL"
        Case "Y_HAKEDÝÞ": Y_BASLIK = "HAKEDÝÞ"
        Case "Y_TEMÝNAT": Y_BASLIK = "TEMÝNAT TAKÝBÝ"
        Case Else: Y_BASLIK = Replace(Replace(Mid$(ad, 3), "_", " "), "RAPOR", "RAPORU")
    End Select
End Function

Private Function Y_TUR(ByVal ad As String) As String
    If ad = "Y_ANA PANEL" Then
        Y_TUR = "PANEL"
    ElseIf ad = "Y_HAKEDÝÞ" Or ad = "Y_TEMÝNAT" Then
        Y_TUR = "BOS"
    ElseIf Right$(ad, 5) = "_STOK" Then
        Y_TUR = "STOK"
    Else
        Y_TUR = "RAPOR"
    End If
End Function

Public Function AKSORA_YATIRIM_SAYFA(ByVal sh As Object) As Boolean
    Dim ad As Variant
    If Not TypeOf sh Is Worksheet Then Exit Function
    If Left$(sh.Name, 2) <> "Y_" Then Exit Function
    For Each ad In Y_ADLAR()
        If sh.Name = CStr(ad) Then AKSORA_YATIRIM_SAYFA = (CStr(sh.Range("BX1").Value2) = Y_ETIKET): Exit Function
    Next ad
End Function

Private Function Y_IL_ADI(ByVal kod As String) As String
    Select Case UCase$(Trim$(kod))
        Case "GIRESUN": Y_IL_ADI = "Giresun"
        Case "TRABZON": Y_IL_ADI = "Trabzon"
        Case "RIZE": Y_IL_ADI = "Rize"
        Case "GUMUSHANE": Y_IL_ADI = "Gümüþhane"
        Case Else: Y_IL_ADI = "Artvin"
    End Select
End Function

Private Function Y_SEKIL(ByVal ws As Worksheet, ByVal ad As String) As Shape
    On Error Resume Next
    Set Y_SEKIL = ws.Shapes(ad)
    On Error GoTo 0
End Function

Private Sub Y_YAZ(ByVal s As Shape, ByVal metin As String)
    If s Is Nothing Then Exit Sub
    On Error Resume Next
    s.TextFrame2.TextRange.Text = metin
    On Error GoTo 0
End Sub

Private Function Y_METIN(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, _
    ByVal x As Double, ByVal y As Double, ByVal w As Double, ByVal h As Double, ByVal punto As Single, _
    ByVal renk As Long, ByVal kalin As Boolean, ByVal hiz As MsoParagraphAlignment) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    s.Name = ad: s.Fill.Visible = msoFalse: s.line.Visible = msoFalse: s.Placement = xlFreeFloating
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoTrue
        .TextRange.Text = metin
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = punto
        .TextRange.Font.bold = IIf(kalin, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = renk
        .TextRange.ParagraphFormat.Alignment = hiz
    End With
    Set Y_METIN = s
End Function

Private Function Y_KUTU(ByVal ws As Worksheet, ByVal ad As String, ByVal x As Double, ByVal y As Double, _
    ByVal w As Double, ByVal h As Double, ByVal dolgu As Long, ByVal kenar As Long) As Shape
    Dim s As Shape
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    s.Name = ad: s.Adjustments.Item(1) = 0.08: s.Placement = xlFreeFloating
    s.Fill.ForeColor.RGB = dolgu: s.Shadow.Visible = msoFalse
    If kenar < 0 Then s.line.Visible = msoFalse Else s.line.ForeColor.RGB = kenar: s.line.Weight = 1
    Set Y_KUTU = s
End Function

Private Sub Y_BAGLA(ByVal s As Shape, ByVal hedef As String)
    If s Is Nothing Then Exit Sub
    If Len(hedef) > 0 Then
        s.AlternativeText = hedef
        s.OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_SAYFA_AC")
    Else
        s.OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_GUNCELLE")
    End If
End Sub

' Eksik Y_ sayfalari musteri sayfalarindan kopyalanarak kurulur (veri tasinmaz).
Private Sub Y_KUR()
    Dim adlar As Variant, i As Long, ws As Worksheet, kaynak As Worksheet, gor As XlSheetVisibility
    Dim da As Boolean, kad As String, adim As String, no As Long, metin As String, yeniAd As String
    Dim onceki As Collection, sh As Object, varmi As Boolean
    adlar = Y_ADLAR()
    da = application.DisplayAlerts: application.DisplayAlerts = False
    On Error GoTo Hata
    For i = 0 To UBound(adlar)
        adim = CStr(adlar(i)) & " / kontrol"
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i)))
        On Error GoTo Hata
        ' Yarim kalmis (etiketsiz) sayfa silinip yeniden kurulur.
        If Not ws Is Nothing Then
            If CStr(ws.Range("BX1").Value2) <> Y_ETIKET Then
                adim = CStr(adlar(i)) & " / yarým sayfayý silme"
                ws.Visible = xlSheetVisible
                ws.Delete
                Set ws = Nothing
            End If
        End If
        If ws Is Nothing Then
            Select Case Y_TUR(CStr(adlar(i)))
                Case "PANEL": kad = "M_ANA PANEL"
                Case "STOK": kad = "M_CDP_STOK"
                Case Else: kad = "M_CDP_RAPOR"
            End Select
            adim = CStr(adlar(i)) & " / kaynak " & kad
            Set kaynak = ThisWorkbook.Worksheets(kad)
            gor = kaynak.Visible
            kaynak.Visible = xlSheetVisible
            ' Kopya kaynagin hemen yanina alinir (sondaki gizli sayfalar etkilemez).
            adim = CStr(adlar(i)) & " / kopyalama"
            Set onceki = New Collection
            For Each sh In ThisWorkbook.Sheets
                onceki.Add sh.Name, sh.Name
            Next sh
            kaynak.Copy after:=kaynak
            ' Yeni sayfa konumdan degil, onceki listede olmamasindan bulunur (gizli sayfa kaymasi).
            Set ws = Nothing
            For Each sh In ThisWorkbook.Sheets
                varmi = False
                On Error Resume Next
                varmi = (Len(onceki(sh.Name)) > 0)
                On Error GoTo Hata
                If Not varmi Then Set ws = sh: Exit For
            Next sh
            If ws Is Nothing Then Err.Raise 9, , "kopya sayfa bulunamadý"
            kaynak.Visible = gor
            Set kaynak = Nothing
            adim = CStr(adlar(i)) & " / ad verme (" & ws.Name & ")"
            ws.Name = CStr(adlar(i))
            ws.Unprotect ""
            ws.Range("BX1:BX40").ClearContents
            adim = CStr(adlar(i)) & " / içerik"
            Select Case Y_TUR(ws.Name)
                Case "PANEL": Y_PANEL_TEMIZLE ws
                Case "STOK": Y_STOK_TEMIZLE ws
                Case "RAPOR": Y_RAPOR_TEMIZLE ws
                Case Else: Y_BOS_KUR ws
            End Select
            ws.Range("BX1").Value2 = Y_ETIKET
            ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next i
    application.DisplayAlerts = da
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not kaynak Is Nothing Then kaynak.Visible = gor
    application.DisplayAlerts = da
    On Error GoTo 0
    Err.Raise no, "Y_KUR", adim & " | " & metin
End Sub

Private Sub Y_PANEL_TEMIZLE(ByVal ws As Worksheet)
    Dim s As Shape, ad As String, i As Long, ilk As Shape, son As Shape, x0 As Double, x1 As Double, bw As Double, g As Double
    Dim adlar As Variant, hedefler As Variant, yeni As Shape
    For Each s In ws.Shapes
        ad = s.Name
        If Left$(ad, 20) = "AKSORA15_KART_DEGER_" Then
            If Right$(ad, 1) = "0" Then Y_YAZ s, "Veri bekleniyor" Else Y_YAZ s, "-"
        End If
        If Left$(ad, 18) = "AKSORA15_KART_LED_" Then s.Fill.ForeColor.RGB = RGB(145, 162, 176)
    Next s
    Y_YAZ Y_SEKIL(ws, "btnBaglantiKontrol"), "TEMÝNAT TAKÝBÝ": Y_BAGLA Y_SEKIL(ws, "btnBaglantiKontrol"), "Y_TEMÝNAT"
    Y_YAZ Y_SEKIL(ws, "btnMalzemeAraPanel"), "HAKEDÝÞ": Y_BAGLA Y_SEKIL(ws, "btnMalzemeAraPanel"), "Y_HAKEDÝÞ"
    Y_YAZ Y_SEKIL(ws, "btnTumStok"), "TÜM STOKLARI GÜNCELLE": Y_BAGLA Y_SEKIL(ws, "btnTumStok"), ""
    Y_YAZ Y_SEKIL(ws, "btnCDP"), "CDP YATIRIM GÜNCELLE": Y_BAGLA Y_SEKIL(ws, "btnCDP"), ""
    Y_YAZ Y_SEKIL(ws, "btnIlce"), "EPP YATIRIM GÜNCELLE": Y_BAGLA Y_SEKIL(ws, "btnIlce"), ""
    Y_YAZ Y_SEKIL(ws, "btnYuk"), "YÜKLENÝCÝ GÜNCELLE": Y_BAGLA Y_SEKIL(ws, "btnYuk"), ""
    Set s = Y_SEKIL(ws, "AKSORA15_IL_DEGISTIR")
    If Not s Is Nothing Then s.OnAction = BIRIM_MAKRO("AKSORA_BIRIM_YATIRIM")
    ' Hizli erisim: 8 sayfa ayni satirda esit aralikla.
    Set ilk = Y_SEKIL(ws, "MUSTERI_HIZLI_0"): Set son = Y_SEKIL(ws, "MUSTERI_HIZLI_5")
    If ilk Is Nothing Or son Is Nothing Then Exit Sub
    x0 = ilk.Left: x1 = son.Left + son.Width: g = 6
    For i = 6 To 7
        Set yeni = son.Duplicate
        yeni.Name = "MUSTERI_HIZLI_" & i
        yeni.Top = son.Top
    Next i
    bw = (x1 - x0 - 7 * g) / 8
    adlar = Array("CDP YATIRIM STOK", "CDP YATIRIM RAPOR", "EPP YATIRIM STOK", "EPP YATIRIM RAPOR", _
        "CDP YÜKLENÝCÝ STOK", "CDP YÜKLENÝCÝ RAPOR", "EPP YÜKLENÝCÝ STOK", "EPP YÜKLENÝCÝ RAPOR")
    hedefler = Array("Y_CDP_YATIRIM_STOK", "Y_CDP_YATIRIM_RAPOR", "Y_EPP_YATIRIM_STOK", "Y_EPP_YATIRIM_RAPOR", _
        "Y_CDP_YÜKLENÝCÝ_STOK", "Y_CDP_YÜKLENÝCÝ_RAPOR", "Y_EPP_YÜKLENÝCÝ_STOK", "Y_EPP_YÜKLENÝCÝ_RAPOR")
    For i = 0 To 7
        Set s = Y_SEKIL(ws, "MUSTERI_HIZLI_" & i)
        If Not s Is Nothing Then
            s.Left = x0 + i * (bw + g): s.Width = bw: s.Top = ilk.Top: s.Height = ilk.Height
            Y_YAZ s, CStr(adlar(i))
            On Error Resume Next
            s.TextFrame2.WordWrap = msoTrue
            s.TextFrame2.MarginLeft = 2: s.TextFrame2.MarginRight = 2
            On Error GoTo 0
            Y_BAGLA s, CStr(hedefler(i))
        End If
    Next i
End Sub

Private Sub Y_STOK_TEMIZLE(ByVal ws As Worksheet)
    Dim son As Long, s As Shape, i As Long
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    son = application.Max(7, ws.Cells(ws.Rows.Count, 5).End(xlUp).Row)
    ws.Range("A7:N" & son).Clear
    ws.Range(ws.Cells(7, 75), ws.Cells(son, 75)).ClearContents
    ws.Range(ws.Rows(7), ws.Rows(son)).RowHeight = ws.StandardHeight
    ws.Range("C2").MergeArea.Cells(1, 1).Value2 = "-"
    ws.Range("C3").MergeArea.Cells(1, 1).Value2 = 0
    ws.Range("C4").MergeArea.Cells(1, 1).Value2 = 0
    ws.Range("A7").Value2 = "Veri bekleniyor: yatýrým depo tanýmlarý yapýlýnca stoklar SAP'den alýnacak."
    ws.Range("A7").Font.Italic = True: ws.Range("A7").Font.Color = RGB(100, 116, 139)
    For Each s In ws.Shapes
        If InStr(s.Name, "STOK_KPI_DEGER_") > 0 Then
            i = Val(Mid$(s.Name, InStrRev(s.Name, "_") + 1))
            On Error Resume Next
            If i >= 0 And i <= 2 Then s.DrawingObject.Formula = "='" & Replace(ws.Name, "'", "''") & "'!C" & CStr(i + 2)
            On Error GoTo 0
        End If
        If InStr(s.Name, "STOK_KPI_ETIKET_2") > 0 Then Y_YAZ s, "Toplam Stok Deðeri"
        If InStr(s.Name, "DEPO_KART_LISTE") > 0 Then
            On Error Resume Next
            s.ControlFormat.RemoveAllItems: s.ControlFormat.AddItem "TÜMÜ": s.ControlFormat.Value = 1
            s.OnAction = ""
            On Error GoTo 0
        End If
    Next s
    Y_BAGLA Y_SEKIL(ws, "btnAnaPanelDon"), "Y_ANA PANEL"
    Y_BAGLA Y_SEKIL(ws, "SAP_STOK_GUNCELLE_BTN"), ""
End Sub

Private Sub Y_RAPOR_TEMIZLE(ByVal ws As Worksheet)
    Dim son As Long, s As Shape, i As Long, ad As String, co As ChartObject
    son = application.Max(60, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 5)
    ws.Range("A9:E" & son).Clear
    ws.Range(ws.Rows(9), ws.Rows(son)).RowHeight = ws.StandardHeight
    ws.Range("A9").Value2 = "Veri bekleniyor: depo tanýmlarý yapýlýnca rapor oluþacak."
    ws.Range("A9").Font.Italic = True: ws.Range("A9").Font.Color = RGB(100, 116, 139)
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 12) = "AKSORA_RP_DM" Then ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
    For Each co In ws.ChartObjects
        On Error Resume Next
        Do While co.Chart.SeriesCollection.Count > 0: co.Chart.SeriesCollection(1).Delete: Loop
        On Error GoTo 0
    Next co
    For Each s In ws.Shapes
        ad = s.Name
        If InStr(ad, "_DEGER_A3_E6") > 0 Then Y_YAZ s, "-"
        If InStr(ad, "_DEGER_F3_J6") > 0 Then Y_YAZ s, "0"
        If InStr(ad, "_DEGER_K3_Q6") > 0 Then Y_YAZ s, "0,00 TL"
        If InStr(ad, "_ETIKET_K3_Q6") > 0 Then Y_YAZ s, "TOPLAM STOK DEÐERÝ"
        If Left$(ad, 21) = "AKSORA_RP_HALKA_RENK_" Or Left$(ad, 19) = "AKSORA_RP_HALKA_AD_" Then s.Visible = msoFalse
        If Left$(ad, 20) = "AKSORA_RP_HALKA_PAY_" And ad <> "AKSORA_RP_HALKA_PAY_BASLIK" Then s.Visible = msoFalse
        If ad = "AKSORA_RP_HALKA_MERKEZ" Then Y_YAZ s, "TOPLAM STOK" & vbLf & "Veri bekleniyor"
    Next s
    Y_BAGLA Y_SEKIL(ws, "btnAnaPanelDon"), "Y_ANA PANEL"
End Sub

' Hakedis ve Teminat: rapor basligi korunur, icerik hedef kartlariyla doldurulur.
Private Sub Y_BOS_KUR(ByVal ws As Worksheet)
    Dim i As Long, ad As String, w As Double, y As Double, s As Shape, x As Double, kw As Double, kutular As Variant, renk As Variant
    Dim silinecek As Collection, sv As Variant
    Set silinecek = New Collection
    For Each s In ws.Shapes
        ad = s.Name
        If Not (ad = "AKT1_RAPOR_UST_ZEMIN" Or ad = "AKT1_RAPOR_BASLIK" Or ad = "btnAnaPanelDon" Or InStr(ad, "FILIGRAN") > 0) Then silinecek.Add ad
    Next s
    On Error Resume Next
    For Each sv In silinecek
        ws.Shapes(CStr(sv)).Delete
    Next sv
    On Error GoTo 0
    ws.Range("A3:Z300").Clear
    ws.Range(ws.Rows(3), ws.Rows(300)).RowHeight = ws.StandardHeight
    Y_BAGLA Y_SEKIL(ws, "btnAnaPanelDon"), "Y_ANA PANEL"
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1200 Else w = s.Width
    y = 70
    If ws.Name = "Y_HAKEDÝÞ" Then
        Y_KUTU ws, "Y_KART_ZEMIN", 12, y, w - 24, 300, RGB(244, 248, 252), RGB(207, 227, 242)
        Y_METIN ws, "Y_KART_ETIKET", "YAPIM AÞAMASINDA", 32, y + 16, 300, 18, 9, RGB(0, 125, 110), True, msoAlignLeft
        Y_METIN ws, "Y_KART_BASLIK", "Otomatik hakediþ", 32, y + 36, w - 64, 30, 20, RGB(12, 45, 79), True, msoAlignLeft
        Y_METIN ws, "Y_KART_METIN", "Sözleþme birim fiyatlarý ve metrajdan hakediþ tutarý otomatik hesaplanýr; önceki hakediþle kümülatif kontrol, " & _
            "kesintiler ve fiyat farký tek ekranda. Hedef: elle hesap hatasýný sýfýra indirmek, hakediþ hazýrlamayý günlerden saniyelere düþürmek.", _
            32, y + 72, w - 64, 48, 11, RGB(51, 65, 85), False, msoAlignLeft
        kutular = Array("Metraj ve birim fiyat", "Kümülatif kontrol", "Kesinti ve fiyat farký", "Günlerden saniyelere")
        kw = (w - 64 - 3 * 12) / 4
        For i = 0 To 3
            x = 32 + i * (kw + 12)
            Y_KUTU ws, "Y_HEDEF_" & i, x, y + 140, kw, 70, RGB(255, 255, 255), RGB(207, 227, 242)
            Y_METIN ws, "Y_HEDEF_NO_" & i, "0" & CStr(i + 1), x + 12, y + 150, 40, 22, 14, RGB(0, 125, 110), True, msoAlignLeft
            Y_METIN ws, "Y_HEDEF_AD_" & i, CStr(kutular(i)), x + 12, y + 174, kw - 24, 26, 11, RGB(12, 45, 79), True, msoAlignLeft
        Next i
        Y_METIN ws, "Y_KART_NOT", "Sayfa tasarýmý netleþtiðinde bu alan hakediþ giriþ ve hesap ekranýna dönüþecek.", 32, y + 240, w - 64, 20, 9, RGB(100, 116, 139), False, msoAlignLeft
    Else
        Y_KUTU ws, "Y_KART_ZEMIN", 12, y, w - 24, 330, RGB(244, 248, 252), RGB(207, 227, 242)
        Y_METIN ws, "Y_KART_ETIKET", "YAPIM AÞAMASINDA", 32, y + 16, 300, 18, 9, RGB(0, 125, 110), True, msoAlignLeft
        Y_METIN ws, "Y_KART_BASLIK", "Teminat ve yüklenici stok riski", 32, y + 36, w - 64, 30, 20, RGB(12, 45, 79), True, msoAlignLeft
        Y_METIN ws, "Y_KART_METIN", "Her yüklenici için belirlenen teminat tutarý, yüklenici deposundaki malzemeyle karþýlaþtýrýlýr. SAP deposu ile fiziki depo yan yana konur; " & _
            "sahada kullanýlmýþ ama hakediþi bekleyen malzemeler seçilip ""hakediþ bekliyor"" olarak iþaretlenir. Geriye kalan açýklanamayan fark yüklenicinin gerçek riskidir.", _
            32, y + 72, w - 64, 48, 11, RGB(51, 65, 85), False, msoAlignLeft
        kutular = Array("SAP'deki yüklenici deposu", "Sayýlan fiziki depo", "Hakediþ bekleyen", "Açýk risk")
        renk = Array(RGB(18, 62, 102), RGB(234, 244, 251), RGB(255, 241, 214), RGB(254, 226, 226))
        kw = (w - 64 - 3 * 34) / 4
        For i = 0 To 3
            x = 32 + i * (kw + 34)
            Set s = Y_KUTU(ws, "Y_HEDEF_" & i, x, y + 140, kw, 56, CLng(renk(i)), -1)
            Y_METIN ws, "Y_HEDEF_AD_" & i, CStr(kutular(i)), x + 8, y + 148, kw - 16, 40, 12, IIf(i = 0, RGB(255, 255, 255), RGB(12, 45, 79)), True, msoAlignCenter
            If i < 3 Then Y_METIN ws, "Y_HEDEF_OP_" & i, IIf(i < 2, ChrW(8722), "="), x + kw, y + 148, 34, 40, 20, RGB(12, 45, 79), True, msoAlignCenter
        Next i
        Set s = Y_KUTU(ws, "Y_ORAN_ZEMIN", 32, y + 214, w - 64, 40, RGB(12, 45, 79), -1)
        Y_METIN ws, "Y_ORAN_METIN", "Risk oraný = Açýk risk ÷ Teminat tutarý   •   yüklenici bazýnda, otomatik", 48, y + 214, w - 96, 40, 13, RGB(255, 255, 255), True, msoAlignLeft
        Y_METIN ws, "Y_KART_NOT", "Proje kapanýþýnda yüklenici bazýnda tek sayfalýk risk raporu. Yüklenici deposu verisi AKSORA'da zaten okunuyor.", 32, y + 268, w - 64, 20, 9, RGB(100, 116, 139), False, msoAlignLeft
    End If
    For Each s In ws.Shapes
        If InStr(s.Name, "FILIGRAN") > 0 Then s.Top = y + 350
    Next s
End Sub

' Secili ile gore basliklar, kart basliklari, alt bilgi ve panel bandi.
Private Sub Y_IL_UYGULA(ByVal kod As String)
    Dim ad As Variant, ws As Worksheet, il As String, s As Shape, korumali As Boolean
    il = Y_IL_ADI(kod)
    For Each ad In Y_ADLAR()
        Set ws = ThisWorkbook.Worksheets(CStr(ad))
        If CStr(ws.Range("BX20").Value2) <> kod Then
            korumali = ws.ProtectContents Or ws.ProtectDrawingObjects
            ws.Unprotect ""
            Select Case Y_TUR(ws.Name)
                Case "PANEL"
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_BASLIK_B"), "CDP YATIRIM / " & AKSORA_IL_TESIS(kod, "CDP")
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_BASLIK_L"), "EPP YATIRIM / " & AKSORA_IL_TESIS(kod, "EPP")
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_BASLIK_Q"), "YÜKLENÝCÝ (CDP + EPP)"
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_ALT_KURUM"), "ÇORUH ELEKTRÝK DAÐITIM A.Þ. | " & il & " Ýl Müdürlüðü"
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_PANEL_TARIHI"), "Yatýrým Birimi  |  Panel kontrolü: " & Format$(Now, "dd.mm.yyyy hh:nn")
                    Set s = Y_SEKIL(ws, "MUSTERI_ANA_SECIM_SERIDI")
                    If Not s Is Nothing Then
                        s.Fill.ForeColor.RGB = RGB(224, 242, 254)
                        Y_YAZ s, il & " Yatýrým Birimi   |   Depo tanýmlarý yapýlýnca CDP ve EPP yatýrým ile yüklenici stoklarý SAP'den alýnacak"
                        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(12, 45, 79)
                    End If
                    Y_BANNER ws, kod
                Case "STOK"
                    Y_YAZ Y_SEKIL(ws, "AKT1_STOK_BASLIK"), Y_BASLIK(ws.Name) & "  |  " & il & " Ýl Müdürlüðü"
                Case Else
                    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | " & Y_BASLIK(ws.Name) & "  |  " & il
                    Y_YAZ Y_SEKIL(ws, "AKSORA_RP_ALT"), "AKSORA  |  " & Y_BASLIK(ws.Name) & "  |  " & il & " Ýl Müdürlüðü"
            End Select
            ws.Range("BX20").Value2 = kod
            If korumali Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next ad
End Sub

' Panel bandi: AKSORA_YATIRIM_PANEL_BANTLARI.png varsa o kullanilir (5 il, musteri bandiyla ayni duzen).
Private Sub Y_BANNER(ByVal ws As Worksheet, ByVal kod As String)
    Dim yol As String, eski As Shape, s As Shape, h As Double, bolum As Long, sinirlar As Variant
    yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_YATIRIM_PANEL_BANTLARI.png"
    If Len(Dir$(yol)) = 0 Then yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_MUSTERI_PANEL_BANTLARI.png"
    If Len(Dir$(yol)) = 0 Then Exit Sub
    Set eski = Y_SEKIL(ws, "MUSTERI_ANA_BANNER")
    If eski Is Nothing Then Exit Sub
    Select Case UCase$(kod)
        Case "GIRESUN": bolum = 0
        Case "TRABZON": bolum = 1
        Case "RIZE": bolum = 2
        Case "GUMUSHANE": bolum = 3
        Case Else: bolum = 4
    End Select
    Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 0, -1, -1)
    h = s.Height
    sinirlar = Array(0, 149, 297, 448, 597, 742)
    s.PictureFormat.CropTop = h * CDbl(sinirlar(bolum)) / 742
    s.PictureFormat.CropBottom = h * (742 - CDbl(sinirlar(bolum + 1))) / 742
    s.LockAspectRatio = msoFalse
    s.Left = eski.Left: s.Top = eski.Top: s.Width = eski.Width: s.Height = eski.Height
    s.Placement = xlFreeFloating
    eski.Delete
    s.Name = "MUSTERI_ANA_BANNER"
    s.ZOrder msoBringToFront
End Sub

' Sayfa gorunumu: pencereye sigar, sekmeler acik, kaydirma basa.
Public Sub AKSORA_YATIRIM_GORUNUM(ByVal ws As Worksheet)
    Dim s As Shape, gen As Double, alt As Double, z As Double
    If Not (ActiveSheet Is ws) Then Exit Sub
    For Each s In ws.Shapes
        If s.Visible = msoTrue Then
            gen = application.Max(gen, s.Left + s.Width)
            alt = application.Max(alt, s.Top + s.Height)
        End If
    Next s
    If gen <= 0 Then gen = ws.Range("A1:L1").Width
    With ThisWorkbook.Windows(1)
        If AKSORA_PENCERE_KUCUK(ThisWorkbook.Windows(1)) Then Exit Sub
        .DisplayWorkbookTabs = True: .DisplayGridlines = False: .DisplayHeadings = False
        .DisplayHorizontalScrollBar = False: .DisplayVerticalScrollBar = True
        z = (.UsableWidth - 8) / gen * 100
        If ws.Name = "Y_ANA PANEL" And alt > 0 Then z = application.Min(z, (.UsableHeight - 8) / alt * 100)
        z = application.Max(10, application.Min(400, Fix(z)))
        If .Zoom <> z Then .Zoom = z
        .ScrollRow = 1: .ScrollColumn = 1
    End With
    On Error Resume Next
    ws.Range("A1").Select
    On Error GoTo 0
End Sub

Public Sub AKSORA_YATIRIM_SAYFA_AC()
    Dim hedef As Worksheet, ad As String, ev As Boolean
    On Error GoTo Hata
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    ad = ActiveSheet.Shapes(CStr(application.Caller)).AlternativeText
    Set hedef = ThisWorkbook.Worksheets(ad)
    ev = application.EnableEvents
    application.ScreenUpdating = False: application.EnableEvents = False
    hedef.Visible = xlSheetVisible
    hedef.Select Replace:=True
    AKSORA_YATIRIM_GORUNUM hedef
    application.EnableEvents = ev: application.ScreenUpdating = True
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_YATIRIM_GUNCELLE()
    MsgBox "Yatýrým biriminde depo tanýmlarý henüz yapýlmadý." & vbCrLf & _
        "Tanýmlar verildiðinde CDP ve EPP yatýrým ile yüklenici stoklarý SAP'den tek tuþla alýnacak.", vbInformation, "AKSORA | Yatýrým Birimi"
End Sub

' Il secim ekranindan giris: sayfalar kurulur, secili ile uyarlanir, yalniz Y_ sayfalari gorunur.
Private Sub Y_GIRIS(ByVal kod As String)
    Dim ev As Boolean, sc As Boolean, ad As Variant, k As Worksheet, panel As Worksheet, no As Long, metin As String
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    Y_KUR
    Y_IL_UYGULA kod
    For Each ad In Y_ADLAR()
        ThisWorkbook.Worksheets(CStr(ad)).Visible = xlSheetVisible
    Next ad
    Set panel = ThisWorkbook.Worksheets("Y_ANA PANEL")
    panel.Select Replace:=True
    For Each k In ThisWorkbook.Worksheets
        If Not AKSORA_YATIRIM_SAYFA(k) Then
            If k.Visible = xlSheetVisible Then k.Visible = xlSheetVeryHidden
        End If
    Next k
    AKSORA_YATIRIM_GORUNUM panel
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = True
    If no <> 0 Then MsgBox "Yatýrým birimi açýlamadý." & vbCrLf & metin & vbCrLf & "(Hata " & no & ")", vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    application.DisplayAlerts = True
    Resume Cikis
End Sub
