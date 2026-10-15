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
Private Const Y_ETIKET As String = "AKSORA_YATIRIM_V2"
Private mYCalisiyor As Boolean
Private mYToplu As Boolean
Private mYOdak As String
Private Const YT_ILK As Long = 10
Private Const YH_KALEM_ILK As Long = 23
Private Const YH_KALEM_ADET As Long = 30
Private Const YH_GECMIS_ILK As Long = 58

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
    application.MoveAfterReturn = (sh.Name <> "M_MALZEME_ARA" And sh.Name <> "Y_MALZEME_ARA")
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
    Dim ws As Worksheet, kilit As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(MUK)
    If Len(CStr(ws.Range("AZ1").Value2)) = 0 Then
        MsgBox "Lütfen bir il seçiniz.", vbInformation, "AKSORA"
    ElseIf Not AKSORA_MUSTERI_IL_DESTEK(CStr(ws.Range("AZ1").Value2)) Then
        MsgBox "Seçilen ilin müþteri birimi þu an aktif deðil. Aktif iller: Artvin, Trabzon, Gümüþhane, Giresun, Rize.", vbInformation, "AKSORA"
    Else
        ' Secilen ilin stoklari ekrana yuklenir (onceki ilin verisi kendi kaydina saklanir).
        ' Gecis boyunca ekran ve pencere kilitli: ara sayfalar gorunmez, tek seferde panel gelir.
        application.ScreenUpdating = False
        kilit = AKSORA_PENCERE_KILIT(True)
        AKSORA_MUSTERI_IL_DEGIS CStr(ws.Range("AZ1").Value2)
        application.ScreenUpdating = False
        AKSORA_MUSTERI_TEMA_AC
        If kilit Then AKSORA_PENCERE_KILIT False
        application.ScreenUpdating = True
    End If
    Exit Sub
Hata:
    If kilit Then AKSORA_PENCERE_KILIT False
    application.ScreenUpdating = True
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
    If Len(key) > 0 Then
        If AKSORA_YATIRIM_IL_SON_TARIH(key) > 0 Then
            ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = "Son güncelleme: " & Format$(AKSORA_YATIRIM_IL_SON_TARIH(key), "dd.mm.yyyy hh:nn")
        Else
            ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = "Son güncelleme: henüz yapýlmadý"
        End If
    Else
        ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = "Yatýrým birimi"
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
    Y_ADLAR = Array("Y_ANA PANEL", "Y_TEMÝNAT", "Y_HAKEDÝÞ", "Y_MALZEME_ARA", _
        "Y_CDP_YATIRIM_STOK", "Y_CDP_YATIRIM_RAPOR", "Y_EPP_YATIRIM_STOK", "Y_EPP_YATIRIM_RAPOR", _
        "Y_CDP_YÜKLENÝCÝ_STOK", "Y_CDP_YÜKLENÝCÝ_RAPOR", "Y_EPP_YÜKLENÝCÝ_STOK", "Y_EPP_YÜKLENÝCÝ_RAPOR")
End Function

Private Function Y_BASLIK(ByVal ad As String) As String
    Select Case ad
        Case "Y_ANA PANEL": Y_BASLIK = "ANA PANEL"
        Case "Y_MALZEME_ARA": Y_BASLIK = "MALZEME ARAMA"
        Case "Y_HAKEDÝÞ": Y_BASLIK = "HAKEDÝÞ"
        Case "Y_TEMÝNAT": Y_BASLIK = "TEMÝNAT TAKÝBÝ"
        Case Else: Y_BASLIK = Replace(Replace(Mid$(ad, 3), "_", " "), "RAPOR", "RAPORU")
    End Select
End Function

Private Function Y_TUR(ByVal ad As String) As String
    If ad = "Y_ANA PANEL" Then
        Y_TUR = "PANEL"
    ElseIf ad = "Y_MALZEME_ARA" Then
        Y_TUR = "ARAMA"
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
    Dim duzen As Variant
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
                Case "ARAMA": kad = "M_MALZEME_ARA"
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
            duzen = ws.Range("BX8").Value2
            ws.Range("BX1:BX40").ClearContents
            ws.Range("BX8").Value2 = duzen
            adim = CStr(adlar(i)) & " / içerik"
            Select Case Y_TUR(ws.Name)
                Case "PANEL": Y_PANEL_TEMIZLE ws
                Case "ARAMA": Y_ARAMA_TEMIZLE ws
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
    For i = 6 To 8
        Set yeni = son.Duplicate
        yeni.Name = "MUSTERI_HIZLI_" & i
        yeni.Top = son.Top
    Next i
    bw = (x1 - x0 - 8 * g) / 9
    adlar = Array("MALZEME ARAMA", "CDP YATIRIM STOK", "CDP YATIRIM RAPOR", "EPP YATIRIM STOK", "EPP YATIRIM RAPOR", _
        "CDP YÜKLENÝCÝ STOK", "CDP YÜKLENÝCÝ RAPOR", "EPP YÜKLENÝCÝ STOK", "EPP YÜKLENÝCÝ RAPOR")
    hedefler = Array("Y_MALZEME_ARA", "Y_CDP_YATIRIM_STOK", "Y_CDP_YATIRIM_RAPOR", "Y_EPP_YATIRIM_STOK", "Y_EPP_YATIRIM_RAPOR", _
        "Y_CDP_YÜKLENÝCÝ_STOK", "Y_CDP_YÜKLENÝCÝ_RAPOR", "Y_EPP_YÜKLENÝCÝ_STOK", "Y_EPP_YÜKLENÝCÝ_RAPOR")
    For i = 0 To 8
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
                        Y_YAZ s, il & " Yatýrým Birimi"
                        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(12, 45, 79)
                    End If
                    Y_BANNER ws, kod
                Case "STOK"
                    Y_YAZ Y_SEKIL(ws, "AKT1_STOK_BASLIK"), Y_BASLIK(ws.Name) & "  |  " & il & " Ýl Müdürlüðü"
                Case "ARAMA"
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
    yol = AKSORA_GORSEL("AKSORA_YATIRIM_PANEL_BANTLARI.png")
    If Len(Dir$(yol)) = 0 Then yol = AKSORA_GORSEL("AKSORA_MUSTERI_PANEL_BANTLARI.png")
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

' Dugmeye gore kapsam: 1 CDP yatirim, 2 EPP yatirim, 3 CDP yuklenici, 4 EPP yuklenici.
Public Sub AKSORA_YATIRIM_GUNCELLE()
    Dim kapsam As String, ad As String, cagiran As String, bilgi As String
    If AKSORA_V11_MESGUL Or mYCalisiyor Then Exit Sub
    On Error Resume Next
    cagiran = CStr(application.Caller)
    ad = ActiveSheet.Name
    On Error GoTo 0
    Select Case ad
        Case "Y_CDP_YATIRIM_STOK", "Y_CDP_YATIRIM_RAPOR": kapsam = "1"
        Case "Y_EPP_YATIRIM_STOK", "Y_EPP_YATIRIM_RAPOR": kapsam = "2"
        Case "Y_CDP_YÜKLENÝCÝ_STOK", "Y_CDP_YÜKLENÝCÝ_RAPOR": kapsam = "3"
        Case "Y_EPP_YÜKLENÝCÝ_STOK", "Y_EPP_YÜKLENÝCÝ_RAPOR": kapsam = "4"
        Case Else
            Select Case cagiran
                Case "btnCDP": kapsam = "1"
                Case "btnIlce": kapsam = "2"
                Case "btnYuk": kapsam = "34"
                Case Else: kapsam = "1234"
            End Select
    End Select
    Y_GUNCELLE Y_AKTIF_IL(), kapsam, False, bilgi
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
    Y_IL_VERI kod
    Y_PANEL_YENILE
    Y_ARAMA_TARIH
    Y_ALT_SAYFALAR kod
    For Each ad In Y_ADLAR()
        ThisWorkbook.Worksheets(CStr(ad)).Visible = xlSheetVisible
    Next ad
    Y_SEKME_SIRASI
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

'==========================================================
' YATIRIM STOKLARI (SAP MB52 -> Y_ stok sayfalari)
' CDP ve EPP yatirim: D002, D011. Yuklenici: malzeme stok birimindeki
' ilin yuklenici depolari. Her il kendi kaydinda (_YA_<IL>_<n>) saklanir;
' arama yalnizca yatirim verisini tarar.
'==========================================================

Private Function Y_STOK_ADLARI() As Variant
    Y_STOK_ADLARI = Array("Y_CDP_YATIRIM_STOK", "Y_EPP_YATIRIM_STOK", "Y_CDP_YÜKLENÝCÝ_STOK", "Y_EPP_YÜKLENÝCÝ_STOK")
End Function

Private Function Y_AKTIF_IL() As String
    On Error Resume Next
    Y_AKTIF_IL = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("Y_ANA PANEL").Range("BX20").Value2)))
    On Error GoTo 0
    If Len(Y_AKTIF_IL) = 0 Then Y_AKTIF_IL = "ARTVIN"
End Function

Private Function Y_DEPO_UYGUN(ByVal depo As String, ByVal i As Long, ByVal il As String) As Boolean
    depo = UCase$(Trim$(depo))
    If i <= 2 Then
        Y_DEPO_UYGUN = (depo = "D002" Or depo = "D011")
    Else
        Y_DEPO_UYGUN = AKSORA_IL_DEPO_GECERLI(depo, "YUKLENICI", il)
    End If
End Function

' MB52 PANEL duzeni (12 sutun) -> musteri stok duzeni (14 sutun), depo ve malzemeye gore sirali.
Private Function Y_DONUSTUR(ByVal v As Variant, ByVal il As String, ByVal i As Long) As Variant
    Dim tesis As String, r As Long, n As Long, k As Long, idx() As Long, anahtar() As String
    Dim sonuc() As Variant, miktar As Double, deger As Double
    tesis = AKSORA_IL_TESIS(il, IIf(i Mod 2 = 1, "CDP", "EPP"))
    ReDim idx(1 To UBound(v, 1)): ReDim anahtar(1 To UBound(v, 1))
    For r = 1 To UBound(v, 1)
        If Trim$(CStr(v(r, 1))) = tesis And Len(Trim$(CStr(v(r, 7)))) > 0 Then
            If Y_DEPO_UYGUN(CStr(v(r, 2)), i, il) Then
                n = n + 1: idx(n) = r
                anahtar(n) = UCase$(Trim$(CStr(v(r, 2)))) & "|" & Trim$(CStr(v(r, 7))) & "|" & Trim$(CStr(v(r, 6)))
            End If
        End If
    Next r
    If n = 0 Then Exit Function
    Y_SIRALA anahtar, idx, 1, n
    ReDim sonuc(1 To n, 1 To 14)
    For k = 1 To n
        r = idx(k)
        miktar = AKSORA_YATIRIM_SAYI(v(r, 9)): deger = AKSORA_YATIRIM_SAYI(v(r, 11))
        sonuc(k, 1) = Trim$(CStr(v(r, 1))): sonuc(k, 2) = UCase$(Trim$(CStr(v(r, 2))))
        sonuc(k, 3) = Trim$(CStr(v(r, 3))): sonuc(k, 4) = Trim$(CStr(v(r, 4)))
        sonuc(k, 5) = Trim$(CStr(v(r, 7))): sonuc(k, 6) = Trim$(CStr(v(r, 8)))
        sonuc(k, 7) = miktar: sonuc(k, 8) = Trim$(CStr(v(r, 10)))
        sonuc(k, 9) = Trim$(CStr(v(r, 6))): sonuc(k, 10) = vbNullString
        sonuc(k, 11) = Round(deger, 2)
        If miktar <> 0 Then sonuc(k, 12) = Round(deger / miktar, 2) Else sonuc(k, 12) = 0#
        sonuc(k, 13) = "SAP"
    Next k
    Y_DONUSTUR = sonuc
End Function

Private Sub Y_SIRALA(ByRef k() As String, ByRef idx() As Long, ByVal ilk As Long, ByVal son As Long)
    Dim i As Long, j As Long, p As String, ts As String, tl As Long
    i = ilk: j = son: p = k((ilk + son) \ 2)
    Do While i <= j
        Do While StrComp(k(i), p, vbTextCompare) < 0: i = i + 1: Loop
        Do While StrComp(k(j), p, vbTextCompare) > 0: j = j - 1: Loop
        If i <= j Then
            ts = k(i): k(i) = k(j): k(j) = ts
            tl = idx(i): idx(i) = idx(j): idx(j) = tl
            i = i + 1: j = j - 1
        End If
    Loop
    If ilk < j Then Y_SIRALA k, idx, ilk, j
    If i < son Then Y_SIRALA k, idx, i, son
End Sub

'---------- Il kayitlari ----------
Private Function Y_ARSIV_SAYFA(ByVal il As String, ByVal i As Long, ByVal olustur As Boolean) As Worksheet
    Dim ad As String, ws As Worksheet, onceki As Object, sc As Boolean, ev As Boolean
    ad = "_YA_" & UCase$(Trim$(il)) & "_" & CStr(i)
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
    If ws Is Nothing And olustur Then
        sc = application.ScreenUpdating: ev = application.EnableEvents
        application.ScreenUpdating = False: application.EnableEvents = False
        Set onceki = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ad
        ws.Visible = xlSheetVeryHidden
        If Not onceki Is Nothing Then onceki.Activate
        application.EnableEvents = ev: application.ScreenUpdating = sc
    End If
    Set Y_ARSIV_SAYFA = ws
End Function

Public Function AKSORA_YATIRIM_ARSIV_SAYFA(ByVal il As String, ByVal i As Long) As Worksheet
    Set AKSORA_YATIRIM_ARSIV_SAYFA = Y_ARSIV_SAYFA(il, i, False)
End Function

Private Sub Y_ARSIV_KAYDET(ByVal il As String, ByVal i As Long, ByVal a As Variant, ByVal tarih As Date)
    Dim ar As Worksheet
    Set ar = Y_ARSIV_SAYFA(il, i, True)
    ar.Range("A1").Value = tarih
    ar.Range("B1").Value2 = UCase$(Trim$(il))
    AKSORA_PAKET_YAZ ar, a
End Sub

' Ilin son yatirim guncellemesi (kayit yoksa 0).
Public Function AKSORA_YATIRIM_IL_SON_TARIH(ByVal il As String) As Date
    Dim i As Long, ar As Worksheet, v As Variant
    For i = 1 To 4
        Set ar = Y_ARSIV_SAYFA(il, i, False)
        If Not ar Is Nothing Then
            v = ar.Range("A1").Value
            If Not IsError(v) Then
                If IsDate(v) Then
                    If CDate(v) > AKSORA_YATIRIM_IL_SON_TARIH Then AKSORA_YATIRIM_IL_SON_TARIH = CDate(v)
                End If
            End If
        End If
    Next i
End Function

'---------- Stok ve rapor sayfalari ----------
Private Sub Y_KORU(ByVal ws As Worksheet)
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
End Sub

Private Sub Y_STOK_YAZ(ByVal ws As Worksheet, ByVal a As Variant, ByVal il As String, ByVal tarih As Variant)
    Dim son As Long, n As Long
    If IsEmpty(a) Then n = 0 Else n = UBound(a, 1)
    son = application.Max(7, AKSORA_M_SON(ws))
    ws.Unprotect ""
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    ws.Range("A7:N" & application.Max(son, n + 6)).ClearContents
    ws.Range("A7").Font.Italic = False: ws.Range("A7").Font.ColorIndex = xlColorIndexAutomatic
    If n > 0 Then
        ws.Range("A7:F" & n + 6).NumberFormat = "@"
        ws.Range("H7:J" & n + 6).NumberFormat = "@"
        ws.Range("A7").Resize(n, 14).Value2 = a
    End If
    ws.Range("K6:N6").Value = Array("Deðer (TL)", "Birim stok deðeri (TL)", "Fiyat durumu", "Fiyat eksik")
    ws.Columns("K:N").Hidden = True
    If n > 0 Then ws.Range("K7:L" & n + 6).NumberFormat = "#,##0.00 ""TL"""
    If IsDate(tarih) Then
        ws.Range("C2").Value = CDate(tarih): ws.Range("C2").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    Else
        ws.Range("C2").Value2 = "-"
    End If
    ws.Range("BX19").Value2 = UCase$(Trim$(il))
    AKSORA_ARAMA_FAZLA_SATIR ws, n + 8
    Y_STOK_YENIDEN ws, il
    If n = 0 Then
        ws.Unprotect ""
        If IsDate(tarih) Then
            ws.Range("A7").Value2 = "Bu depolarda stok bulunamadý (" & Format$(CDate(tarih), "dd.mm.yyyy hh:nn") & ")."
        Else
            ws.Range("A7").Value2 = "Veri bekleniyor: SAP STOKLARINI GÜNCELLE ile stoklar SAP'den alýnýr."
        End If
        ws.Range("A7").Font.Italic = True: ws.Range("A7").Font.Color = RGB(100, 116, 139)
        Y_KORU ws
    End If
End Sub

Private Sub Y_STOK_YENIDEN(ByVal ws As Worksheet, ByVal il As String)
    Dim r As Worksheet
    ws.Unprotect ""
    ws.Range("BX9").ClearContents
    AKSORA_MUSTERI_STOK_KUR_DIS ws
    Y_YAZ Y_SEKIL(ws, "AKT1_STOK_BASLIK"), Y_BASLIK(ws.Name) & "  |  " & Y_IL_ADI(il) & " Ýl Müdürlüðü"
    Y_BAGLA Y_SEKIL(ws, "btnAnaPanelDon"), "Y_ANA PANEL"
    Y_BAGLA Y_SEKIL(ws, "SAP_STOK_GUNCELLE_BTN"), ""
    Y_YAZ Y_SEKIL(ws, "SAP_STOK_GUNCELLE_BTN"), "SAP STOKLARINI GÜNCELLE"
    Y_KORU ws
    Set r = ThisWorkbook.Worksheets(Left$(ws.Name, Len(ws.Name) - 5) & "_RAPOR")
    r.Unprotect ""
    Y_YAZ Y_SEKIL(r, "AKT1_RAPOR_BASLIK"), "AKSORA | " & Y_BASLIK(r.Name) & "  |  " & Y_IL_ADI(il)
    AKSORA_MUSTERI_RAPOR_KUR_DIS r
    On Error Resume Next
    AKSORA_YATIRIM_DM_DUZEN r
    On Error GoTo 0
    Y_BAGLA Y_SEKIL(r, "btnAnaPanelDon"), "Y_ANA PANEL"
    Y_KORU r
End Sub

' Il degisince ekrandaki dort stok sayfasi o ilin kaydindan yuklenir (kayit yoksa bos).
Private Sub Y_IL_VERI(ByVal kod As String)
    Dim adlar As Variant, i As Long, ws As Worksheet, ar As Worksheet, son As Long, a As Variant, tarih As Variant
    kod = UCase$(Trim$(kod))
    adlar = Y_STOK_ADLARI()
    For i = 1 To 4
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i - 1)))
        If CStr(ws.Range("BX19").Value2) <> kod Then
            a = Empty: tarih = Empty
            Set ar = Y_ARSIV_SAYFA(kod, i, False)
            If Not ar Is Nothing Then
                tarih = ar.Range("A1").Value
                a = AKSORA_PAKET_OKU(ar)
            End If
            Y_STOK_YAZ ws, a, kod, tarih
        End If
    Next i
End Sub

'---------- Guncelleme ----------
Private Sub Y_DURUM(ByVal metin As String)
    application.StatusBar = "AKSORA | Yatýrým | " & metin
    If mYToplu Then Y_TOPLU_DURUM metin Else Y_SERIT metin, 1
    DoEvents
End Sub

Private Function Y_GUNCELLE(ByVal il As String, ByVal kapsam As String, ByVal sessiz As Boolean, ByRef bilgi As String) As Boolean
    Dim vC As Variant, vE As Variant, mesaj As String, i As Long, a As Variant, tarih As Date
    Dim bas As Double, sure As Double, adet(1 To 4) As Long, ev As Boolean, sc As Boolean, hesap As XlCalculation
    Dim ekran As Boolean, ozet As String, adlar As Variant, mesgulOnce As Boolean, hMesaj As String
    If mYCalisiyor Then Exit Function
    mYCalisiyor = True: bas = Timer
    mesgulOnce = AKSORA_V11_MESGUL
    ev = application.EnableEvents: sc = application.ScreenUpdating: hesap = application.Calculation
    On Error GoTo Hata
    il = UCase$(Trim$(il))
    ekran = (CStr(ThisWorkbook.Worksheets("Y_CDP_YATIRIM_STOK").Range("BX19").Value2) = il)
    AKSORA_V11_MESGUL = True
    If InStr(kapsam, "1") > 0 Or InStr(kapsam, "3") > 0 Then
        Y_DURUM Y_IL_ADI(il) & " | CDP " & AKSORA_IL_TESIS(il, "CDP") & " stoklarý SAP'den alýnýyor..."
        vC = AKSORA_YATIRIM_MB52("CDP", il, mesaj)
        If Not IsArray(vC) Then Err.Raise vbObjectError + 2101, , "CDP " & AKSORA_IL_TESIS(il, "CDP") & ": " & mesaj
    End If
    If InStr(kapsam, "2") > 0 Or InStr(kapsam, "4") > 0 Then
        Y_DURUM Y_IL_ADI(il) & " | EPP " & AKSORA_IL_TESIS(il, "EPP") & " stoklarý SAP'den alýnýyor..."
        mesaj = vbNullString
        vE = AKSORA_YATIRIM_MB52("EPP", il, mesaj)
        If Not IsArray(vE) Then Err.Raise vbObjectError + 2102, , "EPP " & AKSORA_IL_TESIS(il, "EPP") & ": " & mesaj
    End If
    Y_DURUM Y_IL_ADI(il) & " | Sayfalar yazýlýyor..."
    AKSORA_V11_MESGUL = mesgulOnce
    application.EnableEvents = False: application.Calculation = xlCalculationManual
    application.ScreenUpdating = False
    tarih = Now
    adlar = Y_STOK_ADLARI()
    For i = 1 To 4
        If InStr(kapsam, CStr(i)) > 0 Then
            If i Mod 2 = 1 Then a = Y_DONUSTUR(vC, il, i) Else a = Y_DONUSTUR(vE, il, i)
            If IsEmpty(a) Then adet(i) = 0 Else adet(i) = UBound(a, 1)
            Y_ARSIV_KAYDET il, i, a, tarih
            If ekran Then Y_STOK_YAZ ThisWorkbook.Worksheets(CStr(adlar(i - 1))), a, il, tarih
            ozet = ozet & vbCrLf & Choose(i, "CDP yatýrým", "EPP yatýrým", "CDP yüklenici", "EPP yüklenici") & _
                ": " & AKSORA_SAYI_METNI(adet(i), "#,##0") & " satýr"
        End If
    Next i
    sure = Timer - bas: If sure < 0 Then sure = sure + 86400#
    ThisWorkbook.names.Add Name:="_AKSORA_YATIRIM_SURE_" & il, RefersTo:="=" & Trim$(Str$(Round(sure, 2))), Visible:=False
    application.Calculation = hesap: application.EnableEvents = ev
    If ekran Then
        Y_PANEL_YENILE
        Y_ARAMA_TARIH
        On Error Resume Next
        Y_TEMINAT_KUR il
        On Error GoTo Hata
    End If
    application.ScreenUpdating = sc
    bilgi = Y_IL_ADI(il) & ": güncellendi (" & AKSORA_SAYI_METNI(sure, "0.0") & " sn)"
    Y_GUNCELLE = True
    If Not mYToplu Then Y_SERIT "Güncelleme tamamlandý  |  " & Y_SERIT_METNI(), 2
    If Not mYToplu And Not sessiz Then Y_PANEL_GOSTER
    If Not sessiz Then MsgBox "Yatýrým birimi | " & Y_IL_ADI(il) & " güncellendi." & vbCrLf & ozet & vbCrLf & vbCrLf & _
        "CDP / EPP yatýrým depolarý: D002, D011" & vbCrLf & _
        "Yüklenici depolarý: " & Join(AKSORA_IL_YUKLENICI(il), ", ") & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(sure, "0.00") & " sn", vbInformation, "AKSORA | Yatýrým Birimi"
Cikis:
    AKSORA_V11_MESGUL = mesgulOnce
    application.Calculation = hesap: application.EnableEvents = ev: application.ScreenUpdating = sc
    application.StatusBar = False
    mYCalisiyor = False
    Exit Function
Hata:
    hMesaj = "Hata " & CStr(Err.Number) & ": " & Err.Description
    bilgi = Y_IL_ADI(il) & ": güncellenemedi; " & hMesaj
    On Error Resume Next
    application.Calculation = hesap: application.EnableEvents = ev: application.ScreenUpdating = True
    If Not mYToplu Then Y_SERIT "Güncelleme tamamlanamadý  |  " & hMesaj, 3
    If Not sessiz Then MsgBox "Yatýrým stoklarý güncellenemedi." & vbCrLf & hMesaj, vbExclamation, "AKSORA | Yatýrým Birimi"
    On Error GoTo 0
    Resume Cikis
End Function

'---------- Ana panel ----------
Private Sub Y_OZET(ByVal sh As Worksheet, ByVal d As Object, ByRef satir As Long, ByRef deger As Double, ByRef tarih As Date)
    Dim son As Long, v As Variant, r As Long, t As Variant
    t = sh.Range("C2").Value
    If Not IsError(t) Then
        If IsDate(t) Then
            If CDate(t) > tarih Then tarih = CDate(t)
        End If
    End If
    son = AKSORA_M_SON(sh)
    If son < 7 Then Exit Sub
    v = sh.Range("A7:N" & son).Value2
    For r = 1 To UBound(v, 1)
        If Not IsError(v(r, 5)) Then
            If Len(Trim$(CStr(v(r, 5)))) > 0 Then
                satir = satir + 1
                If Not IsError(v(r, 2)) Then d(Trim$(CStr(v(r, 2)))) = True
                If Not IsError(v(r, 11)) Then
                    If VarType(v(r, 11)) <> vbString And IsNumeric(v(r, 11)) Then deger = deger + CDbl(v(r, 11))
                End If
            End If
        End If
    Next r
End Sub

Private Sub Y_KART_DURUM(ByVal ws As Worksheet, ByVal k As String, ByVal metin As String, ByVal renk As Long, ByVal koyu As Long)
    Dim s As Shape
    Set s = Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "0")
    If Not s Is Nothing Then
        Y_YAZ s, metin
        On Error Resume Next
        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = koyu
        On Error GoTo 0
    End If
    Set s = Y_SEKIL(ws, "AKSORA15_KART_LED_" & k)
    If Not s Is Nothing Then
        On Error Resume Next
        s.Fill.Visible = msoTrue: s.Fill.Solid: s.Fill.ForeColor.RGB = renk
        s.Glow.Color.RGB = renk: s.Glow.Radius = 4: s.Glow.Transparency = 0.5
        On Error GoTo 0
    End If
End Sub

Private Function Y_TL(ByVal x As Double) As String
    Y_TL = AKSORA_SAYI_METNI(x, "#,##0.00") & " TL"
End Function

' Serit metni: il, son guncelleme, sureler, satirlar ve toplam deger.
Private Function Y_SERIT_METNI() As String
    Dim adlar As Variant, i As Long, d As Object, satir As Long, deger As Double, tarih As Date
    Dim enSon As Date, toplam As Double, metin As String, il As String, t As String, sure As Double
    adlar = Y_STOK_ADLARI()
    il = Y_AKTIF_IL()
    For i = 0 To 3
        Set d = CreateObject("Scripting.Dictionary")
        satir = 0: deger = 0: tarih = 0
        Y_OZET ThisWorkbook.Worksheets(CStr(adlar(i))), d, satir, deger, tarih
        If tarih > enSon Then enSon = tarih
        toplam = toplam + deger
        metin = metin & "  |  " & Choose(i + 1, "CDP Yatýrým ", "EPP Yatýrým ", "CDP Yüklenici ", "EPP Yüklenici ") & _
            AKSORA_SAYI_METNI(satir, "#,##0") & " satýr"
    Next i
    If enSon = 0 Then
        Y_SERIT_METNI = "Henüz SAP güncellemesi yapýlmadý  |  CDP ve EPP yatýrým: D002, D011  |  Yüklenici: malzeme stok birimindeki yüklenici depolarý"
        Exit Function
    End If
    If AKSORA_TOPLU_SON_SURE("_AKSORA_YATIRIM_SON_SURE") > 0 Then _
        metin = "  |  Bütün iller " & AKSORA_SURE_METNI(AKSORA_TOPLU_SON_SURE("_AKSORA_YATIRIM_SON_SURE")) & metin
    On Error Resume Next
    t = ThisWorkbook.names("_AKSORA_YATIRIM_SURE_" & il).RefersTo
    On Error GoTo 0
    If Len(t) > 1 Then sure = Val(Mid$(t, 2))
    If sure > 0 Then metin = "  |  Ýl güncelleme süresi " & AKSORA_SURE_METNI(sure) & metin
    Y_SERIT_METNI = "Son güncelleme " & Format$(enSon, "dd.mm.yyyy hh:nn") & metin & _
        "  |  Toplam stok deðeri " & Y_TL(toplam)
End Function

' durum: 0 bilgi, 1 calisiyor, 2 tamam, 3 hata
' Seridin ustune binen IL DEGISTIR dugmesinin sol kenari (yoksa 0).
Private Function Y_SERIT_SAG(ByVal ws As Worksheet, ByVal s As Shape) As Double
    Dim b As Shape
    Set b = Y_SEKIL(ws, "AKSORA15_IL_DEGISTIR")
    If b Is Nothing Then Exit Function
    If b.Left > s.Left And b.Left < s.Left + s.Width And b.Top < s.Top + s.Height And b.Top + b.Height > s.Top Then Y_SERIT_SAG = b.Left
End Function

Private Sub Y_SERIT(ByVal metin As String, ByVal durum As Long)
    Dim ws As Worksheet, s As Shape, korumali As Boolean, pt As Double, gen As Double, ust As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Y_ANA PANEL")
    Set s = ws.Shapes("MUSTERI_ANA_SECIM_SERIDI")
    If s Is Nothing Then Exit Sub
    korumali = ws.ProtectDrawingObjects
    If korumali Then ws.Unprotect ""
    s.Fill.ForeColor.RGB = Choose(durum + 1, RGB(224, 242, 254), RGB(254, 243, 199), RGB(220, 252, 231), RGB(254, 226, 226))
    s.TextFrame2.TextRange.Text = metin
    If Val(s.AlternativeText) <= 0 Then s.AlternativeText = CStr(s.TextFrame2.TextRange.Font.Size)
    pt = Val(s.AlternativeText)
    gen = s.Width - s.TextFrame2.MarginLeft - s.TextFrame2.MarginRight - 8
    ust = Y_SERIT_SAG(ws, s)
    If ust > 0 Then gen = application.Min(gen, ust - s.Left - s.TextFrame2.MarginLeft - 8)
    Do While pt > 7 And Len(metin) * pt * 0.56 > gen
        pt = pt - 0.5
    Loop
    s.TextFrame2.TextRange.Font.Size = pt
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Choose(durum + 1, RGB(12, 45, 79), RGB(146, 64, 14), RGB(21, 128, 61), RGB(153, 27, 27))
    If korumali Then Y_KORU ws
    If durum = 1 Then DoEvents
    On Error GoTo 0
End Sub

' Guncelleme bitince ana panel ekrana getirilir.
Private Sub Y_PANEL_GOSTER()
    Dim ws As Worksheet, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Y_ANA PANEL")
    If ws Is Nothing Then Exit Sub
    ev = application.EnableEvents
    application.ScreenUpdating = False: application.EnableEvents = False
    mYonlendirme = True
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    mYonlendirme = False
    application.EnableEvents = ev
    AKSORA_YATIRIM_GORUNUM ws
    application.ScreenUpdating = True
    DoEvents
    ' Excel islem bitince baska sayfayi one getirirse panel tekrar secilir.
    mYOdak = ws.Name
    application.OnTime Now + TimeSerial(0, 0, 1), BIRIM_MAKRO("AKSORA_YATIRIM_ODAK")
    On Error GoTo 0
End Sub

Private Sub Y_SAYFA_SEC(ByVal ws As Worksheet)
    Dim ev As Boolean, sc As Boolean
    On Error Resume Next
    If ws Is Nothing Then Exit Sub
    If ActiveSheet Is ws Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    application.EnableEvents = False
    mYonlendirme = True
    ws.Select Replace:=True
    mYonlendirme = False
    application.EnableEvents = ev
    AKSORA_YATIRIM_GORUNUM ws
    application.ScreenUpdating = sc
    On Error GoTo 0
End Sub

Public Sub AKSORA_YATIRIM_ODAK()
    Dim ad As String
    ad = mYOdak: mYOdak = vbNullString
    If Len(ad) = 0 Or mYCalisiyor Or mYToplu Then Exit Sub
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name = ad Then Exit Sub
    If Not AKSORA_YATIRIM_SAYFA(ActiveSheet) Then Exit Sub
    Y_SAYFA_SEC ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
End Sub

' Sekmeler: Ana panel, Teminat, Hakedis, Malzeme arama, ardindan stok/rapor ciftleri.
Private Sub Y_SEKME_SIRASI()
    Dim adlar As Variant, i As Long, onceki As Worksheet, ws As Worksheet
    adlar = Y_ADLAR()
    For i = 1 To UBound(adlar)
        Set onceki = ThisWorkbook.Worksheets(CStr(adlar(i - 1)))
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i)))
        If ws.Index <> onceki.Index + 1 Then ws.Move after:=onceki
    Next i
End Sub

' Rapor alt bolumu: depo secim listesi baslik satirinda tablo basligina binmez.
Public Sub AKSORA_YATIRIM_DM_DUZEN(ByVal r As Worksheet)
    Dim bas As Long, ust As Double, eskiH As Double, fark As Double, s As Shape, t As String
    bas = Val(CStr(r.Range("BX31").Value2))
    If bas < 9 Then Exit Sub
    ust = r.Rows(bas).Top: eskiH = r.Rows(bas).Height
    ' Form listesi dusuk yakinlastirmada en az ~22 piksel cizilir; satir bu payi tasir,
    ' liste satirin ustune yaslanir ve tablo basligina asla inmez.
    If eskiH < 56 Then
        fark = 56 - eskiH
        For Each s In r.Shapes
            If Left$(s.Name, 12) = "AKSORA_RP_DM" And s.Top >= ust + eskiH - 1 Then s.Top = s.Top + fark
        Next s
        r.Rows(bas).RowHeight = 56
    End If
    For Each s In r.Shapes
        t = s.Name
        If t = "AKSORA_RP_DM_BASLIK" Or t = "AKSORA_RP_DM_SECIM_ETIKET" Then
            s.Top = ust + 2: s.Height = 24
            On Error Resume Next
            s.TextFrame2.VerticalAnchor = msoAnchorMiddle
            On Error GoTo 0
        ElseIf t = "AKSORA_RP_DM_LISTE" Then
            s.Top = ust + 4: s.Height = 18
        ElseIf t = "AKSORA_RP_DM_ZEMIN" Then
            If s.Top < ust + 1 And s.Top > ust - 1 Then s.Height = s.Height + fark
        End If
    Next s
End Sub

Private Sub Y_PANEL_YENILE()
    Dim ws As Worksheet, kartlar As Variant, adlar As Variant, i As Long, k As String, d As Object
    Dim satir As Long, deger As Double, tarih As Date, korumali As Boolean
    Set ws = ThisWorkbook.Worksheets("Y_ANA PANEL")
    korumali = ws.ProtectContents Or ws.ProtectDrawingObjects
    ws.Unprotect ""
    adlar = Y_STOK_ADLARI()
    kartlar = Array("B", "L", "Q")
    For i = 0 To 2
        k = CStr(kartlar(i))
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
        satir = 0: deger = 0: tarih = 0
        If i < 2 Then
            Y_OZET ThisWorkbook.Worksheets(CStr(adlar(i))), d, satir, deger, tarih
        Else
            Y_OZET ThisWorkbook.Worksheets(CStr(adlar(2))), d, satir, deger, tarih
            Y_OZET ThisWorkbook.Worksheets(CStr(adlar(3))), d, satir, deger, tarih
        End If
        If tarih > 0 Then
            Y_KART_DURUM ws, k, "GÜNCEL", RGB(34, 197, 94), RGB(21, 128, 61)
        Else
            Y_KART_DURUM ws, k, "Veri bekleniyor", RGB(145, 162, 176), RGB(29, 78, 216)
        End If
        If tarih > 0 Then
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "1"), Format$(tarih, "dd.mm.yyyy hh:nn")
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "2"), AKSORA_SAYI_METNI(satir, "#,##0")
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "3"), Y_TL(deger)
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "4"), AKSORA_SAYI_METNI(d.Count, "#,##0")
        Else
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "1"), "-"
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "2"), "-"
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "3"), "-"
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "4"), "-"
        End If
    Next i
    Y_YAZ Y_SEKIL(ws, "AKSORA15_PANEL_TARIHI"), "Yatýrým Birimi  |  Panel kontrolü: " & Format$(Now, "dd.mm.yyyy hh:nn")
    If korumali Then Y_KORU ws
    Y_SERIT Y_SERIT_METNI(), 0
End Sub

'---------- Malzeme arama (yalniz yatirim verisi) ----------
Private Sub Y_ARAMA_TEMIZLE(ByVal ws As Worksheet)
    Dim son As Long, s As Shape
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    son = application.Max(13, AKSORA_M_SON(ws))
    ws.Range("A13:N" & son).ClearContents
    ws.Range("B6").MergeArea.ClearContents
    ws.Range(ws.Cells(13, 75), ws.Cells(son, 75)).ClearContents
    AKSORA_ARAMA_FAZLA_SATIR ws, 13
    ws.Range("BX21").Value2 = "TUMU"
    Y_BAGLA Y_SEKIL(ws, "btnMalzemeAnaPanel"), "Y_ANA PANEL"
    Set s = Y_SEKIL(ws, "btnMalzemeAra")
    If Not s Is Nothing Then s.OnAction = BIRIM_MAKRO("AKSORA_MUSTERI_ARAMA_HAZIRLIK")
    Set s = Y_SEKIL(ws, "AKSORA_TOPLU_GUNCELLE")
    If Not s Is Nothing Then
        Y_YAZ s, "BÜTÜN ÝLLERÝ GÜNCELLE" & vbLf & "MB52  |  5 ÝL"
        s.OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_BUTUN_ILLER")
    End If
    Set s = Y_SEKIL(ws, "AKSORA_TOPLU_VURGU")
    If Not s Is Nothing Then s.OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_BUTUN_ILLER")
    Y_YAZ Y_SEKIL(ws, "AKSORA_TOPLU_DURUM"), "Arama, her ilin yukarýdaki tarihli yatýrým stoklarýný kullanýr."
    Y_YAZ Y_SEKIL(ws, "AKSORA_ARAMA_SABIT_ACIKLAMA"), "Yalnýzca yatýrým birimi stoklarý aranýr."
End Sub

Private Sub Y_ARAMA_TARIH()
    Dim ws As Worksheet, s As Shape, kod As String, isim As String, t As Date, korumali As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Y_MALZEME_ARA")
    If ws Is Nothing Then Exit Sub
    korumali = ws.ProtectContents Or ws.ProtectDrawingObjects
    If korumali Then ws.Unprotect ""
    For Each s In ws.Shapes
        kod = s.AlternativeText: isim = vbNullString
        Select Case kod
            Case "GIRESUN": isim = "GÝRESUN"
            Case "TRABZON": isim = "TRABZON"
            Case "RIZE": isim = "RÝZE"
            Case "GUMUSHANE": isim = "GÜMÜÞHANE"
            Case "ARTVIN": isim = "ARTVÝN"
        End Select
        If Len(isim) > 0 And Left$(s.Name, 10) = "AKSORA_V4_" Then
            t = AKSORA_YATIRIM_IL_SON_TARIH(kod)
            s.TextFrame2.TextRange.Text = isim & vbLf & IIf(t > 0, Format$(t, "dd.mm.yyyy hh:nn"), "Henüz güncellenmedi")
        End If
    Next s
    If korumali Then Y_KORU ws
    On Error GoTo 0
End Sub

Private Sub Y_TOPLU_DURUM(ByVal metin As String)
    Dim ws As Worksheet, s As Shape, korumali As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Y_MALZEME_ARA")
    Set s = ws.Shapes("AKSORA_TOPLU_DURUM")
    If s Is Nothing Then Exit Sub
    korumali = ws.ProtectDrawingObjects
    If korumali Then ws.Unprotect ""
    s.TextFrame2.TextRange.Text = metin
    If korumali Then Y_KORU ws
    application.StatusBar = "AKSORA | Yatýrým | " & Replace(metin, vbLf, " | ")
    DoEvents
    On Error GoTo 0
End Sub

' Bes ilin yatirim ve yuklenici stoklari sirayla alinir; ekrandaki il sayfalara, digerleri kendi kaydina yazilir.
Public Sub AKSORA_YATIRIM_BUTUN_ILLER()
    Dim il As Variant, bilgi As String, sonuc As String, adet As Long, bas As Double, sure As Double
    Dim hataMetni As String, metin As String, sc As Boolean, baslangic As Worksheet
    If AKSORA_V11_MESGUL Or mYCalisiyor Or mYToplu Then Exit Sub
    If MsgBox("Beþ ilin (Artvin, Trabzon, Gümüþhane, Giresun, Rize) yatýrým ve yüklenici stoklarý SAP'den güncellenecek. Devam edilsin mi?", _
        vbYesNo + vbQuestion, "AKSORA | Yatýrým Birimi") <> vbYes Then Exit Sub
    sc = application.ScreenUpdating
    If TypeOf ActiveSheet Is Worksheet Then Set baslangic = ActiveSheet
    mYToplu = True: bas = Timer
    application.ScreenUpdating = True
    On Error GoTo Hata
    For Each il In Array("ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
        bilgi = vbNullString
        Y_TOPLU_DURUM Y_IL_ADI(CStr(il)) & " güncelleniyor..."
        If Y_GUNCELLE(CStr(il), "1234", True, bilgi) Then adet = adet + 1
        sonuc = sonuc & bilgi & vbCrLf
        Y_SAYFA_SEC baslangic
    Next il
Bitir:
    On Error Resume Next
    mYToplu = False
    sure = Timer - bas: If sure < 0 Then sure = sure + 86400#
    ThisWorkbook.names.Add Name:="_AKSORA_YATIRIM_SON_SURE", RefersTo:="=" & Trim$(Str$(Round(sure, 2))), Visible:=False
    metin = IIf(Len(hataMetni) > 0, "Ýþlem durduruldu. ", "Ýþlem tamamlandý. ") & adet & "/5 il güncellendi." & vbLf & _
        "Geçen süre: " & Format$(CLng(Fix(sure)) \ 60, "00") & ":" & Format$(CLng(Fix(sure)) Mod 60, "00")
    Y_TOPLU_DURUM metin
    Y_PANEL_YENILE
    Y_ARAMA_TARIH
    application.StatusBar = False
    Y_PANEL_GOSTER
    application.ScreenUpdating = sc
    On Error GoTo 0
    MsgBox "Yatýrým birimi | Bütün iller" & vbCrLf & vbCrLf & sonuc & IIf(Len(hataMetni) > 0, vbCrLf & hataMetni & vbCrLf, "") & vbCrLf & _
        "Toplam süre: " & AKSORA_SURE_METNI(sure), IIf(adet = 5, vbInformation, vbExclamation), "AKSORA | Yatýrým Birimi"
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Bitir
End Sub

'==========================================================
' YATIRIM - TEMINAT TAKIBI VE OTOMATIK HAKEDIS
' Ortak: rapor sayfalariyla ayni ust bant, numarali KPI kartlari, lacivert tablo
' basliklari. Giris hucreleri acik sari; hesaplar canli formul. Girilen degerler
' il ve yuklenici bazinda gizli kayit sayfalarinda saklanir (_Y_TEMINAT, _Y_HAKEDIS).
'==========================================================


Private Function Y_GIRIS_RENK() As Long
    Y_GIRIS_RENK = RGB(255, 249, 219)
End Function

' Gizli anahtar-deger kayit sayfasi (A: anahtar, B..: degerler).
Private Function Y_KAYIT_SAYFA(ByVal ad As String) As Worksheet
    Dim ws As Worksheet, onceki As Object, sc As Boolean, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
    If ws Is Nothing Then
        sc = application.ScreenUpdating: ev = application.EnableEvents
        application.ScreenUpdating = False: application.EnableEvents = False
        Set onceki = ActiveSheet
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ad
        ws.Visible = xlSheetVeryHidden
        If Not onceki Is Nothing Then onceki.Activate
        application.EnableEvents = ev: application.ScreenUpdating = sc
    End If
    Set Y_KAYIT_SAYFA = ws
End Function

Private Function Y_KAYIT_SATIR(ByVal ws As Worksheet, ByVal anahtar As String, ByVal olustur As Boolean) As Long
    Dim m As Variant, son As Long
    son = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If son >= 1 And Len(CStr(ws.Cells(1, 1).Value2)) > 0 Then
        m = application.Match(anahtar, ws.Range("A1:A" & son), 0)
        If Not IsError(m) Then Y_KAYIT_SATIR = CLng(m): Exit Function
    Else
        son = 0
    End If
    If olustur Then
        ws.Cells(son + 1, 1).NumberFormat = "@"
        ws.Cells(son + 1, 1).Value2 = anahtar
        Y_KAYIT_SATIR = son + 1
    End If
End Function

Private Function Y_KOD(ByVal v As Variant) As String
    Select Case VarType(v)
        Case vbEmpty, vbNull: Y_KOD = vbNullString
        Case vbString: Y_KOD = "s" & Replace(Replace(Replace(CStr(v), vbTab, " "), vbLf, " "), vbCr, " ")
        Case vbBoolean: Y_KOD = IIf(v, "b1", "b0")
        Case vbError: Y_KOD = vbNullString
        Case Else: Y_KOD = "n" & Trim$(Str$(CDbl(v)))
    End Select
End Function

Private Function Y_COZ(ByVal t As String) As Variant
    If Len(t) = 0 Then Exit Function
    Select Case Left$(t, 1)
        Case "s": Y_COZ = Mid$(t, 2)
        Case "n": Y_COZ = Val(Mid$(t, 2))
        Case "b": Y_COZ = (Mid$(t, 2) = "1")
    End Select
End Function

' Sutun genislikleri nokta cinsinden.
Private Sub Y_KOLONLAR(ByVal ws As Worksheet, ByVal gen As Variant)
    Dim i As Long, k As Long, c As Range
    For k = 1 To 2
        For i = 0 To UBound(gen)
            Set c = ws.Columns(i + 1)
            If c.ColumnWidth > 0 And c.Width > 0 Then
                c.ColumnWidth = application.Min(255, c.ColumnWidth * CDbl(gen(i)) / c.Width)
            Else
                c.ColumnWidth = CDbl(gen(i)) / 5.6
            End If
        Next i
    Next k
End Sub

Private Sub Y_SAYFA_SIFIRLA(ByVal ws As Worksheet)
    Dim s As Shape, ad As String, sil As Collection, sv As Variant
    Set sil = New Collection
    For Each s In ws.Shapes
        ad = s.Name
        If Not (ad = "AKT1_RAPOR_UST_ZEMIN" Or ad = "AKT1_RAPOR_BASLIK" Or ad = "btnAnaPanelDon" Or InStr(ad, "FILIGRAN") > 0) Then sil.Add ad
    Next s
    On Error Resume Next
    For Each sv In sil
        ws.Shapes(CStr(sv)).Delete
    Next sv
    On Error GoTo 0
    ws.Cells.FormatConditions.Delete
    ws.Range("A3:Z400").UnMerge
    ws.Range("A3:Z400").Clear
    ws.Range("BZ1:BZ40").Clear
    ws.Range(ws.Rows(3), ws.Rows(400)).RowHeight = ws.StandardHeight
    ws.Columns("L:Z").ColumnWidth = 8.43
    ws.Columns("BZ").Hidden = True
    Y_BAGLA Y_SEKIL(ws, "btnAnaPanelDon"), "Y_ANA PANEL"
End Sub

' Rapor sayfalarindaki numarali kart: rozet + etiket + hucreye bagli deger.
Private Sub Y_KPI(ByVal ws As Worksheet, ByVal i As Long, ByVal x As Double, ByVal y As Double, ByVal w As Double, _
    ByVal h As Double, ByVal etiket As String, ByVal hucre As String, ByVal renk As Long)
    Dim s As Shape
    Set s = Y_KUTU(ws, "Y_KPI_KART_" & i, x, y, w, h, RGB(242, 249, 255), RGB(196, 220, 240))
    Set s = Y_KUTU(ws, "Y_KPI_IKON_" & i, x + 8, y + 9, 40, h - 18, renk, -1)
    Y_YAZ s, Format$(i, "00")
    With s.TextFrame2
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 13: .TextRange.Font.bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter: .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 0: .MarginRight = 0
    End With
    Y_METIN ws, "Y_KPI_ETIKET_" & i, etiket, x + 58, y + 8, w - 66, 18, 10, RGB(28, 54, 83), True, msoAlignLeft
    Set s = Y_METIN(ws, "Y_KPI_DEGER_" & i, "-", x + 58, y + 30, w - 66, 28, 17, RGB(22, 58, 96), True, msoAlignLeft)
    s.DrawingObject.Formula = "='" & Replace(ws.Name, "'", "''") & "'!" & hucre
End Sub

Private Sub Y_BASLIK_SATIRI(ByVal r As Range, ByVal basliklar As Variant)
    Dim i As Long
    For i = 0 To UBound(basliklar)
        r.Cells(1, i + 1).Value2 = basliklar(i)
    Next i
    With r
        .Interior.Color = RGB(26, 75, 120): .Font.Color = vbWhite
        .Font.Name = "Aptos": .Font.Size = 10: .Font.bold = True
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
End Sub

Private Sub Y_TABLO_GOVDE(ByVal r As Range)
    Dim i As Long
    With r
        .Font.Name = "Aptos": .Font.Size = 10: .Font.Color = RGB(30, 41, 59)
        .VerticalAlignment = xlCenter
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(214, 224, 236)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(214, 224, 236)
        End With
    End With
    For i = 2 To r.Rows.Count Step 2
        r.Rows(i).Interior.Color = RGB(244, 248, 252)
    Next i
End Sub

Private Sub Y_GIRIS_HUCRE(ByVal r As Range)
    r.Interior.Color = Y_GIRIS_RENK()
    r.Locked = False
    With r.Borders
        .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(230, 200, 120)
    End With
End Sub

Private Sub Y_TOPLAM_SATIRI(ByVal r As Range)
    With r
        .Font.bold = True: .Interior.Color = RGB(226, 236, 247)
        .Font.Name = "Aptos": .Font.Size = 10: .VerticalAlignment = xlCenter
        With .Borders(xlEdgeTop)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(26, 75, 120)
        End With
        With .Borders(xlEdgeBottom)
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(26, 75, 120)
        End With
    End With
End Sub

Private Function Y_DUGME(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, ByVal x As Double, ByVal y As Double, _
    ByVal w As Double, ByVal h As Double, ByVal renk As Long, ByVal makro As String) As Shape
    Dim s As Shape
    Set s = Y_KUTU(ws, ad, x, y, w, h, renk, -1)
    Y_YAZ s, metin
    With s.TextFrame2
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 10: .TextRange.Font.bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter: .VerticalAnchor = msoAnchorMiddle
    End With
    s.OnAction = BIRIM_MAKRO(makro)
    Set Y_DUGME = s
End Function

Private Sub Y_DURUM_KF(ByVal r As Range, ByVal metin As String, ByVal zemin As Long, ByVal yazi As Long)
    Dim fc As FormatCondition
    Set fc = r.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & metin & """")
    fc.Interior.Color = zemin: fc.Font.Color = yazi: fc.Font.bold = True
End Sub

' Ilin yuklenici depolari: depo -> (ad, CDP degeri, EPP degeri). Liste sirasi malzeme biriminin sirasi.
Private Function Y_YUKLENICI_OZET(ByVal il As String) As Object
    Dim d As Object, x As Variant, i As Long, sh As Worksheet, son As Long, v As Variant, r As Long, dp As String, t As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For Each x In AKSORA_IL_YUKLENICI(il)
        d(CStr(x)) = Array(vbNullString, 0#, 0#)
    Next x
    For i = 3 To 4
        Set sh = ThisWorkbook.Worksheets(CStr(Y_STOK_ADLARI()(i - 1)))
        son = AKSORA_M_SON(sh)
        If son >= 7 Then
            v = sh.Range("A7:N" & son).Value2
            For r = 1 To UBound(v, 1)
                If Not IsError(v(r, 2)) And Not IsError(v(r, 5)) Then
                    dp = UCase$(Trim$(CStr(v(r, 2))))
                    If Len(dp) > 0 And Len(Trim$(CStr(v(r, 5)))) > 0 Then
                        If Not d.Exists(dp) Then d(dp) = Array(vbNullString, 0#, 0#)
                        t = d(dp)
                        If Len(t(0)) = 0 And Not IsError(v(r, 3)) Then t(0) = Trim$(CStr(v(r, 3)))
                        If Not IsError(v(r, 11)) Then
                            If VarType(v(r, 11)) <> vbString And IsNumeric(v(r, 11)) Then t(i - 2) = t(i - 2) + CDbl(v(r, 11))
                        End If
                        d(dp) = t
                    End If
                End If
            Next r
        End If
    Next i
    Set Y_YUKLENICI_OZET = d
End Function

'---------- TEMINAT ----------
Private Sub Y_TEMINAT_KAYDET(ByVal ws As Worksheet)
    Dim kayit As Worksheet, il As String, r As Long, son As Long, dp As String, sr As Long
    il = UCase$(Trim$(CStr(ws.Range("BX22").Value2)))
    If Len(il) = 0 Then Exit Sub
    son = Val(CStr(ws.Range("BX23").Value2))
    If son < YT_ILK Then Exit Sub
    Set kayit = Y_KAYIT_SAYFA("_Y_TEMINAT")
    For r = YT_ILK To son
        dp = UCase$(Trim$(CStr(ws.Cells(r, 1).Value2)))
        If Len(dp) > 0 Then
            sr = Y_KAYIT_SATIR(kayit, il & "|" & dp, True)
            kayit.Cells(sr, 2).Value2 = ws.Cells(r, 6).Value2
            kayit.Cells(sr, 3).Value = Now
        End If
    Next r
End Sub

Private Function Y_TEMINAT_DEGER(ByVal il As String, ByVal dp As String) As Variant
    Dim kayit As Worksheet, sr As Long
    Set kayit = Y_KAYIT_SAYFA("_Y_TEMINAT")
    sr = Y_KAYIT_SATIR(kayit, il & "|" & dp, False)
    If sr > 0 Then Y_TEMINAT_DEGER = kayit.Cells(sr, 2).Value2
End Function

Private Sub Y_TEMINAT_KUR(ByVal il As String)
    Dim ws As Worksheet, d As Object, k As Variant, n As Long, r As Long, son As Long, t As Variant
    Dim w As Double, x As Double, kw As Double, i As Long, y As Double, s As Shape, tarih As Date, t2 As Variant
    Dim rr As Range, db As Databar
    Set ws = ThisWorkbook.Worksheets("Y_TEMÝNAT")
    il = UCase$(Trim$(il))
    ws.Unprotect ""
    Y_TEMINAT_KAYDET ws
    Y_SAYFA_SIFIRLA ws
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    Y_KOLONLAR ws, Array(70, w - 1244, 150, 150, 160, 170, 160, 214, 170)
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10
    ws.Rows(8).RowHeight = 30
    ws.Rows(9).RowHeight = 30
    Set d = Y_YUKLENICI_OZET(il)
    n = d.Count
    son = YT_ILK + application.Max(n, 1) - 1
    ws.Range(ws.Rows(YT_ILK), ws.Rows(son + 1)).RowHeight = 22
    ws.Rows(son + 2).RowHeight = 12
    ws.Range(ws.Rows(son + 3), ws.Rows(son + 6)).RowHeight = 18
    ' Baslik ve tablo
    ws.Range("A8").Value2 = "YÜKLENÝCÝ BAZINDA TEMÝNAT  –  DEPO KARÞILAÞTIRMASI"
    With ws.Range("A8:E8")
        .Merge: .Font.Name = "Aptos": .Font.Size = 13: .Font.bold = True: .Font.Color = RGB(12, 45, 79)
        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlLeft
    End With
    ws.Range("F8").Value2 = "Sarý hücrelere yüklenicinin teminat tutarýný girin; hesaplar anýnda güncellenir."
    With ws.Range("F8:I8")
        .Merge: .Font.Name = "Aptos": .Font.Size = 9: .Font.Italic = True: .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlRight
    End With
    Y_BASLIK_SATIRI ws.Range("A9:I9"), Array("Depo", "Yüklenici / Depo Adý", "CDP Depo Deðeri", "EPP Depo Deðeri", _
        "Depodaki Malzeme", "Teminat Tutarý", "Kalan Teminat", "Teminat Kullanýmý", "Durum")
    r = YT_ILK
    For Each k In d.keys
        t = d(k)
        ws.Cells(r, 1).NumberFormat = "@": ws.Cells(r, 1).Value2 = CStr(k)
        ws.Cells(r, 2).Value2 = IIf(Len(t(0)) > 0, t(0), "(depoda stok yok)")
        ws.Cells(r, 3).Value2 = t(1): ws.Cells(r, 4).Value2 = t(2)
        ws.Cells(r, 5).Formula = "=C" & r & "+D" & r
        t2 = Y_TEMINAT_DEGER(il, CStr(k))
        If Not IsEmpty(t2) Then ws.Cells(r, 6).Value2 = t2
        ws.Cells(r, 7).Formula = "=IF(N(F" & r & ")>0,F" & r & "-E" & r & ","""")"
        ws.Cells(r, 8).Formula = "=IF(N(F" & r & ")>0,E" & r & "/F" & r & ","""")"
        ws.Cells(r, 9).Formula = "=IF(N(F" & r & ")<=0,""Teminat girilmedi"",IF(H" & r & ">1,""AÞIM"",IF(H" & r & ">=0.8,""DÝKKAT"",""GÜVENLÝ"")))"
        r = r + 1
    Next k
    If n = 0 Then ws.Cells(YT_ILK, 2).Value2 = "Bu ilde tanýmlý yüklenici deposu yok."
    Set rr = ws.Range(ws.Cells(YT_ILK, 1), ws.Cells(son, 9))
    Y_TABLO_GOVDE rr
    ws.Range(ws.Cells(YT_ILK, 1), ws.Cells(son, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(YT_ILK, 2), ws.Cells(son, 2)).IndentLevel = 1
    ws.Range(ws.Cells(YT_ILK, 3), ws.Cells(son + 1, 7)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(YT_ILK, 3), ws.Cells(son + 1, 7)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(YT_ILK, 8), ws.Cells(son + 1, 8)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(YT_ILK, 8), ws.Cells(son + 1, 9)).HorizontalAlignment = xlCenter
    If n > 0 Then Y_GIRIS_HUCRE ws.Range(ws.Cells(YT_ILK, 6), ws.Cells(son, 6))
    ' Toplam
    ws.Cells(son + 1, 1).Value2 = "TOPLAM"
    ws.Cells(son + 1, 3).Formula = "=SUM(C" & YT_ILK & ":C" & son & ")"
    ws.Cells(son + 1, 4).Formula = "=SUM(D" & YT_ILK & ":D" & son & ")"
    ws.Cells(son + 1, 5).Formula = "=SUM(E" & YT_ILK & ":E" & son & ")"
    ws.Cells(son + 1, 6).Formula = "=SUM(F" & YT_ILK & ":F" & son & ")"
    ws.Cells(son + 1, 7).Formula = "=IF(F" & son + 1 & ">0,F" & son + 1 & "-SUMIF(F" & YT_ILK & ":F" & son & ","">0"",E" & YT_ILK & ":E" & son & "),"""")"
    ws.Cells(son + 1, 8).Formula = "=IF(F" & son + 1 & ">0,SUMIF(F" & YT_ILK & ":F" & son & ","">0"",E" & YT_ILK & ":E" & son & ")/F" & son + 1 & ","""")"
    Y_TOPLAM_SATIRI ws.Range(ws.Cells(son + 1, 1), ws.Cells(son + 1, 9))
    ws.Cells(son + 1, 6).NumberFormat = "#,##0.00 ""TL"""
    ' Kullanim cubugu ve durum renkleri
    Set rr = ws.Range(ws.Cells(YT_ILK, 8), ws.Cells(son, 8))
    On Error Resume Next
    Set db = rr.FormatConditions.AddDatabar
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1.2
    db.BarColor.Color = RGB(47, 144, 217)
    db.BarFillType = xlDataBarFillSolid
    On Error GoTo 0
    Set rr = ws.Range(ws.Cells(YT_ILK, 9), ws.Cells(son, 9))
    Y_DURUM_KF rr, "AÞIM", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF rr, "DÝKKAT", RGB(254, 243, 199), RGB(180, 83, 9)
    Y_DURUM_KF rr, "GÜVENLÝ", RGB(220, 252, 231), RGB(21, 128, 61)
    Y_DURUM_KF rr, "Teminat girilmedi", RGB(241, 245, 249), RGB(100, 116, 139)
    ' Aciklama
    tarih = 0
    t2 = ThisWorkbook.Worksheets(CStr(Y_STOK_ADLARI()(2))).Range("C2").Value
    If Not IsError(t2) Then If IsDate(t2) Then tarih = CDate(t2)
    ws.Cells(son + 3, 1).Value2 = "Teminat kullanýmý = Depodaki malzeme deðeri ÷ Teminat tutarý.   %80 ve üzeri DÝKKAT, %100 üzeri teminat AÞIMI."
    ws.Cells(son + 4, 1).Value2 = "Depo deðerleri SAP MB52'den alýnýr (CDP + EPP yüklenici depolarý" & _
        IIf(tarih > 0, ", son güncelleme " & Format$(tarih, "dd.mm.yyyy hh:nn"), ", henüz güncellenmedi") & ")."
    ws.Cells(son + 5, 1).Value2 = "Girilen teminat tutarlarý il ve yüklenici bazýnda saklanýr; il deðiþtirilince o ilin tutarlarý gelir."
    With ws.Range(ws.Cells(son + 3, 1), ws.Cells(son + 5, 9))
        .Font.Name = "Aptos": .Font.Size = 9: .Font.Color = RGB(71, 85, 105)
    End With
    ' KPI hucreleri ve kartlar
    ws.Range("BZ1").Formula = "=F" & son + 1
    ws.Range("BZ2").Formula = "=E" & son + 1
    ws.Range("BZ3").Formula = "=IF(F" & son + 1 & ">0,G" & son + 1 & ",""Teminat girilmedi"")"
    ws.Range("BZ4").Formula = "=COUNTIF(I" & YT_ILK & ":I" & son & ",""AÞIM"")&"" aþým   •   ""&COUNTIF(I" & YT_ILK & ":I" & son & ",""DÝKKAT"")&"" dikkat   /   " & n & " yüklenici"""
    ws.Range("BZ1:BZ3").NumberFormat = "#,##0.00 ""TL"""
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "TOPLAM TEMÝNAT", "$BZ$1", RGB(35, 119, 204)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "DEPODAKÝ MALZEME (YÜKLENÝCÝ)", "$BZ$2", RGB(39, 160, 94)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "KALAN TEMÝNAT", "$BZ$3", RGB(35, 119, 204)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "RÝSK DURUMU", "$BZ$4", RGB(220, 38, 38)
    Set s = Y_SEKIL(ws, "Y_KPI_DEGER_4")
    If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 14
    ws.Range("BX22").Value2 = il
    ws.Range("BX23").Value2 = son
    ws.Range("BX20").Value2 = il
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | TEMÝNAT TAKÝBÝ  |  " & Y_IL_ADI(il)
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlUnlockedCells
End Sub

'---------- HAKEDIS ----------
Private Function Y_HK_ANAHTAR(ByVal ws As Worksheet) As String
    Y_HK_ANAHTAR = UCase$(Trim$(CStr(ws.Range("BX22").Value2))) & "|" & UCase$(Trim$(CStr(ws.Range("BX24").Value2)))
End Function

' Sayfadaki girisler kayda yazilir (ust bilgi B..M, kalemler N).
Private Sub Y_HAKEDIS_KAYDET(ByVal ws As Worksheet)
    Dim kayit As Worksheet, sr As Long, anahtar As String, r As Long, c As Long, alan(0 To 6) As String, satirlar() As String
    Dim ust As Variant, i As Long
    If Len(CStr(ws.Range("BX24").Value2)) = 0 Or Len(CStr(ws.Range("BX22").Value2)) = 0 Then Exit Sub
    anahtar = Y_HK_ANAHTAR(ws)
    Set kayit = Y_KAYIT_SAYFA("_Y_HAKEDIS")
    sr = Y_KAYIT_SATIR(kayit, anahtar, True)
    ust = Array("C10", "C11", "C12", "C13", "K10", "J12", "J13", "J14", "J15", "K16", "K17")
    For i = 0 To UBound(ust)
        kayit.Cells(sr, i + 2).NumberFormat = "@"
        kayit.Cells(sr, i + 2).Value2 = Y_KOD(ws.Range(CStr(ust(i))).Value2)
    Next i
    ReDim satirlar(0 To YH_KALEM_ADET - 1)
    For r = 0 To YH_KALEM_ADET - 1
        For c = 0 To 6
            alan(c) = Y_KOD(ws.Cells(YH_KALEM_ILK + r, c + 1).Value2)
        Next c
        satirlar(r) = Join(alan, vbTab)
    Next r
    kayit.Cells(sr, 14).NumberFormat = "@"
    kayit.Cells(sr, 14).Value2 = Join(satirlar, vbLf)
End Sub

Private Sub Y_HAKEDIS_YUKLE(ByVal ws As Worksheet)
    Dim kayit As Worksheet, sr As Long, ust As Variant, i As Long, satirlar() As String, alan() As String, r As Long, c As Long
    ust = Array("C10", "C11", "C12", "C13", "K10", "J12", "J13", "J14", "J15", "K16", "K17")
    ws.Range(ws.Cells(YH_KALEM_ILK, 1), ws.Cells(YH_KALEM_ILK + YH_KALEM_ADET - 1, 7)).ClearContents
    For i = 0 To UBound(ust): ws.Range(CStr(ust(i))).ClearContents: Next i
    Set kayit = Y_KAYIT_SAYFA("_Y_HAKEDIS")
    sr = Y_KAYIT_SATIR(kayit, Y_HK_ANAHTAR(ws), False)
    If sr = 0 Then
        ' Yeni yuklenici: varsayilan oranlar
        ws.Range("C12").Value2 = 1
        ws.Range("J12").Value2 = 0.2
        ws.Range("J13").Value2 = 0.00948
        ws.Range("J14").Value2 = 0
        ws.Range("J15").Value2 = 0
        Exit Sub
    End If
    For i = 0 To UBound(ust)
        ws.Range(CStr(ust(i))).Value2 = Y_COZ(CStr(kayit.Cells(sr, i + 2).Value2))
    Next i
    satirlar = Split(CStr(kayit.Cells(sr, 14).Value2), vbLf)
    For r = 0 To application.Min(UBound(satirlar), YH_KALEM_ADET - 1)
        alan = Split(satirlar(r), vbTab)
        For c = 0 To application.Min(UBound(alan), 6)
            ws.Cells(YH_KALEM_ILK + r, c + 1).Value2 = Y_COZ(alan(c))
        Next c
    Next r
End Sub

' Gecmis tablosu (secili yuklenicinin kayitli hakedisleri).
Private Sub Y_HAKEDIS_GECMIS_YAZ(ByVal ws As Worksheet)
    Dim g As Worksheet, son As Long, v As Variant, r As Long, k As Long, anahtar As String
    ws.Range(ws.Cells(YH_GECMIS_ILK, 1), ws.Cells(YH_GECMIS_ILK + 11, 11)).ClearContents
    Set g = Y_KAYIT_SAYFA("_Y_HAKEDIS_GECMIS")
    anahtar = Y_HK_ANAHTAR(ws)
    son = g.Cells(g.Rows.Count, 1).End(xlUp).Row
    If son >= 1 And Len(CStr(g.Cells(1, 1).Value2)) > 0 Then
        v = g.Range("A1:G" & son).Value2
        For r = UBound(v, 1) To 1 Step -1
            If CStr(v(r, 1)) = anahtar Then
                ws.Cells(YH_GECMIS_ILK + k, 1).Value2 = v(r, 2)
                ws.Cells(YH_GECMIS_ILK + k, 2).Value2 = v(r, 3)
                ws.Cells(YH_GECMIS_ILK + k, 4).Value2 = v(r, 4)
                ws.Cells(YH_GECMIS_ILK + k, 9).Value2 = v(r, 5)
                ws.Cells(YH_GECMIS_ILK + k, 10).Value2 = v(r, 6)
                ws.Cells(YH_GECMIS_ILK + k, 11).Value2 = v(r, 7)
                k = k + 1
                If k >= 12 Then Exit For
            End If
        Next r
    End If
    If k = 0 Then ws.Cells(YH_GECMIS_ILK, 2).Value2 = "Bu yüklenici için kayýtlý hakediþ yok."
End Sub

Private Function Y_HK_LISTE(ByVal il As String) As Variant
    Dim d As Object, k As Variant, a() As String, i As Long, t As Variant
    Set d = Y_YUKLENICI_OZET(il)
    If d.Count = 0 Then Exit Function
    ReDim a(0 To d.Count - 1)
    For Each k In d.keys
        t = d(k)
        a(i) = CStr(k) & IIf(Len(t(0)) > 0, "  -  " & t(0), "")
        i = i + 1
    Next k
    Y_HK_LISTE = a
End Function

Private Sub Y_HAKEDIS_KUR(ByVal il As String)
    Dim ws As Worksheet, w As Double, kw As Double, y As Double, s As Shape, liste As Shape, a As Variant, i As Long, p As Long
    Dim secim As String, r As Long, son As Long, rr As Range, etiketler As Variant, gen As Variant
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    il = UCase$(Trim$(il))
    Y_KAYIT_SAYFA "_Y_HAKEDIS"
    Y_KAYIT_SAYFA "_Y_HAKEDIS_GECMIS"
    ws.Unprotect ""
    Y_HAKEDIS_KAYDET ws
    secim = CStr(ws.Range("BX24").Value2)
    If CStr(ws.Range("BX22").Value2) <> il Then secim = vbNullString
    Y_SAYFA_SIFIRLA ws
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    gen = Array(70, w - 1134, 70, 130, 110, 110, 110, 110, 150, 160, 114)
    Y_KOLONLAR ws, gen
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10
    ws.Rows(8).RowHeight = 28
    ws.Range(ws.Rows(9), ws.Rows(19)).RowHeight = 21
    ws.Rows(20).RowHeight = 12
    ws.Rows(21).RowHeight = 28
    ws.Rows(22).RowHeight = 32
    ws.Range(ws.Rows(YH_KALEM_ILK), ws.Rows(YH_KALEM_ILK + YH_KALEM_ADET)).RowHeight = 20
    son = YH_KALEM_ILK + YH_KALEM_ADET - 1
    ws.Rows(son + 2).RowHeight = 12
    ws.Rows(YH_GECMIS_ILK - 2).RowHeight = 28
    ws.Rows(YH_GECMIS_ILK - 1).RowHeight = 24
    ws.Range(ws.Rows(YH_GECMIS_ILK), ws.Rows(YH_GECMIS_ILK + 11)).RowHeight = 19
    ' Bolum basliklari
    ws.Range("A8").Value2 = "SÖZLEÞME VE DÖNEM"
    ws.Range("H8").Value2 = "HAKEDÝÞ ÖZETÝ VE KESÝNTÝLER"
    With ws.Range("A8,H8")
        .Font.Name = "Aptos": .Font.Size = 12: .Font.bold = True: .Font.Color = RGB(12, 45, 79): .VerticalAlignment = xlCenter
    End With
    ' Sol blok: etiket A:B, deger C:E
    etiketler = Array("Yüklenici", "Sözleþme No", "Sözleþme Bedeli", "Hakediþ No", "Hakediþ Dönemi", "Önceki Hakediþler (ödenecek)", "Sözleþme Kullanýmý")
    For i = 0 To 6
        r = 9 + i
        ws.Range(ws.Cells(r, 1), ws.Cells(r, 2)).Merge
        ws.Cells(r, 1).Value2 = etiketler(i)
        ws.Range(ws.Cells(r, 3), ws.Cells(r, 5)).Merge
    Next i
    With ws.Range("A9:E15")
        .Font.Name = "Aptos": .Font.Size = 10: .VerticalAlignment = xlCenter
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(214, 224, 236)
        End With
    End With
    ws.Range("A9:A15").Font.bold = True: ws.Range("A9:A15").Font.Color = RGB(28, 54, 83): ws.Range("A9:A15").IndentLevel = 1
    ws.Range("A9:B15").Interior.Color = RGB(242, 249, 255)
    Y_GIRIS_HUCRE ws.Range("C10:E13")
    ws.Range("C11").NumberFormat = "#,##0.00 ""TL"""
    ws.Range("C12").NumberFormat = "0"
    ws.Range("C13").NumberFormat = "@"
    ws.Range("C10:E15").HorizontalAlignment = xlLeft
    ws.Range("C10:E15").IndentLevel = 1
    ws.Range("C14").Formula = "=SUMIFS('_Y_HAKEDIS_GECMIS'!F:F,'_Y_HAKEDIS_GECMIS'!A:A,$BZ$10,'_Y_HAKEDIS_GECMIS'!B:B,""<""&N($C$12))"
    ws.Range("C14").NumberFormat = "#,##0.00 ""TL"""
    ws.Range("C15").Formula = "=IF(N(C11)>0,J" & son + 1 & "/C11,"""")"
    ws.Range("C15").NumberFormat = "0.0%"
    ws.Range("C15").Font.bold = True
    Set rr = ws.Range("C15")
    With rr.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=1")
        .Font.Color = RGB(185, 28, 28): .Interior.Color = RGB(254, 226, 226)
    End With
    ' Sag blok: etiket H:I, oran J, tutar K
    etiketler = Array("Bu dönem iþ tutarý", "Fiyat farký", "Toplam (KDV hariç)", "KDV", "Damga vergisi", "Teminat kesintisi", _
        "KDV tevkifatý", "Malzeme kesintisi", "Diðer kesintiler", "Toplam kesinti", "ÖDENECEK TUTAR")
    For i = 0 To 10
        r = 9 + i
        ws.Range(ws.Cells(r, 8), ws.Cells(r, 9)).Merge
        ws.Cells(r, 8).Value2 = etiketler(i)
    Next i
    With ws.Range("H9:K19")
        .Font.Name = "Aptos": .Font.Size = 10: .VerticalAlignment = xlCenter
        With .Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous: .Weight = xlHairline: .Color = RGB(214, 224, 236)
        End With
    End With
    ws.Range("H9:H19").IndentLevel = 1: ws.Range("H9:H19").Font.Color = RGB(28, 54, 83)
    ws.Range("H9:I19").Interior.Color = RGB(242, 249, 255)
    ws.Range("K9").Formula = "=I" & son + 1
    ws.Range("K11").Formula = "=K9+N(K10)"
    ws.Range("K12").Formula = "=K11*N(J12)"
    ws.Range("K13").Formula = "=K11*N(J13)"
    ws.Range("K14").Formula = "=K11*N(J14)"
    ws.Range("K15").Formula = "=K12*N(J15)"
    ws.Range("K18").Formula = "=SUM(K13:K15)+N(K16)+N(K17)"
    ws.Range("K19").Formula = "=K11+K12-K18"
    ws.Range("K9:K19").NumberFormat = "#,##0.00 ""TL"""
    ws.Range("K9:K19").HorizontalAlignment = xlRight
    ws.Range("J12:J15").NumberFormat = "0.00%"
    ws.Range("J12:J15").HorizontalAlignment = xlCenter
    Y_GIRIS_HUCRE ws.Range("K10")
    Y_GIRIS_HUCRE ws.Range("J12:J15")
    Y_GIRIS_HUCRE ws.Range("K16:K17")
    ws.Range("J9").Value2 = "Oran": ws.Range("J9").Font.Size = 8: ws.Range("J9").Font.Color = RGB(100, 116, 139): ws.Range("J9").HorizontalAlignment = xlCenter
    With ws.Range("H11:K11,H18:K18")
        .Font.bold = True
    End With
    With ws.Range("H19:K19")
        .Font.bold = True: .Font.Size = 12: .Font.Color = vbWhite: .Interior.Color = RGB(0, 125, 110)
    End With
    ' Kalemler
    ws.Range("A21").Value2 = "HAKEDÝÞ KALEMLERÝ  –  sözleþme birim fiyatý × metraj"
    With ws.Range("A21:F21")
        .Merge: .Font.Name = "Aptos": .Font.Size = 12: .Font.bold = True: .Font.Color = RGB(12, 45, 79): .VerticalAlignment = xlCenter
    End With
    ws.Range("G21").Value2 = "Sarý hücreler giriþ alanýdýr. Kümülatif miktar sözleþmeyi aþarsa satýr kýrmýzý uyarý verir."
    With ws.Range("G21:K21")
        .Merge: .Font.Name = "Aptos": .Font.Size = 9: .Font.Italic = True: .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = xlCenter: .HorizontalAlignment = xlRight
    End With
    Y_BASLIK_SATIRI ws.Range("A22:K22"), Array("Poz No", "Ýþ Kalemi Tanýmý", "Birim", "Birim Fiyat", "Sözleþme Miktarý", _
        "Önceki Kümülatif", "Bu Dönem", "Kümülatif Miktar", "Bu Dönem Tutar", "Kümülatif Tutar", "Gerçekleþme")
    For r = YH_KALEM_ILK To son
        ws.Cells(r, 8).Formula = "=IF(COUNT(F" & r & ":G" & r & ")=0,"""",N(F" & r & ")+N(G" & r & "))"
        ws.Cells(r, 9).Formula = "=IF(N(G" & r & ")=0,"""",N(G" & r & ")*N(D" & r & "))"
        ws.Cells(r, 10).Formula = "=IF(H" & r & "="""","""",N(H" & r & ")*N(D" & r & "))"
        ws.Cells(r, 11).Formula = "=IF(N(E" & r & ")>0,IF(H" & r & "="""","""",N(H" & r & ")/E" & r & "),"""")"
    Next r
    Set rr = ws.Range(ws.Cells(YH_KALEM_ILK, 1), ws.Cells(son, 11))
    Y_TABLO_GOVDE rr
    Y_GIRIS_HUCRE ws.Range(ws.Cells(YH_KALEM_ILK, 1), ws.Cells(son, 7))
    ws.Range(ws.Cells(YH_KALEM_ILK, 1), ws.Cells(son, 1)).NumberFormat = "@"
    ws.Range(ws.Cells(YH_KALEM_ILK, 1), ws.Cells(son, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(YH_KALEM_ILK, 3), ws.Cells(son, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(YH_KALEM_ILK, 4), ws.Cells(son + 1, 4)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(YH_KALEM_ILK, 5), ws.Cells(son + 1, 8)).NumberFormat = "#,##0.###"
    ws.Range(ws.Cells(YH_KALEM_ILK, 9), ws.Cells(son + 1, 10)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(YH_KALEM_ILK, 11), ws.Cells(son + 1, 11)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(YH_KALEM_ILK, 4), ws.Cells(son + 1, 10)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(YH_KALEM_ILK, 11), ws.Cells(son + 1, 11)).HorizontalAlignment = xlCenter
    ' Kumulatif kontrol: sozlesme miktari asilirsa satir uyarisi
    Set rr = ws.Range(ws.Cells(YH_KALEM_ILK, 8), ws.Cells(son, 11))
    With rr.FormatConditions.Add(Type:=xlExpression, Formula1:="=AND(N($E" & YH_KALEM_ILK & ")>0,N($H" & YH_KALEM_ILK & ")>$E" & YH_KALEM_ILK & ")")
        .Interior.Color = RGB(254, 226, 226): .Font.Color = RGB(185, 28, 28): .Font.bold = True
    End With
    ws.Cells(son + 1, 1).Value2 = "TOPLAM"
    ws.Cells(son + 1, 9).Formula = "=SUM(I" & YH_KALEM_ILK & ":I" & son & ")"
    ws.Cells(son + 1, 10).Formula = "=SUM(J" & YH_KALEM_ILK & ":J" & son & ")"
    ws.Cells(son + 1, 11).Formula = "=IF(N(C11)>0,J" & son + 1 & "/C11,"""")"
    Y_TOPLAM_SATIRI ws.Range(ws.Cells(son + 1, 1), ws.Cells(son + 1, 11))
    ' Gecmis
    ws.Cells(YH_GECMIS_ILK - 2, 1).Value2 = "KAYITLI HAKEDÝÞLER"
    With ws.Cells(YH_GECMIS_ILK - 2, 1)
        .Font.Name = "Aptos": .Font.Size = 12: .Font.bold = True: .Font.Color = RGB(12, 45, 79)
    End With
    Y_BASLIK_SATIRI ws.Range(ws.Cells(YH_GECMIS_ILK - 1, 1), ws.Cells(YH_GECMIS_ILK - 1, 11)), _
        Array("No", "Dönem", "", "Bu Dönem Ýþ Tutarý", "", "", "", "", "Toplam Kesinti", "Ödenecek", "Kayýt Tarihi")
    Set rr = ws.Range(ws.Cells(YH_GECMIS_ILK, 1), ws.Cells(YH_GECMIS_ILK + 11, 11))
    Y_TABLO_GOVDE rr
    ws.Range(ws.Cells(YH_GECMIS_ILK, 1), ws.Cells(YH_GECMIS_ILK + 11, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(YH_GECMIS_ILK, 4), ws.Cells(YH_GECMIS_ILK + 11, 10)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(YH_GECMIS_ILK, 11), ws.Cells(YH_GECMIS_ILK + 11, 11)).NumberFormat = "dd.mm.yyyy hh:nn"
    ws.Range(ws.Cells(YH_GECMIS_ILK, 11), ws.Cells(YH_GECMIS_ILK + 11, 11)).HorizontalAlignment = xlCenter
    ' KPI kartlari
    ws.Range("BZ1").Formula = "=IF(N(C11)>0,C11,""Sözleþme bedeli girilmedi"")"
    ws.Range("BZ2").Formula = "=J" & son + 1
    ws.Range("BZ3").Formula = "=K19"
    ws.Range("BZ4").Formula = "=""No ""&N(C12)&IF(LEN(C13)>0,""   •   ""&C13,"""")"
    ws.Range("BZ1:BZ3").NumberFormat = "#,##0.00 ""TL"""
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "SÖZLEÞME BEDELÝ", "$BZ$1", RGB(35, 119, 204)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "KÜMÜLATÝF ÝÞ TUTARI", "$BZ$2", RGB(39, 160, 94)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "BU DÖNEM ÖDENECEK", "$BZ$3", RGB(0, 125, 110)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "HAKEDÝÞ", "$BZ$4", RGB(35, 119, 204)
    ' Yuklenici secimi
    a = Y_HK_LISTE(il)
    Set liste = ws.Shapes.AddFormControl(xlDropDown, ws.Range("C9").Left + 2, ws.Range("C9").Top + 2, ws.Range("C9:E9").Width - 4, ws.Rows(9).Height - 4)
    liste.Name = "Y_HK_YUKLENICI": liste.Placement = xlFreeFloating
    p = 1
    If IsArray(a) Then
        For i = 0 To UBound(a)
            liste.ControlFormat.AddItem CStr(a(i))
            If Len(secim) > 0 And Left$(CStr(a(i)), Len(secim)) = secim Then p = i + 1
        Next i
        liste.ControlFormat.Value = p
        secim = Trim$(Split(CStr(a(p - 1)) & "  -  ", "  -  ")(0))
    Else
        liste.ControlFormat.AddItem "Yüklenici deposu yok"
        liste.ControlFormat.Value = 1
        secim = vbNullString
    End If
    liste.OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_HAKEDIS_SEC")
    ' Dugmeler (sol blok altinda)
    Y_DUGME ws, "Y_HK_KAYDET", "HAKEDÝÞÝ KAYDET", ws.Range("A17").Left, ws.Range("A17").Top, 170, 32, RGB(0, 125, 110), "AKSORA_YATIRIM_HAKEDIS_KAYDET"
    Y_DUGME ws, "Y_HK_YENI", "YENÝ DÖNEME GEÇ", ws.Range("A17").Left + 180, ws.Range("A17").Top, 170, 32, RGB(16, 57, 94), "AKSORA_YATIRIM_HAKEDIS_YENI"
    Y_DUGME ws, "Y_HK_PDF", "PDF AL", ws.Range("A17").Left + 360, ws.Range("A17").Top, 120, 32, RGB(35, 119, 204), "AKSORA_YATIRIM_HAKEDIS_PDF"
    ws.Range("BX22").Value2 = il
    ws.Range("BX24").Value2 = secim
    ws.Range("BZ10").Value2 = Y_HK_ANAHTAR(ws)
    ws.Range("BX20").Value2 = il
    Y_HAKEDIS_YUKLE ws
    Y_HAKEDIS_GECMIS_YAZ ws
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | OTOMATÝK HAKEDÝÞ  |  " & Y_IL_ADI(il)
    ' Yazici yoksa sayfa ayari atlanir.
    On Error Resume Next
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintArea = "$A$1:$K$" & son + 1
    On Error GoTo 0
    ws.Range("BX25").Value2 = "HK_V1"
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlUnlockedCells
End Sub

Public Sub AKSORA_YATIRIM_HAKEDIS_SEC()
    Dim ws As Worksheet, s As Shape, t As String, ev As Boolean, sc As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    Set s = ws.Shapes("Y_HK_YUKLENICI")
    t = CStr(s.ControlFormat.List(s.ControlFormat.Value))
    t = Trim$(Split(t & "  -  ", "  -  ")(0))
    ev = application.EnableEvents: sc = application.ScreenUpdating
    application.EnableEvents = False: application.ScreenUpdating = False
    ws.Unprotect ""
    Y_HAKEDIS_KAYDET ws
    ws.Range("BX24").Value2 = t
    ws.Range("BZ10").Value2 = Y_HK_ANAHTAR(ws)
    Y_HAKEDIS_YUKLE ws
    Y_HAKEDIS_GECMIS_YAZ ws
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev: application.ScreenUpdating = sc
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

' Donem kaydi: ayni hakedis no varsa uzerine yazilir.
Public Sub AKSORA_YATIRIM_HAKEDIS_KAYDET()
    Dim ws As Worksheet, g As Worksheet, son As Long, r As Long, sr As Long, anahtar As String, no As Long, ev As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    If Len(CStr(ws.Range("BX24").Value2)) = 0 Then MsgBox "Önce yüklenici seçiniz.", vbInformation, "AKSORA": Exit Sub
    no = CLng(Val(CStr(ws.Range("C12").Value2)))
    If no < 1 Then MsgBox "Hakediþ numarasý giriniz.", vbInformation, "AKSORA": Exit Sub
    If Val(CStr(ws.Range("K9").Value2)) = 0 And Val(CStr(ws.Range("K10").Value2)) = 0 Then
        If MsgBox("Bu dönem iþ tutarý sýfýr. Yine de kaydedilsin mi?", vbYesNo + vbQuestion, "AKSORA") <> vbYes Then Exit Sub
    End If
    ev = application.EnableEvents: application.EnableEvents = False
    ws.Unprotect ""
    anahtar = Y_HK_ANAHTAR(ws)
    Set g = Y_KAYIT_SAYFA("_Y_HAKEDIS_GECMIS")
    son = g.Cells(g.Rows.Count, 1).End(xlUp).Row
    If Len(CStr(g.Cells(1, 1).Value2)) = 0 Then son = 0
    For r = 1 To son
        If CStr(g.Cells(r, 1).Value2) = anahtar And CLng(Val(CStr(g.Cells(r, 2).Value2))) = no Then sr = r: Exit For
    Next r
    If sr = 0 Then sr = son + 1
    g.Cells(sr, 1).NumberFormat = "@"
    g.Cells(sr, 1).Value2 = anahtar
    g.Cells(sr, 2).Value2 = no
    g.Cells(sr, 3).NumberFormat = "@"
    g.Cells(sr, 3).Value2 = CStr(ws.Range("C13").Value2)
    g.Cells(sr, 4).Value2 = ws.Range("K11").Value2
    g.Cells(sr, 5).Value2 = ws.Range("K18").Value2
    g.Cells(sr, 6).Value2 = ws.Range("K19").Value2
    g.Cells(sr, 7).Value = Now
    Y_HAKEDIS_KAYDET ws
    Y_HAKEDIS_GECMIS_YAZ ws
    ws.Calculate
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev
    MsgBox "Hakediþ No " & no & " kaydedildi." & vbCrLf & "Ödenecek tutar: " & Format$(ws.Range("K19").Value2, "#,##0.00") & " TL", vbInformation, "AKSORA | Hakediþ"
    Exit Sub
Hata:
    application.EnableEvents = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

' Yeni donem: kumulatif miktarlar "onceki"ye aktarilir, bu donem bosaltilir, hakedis no artar.
Public Sub AKSORA_YATIRIM_HAKEDIS_YENI()
    Dim ws As Worksheet, g As Worksheet, r As Long, son As Long, kayitli As Boolean, anahtar As String, no As Long, ev As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    If Len(CStr(ws.Range("BX24").Value2)) = 0 Then Exit Sub
    no = CLng(Val(CStr(ws.Range("C12").Value2)))
    anahtar = Y_HK_ANAHTAR(ws)
    Set g = Y_KAYIT_SAYFA("_Y_HAKEDIS_GECMIS")
    son = g.Cells(g.Rows.Count, 1).End(xlUp).Row
    For r = 1 To son
        If CStr(g.Cells(r, 1).Value2) = anahtar And CLng(Val(CStr(g.Cells(r, 2).Value2))) = no Then kayitli = True: Exit For
    Next r
    If Not kayitli Then
        Select Case MsgBox("Hakediþ No " & no & " henüz kaydedilmedi. Önce kaydedilsin mi?", vbYesNoCancel + vbQuestion, "AKSORA | Hakediþ")
            Case vbCancel: Exit Sub
            Case vbYes: AKSORA_YATIRIM_HAKEDIS_KAYDET
        End Select
    End If
    ev = application.EnableEvents: application.EnableEvents = False
    ws.Unprotect ""
    For r = YH_KALEM_ILK To YH_KALEM_ILK + YH_KALEM_ADET - 1
        If Len(CStr(ws.Cells(r, 8).Value2)) > 0 And Not IsError(ws.Cells(r, 8).Value2) Then ws.Cells(r, 6).Value2 = ws.Cells(r, 8).Value2
        ws.Cells(r, 7).ClearContents
    Next r
    ws.Range("C12").Value2 = no + 1
    ws.Range("C13").ClearContents
    ws.Range("K10").ClearContents: ws.Range("K16:K17").ClearContents
    Y_HAKEDIS_KAYDET ws
    ws.Calculate
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev
    MsgBox "Hakediþ No " & no + 1 & " hazýr. Kümülatif miktarlar önceki hakediþe aktarýldý; bu dönem miktarlarýný girin.", vbInformation, "AKSORA | Hakediþ"
    Exit Sub
Hata:
    application.EnableEvents = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_YATIRIM_HAKEDIS_PDF()
    Dim ws As Worksheet, yol As String
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    yol = ThisWorkbook.Path & application.PathSeparator & "HAKEDIS_" & CStr(ws.Range("BX22").Value2) & "_" & _
        CStr(ws.Range("BX24").Value2) & "_No" & CLng(Val(CStr(ws.Range("C12").Value2))) & "_" & Format$(Now, "yyyymmdd_hhnn") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=yol, Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    Exit Sub
Hata:
    MsgBox "PDF oluþturulamadý: " & Err.Description, vbExclamation, "AKSORA"
End Sub

' Calisma kitabi SheetChange: teminat ve hakedis girisleri aninda kayda yazilir.
Public Sub AKSORA_YATIRIM_DEGISIM(ByVal sh As Object, ByVal hedef As Range)
    Dim ev As Boolean, c As Range, kayit As Worksheet, sr As Long, il As String, son As Long, alan As Range
    If Not TypeOf sh Is Worksheet Then Exit Sub
    If sh.Name <> "Y_TEMÝNAT" And sh.Name <> "Y_HAKEDÝÞ" Then Exit Sub
    If mYCalisiyor Then Exit Sub
    ev = application.EnableEvents
    On Error GoTo Cikis
    application.EnableEvents = False
    If sh.Name = "Y_TEMÝNAT" Then
        il = UCase$(Trim$(CStr(sh.Range("BX22").Value2)))
        son = Val(CStr(sh.Range("BX23").Value2))
        If son < YT_ILK Or Len(il) = 0 Then GoTo Cikis
        Set alan = Intersect(hedef, sh.Range(sh.Cells(YT_ILK, 6), sh.Cells(son, 6)))
        If alan Is Nothing Then GoTo Cikis
        Set kayit = Y_KAYIT_SAYFA("_Y_TEMINAT")
        For Each c In alan.Cells
            If Len(CStr(sh.Cells(c.Row, 1).Value2)) > 0 Then
                sr = Y_KAYIT_SATIR(kayit, il & "|" & UCase$(Trim$(CStr(sh.Cells(c.Row, 1).Value2))), True)
                kayit.Cells(sr, 2).Value2 = c.Value2
                kayit.Cells(sr, 3).Value = Now
            End If
        Next c
    Else
        Set alan = Intersect(hedef, sh.Range("C10:E13,K10,J12:J15,K16:K17," & _
            sh.Range(sh.Cells(YH_KALEM_ILK, 1), sh.Cells(YH_KALEM_ILK + YH_KALEM_ADET - 1, 7)).Address))
        If Not alan Is Nothing Then Y_HAKEDIS_KAYDET sh
    End If
Cikis:
    application.EnableEvents = ev
End Sub

' Giris ve guncelleme sonrasi: hata olsa bile birim acilir, hata durum seridine yazilir.
Private Sub Y_ALT_SAYFALAR(ByVal il As String)
    On Error Resume Next
    Y_TEMINAT_KUR il
    If Err.Number <> 0 Then Y_SERIT "Teminat sayfasý kurulamadý: " & Err.Description, 3: Err.Clear
    ' Hakedis sayfasi yalniz il veya tasarim surumu degisince yeniden kurulur (girisler korunur).
    With ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
        If CStr(.Range("BX22").Value2) <> UCase$(Trim$(il)) Or CStr(.Range("BX25").Value2) <> "HK_V1" Then
            Y_HAKEDIS_KUR il
        End If
    End With
    If Err.Number <> 0 Then Y_SERIT "Hakediþ sayfasý kurulamadý: " & Err.Description, 3: Err.Clear
    On Error GoTo 0
End Sub
