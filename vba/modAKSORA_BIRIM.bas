Attribute VB_Name = "modAKSORA_BIRIM"
Option Explicit

Private Const HUB As String = "BIRIM SECIM"
Private Const YAT As String = "YATIRIM PANEL"
Private Const MUS As String = "MUSTERI PANEL"
Private Const MUK As String = "MUSTERI IL SECIM"
Private Const YAK As String = "YATIRIM IL SECIM"
Private Const tag As String = "AKSORA_BIRIM_V1"
Private Const PW As String = ""
Private Const CW As Double = 1080
Private Const CH As Double = 660
Private mYonlendirme As Boolean
Private mGorunum As Boolean

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
    If ws.Shapes.Count = 0 Or Val(CStr(ws.Range("BX2").Value2)) < 7 Then BIRIM_CIZ ws, ""
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
    If AKSORA_MUSTERI_TEMA_SAYFA(sh) Then
        AKSORA_MUSTERI_TEMA_GORUNUM sh
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
        ws.Unprotect PW
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
    ws.Unprotect PW
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
    ws.Protect Password:=PW, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
    ws.Visible = xlSheetVisible: ws.Select Replace:=True
    BIRIM_DIGERLERINI_GIZLE MUK
    ThisWorkbook.Worksheets("KAPAK").Visible = xlSheetVeryHidden
    MUSTERI_KAPAK_GORUNUM ws
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=PW, UserInterfaceOnly:=True
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
    Dim ws As Worksheet, ix As Long, keys As Variant
    On Error GoTo Hata
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name <> MUK Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(MUK)
    ix = CLng(ws.Shapes(CStr(application.Caller)).AlternativeText)
    If ix < 1 Or ix > 5 Then Exit Sub
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    ws.Unprotect PW
    ws.Range("AZ1").Value2 = keys(ix - 1)
    MUSTERI_SECIM_CIZ ws
Cikis:
    If Not ws Is Nothing Then ws.Protect Password:=PW, UserInterfaceOnly:=True
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Sub MUSTERI_SECIM_CIZ(ByVal ws As Worksheet)
    Dim keys As Variant, names As Variant, i As Long, key As String
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    names = Array("GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN")
    key = CStr(ws.Range("AZ1").Value2)
    ws.Shapes("AKSORA_V4_STATUS").TextFrame2.TextRange.Text = "SEÇÝLÝ ÝL: -"
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
    Else
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
        ws.Unprotect PW
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
    ws.Unprotect PW
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
    ws.Protect Password:=PW, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
    ws.Visible = xlSheetVisible: ws.Select Replace:=True
    BIRIM_DIGERLERINI_GIZLE YAK
    ThisWorkbook.Worksheets("KAPAK").Visible = xlSheetVeryHidden
    YATIRIM_KAPAK_GORUNUM ws
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:=PW, UserInterfaceOnly:=True
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
    Dim ws As Worksheet, ix As Long, keys As Variant
    On Error GoTo Hata
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name <> YAK Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(YAK)
    ix = CLng(ws.Shapes(CStr(application.Caller)).AlternativeText)
    If ix < 1 Or ix > 5 Then Exit Sub
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    ws.Unprotect PW
    ws.Range("AZ1").Value2 = keys(ix - 1)
    YATIRIM_SECIM_CIZ ws
Cikis:
    If Not ws Is Nothing Then ws.Protect Password:=PW, UserInterfaceOnly:=True
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Sub YATIRIM_SECIM_CIZ(ByVal ws As Worksheet)
    Dim keys As Variant, names As Variant, i As Long, key As String
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    names = Array("GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN")
    key = CStr(ws.Range("AZ1").Value2)
    ws.Shapes("AKSORA_V4_STATUS").TextFrame2.TextRange.Text = "SEÇÝLÝ ÝL: -"
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
        MsgBox "Yatýrým Birimi iþlem ekranlarý hazýrlanýyor.", vbInformation, "AKSORA"
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
    If ws.Shapes.Count = 0 Or Val(CStr(ws.Range("BX2").Value2)) < 7 Then BIRIM_CIZ ws, baslik
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
    oran = genislik / CW
    ws.Unprotect Password:=PW
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
    ws.Protect Password:=PW, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
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
    ws.Unprotect PW
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 6) = "BIRIM_" Then ws.Shapes(i).Delete
    Next i
    ws.Range("A1:BH44").Interior.Color = RGB(236, 248, 254)
    ws.Range("A1:BH44").ClearContents
    ws.Columns("A:BH").ColumnWidth = 2.5
    Do While ws.Range("A1:BH1").Width < CW
        ws.Columns("A:BH").ColumnWidth = ws.Columns("A").ColumnWidth + 0.25
    Loop
    ws.Rows("1:44").RowHeight = 15
    ws.Range("BX1").Value2 = tag
    ws.Range("BX2").Value2 = 7
    ws.Columns("BX").Hidden = True
    ws.ScrollArea = "A1:BH44"
    ws.Tab.Color = RGB(12, 45, 79)
    BIRIM_RECT ws, "ZEMIN", 0, 0, CW, CH, RGB(236, 248, 254), -1, 1
    yol = ThisWorkbook.Path & application.PathSeparator & "AKSORA_PANEL_TEK_PARCA_V7.png"
    If Len(Dir$(yol)) > 0 Then
        Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 126, CW, 483)
        s.Name = "BIRIM_FOTOGRAF": s.Placement = xlFreeFloating: s.Locked = True
    End If
    Set s = BIRIM_RECT(ws, "FOTO_SIS", 0, 126, CW, 483, RGB(235, 248, 255), -1, 1)
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
        .Left = 0: .Top = 0: .Width = CW: .Height = 126
        .Placement = xlFreeFloating: .Locked = True
    End With
    BIRIM_RECT ws, "UST_CIZGI", 0, 124, CW, 3, RGB(34, 185, 191), -1, 1
    BIRIM_RECT ws, "ALT", 0, 609, CW, 51, RGB(12, 45, 79), -1, 1
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
            "Ýþlemlere geç", "AKSORA_BIRIM_MUSTERI", RGB(32, 133, 158), False
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
    ws.Protect Password:=PW, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
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
