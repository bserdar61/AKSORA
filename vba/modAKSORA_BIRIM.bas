Attribute VB_Name = "modAKSORA_BIRIM"
Option Explicit

Private Const HUB As String = "BIRIM SECIM"
Private Const YAT As String = "YATIRIM PANEL"
Private Const MUS As String = "MUSTERI PANEL"
Private Const MUK As String = "MUSTERI IL SECIM"
Private Const YAK As String = "YATIRIM IL SECIM"
Private Const ILCE_KAPAK As String = "ILCE IL SECIM"
Private Const tag As String = "AKSORA_BIRIM_V1"
Private Const pw As String = ""
Private Const cw As Double = 1080
Private Const CH As Double = 660
Private Const ADM As String = "ADMIN PANEL"
Private Const ADM_SIFRE As String = "adminbs1284"
Private Const ADM_ARA_SATIR As Long = 8
Private Const ADM_ARA_HUCRE As String = "C7"
Private Const ADM_IL_HUCRE As String = "C4"
Private mAdmOnbellek As Object
Private mAdmOturum As Boolean
Private Const ADM_KUTU_BASLIK As String = "AKSORA - YONETICI GIRISI"
Private Type ADMW_RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare PtrSafe Function ADMW_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function ADMW_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long
Private Declare PtrSafe Function ADMW_FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
Private Declare PtrSafe Function ADMW_FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As LongPtr, ByVal hWnd2 As LongPtr, ByVal lpsz1 As String, ByVal lpsz2 As String) As LongPtr
Private Declare PtrSafe Function ADMW_SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As LongPtr, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
Private Declare PtrSafe Function ADMW_CreateWindowEx Lib "user32" Alias "CreateWindowExW" (ByVal dwExStyle As Long, ByVal lpClassName As LongPtr, ByVal lpWindowName As LongPtr, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As LongPtr, ByVal hMenu As LongPtr, ByVal hInstance As LongPtr, ByVal lpParam As LongPtr) As LongPtr
Private Declare PtrSafe Function ADMW_GetWindowRect Lib "user32" Alias "GetWindowRect" (ByVal hWnd As LongPtr, lpRect As ADMW_RECT) As Long
Private Declare PtrSafe Function ADMW_MapWindowPoints Lib "user32" Alias "MapWindowPoints" (ByVal hWndFrom As LongPtr, ByVal hWndTo As LongPtr, lpPoints As ADMW_RECT, ByVal cPoints As Long) As Long
Private Declare PtrSafe Function ADMW_SetWindowPos Lib "user32" Alias "SetWindowPos" (ByVal hWnd As LongPtr, ByVal hWndInsertAfter As LongPtr, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Declare PtrSafe Function ADMW_InvalidateRect Lib "user32" Alias "InvalidateRect" (ByVal hWnd As LongPtr, ByVal lpRect As LongPtr, ByVal bErase As Long) As Long
Private Declare PtrSafe Function ADMW_SetFocus Lib "user32" Alias "SetFocus" (ByVal hWnd As LongPtr) As LongPtr
#If Win64 Then
Private Declare PtrSafe Function ADMW_GetWindowLong Lib "user32" Alias "GetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function ADMW_SetWindowLong Lib "user32" Alias "SetWindowLongPtrA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#Else
Private Declare PtrSafe Function ADMW_GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
Private Declare PtrSafe Function ADMW_SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
#End If
Private Declare PtrSafe Function ADMW_LockWindowUpdate Lib "user32" Alias "LockWindowUpdate" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ADMW_RedrawWindow Lib "user32" Alias "RedrawWindow" (ByVal hWnd As LongPtr, ByVal lprcUpdate As LongPtr, ByVal hrgnUpdate As LongPtr, ByVal fuRedraw As Long) As Long
Private mAdmZaman As LongPtr
Private mAdmKutu As LongPtr
Private mAdmGoster As Boolean
Private mAdmYukleme As Boolean
Private mAdmMesgul As Boolean
Private mAdmYukWs As Worksheet
Private mAdmYukBar As Double
Private mAdmYukSon As Double
Public AKSORA_BEKLE_KIPI As Boolean
Public AKSORA_ADMIN_IPTAL As Boolean
Private mAdmTopluKart As Boolean
Private mAdmTopluBas As Double
Private mAdmTopluSon As Double
Private mAdmTopluSira As Long
Private mAdmTopluAd As String
Private mAdmTopluMetin As String
Private mAdmTopluIl As Object
Private mAdmKilit As Boolean
Private Const ADM_ARA_N As Long = 400
Public AKSORA_ADMIN_TOPLU As Boolean
Public AKSORA_ADMIN_SONUC As String
Private mYonlendirme As Boolean
Private mGorunum As Boolean
Private Const Y_ETIKET As String = "AKSORA_YATIRIM_V2"
Private Const I_ETIKET As String = "AKSORA_ILCE_V1"
Private Const I_GECMIS_GUN As Long = 20
Private Const I_TF As String = "dd.mm.yyyy hh:nn:ss"
Private Const I_DEFTER_SURE As Long = 400
Private mIMal As Object
Private mIMalKirli As Boolean
Private mIDefter As Object
Private mYCalisiyor As Boolean
Private mYToplu As Boolean
Private mYOdak As String
Private mYSira As String
Private Const YT_ILK As Long = 10
Private Const YH_KALEM_ILK As Long = 23
Private Const YH_KALEM_ADET As Long = 30
Private Const YH_GECMIS_ILK As Long = 58
' ATIL STOK: kayit, sayfa ve SAP okuma (pano) tanimlari
Private Const ATIL_SAYFA As String = "ATIL_STOK"
Private Const ATIL_KAYIT As String = "_ATIL"
Private Const ATIL_ETIKET As String = "AKSORA_ATIL_V1"
Public AKSORA_ATIL_CALISIYOR As Boolean
Private mAtilKart As Boolean
Private mAtilBas As Double
Private mAtilSon As Double
Private mAtilMetin As String
Private mAtilYuk As Object
Private mAtilEski As Boolean
' Uzun guncellemede kartin "ARKA PLANA AL" dugmesi: Excel diger pencerelerin arkasina gonderilir
' (kucultulmez: pencere durumu degisince sayfa denetimleri hata veriyordu). Is bitince gorev cubugu yanip soner.
Public AKSORA_ARKA_PLAN As Boolean
Private Type ADMW_FLASH
    cbSize As Long
    hWnd As LongPtr
    dwFlags As Long
    uCount As Long
    dwTimeout As Long
End Type
Private Declare PtrSafe Function ADMW_FlashWindowEx Lib "user32" Alias "FlashWindowEx" (pfwi As ADMW_FLASH) As Long
Private Declare PtrSafe Function ATIL_OpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ATIL_EmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
Private Declare PtrSafe Function ATIL_CloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare PtrSafe Function ATIL_SetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function ATIL_GlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function ATIL_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function ATIL_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function ATIL_GlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub ATIL_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal hedef As LongPtr, ByVal kaynak As LongPtr, ByVal adet As LongPtr)
Private Declare PtrSafe Sub ATIL_Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)

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
    yedek = AKSORA_KAYIT_KLASORU() & application.PathSeparator & _
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
    If ws.Shapes.Count = 0 Or Val(CStr(ws.Range("BX2").Value2)) < 14 Then BIRIM_CIZ ws, ""
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
    If sh.Name = ADM Then
        ' Sifre girilmeden (sekme / Ctrl+PgDn ile) yonetim merkezine gecilemez.
        If Not mAdmOturum Then
            AKSORA_BIRIM_OLAY = True
            AKSORA_BIRIM_ANA_MENU
            Exit Function
        End If
        ADM_GORUNUM sh: AKSORA_BIRIM_OLAY = True: Exit Function
    End If
    On Error Resume Next
    application.MoveAfterReturn = (sh.Name <> "M_MALZEME_ARA" And sh.Name <> "Y_MALZEME_ARA" And sh.Name <> "I_MALZEME_ARA")
    On Error GoTo 0
    If AKSORA_MUSTERI_TEMA_SAYFA(sh) Then
        AKSORA_CIKTI_DUGME sh
        AKSORA_MUSTERI_TEMA_GORUNUM sh
        AKSORA_MUSTERI_SECIM_BASA sh
        AKSORA_BIRIM_OLAY = True
        Exit Function
    End If
    If AKSORA_YATIRIM_SAYFA(sh) Or AKSORA_ILCE_SAYFA(sh) Then
        AKSORA_CIKTI_DUGME sh
        AKSORA_YATIRIM_GORUNUM sh
        AKSORA_BIRIM_OLAY = True
        Exit Function
    End If
    If AKSORA_ATIL_SAYFA(sh) Then
        AKSORA_YATIRIM_GORUNUM sh
        AKSORA_BIRIM_OLAY = True
        Exit Function
    End If
    If Not BIRIM_YENI_SAYFA(CStr(sh.Name)) Then Exit Function
    If sh.Name = MUK Or sh.Name = YAK Or sh.Name = ILCE_KAPAK Then MUSTERI_KAPAK_GORUNUM sh Else BIRIM_GORUNUM sh
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
    If AKSORA_YATIRIM_SAYFA(ActiveSheet) Or AKSORA_ILCE_SAYFA(ActiveSheet) Or AKSORA_ATIL_SAYFA(ActiveSheet) Then
        AKSORA_YATIRIM_GORUNUM ActiveSheet
        AKSORA_BIRIM_AKTIF_GORUNUM = True
        Exit Function
    End If
    If Not BIRIM_YENI_SAYFA(CStr(ActiveSheet.Name)) Then Exit Function
    If ActiveSheet.Name = MUK Or ActiveSheet.Name = YAK Or ActiveSheet.Name = ILCE_KAPAK Then MUSTERI_KAPAK_GORUNUM ActiveSheet Else BIRIM_GORUNUM ActiveSheet
    AKSORA_BIRIM_AKTIF_GORUNUM = True
End Function

Private Function BIRIM_YENI_SAYFA(ByVal ad As String) As Boolean
    BIRIM_YENI_SAYFA = (ad = HUB Or ad = YAT Or ad = MUS Or ad = MUK Or ad = YAK Or ad = ILCE_KAPAK)
End Function

Public Sub AKSORA_BIRIM_MALZEME()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  MALZEME STOK BÝRÝMÝ", "Birim açýlýyor...")
    On Error GoTo Son
    BIRIM_MALZEME_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub BIRIM_MALZEME_ASIL()
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
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  YATIRIM BÝRÝMÝ", "Birim açýlýyor...")
    On Error GoTo Son
    BIRIM_YATIRIM_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub BIRIM_YATIRIM_ASIL()
    On Error GoTo Hata
    YATIRIM_KAPAK_AC
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Yatýrým birimi", Err.Number, Err.Description
End Sub
Public Sub AKSORA_BIRIM_MUSTERI()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  MÜÞTERÝ BÝRÝMÝ", "Birim açýlýyor...")
    On Error GoTo Son
    BIRIM_MUSTERI_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub BIRIM_MUSTERI_ASIL()
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
    yol = AKSORA_GORSEL("AKSORA_MUSTERI_BANT.png")
    If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1610, , _
        "AKSORA_MUSTERI_BANT.png AKSORA klasöründe (GORSELLER) bulunmalýdýr."
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
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  MÜÞTERÝ BÝRÝMÝ", "Ýl stoklarý ve ekranlar hazýrlanýyor...")
    On Error GoTo Son
    MUSTERI_GIRIS_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub MUSTERI_GIRIS_ASIL()
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

' Il secim ekrani: yatirim (YAK) ve ilce yonetim (ILCE_KAPAK) ayni KAPAK kopyasini kullanir.
Private Sub YATIRIM_KAPAK_AC(Optional ByVal ilce As Boolean)
    Dim ws As Worksheet, kaynak As Worksheet, s As Shape, foto As Shape, bant As Shape
    Dim yol As String, ev As Boolean, sc As Boolean, no As Long, metin As String
    Dim i As Long, h As Double, kaynakGorunum As XlSheetVisibility, ad As String, bantVar As Boolean
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    ad = IIf(ilce, ILCE_KAPAK, YAK)
    If ilce Then
        yol = AKSORA_GORSEL("AKSORA_ILCE_BANT.png")
        bantVar = (Len(Dir$(yol)) > 0)
    Else
        yol = AKSORA_GORSEL("AKSORA_YATIRIM_BANT.png")
        If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1610, , _
            "AKSORA_YATIRIM_BANT.png AKSORA klasöründe (GORSELLER) bulunmalýdýr."
        bantVar = True
    End If
    BIRIM_SAYFA_KONTROL ad
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo Hata
    If ws Is Nothing Then
        Set kaynak = ThisWorkbook.Worksheets("KAPAK")
        kaynakGorunum = kaynak.Visible
        kaynak.Visible = xlSheetVisible
        kaynak.Copy after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count)
        Set ws = ActiveSheet
        ws.Name = ad
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
        ws.Shapes("AKSORA_V4_ENTER").OnAction = BIRIM_MAKRO(IIf(ilce, "AKSORA_ILCE_GIRIS", "AKSORA_YATIRIM_GIRIS"))
        AKSORA_BIRIM_KAPAK_GERI ad
        ws.Range("BX2").Value2 = 0
    End If
    ws.Unprotect pw
    If Val(CStr(ws.Range("BX2").Value2)) < 7 And bantVar Then
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
    YATIRIM_SECIM_CIZ ws, ilce
    ws.Protect Password:=pw, DrawingObjects:=True, Contents:=True, Scenarios:=True, UserInterfaceOnly:=True
    ws.EnableSelection = xlNoSelection
    ws.Visible = xlSheetVisible: ws.Select Replace:=True
    BIRIM_DIGERLERINI_GIZLE ad
    ThisWorkbook.Worksheets("KAPAK").Visible = xlSheetVeryHidden
    ' Il secim ekrani tek gorunur sayfa: arkada acik kalan sayfalara kayma olmaz.
    Dim yk As Worksheet
    For Each yk In ThisWorkbook.Worksheets
        If yk.Name <> ad And yk.Visible = xlSheetVisible Then yk.Visible = xlSheetVeryHidden
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
    If ActiveSheet.Name <> YAK And ActiveSheet.Name <> ILCE_KAPAK Then Exit Sub
    ev = application.EnableEvents
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(ActiveSheet.Name)
    ix = CLng(ws.Shapes(CStr(application.Caller)).AlternativeText)
    If ix < 1 Or ix > 5 Then Exit Sub
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    ' Secim tek seferde cizilir; olaylar baska sayfaya gecis yaptirmaz.
    application.ScreenUpdating = False: application.EnableEvents = False
    ws.Unprotect pw
    ws.Range("AZ1").Value2 = keys(ix - 1)
    YATIRIM_SECIM_CIZ ws, (ws.Name = ILCE_KAPAK)
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

Private Sub YATIRIM_SECIM_CIZ(ByVal ws As Worksheet, Optional ByVal ilce As Boolean)
    Dim keys As Variant, names As Variant, i As Long, key As String, son As Date
    keys = Array("GIRESUN", "TRABZON", "RIZE", "GUMUSHANE", "ARTVIN")
    names = Array("GÝRESUN", "TRABZON", "RÝZE", "GÜMÜÞHANE", "ARTVÝN")
    key = CStr(ws.Range("AZ1").Value2)
    ws.Shapes("AKSORA_V4_STATUS").TextFrame2.TextRange.Text = "SEÇÝLÝ ÝL: -"
    On Error Resume Next
    If Len(key) > 0 Then
        If ilce Then son = AKSORA_ILCE_IL_SON_TARIH(key) Else son = AKSORA_YATIRIM_IL_SON_TARIH(key)
        If son > 0 Then
            ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = "Son güncelleme: " & Format$(son, "dd.mm.yyyy hh:nn")
        Else
            ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = "Son güncelleme: henüz yapýlmadý"
        End If
    Else
        ws.Shapes("AKSORA_V8_TARIH").TextFrame2.TextRange.Text = IIf(ilce, "Ýlçe Yönetim birimi", "Yatýrým birimi")
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
    Dim ac As Boolean, no As Long, metin As String
    If Len(CStr(ThisWorkbook.Worksheets(YAK).Range("AZ1").Value2)) = 0 Then MsgBox "Lütfen bir il seçiniz.", vbInformation, "AKSORA": Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  YATIRIM BÝRÝMÝ", "Ýl stoklarý ve ekranlar hazýrlanýyor...")
    On Error GoTo Son
    YATIRIM_GIRIS_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub YATIRIM_GIRIS_ASIL()
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
    If ws.Shapes.Count = 0 Or Val(CStr(ws.Range("BX2").Value2)) < 14 Then BIRIM_CIZ ws, baslik
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
    For Each ad In Array(HUB, YAT, MUS, MUK, YAK, ILCE_KAPAK, ADM)
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
    Dim i As Long, s As Shape, yol As String, resimYukseklik As Double, kw As Double
    yol = AKSORA_GORSEL("AKSORA_BIRIM_BANNER_03.png")
    If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1608, , _
        "AKSORA_BIRIM_BANNER_03.png AKSORA klasöründe (GORSELLER) bulunmalýdýr."
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
    ws.Range("BX2").Value2 = 14
    ws.Columns("BX").Hidden = True
    ws.ScrollArea = "A1:BH44"
    ws.Tab.Color = RGB(12, 45, 79)
    BIRIM_RECT ws, "ZEMIN", 0, 0, cw, CH, RGB(236, 248, 254), -1, 1
    yol = AKSORA_GORSEL("AKSORA_PANEL_TEK_PARCA_V7.png")
    If Len(Dir$(yol)) > 0 Then
        Set s = ws.Shapes.AddPicture(yol, msoFalse, msoTrue, 0, 126, cw, 483)
        s.Name = "BIRIM_FOTOGRAF": s.Placement = xlFreeFloating: s.Locked = True
    End If
    Set s = BIRIM_RECT(ws, "FOTO_SIS", 0, 126, cw, 483, RGB(235, 248, 255), -1, 1)
    s.Fill.Transparency = 0.45
    yol = AKSORA_GORSEL("AKSORA_BIRIM_BANNER_03.png")
    If Len(Dir$(yol)) = 0 Then Err.Raise vbObjectError + 1608, , _
        "AKSORA_BIRIM_BANNER_03.png AKSORA klasöründe (GORSELLER) bulunmalýdýr."
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
        kw = (cw - 68 - 3 * 18) / 4
        BIRIM_KART ws, "STOK", 34, "01", "Malzeme Stok Birimi", _
            "", "Stok takibi" & vbLf & "Ýller arasý malzeme arama" & vbLf & _
            "Gün bazlý stok karþýlaþtýrma" & vbLf & "Ýl depolarý raporlarý", _
            "Ýþlemlere geç", "AKSORA_BIRIM_MALZEME", RGB(0, 125, 110), True, kw
        BIRIM_KART ws, "YAT", 34 + (kw + 18), "02", "Yatýrým Birimi", _
            "", "Hakediþ" & vbLf & "Teminat takibi" & vbLf & _
            "Yüklenici stok takibi" & vbLf & "Yatýrým depo takibi", _
            "Ýþlemlere geç", "AKSORA_BIRIM_YATIRIM", RGB(20, 103, 173), True, kw
        BIRIM_KART ws, "MUS", 34 + 2 * (kw + 18), "03", "Müþteri Birimi", _
            "", "Seri numaralý malzeme takibi" & vbLf & "Ýlçe stok kontrolleri" & vbLf & _
            "Yüklenici stok kontrolleri" & vbLf & "Müþteri depo takibi", _
            "Ýþlemlere geç", "AKSORA_BIRIM_MUSTERI", RGB(32, 133, 158), True, kw
        BIRIM_KART ws, "ILC", 34 + 3 * (kw + 18), "04", "Ýlçe Yönetim Birimi", _
            "", "Ýlçe depo stoklarý" & vbLf & "Ýlçeler arasý malzeme arama" & vbLf & _
            "Demontaj malzeme kontrolü" & vbLf & "ÝSG zimmet kontrolü", _
            "Ýþlemlere geç", "AKSORA_BIRIM_ILCE", RGB(79, 70, 160), True, kw
        BIRIM_TEXT ws, "SECIM_NOT", "Ýhtiyacýnýz olan iþlem alanýna tek seçimle ulaþýn.", 35, 565, 700, 24, 10, RGB(87, 121, 142), False
        ' Yonetici girisi (sifreli yonetim merkezi)
        BIRIM_BUTON ws, "ADMIN", "ADMÝN GÝRÝÞÝ  " & ChrW(8250), cw - 34 - 210, 557, 210, 38, RGB(12, 45, 79), "AKSORA_ADMIN_GIRIS"
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
    ByVal renk As Long, ByVal aktif As Boolean, Optional ByVal kw As Double = 324)
    Dim s As Shape
    BIRIM_RECT ws, ad & "_GOLGE", x + 2, 299, kw, 244, RGB(221, 237, 246), -1, msoShapeRoundedRectangle
    Set s = BIRIM_RECT(ws, ad & "_KART", x, 295, kw, 244, RGB(255, 255, 255), RGB(193, 224, 235), msoShapeRoundedRectangle)
    s.OnAction = BIRIM_MAKRO(makro)
    BIRIM_RECT ws, ad & "_VURGU", x + 18, 315, 4, 33, renk, -1, 1
    BIRIM_TEXT ws, ad & "_NO", numara, x + 35, 315, 60, 31, 19, renk, True
    BIRIM_TEXT ws, ad & "_BASLIK", baslik, x + 21, 367, kw - 38, 31, IIf(kw < 300, 16, 18), RGB(12, 45, 79), True
    BIRIM_TEXT ws, ad & "_DETAY", detay, x + 21, 407, kw - 38, 66, 10, RGB(55, 99, 125), False
    With ws.Shapes("BIRIM_" & ad & "_DETAY").TextFrame2
        .VerticalAnchor = msoAnchorTop
        .TextRange.ParagraphFormat.SpaceAfter = 3
    End With
    BIRIM_BUTON ws, ad & "_GIT", buton & "  " & ChrW(8594), x + 21, 484, kw - 42, 36, renk, makro
    If aktif Then
        BIRIM_RECT ws, ad & "_ROZET", x + kw - 100, 318, 78, 25, RGB(222, 248, 235), -1, msoShapeRoundedRectangle
        BIRIM_TEXT ws, ad & "_ROZET_YAZI", "AKTÝF", x + kw - 87, 321, 54, 19, 8, RGB(0, 111, 80), True
    Else
        BIRIM_RECT ws, ad & "_ROZET", x + kw - 159, 318, 137, 25, RGB(184, 45, 55), -1, msoShapeRoundedRectangle
        BIRIM_TEXT ws, ad & "_ROZET_YAZI", "Yapým Aþamasýnda", x + kw - 148, 321, 117, 19, 8, RGB(255, 255, 255), True
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
Private Sub Y_KUR(Optional ByVal ilce As Boolean)
    Dim adlar As Variant, i As Long, ws As Worksheet, kaynak As Worksheet, gor As XlSheetVisibility, etk As String, tur As String
    Dim da As Boolean, kad As String, adim As String, no As Long, metin As String, yeniAd As String
    Dim onceki As Collection, sh As Object, varmi As Boolean
    Dim duzen As Variant
    If ilce Then adlar = I_ADLAR(): etk = I_ETIKET Else adlar = Y_ADLAR(): etk = Y_ETIKET
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
            If CStr(ws.Range("BX1").Value2) <> etk Then
                adim = CStr(adlar(i)) & " / yarým sayfayý silme"
                ws.Visible = xlSheetVisible
                ws.Delete
                Set ws = Nothing
            End If
        End If
        If ws Is Nothing Then
            If ilce Then tur = I_TUR(CStr(adlar(i))) Else tur = Y_TUR(CStr(adlar(i)))
            If tur = "BOS" Then tur = "RAPOR"
            Select Case tur
                Case "PANEL": kad = "M_ANA PANEL"
                Case "ARAMA": kad = "M_MALZEME_ARA"
                Case "STOK": kad = "M_CDP_STOK"
                Case "KARS": kad = "STOK_KARÞILAÞTIRMA"
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
            If ilce Then
                I_TEMIZLE ws
            Else
                Select Case Y_TUR(ws.Name)
                    Case "PANEL": Y_PANEL_TEMIZLE ws
                    Case "ARAMA": Y_ARAMA_TEMIZLE ws
                    Case "STOK": Y_STOK_TEMIZLE ws
                    Case "RAPOR": Y_RAPOR_TEMIZLE ws
                    Case Else: Y_BOS_KUR ws
                End Select
            End If
            ws.Range("BX1").Value2 = etk
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
Private Sub Y_BANNER(ByVal ws As Worksheet, ByVal kod As String, Optional ByVal dosya As String = "AKSORA_YATIRIM_PANEL_BANTLARI.png")
    Dim yol As String, eski As Shape, s As Shape, h As Double, bolum As Long, sinirlar As Variant
    yol = AKSORA_GORSEL(dosya)
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
    Y_DM_KONTROL ws
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
        If (ws.Name = "Y_ANA PANEL" Or ws.Name = "I_ANA PANEL") And alt > 0 Then z = application.Min(z, (.UsableHeight - 8) / alt * 100)
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
    AKSORA_BEKLE_ILERLE 0.45, "Sayfalar hazýrlanýyor..."
    Y_KUR
    Y_IL_UYGULA kod
    AKSORA_BEKLE_ILERLE 0.55, "Ýl stoklarý yükleniyor..."
    Y_IL_VERI kod
    AKSORA_BEKLE_ILERLE 0.7, "Ana panel hazýrlanýyor..."
    Y_PANEL_YENILE
    Y_ARAMA_TARIH
    AKSORA_BEKLE_ILERLE 0.8, "Teminat ve hakediþ hazýrlanýyor..."
    Y_ALT_SAYFALAR kod
    AKSORA_BEKLE_ILERLE 0.95, "Ekran açýlýyor..."
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
    ' 5 / 6: ilce yonetim birimi (CDP / EPP ilce depolari K001-K999).
    If i >= 5 Then Y_DEPO_UYGUN = I_DEPO_UYGUN(depo): Exit Function
    If i <= 2 Then
        Y_DEPO_UYGUN = (depo = "D002" Or depo = "D011")
    Else
        ' Z ile baslayan depolar (or. Z007) yatirim / yuklenici kapsaminda degildir.
        If Left$(depo, 1) = "Z" Then Exit Function
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
    Y_FILIGRAN ws
    Y_KORU ws
    Set r = ThisWorkbook.Worksheets(Left$(ws.Name, Len(ws.Name) - 5) & "_RAPOR")
    r.Unprotect ""
    Y_YAZ Y_SEKIL(r, "AKT1_RAPOR_BASLIK"), "AKSORA | " & Y_BASLIK(r.Name) & "  |  " & Y_IL_ADI(il)
    AKSORA_MUSTERI_RAPOR_KUR_DIS r
    On Error Resume Next
    AKSORA_YATIRIM_DM_DUZEN r
    Y_FILIGRAN r
    On Error GoTo 0
    Y_BAGLA Y_SEKIL(r, "btnAnaPanelDon"), "Y_ANA PANEL"
    If Left$(ws.Name, 2) = "I_" Then I_BAGLA ws: I_BAGLA r
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
    AKSORA_ADMIN_DURUM metin
    If mYToplu Then Y_TOPLU_DURUM metin
    Y_SERIT "Güncelleniyor  •  " & mYSira & metin, 1
    DoEvents
End Sub

' Guncelleme basinda (tek il ve butun iller): ana panel ekrana gelir, kartlar mavi
' "Güncelleniyor...", serit acik sari (musteri ana paneli temasi). Islem panelden izlenir.
Private Sub Y_CANLI_BASLAT(ByVal sayfa As String, ByVal metin As String)
    Dim ws As Worksheet, k As Variant
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sayfa)
    If ws Is Nothing Then Exit Sub
    If Not (ActiveSheet Is ws) Then Y_SAYFA_SEC ws
    ws.Unprotect ""
    For Each k In Array("B", "L", "Q")
        Y_KART_DURUM ws, CStr(k), "Güncelleniyor...", RGB(59, 130, 246), RGB(29, 78, 216)
    Next k
    Y_KORU ws
    application.ScreenUpdating = True
    Y_SERIT "Güncelleniyor  •  " & mYSira & metin, 1, sayfa
    On Error GoTo 0
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
    Y_CANLI_BASLAT "Y_ANA PANEL", Y_IL_ADI(il) & "  •  SAP baðlantýsý hazýrlanýyor..."
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
        Y_HAKEDIS_KUR il
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

' Kart durumu guncelleme gunune gore: ayni gun GUNCEL; gun degisince (23:59 sonrasi) kac gundur guncellenmedigi.
Private Sub Y_KART_TARIH(ByVal ws As Worksheet, ByVal k As String, ByVal tarih As Date)
    Dim gun As Long
    gun = DateDiff("d", DateValue(tarih), Date)
    If gun <= 0 Then
        Y_KART_DURUM ws, k, "GÜNCEL", RGB(34, 197, 94), RGB(21, 128, 61)
    ElseIf gun = 1 Then
        Y_KART_DURUM ws, k, "1 gündür güncellenmedi", RGB(245, 158, 11), RGB(180, 83, 9)
    Else
        Y_KART_DURUM ws, k, gun & " gündür güncellenmedi", RGB(239, 68, 68), RGB(185, 28, 28)
    End If
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

Private Sub Y_SERIT(ByVal metin As String, ByVal durum As Long, Optional ByVal sayfa As String = "Y_ANA PANEL")
    Dim ws As Worksheet, s As Shape, korumali As Boolean, pt As Double, gen As Double, ust As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sayfa)
    Set s = ws.Shapes("MUSTERI_ANA_SECIM_SERIDI")
    If s Is Nothing Then Exit Sub
    korumali = ws.ProtectDrawingObjects
    If korumali Then ws.Unprotect ""
    s.Fill.ForeColor.RGB = Choose(durum + 1, RGB(224, 242, 254), RGB(255, 246, 219), RGB(209, 250, 229), RGB(254, 226, 226))
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
    s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = Choose(durum + 1, RGB(12, 45, 79), RGB(146, 64, 14), RGB(6, 95, 70), RGB(153, 27, 27))
    s.TextFrame2.TextRange.Font.bold = msoTrue
    If korumali Then Y_KORU ws
    If durum = 1 Then DoEvents
    On Error GoTo 0
End Sub

' Guncelleme bitince ana panel ekrana getirilir.
Private Sub Y_PANEL_GOSTER(Optional ByVal ad As String = "Y_ANA PANEL")
    Dim ws As Worksheet, ev As Boolean
    If AKSORA_ADMIN_TOPLU Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
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
    ' Yonetim merkezinden tum birimler guncellenirken ekran yonetim merkezinde kalir.
    If AKSORA_ADMIN_TOPLU Then Exit Sub
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
    If Not (AKSORA_YATIRIM_SAYFA(ActiveSheet) Or AKSORA_ILCE_SAYFA(ActiveSheet)) Then Exit Sub
    Y_SAYFA_SEC ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
End Sub

' Sekmeler: Ana panel, Teminat, Hakedis, Malzeme arama, ardindan stok/rapor ciftleri.
Private Sub Y_SEKME_SIRASI(Optional ByVal ilce As Boolean)
    Dim adlar As Variant, i As Long, onceki As Worksheet, ws As Worksheet
    If ilce Then adlar = I_ADLAR() Else adlar = Y_ADLAR()
    For i = 1 To UBound(adlar)
        Set onceki = ThisWorkbook.Worksheets(CStr(adlar(i - 1)))
        Set ws = ThisWorkbook.Worksheets(CStr(adlar(i)))
        If ws.Index <> onceki.Index + 1 Then ws.Move after:=onceki
    Next i
End Sub

' Depo bazinda dagilim bolumu: tum sekiller satir konumundan yeniden hesaplanir.
Public Sub AKSORA_YATIRIM_DM_DUZEN(ByVal r As Worksheet)
    Dim bas As Long, bit As Long, ust As Double, s As Shape, z As Shape, hz As Shape, co As ChartObject
    Dim x2 As Double, zw As Double, zh As Double, gs As Double, lx As Double, lw As Double, adim As Double
    Dim m As Long, i As Long, y As Double, t As String
    Y_PAY_HIZALA r
    bas = Val(CStr(r.Range("BX31").Value2)): bit = Val(CStr(r.Range("BX32").Value2))
    If bas < 9 Or bit < bas Then Exit Sub
    Set z = Y_SEKIL(r, "AKSORA_RP_DM_ZEMIN")
    If z Is Nothing Then Exit Sub
    On Error Resume Next
    Set co = r.ChartObjects("AKSORA_RP_DM_GRAFIK")
    On Error GoTo 0
    If co Is Nothing Then Exit Sub
    ' Form listesi dusuk yakinlastirmada en az ~22 piksel cizilir; baslik satiri bu payi tasir.
    If r.Rows(bas).Height < 55 Then r.Rows(bas).RowHeight = 56
    ust = r.Rows(bas).Top
    Set hz = Y_SEKIL(r, "AKSORA_RP_HALKA_ZEMIN")
    If hz Is Nothing Then
        x2 = z.Left: zw = z.Width
    Else
        x2 = hz.Left: zw = hz.Width
    End If
    zh = application.Max(366, r.Rows(bit).Top + r.Rows(bit).Height - ust)
    z.Left = x2: z.Top = ust: z.Width = zw: z.Height = zh
    gs = application.Min(zh - 96, zw * 0.42)
    If gs < 120 Then gs = 120
    co.Left = x2 + zw * 0.04: co.Top = ust + 76: co.Width = gs: co.Height = gs
    Do While Not Y_SEKIL(r, "AKSORA_RP_DM_HALKA_AD_" & (m + 1)) Is Nothing
        m = m + 1
    Loop
    If m < 1 Then m = 1
    lx = co.Left + gs + zw * 0.05: lw = x2 + zw - 14 - lx
    adim = application.Min(24, (gs - 20) / m)
    For Each s In r.Shapes
        t = s.Name
        If Left$(t, 12) = "AKSORA_RP_DM" Then
            Select Case True
                Case t = "AKSORA_RP_DM_BASLIK" Or t = "AKSORA_RP_DM_SECIM_ETIKET"
                    s.Top = ust + 2: s.Height = 24
                    On Error Resume Next
                    s.TextFrame2.VerticalAnchor = msoAnchorMiddle
                    On Error GoTo 0
                Case t = "AKSORA_RP_DM_LISTE"
                    s.Top = ust + 4: s.Height = 18
                Case t = "AKSORA_RP_DM_GRAFIK_BASLIK"
                    s.Left = x2 + 12: s.Top = ust + 10: s.Width = zw - 24
                Case t = "AKSORA_RP_DM_MERKEZ"
                    s.Left = co.Left + gs * 0.2: s.Top = co.Top + gs * 0.36: s.Width = gs * 0.6: s.Height = gs * 0.28
                Case t = "AKSORA_RP_DM_LISTE_BASLIK"
                    s.Left = lx + 16: s.Top = co.Top - 4: s.Width = lw - 80
                Case t = "AKSORA_RP_DM_PAY_BASLIK"
                    s.Left = lx + lw - 60: s.Top = co.Top - 4: s.Width = 60
                Case Left$(t, 24) = "AKSORA_RP_DM_HALKA_RENK_"
                    i = Val(Mid$(t, 25)): y = co.Top + 20 + (i - 1) * adim
                    s.Left = lx: s.Top = y + 4: s.Width = 10: s.Height = 10
                Case Left$(t, 22) = "AKSORA_RP_DM_HALKA_AD_"
                    i = Val(Mid$(t, 23)): y = co.Top + 20 + (i - 1) * adim
                    s.Left = lx + 16: s.Top = y: s.Width = lw - 80: s.Height = 18
                Case Left$(t, 23) = "AKSORA_RP_DM_HALKA_PAY_"
                    i = Val(Mid$(t, 24)): y = co.Top + 20 + (i - 1) * adim
                    s.Left = lx + lw - 60: s.Top = y: s.Width = 60: s.Height = 18
            End Select
        End If
    Next s
End Sub

'' Ust halka lejanti: yuzdeler (PAY_n) her zaman PAY basliginin altinda, ayni genislik ve hizada.
Private Sub Y_PAY_HIZALA(ByVal r As Worksheet)
    Dim b As Shape, s As Shape, i As Long
    Set b = Y_SEKIL(r, "AKSORA_RP_HALKA_PAY_BASLIK")
    If b Is Nothing Then Exit Sub
    On Error Resume Next
    For i = 1 To 20
        Set s = Y_SEKIL(r, "AKSORA_RP_HALKA_PAY_" & i)
        If Not s Is Nothing Then
            s.Left = b.Left: s.Width = b.Width
            s.TextFrame2.MarginLeft = b.TextFrame2.MarginLeft: s.TextFrame2.MarginRight = b.TextFrame2.MarginRight
            s.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignRight
        End If
    Next i
    On Error GoTo 0
End Sub

Private Function Y_PAY_KAYIK(ByVal r As Worksheet) As Boolean
    Dim b As Shape, s As Shape
    Set b = Y_SEKIL(r, "AKSORA_RP_HALKA_PAY_BASLIK")
    Set s = Y_SEKIL(r, "AKSORA_RP_HALKA_PAY_1")
    If b Is Nothing Or s Is Nothing Then Exit Function
    Y_PAY_KAYIK = (Abs(s.Left - b.Left) > 1 Or Abs(s.Width - b.Width) > 1)
End Function

' Sayfa acilirken bolum kaymissa (eski dosya ya da Excel'in yeniden konumlamasi) sessizce toparlanir.
Private Sub Y_DM_KONTROL(ByVal r As Worksheet)
    Dim bas As Long, z As Shape, co As ChartObject, ust As Double, korumali As Boolean
    If Right$(r.Name, 6) <> "_RAPOR" Then Exit Sub
    If Y_PAY_KAYIK(r) Then
        On Error Resume Next
        korumali = r.ProtectContents Or r.ProtectDrawingObjects
        If korumali Then r.Unprotect ""
        Y_PAY_HIZALA r
        If korumali Then Y_KORU r
        On Error GoTo 0
    End If
    bas = Val(CStr(r.Range("BX31").Value2))
    If bas < 9 Then Exit Sub
    Set z = Y_SEKIL(r, "AKSORA_RP_DM_ZEMIN")
    If z Is Nothing Then Exit Sub
    On Error Resume Next
    Set co = r.ChartObjects("AKSORA_RP_DM_GRAFIK")
    On Error GoTo 0
    If co Is Nothing Then Exit Sub
    ust = r.Rows(bas).Top
    If Abs(z.Top - ust) < 1 And Abs(co.Top - ust - 76) < 1 And r.Rows(bas).Height >= 55 Then Exit Sub
    On Error Resume Next
    korumali = r.ProtectContents Or r.ProtectDrawingObjects
    If korumali Then r.Unprotect ""
    AKSORA_YATIRIM_DM_DUZEN r
    If korumali Then Y_KORU r
    On Error GoTo 0
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
            Y_KART_TARIH ws, k, tarih
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

Private Sub Y_ARAMA_TARIH(Optional ByVal ilce As Boolean)
    Dim ws As Worksheet, s As Shape, kod As String, isim As String, t As Date, korumali As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(IIf(ilce, "I_MALZEME_ARA", "Y_MALZEME_ARA"))
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
            If ilce Then t = AKSORA_ILCE_IL_SON_TARIH(kod) Else t = AKSORA_YATIRIM_IL_SON_TARIH(kod)
            s.TextFrame2.TextRange.Text = isim & vbLf & IIf(t > 0, Format$(t, "dd.mm.yyyy hh:nn"), "Henüz güncellenmedi")
        End If
    Next s
    If korumali Then Y_KORU ws
    On Error GoTo 0
End Sub

Private Sub Y_TOPLU_DURUM(ByVal metin As String, Optional ByVal sayfa As String = "Y_MALZEME_ARA")
    Dim ws As Worksheet, s As Shape, korumali As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sayfa)
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
    Dim hataMetni As String, metin As String, sc As Boolean, baslangic As Worksheet, adim As Long
    If AKSORA_V11_MESGUL Or mYCalisiyor Or mYToplu Then Exit Sub
    If Not AKSORA_ADMIN_TOPLU Then
        If MsgBox("Beþ ilin (Artvin, Trabzon, Gümüþhane, Giresun, Rize) yatýrým ve yüklenici stoklarý SAP'den güncellenecek. Devam edilsin mi?", _
            vbYesNo + vbQuestion, "AKSORA | Yatýrým Birimi") <> vbYes Then Exit Sub
    End If
    sc = application.ScreenUpdating
    If TypeOf ActiveSheet Is Worksheet Then Set baslangic = ActiveSheet
    mYToplu = True: bas = Timer
    application.ScreenUpdating = True
    On Error GoTo Hata
    For Each il In Array("ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
        If AKSORA_ADMIN_IPTAL Then sonuc = sonuc & "Durduruldu: kalan iller atlandý." & vbCrLf: Exit For
        bilgi = vbNullString
        adim = adim + 1: mYSira = "Bütün iller " & adim & "/5  •  "
        Y_TOPLU_DURUM Y_IL_ADI(CStr(il)) & " güncelleniyor..."
        If Y_GUNCELLE(CStr(il), "1234", True, bilgi) Then adet = adet + 1
        sonuc = sonuc & bilgi & vbCrLf
    Next il
Bitir:
    On Error Resume Next
    mYToplu = False: mYSira = vbNullString
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
    If AKSORA_ADMIN_TOPLU Then AKSORA_ADMIN_SONUC = AKSORA_ADMIN_SONUC & "YATIRIM:" & vbCrLf & sonuc & IIf(Len(hataMetni) > 0, hataMetni & vbCrLf, ""): Exit Sub
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

'' Kosullu bicim formulleri Excel dilinde istenir (Turkce Excel'de VE/N...): gecici hucreyle cevrilir.
Private Function Y_YEREL(ByVal ws As Worksheet, ByVal formul As String) As String
    Dim c As Range
    Set c = ws.Range("BZ40")
    On Error Resume Next
    c.Formula = formul
    Y_YEREL = c.FormulaLocal
    c.ClearContents
    On Error GoTo 0
    If Len(Y_YEREL) = 0 Then Y_YEREL = formul
End Function

Private Sub Y_SAYFA_SIFIRLA(ByVal ws As Worksheet)
    Dim s As Shape, ad As String, sil As Collection, sv As Variant, son As Long
    Set sil = New Collection
    For Each s In ws.Shapes
        ad = s.Name
        If Not (ad = "AKT1_RAPOR_UST_ZEMIN" Or ad = "AKT1_RAPOR_BASLIK" Or ad = "btnAnaPanelDon" Or InStr(ad, "FILIGRAN") > 0 Or Left$(ad, 5) = "ADMY_") Then sil.Add ad
    Next s
    On Error Resume Next
    For Each sv In sil
        ws.Shapes(CStr(sv)).Delete
    Next sv
    On Error GoTo 0
    ws.AutoFilterMode = False
    ' Onceki kurulumdan kalan satirlar (400 sonrasi dahil) tamamen temizlenir.
    son = application.Max(400, ws.UsedRange.Row + ws.UsedRange.Rows.Count)
    ws.Range("A3:Z" & son).Validation.Delete
    ws.Cells.FormatConditions.Delete
    ws.Range("A3:Z" & son).UnMerge
    ws.Range("A3:Z" & son).Clear
    ws.Range("BZ1:BZ40").Clear
    ws.Range(ws.Rows(3), ws.Rows(son)).Hidden = False
    ws.Range(ws.Rows(3), ws.Rows(son)).RowHeight = ws.StandardHeight
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
    For i = 3 To 4
        Set sh = ThisWorkbook.Worksheets(CStr(Y_STOK_ADLARI()(i - 1)))
        son = AKSORA_M_SON(sh)
        If son >= 7 Then
            v = sh.Range("A7:N" & son).Value2
            For r = 1 To UBound(v, 1)
                If Not IsError(v(r, 2)) And Not IsError(v(r, 5)) Then
                    dp = UCase$(Trim$(CStr(v(r, 2))))
                    If Len(dp) > 0 And Len(Trim$(CStr(v(r, 5)))) > 0 And Left$(dp, 1) <> "Z" Then
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
    ' Sira: ilin yuklenici depo listesi sirasi; listede olmayanlar sonda.
    Dim s2 As Object
    Set s2 = CreateObject("Scripting.Dictionary"): s2.CompareMode = vbTextCompare
    For Each x In AKSORA_IL_YUKLENICI(il)
        If d.Exists(CStr(x)) Then s2(CStr(x)) = d(CStr(x))
    Next x
    For Each x In d.keys
        If Not s2.Exists(CStr(x)) Then s2(CStr(x)) = d(x)
    Next x
    Set Y_YUKLENICI_OZET = s2
End Function

'---------- TEMINAT ----------
' Mantik: teminata yalniz DEPODA duran malzeme tabidir. Sahaya cikmis (kullanilmis, hakedis
' bekleyen) malzeme teminati rahatlatir. Satir bazinda Konum: Depoda / Sahada / Kismi (kismi:
' sahadaki adet girilir). Ayni yukleniciye ait depolar tek yuklenici adinda toplanir.
' Sahada 45 gunu gecen malzeme "hakedis gecikiyor" olarak isaretlenir.

Private Function Y_YUK_AD(ByVal depoAdi As String, ByVal depo As String) As String
    Dim p As Long, sol As String
    depoAdi = Trim$(depoAdi)
    If Len(depoAdi) = 0 Then Y_YUK_AD = depo: Exit Function
    p = InStr(depoAdi, " ")
    If p > 0 Then
        sol = Left$(depoAdi, p - 1)
        If Len(sol) - Len(Replace(sol, ".", "")) >= 2 Then Y_YUK_AD = Trim$(Mid$(depoAdi, p + 1)): Exit Function
    End If
    Y_YUK_AD = depoAdi
End Function

Private Function Y_KAYIT_SOZLUK(ByVal ad As String) As Object
    Dim ws As Worksheet, d As Object, son As Long, v As Variant, r As Long
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set ws = Y_KAYIT_SAYFA(ad)
    son = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If son >= 1 And Len(CStr(ws.Cells(1, 1).Value2)) > 0 Then
        v = ws.Range("A1:D" & son).Value2
        For r = 1 To UBound(v, 1)
            If Len(CStr(v(r, 1))) > 0 Then d(CStr(v(r, 1))) = Array(v(r, 2), v(r, 3), v(r, 4))
        Next r
    End If
    Set Y_KAYIT_SOZLUK = d
End Function

Private Sub Y_KAYIT_YAZ(ByVal ad As String, ByVal anahtar As String, ByVal a As Variant, Optional ByVal b As Variant, Optional ByVal c As Variant)
    Dim ws As Worksheet, sr As Long
    Set ws = Y_KAYIT_SAYFA(ad)
    sr = Y_KAYIT_SATIR(ws, anahtar, True)
    ws.Cells(sr, 2).Value2 = a
    If Not IsMissing(b) Then ws.Cells(sr, 3).Value2 = b
    If Not IsMissing(c) Then ws.Cells(sr, 4).Value = c
End Sub

Private Function Y_ARALIK(ByVal ws As Worksheet, ByVal hucre As String, ByRef ilk As Long, ByRef son As Long) As Boolean
    Dim t As String
    t = CStr(ws.Range(hucre).Value2)
    If InStr(t, "|") = 0 Then Exit Function
    ilk = Val(Split(t, "|")(0)): son = Val(Split(t, "|")(1))
    Y_ARALIK = (ilk > 0 And son >= ilk)
End Function

' Sayfadaki tum girisler kayda yazilir (yeniden kurulumdan once).
Private Sub Y_TEMINAT_KAYDET(ByVal ws As Worksheet)
    Dim il As String, a As Long, b As Long, r As Long, t As String, anahtar As String
    il = UCase$(Trim$(CStr(ws.Range("BX22").Value2)))
    If Len(il) = 0 Then Exit Sub
    If CStr(ws.Range("BX25").Value2) <> "TM_V3" Then Exit Sub
    If Y_ARALIK(ws, "BX23", a, b) Then
        For r = a To b
            If Len(CStr(ws.Cells(r, 1).Value2)) > 0 Then Y_KAYIT_YAZ "_Y_TEMINAT", "T|" & il & "|" & UCase$(CStr(ws.Cells(r, 1).Value2)), ws.Cells(r, 4).Value2
        Next r
    End If
    If Y_ARALIK(ws, "BX27", a, b) Then
        For r = a To b
            If Len(CStr(ws.Cells(r, 3).Value2)) > 0 Then
                anahtar = "K|" & il & "|" & UCase$(CStr(ws.Cells(r, 1).Value2)) & "|" & CStr(ws.Cells(r, 3).Value2)
                t = CStr(ws.Cells(r, 7).Value2)
                If t = "Sahada" Or t = "Kýsmi" Then
                    Y_KAYIT_YAZ "_Y_TEMINAT", anahtar, t, ws.Cells(r, 8).Value2, ws.Cells(r, 13).Value2
                Else
                    Y_KAYIT_YAZ "_Y_TEMINAT", anahtar, "Depoda", Empty, Empty
                End If
            End If
        Next r
    End If
End Sub

' Miktar bicimi: tam sayida ondalik ayirac gosterilmez; birim sayinin yaninda yazilir.
Private Sub Y_MIKTAR_BICIM(ByVal c As Range, Optional ByVal birim As String = "")
    Dim ek As String, v As Variant
    v = c.Value2
    If Len(birim) > 0 Then ek = " """ & Replace(birim, """", "") & """"
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        If v = Int(v) Then
            c.NumberFormat = "#,##0" & ek & ";[Red]-#,##0" & ek
        Else
            c.NumberFormat = "#,##0.###" & ek & ";[Red]-#,##0.###" & ek
        End If
    End If
End Sub

' AKSORA amblemi icerigin altina, sag kenara alinir.
Private Sub Y_FILIGRAN(ByVal ws As Worksheet, Optional ByVal y As Double, Optional ByVal w As Double)
    ' AKSORA amblemi yatirim sayfalarinda kullanilmaz (listelerin ustune biniyordu).
    Dim i As Long
    On Error Resume Next
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(ws.Shapes(i).Name, "FILIGRAN") > 0 Then ws.Shapes(i).Delete
    Next i
    On Error GoTo 0
End Sub

Private Sub Y_TEMINAT_KUR(ByVal il As String)
    Dim ws As Worksheet, kayit As Object, ozet As Object, k As Variant, t As Variant, w As Double, kw As Double, y As Double
    Dim s As Shape, r As Long, i As Long, n As Long, sh As Worksheet, son As Long, v As Variant, dp As String, mk As String
    Dim det As Object, dk As Variant, a1 As Long, b1 As Long, a3 As Long, b3 As Long
    Dim satirlar() As Variant, sira() As Long, j As Long, tmp As Long, rr As Range, db As Databar, liste As Shape, x As Variant
    Dim tarih As Variant, nD As Long, nL As Long, sis As String, blok() As Variant
    Set ws = ThisWorkbook.Worksheets("Y_TEMÝNAT")
    il = UCase$(Trim$(il))
    ws.Unprotect ""
    Y_TEMINAT_KAYDET ws
    Set kayit = Y_KAYIT_SOZLUK("_Y_TEMINAT")
    Y_SAYFA_SIFIRLA ws
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    Y_KOLONLAR ws, Array(70, 230, 110, application.Max(80, w - 1284), 140, 140, 140, 110, 120, 120, 104)
    ws.Columns("L:N").Hidden = True
    Set ozet = Y_YUKLENICI_OZET(il)
    ' Malzeme detayi: depo + malzeme toplami (CDP ve EPP birlikte)
    Set det = CreateObject("Scripting.Dictionary"): det.CompareMode = vbTextCompare
    For i = 3 To 4
        Set sh = ThisWorkbook.Worksheets(CStr(Y_STOK_ADLARI()(i - 1)))
        son = AKSORA_M_SON(sh)
        If son >= 7 Then
            v = sh.Range("A7:N" & son).Value2
            For r = 1 To UBound(v, 1)
                If Not IsError(v(r, 2)) And Not IsError(v(r, 5)) Then
                    dp = UCase$(Trim$(CStr(v(r, 2)))): mk = Trim$(CStr(v(r, 5)))
                    If Len(dp) > 0 And Len(mk) > 0 And Left$(dp, 1) <> "Z" Then
                        If Not det.Exists(dp & "|" & mk) Then det(dp & "|" & mk) = Array(dp, mk, Trim$(CStr(v(r, 6))), Trim$(CStr(v(r, 8))), 0#, 0#)
                        t = det(dp & "|" & mk)
                        If Not IsError(v(r, 7)) Then If IsNumeric(v(r, 7)) Then t(4) = t(4) + CDbl(v(r, 7))
                        If Not IsError(v(r, 11)) Then If VarType(v(r, 11)) <> vbString And IsNumeric(v(r, 11)) Then t(5) = t(5) + CDbl(v(r, 11))
                        det(dp & "|" & mk) = t
                    End If
                End If
            Next r
        End If
    Next i
    nD = ozet.Count: nL = det.Count
    a1 = 10: b1 = a1 + application.Max(nD, 1) - 1
    a3 = b1 + 9: b3 = a3 + application.Max(nL, 1) - 1
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10
    ws.Rows(a1 - 2).RowHeight = 30: ws.Rows(a1 - 1).RowHeight = 30
    ' Detay basligi: baslik + depo secimi, arama etiketi, arama kutusu + ARA, aciklama, tablo basligi
    ws.Rows(a3 - 6).RowHeight = 30: ws.Rows(a3 - 5).RowHeight = 20: ws.Rows(a3 - 4).RowHeight = 28
    ws.Rows(a3 - 3).RowHeight = 18: ws.Rows(a3 - 2).RowHeight = 8: ws.Rows(a3 - 1).RowHeight = 30
    ws.Range(ws.Rows(a1), ws.Rows(b1 + 1)).RowHeight = 22
    ws.Rows(b1 + 2).RowHeight = 14
    ws.Range(ws.Rows(a3), ws.Rows(b3)).RowHeight = 19
    Y_BOLUM ws, a1 - 2, "DEPO BAZINDA TEMÝNAT DURUMU", "Her depo ayrý yüklenici / ayrý hakediþ. Teminata yalnýz DEPODA duran malzeme tabidir."
    Y_BOLUM ws, a3 - 6, "MALZEME DETAYI  –  SAHADA / DEPODA", ""
    Y_BASLIK_SATIRI ws.Range(ws.Cells(a1 - 1, 1), ws.Cells(a1 - 1, 11)), Array("Depo", "Depo Adý", "Sistem", "Teminat Tutarý", _
        "Depoda (teminata tabi)", "Sahada (hakediþ bekleyen)", "Kalan Teminat", "Kullaným", "Durum", "Sahada > 45 gün", "")
    Y_BASLIK_SATIRI ws.Range(ws.Cells(a3 - 1, 1), ws.Cells(a3 - 1, 11)), Array("Depo", "Depo Adý", "Malzeme", "Malzeme Tanýmý", _
        "Miktar", "Tutar", "Konum", "Sahadaki Miktar", "Depodaki Tutar", "Sahadaki Tutar", "Sahada (gün)")
    ' Malzeme detayi (depo, tutar azalan)
    If nL > 0 Then
        ReDim satirlar(1 To nL): ReDim sira(1 To nL)
        i = 0
        For Each dk In det.keys
            i = i + 1: satirlar(i) = det(dk): sira(i) = i
        Next dk
        For i = 2 To nL
            tmp = sira(i): j = i - 1
            Do While j >= 1
                If StrComp(satirlar(sira(j))(0), satirlar(tmp)(0), vbTextCompare) < 0 Then Exit Do
                If StrComp(satirlar(sira(j))(0), satirlar(tmp)(0), vbTextCompare) = 0 And satirlar(sira(j))(5) >= satirlar(tmp)(5) Then Exit Do
                sira(j + 1) = sira(j): j = j - 1
            Loop
            sira(j + 1) = tmp
        Next i
        ReDim blok(1 To nL, 1 To 13)
        For i = 1 To nL
            t = satirlar(sira(i))
            blok(i, 1) = t(0)
            If ozet.Exists(CStr(t(0))) Then blok(i, 2) = ozet(CStr(t(0)))(0)
            blok(i, 3) = CStr(t(1)): blok(i, 4) = t(2)
            blok(i, 5) = t(4): blok(i, 6) = Round(t(5), 2)
            mk = "K|" & il & "|" & UCase$(CStr(t(0))) & "|" & CStr(t(1))
            blok(i, 7) = "Depoda"
            If kayit.Exists(mk) Then
                If CStr(kayit(mk)(0)) = "Sahada" Or CStr(kayit(mk)(0)) = "Kýsmi" Then
                    blok(i, 7) = CStr(kayit(mk)(0)): blok(i, 8) = kayit(mk)(1): blok(i, 13) = kayit(mk)(2)
                End If
            End If
        Next i
        ws.Range(ws.Cells(a3, 1), ws.Cells(b3, 1)).NumberFormat = "@"
        ws.Range(ws.Cells(a3, 3), ws.Cells(b3, 3)).NumberFormat = "@"
        ws.Range(ws.Cells(a3, 1), ws.Cells(b3, 13)).Value2 = blok
        ' Satir formulleri tek seferde (sutun bazinda R1C1; ayni formuller, hucre hucre yazimdan hizli).
        ws.Range(ws.Cells(a3, 9), ws.Cells(b3, 9)).FormulaR1C1 = "=RC6-RC10"
        ws.Range(ws.Cells(a3, 10), ws.Cells(b3, 10)).FormulaR1C1 = "=IF(N(RC5)=0,0,RC6*IF(RC7=""Sahada"",1,IF(RC7=""Kýsmi"",MIN(N(RC8),RC5)/RC5,0)))"
        ws.Range(ws.Cells(a3, 11), ws.Cells(b3, 11)).FormulaR1C1 = "=IF(OR(RC7=""Depoda"",N(RC13)=0),"""",TODAY()-INT(RC13))"
        ws.Range(ws.Cells(a3, 14), ws.Cells(b3, 14)).FormulaR1C1 = "=RC3&"" ""&RC4"
        For r = a3 To b3
            Y_MIKTAR_BICIM ws.Cells(r, 5), CStr(satirlar(sira(r - a3 + 1))(3))
        Next r
    Else
        ws.Cells(a3, 2).Value2 = "Yüklenici depolarýnda stok yok (SAP güncellemesi sonrasý dolar)."
    End If
    ' Depolar
    r = a1
    For Each k In ozet.keys
        t = ozet(k)
        ws.Cells(r, 1).NumberFormat = "@": ws.Cells(r, 1).Value2 = CStr(k)
        ws.Cells(r, 2).Value2 = IIf(Len(t(0)) > 0, t(0), "(depoda stok yok)")
        sis = IIf(t(1) <> 0, "CDP", "") & IIf(t(1) <> 0 And t(2) <> 0, " + ", "") & IIf(t(2) <> 0, "EPP", "")
        ws.Cells(r, 3).Value2 = IIf(Len(sis) > 0, sis, "-")
        If kayit.Exists("T|" & il & "|" & UCase$(CStr(k))) Then
            If Not IsError(kayit("T|" & il & "|" & UCase$(CStr(k)))(0)) Then
                If Len(CStr(kayit("T|" & il & "|" & UCase$(CStr(k)))(0))) > 0 Then ws.Cells(r, 4).Value2 = kayit("T|" & il & "|" & UCase$(CStr(k)))(0)
            End If
        End If
        ws.Cells(r, 5).Formula = "=SUMIFS($I$" & a3 & ":$I$" & b3 & ",$A$" & a3 & ":$A$" & b3 & ",A" & r & ")"
        ws.Cells(r, 6).Formula = "=SUMIFS($J$" & a3 & ":$J$" & b3 & ",$A$" & a3 & ":$A$" & b3 & ",A" & r & ")"
        ws.Cells(r, 7).Formula = "=IF(N(D" & r & ")>0,D" & r & "-E" & r & ","""")"
        ws.Cells(r, 8).Formula = "=IF(N(D" & r & ")>0,E" & r & "/D" & r & ","""")"
        ws.Cells(r, 9).Formula = "=IF(N(D" & r & ")<=0,""Teminat girilmedi"",IF(H" & r & ">1,""AÞIM"",IF(H" & r & ">=0.8,""DÝKKAT"",""GÜVENLÝ"")))"
        ws.Cells(r, 10).Formula = "=COUNTIFS($A$" & a3 & ":$A$" & b3 & ",A" & r & ",$K$" & a3 & ":$K$" & b3 & ","">45"")"
        r = r + 1
    Next k
    ws.Cells(b1 + 1, 2).Value2 = "TOPLAM"
    For Each x In Array(4, 5, 6)
        ws.Cells(b1 + 1, CLng(x)).Formula = "=SUM(" & ws.Range(ws.Cells(a1, CLng(x)), ws.Cells(b1, CLng(x))).Address(False, False) & ")"
    Next x
    ws.Cells(b1 + 1, 7).Formula = "=IF(D" & b1 + 1 & ">0,D" & b1 + 1 & "-SUMIF(D" & a1 & ":D" & b1 & ","">0"",E" & a1 & ":E" & b1 & "),"""")"
    ws.Cells(b1 + 1, 8).Formula = "=IF(D" & b1 + 1 & ">0,SUMIF(D" & a1 & ":D" & b1 & ","">0"",E" & a1 & ":E" & b1 & ")/D" & b1 + 1 & ","""")"
    ws.Cells(b1 + 1, 10).Formula = "=SUM(J" & a1 & ":J" & b1 & ")"
    ' Bicimler
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(b1, 11))
    Y_TABLO_GOVDE ws.Range(ws.Cells(a3, 1), ws.Cells(b3, 11))
    Y_TOPLAM_SATIRI ws.Range(ws.Cells(b1 + 1, 1), ws.Cells(b1 + 1, 11))
    ws.Range(ws.Cells(a1, 1), ws.Cells(b1, 1)).Font.bold = True
    ws.Range(ws.Cells(a1, 1), ws.Cells(b1, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 3), ws.Cells(b1, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a3, 1), ws.Cells(b3, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a3, 3), ws.Cells(b3, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a3, 7), ws.Cells(b3, 8)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a3, 11), ws.Cells(b3, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 2), ws.Cells(b1, 2)).IndentLevel = 1
    ws.Range(ws.Cells(a3, 2), ws.Cells(b3, 2)).IndentLevel = 1
    ws.Range(ws.Cells(a3, 4), ws.Cells(b3, 4)).IndentLevel = 1
    ws.Range(ws.Cells(a1, 4), ws.Cells(b1 + 1, 7)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 8), ws.Cells(b1 + 1, 8)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(a1, 8), ws.Cells(b1 + 1, 10)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a3, 5), ws.Cells(b3, 5)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(a3, 8), ws.Cells(b3, 8)).NumberFormat = "General"
    ws.Range(ws.Cells(a3, 6), ws.Cells(b3, 6)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a3, 9), ws.Cells(b3, 10)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a3, 13), ws.Cells(b3, 13)).NumberFormat = "dd.mm.yyyy"
    If nD > 0 Then Y_GIRIS_HUCRE ws.Range(ws.Cells(a1, 4), ws.Cells(b1, 4))
    If nL > 0 Then
        Y_GIRIS_HUCRE ws.Range(ws.Cells(a3, 7), ws.Cells(b3, 8))
        ws.Range("BZ20").Value2 = "Depoda": ws.Range("BZ21").Value2 = "Sahada": ws.Range("BZ22").Value2 = "Kýsmi"
        With ws.Range(ws.Cells(a3, 7), ws.Cells(b3, 7)).Validation
            .Delete
            .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$BZ$20:$BZ$22"
            .IgnoreBlank = True: .InCellDropdown = True
            .ErrorMessage = "Depoda, Sahada veya Kýsmi seçiniz."
        End With
        Set rr = ws.Range(ws.Cells(a3, 7), ws.Cells(b3, 7))
        Y_DURUM_KF rr, "Sahada", RGB(220, 252, 231), RGB(21, 128, 61)
        Y_DURUM_KF rr, "Kýsmi", RGB(254, 243, 199), RGB(180, 83, 9)
        Set rr = ws.Range(ws.Cells(a3, 11), ws.Cells(b3, 11))
        ' Depoda satirlarinda K bos metindir; yalniz sayi > 45 boyanir.
        With rr.FormatConditions.Add(Type:=xlExpression, Formula1:=Y_YEREL(ws, "=AND(ISNUMBER(K" & a3 & "),K" & a3 & ">45)"))
            .Interior.Color = RGB(254, 226, 226): .Font.Color = RGB(185, 28, 28): .Font.bold = True
        End With
    End If
    Set rr = ws.Range(ws.Cells(a1, 8), ws.Cells(b1, 8))
    On Error Resume Next
    Set db = rr.FormatConditions.AddDatabar
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=1.2
    db.BarColor.Color = RGB(47, 144, 217)
    db.BarFillType = xlDataBarFillSolid
    On Error GoTo 0
    Set rr = ws.Range(ws.Cells(a1, 9), ws.Cells(b1, 9))
    Y_DURUM_KF rr, "AÞIM", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF rr, "DÝKKAT", RGB(254, 243, 199), RGB(180, 83, 9)
    Y_DURUM_KF rr, "GÜVENLÝ", RGB(220, 252, 231), RGB(21, 128, 61)
    Y_DURUM_KF rr, "Teminat girilmedi", RGB(241, 245, 249), RGB(100, 116, 139)
    Set rr = ws.Range(ws.Cells(a1, 10), ws.Cells(b1, 10))
    With rr.FormatConditions.Add(Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0")
        .Font.Color = RGB(185, 28, 28): .Font.bold = True
    End With
    ' Filtre ve depo secimi
    If nL > 0 Then ws.Range(ws.Cells(a3 - 1, 1), ws.Cells(b3, 14)).AutoFilter
    ' Malzeme arama (musteri arama sayfasiyla ayni gorunum): etiket, kutu + ARA, aciklama.
    With ws.Range(ws.Cells(a3 - 5, 2), ws.Cells(a3 - 5, 8))
        .HorizontalAlignment = xlCenterAcrossSelection: .VerticalAlignment = xlBottom
    End With
    ws.Cells(a3 - 5, 2).Value2 = "MALZEME ADI VEYA KODU"
    With ws.Cells(a3 - 5, 2).Font
        .Name = "Aptos": .Size = 11: .bold = True: .Color = RGB(64, 74, 89)
    End With
    With ws.Range(ws.Cells(a3 - 4, 2), ws.Cells(a3 - 4, 7))
        .Merge
        .Locked = False: .NumberFormat = "@"
        .Interior.Color = RGB(255, 255, 255)
        .Font.Name = "Aptos": .Font.Size = 11: .Font.Color = RGB(12, 45, 79)
        .HorizontalAlignment = xlLeft: .VerticalAlignment = xlCenter: .IndentLevel = 1
        With .Borders
            .LineStyle = xlContinuous: .Weight = xlThin: .Color = RGB(165, 190, 214)
        End With
    End With
    Y_DUGME ws, "Y_TM_ARA", "ARA", ws.Cells(a3 - 4, 8).Left + 6, ws.Cells(a3 - 4, 8).Top + 1, _
        ws.Cells(a3 - 4, 8).Width - 6, ws.Rows(a3 - 4).Height - 2, RGB(44, 95, 145), "AKSORA_YATIRIM_TEMINAT_ARA"
    Set s = Y_SEKIL(ws, "Y_TM_ARA")
    If Not s Is Nothing Then s.Adjustments.Item(1) = 0.18: s.TextFrame2.TextRange.Font.Size = 11
    With ws.Range(ws.Cells(a3 - 3, 2), ws.Cells(a3 - 3, 8))
        .HorizontalAlignment = xlCenterAcrossSelection: .VerticalAlignment = xlCenter
    End With
    ws.Cells(a3 - 3, 2).Value2 = "Malzeme adý veya kodunun tamamý ya da bir kýsmý ile arama yapabilirsiniz. Boþ býrakýp Enter: tüm liste."
    With ws.Cells(a3 - 3, 2).Font
        .Name = "Aptos": .Size = 9: .Color = RGB(100, 116, 139)
    End With
    ws.Range(ws.Rows(a3 - 5), ws.Rows(a3 - 3)).Interior.Color = RGB(244, 247, 251)
    ws.Range(ws.Cells(a3 - 4, 2), ws.Cells(a3 - 4, 7)).Interior.Color = RGB(255, 255, 255)
    ws.Range("BX29").Value2 = ws.Cells(a3 - 4, 2).Address(False, False)
    ws.Cells(a3 - 6, 7).Value2 = "Depo:"
    ws.Cells(a3 - 6, 7).HorizontalAlignment = xlRight: ws.Cells(a3 - 6, 7).Font.bold = True: ws.Cells(a3 - 6, 7).Font.Color = RGB(28, 54, 83)
    Set liste = ws.Shapes.AddFormControl(xlDropDown, ws.Cells(a3 - 6, 8).Left + 4, ws.Cells(a3 - 6, 8).Top + 5, _
        ws.Range(ws.Cells(a3 - 6, 8), ws.Cells(a3 - 6, 11)).Width - 6, 20)
    liste.Name = "Y_TM_SECIM": liste.Placement = xlFreeFloating
    liste.ControlFormat.AddItem "TÜM DEPOLAR"
    For Each k In ozet.keys
        t = ozet(k)
        liste.ControlFormat.AddItem CStr(k) & IIf(Len(t(0)) > 0, "  -  " & t(0), "")
    Next k
    liste.ControlFormat.Value = 1
    liste.OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_TEMINAT_SEC")
    tarih = ThisWorkbook.Worksheets(CStr(Y_STOK_ADLARI()(2))).Range("C2").Value
    ws.Cells(b3 + 2, 1).Value2 = "Kullaným = Depoda (teminata tabi) ÷ Teminat.  %80 ve üzeri DÝKKAT, %100 üzeri AÞIM.  Konum: Depoda / Sahada / Kýsmi " & _
        "(kýsmide sahadaki miktarý yazýn).  Sahada 45 günü geçen kýrmýzý.  Depo deðerleri SAP MB52" & _
        IIf(IsDate(tarih), " (" & Format$(tarih, "dd.mm.yyyy hh:nn") & ")", "") & "."
    ws.Cells(b3 + 2, 1).Font.Size = 9: ws.Cells(b3 + 2, 1).Font.Color = RGB(71, 85, 105)
    ' KPI
    ws.Range("BZ1").Formula = "=D" & b1 + 1
    ws.Range("BZ2").Formula = "=E" & b1 + 1
    ws.Range("BZ3").Formula = "=F" & b1 + 1
    ws.Range("BZ4").Formula = "=COUNTIF(I" & a1 & ":I" & b1 & ",""AÞIM"")&"" aþým  •  ""&COUNTIF(I" & a1 & ":I" & b1 & ",""DÝKKAT"")&"" dikkat  •  ""&J" & b1 + 1 & "&"" geciken saha"""
    ws.Range("BZ1:BZ3").NumberFormat = "#,##0.00 ""TL"""
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "TOPLAM TEMÝNAT", "$BZ$1", RGB(35, 119, 204)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "DEPODA  –  TEMÝNATA TABÝ", "$BZ$2", RGB(220, 38, 38)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "SAHADA  –  HAKEDÝÞ BEKLEYEN", "$BZ$3", RGB(39, 160, 94)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "RÝSK DURUMU", "$BZ$4", RGB(245, 158, 11)
    Set s = Y_SEKIL(ws, "Y_KPI_DEGER_4")
    If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 13
    Y_FILIGRAN ws, ws.Rows(b3 + 3).Top + 6, w
    ws.Range("BX22").Value2 = il: ws.Range("BX20").Value2 = il
    ws.Range("BX23").Value2 = a1 & "|" & b1
    ws.Range("BX26").ClearContents
    ws.Range("BX27").Value2 = a3 & "|" & b3
    ws.Range("BX25").Value2 = "TM_V3"
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | TEMÝNAT TAKÝBÝ  |  " & Y_IL_ADI(il)
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Sub Y_BOLUM(ByVal ws As Worksheet, ByVal r As Long, ByVal baslik As String, ByVal notu As String, _
    Optional ByVal notSutun As Long = 11)
    ws.Cells(r, 1).Value2 = baslik
    With ws.Cells(r, 1).Font
        .Name = "Aptos": .Size = 13: .bold = True: .Color = RGB(12, 45, 79)
    End With
    ws.Cells(r, 1).VerticalAlignment = xlCenter
    ' Not, tablonun son sutununa saga yaslanir (tablonun disina tasmaz).
    If notSutun <> 11 Then ws.Cells(r, 11).ClearContents
    ws.Cells(r, notSutun).Value2 = notu
    With ws.Cells(r, notSutun)
        .HorizontalAlignment = xlRight: .VerticalAlignment = xlCenter
        .Font.Name = "Aptos": .Font.Size = 9: .Font.Italic = True: .Font.Color = RGB(100, 116, 139)
    End With
End Sub

'' Filtre sonrasi Teminat sayfasinda kalinir (Excel baska yatirim sayfasina gecerse geri alinir).
Private Sub Y_TEMINAT_KAL(ByVal ws As Worksheet)
    On Error Resume Next
    If Not (ActiveSheet Is ws) Then Y_SAYFA_SEC ws
    mYOdak = ws.Name
    application.OnTime Now + TimeSerial(0, 0, 1), BIRIM_MAKRO("AKSORA_YATIRIM_ODAK")
    On Error GoTo 0
End Sub

Public Sub AKSORA_YATIRIM_TEMINAT_ARA()
    Dim ws As Worksheet, ev As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Y_TEMÝNAT")
    If Len(CStr(ws.Range("BX29").Value2)) = 0 Then Exit Sub
    ev = application.EnableEvents
    application.EnableEvents = True
    AKSORA_YATIRIM_DEGISIM ws, ws.Range(CStr(ws.Range("BX29").Value2))
    application.EnableEvents = ev
End Sub

Public Sub AKSORA_YATIRIM_TEMINAT_SEC()
    Dim ws As Worksheet, s As Shape, t As String, a As Long, b As Long
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_TEMÝNAT")
    Set s = ws.Shapes("Y_TM_SECIM")
    t = Trim$(Split(CStr(s.ControlFormat.List(s.ControlFormat.Value)) & "  -  ", "  -  ")(0))
    If Not Y_ARALIK(ws, "BX27", a, b) Then Exit Sub
    ws.Unprotect ""
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(a - 1, 1), ws.Cells(b, 14)).AutoFilter
    If s.ControlFormat.Value = 1 Then
        ws.AutoFilter.Range.AutoFilter Field:=1
    Else
        ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:="=" & t
    End If
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Y_TEMINAT_KAL ws
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

'---------- HAKEDIS: PROJE - STOK KARSILASTIRMA ----------
' Proje malzeme listesi (Excel) yuklenir; kod / tanim / miktar sutunlari otomatik bulunur.
' Secili yuklenici deposu ve yatirim depolari (D002, D011) ile karsilastirilir.
' Hakedis hesabi uzman gorusuyle kesinlesecek; sayfada yeri hazir.

Private Function Y_NORM(ByVal t As String) As String
    t = UCase$(Trim$(t))
    t = Replace(t, "Ý", "I"): t = Replace(t, "ý", "I"): t = Replace(t, "Þ", "S"): t = Replace(t, "Ð", "G")
    t = Replace(t, "Ü", "U"): t = Replace(t, "Ö", "O"): t = Replace(t, "Ç", "C")
    t = Replace(t, "þ", "S"): t = Replace(t, "ð", "G"): t = Replace(t, "ü", "U"): t = Replace(t, "ö", "O"): t = Replace(t, "ç", "C")
    Y_NORM = t
End Function

Private Function Y_KOD_NORM(ByVal v As Variant) As String
    Dim t As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbDouble Or VarType(v) = vbLong Or VarType(v) = vbInteger Or VarType(v) = vbCurrency Then
        If v = Int(v) And Abs(v) < 1E+15 Then t = Format$(v, "0") Else t = CStr(v)
    Else
        t = Trim$(CStr(v))
    End If
    t = Replace(Replace(t, " ", ""), ChrW(160), "")
    If Len(t) > 2 And Right$(t, 2) = ".0" Then t = Left$(t, Len(t) - 2)
    If t Like String$(Len(t), "#") And Len(t) > 0 Then
        Do While Len(t) > 1 And Left$(t, 1) = "0": t = Mid$(t, 2): Loop
    End If
    Y_KOD_NORM = t
End Function

Private Function Y_SAYI_DENE(ByVal v As Variant, ByRef x As Double) As Boolean
    On Error GoTo Yok
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If Len(Trim$(v)) = 0 Then Exit Function
        x = AKSORA_YATIRIM_SAYI(v)
    ElseIf IsNumeric(v) Then
        x = CDbl(v)
    Else
        Exit Function
    End If
    Y_SAYI_DENE = True
Yok:
End Function

' Bir sayfadaki proje listesini okur; bulunan kalem sayisini dondurur.
Private Function Y_PROJE_TARA(ByVal v As Variant, ByVal d As Object, ByRef bilgi As String) As Long
    Dim nr As Long, nc As Long, r As Long, c As Long, h As Long, t As String
    Dim cKod As Long, cTan As Long, cMik As Long, cBir As Long, cFiy As Long, en As Long, sayac As Long
    Dim kod As String, x As Double, f As Double, a As Variant, adet As Long
    If Not IsArray(v) Then Exit Function
    nr = UBound(v, 1): nc = UBound(v, 2)
    For r = 1 To application.Min(40, nr)
        cKod = 0: cTan = 0: cMik = 0: cBir = 0: cFiy = 0
        For c = 1 To nc
            If Not IsError(v(r, c)) Then
                t = Y_NORM(CStr(v(r, c)))
                If Len(t) > 0 And Len(t) < 40 Then
                    If cFiy = 0 And InStr(t, "FIYAT") > 0 Then
                        cFiy = c
                    ElseIf cKod = 0 And (InStr(t, "MALZEME KOD") > 0 Or InStr(t, "MALZEME NO") > 0 Or InStr(t, "STOK KOD") > 0 Or _
                        InStr(t, "SAP KOD") > 0 Or InStr(t, "SAP NO") > 0 Or t = "KOD" Or t = "KODU" Or t = "MALZEME" Or InStr(t, "MATERIAL") > 0) Then
                        cKod = c
                    ElseIf cTan = 0 And (InStr(t, "TANIM") > 0 Or InStr(t, "ACIKLAMA") > 0 Or InStr(t, "MALZEME ADI") > 0 Or _
                        InStr(t, "CINSI") > 0 Or t = "ADI" Or InStr(t, "DESCRIPTION") > 0) Then
                        cTan = c
                    ElseIf cMik = 0 And (InStr(t, "MIKTAR") > 0 Or t = "ADET" Or InStr(t, "METRAJ") > 0 Or t = "QTY" Or InStr(t, "QUANTITY") > 0) Then
                        cMik = c
                    ElseIf cBir = 0 And (t = "BIRIM" Or t = "BIRIMI" Or InStr(t, "OLCU") > 0 Or t = "UNIT") Then
                        cBir = c
                    End If
                End If
            End If
        Next c
        If cKod > 0 And cMik > 0 Then h = r: Exit For
    Next r
    If h = 0 Then
        ' Basliksiz liste: en cok 7-12 haneli sayi iceren sutun kod, sagindaki sayisal sutun miktar.
        For c = 1 To nc
            sayac = 0
            For r = 1 To application.Min(nr, 300)
                kod = Y_KOD_NORM(v(r, c))
                If Len(kod) >= 7 And Len(kod) <= 12 And kod Like String$(Len(kod), "#") Then sayac = sayac + 1
            Next r
            If sayac > en Then en = sayac: cKod = c
        Next c
        If en < 3 Then Exit Function
        For c = cKod + 1 To nc
            sayac = 0
            For r = 1 To application.Min(nr, 300)
                If Y_SAYI_DENE(v(r, c), x) Then sayac = sayac + 1
            Next r
            If sayac >= en * 0.8 Then cMik = c: Exit For
        Next c
        If cMik = 0 Then Exit Function
        For c = cKod + 1 To cMik - 1
            If VarType(v(1, c)) = vbString Then cTan = c: Exit For
        Next c
        h = 0
    End If
    For r = h + 1 To nr
        kod = Y_KOD_NORM(v(r, cKod))
        If Len(kod) >= 4 And Y_SAYI_DENE(v(r, cMik), x) Then
            If x <> 0 Then
                If Not d.Exists(kod) Then d(kod) = Array(vbNullString, vbNullString, 0#, 0#)
                a = d(kod)
                If cTan > 0 And Len(a(0)) = 0 Then If Not IsError(v(r, cTan)) Then a(0) = Trim$(CStr(v(r, cTan)))
                If cBir > 0 And Len(a(1)) = 0 Then If Not IsError(v(r, cBir)) Then a(1) = Trim$(CStr(v(r, cBir)))
                a(2) = a(2) + x
                If cFiy > 0 Then If Y_SAYI_DENE(v(r, cFiy), f) Then a(3) = f
                d(kod) = a: adet = adet + 1
            End If
        End If
    Next r
    bilgi = "kod " & Y_SUTUN(cKod) & ", miktar " & Y_SUTUN(cMik) & IIf(cTan > 0, ", taným " & Y_SUTUN(cTan), "") & _
        IIf(cBir > 0, ", birim " & Y_SUTUN(cBir), "") & IIf(cFiy > 0, ", fiyat " & Y_SUTUN(cFiy), "") & IIf(h > 0, " (baþlýk satýrý " & h & ")", " (baþlýksýz)")
    Y_PROJE_TARA = adet
End Function

Private Function Y_SUTUN(ByVal c As Long) As String
    If c <= 0 Then Exit Function
    Y_SUTUN = Split(ThisWorkbook.Worksheets(1).Cells(1, c).Address(True, False), "$")(0)
End Function

Public Sub AKSORA_YATIRIM_PROJE_YUKLE()
    Dim f As Variant, wb As Workbook, sh As Worksheet, d As Object, en As Object, bilgi As String, enBilgi As String
    Dim n As Long, enN As Long, enAd As String, il As String, ws As Worksheet, g As Worksheet, k As Variant, r As Long, a As Variant
    Dim sc As Boolean, ev As Boolean, da As Boolean, blok() As Variant, guv As Long
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    il = UCase$(Trim$(CStr(ws.Range("BX22").Value2)))
    f = application.GetOpenFilename("Excel ve CSV (*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv),*.xlsx;*.xlsm;*.xls;*.xlsb;*.csv", , "Proje malzeme listesini seçin")
    If VarType(f) = vbBoolean Then Exit Sub
    sc = application.ScreenUpdating: ev = application.EnableEvents: da = application.DisplayAlerts
    application.ScreenUpdating = False: application.EnableEvents = False: application.DisplayAlerts = False
    application.StatusBar = "AKSORA | Proje listesi okunuyor..."
    ' Proje dosyasindaki makrolar calistirilmaz.
    guv = application.AutomationSecurity
    application.AutomationSecurity = msoAutomationSecurityForceDisable
    Set wb = Workbooks.Open(Filename:=CStr(f), ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)
    application.AutomationSecurity = guv
    For Each sh In wb.Worksheets
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
        bilgi = vbNullString
        n = Y_PROJE_TARA(sh.UsedRange.Value2, d, bilgi)
        If n > enN Then enN = n: Set en = d: enBilgi = bilgi: enAd = sh.Name
    Next sh
    wb.Close SaveChanges:=False
    Set wb = Nothing
    application.StatusBar = False
    application.DisplayAlerts = da
    If enN = 0 Then
        application.ScreenUpdating = sc: application.EnableEvents = ev
        MsgBox "Dosyada malzeme kodu ve miktar sütunu bulunamadý." & vbCrLf & _
            "Baþlýklarda 'Malzeme Kodu' ve 'Miktar' (veya 'Adet') olmalý.", vbExclamation, "AKSORA | Proje"
        Exit Sub
    End If
    ' Kayit: _Y_PROJE (A il, B kod, C tanim, D birim, E miktar, F fiyat); ust bilgi "#il" satirinda
    Set g = Y_KAYIT_SAYFA("_Y_PROJE")
    For r = g.Cells(g.Rows.Count, 1).End(xlUp).Row To 1 Step -1
        If CStr(g.Cells(r, 1).Value2) = il Or CStr(g.Cells(r, 1).Value2) = "#" & il Then g.Rows(r).Delete
    Next r
    r = g.Cells(g.Rows.Count, 1).End(xlUp).Row
    If Len(CStr(g.Cells(1, 1).Value2)) = 0 Then r = 0
    ReDim blok(1 To en.Count + 1, 1 To 6)
    blok(1, 1) = "#" & il: blok(1, 2) = Mid$(CStr(f), InStrRev(CStr(f), application.PathSeparator) + 1)
    blok(1, 3) = enAd: blok(1, 4) = enBilgi: blok(1, 5) = Format$(Now, "dd.mm.yyyy hh:nn"): blok(1, 6) = en.Count
    n = 1
    For Each k In en.keys
        n = n + 1: a = en(k)
        blok(n, 1) = il: blok(n, 2) = CStr(k): blok(n, 3) = a(0): blok(n, 4) = a(1): blok(n, 5) = a(2): blok(n, 6) = a(3)
    Next k
    g.Range(g.Cells(r + 1, 1), g.Cells(r + n, 2)).NumberFormat = "@"
    g.Range(g.Cells(r + 1, 1), g.Cells(r + n, 6)).Value2 = blok
    ws.Unprotect ""
    Y_PROJE_KARSILASTIR ws
    Y_HAKEDIS_ALT ws
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    application.ScreenUpdating = sc: application.EnableEvents = ev
    MsgBox "Proje yüklendi: " & blok(1, 2) & vbCrLf & "Sayfa: " & enAd & "   |   " & en.Count & " farklý malzeme" & vbCrLf & _
        "Okunan sütunlar: " & enBilgi, vbInformation, "AKSORA | Proje"
    Exit Sub
Hata:
    On Error Resume Next
    If guv <> 0 Then application.AutomationSecurity = guv
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    application.StatusBar = False
    application.ScreenUpdating = True: application.EnableEvents = True: application.DisplayAlerts = True
    MsgBox "Proje dosyasý okunamadý: " & Err.Description, vbExclamation, "AKSORA | Proje"
End Sub

' Proje ile stok karsilastirmasi (tablo VBA ile yazilir).
Private Sub Y_PROJE_KARSILASTIR(ByVal ws As Worksheet)
    Dim il As String, g As Worksheet, son As Long, v As Variant, r As Long, proje As Object, ust As Variant
    Dim yuk As Object, yat As Object, i As Long, sh As Worksheet, sv As Variant, kod As String, depo As String, secim As String
    Dim t As Variant, k As Variant, n As Long, blok() As Variant, sira() As Long, oncelik() As Long, j As Long, tmp As Long
    Dim a As Long, b As Long, eksikTL As Double, yeterli As Long, eksik As Long, fazla As Long, p As Double, e As Double, y As Double, fy As Double
    Dim satir() As Variant, ad As String
    il = UCase$(Trim$(CStr(ws.Range("BX22").Value2)))
    secim = UCase$(Trim$(CStr(ws.Range("BX24").Value2)))
    a = 13
    ' Eski tablo temizlenir
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    son = application.Max(a + 1, Val(CStr(ws.Range("BX28").Value2)))
    With ws.Range(ws.Cells(a, 1), ws.Cells(son + 2, 11))
        .Clear
    End With
    ws.Range(ws.Rows(a), ws.Rows(son + 2)).RowHeight = ws.StandardHeight
    ' Proje
    Set proje = CreateObject("Scripting.Dictionary"): proje.CompareMode = vbTextCompare
    Set g = Y_KAYIT_SAYFA("_Y_PROJE")
    son = g.Cells(g.Rows.Count, 1).End(xlUp).Row
    If son >= 1 And Len(CStr(g.Cells(1, 1).Value2)) > 0 Then
        v = g.Range("A1:F" & son).Value2
        For r = 1 To UBound(v, 1)
            If CStr(v(r, 1)) = il Then
                proje(CStr(v(r, 2))) = Array(v(r, 3), v(r, 4), v(r, 5), v(r, 6))
            ElseIf CStr(v(r, 1)) = "#" & il Then
                ust = Array(v(r, 2), v(r, 3), v(r, 4), v(r, 5), v(r, 6))
            End If
        Next r
    End If
    ' Stoklar: yuklenici (secili depo veya tumu) ve yatirim depolari
    Set yuk = CreateObject("Scripting.Dictionary"): yuk.CompareMode = vbTextCompare
    Set yat = CreateObject("Scripting.Dictionary"): yat.CompareMode = vbTextCompare
    For i = 1 To 4
        Set sh = ThisWorkbook.Worksheets(CStr(Y_STOK_ADLARI()(i - 1)))
        son = AKSORA_M_SON(sh)
        If son >= 7 Then
            sv = sh.Range("A7:N" & son).Value2
            For r = 1 To UBound(sv, 1)
                kod = Y_KOD_NORM(sv(r, 5))
                If Len(kod) > 0 And Not IsError(sv(r, 2)) Then
                    depo = UCase$(Trim$(CStr(sv(r, 2))))
                    If i <= 2 Then
                        Y_STOK_EKLE yat, kod, sv, r
                    ElseIf Len(secim) = 0 Or secim = depo Then
                        Y_STOK_EKLE yuk, kod, sv, r
                    End If
                End If
            Next r
        End If
    Next i
    ' Satirlar: proje kalemleri + depoda olup projede olmayanlar
    n = proje.Count
    If proje.Count > 0 Then
        For Each k In yuk.keys
            If Not proje.Exists(k) Then n = n + 1
        Next k
    End If
    If n > 0 Then
        ReDim blok(1 To n, 1 To 11): ReDim oncelik(1 To n): ReDim sira(1 To n)
        i = 0
        For Each k In proje.keys
            i = i + 1: t = proje(k)
            p = CDbl(t(2)): e = 0: y = 0: fy = 0
            If yuk.Exists(k) Then e = yuk(k)(2)
            If yat.Exists(k) Then y = yat(k)(2)
            If Not IsError(t(3)) Then
                If IsNumeric(t(3)) Then If CDbl(t(3)) <> 0 Then fy = CDbl(t(3))
            End If
            If fy = 0 And yuk.Exists(k) Then If yuk(k)(2) <> 0 Then fy = yuk(k)(3) / yuk(k)(2)
            If fy = 0 And yat.Exists(k) Then If yat(k)(2) <> 0 Then fy = yat(k)(3) / yat(k)(2)
            ad = CStr(t(0))
            If Len(ad) = 0 And yuk.Exists(k) Then ad = yuk(k)(0)
            If Len(ad) = 0 And yat.Exists(k) Then ad = yat(k)(0)
            blok(i, 1) = CStr(k): blok(i, 2) = ad
            blok(i, 3) = IIf(Len(CStr(t(1))) > 0, t(1), IIf(yuk.Exists(k), yuk(k)(1), IIf(yat.Exists(k), yat(k)(1), "")))
            blok(i, 4) = p: blok(i, 5) = e: blok(i, 6) = y: blok(i, 7) = e - p
            If p > 0 Then blok(i, 8) = application.Min(1, e / p) Else blok(i, 8) = Empty
            If e >= p Then
                blok(i, 9) = "Depoda yeterli": oncelik(i) = 3: yeterli = yeterli + 1
            ElseIf e + y >= p Then
                blok(i, 9) = "Yatýrýmdan tamamlanýr": oncelik(i) = 2
            Else
                blok(i, 9) = "EKSÝK": oncelik(i) = 1: eksik = eksik + 1
            End If
            blok(i, 10) = IIf(fy > 0, fy, Empty)
            blok(i, 11) = application.Max(0, p - e - y) * fy
            eksikTL = eksikTL + blok(i, 11)
        Next k
        For Each k In yuk.keys
            If Not proje.Exists(k) And proje.Count > 0 Then
                i = i + 1: t = yuk(k)
                blok(i, 1) = CStr(k): blok(i, 2) = t(0): blok(i, 3) = t(1)
                blok(i, 4) = 0: blok(i, 5) = t(2): blok(i, 6) = IIf(yat.Exists(k), yat(k)(2), 0): blok(i, 7) = t(2)
                blok(i, 9) = "Projede yok (iade)": oncelik(i) = 4: fazla = fazla + 1
                If t(2) <> 0 Then blok(i, 10) = t(3) / t(2)
                blok(i, 11) = Empty
            End If
        Next k
        For i = 1 To n: sira(i) = i: Next i
        For i = 2 To n
            tmp = sira(i): j = i - 1
            Do While j >= 1
                If oncelik(sira(j)) <= oncelik(tmp) Then Exit Do
                sira(j + 1) = sira(j): j = j - 1
            Loop
            sira(j + 1) = tmp
        Next i
        ReDim satir(1 To n, 1 To 11)
        For i = 1 To n
            For j = 1 To 11: satir(i, j) = blok(sira(i), j): Next j
        Next i
        b = a + n - 1
        ws.Range(ws.Cells(a, 1), ws.Cells(b, 1)).NumberFormat = "@"
        ws.Range(ws.Cells(a, 1), ws.Cells(b, 11)).Value2 = satir
        ws.Range(ws.Rows(a), ws.Rows(b + 1)).RowHeight = 20
        Y_TABLO_GOVDE ws.Range(ws.Cells(a, 1), ws.Cells(b, 11))
        ws.Cells(b + 1, 1).Value2 = "TOPLAM"
        ws.Cells(b + 1, 4).Formula = "=SUM(D" & a & ":D" & b & ")"
        ws.Cells(b + 1, 5).Formula = "=SUM(E" & a & ":E" & b & ")"
        ws.Cells(b + 1, 6).Formula = "=SUM(F" & a & ":F" & b & ")"
        ws.Cells(b + 1, 11).Formula = "=SUM(K" & a & ":K" & b & ")"
        Y_TOPLAM_SATIRI ws.Range(ws.Cells(b + 1, 1), ws.Cells(b + 1, 11))
        ws.Range(ws.Cells(a, 1), ws.Cells(b, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a, 2), ws.Cells(b, 2)).IndentLevel = 1
        ws.Range(ws.Cells(a, 3), ws.Cells(b, 3)).HorizontalAlignment = xlCenter
        For i = a To b + 1
            For j = 4 To 7
                Y_MIKTAR_BICIM ws.Cells(i, j)
            Next j
        Next i
        ws.Range(ws.Cells(a, 8), ws.Cells(b, 8)).NumberFormat = "0%"
        ws.Range(ws.Cells(a, 8), ws.Cells(b, 9)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a, 10), ws.Cells(b + 1, 11)).NumberFormat = "#,##0.00 ""TL"""
        Y_DURUM_KF ws.Range(ws.Cells(a, 9), ws.Cells(b, 9)), "EKSÝK", RGB(254, 226, 226), RGB(185, 28, 28)
        Y_DURUM_KF ws.Range(ws.Cells(a, 9), ws.Cells(b, 9)), "Yatýrýmdan tamamlanýr", RGB(254, 243, 199), RGB(180, 83, 9)
        Y_DURUM_KF ws.Range(ws.Cells(a, 9), ws.Cells(b, 9)), "Depoda yeterli", RGB(220, 252, 231), RGB(21, 128, 61)
        Y_DURUM_KF ws.Range(ws.Cells(a, 9), ws.Cells(b, 9)), "Projede yok (iade)", RGB(237, 233, 254), RGB(91, 33, 182)
        ws.Range(ws.Cells(a - 1, 1), ws.Cells(b, 11)).AutoFilter
    Else
        b = a
        ws.Cells(a, 2).Value2 = "Proje yüklenmedi. 'PROJE EXCELÝ YÜKLE' ile malzeme listesini seçin (kod ve miktar sütunlarý otomatik bulunur)."
        ws.Cells(a, 2).Font.Italic = True: ws.Cells(a, 2).Font.Color = RGB(100, 116, 139)
    End If
    ws.Range("BX28").Value2 = b + 1
    ' KPI degerleri
    If IsArray(ust) Then
        ws.Range("BZ1").Value2 = CStr(ust(0))
        ws.Range("BZ5").Value2 = "Yüklendi: " & CStr(ust(3)) & "   |   " & CStr(ust(2))
    Else
        ws.Range("BZ1").Value2 = "Proje yüklenmedi"
        ws.Range("BZ5").Value2 = ""
    End If
    ws.Range("BZ2").Value2 = proje.Count & " kalem"
    If proje.Count > 0 Then
        ws.Range("BZ3").Value2 = yeterli & " / " & proje.Count & " kalem  (" & Format$(yeterli / proje.Count, "0%") & ")"
    Else
        ws.Range("BZ3").Value2 = "-"
    End If
    ws.Range("BZ4").Value2 = eksik & " kalem  •  " & Format$(eksikTL, "#,##0.00") & " TL"
    ws.Cells(a - 3, 1).Value2 = "Karþýlaþtýrýlan depo: " & IIf(Len(secim) > 0, secim, "TÜM YÜKLENÝCÝ DEPOLARI") & _
        "  +  yatýrým depolarý (D002, D011)" & IIf(fazla > 0, "   |   projede olmayan " & fazla & " kalem depoda (iade / teminat)", "")
    ws.Cells(a - 3, 1).Font.Size = 9: ws.Cells(a - 3, 1).Font.Color = RGB(71, 85, 105)
End Sub

Private Sub Y_STOK_EKLE(ByVal d As Object, ByVal kod As String, ByRef sv As Variant, ByVal r As Long)
    Dim t As Variant
    If Not d.Exists(kod) Then d(kod) = Array(Trim$(CStr(sv(r, 6))), Trim$(CStr(sv(r, 8))), 0#, 0#)
    t = d(kod)
    If Not IsError(sv(r, 7)) Then If IsNumeric(sv(r, 7)) Then t(2) = t(2) + CDbl(sv(r, 7))
    If Not IsError(sv(r, 11)) Then If VarType(sv(r, 11)) <> vbString And IsNumeric(sv(r, 11)) Then t(3) = t(3) + CDbl(sv(r, 11))
    d(kod) = t
End Sub

Private Sub Y_HAKEDIS_KUR(ByVal il As String)
    Dim ws As Worksheet, w As Double, kw As Double, y As Double, s As Shape, liste As Shape, k As Variant, ozet As Object
    Dim secim As String, p As Long, i As Long, x As Double, kutular As Variant, renk As Variant, ust As Double, t As Variant
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    il = UCase$(Trim$(il))
    Y_KAYIT_SAYFA "_Y_PROJE"
    ws.Unprotect ""
    secim = CStr(ws.Range("BX24").Value2)
    If CStr(ws.Range("BX22").Value2) <> il Then secim = vbNullString
    Y_SAYFA_SIFIRLA ws
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    Y_KOLONLAR ws, Array(110, application.Max(80, w - 1214), 70, 120, 130, 130, 120, 100, 160, 120, 124)
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10
    ws.Rows(8).RowHeight = 40
    ws.Rows(9).RowHeight = 12
    ws.Rows(10).RowHeight = 18
    ws.Rows(11).RowHeight = 30
    ws.Rows(12).RowHeight = 32
    ws.Range("BX22").Value2 = il: ws.Range("BX20").Value2 = il
    ws.Range("BX24").Value2 = secim
    ' Dugmeler ve depo secimi
    Y_DUGME ws, "Y_HK_YUKLE", "PROJE EXCELÝ YÜKLE", 0, ws.Rows(8).Top + 4, 220, 32, RGB(0, 125, 110), "AKSORA_YATIRIM_PROJE_YUKLE"
    Y_DUGME ws, "Y_HK_YENILE", "KARÞILAÞTIRMAYI YENÝLE", 232, ws.Rows(8).Top + 4, 220, 32, RGB(16, 57, 94), "AKSORA_YATIRIM_PROJE_YENILE"
    Y_DUGME ws, "Y_HK_PDF", "PDF AL", 464, ws.Rows(8).Top + 4, 120, 32, RGB(35, 119, 204), "AKSORA_YATIRIM_HAKEDIS_PDF"
    Y_METIN ws, "Y_HK_DEPO_ETIKET", "Karþýlaþtýrýlacak yüklenici deposu:", w - 640, ws.Rows(8).Top + 4, 230, 32, 10, RGB(28, 54, 83), True, msoAlignRight
    Set liste = ws.Shapes.AddFormControl(xlDropDown, w - 400, ws.Rows(8).Top + 10, 400, 20)
    liste.Name = "Y_HK_DEPO": liste.Placement = xlFreeFloating
    liste.ControlFormat.AddItem "TÜM YÜKLENÝCÝ DEPOLARI"
    p = 1
    Set ozet = Y_YUKLENICI_OZET(il)
    i = 1
    For Each k In ozet.keys
        i = i + 1: t = ozet(k)
        liste.ControlFormat.AddItem CStr(k) & IIf(Len(t(0)) > 0, "  -  " & t(0), "")
        If UCase$(CStr(k)) = UCase$(secim) Then p = i
    Next k
    liste.ControlFormat.Value = p
    liste.OnAction = BIRIM_MAKRO("AKSORA_YATIRIM_HAKEDIS_SEC")
    ' Karsilastirma bolumu
    Y_BOLUM ws, 11, "PROJE  –  STOK KARÞILAÞTIRMASI", "Sýra: eksikler, yatýrýmdan tamamlananlar, depoda yeterli olanlar, projede olmayanlar (iade)."
    Y_BASLIK_SATIRI ws.Range("A12:K12"), Array("Malzeme", "Malzeme Tanýmý", "Birim", "Proje Miktarý", "Yüklenici Deposu", _
        "Yatýrým Deposu (D002/D011)", "Fark (Depo - Proje)", "Karþýlanma", "Durum", "Birim Fiyat", "Eksik Tutar")
    Y_PROJE_KARSILASTIR ws
    ' KPI
    ws.Range("BZ1:BZ5").NumberFormat = "@"
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "PROJE", "$BZ$1", RGB(35, 119, 204)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "PROJE KALEMÝ", "$BZ$2", RGB(35, 119, 204)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "YÜKLENÝCÝ DEPOSUNDA YETERLÝ", "$BZ$3", RGB(39, 160, 94)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "EKSÝK MALZEME", "$BZ$4", RGB(220, 38, 38)
    For i = 1 To 4
        Set s = Y_SEKIL(ws, "Y_KPI_DEGER_" & i)
        If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 14
    Next i
    ' Hakedis hesabi: uzman gorusu sonrasi (tasarim yeri)
    ust = ws.Rows(CLng(Val(CStr(ws.Range("BX28").Value2))) + 2).Top + 10
    Y_HAKEDIS_TASLAK ws, ust, w
    Y_FILIGRAN ws, ust + 200, w
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | HAKEDÝÞ  |  " & Y_IL_ADI(il)
    ' Yazici yoksa sayfa ayari atlanir.
    On Error Resume Next
    ws.PageSetup.Orientation = xlLandscape
    ws.PageSetup.Zoom = False: ws.PageSetup.FitToPagesWide = 1: ws.PageSetup.FitToPagesTall = False
    ws.PageSetup.PrintArea = "$A$1:$K$" & CLng(Val(CStr(ws.Range("BX28").Value2)))
    On Error GoTo 0
    ws.Range("BX25").Value2 = "HK_V2"
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

' Hakedis hesabi icin ayrilan alan: akis kartlari (uzman gorusuyle doldurulacak).
Private Sub Y_HAKEDIS_TASLAK(ByVal ws As Worksheet, ByVal y As Double, ByVal w As Double)
    Dim s As Shape, i As Long, kw As Double, x As Double, basliklar As Variant, metinler As Variant, renk As Variant
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 8) = "Y_HK_TS_" Then ws.Shapes(i).Delete
    Next i
    Y_KUTU ws, "Y_HK_TS_ZEMIN", 0, y, w, 190, RGB(245, 250, 255), RGB(207, 227, 242)
    Y_METIN ws, "Y_HK_TS_ETIKET", "HAKEDÝÞ HESABI  •  TASARIM", 20, y + 12, 400, 18, 9, RGB(0, 125, 110), True, msoAlignLeft
    Y_METIN ws, "Y_HK_TS_BASLIK", "Projeden hakediþe tek akýþ", 20, y + 32, w - 40, 28, 18, RGB(12, 45, 79), True, msoAlignLeft
    basliklar = Array("1  Proje listesi", "2  Stok karþýlaþtýrmasý", "3  Sahada kullanýlan", "4  Hakediþ tutarý")
    metinler = Array("Proje Exceli yüklenir; kod, taným, miktar otomatik okunur.", _
        "Yüklenici ve yatýrým depolarýyla kalem kalem karþýlaþtýrýlýr.", _
        "Teminat Takibi'nde 'Sahada' iþaretlenen malzeme hakediþe baðlanýr.", _
        "Birim fiyat × metraj, kesintiler ve ödenecek tutar (uzman görüþüyle kesinleþecek).")
    renk = Array(RGB(35, 119, 204), RGB(39, 160, 94), RGB(245, 158, 11), RGB(0, 125, 110))
    kw = (w - 40 - 3 * 14) / 4
    For i = 0 To 3
        x = 20 + i * (kw + 14)
        Y_KUTU ws, "Y_HK_TS_KART_" & i, x, y + 72, kw, 100, RGB(255, 255, 255), RGB(207, 227, 242)
        Y_KUTU ws, "Y_HK_TS_SERIT_" & i, x, y + 72, 6, 100, CLng(renk(i)), -1
        Y_METIN ws, "Y_HK_TS_KB_" & i, CStr(basliklar(i)), x + 18, y + 82, kw - 30, 22, 12, RGB(12, 45, 79), True, msoAlignLeft
        Y_METIN ws, "Y_HK_TS_KM_" & i, CStr(metinler(i)), x + 18, y + 108, kw - 30, 56, 10, RGB(51, 65, 85), False, msoAlignLeft
    Next i
End Sub

Public Sub AKSORA_YATIRIM_HAKEDIS_SEC()
    Dim ws As Worksheet, s As Shape, t As String, ev As Boolean, sc As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    Set s = ws.Shapes("Y_HK_DEPO")
    If s.ControlFormat.Value = 1 Then t = "" Else t = Trim$(Split(CStr(s.ControlFormat.List(s.ControlFormat.Value)) & "  -  ", "  -  ")(0))
    ev = application.EnableEvents: sc = application.ScreenUpdating
    application.EnableEvents = False: application.ScreenUpdating = False
    ws.Unprotect ""
    ws.Range("BX24").Value2 = t
    Y_PROJE_KARSILASTIR ws
    Y_HAKEDIS_ALT ws
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    application.EnableEvents = ev: application.ScreenUpdating = sc
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

' Tablo boyu degisince: tasarim karti tablonun altina alinir, yazdirma alani tabloyu kapsar.
Private Sub Y_HAKEDIS_ALT(ByVal ws As Worksheet)
    Dim s As Shape, w As Double, ust As Double, son As Long
    son = CLng(Val(CStr(ws.Range("BX28").Value2)))
    If son < 1 Then son = 14
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    ust = ws.Rows(son + 2).Top + 10
    Y_HAKEDIS_TASLAK ws, ust, w
    Y_FILIGRAN ws, ust + 200, w
    On Error Resume Next
    ws.PageSetup.PrintArea = "$A$1:$K$" & son
    On Error GoTo 0
End Sub

Public Sub AKSORA_YATIRIM_PROJE_YENILE()
    AKSORA_YATIRIM_HAKEDIS_SEC
End Sub

Public Sub AKSORA_YATIRIM_HAKEDIS_PDF()
    Dim ws As Worksheet, yol As String, kart As Boolean
    kart = AKSORA_BEKLE_AC("AKSORA  |  HAKEDÝÞ", "PDF hazýrlanýyor...", 0.3)
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
    ' Yazdirma alani her zaman guncel tablo boyuna gore (proje sonradan yuklendiyse kesik cikmaz).
    On Error Resume Next
    If Val(CStr(ws.Range("BX28").Value2)) > 0 Then ws.PageSetup.PrintArea = "$A$1:$K$" & CLng(Val(CStr(ws.Range("BX28").Value2)))
    On Error GoTo Hata
    yol = ThisWorkbook.Path & application.PathSeparator & "PROJE_KARSILASTIRMA_" & CStr(ws.Range("BX22").Value2) & "_" & _
        Format$(Now, "yyyymmdd_hhnn") & ".pdf"
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=yol, Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
    If kart Then AKSORA_BEKLE_KAPAT
    Exit Sub
Hata:
    If kart Then AKSORA_BEKLE_KAPAT
    MsgBox "PDF oluþturulamadý: " & Err.Description, vbExclamation, "AKSORA"
End Sub

' Calisma kitabi SheetChange: teminat girisleri aninda kayda yazilir.
Public Sub AKSORA_YATIRIM_DEGISIM(ByVal sh As Object, ByVal hedef As Range)
    Dim ev As Boolean, c As Range, il As String, a As Long, b As Long, alan As Range, t As String, anahtar As String
    If Not TypeOf sh Is Worksheet Then Exit Sub
    If sh.Name <> "Y_TEMÝNAT" Then Exit Sub
    If mYCalisiyor Then Exit Sub
    ev = application.EnableEvents
    On Error GoTo Cikis
    application.EnableEvents = False
    il = UCase$(Trim$(CStr(sh.Range("BX22").Value2)))
    If Len(il) = 0 Then GoTo Cikis
    ' Malzeme arama hucresi
    If Len(CStr(sh.Range("BX29").Value2)) > 0 Then
        If Not Intersect(hedef, sh.Range(CStr(sh.Range("BX29").Value2)).MergeArea) Is Nothing Then
            t = Trim$(CStr(sh.Range(CStr(sh.Range("BX29").Value2)).Value2))
            sh.Unprotect ""
            If Y_ARALIK(sh, "BX27", a, b) Then
                If Not sh.AutoFilterMode Then sh.Range(sh.Cells(a - 1, 1), sh.Cells(b, 14)).AutoFilter
                If Len(t) = 0 Then
                    sh.AutoFilter.Range.AutoFilter Field:=14
                Else
                    sh.AutoFilter.Range.AutoFilter Field:=14, Criteria1:="=*" & t & "*"
                End If
            End If
            sh.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
            Y_TEMINAT_KAL sh
            GoTo Cikis
        End If
    End If
    If Y_ARALIK(sh, "BX23", a, b) Then
        Set alan = Intersect(hedef, sh.Range(sh.Cells(a, 4), sh.Cells(b, 4)))
        If Not alan Is Nothing Then
            For Each c In alan.Cells
                If Len(CStr(sh.Cells(c.Row, 1).Value2)) > 0 Then Y_KAYIT_YAZ "_Y_TEMINAT", "T|" & il & "|" & UCase$(CStr(sh.Cells(c.Row, 1).Value2)), c.Value2
            Next c
        End If
    End If
    If Y_ARALIK(sh, "BX27", a, b) Then
        Set alan = Intersect(hedef, sh.Range(sh.Cells(a, 7), sh.Cells(b, 8)))
        If Not alan Is Nothing Then
            sh.Unprotect ""
            For Each c In alan.Cells
                If Len(CStr(sh.Cells(c.Row, 3).Value2)) > 0 Then
                    t = CStr(sh.Cells(c.Row, 7).Value2)
                    If t <> "Sahada" And t <> "Kýsmi" Then
                        t = "Depoda": sh.Cells(c.Row, 7).Value2 = t
                        sh.Cells(c.Row, 13).ClearContents
                    ElseIf Len(CStr(sh.Cells(c.Row, 13).Value2)) = 0 Then
                        sh.Cells(c.Row, 13).Value = Date
                    End If
                    anahtar = "K|" & il & "|" & UCase$(CStr(sh.Cells(c.Row, 1).Value2)) & "|" & CStr(sh.Cells(c.Row, 3).Value2)
                    Y_KAYIT_YAZ "_Y_TEMINAT", anahtar, t, sh.Cells(c.Row, 8).Value2, sh.Cells(c.Row, 13).Value2
                End If
            Next c
            sh.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    End If
Cikis:
    application.EnableEvents = ev
End Sub

' Giris ve guncelleme sonrasi: hata olsa bile birim acilir, hata durum seridine yazilir.
Private Sub Y_ALT_SAYFALAR(ByVal il As String)
    Dim da As Boolean, ad As Variant, ws As Worksheet, s As Shape, var As Boolean
    ' AKSORA amblemi tum yatirim sayfalarindan kaldirilir.
    On Error Resume Next
    For Each ad In Y_ADLAR()
        Set ws = ThisWorkbook.Worksheets(CStr(ad))
        var = False
        For Each s In ws.Shapes
            If InStr(s.Name, "FILIGRAN") > 0 Then var = True: Exit For
        Next s
        If var Then
            ws.Unprotect ""
            Y_FILIGRAN ws
            ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
        End If
    Next ad
    On Error GoTo 0
    da = application.DisplayAlerts
    application.DisplayAlerts = False
    On Error Resume Next
    Y_TEMINAT_KUR il
    If Err.Number <> 0 Then Y_SERIT "Teminat sayfasý kurulamadý: " & Err.Description, 3: Err.Clear
    ' Hakedis sayfasi yalniz il veya tasarim surumu degisince yeniden kurulur (girisler korunur).
    With ThisWorkbook.Worksheets("Y_HAKEDÝÞ")
        If CStr(.Range("BX22").Value2) <> UCase$(Trim$(il)) Or CStr(.Range("BX25").Value2) <> "HK_V2" Then
            Y_HAKEDIS_KUR il
        End If
    End With
    If Err.Number <> 0 Then Y_SERIT "Hakediþ sayfasý kurulamadý: " & Err.Description, 3: Err.Clear
    On Error GoTo 0
    application.DisplayAlerts = da
End Sub

'==========================================================
' ILCE YONETIM BIRIMI (I_ sayfalari)
' CDP (7151-7155) ve EPP (7051-7055) uretim yerlerinin K001-K999 ilce depolari.
' Ana panel, arama, stok ve rapor sayfalari musteri biriminin kilitli tasarimindan
' kopyalanir; karsilastirma ve ISG sayfalari ayni rapor basligiyla kurulur.
' Kayitlar: _IA_<IL> (son stok, paketli), _IH_<IL> (gunluk stok gecmisi, her sutun bir gun).
'==========================================================

Private Function I_ADLAR() As Variant
    I_ADLAR = Array("I_ANA PANEL", "I_MALZEME_ARA", "I_STOK_KARÞILAÞTIRMA", "I_AYLIK_DEMONTAJ", "I_CDP_ÝLÇELER_STOK", _
        "I_CDP_ÝLÇELER_RAPOR", "I_EPP_ÝLÇELER_STOK", "I_EPP_ÝLÇELER_RAPOR", "I_ÝSG_MALZEMELERÝ")
End Function

Private Function I_TUR(ByVal ad As String) As String
    Select Case ad
        Case "I_ANA PANEL": I_TUR = "PANEL"
        Case "I_MALZEME_ARA": I_TUR = "ARAMA"
        Case "I_CDP_ÝLÇELER_STOK", "I_EPP_ÝLÇELER_STOK": I_TUR = "STOK"
        Case "I_CDP_ÝLÇELER_RAPOR", "I_EPP_ÝLÇELER_RAPOR": I_TUR = "RAPOR"
        Case "I_STOK_KARÞILAÞTIRMA": I_TUR = "KARS"
        Case Else: I_TUR = "BOS"
    End Select
End Function

Public Function AKSORA_ILCE_SAYFA(ByVal sh As Object) As Boolean
    Dim ad As Variant
    If Not TypeOf sh Is Worksheet Then Exit Function
    If Left$(sh.Name, 2) <> "I_" Then Exit Function
    For Each ad In I_ADLAR()
        If sh.Name = CStr(ad) Then AKSORA_ILCE_SAYFA = (CStr(sh.Range("BX1").Value2) = I_ETIKET): Exit Function
    Next ad
End Function

Public Sub AKSORA_BIRIM_ILCE()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  ÝLÇE YÖNETÝM BÝRÝMÝ", "Birim açýlýyor...")
    On Error GoTo Son
    BIRIM_ILCE_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub BIRIM_ILCE_ASIL()
    On Error GoTo Hata
    YATIRIM_KAPAK_AC True
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Ýlçe Yönetim birimi", Err.Number, Err.Description
End Sub

Public Sub AKSORA_ILCE_GIRIS()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  ÝLÇE YÖNETÝM BÝRÝMÝ", "Ýl stoklarý ve ekranlar hazýrlanýyor...")
    On Error GoTo Son
    ILCE_GIRIS_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub ILCE_GIRIS_ASIL()
    Dim ws As Worksheet
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(ILCE_KAPAK)
    If Len(CStr(ws.Range("AZ1").Value2)) = 0 Then
        MsgBox "Lütfen bir il seçiniz.", vbInformation, "AKSORA"
    Else
        I_GIRIS CStr(ws.Range("AZ1").Value2)
    End If
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

' Kopyalanan sayfanin dugmeleri ilce birimine baglanir.
Private Sub I_TEMIZLE(ByVal ws As Worksheet)
    Select Case I_TUR(ws.Name)
        Case "PANEL": I_PANEL_TEMIZLE ws
        Case "ARAMA": Y_ARAMA_TEMIZLE ws
        Case "STOK": Y_STOK_TEMIZLE ws
        Case "RAPOR": Y_RAPOR_TEMIZLE ws
        Case "KARS": I_KARS_HAZIRLA ws
        Case Else: Y_SAYFA_SIFIRLA ws
    End Select
    If I_TUR(ws.Name) = "STOK" Then ws.Range("A7").Value2 = "Veri bekleniyor: SAP STOKLARINI GÜNCELLE ile ilçe depolarý (K001-K999) SAP'den alýnýr."
    If I_TUR(ws.Name) = "RAPOR" Then ws.Range("A9").Value2 = "Veri bekleniyor: ilçe stoklarý güncellenince rapor oluþacak."
    I_BAGLA ws
End Sub

Private Sub I_BAGLA(ByVal ws As Worksheet)
    Dim s As Shape
    Y_BAGLA Y_SEKIL(ws, "btnAnaPanelDon"), "I_ANA PANEL"
    Y_BAGLA Y_SEKIL(ws, "btnMalzemeAnaPanel"), "I_ANA PANEL"
    Set s = Y_SEKIL(ws, "SAP_STOK_GUNCELLE_BTN")
    If Not s Is Nothing Then s.OnAction = BIRIM_MAKRO("AKSORA_ILCE_GUNCELLE")
    If ws.Name = "I_MALZEME_ARA" Then
        Set s = Y_SEKIL(ws, "AKSORA_TOPLU_GUNCELLE")
        If Not s Is Nothing Then
            Y_YAZ s, "BÜTÜN ÝLLERÝN ÝLÇELERÝNÝ GÜNCELLE" & vbLf & "MB52  |  5 ÝL"
            s.OnAction = BIRIM_MAKRO("AKSORA_ILCE_BUTUN_ILLER")
        End If
        Set s = Y_SEKIL(ws, "AKSORA_TOPLU_VURGU")
        If Not s Is Nothing Then s.OnAction = BIRIM_MAKRO("AKSORA_ILCE_BUTUN_ILLER")
        Y_YAZ Y_SEKIL(ws, "AKSORA_TOPLU_DURUM"), "Arama, her ilin yukarýdaki tarihli ilçe depo stoklarýný kullanýr."
        Y_YAZ Y_SEKIL(ws, "AKSORA_ARAMA_SABIT_ACIKLAMA"), "Yalnýzca ilçe depolarý (CDP ve EPP, K001-K999) aranýr."
    End If
End Sub

Private Sub I_PANEL_TEMIZLE(ByVal ws As Worksheet)
    Dim s As Shape, ad As String, b As Variant
    For Each s In ws.Shapes
        ad = s.Name
        If Left$(ad, 20) = "AKSORA15_KART_DEGER_" Then
            If Right$(ad, 1) = "0" Then Y_YAZ s, "Veri bekleniyor" Else Y_YAZ s, "-"
        End If
        If Left$(ad, 18) = "AKSORA15_KART_LED_" Then s.Fill.ForeColor.RGB = RGB(145, 162, 176)
    Next s
    Y_YAZ Y_SEKIL(ws, "btnBaglantiKontrol"), "STOK / DEMONTAJ KONTROLÜ": Y_BAGLA Y_SEKIL(ws, "btnBaglantiKontrol"), "I_STOK_KARÞILAÞTIRMA"
    Y_YAZ Y_SEKIL(ws, "btnMalzemeAraPanel"), "MALZEME ARAMA": Y_BAGLA Y_SEKIL(ws, "btnMalzemeAraPanel"), "I_MALZEME_ARA"
    Y_YAZ Y_SEKIL(ws, "btnTumStok"), "TÜM ÝLÇE STOKLARINI GÜNCELLE"
    Y_YAZ Y_SEKIL(ws, "btnCDP"), "CDP ÝLÇE GÜNCELLE"
    Y_YAZ Y_SEKIL(ws, "btnIlce"), "EPP ÝLÇE GÜNCELLE"
    For Each b In Array("btnTumStok", "btnCDP", "btnIlce")
        Set s = Y_SEKIL(ws, CStr(b))
        If Not s Is Nothing Then s.OnAction = BIRIM_MAKRO("AKSORA_ILCE_GUNCELLE")
    Next b
    Y_YAZ Y_SEKIL(ws, "btnYuk"), "ÝSG ZÝMMET KONTROLÜ": Y_BAGLA Y_SEKIL(ws, "btnYuk"), "I_ÝSG_MALZEMELERÝ"
    Set s = Y_SEKIL(ws, "AKSORA15_IL_DEGISTIR")
    If Not s Is Nothing Then s.OnAction = BIRIM_MAKRO("AKSORA_BIRIM_ILCE")
    I_HIZLI ws
End Sub

' Buyuk dugmeler: ust sira islem sayfalari ayni beyaz temada (STOK / DEMONTAJ | MALZEME ARAMA |
' ISG ZIMMET), alt sira guncellemeler (CDP | TUM ILCE - yesil, farkli | EPP). Bir kez uygulanir (BX27).
Private Sub I_PANEL_DUGME_DUZEN(ByVal ws As Worksheet)
    Dim l1 As Shape, m1 As Shape, r1 As Shape, l2 As Shape, m2 As Shape, r2 As Shape, g As Variant
    Set l1 = Y_SEKIL(ws, "btnBaglantiKontrol"): Set m1 = Y_SEKIL(ws, "btnTumStok"): Set r1 = Y_SEKIL(ws, "btnMalzemeAraPanel")
    Set l2 = Y_SEKIL(ws, "btnCDP"): Set m2 = Y_SEKIL(ws, "btnIlce"): Set r2 = Y_SEKIL(ws, "btnYuk")
    If l1 Is Nothing Or m1 Is Nothing Or r1 Is Nothing Or l2 Is Nothing Or m2 Is Nothing Or r2 Is Nothing Then Exit Sub
    g = Array(m1.Left, m1.Top, m1.Width, m1.Height, r1.Left, r1.Top, r1.Width, r1.Height, _
        m2.Left, m2.Top, m2.Width, m2.Height, r2.Left, r2.Top, r2.Width, r2.Height)
    On Error Resume Next
    r1.Left = g(0): r1.Top = g(1): r1.Width = g(2): r1.Height = g(3)
    l1.PickUp: r2.Apply
    r2.Left = g(4): r2.Top = g(5): r2.Width = g(6): r2.Height = g(7)
    With r2.TextFrame2.TextRange.Font
        .Size = l1.TextFrame2.TextRange.Font.Size
        .bold = l1.TextFrame2.TextRange.Font.bold
        .Fill.ForeColor.RGB = l1.TextFrame2.TextRange.Font.Fill.ForeColor.RGB
    End With
    m1.Left = g(8): m1.Top = g(9): m1.Width = g(10): m1.Height = g(11)
    m1.TextFrame2.TextRange.Font.Size = l2.TextFrame2.TextRange.Font.Size
    m2.Left = g(12): m2.Top = g(13): m2.Width = g(14): m2.Height = g(15)
    On Error GoTo 0
    ws.Range("BX27").Value2 = "DUGME_V1"
End Sub

' Hizli erisim: 7 sayfa ayni satirda esit aralikla (eksik dugme ilk dugmeden cogaltilir).
Private Sub I_HIZLI(ByVal ws As Worksheet)
    Dim s As Shape, ilk As Shape, i As Long, x0 As Double, x1 As Double, bw As Double, g As Double
    Dim adlar As Variant, hedefler As Variant, yeni As Shape
    Set ilk = Y_SEKIL(ws, "MUSTERI_HIZLI_0")
    If ilk Is Nothing Then Exit Sub
    x0 = ilk.Left: x1 = CDbl(Val(CStr(ws.Range("BX26").Value2))): g = 6
    If x1 <= x0 Then
        For Each s In ws.Shapes
            If Left$(s.Name, 14) = "MUSTERI_HIZLI_" Then x1 = application.Max(x1, s.Left + s.Width)
        Next s
        ws.Range("BX26").Value2 = x1
    End If
    For i = 1 To 3
        If Y_SEKIL(ws, "MUSTERI_HIZLI_" & i) Is Nothing Then
            Set yeni = ilk.Duplicate
            yeni.Name = "MUSTERI_HIZLI_" & i
        End If
    Next i
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 14) = "MUSTERI_HIZLI_" Then
            If Val(Mid$(ws.Shapes(i).Name, 15)) > 3 Then ws.Shapes(i).Delete
        End If
    Next i
    bw = (x1 - x0 - 3 * g) / 4
    ' Buyuk dugmesi olan sayfalar (arama, stok/demontaj, zimmet) burada tekrarlanmaz.
    adlar = Array("CDP ÝLÇE STOK", "CDP ÝLÇE RAPOR", "EPP ÝLÇE STOK", "EPP ÝLÇE RAPOR")
    hedefler = Array("I_CDP_ÝLÇELER_STOK", "I_CDP_ÝLÇELER_RAPOR", "I_EPP_ÝLÇELER_STOK", "I_EPP_ÝLÇELER_RAPOR")
    Y_YAZ Y_SEKIL(ws, "btnBaglantiKontrol"), "STOK / DEMONTAJ KONTROLÜ"
    Y_YAZ Y_SEKIL(ws, "btnYuk"), "ÝSG ZÝMMET KONTROLÜ"
    Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_BASLIK_Q"), "ÝSG ZÝMMET KONTROLÜ"
    For i = 0 To 3
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
    If CStr(ws.Range("BX27").Value2) <> "DUGME_V1" Then I_PANEL_DUGME_DUZEN ws
    ws.Range("BX25").Value2 = "IP_V4"
End Sub

' Secili ile gore basliklar, kart basliklari, alt bilgi ve panel bandi.
Private Sub I_IL_UYGULA(ByVal kod As String)
    Dim ad As Variant, ws As Worksheet, il As String, s As Shape
    il = Y_IL_ADI(kod)
    For Each ad In I_ADLAR()
        Set ws = ThisWorkbook.Worksheets(CStr(ad))
        If CStr(ws.Range("BX20").Value2) <> kod Then
            ws.Unprotect ""
            Select Case I_TUR(ws.Name)
                Case "PANEL"
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_BASLIK_B"), "CDP ÝLÇELER / " & AKSORA_IL_TESIS(kod, "CDP")
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_BASLIK_L"), "EPP ÝLÇELER / " & AKSORA_IL_TESIS(kod, "EPP")
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_BASLIK_Q"), "ÝSG ZÝMMET KONTROLÜ"
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_ALT_KURUM"), "ÇORUH ELEKTRÝK DAÐITIM A.Þ. | " & il & " Ýl Müdürlüðü"
                    Y_YAZ Y_SEKIL(ws, "AKSORA15_PANEL_TARIHI"), "Ýlçe Yönetim Birimi  |  Panel kontrolü: " & Format$(Now, "dd.mm.yyyy hh:nn")
                    Set s = Y_SEKIL(ws, "MUSTERI_ANA_SECIM_SERIDI")
                    If Not s Is Nothing Then
                        s.Fill.ForeColor.RGB = RGB(224, 242, 254)
                        Y_YAZ s, il & " Ýlçe Yönetim Birimi"
                        s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(12, 45, 79)
                    End If
                    Y_BANNER ws, kod, "AKSORA_ILCE_PANEL_BANTLARI.png"
                Case "STOK"
                    Y_YAZ Y_SEKIL(ws, "AKT1_STOK_BASLIK"), Y_BASLIK(ws.Name) & "  |  " & il & " Ýl Müdürlüðü"
                Case "ARAMA"
                Case Else
                    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | " & Y_BASLIK(ws.Name) & "  |  " & il
                    Y_YAZ Y_SEKIL(ws, "AKSORA_RP_ALT"), "AKSORA  |  " & Y_BASLIK(ws.Name) & "  |  " & il & " Ýl Müdürlüðü"
            End Select
            ws.Range("BX20").Value2 = kod
            Y_KORU ws
        End If
    Next ad
End Sub

' Il secim ekranindan giris: sayfalar kurulur, secili ile uyarlanir, yalniz I_ sayfalari gorunur.
Private Sub I_GIRIS(ByVal kod As String)
    Dim ev As Boolean, ad As Variant, k As Worksheet, panel As Worksheet, no As Long, metin As String
    ev = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    mYonlendirme = True
    kod = UCase$(Trim$(kod))
    I_ESKI_KARS_SIL
    Y_KUR True
    I_IL_UYGULA kod
    I_IL_VERI kod
    I_ALT_SAYFALAR kod, True
    I_PANEL_YENILE
    Y_ARAMA_TARIH True
    For Each ad In I_ADLAR()
        ThisWorkbook.Worksheets(CStr(ad)).Visible = xlSheetVisible
    Next ad
    Y_SEKME_SIRASI True
    Set panel = ThisWorkbook.Worksheets("I_ANA PANEL")
    panel.Select Replace:=True
    For Each k In ThisWorkbook.Worksheets
        If Not AKSORA_ILCE_SAYFA(k) Then
            If k.Visible = xlSheetVisible Then k.Visible = xlSheetVeryHidden
        End If
    Next k
    AKSORA_YATIRIM_GORUNUM panel
Cikis:
    mYonlendirme = False
    application.EnableEvents = ev: application.ScreenUpdating = True
    If no <> 0 Then MsgBox "Ýlçe Yönetim birimi açýlamadý." & vbCrLf & metin & vbCrLf & "(Hata " & no & ")", vbExclamation, "AKSORA"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    application.DisplayAlerts = True
    Resume Cikis
End Sub

' Ilk surumdeki (rapor kopyasi) karsilastirma sayfasi kaldirilir; yerine malzeme biriminin kopyasi kurulur.
Private Sub I_ESKI_KARS_SIL()
    Dim ws As Worksheet, da As Boolean
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("I_STOK_KARÞILAÞTIRMA")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If Left$(CStr(ws.Range("BX25").Value2), 3) = "IK_" Then Exit Sub
    da = application.DisplayAlerts: application.DisplayAlerts = False
    ws.Visible = xlSheetVisible
    ws.Delete
    application.DisplayAlerts = da
End Sub

' Karsilastirma ve ISG sayfalari: hata olsa bile birim acilir, hata seride yazilir.
Private Sub I_ALT_SAYFALAR(ByVal il As String, ByVal sifirla As Boolean)
    Dim da As Boolean, ad As Variant, ws As Worksheet
    da = application.DisplayAlerts
    application.DisplayAlerts = False
    On Error Resume Next
    I_KARS_KUR il, sifirla
    If Err.Number <> 0 Then Y_SERIT "Stok karþýlaþtýrma sayfasý kurulamadý: " & Err.Description, 3, "I_ANA PANEL": Err.Clear
    I_ISG_KUR il
    If Err.Number <> 0 Then Y_SERIT "ÝSG sayfasý kurulamadý: " & Err.Description, 3, "I_ANA PANEL": Err.Clear
    I_AYLIK_KUR il
    If Err.Number <> 0 Then Y_SERIT "Aylýk demontaj sayfasý kurulamadý: " & Err.Description, 3, "I_ANA PANEL": Err.Clear
    ' AKSORA amblemi ilce sayfalarinda kullanilmaz (listelerin ustune biniyordu).
    For Each ad In I_ADLAR()
        Set ws = ThisWorkbook.Worksheets(CStr(ad))
        ws.Unprotect ""
        Y_FILIGRAN ws
        Y_KORU ws
    Next ad
    On Error GoTo 0
    application.DisplayAlerts = da
End Sub

'---------- Ilce stoklari (SAP MB52) ----------
Private Function I_DEPO_UYGUN(ByVal depo As String) As Boolean
    depo = UCase$(Trim$(depo))
    If Len(depo) <> 4 Or Left$(depo, 1) <> "K" Then Exit Function
    If Not (Mid$(depo, 2, 3) Like "###") Then Exit Function
    I_DEPO_UYGUN = (Val(Mid$(depo, 2, 3)) >= 1)
End Function

Private Function I_AKTIF_IL() As String
    On Error Resume Next
    I_AKTIF_IL = UCase$(Trim$(CStr(ThisWorkbook.Worksheets("I_ANA PANEL").Range("BX20").Value2)))
    On Error GoTo 0
    If Len(I_AKTIF_IL) = 0 Then I_AKTIF_IL = "ARTVIN"
End Function

Private Function I_ARSIV_SAYFA(ByVal il As String, ByVal olustur As Boolean) As Worksheet
    Dim ad As String
    ad = "_IA_" & UCase$(Trim$(il))
    If olustur Then Set I_ARSIV_SAYFA = Y_KAYIT_SAYFA(ad): Exit Function
    On Error Resume Next
    Set I_ARSIV_SAYFA = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
End Function

Public Function AKSORA_ILCE_ARSIV_SAYFA(ByVal il As String) As Worksheet
    Set AKSORA_ILCE_ARSIV_SAYFA = I_ARSIV_SAYFA(il, False)
End Function

' Ilin son ilce guncellemesi (kayit yoksa 0).
Public Function AKSORA_ILCE_IL_SON_TARIH(ByVal il As String) As Date
    Dim ar As Worksheet, v As Variant
    Set ar = I_ARSIV_SAYFA(il, False)
    If ar Is Nothing Then Exit Function
    v = ar.Range("A1").Value
    If Not IsError(v) Then
        If IsDate(v) Then AKSORA_ILCE_IL_SON_TARIH = CDate(v)
    End If
End Function

Private Function I_SAYI(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) = vbString Then
        If IsNumeric(v) Then I_SAYI = CDbl(v)
    ElseIf IsNumeric(v) Then
        I_SAYI = CDbl(v)
    End If
End Function

' Kayittaki bir uretim yerinin satirlari (tek sistem guncellenince digeri korunur).
Private Function I_SISTEM_SATIRLARI(ByVal v As Variant, ByVal tesis As String) As Variant
    Dim r As Long, c As Long, n As Long, s() As Variant
    If Not IsArray(v) Then Exit Function
    For r = 1 To UBound(v, 1)
        If Trim$(CStr(v(r, 1))) = tesis Then n = n + 1
    Next r
    If n = 0 Then Exit Function
    ReDim s(1 To n, 1 To 14)
    n = 0
    For r = 1 To UBound(v, 1)
        If Trim$(CStr(v(r, 1))) = tesis Then
            n = n + 1
            For c = 1 To application.Min(14, UBound(v, 2)): s(n, c) = v(r, c): Next c
        End If
    Next r
    I_SISTEM_SATIRLARI = s
End Function

' CDP ve EPP satirlari tek listede: depo, uretim yeri, malzeme sirasiyla.
Private Function I_BIRLESTIR(ByVal a As Variant, ByVal b As Variant) As Variant
    Dim na As Long, nb As Long, n As Long, k() As String, idx() As Long, r As Long, c As Long, j As Long, s() As Variant
    If IsArray(a) Then na = UBound(a, 1)
    If IsArray(b) Then nb = UBound(b, 1)
    n = na + nb
    If n = 0 Then Exit Function
    ReDim k(1 To n): ReDim idx(1 To n)
    For r = 1 To n
        idx(r) = r
        If r <= na Then
            k(r) = CStr(a(r, 2)) & "|" & CStr(a(r, 1)) & "|" & CStr(a(r, 5)) & "|" & CStr(a(r, 9))
        Else
            k(r) = CStr(b(r - na, 2)) & "|" & CStr(b(r - na, 1)) & "|" & CStr(b(r - na, 5)) & "|" & CStr(b(r - na, 9))
        End If
    Next r
    Y_SIRALA k, idx, 1, n
    ReDim s(1 To n, 1 To 14)
    For r = 1 To n
        j = idx(r)
        For c = 1 To 14
            If j <= na Then s(r, c) = a(j, c) Else s(r, c) = b(j - na, c)
        Next c
    Next r
    I_BIRLESTIR = s
End Function

' Il degisince stok sayfasi o ilin kaydindan yuklenir (kayit yoksa bos).
Private Sub I_IL_VERI(ByVal kod As String)
    Dim ar As Worksheet, a As Variant, tarih As Variant
    If CStr(ThisWorkbook.Worksheets("I_CDP_ÝLÇELER_STOK").Range("BX19").Value2) = kod And _
        CStr(ThisWorkbook.Worksheets("I_EPP_ÝLÇELER_STOK").Range("BX19").Value2) = kod Then Exit Sub
    Set ar = I_ARSIV_SAYFA(kod, False)
    If Not ar Is Nothing Then
        tarih = ar.Range("A1").Value
        a = AKSORA_PAKET_OKU(ar)
        I_DEPO_KAYDET kod, a
    End If
    I_STOK_YAZ kod, a, tarih
End Sub

' CDP ve EPP ilce stoklari ayri sayfalara (kayit tek: _IA_<IL>).
Private Sub I_STOK_YAZ(ByVal il As String, ByVal a As Variant, ByVal tarih As Variant)
    Y_STOK_YAZ ThisWorkbook.Worksheets("I_CDP_ÝLÇELER_STOK"), I_SISTEM_SATIRLARI(a, AKSORA_IL_TESIS(il, "CDP")), il, tarih
    Y_STOK_YAZ ThisWorkbook.Worksheets("I_EPP_ÝLÇELER_STOK"), I_SISTEM_SATIRLARI(a, AKSORA_IL_TESIS(il, "EPP")), il, tarih
End Sub

' Ilin ilce depolari kalici listede tutulur (stogu olmayan depo da sayilir).
Private Sub I_DEPO_KAYDET(ByVal il As String, ByVal a As Variant)
    Dim d As Object, r As Long, k As String, kayit As Object, x As Variant
    If Not IsArray(a) Then Exit Sub
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For r = 1 To UBound(a, 1)
        k = UCase$(I_TEMIZ(a(r, 2)))
        If Len(k) > 0 Then
            If Not d.Exists(k) Then d(k) = I_TEMIZ(a(r, 3))
            If Len(CStr(d(k))) = 0 Then d(k) = I_TEMIZ(a(r, 3))
        End If
    Next r
    Set kayit = Y_KAYIT_SOZLUK("_I_DEPO")
    For Each x In d.keys
        If Not kayit.Exists(UCase$(il) & "|" & CStr(x)) Then
            Y_KAYIT_YAZ "_I_DEPO", UCase$(il) & "|" & CStr(x), CStr(d(x))
        ElseIf Len(CStr(d(x))) > 0 And CStr(kayit(UCase$(il) & "|" & CStr(x))(0)) <> CStr(d(x)) Then
            Y_KAYIT_YAZ "_I_DEPO", UCase$(il) & "|" & CStr(x), CStr(d(x))
        End If
    Next x
End Sub

' Ilin bilinen ilce depolari: depo -> ad.
Private Function I_DEPOLAR(ByVal il As String) As Object
    Dim d As Object, kayit As Object, x As Variant, onek As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set kayit = Y_KAYIT_SOZLUK("_I_DEPO")
    onek = UCase$(Trim$(il)) & "|"
    For Each x In kayit.keys
        If Left$(CStr(x), Len(onek)) = onek Then d(Mid$(CStr(x), Len(onek) + 1)) = CStr(kayit(x)(0))
    Next x
    Set I_DEPOLAR = d
End Function

Private Sub I_DURUM(ByVal metin As String)
    application.StatusBar = "AKSORA | Ýlçe Yönetim | " & metin
    AKSORA_ADMIN_DURUM metin
    If mYToplu Then Y_TOPLU_DURUM metin, "I_MALZEME_ARA"
    Y_SERIT "Güncelleniyor  •  " & mYSira & metin, 1, "I_ANA PANEL"
    DoEvents
End Sub

' Dugmeye gore kapsam: C = CDP ilceler, E = EPP ilceler.
Public Sub AKSORA_ILCE_GUNCELLE()
    Dim cagiran As String, kapsam As String, bilgi As String
    If AKSORA_V11_MESGUL Or mYCalisiyor Then Exit Sub
    On Error Resume Next
    cagiran = CStr(application.Caller)
    On Error GoTo 0
    Select Case cagiran
        Case "btnCDP": kapsam = "C"
        Case "btnIlce": kapsam = "E"
        Case Else: kapsam = "CE"
    End Select
    On Error Resume Next
    If Left$(ActiveSheet.Name, 6) = "I_CDP_" Then kapsam = "C"
    If Left$(ActiveSheet.Name, 6) = "I_EPP_" Then kapsam = "E"
    On Error GoTo 0
    I_GUNCELLE I_AKTIF_IL(), kapsam, False, bilgi
End Sub

Private Function I_GUNCELLE(ByVal il As String, ByVal kapsam As String, ByVal sessiz As Boolean, ByRef bilgi As String) As Boolean
    Dim vC As Variant, vE As Variant, aC As Variant, aE As Variant, eski As Variant, a As Variant, mesaj As String
    Dim tarih As Date, bas As Double, sure As Double, ev As Boolean, sc As Boolean, hesap As XlCalculation
    Dim ekran As Boolean, mesgulOnce As Boolean, hMesaj As String, ar As Worksheet, n As Long, ws As Worksheet
    Dim nC As Long, nE As Long
    If mYCalisiyor Then Exit Function
    mYCalisiyor = True: bas = Timer
    mesgulOnce = AKSORA_V11_MESGUL
    ev = application.EnableEvents: sc = application.ScreenUpdating: hesap = application.Calculation
    On Error GoTo Hata
    il = UCase$(Trim$(il))
    Set ws = ThisWorkbook.Worksheets("I_CDP_ÝLÇELER_STOK")
    ekran = (CStr(ws.Range("BX19").Value2) = il)
    Y_CANLI_BASLAT "I_ANA PANEL", Y_IL_ADI(il) & "  •  SAP baðlantýsý hazýrlanýyor..."
    AKSORA_V11_MESGUL = True
    If InStr(kapsam, "C") > 0 Then
        I_DURUM Y_IL_ADI(il) & " | CDP " & AKSORA_IL_TESIS(il, "CDP") & " ilçe depolarý SAP'den alýnýyor..."
        vC = AKSORA_YATIRIM_MB52("CDP", il, mesaj)
        If Not IsArray(vC) Then Err.Raise vbObjectError + 2111, , "CDP " & AKSORA_IL_TESIS(il, "CDP") & ": " & mesaj
        aC = Y_DONUSTUR(vC, il, 5)
    End If
    If InStr(kapsam, "E") > 0 Then
        I_DURUM Y_IL_ADI(il) & " | EPP " & AKSORA_IL_TESIS(il, "EPP") & " ilçe depolarý SAP'den alýnýyor..."
        mesaj = vbNullString
        vE = AKSORA_YATIRIM_MB52("EPP", il, mesaj)
        If Not IsArray(vE) Then Err.Raise vbObjectError + 2112, , "EPP " & AKSORA_IL_TESIS(il, "EPP") & ": " & mesaj
        aE = Y_DONUSTUR(vE, il, 6)
    End If
    ' SAP'den ilce deposu satiri gelmediyse (rapor duzeni / varyant degisimi) ve onceki kayitta
    ' o tesisin satirlari varsa bos sonuc kaydedilmez: aksi halde tum stok "kullanildi" sayilir.
    If (InStr(kapsam, "C") > 0 And Not IsArray(aC)) Or (InStr(kapsam, "E") > 0 And Not IsArray(aE)) Then
        Set ar = I_ARSIV_SAYFA(il, False)
        If Not ar Is Nothing Then eski = AKSORA_PAKET_OKU(ar)
        If InStr(kapsam, "C") > 0 And Not IsArray(aC) Then
            If IsArray(I_SISTEM_SATIRLARI(eski, AKSORA_IL_TESIS(il, "CDP"))) Then Err.Raise vbObjectError + 2113, , _
                "CDP " & AKSORA_IL_TESIS(il, "CDP") & ": SAP'den ilçe deposu satýrý gelmedi; önceki kayýt korundu."
        End If
        If InStr(kapsam, "E") > 0 And Not IsArray(aE) Then
            If IsArray(I_SISTEM_SATIRLARI(eski, AKSORA_IL_TESIS(il, "EPP"))) Then Err.Raise vbObjectError + 2114, , _
                "EPP " & AKSORA_IL_TESIS(il, "EPP") & ": SAP'den ilçe deposu satýrý gelmedi; önceki kayýt korundu."
        End If
        eski = Empty: Set ar = Nothing
    End If
    I_DURUM Y_IL_ADI(il) & " | Sayfalar yazýlýyor..."
    AKSORA_V11_MESGUL = mesgulOnce
    application.EnableEvents = False: application.Calculation = xlCalculationManual
    application.ScreenUpdating = False
    ' Yalniz bir sistem guncellenince digerinin son kaydi korunur.
    If InStr(kapsam, "C") = 0 Or InStr(kapsam, "E") = 0 Then
        Set ar = I_ARSIV_SAYFA(il, False)
        If Not ar Is Nothing Then eski = AKSORA_PAKET_OKU(ar)
        If InStr(kapsam, "C") = 0 Then aC = I_SISTEM_SATIRLARI(eski, AKSORA_IL_TESIS(il, "CDP"))
        If InStr(kapsam, "E") = 0 Then aE = I_SISTEM_SATIRLARI(eski, AKSORA_IL_TESIS(il, "EPP"))
    End If
    If IsArray(aC) Then nC = UBound(aC, 1)
    If IsArray(aE) Then nE = UBound(aE, 1)
    a = I_BIRLESTIR(aC, aE)
    If IsArray(a) Then n = UBound(a, 1)
    tarih = Now
    Set ar = I_ARSIV_SAYFA(il, True)
    ar.Range("A1").Value = tarih: ar.Range("B1").Value2 = il
    AKSORA_PAKET_YAZ ar, a
    I_GECMIS_EKLE il, a, tarih
    I_MAL_DIZI a
    I_DEFTER_GUNCELLE il
    I_DEPO_KAYDET il, a
    sure = Timer - bas: If sure < 0 Then sure = sure + 86400#
    ThisWorkbook.names.Add Name:="_AKSORA_ILCE_SURE_" & il, RefersTo:="=" & Trim$(Str$(Round(sure, 2))), Visible:=False
    If ekran Then
        I_STOK_YAZ il, a, tarih
        I_ALT_SAYFALAR il, True
        I_PANEL_YENILE
        Y_ARAMA_TARIH True
    End If
    application.Calculation = hesap: application.EnableEvents = ev
    application.ScreenUpdating = sc
    bilgi = Y_IL_ADI(il) & ": güncellendi (" & AKSORA_SAYI_METNI(sure, "0.0") & " sn)"
    I_GUNCELLE = True
    If Not mYToplu Then Y_SERIT "Güncelleme tamamlandý  |  " & I_SERIT_METNI(), 2, "I_ANA PANEL"
    If Not mYToplu And Not sessiz Then Y_PANEL_GOSTER "I_ANA PANEL"
    If Not sessiz Then MsgBox "Ýlçe Yönetim birimi | " & Y_IL_ADI(il) & " güncellendi." & vbCrLf & vbCrLf & _
        "CDP " & AKSORA_IL_TESIS(il, "CDP") & " ilçe depolarý: " & AKSORA_SAYI_METNI(nC, "#,##0") & " satýr" & vbCrLf & _
        "EPP " & AKSORA_IL_TESIS(il, "EPP") & " ilçe depolarý: " & AKSORA_SAYI_METNI(nE, "#,##0") & " satýr" & vbCrLf & _
        "Depolar: K001 - K999" & vbCrLf & _
        "Süre: " & AKSORA_SAYI_METNI(sure, "0.00") & " sn", vbInformation, "AKSORA | Ýlçe Yönetim Birimi"
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
    If Not mYToplu Then Y_SERIT "Güncelleme tamamlanamadý  |  " & hMesaj, 3, "I_ANA PANEL"
    If Not sessiz Then MsgBox "Ýlçe stoklarý güncellenemedi." & vbCrLf & hMesaj, vbExclamation, "AKSORA | Ýlçe Yönetim Birimi"
    On Error GoTo 0
    Resume Cikis
End Function

' Bes ilin ilce depolari sirayla alinir; ekrandaki il sayfalara, digerleri kendi kaydina yazilir.
Public Sub AKSORA_ILCE_BUTUN_ILLER()
    Dim il As Variant, bilgi As String, sonuc As String, adet As Long, bas As Double, sure As Double
    Dim hataMetni As String, metin As String, sc As Boolean, baslangic As Worksheet, adim As Long
    If AKSORA_V11_MESGUL Or mYCalisiyor Or mYToplu Then Exit Sub
    If Not AKSORA_ADMIN_TOPLU Then
        If MsgBox("Beþ ilin (Artvin, Trabzon, Gümüþhane, Giresun, Rize) CDP ve EPP ilçe depo stoklarý SAP'den güncellenecek. Devam edilsin mi?", _
            vbYesNo + vbQuestion, "AKSORA | Ýlçe Yönetim Birimi") <> vbYes Then Exit Sub
    End If
    sc = application.ScreenUpdating
    If TypeOf ActiveSheet Is Worksheet Then Set baslangic = ActiveSheet
    mYToplu = True: bas = Timer
    application.ScreenUpdating = True
    On Error GoTo Hata
    For Each il In Array("ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
        If AKSORA_ADMIN_IPTAL Then sonuc = sonuc & "Durduruldu: kalan iller atlandý." & vbCrLf: Exit For
        bilgi = vbNullString
        adim = adim + 1: mYSira = "Bütün iller " & adim & "/5  •  "
        Y_TOPLU_DURUM Y_IL_ADI(CStr(il)) & " güncelleniyor...", "I_MALZEME_ARA"
        If I_GUNCELLE(CStr(il), "CE", True, bilgi) Then adet = adet + 1
        sonuc = sonuc & bilgi & vbCrLf
    Next il
Bitir:
    On Error Resume Next
    mYToplu = False: mYSira = vbNullString
    sure = Timer - bas: If sure < 0 Then sure = sure + 86400#
    ThisWorkbook.names.Add Name:="_AKSORA_ILCE_SON_SURE", RefersTo:="=" & Trim$(Str$(Round(sure, 2))), Visible:=False
    metin = IIf(Len(hataMetni) > 0, "Ýþlem durduruldu. ", "Ýþlem tamamlandý. ") & adet & "/5 il güncellendi." & vbLf & _
        "Geçen süre: " & Format$(CLng(Fix(sure)) \ 60, "00") & ":" & Format$(CLng(Fix(sure)) Mod 60, "00")
    Y_TOPLU_DURUM metin, "I_MALZEME_ARA"
    I_PANEL_YENILE
    Y_ARAMA_TARIH True
    application.StatusBar = False
    Y_PANEL_GOSTER "I_ANA PANEL"
    application.ScreenUpdating = sc
    On Error GoTo 0
    If AKSORA_ADMIN_TOPLU Then AKSORA_ADMIN_SONUC = AKSORA_ADMIN_SONUC & "ÝLÇE:" & vbCrLf & sonuc & IIf(Len(hataMetni) > 0, hataMetni & vbCrLf, ""): Exit Sub
    MsgBox "Ýlçe Yönetim birimi | Bütün iller" & vbCrLf & vbCrLf & sonuc & IIf(Len(hataMetni) > 0, vbCrLf & hataMetni & vbCrLf, "") & vbCrLf & _
        "Toplam süre: " & AKSORA_SURE_METNI(sure), IIf(adet = 5, vbInformation, vbExclamation), "AKSORA | Ýlçe Yönetim Birimi"
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Bitir
End Sub

'---------- Ana panel ----------
Private Sub I_OZET(ByVal sh As Worksheet, ByVal tesis As String, ByVal d As Object, ByRef satir As Long, ByRef deger As Double, ByRef tarih As Date)
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
        If Not IsError(v(r, 5)) And Not IsError(v(r, 1)) Then
            If Len(Trim$(CStr(v(r, 5)))) > 0 And (Len(tesis) = 0 Or Trim$(CStr(v(r, 1))) = tesis) Then
                satir = satir + 1
                If Not IsError(v(r, 2)) Then d(Trim$(CStr(v(r, 2)))) = True
                deger = deger + I_SAYI(v(r, 11))
            End If
        End If
    Next r
End Sub

Private Function I_SERIT_METNI() As String
    Dim sh As Worksheet, d As Object, satir As Long, deger As Double, tarih As Date, metin As String, i As Long
    Dim toplam As Double, il As String, t As String, sure As Double, sistem As Variant
    il = I_AKTIF_IL()
    For Each sistem In Array("CDP", "EPP")
        Set sh = ThisWorkbook.Worksheets("I_" & CStr(sistem) & "_ÝLÇELER_STOK")
        Set d = CreateObject("Scripting.Dictionary")
        satir = 0: deger = 0
        I_OZET sh, AKSORA_IL_TESIS(il, CStr(sistem)), d, satir, deger, tarih
        toplam = toplam + deger
        metin = metin & "  |  " & CStr(sistem) & " Ýlçe " & AKSORA_SAYI_METNI(satir, "#,##0") & " satýr"
    Next sistem
    If tarih = 0 Then
        I_SERIT_METNI = "Henüz SAP güncellemesi yapýlmadý  |  CDP " & AKSORA_IL_TESIS(il, "CDP") & " ve EPP " & _
            AKSORA_IL_TESIS(il, "EPP") & " ilçe depolarý: K001 - K999"
        Exit Function
    End If
    If AKSORA_TOPLU_SON_SURE("_AKSORA_ILCE_SON_SURE") > 0 Then _
        metin = "  |  Bütün iller " & AKSORA_SURE_METNI(AKSORA_TOPLU_SON_SURE("_AKSORA_ILCE_SON_SURE")) & metin
    On Error Resume Next
    t = ThisWorkbook.names("_AKSORA_ILCE_SURE_" & il).RefersTo
    On Error GoTo 0
    If Len(t) > 1 Then sure = Val(Mid$(t, 2))
    If sure > 0 Then metin = "  |  Ýl güncelleme süresi " & AKSORA_SURE_METNI(sure) & metin
    I_SERIT_METNI = "Son güncelleme " & Format$(tarih, "dd.mm.yyyy hh:nn") & "  |  " & I_DEPOLAR(il).Count & " ilçe deposu" & metin & _
        "  |  Toplam stok deðeri " & Y_TL(toplam)
End Function

Private Sub I_PANEL_YENILE()
    Dim ws As Worksheet, sh As Worksheet, isg As Worksheet, i As Long, k As String, d As Object, il As String
    Dim satir As Long, deger As Double, tarih As Date, depo As Long
    Set ws = ThisWorkbook.Worksheets("I_ANA PANEL")
    Set isg = ThisWorkbook.Worksheets("I_ÝSG_MALZEMELERÝ")
    il = I_AKTIF_IL()
    ws.Unprotect ""
    ' Depo sayisi: ilin butun ilce depolari (stogu olmasa da).
    depo = I_DEPOLAR(il).Count
    If CStr(ws.Range("BX25").Value2) <> "IP_V4" Then I_HIZLI ws
    For i = 0 To 2
        k = Choose(i + 1, "B", "L", "Q")
        Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
        satir = 0: deger = 0: tarih = 0
        If i < 2 Then
            Set sh = ThisWorkbook.Worksheets(IIf(i = 0, "I_CDP_ÝLÇELER_STOK", "I_EPP_ÝLÇELER_STOK"))
            I_OZET sh, AKSORA_IL_TESIS(il, IIf(i = 0, "CDP", "EPP")), d, satir, deger, tarih
            If satir = 0 Then tarih = 0
        Else
            I_OZET ThisWorkbook.Worksheets("I_CDP_ÝLÇELER_STOK"), "", d, satir, deger, tarih
            I_OZET ThisWorkbook.Worksheets("I_EPP_ÝLÇELER_STOK"), "", d, satir, deger, tarih
            satir = CLng(I_SAYI(isg.Range("BZ21").Value2))
            deger = I_SAYI(isg.Range("BZ22").Value2)
        End If
        If tarih > 0 Then
            Y_KART_TARIH ws, k, tarih
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "1"), Format$(tarih, "dd.mm.yyyy hh:nn")
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "2"), AKSORA_SAYI_METNI(satir, "#,##0")
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "3"), Y_TL(deger)
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "4"), AKSORA_SAYI_METNI(depo, "#,##0")
        Else
            Y_KART_DURUM ws, k, "Veri bekleniyor", RGB(145, 162, 176), RGB(29, 78, 216)
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "1"), "-"
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "2"), "-"
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "3"), "-"
            Y_YAZ Y_SEKIL(ws, "AKSORA15_KART_DEGER_" & k & "4"), "-"
        End If
    Next i
    Y_YAZ Y_SEKIL(ws, "AKSORA15_PANEL_TARIHI"), "Ýlçe Yönetim Birimi  |  Panel kontrolü: " & Format$(Now, "dd.mm.yyyy hh:nn")
    Y_KORU ws
    Y_SERIT I_SERIT_METNI(), 0, "I_ANA PANEL"
End Sub

'---------- Gunluk stok gecmisi (karsilastirma icin) ----------
' _IH_<IL>: her sutun bir gun (gunun son guncellemesi). 1. satir tarih, alti metin bloklari.
' Satir: depo TAB uretim yeri TAB malzeme TAB degerleme TAB depo adi TAB tanim TAB birim TAB miktar TAB deger.
' (Eski 8 alanli kayitlarda degerleme yoktur.)

Private Function I_TEMIZ(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    I_TEMIZ = Replace(Replace(Replace(Trim$(CStr(v)), vbTab, " "), vbLf, " "), vbCr, " ")
End Function

Private Function I_GECMIS_SON(ByVal ws As Worksheet) As Long
    I_GECMIS_SON = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If I_GECMIS_SON = 1 And Len(CStr(ws.Cells(1, 1).Value2)) = 0 Then I_GECMIS_SON = 0
End Function

Private Sub I_GECMIS_EKLE(ByVal il As String, ByVal a As Variant, ByVal tarih As Date)
    Dim ws As Worksheet, d As Object, r As Long, k As String, t As Variant, c As Long, son As Long
    Dim satir As String, tampon As String, j As Long, x As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    If IsArray(a) Then
        For r = 1 To UBound(a, 1)
            k = I_TEMIZ(a(r, 2)) & vbTab & I_TEMIZ(a(r, 1)) & vbTab & I_TEMIZ(a(r, 5)) & vbTab & I_TEMIZ(a(r, 9))
            If d.Exists(k) Then
                t = d(k)
            Else
                t = Array(I_TEMIZ(a(r, 3)), I_TEMIZ(a(r, 6)), I_TEMIZ(a(r, 8)), 0#, 0#)
            End If
            t(3) = t(3) + I_SAYI(a(r, 7)): t(4) = t(4) + I_SAYI(a(r, 11))
            d(k) = t
        Next r
    End If
    Set ws = Y_KAYIT_SAYFA("_IH_" & UCase$(Trim$(il)))
    ' Ayni gunun onceki kaydi yerine yenisi yazilir; en fazla I_GECMIS_GUN gun tutulur.
    son = I_GECMIS_SON(ws)
    For c = son To 1 Step -1
        x = ws.Cells(1, c).Value
        If IsDate(x) Then
            If Int(CDbl(CDate(x))) = Int(CDbl(tarih)) Then ws.Columns(c).Delete
        Else
            ws.Columns(c).Delete
        End If
    Next c
    son = I_GECMIS_SON(ws)
    Do While son >= I_GECMIS_GUN
        ws.Columns(1).Delete
        son = son - 1
    Loop
    c = son + 1
    ws.Columns(c).NumberFormat = "@"
    ws.Cells(1, c).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    ws.Cells(1, c).Value = tarih
    j = 1
    For Each x In d.keys
        t = d(x)
        satir = CStr(x) & vbTab & t(0) & vbTab & t(1) & vbTab & t(2) & vbTab & Trim$(Str$(t(3))) & vbTab & Trim$(Str$(Round(t(4), 2)))
        If Len(tampon) + Len(satir) + 1 > 30000 And Len(tampon) > 0 Then
            j = j + 1: ws.Cells(j, c).Value2 = tampon: tampon = vbNullString
        End If
        If Len(tampon) > 0 Then tampon = tampon & vbLf
        tampon = tampon & satir
    Next x
    If Len(tampon) > 0 Then j = j + 1: ws.Cells(j, c).Value2 = tampon
End Sub

' Kayitli gunler, en yeni ustte.
Private Function I_GECMIS_TARIHLER(ByVal il As String) As Collection
    Dim ws As Worksheet, son As Long, c As Long, x As Variant, sonuc As Collection
    Set sonuc = New Collection
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_IH_" & UCase$(Trim$(il)))
    On Error GoTo 0
    If Not ws Is Nothing Then
        son = I_GECMIS_SON(ws)
        For c = son To 1 Step -1
            x = ws.Cells(1, c).Value
            If IsDate(x) Then sonuc.Add CDate(x)
        Next c
    End If
    Set I_GECMIS_TARIHLER = sonuc
End Function

' Bir gunun stoku: depo|uretim yeri|malzeme|degerleme -> (depo adi, tanim, birim, miktar, deger).
Private Function I_GECMIS_OKU(ByVal il As String, ByVal metin As String) As Object
    Dim ws As Worksheet, son As Long, c As Long, x As Variant, d As Object, r As Long, satirlar() As String, f() As String
    Dim i As Long
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set I_GECMIS_OKU = d
    If Len(metin) = 0 Then Exit Function
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_IH_" & UCase$(Trim$(il)))
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    son = I_GECMIS_SON(ws)
    For c = 1 To son
        x = ws.Cells(1, c).Value
        If IsDate(x) Then
            If Format$(CDate(x), I_TF) = metin Then Exit For
        End If
    Next c
    If c > son Then Exit Function
    r = 2
    Do While Len(CStr(ws.Cells(r, c).Value2)) > 0
        AKSORA_BEKLE_NABIZ
        satirlar = Split(CStr(ws.Cells(r, c).Value2), vbLf)
        For i = 0 To UBound(satirlar)
            f = Split(satirlar(i), vbTab)
            If UBound(f) >= 8 Then
                d(f(0) & "|" & f(1) & "|" & f(2) & "|" & f(3)) = Array(f(4), f(5), f(6), Val(f(7)), Val(f(8)))
            ElseIf UBound(f) = 7 Then
                d(f(0) & "|" & f(1) & "|" & f(2) & "|") = Array(f(3), f(4), f(5), Val(f(6)), Val(f(7)))
            End If
        Next i
        r = r + 1
    Loop
End Function

'---------- AYLIK DEMONTAJ (ilce) ----------
' Demontaj defteri _ID_<IL>: her sutun bir gun. 1. satir gun, 2. satir kaynak (eski|yeni guncelleme),
' alti metin bloklari. Satir: depo TAB uretim yeri TAB malzeme TAB kullanilan TAB demontaj girilen TAB kullanilan TL.
' Gun degeri, o gunun son guncellemesi ile bir onceki kayitli gunun son guncellemesi arasindaki farktir
' (gunluk STOK/DEMONTAJ kontroluyle ayni kural). Aylik eksik = kullanilan - girilen (depo|uretim yeri|malzeme);
' ay icindeki gec girisler de sayilir. Malzeme bilgisi (mal grubu, tanim, birim) _I_MALZEME sayfasinda birikir.

' Demontaj kontrolu disi: ISG malzemeleri ve sayac / modem / aparat-muhur mal gruplari. Bos = kontrol edilir.
Private Function I_DEM_DISI(ByVal mal As String, ByVal isgL As Object) As String
    Dim d As Object, k As String, mg As String
    If isgL.Exists(Trim$(mal)) Or isgL.Exists(Y_KOD_NORM(mal)) Then I_DEM_DISI = "ÝSG malzemesi (demontaj gerekmez)": Exit Function
    Set d = I_MAL_HARITA()
    k = Y_KOD_NORM(mal)
    If d.Exists(k) Then mg = CStr(d(k)(0))
    Select Case mg
        Case "M03550000": I_DEM_DISI = "Sayaç (demontaj kontrolü dýþý)"
        Case "M03250000": I_DEM_DISI = "Modem (demontaj kontrolü dýþý)"
        Case "M03550100": I_DEM_DISI = "Aparat / mühür (demontaj kontrolü dýþý)"
    End Select
End Function

' Malzeme -> Array(mal grubu, tanim, birim). Bos ise il kayitlarindan (_IA_) doldurulur.
Private Function I_MAL_HARITA() As Object
    Dim ws As Worksheet, v As Variant, r As Long, son As Long, il As Variant, a As Variant
    If Not mIMal Is Nothing Then Set I_MAL_HARITA = mIMal: Exit Function
    Set mIMal = CreateObject("Scripting.Dictionary"): mIMal.CompareMode = vbTextCompare
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_I_MALZEME")
    On Error GoTo 0
    If Not ws Is Nothing Then
        son = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If Len(CStr(ws.Cells(son, 1).Value2)) > 0 Then
            v = ws.Range("A1:D" & son).Value2
            For r = 1 To UBound(v, 1)
                If Len(CStr(v(r, 1))) > 0 Then mIMal(CStr(v(r, 1))) = Array(CStr(v(r, 2)), CStr(v(r, 3)), CStr(v(r, 4)))
            Next r
        End If
    End If
    If mIMal.Count = 0 Then
        For Each il In ADM_ILLER()
            a = Empty
            On Error Resume Next
            a = AKSORA_PAKET_OKU(I_ARSIV_SAYFA(CStr(il), False))
            On Error GoTo 0
            I_MAL_DIZI a
        Next il
        If mIMalKirli Then I_MAL_KAYDET
    End If
    Set I_MAL_HARITA = mIMal
End Function

' Ilce stok dizisi (4 mal grubu, 5 malzeme, 6 tanim, 8 birim) haritaya eklenir.
Private Sub I_MAL_DIZI(ByVal a As Variant)
    Dim r As Long, k As String, mg As String, t As Variant, d As Object
    If Not IsArray(a) Then Exit Sub
    On Error Resume Next
    Set d = I_MAL_HARITA()
    For r = 1 To UBound(a, 1)
        k = Y_KOD_NORM(a(r, 5))
        If Len(k) > 0 Then
            mg = UCase$(I_TEMIZ(a(r, 4)))
            If d.Exists(k) Then
                t = d(k)
                If (Len(mg) > 0 And CStr(t(0)) <> mg) Or Len(CStr(t(1))) = 0 Then
                    If Len(mg) > 0 Then t(0) = mg
                    If Len(CStr(t(1))) = 0 Then t(1) = I_TEMIZ(a(r, 6)): t(2) = I_TEMIZ(a(r, 8))
                    d(k) = t: mIMalKirli = True
                End If
            Else
                d(k) = Array(mg, I_TEMIZ(a(r, 6)), I_TEMIZ(a(r, 8))): mIMalKirli = True
            End If
        End If
    Next r
End Sub

' Gecmisten gelen tanim / birim (artik stokta olmayan malzemeler icin).
Private Sub I_MAL_NOT(ByVal mal As String, ByVal tanim As String, ByVal birim As String)
    Dim d As Object, k As String
    Set d = I_MAL_HARITA()
    k = Y_KOD_NORM(mal)
    If Len(k) = 0 Then Exit Sub
    If Not d.Exists(k) Then d(k) = Array("", tanim, birim): mIMalKirli = True
End Sub

Private Sub I_MAL_KAYDET()
    Dim ws As Worksheet, v() As Variant, k As Variant, i As Long, t As Variant
    If mIMal Is Nothing Then Exit Sub
    If mIMal.Count = 0 Then Exit Sub
    Set ws = Y_KAYIT_SAYFA("_I_MALZEME")
    ws.Cells.ClearContents
    ReDim v(1 To mIMal.Count, 1 To 4)
    For Each k In mIMal.keys
        i = i + 1: t = mIMal(k)
        v(i, 1) = CStr(k): v(i, 2) = t(0): v(i, 3) = t(1): v(i, 4) = t(2)
    Next k
    ws.Range("A1").Resize(i, 4).NumberFormat = "@"
    ws.Range("A1").Resize(i, 4).Value2 = v
    mIMalKirli = False
End Sub

Private Function I_MAL_BILGI(ByVal mal As String) As Variant
    Dim d As Object, k As String
    Set d = I_MAL_HARITA()
    k = Y_KOD_NORM(mal)
    If d.Exists(k) Then I_MAL_BILGI = d(k) Else I_MAL_BILGI = Array("", "", "")
End Function

' Gecmisteki her gun cifti icin deftere gun yazilir (yeni veya kaynagi degismis gunler).
Private Sub I_DEFTER_GUNCELLE(ByVal il As String)
    Dim tarihler As Collection, ws As Worksheet, i As Long, c As Long, son As Long, imza As String, x As Variant
    Dim gunler As Object, gun As Long, isgL As Object, sinir As Long, yaz As Boolean
    Dim eskiT As String, yeniT As String, onceT As String, dOnce As Object, dE As Object, dY As Object
    On Error GoTo Bitti
    il = UCase$(Trim$(il))
    Set tarihler = I_GECMIS_TARIHLER(il)
    If tarihler.Count < 2 Then GoTo Bitti
    Set ws = Y_KAYIT_SAYFA("_ID_" & il)
    Set gunler = CreateObject("Scripting.Dictionary")
    son = I_GECMIS_SON(ws)
    For c = 1 To son
        x = ws.Cells(1, c).Value
        If IsDate(x) Then gunler(CLng(Int(CDbl(CDate(x))))) = c
    Next c
    Set isgL = I_ISG_LISTE()
    For i = tarihler.Count - 1 To 1 Step -1
        eskiT = Format$(tarihler(i + 1), I_TF): yeniT = Format$(tarihler(i), I_TF)
        imza = eskiT & "|" & yeniT
        gun = CLng(Int(CDbl(tarihler(i))))
        c = 0: yaz = True
        If gunler.Exists(gun) Then
            c = gunler(gun)
            If CStr(ws.Cells(2, c).Value2) = imza Then yaz = False
        End If
        If yaz Then
            AKSORA_BEKLE_NABIZ
            If c = 0 Then
                c = I_GECMIS_SON(ws) + 1
                gunler(gun) = c
            End If
            If eskiT = onceT And Not dOnce Is Nothing Then Set dE = dOnce Else Set dE = I_GECMIS_OKU(il, eskiT)
            Set dY = I_GECMIS_OKU(il, yeniT)
            I_DEFTER_GUN ws, c, gun, imza, dE, dY, isgL
            onceT = yeniT: Set dOnce = dY
        End If
    Next i
    ' En fazla I_DEFTER_SURE gun tutulur.
    sinir = CLng(Int(CDbl(Date))) - I_DEFTER_SURE
    For c = I_GECMIS_SON(ws) To 1 Step -1
        x = ws.Cells(1, c).Value
        If IsDate(x) Then
            If CLng(Int(CDbl(CDate(x)))) < sinir Then ws.Columns(c).Delete
        End If
    Next c
Bitti:
    On Error Resume Next
    If mIMalKirli Then I_MAL_KAYDET
    On Error GoTo 0
End Sub

Private Sub I_DEFTER_GUN(ByVal ws As Worksheet, ByVal c As Long, ByVal gun As Long, ByVal imza As String, _
    ByVal dE As Object, ByVal dY As Object, ByVal isgL As Object)
    Dim anahtar As Object, grup As Object, k As Variant, x As Variant, p() As String, gk As String
    Dim qE As Double, qY As Double, qV As Double, t As Variant, g As Variant
    Dim satir As String, tampon As String, j As Long, sonR As Long
    Set anahtar = CreateObject("Scripting.Dictionary"): anahtar.CompareMode = vbTextCompare
    Set grup = CreateObject("Scripting.Dictionary"): grup.CompareMode = vbTextCompare
    For Each x In Array(dE, dY)
        For Each k In x.keys
            anahtar(k) = True
        Next k
    Next x
    For Each k In anahtar.keys
        qE = 0: qY = 0: qV = 0
        If dE.Exists(k) Then t = dE(k): qE = t(3): qV = t(4)
        If dY.Exists(k) Then t = dY(k): qY = t(3)
        p = Split(CStr(k), "|")
        If UBound(p) >= 3 And Abs(qY - qE) > 0.0005 Then
            I_MAL_NOT p(2), CStr(t(1)), CStr(t(2))
            If Len(I_DEM_DISI(p(2), isgL)) = 0 Then
                gk = p(0) & "|" & p(1) & "|" & p(2)
                If Not grup.Exists(gk) Then grup(gk) = Array(0#, 0#, 0#)
                g = grup(gk)
                If I_DEMONTAJ_MI(p(3)) Then
                    If qY > qE Then g(1) = g(1) + (qY - qE)
                ElseIf qY < qE Then
                    g(0) = g(0) + (qE - qY)
                    If qE > 0 Then g(2) = g(2) + (qE - qY) * qV / qE
                End If
                grup(gk) = g
            End If
        End If
    Next k
    sonR = ws.Cells(ws.Rows.Count, c).End(xlUp).Row
    If sonR >= 2 Then ws.Range(ws.Cells(2, c), ws.Cells(sonR, c)).ClearContents
    ws.Columns(c).NumberFormat = "@"
    ws.Cells(1, c).NumberFormat = "dd.mm.yyyy"
    ws.Cells(1, c).Value = CDate(gun)
    j = 2
    For Each k In grup.keys
        g = grup(k)
        If g(0) > 0.0005 Or g(1) > 0.0005 Then
            satir = Replace(CStr(k), "|", vbTab) & vbTab & Trim$(Str$(Round(g(0), 4))) & vbTab & _
                Trim$(Str$(Round(g(1), 4))) & vbTab & Trim$(Str$(Round(g(2), 2)))
            If Len(tampon) + Len(satir) + 1 > 30000 And Len(tampon) > 0 Then
                j = j + 1: ws.Cells(j, c).Value2 = tampon: tampon = vbNullString
            End If
            If Len(tampon) > 0 Then tampon = tampon & vbLf
            tampon = tampon & satir
        End If
    Next k
    If Len(tampon) > 0 Then j = j + 1: ws.Cells(j, c).Value2 = tampon
    ' Imza en son yazilir: yazim yarida kalirsa gun bir sonraki guncellemede yeniden hesaplanir.
    ws.Cells(2, c).Value2 = imza
End Sub

' Defter: ay (yyyymm) -> Array(gun sayisi, ilk gun, son gun, kalemler)
'   kalemler: depo|uretim yeri|malzeme -> Array(kullanilan, girilen, kullanilan TL)
Private Function I_DEFTER_OKU(ByVal il As String) As Object
    Dim ws As Worksheet, son As Long, c As Long, x As Variant, r As Long, satirlar() As String, f() As String, i As Long
    Dim aylar As Object, ay As Long, o As Variant, kal As Object, k As String, t As Variant, imza As String, v As Variant
    If mIDefter Is Nothing Then Set mIDefter = CreateObject("Scripting.Dictionary")
    il = UCase$(Trim$(il))
    Set aylar = CreateObject("Scripting.Dictionary")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("_ID_" & il)
    On Error GoTo 0
    If ws Is Nothing Then Set I_DEFTER_OKU = aylar: Exit Function
    son = I_GECMIS_SON(ws)
    If son = 0 Then Set I_DEFTER_OKU = aylar: Exit Function
    v = ws.Range(ws.Cells(2, 1), ws.Cells(2, son + 1)).Value2
    imza = CStr(son)
    For c = 1 To son
        imza = imza & ";" & CStr(v(1, c))
    Next c
    If mIDefter.Exists(il) Then
        o = mIDefter(il)
        If CStr(o(0)) = imza Then Set I_DEFTER_OKU = o(1): Exit Function
    End If
    For c = 1 To son
        x = ws.Cells(1, c).Value
        If IsDate(x) Then
            ay = Year(CDate(x)) * 100 + Month(CDate(x))
            If Not aylar.Exists(ay) Then
                Set kal = CreateObject("Scripting.Dictionary"): kal.CompareMode = vbTextCompare
                aylar.Add ay, Array(0&, CDate(x), CDate(x), kal)
            End If
            o = aylar(ay)
            o(0) = o(0) + 1
            If CDate(x) < o(1) Then o(1) = CDate(x)
            If CDate(x) > o(2) Then o(2) = CDate(x)
            Set kal = o(3)
            aylar(ay) = o
            AKSORA_BEKLE_NABIZ
            r = 3
            Do While Len(CStr(ws.Cells(r, c).Value2)) > 0
                satirlar = Split(CStr(ws.Cells(r, c).Value2), vbLf)
                For i = 0 To UBound(satirlar)
                    f = Split(satirlar(i), vbTab)
                    If UBound(f) >= 5 Then
                        k = f(0) & "|" & f(1) & "|" & f(2)
                        If kal.Exists(k) Then t = kal(k) Else t = Array(0#, 0#, 0#)
                        t(0) = t(0) + Val(f(3)): t(1) = t(1) + Val(f(4)): t(2) = t(2) + Val(f(5))
                        kal(k) = t
                    End If
                Next i
                r = r + 1
            Loop
        End If
    Next c
    If mIDefter.Exists(il) Then mIDefter.Remove il
    mIDefter.Add il, Array(imza, aylar)
    Set I_DEFTER_OKU = aylar
End Function

' Bir ayin kalemlerinden ozet. Donus: Array(top, depolar, malzemeler, detay)
'   top / depo: Array(kullanilan TL, eksik TL, kullanilan kalem, eksik kalem, tamam kalem, giris yok, eksik giris, girilen TL,
'               kritik kullanilan TL, kritik eksik TL, kritik kullanilan kalem, kritik eksik kalem)
'   malzeme: Array(kullanilan, girilen, eksik, eksik TL, kullanilan TL, ilce sayisi, girilen (ham))
'   detay (depo|malzeme): Array(kullanilan, girilen, eksik, eksik TL, kullanilan TL, girilen (ham))
' Kritik: kablo, direk, pano, hucre, trafo (I_KATEGORI).
Private Function I_AY_HESAP(ByVal kal As Object) As Variant
    Dim top As Variant, dD As Object, dM As Object, dX As Object, mIl As Object, k As Variant, p() As String, kr As Boolean
    Dim t As Variant, o As Variant, u As Double, e As Double, v As Double, ek As Double, ekTl As Double, gi As Double, xk As String
    top = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    Set dD = CreateObject("Scripting.Dictionary"): dD.CompareMode = vbTextCompare
    Set dM = CreateObject("Scripting.Dictionary"): dM.CompareMode = vbTextCompare
    Set dX = CreateObject("Scripting.Dictionary"): dX.CompareMode = vbTextCompare
    Set mIl = CreateObject("Scripting.Dictionary"): mIl.CompareMode = vbTextCompare
    For Each k In kal.keys
        t = kal(k): u = t(0): e = t(1): v = t(2)
        If u > 0.0005 Then
            p = Split(CStr(k), "|")
            ek = u - e
            If ek < 0.0005 Then ek = 0
            gi = u - ek
            ekTl = ek * v / u
            kr = (Len(I_KATEGORI(p(2))) > 0)
            I_AY_EKLE top, v, ekTl, ek, e, gi, kr
            If dD.Exists(p(0)) Then o = dD(p(0)) Else o = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
            I_AY_EKLE o, v, ekTl, ek, e, gi, kr
            dD(p(0)) = o
            If dM.Exists(p(2)) Then o = dM(p(2)) Else o = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#)
            o(0) = o(0) + u: o(1) = o(1) + gi: o(2) = o(2) + ek: o(3) = o(3) + ekTl: o(4) = o(4) + v: o(6) = o(6) + e
            If Not mIl.Exists(p(2) & "|" & p(0)) Then mIl(p(2) & "|" & p(0)) = True: o(5) = o(5) + 1
            dM(p(2)) = o
            xk = p(0) & "|" & p(2)
            If dX.Exists(xk) Then o = dX(xk) Else o = Array(0#, 0#, 0#, 0#, 0#, 0#)
            o(0) = o(0) + u: o(1) = o(1) + gi: o(2) = o(2) + ek: o(3) = o(3) + ekTl: o(4) = o(4) + v: o(5) = o(5) + e
            dX(xk) = o
        End If
    Next k
    I_AY_HESAP = Array(top, dD, dM, dX)
End Function

Private Sub I_AY_EKLE(ByRef o As Variant, ByVal v As Double, ByVal ekTl As Double, ByVal ek As Double, ByVal e As Double, _
    ByVal gi As Double, ByVal kritik As Boolean)
    o(0) = o(0) + v: o(1) = o(1) + ekTl: o(2) = o(2) + 1
    If ek > 0 Then
        o(3) = o(3) + 1
        If e <= 0.0005 Then o(5) = o(5) + 1 Else o(6) = o(6) + 1
    Else
        o(4) = o(4) + 1
    End If
    If v > 0 And gi > 0 Then o(7) = o(7) + v - ekTl
    If kritik Then
        o(8) = o(8) + v: o(9) = o(9) + ekTl: o(10) = o(10) + 1
        If ek > 0 Then o(11) = o(11) + 1
    End If
End Sub

' Oncelikli (kritik) demontaj kategorisi: kablo, direk, pano, hucre, trafo. Digerleri (sigorta, ampul vb.) bos.
Private Function I_KATEGORI(ByVal mal As String) As String
    Dim b As Variant, mg As String, tn As String
    b = I_MAL_BILGI(mal)
    mg = UCase$(Trim$(CStr(b(0)))): tn = UCase$(Trim$(CStr(b(1))))
    If Left$(tn, 6) = "IPTAL." Then tn = Trim$(Mid$(tn, 7))
    Select Case True
        Case Left$(mg, 5) = "M0316": I_KATEGORI = "KABLO"
        Case Left$(mg, 5) = "M0308": I_KATEGORI = "DÝREK"
        Case Left$(mg, 5) = "M0327": I_KATEGORI = "PANO"
        Case Left$(mg, 5) = "M0314": I_KATEGORI = "HÜCRE"
        Case Left$(mg, 5) = "M0340": I_KATEGORI = "TRAFO"
        Case Len(mg) > 0: I_KATEGORI = vbNullString
        Case Left$(tn, 6) = "KABLO ": I_KATEGORI = "KABLO"
        Case Left$(tn, 6) = "DIREK ": I_KATEGORI = "DÝREK"
        Case Left$(tn, 5) = "PANO ": I_KATEGORI = "PANO"
        Case Left$(tn, 6) = "HUCRE ": I_KATEGORI = "HÜCRE"
        Case Left$(tn, 6) = "TRAFO ": I_KATEGORI = "TRAFO"
    End Select
End Function

Private Function I_KAT_SIRA(ByVal kat As String) As Long
    Select Case kat
        Case "KABLO": I_KAT_SIRA = 1
        Case "DÝREK": I_KAT_SIRA = 2
        Case "PANO": I_KAT_SIRA = 3
        Case "HÜCRE": I_KAT_SIRA = 4
        Case "TRAFO": I_KAT_SIRA = 5
        Case Else: I_KAT_SIRA = 9
    End Select
End Function

' Kategori ozeti (malzeme listesinden): kat -> Array(cesit, kullanilan, girilen, eksik, kullanilan TL, eksik TL,
'   eksik cesit, birim, birim tutarli, girilen (ham)). Kritik olmayanlar "DÝÐER".
Private Function I_KAT_OZET(ByVal dM As Object) As Object
    Dim d As Object, k As Variant, t As Variant, kat As String, o As Variant, b As Variant
    Set d = CreateObject("Scripting.Dictionary")
    For Each k In dM.keys
        t = dM(k): kat = I_KATEGORI(CStr(k))
        If Len(kat) = 0 Then kat = "DÝÐER"
        b = I_MAL_BILGI(CStr(k))
        If d.Exists(kat) Then o = d(kat) Else o = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, CStr(b(2)), True, 0#)
        o(0) = o(0) + 1: o(1) = o(1) + t(0): o(2) = o(2) + t(1): o(3) = o(3) + t(2)
        o(4) = o(4) + t(4): o(5) = o(5) + t(3): o(9) = o(9) + t(6)
        If t(2) > 0.0005 Then o(6) = o(6) + 1
        If UCase$(CStr(o(7))) <> UCase$(CStr(b(2))) Then o(8) = False
        d(kat) = o
    Next k
    Set I_KAT_OZET = d
End Function

' Kritik malzeme uyumu (%): deger agirlikli; deger yoksa kalem bazli.
Private Function I_UYUM_K(ByVal t As Variant) As Double
    If t(8) > 0.005 Then
        I_UYUM_K = (t(8) - t(9)) / t(8) * 100
    ElseIf t(10) > 0 Then
        I_UYUM_K = (t(10) - t(11)) / t(10) * 100
    Else
        I_UYUM_K = 100
    End If
End Function

' Kategori dizisinin uyumu (%).
Private Function I_UYUM_KAT(ByVal o As Variant) As Double
    If o(4) > 0.005 Then
        I_UYUM_KAT = (o(4) - o(5)) / o(4) * 100
    ElseIf o(1) > 0 Then
        I_UYUM_KAT = (o(1) - o(3)) / o(1) * 100
    Else
        I_UYUM_KAT = 100
    End If
End Function

Private Function I_OZET_DURUM(ByVal t As Variant) As String
    If t(11) > 0 Then
        I_OZET_DURUM = "KRÝTÝK EKSÝK"
    ElseIf t(3) > 0 Then
        I_OZET_DURUM = "DÝÐER EKSÝK"
    Else
        I_OZET_DURUM = "TAMAM"
    End If
End Function

Private Function I_KAT_KAPSAM(ByVal kat As String) As String
    Select Case kat
        Case "KABLO": I_KAT_KAPSAM = "Kablo çeþitleri"
        Case "DÝREK": I_KAT_KAPSAM = "Direk çeþitleri"
        Case "PANO": I_KAT_KAPSAM = "Panolar"
        Case "HÜCRE": I_KAT_KAPSAM = "Hücreler"
        Case "TRAFO": I_KAT_KAPSAM = "Trafolar"
        Case Else: I_KAT_KAPSAM = "Sigorta, ampul vb. (ikinci planda)"
    End Select
End Function

' Ayin ozeti: Array(top, depolar, malzemeler, detay, gun sayisi, ilk gun, son gun); kayit yoksa Empty.
Private Function I_AY_OZET(ByVal aylar As Object, ByVal ay As Long) As Variant
    Dim o As Variant, h As Variant
    If aylar Is Nothing Then Exit Function
    If Not aylar.Exists(ay) Then Exit Function
    o = aylar(ay)
    h = I_AY_HESAP(o(3))
    I_AY_OZET = Array(h(0), h(1), h(2), h(3), o(0), o(1), o(2))
End Function

' Demontaj uyumu (%): deger agirlikli; deger yoksa kalem bazli.
Private Function I_UYUM(ByVal t As Variant) As Double
    If t(0) > 0.005 Then
        I_UYUM = (t(0) - t(1)) / t(0) * 100
    ElseIf t(2) > 0 Then
        I_UYUM = t(4) / t(2) * 100
    Else
        I_UYUM = 100
    End If
End Function

Private Function I_AY_ADI(ByVal ay As Long) As String
    Dim m As Long
    m = ay Mod 100
    If m < 1 Or m > 12 Then Exit Function
    I_AY_ADI = Choose(m, "OCAK", "ÞUBAT", "MART", "NÝSAN", "MAYIS", "HAZÝRAN", "TEMMUZ", "AÐUSTOS", "EYLÜL", "EKÝM", "KASIM", "ARALIK") & " " & (ay \ 100)
End Function

Private Function I_AY_ONCEKI(ByVal ay As Long) As Long
    If ay Mod 100 = 1 Then I_AY_ONCEKI = (ay \ 100 - 1) * 100 + 12 Else I_AY_ONCEKI = ay - 1
End Function

' Aylar (en yeni ilk). Kayit yoksa Empty.
Private Function I_AY_LISTE(ByVal aylar As Object) As Variant
    Dim v() As Long, k As Variant, i As Long, j As Long, t As Long
    If aylar Is Nothing Then Exit Function
    If aylar.Count = 0 Then Exit Function
    ReDim v(1 To aylar.Count)
    For Each k In aylar.keys
        i = i + 1
        v(i) = CLng(k)
    Next k
    For i = 1 To UBound(v) - 1
        For j = i + 1 To UBound(v)
            If v(j) > v(i) Then t = v(i): v(i) = v(j): v(j) = t
        Next j
    Next i
    I_AY_LISTE = v
End Function

' Uyum yuzdesi: trafik isigi ikonu (yesil >= 90, sari >= 70, kirmizi < 70) + yesil veri cubugu.
Private Sub I_UYUM_CUBUK(ByVal r As Range)
    Dim db As Object, ic As Object
    r.NumberFormat = """%""0.0"
    r.HorizontalAlignment = xlCenter
    On Error Resume Next
    Set ic = r.FormatConditions.AddIconSetCondition
    ic.IconSet = ThisWorkbook.IconSets(xl3TrafficLights1)
    ic.IconCriteria(2).Type = xlConditionValueNumber: ic.IconCriteria(2).Value = 70: ic.IconCriteria(2).Operator = xlGreaterEqual
    ic.IconCriteria(3).Type = xlConditionValueNumber: ic.IconCriteria(3).Value = 90: ic.IconCriteria(3).Operator = xlGreaterEqual
    Set db = r.FormatConditions.AddDatabar
    db.MinPoint.Modify newtype:=xlConditionValueNumber, newvalue:=0
    db.MaxPoint.Modify newtype:=xlConditionValueNumber, newvalue:=100
    db.BarColor.Color = RGB(74, 197, 125)
    On Error GoTo 0
End Sub

' Fire degeri sutunu: beyazdan kirmiziya renk skalasi.
Private Sub I_FIRE_ISI(ByVal r As Range)
    Dim cs As Object
    On Error Resume Next
    Set cs = r.FormatConditions.AddColorScale(ColorScaleType:=2)
    cs.ColorScaleCriteria(1).Type = xlConditionValueLowestValue
    cs.ColorScaleCriteria(1).FormatColor.Color = RGB(255, 255, 255)
    cs.ColorScaleCriteria(2).Type = xlConditionValueHighestValue
    cs.ColorScaleCriteria(2).FormatColor.Color = RGB(252, 165, 165)
    On Error GoTo 0
End Sub

Private Function I_KAT_RENK(ByVal kat As String) As Long
    Select Case kat
        Case "KABLO": I_KAT_RENK = RGB(37, 99, 235)
        Case "DÝREK": I_KAT_RENK = RGB(180, 83, 9)
        Case "PANO": I_KAT_RENK = RGB(124, 58, 237)
        Case "HÜCRE": I_KAT_RENK = RGB(13, 148, 136)
        Case "TRAFO": I_KAT_RENK = RGB(220, 38, 38)
        Case Else: I_KAT_RENK = RGB(148, 163, 184)
    End Select
End Function

Private Function I_DEGERLENDIR(ByVal u As Double) As String
    If u >= 95 Then
        I_DEGERLENDIR = "ÇOK ÝYÝ"
    ElseIf u >= 80 Then
        I_DEGERLENDIR = "ÝYÝ"
    ElseIf u >= 50 Then
        I_DEGERLENDIR = "ORTA"
    Else
        I_DEGERLENDIR = "ZAYIF"
    End If
End Function

' Ilce karnesi (kritik malzemeler): depo -> Array(kritik kullanilan TL, fire TL, kullanilan kalem, eksik kalem,
'   kablo eksik, direk eksik, pano eksik, hucre eksik, trafo eksik, uyum %). Yalniz kritik kullanimi olan depolar.
Private Function I_KARNE(ByVal dD As Object, ByVal dX As Object) As Object
    Dim d As Object, k As Variant, t As Variant, o As Variant, p() As String, kat As String, s As Long
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For Each k In dD.keys
        t = dD(k)
        If t(10) > 0 Then d(CStr(k)) = Array(t(8), t(9), t(10), t(11), 0#, 0#, 0#, 0#, 0#, I_UYUM_K(t))
    Next k
    For Each k In dX.keys
        p = Split(CStr(k), "|")
        If d.Exists(p(0)) Then
            kat = I_KATEGORI(p(1))
            s = I_KAT_SIRA(kat)
            If s <= 5 Then
                t = dX(k): o = d(p(0))
                o(3 + s) = o(3 + s) + t(2)
                d(p(0)) = o
            End If
        End If
    Next k
    Set I_KARNE = d
End Function

Private Sub I_DURUM_RENK(ByVal r As Range)
    Y_DURUM_KF r, "TAMAM", RGB(220, 252, 231), RGB(21, 128, 61)
    Y_DURUM_KF r, "KRÝTÝK EKSÝK", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF r, "DÝÐER EKSÝK", RGB(254, 243, 199), RGB(180, 83, 9)
    Y_DURUM_KF r, "EKSÝK VAR", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF r, "GÝRÝÞ YOK", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF r, "EKSÝK GÝRÝÞ", RGB(254, 243, 199), RGB(180, 83, 9)
    Y_DURUM_KF r, "KULLANIM YOK", RGB(241, 245, 249), RGB(100, 116, 139)
    Y_DURUM_KF r, "ÇOK ÝYÝ", RGB(187, 247, 208), RGB(21, 128, 61)
    Y_DURUM_KF r, "ÝYÝ", RGB(220, 252, 231), RGB(22, 101, 52)
    Y_DURUM_KF r, "ORTA", RGB(254, 243, 199), RGB(180, 83, 9)
    Y_DURUM_KF r, "ZAYIF", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF r, ChrW(9733) & " EN ÝYÝ", RGB(22, 163, 74), RGB(255, 255, 255)
    Y_DURUM_KF r, ChrW(9660) & " EN KÖTÜ", RGB(220, 38, 38), RGB(255, 255, 255)
    r.HorizontalAlignment = xlCenter
End Sub

Private Function I_SATIR_DURUM(ByVal eksik As Double, ByVal girilenHam As Double) As String
    If eksik <= 0.0005 Then
        I_SATIR_DURUM = "TAMAM"
    ElseIf girilenHam <= 0.0005 Then
        I_SATIR_DURUM = "GÝRÝÞ YOK"
    Else
        I_SATIR_DURUM = "EKSÝK GÝRÝÞ"
    End If
End Function

Private Function I_SIRA_ANAHTAR(ByVal x As Double) As String
    I_SIRA_ANAHTAR = Format$(1E+15 - Round(x, 2) * 100, "0000000000000000")
End Function

Public Sub AKSORA_ILCE_AY_SEC()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  AYLIK DEMONTAJ", "Ay raporu hazýrlanýyor...")
    On Error GoTo Son
    ILCE_AY_SEC_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub ILCE_AY_SEC_ASIL()
    Dim ws As Worksheet, ay As String, ev As Boolean, no As Long, metin As String
    ev = application.EnableEvents
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("I_AYLIK_DEMONTAJ")
    ay = ws.Shapes(CStr(application.Caller)).AlternativeText
    application.ScreenUpdating = False: application.EnableEvents = False
    ws.Unprotect ""
    ws.Range("BX30").Value2 = ay
    I_AYLIK_KUR CStr(ws.Range("BX22").Value2)
    application.EnableEvents = ev: application.ScreenUpdating = True
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    application.EnableEvents = ev: application.ScreenUpdating = True
    Err.Raise no, , metin
End Sub

' I_AYLIK_DEMONTAJ (ilce birimi): yalniz kritik malzemeler (kablo, direk, pano, hucre, trafo).
' KPI: uyum, fire degeri, en iyi ve en kotu ilce. Tablolar: kritik malzemeler, ilce karnesi, kritik fire listesi.
Private Sub I_AYLIK_KUR(ByVal il As String)
    Dim ws As Worksheet, w As Double, sw As Double, s As Shape, aylar As Object, liste As Variant, secAy As Long
    Dim ayrim As Boolean
    Dim h As Variant, o As Variant, top As Variant, dD As Object, dM As Object, dX As Object, depolar As Object, karne As Object
    Dim k As Variant, t As Variant, mb As Variant, p() As String, r As Long, a1 As Long, i As Long, j As Long, n As Long
    Dim y As Double, bw As Double, kw As Double, anah() As String, idx() As Long, kod() As String, sonuc() As Variant
    Dim ad As String, ilk As Date, ilkVar As Boolean, dK As Object, kat As String, kats As Variant, fireli As Object
    Dim enIyi As String, enKotu As String, bos As String, br() As String
    Set ws = ThisWorkbook.Worksheets("I_AYLIK_DEMONTAJ")
    il = UCase$(Trim$(il))
    I_DEFTER_GUNCELLE il
    ws.Unprotect ""
    secAy = CLng(Val(CStr(ws.Range("BX30").Value2)))
    If CStr(ws.Range("BX22").Value2) <> il Then secAy = 0
    Y_SAYFA_SIFIRLA ws
    I_BAGLA ws
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    sw = (w - 630) / 7
    If sw < 70 Then sw = 70
    Y_KOLONLAR ws, Array(70, 180, 100, 280, sw, sw, sw, sw, sw, sw, sw)
    Set aylar = I_DEFTER_OKU(il)
    Set depolar = I_DEPOLAR(il)
    liste = I_AY_LISTE(aylar)
    If IsArray(liste) Then
        For i = 1 To UBound(liste)
            If liste(i) = secAy Then Exit For
        Next i
        If i > UBound(liste) Then secAy = liste(1)
        For Each k In aylar.keys
            o = aylar(k)
            If Not ilkVar Or o(1) < ilk Then ilk = o(1): ilkVar = True
        Next k
    Else
        secAy = 0
    End If
    ws.Range("BX22").Value2 = il: ws.Range("BX20").Value2 = il: ws.Range("BX30").Value2 = secAy
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10: ws.Rows(8).RowHeight = 30: ws.Rows(9).RowHeight = 10: ws.Rows(10).RowHeight = 10
    If IsArray(liste) Then
        n = application.Min(12, UBound(liste))
        bw = application.Min(130, (w - 11 * 6) / 12)
        y = ws.Rows(8).Top + 1
        For i = 1 To n
            Set s = Y_DUGME(ws, "I_AY_" & i, I_AY_ADI(CLng(liste(i))), (i - 1) * (bw + 6), y, bw, 28, _
                IIf(liste(i) = secAy, RGB(79, 70, 160), RGB(148, 163, 184)), "AKSORA_ILCE_AY_SEC")
            s.AlternativeText = CStr(liste(i))
        Next i
    End If
    h = I_AY_OZET(aylar, secAy)
    If IsArray(h) Then
        top = h(0): Set dD = h(1): Set dM = h(2): Set dX = h(3)
    Else
        top = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        Set dD = CreateObject("Scripting.Dictionary"): Set dM = CreateObject("Scripting.Dictionary"): Set dX = CreateObject("Scripting.Dictionary")
    End If
    Set dK = I_KAT_OZET(dM)
    Set karne = I_KARNE(dD, dX)
    ' ilce siralamasi: uyum yuksekten dusuge, esitlikte fire degeri az olan once
    n = karne.Count
    If n > 0 Then
        ReDim anah(1 To n): ReDim idx(1 To n): ReDim kod(1 To n)
        i = 0
        For Each k In karne.keys
            t = karne(k): i = i + 1: idx(i) = i: kod(i) = CStr(k)
            anah(i) = I_SIRA_ANAHTAR(t(9)) & Format$(Round(t(1), 2) * 100, "0000000000000000")
        Next k
        Y_SIRALA anah, idx, 1, n
        ' En iyi / en kotu ancak aralarinda gercek fark varsa gosterilir (tek ilce veya hepsi esitse yok).
        ayrim = False
        If n > 1 Then ayrim = (karne(kod(idx(1)))(9) > karne(kod(idx(n)))(9) + 0.05) Or (karne(kod(idx(n)))(1) > karne(kod(idx(1)))(1) + 0.005)
        ad = vbNullString
        If depolar.Exists(kod(idx(1))) Then ad = I_DEPO_KISA(CStr(depolar(kod(idx(1)))))
        enIyi = IIf(Len(ad) > 0, ad, kod(idx(1))) & "  •  %" & AKSORA_SAYI_METNI(karne(kod(idx(1)))(9), "0.0")
        ad = vbNullString
        If depolar.Exists(kod(idx(n))) Then ad = I_DEPO_KISA(CStr(depolar(kod(idx(n)))))
        enKotu = IIf(Len(ad) > 0, ad, kod(idx(n))) & "  •  %" & AKSORA_SAYI_METNI(karne(kod(idx(n)))(9), "0.0")
        If Not ayrim Then enKotu = "-"
    Else
        enIyi = "-": enKotu = "-"
    End If
    ' KPI
    ws.Range("BZ21").Value2 = "%" & AKSORA_SAYI_METNI(I_UYUM_K(top), "0.0")
    ws.Range("BZ22").Value2 = Round(top(9), 2): ws.Range("BZ22").NumberFormat = "#,##0.00 ""TL"""
    ws.Range("BZ23").Value2 = enIyi
    ws.Range("BZ24").Value2 = enKotu
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "KRÝTÝK DEMONTAJ UYUMU", "$BZ$21", RGB(39, 160, 94)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "KRÝTÝK FÝRE (EKSÝK DEMONTAJ)", "$BZ$22", RGB(220, 38, 38)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, ChrW(9733) & " EN ÝYÝ ÝLÇE", "$BZ$23", RGB(22, 163, 74)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, ChrW(9660) & " EN KÖTÜ ÝLÇE", "$BZ$24", RGB(185, 28, 28)
    For i = 3 To 4
        Set s = Y_SEKIL(ws, "Y_KPI_DEGER_" & i)
        If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 14
    Next i

    '--- kritik malzemeler
    r = 11
    Y_BOLUM ws, r, "KRÝTÝK MALZEMELER  –  " & I_AY_ADI(secAy), "Ay boyunca kullanýlan, demontajý girilen ve eksik (fire) miktarlar.", 11
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), Array("", "Kategori", "Malzeme çeþidi", "Demontaj oraný", _
        "Kullanýlan", "Demontajý girilen", "Eksik (fire)", "Fire deðeri", "Kullanýlan deðer", "Fireli ilçe", "Durum")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    Set fireli = CreateObject("Scripting.Dictionary"): fireli.CompareMode = vbTextCompare
    For Each k In dX.keys
        p = Split(CStr(k), "|")
        If dX(k)(2) > 0.0005 Then
            kat = I_KATEGORI(p(1))
            If Len(kat) > 0 Then fireli(kat & "|" & p(0)) = True
        End If
    Next k
    kats = Array("KABLO", "DÝREK", "PANO", "HÜCRE", "TRAFO")
    For i = 0 To 4
        kat = CStr(kats(i))
        r = r + 1
        ws.Cells(r, 1).Value2 = ChrW(9632): ws.Cells(r, 1).Font.Color = I_KAT_RENK(kat): ws.Cells(r, 1).Font.Size = 16
        ws.Cells(r, 2).Value2 = kat & "   " & I_KAT_KAPSAM(kat)
        If dK.Exists(kat) Then
            o = dK(kat)
            ws.Cells(r, 3).Value2 = o(0)
            ws.Cells(r, 4).Value2 = Round(I_UYUM_KAT(o), 1)
            If o(8) Then
                ws.Cells(r, 5).Value2 = o(1): ws.Cells(r, 6).Value2 = o(2): ws.Cells(r, 7).Value2 = o(3)
                For j = 5 To 7
                    Y_MIKTAR_BICIM ws.Cells(r, j), CStr(o(7))
                Next j
            Else
                ws.Cells(r, 5).Value2 = "farklý birimler"
            End If
            ws.Cells(r, 8).Value2 = Round(o(5), 2): ws.Cells(r, 9).Value2 = Round(o(4), 2)
            n = 0
            For Each k In fireli.keys
                If Left$(CStr(k), Len(kat) + 1) = kat & "|" Then n = n + 1
            Next k
            ws.Cells(r, 10).Value2 = n
            ws.Cells(r, 11).Value2 = IIf(o(6) > 0, "KRÝTÝK EKSÝK", "TAMAM")
            If o(3) > 0.0005 Then ws.Cells(r, 7).Font.Color = RGB(185, 28, 28)
        Else
            ws.Cells(r, 11).Value2 = "KULLANIM YOK"
        End If
    Next i
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 11))
    ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 24
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 2), ws.Cells(r, 2)).Font.bold = True
    ws.Range(ws.Cells(a1, 2), ws.Cells(r, 2)).IndentLevel = 1
    ws.Range(ws.Cells(a1, 3), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 5), ws.Cells(r, 7)).HorizontalAlignment = xlRight
    ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).Font.bold = True
    ws.Range(ws.Cells(a1, 8), ws.Cells(r, 9)).NumberFormat = "#,##0.00 ""TL"""
    I_UYUM_CUBUK ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4))
    I_FIRE_ISI ws.Range(ws.Cells(a1, 8), ws.Cells(r, 8))
    I_DURUM_RENK ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11))

    '--- ilce karnesi
    r = r + 2
    Y_BOLUM ws, r, "ÝLÇE KARNESÝ  –  " & I_AY_ADI(secAy) & "  (en iyiden en kötüye)", "Sýra: kritik demontaj uyumu; eþitlikte fire deðeri az olan önde.", 11
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), Array("Sýra", "Ýlçe", "Depo", "Kritik demontaj uyumu", _
        "Kablo eksik", "Direk eksik", "Pano eksik", "Hücre eksik", "Trafo eksik", "Fire deðeri", "Deðerlendirme")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    n = karne.Count
    If n > 0 Then
        For i = 1 To n
            k = kod(idx(i)): t = karne(k)
            r = r + 1
            ad = vbNullString
            If depolar.Exists(k) Then ad = I_DEPO_KISA(CStr(depolar(k)))
            ws.Cells(r, 1).Value2 = i
            ws.Cells(r, 2).Value2 = ad
            ws.Cells(r, 3).NumberFormat = "@": ws.Cells(r, 3).Value2 = CStr(k)
            ws.Cells(r, 4).Value2 = Round(t(9), 1)
            For j = 0 To 4
                ws.Cells(r, 5 + j).Value2 = t(4 + j)
                If t(4 + j) > 0.0005 Then ws.Cells(r, 5 + j).Font.Color = RGB(185, 28, 28): ws.Cells(r, 5 + j).Font.bold = True
            Next j
            ws.Cells(r, 10).Value2 = Round(t(1), 2)
            If i = 1 And ayrim Then
                ws.Cells(r, 11).Value2 = ChrW(9733) & " EN ÝYÝ"
            ElseIf i = n And ayrim Then
                ws.Cells(r, 11).Value2 = ChrW(9660) & " EN KÖTÜ"
            Else
                ws.Cells(r, 11).Value2 = I_DEGERLENDIR(t(9))
            End If
        Next i
        Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 11))
        ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 22
        If ayrim Then
            ws.Range(ws.Cells(a1, 1), ws.Cells(a1, 10)).Interior.Color = RGB(240, 253, 244)
            ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)).Interior.Color = RGB(254, 242, 242)
        End If
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 2)).Font.bold = True
        ws.Range(ws.Cells(a1, 2), ws.Cells(r, 2)).IndentLevel = 1
        ws.Range(ws.Cells(a1, 3), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 9)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).NumberFormat = "#,##0.## ""m"";-#,##0.## ""m"";""–"""
        ws.Range(ws.Cells(a1, 6), ws.Cells(r, 9)).NumberFormat = "#,##0.## ""adet"";-#,##0.## ""adet"";""–"""
        ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).NumberFormat = "#,##0.00 ""TL"""
        I_UYUM_CUBUK ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4))
        I_FIRE_ISI ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10))
        I_DURUM_RENK ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11))
    Else
        r = r + 1
        ws.Cells(r, 2).Value2 = "Bu ay ilçe depolarýnda kritik malzeme kullanýmý yok."
        ws.Cells(r, 2).Font.Italic = True: ws.Cells(r, 2).Font.Color = RGB(100, 116, 139)
    End If
    ' kritik kullanimi olmayan ilceler (tek satir)
    For Each k In depolar.keys
        If Not karne.Exists(CStr(k)) Then bos = bos & IIf(Len(bos) > 0, ", ", "") & I_DEPO_KISA(CStr(depolar(k)))
    Next k
    If Len(bos) > 0 Then
        r = r + 1
        ws.Cells(r, 2).Value2 = "Bu ay kritik malzeme kullanmayan ilçeler: " & bos
        ws.Cells(r, 2).Font.Size = 9: ws.Cells(r, 2).Font.Italic = True: ws.Cells(r, 2).Font.Color = RGB(100, 116, 139)
    End If

    '--- kritik fire listesi (yalniz eksik olanlar)
    r = r + 2
    Y_BOLUM ws, r, "KRÝTÝK FÝRE LÝSTESÝ  –  " & I_AY_ADI(secAy), "En yüksek fire deðeri üstte. Baþlýklardaki filtrelerle ilçe veya kategori seçilebilir.", 11
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 11)), Array("Depo", "Ýlçe", "Malzeme", "Malzeme tanýmý", "Kategori", _
        "Kullanýlan", "Demontajý girilen", "Eksik (fire)", "Fire deðeri", "Demontaj oraný", "Durum")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    n = 0
    If dX.Count > 0 Then
        ReDim anah(1 To dX.Count): ReDim idx(1 To dX.Count): ReDim kod(1 To dX.Count)
        For Each k In dX.keys
            t = dX(k): p = Split(CStr(k), "|")
            If t(2) > 0.0005 Then
                If Len(I_KATEGORI(p(1))) > 0 Then
                    n = n + 1: idx(n) = n: kod(n) = CStr(k)
                    anah(n) = I_SIRA_ANAHTAR(t(3)) & I_SIRA_ANAHTAR(t(2))
                End If
            End If
        Next k
    End If
    If n > 0 Then
        ReDim Preserve anah(1 To n): ReDim Preserve idx(1 To n)
        Y_SIRALA anah, idx, 1, n
        ReDim sonuc(1 To n, 1 To 11): ReDim br(1 To n)
        For i = 1 To n
            t = dX(kod(idx(i))): p = Split(kod(idx(i)), "|"): mb = I_MAL_BILGI(p(1))
            ad = vbNullString
            If depolar.Exists(p(0)) Then ad = I_DEPO_KISA(CStr(depolar(p(0))))
            sonuc(i, 1) = p(0): sonuc(i, 2) = ad: sonuc(i, 3) = p(1): sonuc(i, 4) = mb(1): sonuc(i, 5) = I_KATEGORI(p(1))
            sonuc(i, 6) = t(0): sonuc(i, 7) = t(1): sonuc(i, 8) = t(2): sonuc(i, 9) = Round(t(3), 2)
            sonuc(i, 10) = Round((t(0) - t(2)) / t(0) * 100, 1)
            sonuc(i, 11) = I_SATIR_DURUM(t(2), t(5))
            br(i) = CStr(mb(2))
        Next i
        ws.Range(ws.Cells(a1, 1), ws.Cells(a1 + n - 1, 5)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 1), ws.Cells(a1 + n - 1, 11)).Value2 = sonuc
        r = a1 + n - 1
        For i = 1 To n
            For j = 6 To 8
                Y_MIKTAR_BICIM ws.Cells(a1 + i - 1, j), br(i)
            Next j
            ws.Cells(a1 + i - 1, 5).Font.Color = I_KAT_RENK(CStr(sonuc(i, 5)))
        Next i
        Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 11))
        ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 19
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 3), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).Font.bold = True
        ws.Range(ws.Cells(a1, 2), ws.Cells(r, 2)).IndentLevel = 1
        ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).IndentLevel = 1
        ws.Range(ws.Cells(a1, 6), ws.Cells(r, 8)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(a1, 8), ws.Cells(r, 8)).Font.bold = True
        ws.Range(ws.Cells(a1, 8), ws.Cells(r, 8)).Font.Color = RGB(185, 28, 28)
        ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)).NumberFormat = "#,##0.00 ""TL"""
        I_FIRE_ISI ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9))
        I_UYUM_CUBUK ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10))
        I_DURUM_RENK ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11))
        ws.Range(ws.Cells(a1 - 1, 1), ws.Cells(r, 11)).AutoFilter
    Else
        r = r + 1
        ws.Cells(r, 2).Value2 = IIf(secAy > 0, "Bu ay kritik malzemelerde fire yok.", "Henüz kayýt yok.")
        ws.Cells(r, 2).Font.Italic = True: ws.Cells(r, 2).Font.Color = RGB(21, 128, 61)
    End If
    r = r + 2
    ws.Cells(r, 1).Value2 = "Kritik: kablo, direk, pano, hücre, trafo. Fire: kullanýlýp demontajý girilmeyen miktar (ay içindeki geç giriþler düþülür)." & _
        IIf(ilkVar, " Kayýt baþlangýcý: " & Format$(ilk, "dd.mm.yyyy") & ".", "")
    ws.Cells(r, 1).Font.Size = 9: ws.Cells(r, 1).Font.Color = RGB(71, 85, 105)
    Y_FILIGRAN ws
    ad = "AKSORA | AYLIK DEMONTAJ  |  " & Y_IL_ADI(il)
    If secAy > 0 Then ad = ad & "  |  " & I_AY_ADI(secAy)
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), ad
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

' Admin aylik satir: ad | donem | kritik eksik kalem | diger eksik kalem | kritik kullanilan TL | kritik eksik TL | kritik uyum | onceki ay | durum
Private Sub ADM_AY_SATIR(ByVal ws As Worksheet, ByVal r As Long, ByVal ad As String, ByVal donem As String, _
    ByVal t As Variant, ByVal tOn As Variant, ByVal toplam As Boolean)
    Dim drm As String
    ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = IIf(toplam, RGB(226, 236, 247), vbWhite)
    ws.Cells(r, 2).NumberFormat = "@": ws.Cells(r, 2).Value2 = ad: ws.Cells(r, 2).Font.bold = True
    ws.Range(ADM_ADRES(r, 3, r, 4)).Merge: ws.Cells(r, 3).NumberFormat = "@": ws.Cells(r, 3).Value2 = donem
    ws.Range(ADM_ADRES(r, 7, r, 8)).Merge: ws.Range(ADM_ADRES(r, 9, r, 10)).Merge
    If IsArray(t) Then
        ws.Cells(r, 5).Value2 = t(11): ws.Cells(r, 6).Value2 = t(3) - t(11)
        ws.Cells(r, 7).Value2 = Round(t(8), 2): ws.Cells(r, 9).Value2 = Round(t(9), 2)
        ws.Cells(r, 11).Value2 = Round(I_UYUM_K(t), 1)
        If IsArray(tOn) Then ws.Cells(r, 12).Value2 = Round(I_UYUM_K(tOn), 1)
        ws.Range(ADM_ADRES(r, 5, r, 6)).NumberFormat = "#,##0"
        ws.Cells(r, 7).NumberFormat = "#,##0.00 ""TL""": ws.Cells(r, 9).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ADM_ADRES(r, 11, r, 12)).NumberFormat = """%""0.0"
        ws.Cells(r, 6).Font.Color = RGB(100, 116, 139)
        drm = I_OZET_DURUM(t)
        Select Case drm
            Case "KRÝTÝK EKSÝK"
                ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(254, 226, 226), RGB(185, 28, 28), 8, True, xlCenter
                ws.Cells(r, 5).Font.Color = RGB(185, 28, 28): ws.Cells(r, 9).Font.Color = RGB(185, 28, 28)
                ws.Cells(r, 5).Font.bold = True: ws.Cells(r, 9).Font.bold = True
            Case "DÝÐER EKSÝK"
                ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(254, 243, 199), RGB(180, 83, 9), 8, True, xlCenter
            Case Else
                ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(220, 252, 231), RGB(21, 128, 61), 8, True, xlCenter
        End Select
    Else
        If Len(donem) = 0 Then ws.Cells(r, 3).Value2 = "kayýt yok"
        ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), "VERÝ YOK", RGB(241, 245, 249), RGB(100, 116, 139), 8, True, xlCenter
    End If
    ws.Range(ADM_ADRES(r, 2, r, 12)).HorizontalAlignment = xlCenter
    ws.Range(ADM_ADRES(r, 7, r, 10)).ShrinkToFit = True
End Sub

Private Sub ADM_TOP_TOPLA(ByRef hedef As Variant, ByVal t As Variant)
    Dim i As Long
    For i = 0 To UBound(t)
        hedef(i) = hedef(i) + t(i)
    Next i
End Sub

Private Sub ADM_MAL_TOPLA(ByVal hedef As Object, ByVal dM As Object)
    Dim k As Variant, t As Variant, o As Variant, i As Long
    For Each k In dM.keys
        t = dM(k)
        If hedef.Exists(k) Then
            o = hedef(k)
            For i = 0 To UBound(t)
                o(i) = o(i) + t(i)
            Next i
            hedef(k) = o
        Else
            hedef(k) = t
        End If
    Next k
End Sub

' En ust N malzeme (alan: 3 eksik TL, 4 kullanilan TL). sadeceEksik / sadeceKritik ile suzulur.
Private Function ADM_MAL_SIRA(ByVal dM As Object, ByVal alan As Long, ByVal adet As Long, ByVal sadeceEksik As Boolean, _
    ByVal sadeceKritik As Boolean) As Variant
    Dim anah() As String, idx() As Long, kod() As String, k As Variant, t As Variant, n As Long, i As Long, sonuc() As String
    If dM.Count = 0 Then Exit Function
    ReDim anah(1 To dM.Count): ReDim idx(1 To dM.Count): ReDim kod(1 To dM.Count)
    For Each k In dM.keys
        t = dM(k)
        If (Not sadeceEksik Or t(2) > 0.0005) And (Not sadeceKritik Or Len(I_KATEGORI(CStr(k))) > 0) Then
            n = n + 1: idx(n) = n: kod(n) = CStr(k)
            anah(n) = I_SIRA_ANAHTAR(t(alan)) & I_SIRA_ANAHTAR(t(4))
        End If
    Next k
    If n = 0 Then Exit Function
    ReDim Preserve anah(1 To n): ReDim Preserve idx(1 To n)
    Y_SIRALA anah, idx, 1, n
    ReDim sonuc(1 To application.Min(adet, n))
    For i = 1 To UBound(sonuc)
        sonuc(i) = kod(idx(i))
    Next i
    ADM_MAL_SIRA = sonuc
End Function

' Admin gunluk eksik listesi siralamasi: kritik malzemeler once, sonra eksik deger.
Private Function ADM_DEM_SKOR(ByVal t As Variant) As Double
    ADM_DEM_SKOR = t(8)
    If Len(I_KATEGORI(CStr(t(3)))) > 0 Then ADM_DEM_SKOR = ADM_DEM_SKOR + 1E+12
End Function

Private Function ADM_DEM_DURUM(ByVal t As Variant) As String
    Dim kat As String
    kat = I_KATEGORI(CStr(t(3)))
    ADM_DEM_DURUM = CStr(t(7))
    If Len(kat) > 0 Then ADM_DEM_DURUM = ADM_DEM_DURUM & "  •  " & kat
End Function

' Admin aylik demontaj (yalniz kritik malzemeler): il ozeti, kritik malzemeler, ilce karnesi (en kotu / en iyi).
Private Sub ADM_AYLIK_CIZ(ByVal ws As Worksheet, ByRef r As Long, ByVal aylik As Object, ByVal ayCur As Long, _
    ByVal secIl As String, ByVal ilMetin As String)
    Dim iller As Variant, p As Long, il As String, hA As Variant, hB As Variant, aT As Variant, aTO As Variant
    Dim aMal As Object, k As Variant, t As Variant, i As Long, j As Long, satirNo As Long, ayOnc As Long, n As Long
    Dim depolar As Object, anah() As String, idx() As Long, kod() As Variant, ad As String, dK As Object
    Dim kats As Variant, kat As String, o As Variant, karne As Object, tum As Object, ilE As Variant, fireli As Object
    Dim adet As Long, sec As Collection, x As Variant, drm As String, ayrim As Boolean
    iller = ADM_ILLER()
    ayOnc = I_AY_ONCEKI(ayCur)
    Set aMal = CreateObject("Scripting.Dictionary"): aMal.CompareMode = vbTextCompare
    Set tum = CreateObject("Scripting.Dictionary"): tum.CompareMode = vbTextCompare
    Set fireli = CreateObject("Scripting.Dictionary"): fireli.CompareMode = vbTextCompare
    r = r + 2
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "AYLIK DEMONTAJ (KRÝTÝK MALZEMELER)   •   " & ilMetin & "   •   " & I_AY_ADI(ayCur) & "   (kablo, direk, pano, hücre, trafo)"
    r = r + 1: satirNo = r
    aT = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#): aTO = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
    If Len(secIl) = 0 Then
        ADM_BASLIKLAR ws, r, 2, Array("Ýl", "Kritik kullaným deðeri", "", "Kritik fire deðeri", "", "Kritik uyum", "", "Kablo eksik", "Direk eksik", "Pano / hücre / trafo eksik", "Önceki ay", "Durum")
        ws.Range(ADM_ADRES(r, 3, r, 4)).Merge: ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 7, r, 8)).Merge
    End If
    For p = 0 To 4
        il = CStr(iller(p))
        If Len(secIl) = 0 Or il = secIl Then
            hA = I_AY_OZET(aylik(il), ayCur): hB = I_AY_OZET(aylik(il), ayOnc)
            o = Array(0#, 0#, 0#, 0#, 0#)
            If IsArray(hA) Then
                ADM_TOP_TOPLA aT, hA(0): ADM_MAL_TOPLA aMal, hA(2)
                If IsArray(hB) Then ADM_TOP_TOPLA aTO, hB(0)
                Set karne = I_KARNE(hA(1), hA(3))
                Set depolar = I_DEPOLAR(il)
                For Each k In karne.keys
                    t = karne(k)
                    ad = vbNullString
                    If depolar.Exists(k) Then ad = I_DEPO_KISA(CStr(depolar(k)))
                    tum(il & "|" & k) = Array(il, CStr(k), ad, t)
                    For j = 0 To 4
                        o(j) = o(j) + t(4 + j)
                        If t(4 + j) > 0.0005 Then fireli(j & "|" & il & "|" & k) = True
                    Next j
                Next k
            End If
            If Len(secIl) = 0 Then
                r = r + 1
                ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
                ws.Cells(r, 2).NumberFormat = "@": ws.Cells(r, 2).Value2 = Y_IL_ADI(il): ws.Cells(r, 2).Font.bold = True
                ws.Range(ADM_ADRES(r, 3, r, 4)).Merge: ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 7, r, 8)).Merge
                If IsArray(hA) Then
                    t = hA(0)
                    ws.Cells(r, 3).Value2 = Round(t(8), 2): ws.Cells(r, 5).Value2 = Round(t(9), 2)
                    ws.Cells(r, 9).Value2 = o(0): ws.Cells(r, 10).Value2 = o(1): ws.Cells(r, 11).Value2 = o(2) + o(3) + o(4)
                    If IsArray(hB) Then
                        If hB(0)(10) > 0 Then ws.Cells(r, 12).Value2 = Round(I_UYUM_K(hB(0)), 1)
                    End If
                    ' Kritik kullanim yoksa uyum %100 gosterilmez (yaniltici olur).
                    If t(10) > 0 Then
                        ws.Cells(r, 7).Value2 = Round(I_UYUM_K(t), 1)
                        drm = IIf(t(11) > 0, "KRÝTÝK EKSÝK", "TAMAM")
                        ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, IIf(t(11) > 0, RGB(254, 226, 226), RGB(220, 252, 231)), IIf(t(11) > 0, RGB(185, 28, 28), RGB(21, 128, 61)), 8, True, xlCenter
                    Else
                        ws.Cells(r, 7).Value2 = "–": ws.Cells(r, 7).Font.Color = RGB(148, 163, 184)
                        ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), "KULLANIM YOK", RGB(241, 245, 249), RGB(100, 116, 139), 8, True, xlCenter
                    End If
                Else
                    ws.Cells(r, 3).Value2 = "kayýt yok": ws.Cells(r, 3).Font.Color = RGB(148, 163, 184)
                    ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), "VERÝ YOK", RGB(241, 245, 249), RGB(100, 116, 139), 8, True, xlCenter
                End If
            End If
        End If
    Next p
    If Len(secIl) = 0 Then
        r = r + 1
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = RGB(226, 236, 247)
        ws.Range(ADM_ADRES(r, 3, r, 4)).Merge: ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 7, r, 8)).Merge
        ws.Cells(r, 2).Value2 = "TOPLAM": ws.Range(ADM_ADRES(r, 2, r, 13)).Font.bold = True
        ws.Cells(r, 3).Value2 = Round(aT(8), 2): ws.Cells(r, 5).Value2 = Round(aT(9), 2): ws.Cells(r, 7).Value2 = Round(I_UYUM_K(aT), 1)
        If aTO(10) > 0 Then ws.Cells(r, 12).Value2 = Round(I_UYUM_K(aTO), 1)
        ws.Range(ADM_ADRES(satirNo + 1, 3, r, 6)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ADM_ADRES(satirNo + 1, 9, r, 9)).NumberFormat = "#,##0.## ""m"";-#,##0.## ""m"";""–"""
        ws.Range(ADM_ADRES(satirNo + 1, 10, r, 11)).NumberFormat = "#,##0.## ""adet"";-#,##0.## ""adet"";""–"""
        ws.Range(ADM_ADRES(satirNo + 1, 2, r, 12)).HorizontalAlignment = xlCenter
        ws.Range(ADM_ADRES(satirNo + 1, 3, r, 6)).ShrinkToFit = True
        ws.Range(ADM_ADRES(satirNo + 1, 5, r, 6)).Font.Color = RGB(185, 28, 28)
        I_UYUM_CUBUK ws.Range(ADM_ADRES(satirNo + 1, 7, r, 8))
        I_UYUM_CUBUK ws.Range(ADM_ADRES(satirNo + 1, 12, r, 12))
        I_FIRE_ISI ws.Range(ADM_ADRES(satirNo + 1, 5, r - 1, 5))
        ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 13)), True
        r = r + 2
        ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "KRÝTÝK MALZEMELER   •   " & ilMetin & "   •   " & I_AY_ADI(ayCur)
        r = r + 1: satirNo = r
    End If
    '--- kritik malzemeler
    Set dK = I_KAT_OZET(aMal)
    ADM_BASLIKLAR ws, r, 2, Array("Kategori", "", "Çeþit", "Demontaj oraný", "", "Kullanýlan", "Girilen", "Eksik (fire)", "Fire deðeri", "", "Fireli ilçe", "Durum")
    ws.Range(ADM_ADRES(r, 2, r, 3)).Merge: ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 10, r, 11)).Merge
    kats = Array("KABLO", "DÝREK", "PANO", "HÜCRE", "TRAFO")
    For i = 0 To 4
        kat = CStr(kats(i))
        r = r + 1
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
        ws.Range(ADM_ADRES(r, 2, r, 3)).Merge: ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 10, r, 11)).Merge
        ws.Cells(r, 2).Value2 = ChrW(9632) & "  " & kat
        ws.Cells(r, 2).Font.bold = True: ws.Cells(r, 2).Font.Color = I_KAT_RENK(kat): ws.Cells(r, 2).HorizontalAlignment = xlLeft
        If dK.Exists(kat) Then
            o = dK(kat)
            ws.Cells(r, 4).Value2 = o(0)
            ws.Cells(r, 5).Value2 = Round(I_UYUM_KAT(o), 1)
            If o(8) Then
                ws.Cells(r, 7).Value2 = o(1): ws.Cells(r, 8).Value2 = o(2): ws.Cells(r, 9).Value2 = o(3)
                For j = 7 To 9
                    Y_MIKTAR_BICIM ws.Cells(r, j), CStr(o(7))
                Next j
            Else
                ws.Cells(r, 7).Value2 = "farklý birim"
            End If
            ws.Cells(r, 10).Value2 = Round(o(5), 2): ws.Cells(r, 10).NumberFormat = "#,##0.00 ""TL"""
            n = 0
            For Each k In fireli.keys
                If Left$(CStr(k), 2) = i & "|" Then n = n + 1
            Next k
            ws.Cells(r, 12).Value2 = n
            If o(6) > 0 Then
                ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), "KRÝTÝK EKSÝK", RGB(254, 226, 226), RGB(185, 28, 28), 8, True, xlCenter
                ws.Cells(r, 9).Font.Color = RGB(185, 28, 28): ws.Cells(r, 9).Font.bold = True
            Else
                ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), "TAMAM", RGB(220, 252, 231), RGB(21, 128, 61), 8, True, xlCenter
            End If
        Else
            ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), "KULLANIM YOK", RGB(241, 245, 249), RGB(100, 116, 139), 8, True, xlCenter
        End If
        ws.Range(ADM_ADRES(r, 4, r, 12)).HorizontalAlignment = xlCenter
        ws.Range(ADM_ADRES(r, 10, r, 11)).ShrinkToFit = True
    Next i
    I_UYUM_CUBUK ws.Range(ADM_ADRES(r - 4, 5, r, 6))
    I_FIRE_ISI ws.Range(ADM_ADRES(r - 4, 10, r, 10))
    ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 13)), True
    '--- ilce karnesi (en iyiden en kotuye)
    r = r + 2: satirNo = r
    n = tum.Count
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "ÝLÇE KARNESÝ   •   " & ilMetin & "   •   " & I_AY_ADI(ayCur) & IIf(Len(secIl) = 0 And n > 15, "   (en iyi 5 ve en kötü 10)", "   (en iyiden en kötüye)")
    r = r + 1
    ADM_BASLIKLAR ws, r, 2, Array("Sýra", "Ýl", "Ýlçe", "Kritik uyum", "", "Kablo eksik", "Direk eksik", "Pano eksik", "Hücre eksik", "Trafo eksik", "Fire deðeri", "Deðerlendirme")
    ws.Range(ADM_ADRES(r, 5, r, 6)).Merge
    If n > 0 Then
        ReDim anah(1 To n): ReDim idx(1 To n): ReDim kod(1 To n)
        i = 0
        For Each k In tum.keys
            t = tum(k)(3): i = i + 1: idx(i) = i: kod(i) = k
            anah(i) = I_SIRA_ANAHTAR(t(9)) & Format$(Round(t(1), 2) * 100, "0000000000000000")
        Next k
        Y_SIRALA anah, idx, 1, n
        ayrim = False
        If n > 1 Then ayrim = (tum(kod(idx(1)))(3)(9) > tum(kod(idx(n)))(3)(9) + 0.05) Or (tum(kod(idx(n)))(3)(1) > tum(kod(idx(1)))(3)(1) + 0.005)
        Set sec = New Collection
        If Len(secIl) = 0 And n > 15 Then
            For i = 1 To 5
                sec.Add i
            Next i
            sec.Add 0
            For i = n - 9 To n
                sec.Add i
            Next i
        Else
            For i = 1 To n
                sec.Add i
            Next i
        End If
        For Each x In sec
            r = r + 1
            ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
            ws.Range(ADM_ADRES(r, 5, r, 6)).Merge
            If x = 0 Then
                ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = RGB(244, 248, 252)
                ws.Cells(r, 4).Value2 = "…  " & (n - 15) & " ilçe  …": ws.Cells(r, 4).Font.Color = RGB(148, 163, 184)
                ws.Cells(r, 4).HorizontalAlignment = xlCenter
            Else
                ilE = tum(kod(idx(x))): t = ilE(3)
                ws.Cells(r, 2).Value2 = x: ws.Cells(r, 2).Font.bold = True
                ws.Cells(r, 3).NumberFormat = "@": ws.Cells(r, 3).Value2 = Y_IL_ADI(CStr(ilE(0)))
                ws.Cells(r, 4).NumberFormat = "@": ws.Cells(r, 4).Value2 = IIf(Len(ilE(2)) > 0, ilE(2), ilE(1)): ws.Cells(r, 4).Font.bold = True
                ws.Cells(r, 5).Value2 = Round(t(9), 1)
                For j = 0 To 4
                    ws.Cells(r, 7 + j).Value2 = t(4 + j)
                    If t(4 + j) > 0.0005 Then ws.Cells(r, 7 + j).Font.Color = RGB(185, 28, 28): ws.Cells(r, 7 + j).Font.bold = True
                Next j
                ws.Cells(r, 12).Value2 = Round(t(1), 2)
                If x = 1 And ayrim Then
                    drm = ChrW(9733) & " EN ÝYÝ"
                ElseIf x = n And ayrim Then
                    drm = ChrW(9660) & " EN KÖTÜ"
                Else
                    drm = I_DEGERLENDIR(t(9))
                End If
                Select Case drm
                    Case ChrW(9733) & " EN ÝYÝ": ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(22, 163, 74), vbWhite, 8, True, xlCenter
                    Case ChrW(9660) & " EN KÖTÜ": ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(220, 38, 38), vbWhite, 8, True, xlCenter
                    Case "ÇOK ÝYÝ", "ÝYÝ": ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(220, 252, 231), RGB(21, 128, 61), 8, True, xlCenter
                    Case "ORTA": ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(254, 243, 199), RGB(180, 83, 9), 8, True, xlCenter
                    Case Else: ADM_BLOK ws, ADM_ADRES(r, 13, r, 13), drm, RGB(254, 226, 226), RGB(185, 28, 28), 8, True, xlCenter
                End Select
                If x = 1 And ayrim Then ws.Range(ADM_ADRES(r, 2, r, 12)).Interior.Color = RGB(240, 253, 244)
                If x = n And ayrim Then ws.Range(ADM_ADRES(r, 2, r, 12)).Interior.Color = RGB(254, 242, 242)
            End If
        Next x
        ws.Range(ADM_ADRES(satirNo + 2, 2, r, 3)).HorizontalAlignment = xlCenter
        ws.Range(ADM_ADRES(satirNo + 2, 7, r, 11)).HorizontalAlignment = xlCenter
        ws.Range(ADM_ADRES(satirNo + 2, 7, r, 7)).NumberFormat = "#,##0.## ""m"";-#,##0.## ""m"";""–"""
        ws.Range(ADM_ADRES(satirNo + 2, 8, r, 11)).NumberFormat = "#,##0.## ""adet"";-#,##0.## ""adet"";""–"""
        ws.Range(ADM_ADRES(satirNo + 2, 12, r, 12)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ADM_ADRES(satirNo + 2, 12, r, 12)).ShrinkToFit = True
        I_UYUM_CUBUK ws.Range(ADM_ADRES(satirNo + 2, 5, r, 6))
        I_FIRE_ISI ws.Range(ADM_ADRES(satirNo + 2, 12, r, 12))
    Else
        r = r + 1
        ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Bu ay ilçe depolarýnda kritik malzeme kullanýmý yok.", vbWhite, RGB(100, 116, 139), 10, False, xlLeft
    End If
    ADM_KENAR ws.Range(ADM_ADRES(satirNo + 1, 2, r, 13)), True
    r = r + 1
    ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Kritik uyum: kablo, direk, pano, hücre ve trafo kullaným deðerinin demontajý girilen oraný. Fire: kullanýlýp demontajý girilmeyen miktar. " & _
        "Ayrýntý: Ýlçe Yönetim birimi > AYLIK DEMONTAJ.", RGB(244, 248, 252), RGB(100, 116, 139), 8, False, xlLeft
    ws.Cells(r, 2).Font.Italic = True
End Sub

'---------- STOK KARSILASTIRMA (ilce) ----------
' Sayfa malzeme stok biriminin STOK_KARÞILAÞTIRMA sayfasinin kopyasidir (ayni gorunum).
' Il dugmeleri yerine ilce depo dugmeleri; durumlar AZALDI / ARTTI / DEMONTAJ GÝRÝÞ.
' Demontaj kontrolu: azalan her malzeme icin ayni depo + uretim yeri + malzemede
' CIKMA BEDELSIZ degerlemeli stok artisi aranir. Yoksa DEMONTAJ GÝRÝÞÝ YAPINIZ (kirmizi).

Private Function I_DEMONTAJ_MI(ByVal degerleme As String) As Boolean
    Dim t As String
    t = Y_NORM(degerleme)
    t = Replace(Replace(Replace(Replace(Replace(t, " ", ""), "_", ""), "-", ""), ".", ""), "/", "")
    ' SAP degerleme adi: CIKMA/BDSZ (cikma bedelsiz).
    I_DEMONTAJ_MI = (InStr(t, "CIKMA") > 0 And (InStr(t, "BDSZ") > 0 Or InStr(t, "BEDELSIZ") > 0))
End Function

' Kopya sayfa ilk kurulumda ilce birimine uyarlanir.
Private Sub I_KARS_HAZIRLA(ByVal ws As Worksheet)
    Dim i As Long, s As Shape, adlar As Variant, metinler As Variant, makrolar As Variant, son As Long
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    For i = ws.Shapes.Count To 1 Step -1
        Set s = ws.Shapes(i)
        If Left$(s.Name, 7) = "AKSKV1_" Or s.Name = "AKSORA_GECMIS_KAYIT_NOTU" Then s.Delete
    Next i
    son = application.Max(400, ws.UsedRange.Row + ws.UsedRange.Rows.Count)
    ws.Range("A16:N" & son).Clear
    ws.Range("X1:AD" & son).ClearContents
    ws.Range("B6").MergeArea.ClearContents: ws.Range("I6").MergeArea.ClearContents
    ws.Range("B7").MergeArea.ClearContents: ws.Range("I7").MergeArea.ClearContents
    adlar = Array("btnKarsTum", "btnKarsAzalan", "btnKarsArtan", "btnKarsYeni", "btnKarsSifir", "btnKarsKarsilastir")
    metinler = Array("TÜMÜ", "AZALAN", "ARTAN", "DEMONTAJ GÝRÝÞ", "DEMONTAJ YAPINIZ", "KARÞILAÞTIR")
    makrolar = Array("AKSORA_ILCE_KF_TUMU", "AKSORA_ILCE_KF_AZALAN", "AKSORA_ILCE_KF_ARTAN", _
        "AKSORA_ILCE_KF_DEMONTAJ", "AKSORA_ILCE_KF_EKSIK", "AKSORA_ILCE_KARS_YENILE")
    For i = 0 To 5
        Set s = Y_SEKIL(ws, CStr(adlar(i)))
        If Not s Is Nothing Then
            Y_YAZ s, CStr(metinler(i))
            s.OnAction = BIRIM_MAKRO(CStr(makrolar(i)))
            If i = 4 Then
                s.Fill.ForeColor.RGB = RGB(185, 28, 28)
                On Error Resume Next
                s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                On Error GoTo 0
            End If
        End If
    Next i
    ws.Range("G9").Value2 = "DEMONTAJ GÝRÝÞ"
    ws.Range("J9").Value2 = "DEMONTAJ GÝRÝÞÝ YAPINIZ"
    ws.Range("O6").Value2 = "KARÞILAÞTIRILACAK ÝLÇE"
    ws.Range("F15").Value2 = "DEÐERLEME"
    ' Demontaj kontrolu sutunu (N): tablo basligi M ile ayni bicimde.
    ws.Columns("N").ColumnWidth = 34
    With ws.Range("N15")
        .Interior.Color = ws.Range("M15").Interior.Color
        .Font.Name = ws.Range("M15").Font.Name: .Font.Size = ws.Range("M15").Font.Size
        .Font.bold = ws.Range("M15").Font.bold: .Font.Color = ws.Range("M15").Font.Color
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Borders.LineStyle = ws.Range("M15").Borders(xlEdgeBottom).LineStyle
    End With
    ws.Range("N15").Value2 = "DEMONTAJ KONTROLÜ"
    ws.Range("BX25").Value2 = "IK_V2"
End Sub

' Tablo tek parca: baslik, tarih ve ozet bantlari demontaj sutununu (N) da kapsar;
' filtre dugmeleri A:N boyunca esit, ANA PANEL sagda.
Private Sub I_KARS_DUZEN(ByVal ws As Worksheet)
    Dim r As Long, b As Variant, s As Shape, i As Long, gen As Double, da As Boolean, adlar As Variant
    da = application.DisplayAlerts: application.DisplayAlerts = False
    ws.Columns("N").ColumnWidth = 30
    For r = 1 To 14
        If ws.Cells(r, 13).Interior.Pattern = xlNone Then
            ws.Cells(r, 14).Interior.Pattern = xlNone
        Else
            ws.Cells(r, 14).Interior.Color = ws.Cells(r, 13).Interior.Color
        End If
    Next r
    For Each b In Array("A1:M2", "A3:M3", "H5:M5", "I6:M6", "I7:M7", "J9:M9", "J10:M10", "J11:M11", "J12:M12")
        ws.Range(CStr(b)).UnMerge
        ws.Range(Replace(CStr(b), "M", "N")).Merge
    Next b
    adlar = Array("btnKarsTum", "btnKarsAzalan", "btnKarsArtan", "btnKarsYeni", "btnKarsSifir", "btnKarsKarsilastir")
    gen = (ws.Range("A1:N1").Width - 5 * 2) / 6
    For i = 0 To 5
        Set s = Y_SEKIL(ws, CStr(adlar(i)))
        If Not s Is Nothing Then
            s.Placement = xlFreeFloating
            s.Left = i * (gen + 2): s.Width = gen
        End If
    Next i
    Set s = Y_SEKIL(ws, "btnAnaPanelDon")
    If Not s Is Nothing Then s.Placement = xlFreeFloating: s.Left = ws.Columns("O").Left + 12
    application.DisplayAlerts = da
    ws.Range("BX25").Value2 = "IK_V3"
End Sub

'' Tablo surum 4: DEPO'nun yaninda ILCE sutunu; BIRIM sutunu kalkar (birim miktarlarin yaninda yazilir).
Private Sub I_KARS_SUTUN(ByVal ws As Worksheet)
    Dim gen As Variant, i As Long, s As Shape, adlar As Variant, w As Double
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    gen = Array(16.8, 7.8, 8.8, 15.8, 12.8, 42.8, 12.8, 13.8, 13.8, 11.8, 15.8, 15.8, 15.8, 32.8)
    For i = 0 To 13
        ws.Columns(i + 1).ColumnWidth = gen(i)
    Next i
    ws.Range("A15:N15").Value2 = Array("DURUM", "ÜY", "DEPO", "ÝLÇE", "MALZEME", "MALZEME AÇIKLAMASI", "DEÐERLEME", _
        "ESKÝ MÝKTAR", "YENÝ MÝKTAR", "FARK", "ESKÝ DEÐER", "YENÝ DEÐER", "DEÐER FARKI", "DEMONTAJ KONTROLÜ")
    ws.Range("A15:N15").HorizontalAlignment = xlCenter
    adlar = Array("btnKarsTum", "btnKarsAzalan", "btnKarsArtan", "btnKarsYeni", "btnKarsSifir", "btnKarsKarsilastir")
    w = (ws.Range("A1:N1").Width - 5 * 2) / 6
    For i = 0 To 5
        Set s = Y_SEKIL(ws, CStr(adlar(i)))
        If Not s Is Nothing Then s.Placement = xlFreeFloating: s.Left = i * (w + 2): s.Width = w
    Next i
    Set s = Y_SEKIL(ws, "btnAnaPanelDon")
    If Not s Is Nothing Then s.Placement = xlFreeFloating: s.Left = ws.Columns("O").Left + 12
    ws.Range("BX25").Value2 = "IK_V4"
End Sub

' Ilce secimi: tek acilir liste (ilce sayisi arttikca sayfa asagi uzamaz).
Private Sub I_KARS_SECICI(ByVal ws As Worksheet, ByVal depolar As Object, ByVal sec As String)
    Dim i As Long, s As Shape, k As Variant, n As Long, secIx As Long, a As Shape
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 7) = "I_KV_D_" Or ws.Shapes(i).Name = "I_KV_SECIM" Then ws.Shapes(i).Delete
    Next i
    Set s = ws.Shapes.AddFormControl(xlDropDown, ws.Range("O8").Left + 2, ws.Range("O8").Top + 6, ws.Range("O8:P8").Width - 4, 24)
    s.Name = "I_KV_SECIM": s.Placement = xlFreeFloating
    s.ControlFormat.AddItem "TÜM ÝLÇELER  (" & depolar.Count & " depo)"
    secIx = 1: n = 1
    For Each k In depolar.keys
        n = n + 1
        s.ControlFormat.AddItem CStr(k) & "  -  " & I_DEPO_KISA(CStr(depolar(k)))
        If UCase$(CStr(k)) = sec Then secIx = n
    Next k
    s.ControlFormat.DropDownLines = application.Min(30, n)
    s.ControlFormat.Value = secIx
    s.OnAction = BIRIM_MAKRO("AKSORA_ILCE_KARS_LISTE")
    Set a = Y_SEKIL(ws, "AKSORA_KARS_SABIT_ACIKLAMA")
    If Not a Is Nothing Then
        a.Left = s.Left: a.Width = s.Width: a.Top = s.Top + s.Height + 6: a.Height = 30
        Y_YAZ a, "Listeden ilçe deposunu seçin; karþýlaþtýrma seçilen depo için yapýlýr."
        a.TextFrame2.WordWrap = msoTrue
    End If
End Sub

Public Sub AKSORA_ILCE_KARS_LISTE()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  STOK KARÞILAÞTIRMA", "Karþýlaþtýrma hazýrlanýyor...")
    On Error GoTo Son
    ILCE_KARS_LISTE_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub ILCE_KARS_LISTE_ASIL()
    Dim ws As Worksheet, s As Shape, k As String, ev As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("I_STOK_KARÞILAÞTIRMA")
    Set s = ws.Shapes("I_KV_SECIM")
    If s.ControlFormat.Value > 1 Then k = Trim$(Split(CStr(s.ControlFormat.List(s.ControlFormat.Value)) & "  -  ", "  -  ")(0))
    ev = application.EnableEvents
    application.ScreenUpdating = False: application.EnableEvents = False
    ws.Unprotect ""
    ws.Range("BX30").Value2 = k
    I_KARS_KUR CStr(ws.Range("BX22").Value2), False
    application.EnableEvents = ev: application.ScreenUpdating = True
    Y_TEMINAT_KAL ws
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub I_KARS_KUR(ByVal il As String, ByVal sifirla As Boolean)
    Dim ws As Worksheet, tarihler As Collection, nT As Long, i As Long, j As Long, eskiT As String, yeniT As String
    Dim dTum As Object
    Dim dE As Object, dY As Object, anahtar As Object, grup As Object, depolar As Object, k As Variant, x As Variant, p() As String
    Dim sec As String, gk As String, t As Variant, n As Long, son As Long, r As Long
    Dim mEski As Double, mYeni As Double, vEski As Double, vYeni As Double, tE As Double, tY As Double
    Dim blok() As Variant, anah() As String, idx() As Long, sonuc() As Variant, sira As Long, g As Variant
    Dim nAz As Long, nAr As Long, nDm As Long, nEk As Long, durum As String, kontrol As String, demontaj As Boolean
    Dim birimler() As String, isgL As Object
    Set ws = ThisWorkbook.Worksheets("I_STOK_KARÞILAÞTIRMA")
    Set isgL = I_ISG_LISTE()
    il = UCase$(Trim$(il))
    ws.Unprotect ""
    If Left$(CStr(ws.Range("BX25").Value2), 3) <> "IK_" Or CStr(ws.Range("BX25").Value2) = "IK_V1" Then I_KARS_HAZIRLA ws
    If CStr(ws.Range("BX25").Value2) = "IK_V2" Then I_KARS_DUZEN ws
    If CStr(ws.Range("BX25").Value2) = "IK_V3" Then I_KARS_SUTUN ws
    If CStr(ws.Range("BX22").Value2) <> il Then sifirla = True: ws.Range("BX30").ClearContents
    If ws.FilterMode Then ws.ShowAllData
    ws.AutoFilterMode = False
    ' Gunler (en yeni ustte): X sutunu tarih listesi
    Set tarihler = I_GECMIS_TARIHLER(il)
    nT = tarihler.Count
    ws.Range("X1:X200").ClearContents
    ws.Range("X1:X200").NumberFormat = "@"
    ws.Range("X1").Value2 = "GÖRÜNEN"
    For i = 1 To nT
        ws.Cells(i + 1, 24).Value2 = Format$(tarihler(i), I_TF)
    Next i
    eskiT = Trim$(CStr(ws.Range("B6").Value2)): yeniT = Trim$(CStr(ws.Range("I6").Value2))
    If Not sifirla Then
        If Not I_LISTEDE(tarihler, eskiT) Or Not I_LISTEDE(tarihler, yeniT) Then sifirla = True
    End If
    If sifirla Then
        eskiT = vbNullString: yeniT = vbNullString
        If nT > 0 Then
            ' Varsayilan: yeni = son guncelleme, eski = bir onceki gunun son guncellemesi.
            yeniT = Format$(tarihler(1), I_TF)
            eskiT = yeniT
            For i = 2 To nT
                If Int(CDbl(tarihler(i))) < Int(CDbl(tarihler(1))) Then eskiT = Format$(tarihler(i), I_TF): Exit For
            Next i
        End If
    End If
    For Each x In Array("B6", "I6")
        With ws.Range(CStr(x))
            .MergeArea.NumberFormat = "@"
            ' Tarih secim kutulari koruma altinda da degistirilebilir.
            .MergeArea.Locked = False
            .Validation.Delete
            If nT > 0 Then
                .Validation.Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=$X$2:$X$" & (nT + 1)
                .Validation.IgnoreBlank = True: .Validation.InCellDropdown = True
            End If
        End With
    Next x
    ws.Range("B6").Value2 = eskiT: ws.Range("I6").Value2 = yeniT
    ws.Range("BX34").Value2 = "B6": ws.Range("BX35").Value2 = "I6"
    Set dE = I_GECMIS_OKU(il, eskiT)
    Set dY = I_GECMIS_OKU(il, yeniT)
    ' Ilce depolari (iki gunun birlesimi)
    Set depolar = CreateObject("Scripting.Dictionary"): depolar.CompareMode = vbTextCompare
    For Each x In Array(dY, dE)
        For Each k In x.keys
            p = Split(CStr(k), "|")
            If Not depolar.Exists(p(0)) Then depolar(p(0)) = x(k)(0)
            If Len(CStr(depolar(p(0)))) = 0 Then depolar(p(0)) = x(k)(0)
        Next k
    Next x
    Set dTum = I_DEPOLAR(il)
    For Each k In dTum.keys
        If Not depolar.Exists(CStr(k)) Then depolar(CStr(k)) = dTum(k)
    Next k
    I_ANAHTAR_SIRALA depolar
    sec = UCase$(Trim$(CStr(ws.Range("BX30").Value2)))
    If Len(sec) > 0 And Not depolar.Exists(sec) Then sec = vbNullString
    ws.Range("BX30").Value2 = sec
    I_KARS_SECICI ws, depolar, sec
    ' Kalemler ve demontaj gruplari (depo|uretim yeri|malzeme)
    Set anahtar = CreateObject("Scripting.Dictionary"): anahtar.CompareMode = vbTextCompare
    Set grup = CreateObject("Scripting.Dictionary"): grup.CompareMode = vbTextCompare
    For Each x In Array(dE, dY)
        For Each k In x.keys
            If Len(sec) = 0 Or UCase$(Split(CStr(k), "|")(0)) = sec Then anahtar(k) = True
        Next k
    Next x
    For Each k In anahtar.keys
        mEski = 0: mYeni = 0: vEski = 0: vYeni = 0
        If dE.Exists(k) Then t = dE(k): mEski = t(3): vEski = t(4)
        If dY.Exists(k) Then t = dY(k): mYeni = t(3): vYeni = t(4)
        tE = tE + vEski: tY = tY + vYeni
        p = Split(CStr(k), "|")
        gk = p(0) & "|" & p(1) & "|" & p(2)
        If Not grup.Exists(gk) Then grup(gk) = Array(0#, 0#)
        g = grup(gk)
        If I_DEMONTAJ_MI(p(3)) Then
            If mYeni > mEski Then g(1) = g(1) + (mYeni - mEski)
        ElseIf mYeni < mEski Then
            g(0) = g(0) + (mEski - mYeni)
        End If
        grup(gk) = g
    Next k
    n = 0
    If anahtar.Count > 0 Then
        ReDim blok(1 To anahtar.Count, 1 To 15): ReDim anah(1 To anahtar.Count): ReDim idx(1 To anahtar.Count)
    End If
    For Each k In anahtar.keys
        mEski = 0: mYeni = 0: vEski = 0: vYeni = 0
        If dE.Exists(k) Then t = dE(k): mEski = t(3): vEski = t(4)
        If dY.Exists(k) Then t = dY(k): mYeni = t(3): vYeni = t(4)
        If Abs(mYeni - mEski) > 0.0005 Then
            p = Split(CStr(k), "|")
            demontaj = I_DEMONTAJ_MI(p(3))
            kontrol = vbNullString
            If mYeni > mEski Then
                If demontaj Then
                    durum = "DEMONTAJ GÝRÝÞ": nDm = nDm + 1: sira = 5: kontrol = "Demontaj giriþi"
                Else
                    durum = "ARTTI": nAr = nAr + 1: sira = 4
                End If
            Else
                durum = "AZALDI": nAz = nAz + 1: sira = 3
                ' ISG malzemeleri ile sayac / modem / aparat-muhur gruplari demontaj kontrolune girmez.
                kontrol = I_DEM_DISI(p(2), isgL)
                If Len(kontrol) > 0 Then
                    sira = 3
                ElseIf Not demontaj Then
                    g = grup(p(0) & "|" & p(1) & "|" & p(2))
                    If g(1) <= 0.0005 Then
                        kontrol = "DEMONTAJ GÝRÝÞÝ YAPINIZ  •  " & I_MIKTAR_METNI(mEski - mYeni) & " " & CStr(t(2))
                        sira = 1: nEk = nEk + 1
                    ElseIf g(1) + 0.0005 < g(0) Then
                        kontrol = "EKSÝK DEMONTAJ  •  " & I_MIKTAR_METNI(g(0) - g(1)) & " " & CStr(t(2)) & " eksik (" & _
                            I_MIKTAR_METNI(g(1)) & " / " & I_MIKTAR_METNI(g(0)) & ")"
                        sira = 2: nEk = nEk + 1
                    Else
                        kontrol = "Demontaj girildi"
                    End If
                End If
            End If
            n = n + 1
            ' Sutunlar: durum, UY, depo, ilce, malzeme, tanim, degerleme, eski/yeni/fark miktar (birimiyle),
            ' eski/yeni/fark deger, demontaj kontrolu. Birim ayri sutun yerine miktarlarin yaninda yazilir.
            blok(n, 1) = durum: blok(n, 2) = p(1): blok(n, 3) = p(0)
            blok(n, 4) = I_DEPO_KISA(CStr(depolar(p(0)))): blok(n, 5) = p(2)
            blok(n, 6) = t(1): blok(n, 7) = p(3)
            blok(n, 8) = mEski: blok(n, 9) = mYeni: blok(n, 10) = mYeni - mEski
            blok(n, 11) = Round(vEski, 2): blok(n, 12) = Round(vYeni, 2): blok(n, 13) = Round(vYeni - vEski, 2)
            blok(n, 14) = kontrol: blok(n, 15) = CStr(t(2))
            anah(n) = Left$(UCase$(CStr(blok(n, 4))) & Space$(40), 40) & CStr(sira) & Format$(1E+15 - Abs(vYeni - vEski) * 100, "0000000000000000")
            idx(n) = n
        End If
    Next k
    ' Tablo
    son = application.Max(16, ws.Cells(ws.Rows.Count, 4).End(xlUp).Row, ws.Cells(ws.Rows.Count, 14).End(xlUp).Row)
    ws.Range("A16:N" & son + 1).Clear
    If n > 0 Then
        Y_SIRALA anah, idx, 1, n
        ReDim sonuc(1 To n, 1 To 14): ReDim birimler(1 To n)
        For i = 1 To n
            For j = 1 To 14: sonuc(i, j) = blok(idx(i), j): Next j
            birimler(i) = CStr(blok(idx(i), 15))
        Next i
        ws.Range("B16:G" & n + 15).NumberFormat = "@"
        ws.Range("A16").Resize(n, 14).Value2 = sonuc
        With ws.Range("A16:N" & n + 15)
            .Interior.Color = RGB(255, 255, 255)
            .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(225, 230, 236): .Borders.Weight = xlThin
            .VerticalAlignment = xlCenter
            .Font.Name = "Aptos": .Font.Size = 10
        End With
        ws.Range("A16:E" & n + 15).HorizontalAlignment = xlCenter
        ws.Range("D16:D" & n + 15).HorizontalAlignment = xlLeft: ws.Range("D16:D" & n + 15).IndentLevel = 1
        ws.Range("G16:G" & n + 15).HorizontalAlignment = xlCenter
        ws.Range("H16:J" & n + 15).HorizontalAlignment = xlRight
        ws.Range("K16:M" & n + 15).NumberFormat = "#,##0.00 ""TL"";[Red]-#,##0.00 ""TL"""
        ws.Range("A16:A" & n + 15).Font.bold = True
        ws.Range("N16:N" & n + 15).HorizontalAlignment = xlCenter
        For r = 16 To n + 15
            For j = 8 To 10
                Y_MIKTAR_BICIM ws.Cells(r, j), birimler(r - 15)
            Next j
            Select Case CStr(ws.Cells(r, 1).Value2)
                Case "AZALDI": ws.Cells(r, 1).Font.Color = RGB(192, 0, 0)
                Case "ARTTI": ws.Cells(r, 1).Font.Color = RGB(0, 140, 75)
                Case Else: ws.Cells(r, 1).Font.Color = RGB(29, 78, 216)
            End Select
            kontrol = CStr(ws.Cells(r, 14).Value2)
            If Left$(kontrol, 8) = "DEMONTAJ" And InStr(kontrol, "YAPINIZ") > 0 Then
                ws.Cells(r, 14).Interior.Color = RGB(220, 38, 38): ws.Cells(r, 14).Font.Color = RGB(255, 255, 255): ws.Cells(r, 14).Font.bold = True
            ElseIf Left$(kontrol, 5) = "EKSÝK" Then
                ws.Cells(r, 14).Interior.Color = RGB(254, 243, 199): ws.Cells(r, 14).Font.Color = RGB(180, 83, 9): ws.Cells(r, 14).Font.bold = True
            ElseIf kontrol = "Demontaj girildi" Then
                ws.Cells(r, 14).Font.Color = RGB(21, 128, 61)
            ElseIf kontrol = "Demontaj giriþi" Then
                ws.Cells(r, 14).Font.Color = RGB(29, 78, 216)
            ElseIf InStr(kontrol, "gerekmez") > 0 Or InStr(kontrol, "kontrolü dýþý") > 0 Then
                ws.Cells(r, 14).Font.Color = RGB(100, 116, 139): ws.Cells(r, 14).Font.Italic = True
            End If
        Next r
        ws.Range("A15:N" & n + 15).AutoFilter
    Else
        ws.Range("A16").Value2 = IIf(nT < 2, "Karþýlaþtýrma için bu ilde en az iki farklý günde ilçe stok güncellemesi yapýlmalýdýr" & _
            IIf(nT = 1, " (mevcut kayýt: " & yeniT & ").", "."), "Seçilen tarihler arasýnda stok deðiþikliði yok.")
        ws.Range("A16").Font.Italic = True: ws.Range("A16").Font.Color = RGB(100, 116, 139)
    End If
    ' Ozet kutulari
    ws.Range("A10").Value = nAz: ws.Range("D10").Value = nAr: ws.Range("G10").Value = nDm: ws.Range("J10").Value = nEk
    ws.Range("J10").Font.Color = IIf(nEk > 0, RGB(192, 0, 0), RGB(24, 55, 88))
    ws.Range("A12").Value = Round(tE, 2): ws.Range("D12").Value = Round(tY, 2)
    ws.Range("G12").Value = Round(tY - tE, 2): ws.Range("J12").Value = n
    ws.Range("A12,D12,G12").NumberFormat = "#,##0.00 ""TL"""
    ws.Range("G12").Font.Color = IIf(tY - tE < 0, RGB(192, 0, 0), IIf(tY - tE > 0, RGB(0, 140, 75), RGB(24, 55, 88)))
    ws.Range("B7").Value2 = IIf(Len(eskiT) > 0, "Son karþýlaþtýrýlan: " & eskiT, "")
    ws.Range("I7").Value2 = IIf(Len(yeniT) > 0, "Son karþýlaþtýrýlan: " & yeniT, "")
    ws.Range("A1").Value2 = "STOK / DEMONTAJ KONTROL MERKEZÝ"
    Y_FILIGRAN ws
    ws.Range("A3").Value2 = UCase$(Y_IL_ADI(il)) & " | CDP " & AKSORA_IL_TESIS(il, "CDP") & " + EPP " & AKSORA_IL_TESIS(il, "EPP") & _
        " | Ýlçe Depolarý (K001-K999)" & IIf(Len(sec) > 0, " | " & sec & " " & I_DEPO_KISA(CStr(depolar(sec))), "")
    ws.Range("BX22").Value2 = il: ws.Range("BX20").Value2 = il
    I_BAGLA ws
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Function I_MIKTAR_METNI(ByVal x As Double) As String
    If x = Int(x) Then I_MIKTAR_METNI = AKSORA_SAYI_METNI(x, "#,##0") Else I_MIKTAR_METNI = AKSORA_SAYI_METNI(x, "#,##0.###")
End Function

Private Function I_LISTEDE(ByVal c As Collection, ByVal metin As String) As Boolean
    Dim x As Variant
    For Each x In c
        If Format$(x, I_TF) = metin Then I_LISTEDE = True: Exit Function
    Next x
End Function

' Hizli filtreler: TUMU, AZALDI, ARTTI, DEMONTAJ GIRIS, DEMONTAJ YAPINIZ (eksik/yok).
Private Sub I_KARS_FILTRE(ByVal tur As String)
    Dim ws As Worksheet, son As Long
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("I_STOK_KARÞILAÞTIRMA")
    ws.Unprotect ""
    If ws.FilterMode Then ws.ShowAllData
    son = ws.Cells(ws.Rows.Count, 4).End(xlUp).Row
    If son >= 16 Then
        If Not ws.AutoFilterMode Then ws.Range("A15:N" & son).AutoFilter
        Select Case tur
            Case "AZALDI", "ARTTI", "DEMONTAJ GÝRÝÞ"
                ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:=tur
            Case "EKSIK"
                ws.AutoFilter.Range.AutoFilter Field:=14, Criteria1:="=DEMONTAJ GÝRÝÞÝ YAPINIZ*", Operator:=xlOr, Criteria2:="=EKSÝK DEMONTAJ*"
        End Select
    End If
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Y_TEMINAT_KAL ws
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_ILCE_KF_TUMU(): I_KARS_FILTRE "": End Sub
Public Sub AKSORA_ILCE_KF_AZALAN(): I_KARS_FILTRE "AZALDI": End Sub
Public Sub AKSORA_ILCE_KF_ARTAN(): I_KARS_FILTRE "ARTTI": End Sub
Public Sub AKSORA_ILCE_KF_DEMONTAJ(): I_KARS_FILTRE "DEMONTAJ GÝRÝÞ": End Sub
Public Sub AKSORA_ILCE_KF_EKSIK(): I_KARS_FILTRE "EKSIK": End Sub

' Depo kodlari artan sirada (sozluk yeniden kurulur).
Private Sub I_ANAHTAR_SIRALA(ByRef d As Object)
    Dim k() As String, idx() As Long, i As Long, n As Long, x As Variant, y As Object
    n = d.Count
    If n < 2 Then Exit Sub
    ReDim k(1 To n): ReDim idx(1 To n)
    For Each x In d.keys
        i = i + 1: k(i) = CStr(x): idx(i) = i
    Next x
    Y_SIRALA k, idx, 1, n
    Set y = CreateObject("Scripting.Dictionary"): y.CompareMode = vbTextCompare
    For i = 1 To n
        y(k(i)) = d(k(i))
    Next i
    Set d = y
End Sub

Private Function I_DUGME_LISTESI(ByVal d As Object) As Collection
    Dim c As Collection, x As Variant
    Set c = New Collection
    c.Add vbNullString
    For Each x In d.keys
        c.Add CStr(x)
    Next x
    Set I_DUGME_LISTESI = c
End Function

' "ARHAVI ILCE DEPOSU" gibi adlardan kisa ilce adi.
Private Function I_DEPO_KISA(ByVal ad As String) As String
    Dim t As String, x As Variant
    t = Trim$(ad)
    For Each x In Array(" ÝLÇE DEPOSU", " ILCE DEPOSU", " ÝLÇE DEPO", " ILCE DEPO", " DEPOSU", " DEPO")
        If UCase$(Right$(t, Len(x))) = UCase$(x) Then t = Trim$(Left$(t, Len(t) - Len(x))): Exit For
    Next x
    If Len(t) = 0 Then t = ad
    I_DEPO_KISA = t
End Function

Public Sub AKSORA_ILCE_KARS_SEC()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  STOK KARÞILAÞTIRMA", "Karþýlaþtýrma hazýrlanýyor...")
    On Error GoTo Son
    ILCE_KARS_SEC_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub ILCE_KARS_SEC_ASIL()
    Dim ws As Worksheet, ev As Boolean, k As String
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("I_STOK_KARÞILAÞTIRMA")
    k = ws.Shapes(CStr(application.Caller)).AlternativeText
    ev = application.EnableEvents
    application.ScreenUpdating = False: application.EnableEvents = False
    ws.Unprotect ""
    ws.Range("BX30").Value2 = k
    I_KARS_KUR CStr(ws.Range("BX22").Value2), False
    application.EnableEvents = ev: application.ScreenUpdating = True
    Y_TEMINAT_KAL ws
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_ILCE_KARS_YENILE()
    Dim ac As Boolean, no As Long, metin As String
    ac = AKSORA_BEKLE_AC("AKSORA  |  STOK KARÞILAÞTIRMA", "Karþýlaþtýrma hazýrlanýyor...")
    On Error GoTo Son
    ILCE_KARS_YENILE_ASIL
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox metin, vbExclamation, "AKSORA"
End Sub

Private Sub ILCE_KARS_YENILE_ASIL()
    Dim ws As Worksheet, ev As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("I_STOK_KARÞILAÞTIRMA")
    If Len(CStr(ws.Range("BX34").Value2)) = 0 Then Exit Sub
    ev = application.EnableEvents
    application.ScreenUpdating = False: application.EnableEvents = False
    ws.Unprotect ""
    ws.Range("BX32").Value2 = Trim$(CStr(ws.Range(CStr(ws.Range("BX34").Value2)).Value2))
    ws.Range("BX33").Value2 = Trim$(CStr(ws.Range(CStr(ws.Range("BX35").Value2)).Value2))
    I_KARS_KUR CStr(ws.Range("BX22").Value2), False
    application.EnableEvents = ev: application.ScreenUpdating = True
    Y_TEMINAT_KAL ws
    Exit Sub
Hata:
    application.EnableEvents = True: application.ScreenUpdating = True
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

' Ilce ISG malzemeleri (docs/isg/ISG_ISLEM.csv): kod + Z (kisisel zimmet) / M (manuel sarf).
Private Function I_ISG_LISTE() As Object
    Static d As Object
    Dim s As String, x As Variant
    If Not d Is Nothing Then Set I_ISG_LISTE = d: Exit Function
    s = s & "7500050332Z;7500051142Z;7500051144Z;7500050334Z;7500051145Z;7500050240Z;7500000038Z;7500050254Z;7500051153Z;7500051152Z;7500050908Z;7500051135Z;7500000042Z;7500000233Z;"
    s = s & "7500000067Z;7500050900Z;7500050977Z;7500050266Z;7500000060Z;7500050335Z;7500000253Z;7500001730Z;7500000709Z;7500000710Z;7500000711Z;7500000712Z;7500000713Z;7500000714Z;"
    s = s & "7500000715Z;7500000716Z;7500000717Z;7500000729Z;7500000733Z;7500000735Z;7500000737Z;7500000739Z;7500000740Z;7500000744Z;7500000747Z;7500000752Z;7500000762Z;7500000769Z;"
    s = s & "7500000775Z;7500000777Z;7500000785Z;7500000790Z;7500000833Z;7500000837Z;7500000842Z;7500000843Z;7500000844Z;7500000847Z;7500000996Z;7500001021Z;7500001040Z;7500001043Z;"
    s = s & "7500001045Z;7500001047Z;7500001049Z;7500001051Z;7500001157Z;7500000665Z;7500000666Z;7500000667Z;7500000670Z;7500000671Z;7500001583Z;7500001591Z;7500000683Z;7500000684Z;"
    s = s & "7500051420Z;7500000687Z;7500001644Z;7500000697Z;7500000699Z;7500000700Z;7500000703Z;7500050815Z;7500050816Z;7500050920Z;7500050858Z;7500050883Z;7500053995M;7500050916Z;"
    s = s & "7500001435Z;7500001312Z;7500001314Z;7500051046Z;7500050960Z;7500050910Z;7500050909Z;7500050967Z;7500050905Z;7500050787Z;7500051010Z;7500051009Z;7500053643M;7500000381Z;"
    s = s & "7500000382Z;7500050895Z;7500000473Z;7500050849Z;7500051455Z;7500050398Z;7500051423Z;7500000370Z;7500050931Z;7500050861Z;7500050860Z;7500050887Z;7500051422Z;7500050377Z;"
    s = s & "7500050876Z;7500000338Z;7500050992Z;7500050875Z;7500050934Z;7500050397Z;7500050921Z;7500050949Z;7500000504Z;7500000507Z;7500000556Z;7500000559Z;7500000496Z;7500000564Z;"
    s = s & "7500050431Z;7500000616Z;7500000617Z;7500000618Z;7500000458Z;7500050375Z;7500050961Z;7500050962Z;7500050954Z;7500051011Z;7500050938M;7500052804Z;7500053648M;7500051020Z;"
    s = s & "7500051022Z;7500051023Z;7500051026Z;7500051025Z;7500051131M;7500050843Z;7500050944Z;7500050888Z;7500050790Z;7500050404Z;7500050974Z;7500050871Z;7500050933Z;7500050870Z;"
    s = s & "7500050868Z;7500050918M;7500050857M;7500050856M;7500050940M;7500050919M;7500001504M;7500001505M;7500050356M;7500001510M;7500001618M;7500050388Z;7500050405M;7500050917Z;"
    s = s & "7500050981Z;7500001512M;7500050852Z;7500050903M;7500050391Z;7500001475Z;7500050785Z;7500050889Z;7500050929Z;7500050351M;7500001514M;7500001469M;7500000506Z;7500000558Z;"
    s = s & "7500000498Z;7500000501Z;7500000565Z;7500000568Z;7500000481Z;7500001581Z;7500001580Z;7500050913M;7500050899Z;7500050872Z;7500050348M;7500051001M;7500050345M;7500051000M;"
    s = s & "7500050924M;7500051003M;7500050923Z;7500050932Z;7500050341Z;7500050912Z;7500051136Z;7500051514Z;7500051513Z;7500050267Z;7500052777Z;7500051147Z;7500051148Z;7500051149Z;"
    s = s & "7500051150Z;7500026511Z;7500050299Z;7500050243Z;7500000006Z;7500050797Z;7500051424Z;7500051516Z;7500051517Z;7500051134Z;7500052669Z;7500052579Z;7500050926Z;7500050271Z;"
    s = s & "7500051432Z;7500050265Z;7500001993Z;7500052616Z;7500051028Z;7500050950Z;7500050874Z;7500030895Z;7500050904Z;7500050836Z;7500000080Z;7500050902Z;7500050825Z;7500051137Z;"
    s = s & "7500053641M;7500051024Z;7500051456Z;7500050988Z;7500050810Z;7500050942Z;7500051073Z;7500050823Z;7500050827Z;7500050824Z;7500050813Z;7500053644M;7500050826Z;7500000226Z;"
    s = s & "7500000228Z;7500000229Z;7500000232Z;7500000246Z;7500051014Z;7500050833Z;7500050822Z;7500050821Z;7500050820Z;7500000249Z;7500000250Z;7500050247Z;7500050835Z;7500050834Z;"
    s = s & "7500051138Z;7500050878Z;7500000012Z;7500052668Z;7500000313Z;7500051133Z;7500050809Z;7500050630M;7500050286Z;7500050293Z;7500050290Z;7500050295Z;7500050288Z;7500050268Z;"
    s = s & "7500050289Z;7500053129Z;7500050287Z;7500050294Z;7500050291Z;7500050296Z;7500050292Z;7500051452M;7500050971Z;7500050914Z;7500051130Z;7500050845Z;7500050968Z;"
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For Each x In Split(s, ";")
        If Len(x) > 1 Then d(Left$(CStr(x), Len(x) - 1)) = Right$(CStr(x), 1)
    Next x
    Set I_ISG_LISTE = d
End Function

'---------- ISG MALZEMELERI ----------
' Ilce depolarindaki ISG malzemeleri (ana liste: docs/isg). Depoda ISG malzemesi bekletilmez:
' kisisel zimmet listesindekiler icin ZÝMMET YAPIN, digerleri icin MANUEL SARF YAPIN.
' Ust bolum ilce ozeti, alt bolum malzeme listesi (ilce secimi ve islem filtreleri).

Private Sub I_ISG_KUR(ByVal il As String)
    Dim ws As Worksheet, w As Double, s As Shape, v As Variant, r As Long, i As Long, j As Long, liste As Object
    Dim mal As Object, ozet As Object, depolar As Object, k As Variant, t As Variant, mk As String, dp As String
    Dim islem As String, a1 As Long, b1 As Long, a2 As Long, b2 As Long, n As Long, nD As Long
    Dim blok() As Variant, anah() As String, idx() As Long, sonuc() As Variant, y As Double, kw As Double
    Dim zK As Long, zM As Double, sK As Long, sM As Double, toplam As Double, rr As Range, bw As Double
    Dim dd As Shape, secim As String, birim As String, o As Variant
    Set ws = ThisWorkbook.Worksheets("I_ÝSG_MALZEMELERÝ")
    il = UCase$(Trim$(il))
    ws.Unprotect ""
    secim = CStr(ws.Range("BX30").Value2)
    If CStr(ws.Range("BX22").Value2) <> il Then secim = vbNullString
    Y_SAYFA_SIFIRLA ws
    I_BAGLA ws
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    Y_KOLONLAR ws, Array(70, 190, 64, 100, 100, application.Max(120, w - 1034), 110, 130, 270)
    Set liste = I_ISG_LISTE()
    ' Ilce depo + malzeme toplamlari (CDP ve EPP birlikte, butun degerlemeler)
    Set mal = CreateObject("Scripting.Dictionary"): mal.CompareMode = vbTextCompare
    Set depolar = I_DEPOLAR(il)
    v = AKSORA_PAKET_OKU(I_ARSIV_SAYFA(il, False))
    If IsArray(v) Then
        For r = 1 To UBound(v, 1)
            mk = Y_KOD_NORM(v(r, 5))
            If liste.Exists(mk) Then
                dp = UCase$(I_TEMIZ(v(r, 2)))
                If Len(dp) > 0 And I_SAYI(v(r, 7)) > 0 Then
                    If Not depolar.Exists(dp) Then depolar(dp) = I_TEMIZ(v(r, 3))
                    k = dp & "|" & I_TEMIZ(v(r, 1)) & "|" & mk
                    If Not mal.Exists(k) Then
                        mal(k) = Array(dp, I_TEMIZ(v(r, 3)), IIf(Left$(I_TEMIZ(v(r, 1)), 2) = "70", "EPP", "CDP"), _
                            I_TEMIZ(v(r, 4)), mk, I_TEMIZ(v(r, 6)), 0#, I_TEMIZ(v(r, 8)), 0#, _
                            IIf(liste(mk) = "Z", "ZÝMMET YAPIN", "MANUEL SARF YAPIN"))
                    End If
                    t = mal(k)
                    t(6) = t(6) + I_SAYI(v(r, 7)): t(8) = t(8) + I_SAYI(v(r, 11))
                    mal(k) = t
                End If
            End If
        Next r
    End If
    I_ANAHTAR_SIRALA depolar
    ' Ilce ozeti: depo -> (ad, zimmet kalem, zimmet miktar, sarf kalem, sarf miktar, deger)
    Set ozet = CreateObject("Scripting.Dictionary"): ozet.CompareMode = vbTextCompare
    For Each k In depolar.keys
        ozet(CStr(k)) = Array(CStr(depolar(k)), 0, 0#, 0, 0#, 0#)
    Next k
    For Each k In mal.keys
        t = mal(k)
        If ozet.Exists(CStr(t(0))) Then
            o = ozet(CStr(t(0)))
            If Len(CStr(o(0))) = 0 Then o(0) = t(1)
            If t(9) = "ZÝMMET YAPIN" Then
                o(1) = o(1) + 1: o(2) = o(2) + t(6): zK = zK + 1: zM = zM + t(6)
            Else
                o(3) = o(3) + 1: o(4) = o(4) + t(6): sK = sK + 1: sM = sM + t(6)
            End If
            o(5) = o(5) + t(8): toplam = toplam + t(8)
            ozet(CStr(t(0))) = o
        End If
    Next k
    nD = ozet.Count: n = mal.Count
    a1 = 10: b1 = a1 + application.Max(nD, 1) - 1
    a2 = b1 + 7: b2 = a2 + application.Max(n, 1) - 1
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10
    ws.Rows(a1 - 2).RowHeight = 30: ws.Rows(a1 - 1).RowHeight = 30
    ws.Range(ws.Rows(a1), ws.Rows(b1 + 1)).RowHeight = 21
    ws.Rows(b1 + 2).RowHeight = 14
    ws.Rows(a2 - 4).RowHeight = 30: ws.Rows(a2 - 3).RowHeight = 32: ws.Rows(a2 - 2).RowHeight = 8: ws.Rows(a2 - 1).RowHeight = 30
    ws.Range(ws.Rows(a2), ws.Rows(b2)).RowHeight = 19
    Y_BOLUM ws, a1 - 2, "ÝLÇE BAZINDA ÝSG DURUMU", "Depoda ÝSG malzemesi bekletilmez: kiþisel zimmet ya da manuel sarf yapýlmalýdýr.", 9
    Y_BASLIK_SATIRI ws.Range(ws.Cells(a1 - 1, 1), ws.Cells(a1 - 1, 9)), Array("Depo", "Depo Adý", "Zimmet Kalem", _
        "Zimmet Miktar", "Manuel Sarf Kalem", "Manuel Sarf Miktar", "", "ÝSG Stok Deðeri", "Durum")
    r = a1
    For Each k In ozet.keys
        o = ozet(k)
        ws.Cells(r, 1).NumberFormat = "@": ws.Cells(r, 1).Value2 = CStr(k)
        ws.Cells(r, 2).Value2 = o(0)
        ws.Cells(r, 3).Value2 = o(1): ws.Cells(r, 4).Value2 = o(2)
        ws.Cells(r, 5).Value2 = o(3): ws.Cells(r, 6).Value2 = o(4)
        ws.Cells(r, 8).Value2 = Round(o(5), 2)
        If o(1) > 0 And o(3) > 0 Then
            ws.Cells(r, 9).Value2 = "ZÝMMET + SARF BEKLÝYOR"
        ElseIf o(1) > 0 Then
            ws.Cells(r, 9).Value2 = "ZÝMMET BEKLÝYOR"
        ElseIf o(3) > 0 Then
            ws.Cells(r, 9).Value2 = "SARF BEKLÝYOR"
        Else
            ws.Cells(r, 9).Value2 = "TEMÝZ"
        End If
        Y_MIKTAR_BICIM ws.Cells(r, 4): Y_MIKTAR_BICIM ws.Cells(r, 6)
        r = r + 1
    Next k
    If nD = 0 Then ws.Cells(a1, 2).Value2 = "Ýlçe stoklarý güncellenince ilçeler listelenir."
    ws.Cells(b1 + 1, 2).Value2 = "TOPLAM"
    ws.Cells(b1 + 1, 3).Value2 = zK: ws.Cells(b1 + 1, 4).Value2 = zM
    ws.Cells(b1 + 1, 5).Value2 = sK: ws.Cells(b1 + 1, 6).Value2 = sM
    ws.Cells(b1 + 1, 8).Value2 = Round(toplam, 2)
    Y_MIKTAR_BICIM ws.Cells(b1 + 1, 4): Y_MIKTAR_BICIM ws.Cells(b1 + 1, 6)
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(b1, 9))
    Y_TOPLAM_SATIRI ws.Range(ws.Cells(b1 + 1, 1), ws.Cells(b1 + 1, 9))
    ws.Range(ws.Cells(a1, 1), ws.Cells(b1 + 1, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 3), ws.Cells(b1 + 1, 6)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 9), ws.Cells(b1, 9)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 2), ws.Cells(b1, 2)).IndentLevel = 1
    ws.Range(ws.Cells(a1, 8), ws.Cells(b1 + 1, 8)).NumberFormat = "#,##0.00 ""TL"""
    Set rr = ws.Range(ws.Cells(a1, 9), ws.Cells(b1, 9))
    Y_DURUM_KF rr, "ZÝMMET + SARF BEKLÝYOR", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF rr, "ZÝMMET BEKLÝYOR", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF rr, "SARF BEKLÝYOR", RGB(219, 234, 254), RGB(29, 78, 216)
    Y_DURUM_KF rr, "TEMÝZ", RGB(220, 252, 231), RGB(21, 128, 61)
    ' Malzeme listesi: baslik + ilce secimi, islem filtreleri, tablo
    Y_BOLUM ws, a2 - 4, "ÝSG MALZEME LÝSTESÝ  –  YAPILACAK ÝÞLEM", ""
    ws.Cells(a2 - 4, 7).Value2 = "Ýlçe:"
    ws.Cells(a2 - 4, 7).HorizontalAlignment = xlRight: ws.Cells(a2 - 4, 7).Font.bold = True: ws.Cells(a2 - 4, 7).Font.Color = RGB(28, 54, 83)
    Set dd = ws.Shapes.AddFormControl(xlDropDown, ws.Cells(a2 - 4, 8).Left + 4, ws.Cells(a2 - 4, 8).Top + 4, _
        ws.Range(ws.Cells(a2 - 4, 8), ws.Cells(a2 - 4, 9)).Width - 6, 22)
    dd.Name = "I_ISG_SECIM": dd.Placement = xlFreeFloating
    dd.ControlFormat.AddItem "TÜM ÝLÇELER"
    i = 1
    For Each k In ozet.keys
        dd.ControlFormat.AddItem CStr(k) & "  -  " & I_DEPO_KISA(CStr(ozet(k)(0)))
        i = i + 1
        If CStr(k) = secim Then dd.ControlFormat.Value = i
    Next k
    If dd.ControlFormat.Value < 1 Then dd.ControlFormat.Value = 1
    dd.ControlFormat.DropDownLines = application.Min(30, i)
    dd.OnAction = BIRIM_MAKRO("AKSORA_ILCE_ISG_SEC")
    bw = (w - 2 * 8) / 3
    y = ws.Rows(a2 - 3).Top + 2
    Set s = Y_DUGME(ws, "I_ISG_F_TUMU", "TÜMÜ  (" & n & " kalem)", 0, y, bw, 28, RGB(45, 95, 145), "AKSORA_ILCE_ISG_TUMU")
    Set s = Y_DUGME(ws, "I_ISG_F_ZIMMET", "ZÝMMET YAPIN  (" & zK & " kalem)", bw + 8, y, bw, 28, RGB(185, 28, 28), "AKSORA_ILCE_ISG_ZIMMET")
    Set s = Y_DUGME(ws, "I_ISG_F_SARF", "MANUEL SARF YAPIN  (" & sK & " kalem)", 2 * (bw + 8), y, bw, 28, RGB(29, 78, 216), "AKSORA_ILCE_ISG_SARF")
    Y_BASLIK_SATIRI ws.Range(ws.Cells(a2 - 1, 1), ws.Cells(a2 - 1, 9)), Array("Depo", "Depo Adý", "Sistem", "Mal Grubu", _
        "Malzeme", "Malzeme Tanýmý", "Miktar", "Stok Deðeri", "YAPILACAK ÝÞLEM")
    If n > 0 Then
        ReDim blok(1 To n, 1 To 10): ReDim anah(1 To n): ReDim idx(1 To n)
        i = 0
        For Each k In mal.keys
            t = mal(k): i = i + 1: idx(i) = i
            blok(i, 1) = t(0): blok(i, 2) = t(1): blok(i, 3) = t(2): blok(i, 4) = t(3)
            blok(i, 5) = t(4): blok(i, 6) = t(5): blok(i, 7) = t(6): blok(i, 8) = Round(t(8), 2): blok(i, 9) = t(9): blok(i, 10) = t(7)
            anah(i) = CStr(t(0)) & "|" & IIf(t(9) = "ZÝMMET YAPIN", "1", "2") & "|" & Y_NORM(CStr(t(5))) & "|" & CStr(t(2))
        Next k
        Y_SIRALA anah, idx, 1, n
        ReDim sonuc(1 To n, 1 To 9)
        For i = 1 To n
            For j = 1 To 9: sonuc(i, j) = blok(idx(i), j): Next j
        Next i
        ws.Range(ws.Cells(a2, 1), ws.Cells(b2, 5)).NumberFormat = "@"
        ws.Range(ws.Cells(a2, 1), ws.Cells(b2, 9)).Value2 = sonuc
        For i = 1 To n
            birim = CStr(blok(idx(i), 10))
            Y_MIKTAR_BICIM ws.Cells(a2 + i - 1, 7), birim
        Next i
        Y_TABLO_GOVDE ws.Range(ws.Cells(a2, 1), ws.Cells(b2, 9))
        ws.Range(ws.Cells(a2, 1), ws.Cells(b2, 1)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a2, 3), ws.Cells(b2, 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a2, 7), ws.Cells(b2, 7)).HorizontalAlignment = xlRight
        ws.Range(ws.Cells(a2, 9), ws.Cells(b2, 9)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a2, 2), ws.Cells(b2, 2)).IndentLevel = 1
        ws.Range(ws.Cells(a2, 6), ws.Cells(b2, 6)).IndentLevel = 1
        ws.Range(ws.Cells(a2, 8), ws.Cells(b2, 8)).NumberFormat = "#,##0.00 ""TL"""
        Set rr = ws.Range(ws.Cells(a2, 9), ws.Cells(b2, 9))
        Y_DURUM_KF rr, "ZÝMMET YAPIN", RGB(220, 38, 38), RGB(255, 255, 255)
        Y_DURUM_KF rr, "MANUEL SARF YAPIN", RGB(29, 78, 216), RGB(255, 255, 255)
        ws.Range(ws.Cells(a2 - 1, 1), ws.Cells(b2, 9)).AutoFilter
        If Len(secim) > 0 Then ws.AutoFilter.Range.AutoFilter Field:=1, Criteria1:="=" & secim
    Else
        ws.Cells(a2, 2).Value2 = IIf(IsArray(v) Or nD > 0, "Ýlçe depolarýnda ÝSG malzemesi yok.", "Veri bekleniyor: ilçe stoklarý güncellenince ÝSG malzemeleri listelenir.")
        ws.Cells(a2, 2).Font.Italic = True: ws.Cells(a2, 2).Font.Color = RGB(100, 116, 139)
    End If
    ws.Cells(b2 + 2, 1).Value2 = "ÝSG malzemesi ana listesi ve kiþisel zimmet listesi AKSORA'da kayýtlýdýr (" & liste.Count & _
        " ÝSG malzemesi). Zimmet listesindekiler kiþiye zimmetlenir, diðerleri manuel sarf edilir."
    ws.Cells(b2 + 2, 1).Font.Size = 9: ws.Cells(b2 + 2, 1).Font.Color = RGB(71, 85, 105)
    ' KPI (panel karti BZ21 / BZ22 kullanir)
    ws.Range("BZ21").Value2 = n
    ws.Range("BZ22").Value2 = Round(toplam, 2)
    ws.Range("BZ23").Value2 = zK & " kalem  •  " & I_MIKTAR_METNI(zM) & " adet"
    ws.Range("BZ24").Value2 = sK & " kalem  •  " & I_MIKTAR_METNI(sM) & " adet"
    ws.Range("BZ21").NumberFormat = "#,##0"
    ws.Range("BZ22").NumberFormat = "#,##0.00 ""TL"""
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "ÝLÇE DEPOLARINDA ÝSG KALEMÝ", "$BZ$21", RGB(35, 119, 204)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "ÝSG STOK DEÐERÝ", "$BZ$22", RGB(39, 160, 94)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "ZÝMMET YAPILACAK", "$BZ$23", RGB(220, 38, 38)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "MANUEL SARF YAPILACAK", "$BZ$24", RGB(29, 78, 216)
    For i = 3 To 4
        Set s = Y_SEKIL(ws, "Y_KPI_DEGER_" & i)
        If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 14
    Next i
    ws.Range("BX22").Value2 = il: ws.Range("BX20").Value2 = il
    ws.Range("BX30").Value2 = secim
    ws.Range("BX27").Value2 = a2 & "|" & b2
    Y_FILIGRAN ws
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | ÝSG ZÝMMET KONTROLÜ  |  " & Y_IL_ADI(il)
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Sub I_ISG_FILTRE(ByVal alan As Long, ByVal deger As String)
    Dim ws As Worksheet, a As Long, b As Long
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("I_ÝSG_MALZEMELERÝ")
    If Not Y_ARALIK(ws, "BX27", a, b) Then Exit Sub
    ws.Unprotect ""
    If Not ws.AutoFilterMode Then ws.Range(ws.Cells(a - 1, 1), ws.Cells(b, 9)).AutoFilter
    If Len(deger) = 0 Then
        ws.AutoFilter.Range.AutoFilter Field:=alan
    Else
        ws.AutoFilter.Range.AutoFilter Field:=alan, Criteria1:="=" & deger
    End If
    If alan = 1 Then ws.Range("BX30").Value2 = deger
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Y_TEMINAT_KAL ws
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_ILCE_ISG_SEC()
    Dim ws As Worksheet, s As Shape, k As String
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("I_ÝSG_MALZEMELERÝ")
    Set s = ws.Shapes("I_ISG_SECIM")
    If s.ControlFormat.Value > 1 Then k = Trim$(Split(CStr(s.ControlFormat.List(s.ControlFormat.Value)) & "  -  ", "  -  ")(0))
    I_ISG_FILTRE 1, k
    Exit Sub
Hata:
    MsgBox Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_ILCE_ISG_TUMU(): I_ISG_FILTRE 9, "": End Sub
Public Sub AKSORA_ILCE_ISG_ZIMMET(): I_ISG_FILTRE 9, "ZÝMMET YAPIN": End Sub
Public Sub AKSORA_ILCE_ISG_SARF(): I_ISG_FILTRE 9, "MANUEL SARF YAPIN": End Sub

' Karsilastirma tarihleri degisince sayfa yenilenir.
Public Sub AKSORA_ILCE_DEGISIM(ByVal sh As Object, ByVal hedef As Range)
    Dim ev As Boolean, a As String
    If Not TypeOf sh Is Worksheet Then Exit Sub
    If mYCalisiyor Or mYonlendirme Then Exit Sub
    If Len(CStr(sh.Range("BX22").Value2)) = 0 Then Exit Sub
    ev = application.EnableEvents
    On Error GoTo Cikis
    If sh.Name = "I_STOK_KARÞILAÞTIRMA" Then
        a = CStr(sh.Range("BX34").Value2)
        If Len(a) = 0 Then GoTo Cikis
        If Intersect(hedef, sh.Range(a & "," & CStr(sh.Range("BX35").Value2))) Is Nothing Then GoTo Cikis
        application.EnableEvents = False
        AKSORA_ILCE_KARS_YENILE
    End If
Cikis:
    application.EnableEvents = ev
End Sub

'==================== ADMIN | YONETIM MERKEZI ====================
' Sifreli yonetici (CEO) raporu. Il filtresi (TUMU veya tek il): her birimin ekrandaki ili canli sayfalardan,
' diger iller birimlerin il kayitlarindan (_AKLIVE_, _YA_, _MA_, _IA_) okunur. Okunan veri imza (tarih+satir)
' degismedikce bellekte tutulur. Birimler arasi ayni depo (il+UY+depo) tekil toplamda bir kez sayilir.

' Yonetici girisi: sifre kutusu (Windows), yazilan sifre yildizla gizlenir.
Public Sub AKSORA_ADMIN_GIRIS()
    Dim s As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    s = ADM_SIFRE_SOR()
    If Len(s) = 0 Then Exit Sub
    If s <> ADM_SIFRE Then
        ADM_LOG_YAZ "HATALI"
        MsgBox "Þifre hatalý." & vbCrLf & "Yönetim merkezine yalnýzca yetkili kullanýcýlar girebilir." & vbCrLf & _
            "Deneme kayýt altýna alýndý.", vbCritical, "AKSORA | ADMÝN"
        Exit Sub
    End If
    ADM_LOG_YAZ "GÝRÝÞ"
    mAdmOturum = True
    ADM_AC
End Sub

' Sifre kutusu: InputBox acilinca zamanlayici yazilani yildizli yapar ve
' kutunun altina "Sifreyi goster" secenegini ekler. Her giriste gizli baslar.
Private Function ADM_SIFRE_SOR() As String
    On Error Resume Next
    mAdmZaman = 0: mAdmKutu = 0: mAdmGoster = False
    mAdmZaman = ADMW_SetTimer(0, 0, 40, AddressOf ADM_MASKE_ZAMAN)
    On Error GoTo 0
    ADM_SIFRE_SOR = InputBox("AKSORA YÖNETÝM MERKEZÝ" & vbCrLf & vbCrLf & "Yönetici þifresini girin:", ADM_KUTU_BASLIK)
    On Error Resume Next
    If mAdmZaman <> 0 Then ADMW_KillTimer 0, mAdmZaman
    mAdmZaman = 0: mAdmKutu = 0: mAdmGoster = False
    On Error GoTo 0
End Function

Public Sub ADM_MASKE_ZAMAN(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    Dim dlg As LongPtr, ed As LongPtr, goster As Boolean
    On Error Resume Next
    dlg = ADMW_FindWindow(vbNullString, ADM_KUTU_BASLIK)
    If dlg = 0 Then Exit Sub
    ed = ADMW_FindWindowEx(dlg, 0, "Edit", vbNullString)
    If ed = 0 Then Exit Sub
    If mAdmKutu = 0 Then
        ADMW_SendMessage ed, &HCC, 42, 0
        ADM_KUTU_EKLE dlg, ed
        Exit Sub
    End If
    goster = (ADMW_SendMessage(mAdmKutu, &HF0, 0, 0) = 1)
    If goster <> mAdmGoster Then
        mAdmGoster = goster
        ADMW_SendMessage ed, &HCC, IIf(goster, 0, 42), 0
        ADMW_InvalidateRect ed, 0, 1
        ADMW_SetFocus ed
    End If
End Sub

Private Sub ADM_KUTU_EKLE(ByVal dlg As LongPtr, ByVal ed As LongPtr)
    Dim rd As ADMW_RECT, re As ADMW_RECT, h As Long, metin As String, sinif As String
    ADMW_GetWindowRect dlg, rd
    ADMW_GetWindowRect ed, re
    ADMW_MapWindowPoints 0, dlg, re, 2
    h = re.Bottom - re.Top
    If h < 16 Then h = 16
    ADMW_SetWindowPos dlg, 0, 0, 0, rd.Right - rd.Left, rd.Bottom - rd.Top + h + 10, &H16
    sinif = "Button"
    metin = ChrW(350) & "ifreyi g" & ChrW(246) & "ster"
    mAdmKutu = ADMW_CreateWindowEx(0, StrPtr(sinif), StrPtr(metin), &H50010003, re.Left, re.Bottom + 6, re.Right - re.Left, h, dlg, 4001, 0, 0)
    If mAdmKutu = 0 Then mAdmKutu = -1: Exit Sub
    ADMW_SendMessage mAdmKutu, &H30, ADMW_SendMessage(ed, &H31, 0, 0), 1
End Sub

Private Sub ADM_AC()
    Dim ws As Worksheet, ev As Boolean, sc As Boolean, no As Long, metin As String, hesap As Long
    If mAdmMesgul Then Exit Sub
    mAdmMesgul = True
    ev = application.EnableEvents: sc = application.ScreenUpdating: hesap = application.Calculation
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    If hesap <> xlCalculationManual Then application.Calculation = xlCalculationManual
    AKSORA_BEKLE_AC "AKSORA  |  YÖNETÝM MERKEZÝ", "Yönetim merkezi hazýrlanýyor...", 0.03
    application.Cursor = xlWait
    mYonlendirme = True
    Set ws = ADM_SAYFA()
    ADM_KUR ws
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    BIRIM_DIGERLERINI_GIZLE ADM
    ADM_GORUNUM ws
Cikis:
    AKSORA_BEKLE_KAPAT
    mAdmMesgul = False
    mYonlendirme = False
    application.StatusBar = False
    application.Cursor = xlDefault
    If application.Calculation <> hesap Then application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = True
    application.Interactive = True
    If no <> 0 Then MsgBox "Yönetim merkezi açýlamadý." & vbCrLf & vbCrLf & metin, vbExclamation, "AKSORA | ADMÝN"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

'---------- bekleme karti (tum uzun gecisler) ----------
' AKSORA penceresinin tam ortasinda, basliksiz kucuk kart (frmSplash kart olarak yeniden bicimlenir).
' Kart kendi penceresinde cizildigi icin Excel ekrani donsa da hemen gorunur; AKSORA hangi ekrandaysa orada acilir.
Public Function AKSORA_BEKLE_AC(ByVal baslik As String, ByVal mesaj As String, Optional ByVal y0 As Double = 0.4, _
    Optional ByVal durdur As Boolean = False) As Boolean
    If AKSORA_V13_ACILIS Or AKSORA_ADMIN_TOPLU Or AKSORA_V11_MESGUL Then Exit Function
    If mAdmYukleme Then AKSORA_BEKLE_KAPAT
    mAdmYukleme = False
    On Error GoTo Son
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Function
    If application.WindowState = xlMinimized Then Exit Function
    AKSORA_BEKLE_KIPI = True
    Load frmSplash
    AKSORA_BEKLE_KIPI = False
    BEKLE_KART_BICIM baslik, durdur
    frmSplash.Show vbModeless
    BEKLE_KART_KONUM
    mAdmYukleme = True: mAdmYukSon = 0
    AKSORA_BEKLE_AC = True
    AKSORA_BEKLE_ILERLE y0, mesaj
    Exit Function
Son:
    AKSORA_BEKLE_KIPI = False
    On Error Resume Next
    Unload frmSplash
    mAdmYukleme = False
    On Error GoTo 0
End Function

Private Sub BEKLE_KART_BICIM(ByVal baslik As String, Optional ByVal durdur As Boolean = False)
    Dim w As Single
    On Error Resume Next
    With frmSplash
        .Caption = "AKSORA"
        .StartUpPosition = 0
        .Width = 420: .Height = IIf(durdur, 182, 150)
        .Left = application.Left + (application.Width - .Width) / 2
        .Top = application.Top + (application.Height - .Height) / 2
        .BackColor = RGB(255, 255, 255)
        .BorderStyle = fmBorderStyleSingle: .BorderColor = RGB(20, 103, 173)
        w = .InsideWidth
        .imgLogo.Visible = False
        .lblImza.Visible = durdur
        With .lblBaslik
            .Caption = baslik: .Left = 0: .Top = 18: .Width = w: .Height = 26
            .Font.Name = "Segoe UI": .Font.Size = 15: .Font.bold = True: .Font.Italic = False
            .ForeColor = RGB(12, 45, 79): .TextAlign = fmTextAlignCenter: .WordWrap = False: .BackStyle = fmBackStyleTransparent
        End With
        With .lblDurum
            ' Durdurulabilir islerde iki satir: adim + gecen sure.
            .Caption = vbNullString: .Left = 8: .Top = IIf(durdur, 50, 58): .Width = w - 16: .Height = IIf(durdur, 34, 18)
            .Font.Name = "Segoe UI": .Font.Size = 10: .Font.bold = False: .Font.Italic = False
            .ForeColor = RGB(20, 103, 173): .TextAlign = fmTextAlignCenter: .WordWrap = durdur: .BackStyle = fmBackStyleTransparent
        End With
        With .lblBarArka
            .Left = 36: .Top = IIf(durdur, 92, 88): .Width = w - 72: .Height = 8: .BackColor = RGB(224, 232, 240): .BackStyle = fmBackStyleOpaque
        End With
        With .lblBar
            .Left = 36: .Top = IIf(durdur, 92, 88): .Width = 1: .Height = 8: .BackColor = RGB(0, 125, 110): .BackStyle = fmBackStyleOpaque
            .ZOrder 0
        End With
        If durdur Then
            ' DURDUR (lblImza) ve ARKA PLANA AL (lblAltBaslik) yan yana dugme.
            With .lblImza
                .Caption = "DURDUR": .Left = w / 2 - 166: .Top = 114: .Width = 160: .Height = 28
                .Font.Name = "Segoe UI": .Font.Size = 11: .Font.bold = True: .Font.Italic = False
                .ForeColor = vbWhite: .BackColor = RGB(220, 38, 38): .BackStyle = fmBackStyleOpaque
                .TextAlign = fmTextAlignCenter: .WordWrap = False: .BorderStyle = fmBorderStyleNone
            End With
            With .lblAltBaslik
                .Caption = "ARKA PLANA AL": .Left = w / 2 + 6: .Top = 114: .Width = 160: .Height = 28
                .Font.Name = "Segoe UI": .Font.Size = 11: .Font.bold = True: .Font.Italic = False
                .ForeColor = vbWhite: .BackColor = RGB(20, 103, 173): .BackStyle = fmBackStyleOpaque
                .TextAlign = fmTextAlignCenter: .WordWrap = False: .BorderStyle = fmBorderStyleNone
            End With
        Else
            With .lblAltBaslik
                .Caption = "Lütfen bekleyin, AKSORA verileri hazýrlýyor": .Left = 0: .Top = 106: .Width = w: .Height = 16
                .Font.Name = "Segoe UI": .Font.Size = 8.5: .Font.bold = False: .Font.Italic = False
                .ForeColor = RGB(100, 116, 139): .TextAlign = fmTextAlignCenter: .WordWrap = False: .BackStyle = fmBackStyleTransparent
            End With
        End If
    End With
    On Error GoTo 0
End Sub

' Kart penceresinin basligi kaldirilir ve AKSORA penceresinin tam ortasina alinir.
Private Sub BEKLE_KART_KONUM()
    Dim h As LongPtr, st As LongPtr, rx As ADMW_RECT, rf As ADMW_RECT, x As Long, y As Long
    On Error Resume Next
    h = ADMW_FindWindow("ThunderDFrame", "AKSORA")
    If h = 0 Then h = ADMW_FindWindow("ThunderXFrame", "AKSORA")
    If h = 0 Then Exit Sub
    st = ADMW_GetWindowLong(h, -16)
    st = st And Not &HC40000
    ADMW_SetWindowLong h, -16, st
    ADMW_SetWindowPos h, 0, 0, 0, 0, 0, &H37
    ADMW_GetWindowRect application.hWnd, rx
    ADMW_GetWindowRect h, rf
    x = (rx.Left + rx.Right) \ 2 - (rf.Right - rf.Left) \ 2
    y = (rx.Top + rx.Bottom) \ 2 - (rf.Bottom - rf.Top) \ 2
    ADMW_SetWindowPos h, 0, x, y, 0, 0, &H15
    frmSplash.Repaint
    On Error GoTo 0
End Sub

' Karti ilerletir; kart kendi penceresinde hemen cizilir. Uzun islemlerde pencere "yanit vermiyor"a
' dusmesin diye yarim saniyede bir mesaj kuyrugu bosaltilir (bu sirada tiklamalar engelli).
Public Sub AKSORA_BEKLE_ILERLE(ByVal y As Double, ByVal mesaj As String)
    Dim iv As Boolean
    If Not mAdmYukleme Then Exit Sub
    On Error Resume Next
    If y > 1 Then y = 1
    If y < 0 Then y = 0
    frmSplash.lblDurum.Caption = mesaj
    frmSplash.lblBar.Width = application.Max(1, frmSplash.lblBarArka.Width * y)
    frmSplash.Repaint
    If mAdmYukSon = 0 Or Abs(Timer - mAdmYukSon) > 0.5 Then
        iv = application.Interactive
        application.Interactive = False
        DoEvents
        application.Interactive = iv
        ARKA_PLAN_KORU
        mAdmYukSon = Timer
    End If
    On Error GoTo 0
End Sub

' Uzun dongulerde cagrilir: kart acikken yarim saniyede bir mesaj kuyrugunu bosaltir (yazilar degismez).
Public Sub AKSORA_BEKLE_NABIZ()
    Dim iv As Boolean
    If Not mAdmYukleme Then Exit Sub
    If Abs(Timer - mAdmYukSon) <= 0.5 Then Exit Sub
    On Error Resume Next
    iv = application.Interactive
    application.Interactive = False
    DoEvents
    application.Interactive = iv
    ARKA_PLAN_KORU
    mAdmYukSon = Timer
    AKSORA_ADMIN_DURUM vbNullString
    On Error GoTo 0
End Sub

' Kartin "ARKA PLANA AL" dugmesi (frmSplash lblAltBaslik_Click): kart ve Excel penceresi z-sirasinda en alta
' gonderilir; pencere boyutu/durumu degismez, giris kapali kalir. Kullanici baska programlarla calisir.
Public Sub AKSORA_ARKA_PLANA_AL()
    Dim h As LongPtr
    On Error Resume Next
    AKSORA_ARKA_PLAN = True
    h = ADMW_FindWindow("ThunderDFrame", "AKSORA")
    If h = 0 Then h = ADMW_FindWindow("ThunderXFrame", "AKSORA")
    ' SWP_NOSIZE Or SWP_NOMOVE Or SWP_NOACTIVATE; HWND_BOTTOM = 1
    If h <> 0 Then ADMW_SetWindowPos h, 1, 0, 0, 0, 0, &H13
    ADMW_SetWindowPos application.hWnd, 1, 0, 0, 0, 0, &H13
    On Error GoTo 0
End Sub

' Uzun is surerken Excel simge durumuna alinmissa geri acilir ve arkaya gonderilir (kucuk pencerede sayfa denetimleri yazilamiyor).
Private Sub ARKA_PLAN_KORU()
    If Not (AKSORA_ADMIN_TOPLU Or AKSORA_ATIL_CALISIYOR) Then Exit Sub
    On Error Resume Next
    If application.WindowState = xlMinimized Then
        application.WindowState = xlMaximized
        AKSORA_ARKA_PLANA_AL
    End If
    On Error GoTo 0
End Sub

' Is bitince Excel gorev cubugunda yanip soner (pencere one zorla getirilmez); sonuc mesaji Excel'de bekler.
Public Sub AKSORA_ARKA_PLAN_BITIR()
    Dim f As ADMW_FLASH
    If Not AKSORA_ARKA_PLAN Then Exit Sub
    AKSORA_ARKA_PLAN = False
    On Error Resume Next
    f.cbSize = LenB(f): f.hWnd = application.hWnd: f.dwFlags = 15: f.uCount = 0: f.dwTimeout = 0
    ADMW_FlashWindowEx f
    On Error GoTo 0
End Sub

Public Sub AKSORA_BEKLE_KAPAT()
    On Error Resume Next
    ADM_EKRAN_KILIT False
    If mAdmYukleme Then Unload frmSplash
    mAdmYukleme = False: Set mAdmYukWs = Nothing
    On Error GoTo 0
End Sub

' Koruma ayarlarini saklayip korumayi gecici kaldirir (korumasizsa ayar bos kalir).
Private Function BEKLE_SERBEST(ByVal ws As Worksheet, ByRef ayar As Variant) As Boolean
    ayar = Empty
    If Not (ws.ProtectContents Or ws.ProtectDrawingObjects Or ws.ProtectScenarios) Then BEKLE_SERBEST = True: Exit Function
    With ws.Protection
        ayar = Array(ws.ProtectContents, ws.ProtectDrawingObjects, ws.ProtectScenarios, .AllowFiltering, .AllowSorting, _
            .AllowFormattingCells, .AllowFormattingColumns, .AllowFormattingRows)
    End With
    On Error Resume Next
    ws.Unprotect pw
    BEKLE_SERBEST = (Err.Number = 0)
    On Error GoTo 0
End Function

Private Sub BEKLE_KORU(ByVal ws As Worksheet, ByVal ayar As Variant)
    If Not IsArray(ayar) Then Exit Sub
    On Error Resume Next
    ws.Protect Password:=pw, Contents:=ayar(0), DrawingObjects:=ayar(1), Scenarios:=ayar(2), UserInterfaceOnly:=True, _
        AllowFiltering:=ayar(3), AllowSorting:=ayar(4), AllowFormattingCells:=ayar(5), AllowFormattingColumns:=ayar(6), AllowFormattingRows:=ayar(7)
    On Error GoTo 0
End Sub

Private Sub BEKLE_METIN(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, ByVal x As Double, ByVal y As Double, _
    ByVal w As Double, ByVal h As Double, ByVal boyut As Double, ByVal renk As Long, ByVal kalin As Boolean)
    Dim s As Shape
    Set s = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, w, h)
    s.Name = ad: s.Fill.Visible = msoFalse: s.line.Visible = msoFalse
    With s.TextFrame2
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
        .WordWrap = msoFalse: .AutoSize = msoAutoSizeNone
        .TextRange.Text = metin
        .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = boyut: .TextRange.Font.bold = IIf(kalin, msoTrue, msoFalse)
        .TextRange.Font.Fill.ForeColor.RGB = renk
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
End Sub

Private Sub BEKLE_SIL(ByVal ws As Worksheet)
    Dim i As Long
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) = "ADMY_" Then ws.Shapes(i).Delete
    Next i
End Sub

Private Sub ADM_ILERLE(ByVal y As Double, ByVal mesaj As String)
    AKSORA_BEKLE_ILERLE y, mesaj
End Sub

Public Sub AKSORA_ADMIN_YENILE()
    ADM_AC
End Sub

Public Sub AKSORA_ADMIN_CIKIS()
    mAdmOturum = False
    AKSORA_BIRIM_ANA_MENU
End Sub

Public Sub AKSORA_ADMIN_CIKTI()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ADM)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    AKSORA_CIKTI_SAYFA ws
End Sub

Private Function ADM_SAYFA() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ADM)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = ADM
    End If
    Set ADM_SAYFA = ws
End Function

Private Sub ADM_GORUNUM(ByVal ws As Worksheet)
    Dim z As Double
    If Not (ActiveSheet Is ws) Then Exit Sub
    On Error Resume Next
    With ThisWorkbook.Windows(1)
        If AKSORA_PENCERE_KUCUK(ThisWorkbook.Windows(1)) Then Exit Sub
        .DisplayGridlines = False: .DisplayHeadings = False: .DisplayWorkbookTabs = False
        .DisplayHorizontalScrollBar = False: .DisplayVerticalScrollBar = True
        z = (.UsableWidth - 20) / ws.Range("A1:N1").Width * 100
        z = application.Max(10, application.Min(400, Fix(z)))
        If .Zoom <> z Then .Zoom = z
        .ScrollRow = 1: .ScrollColumn = 1
    End With
    ws.Range(ADM_ARA_HUCRE).Select
    On Error GoTo 0
End Sub

' Il secimi (C4) raporu yeniden kurar; arama kutusu (C30) aramayi calistirir.
Public Sub AKSORA_ADMIN_DEGISIM(ByVal hedef As Range)
    Dim ev As Boolean, ws As Worksheet, il As String
    If mYonlendirme Then Exit Sub
    Set ws = hedef.Worksheet
    ev = application.EnableEvents
    application.EnableEvents = False
    On Error GoTo Son
    If Not Intersect(hedef, ws.Range(ADM_IL_HUCRE)) Is Nothing Then
        il = ADM_IL_KOD(CStr(ws.Range(ADM_IL_HUCRE).Value2))
        ws.Unprotect ""
        ws.Range("BX3").Value2 = il
        application.EnableEvents = ev
        ADM_AC
        Exit Sub
    End If
    If Not Intersect(hedef, ws.Range(ADM_ARA_HUCRE)) Is Nothing Then AKSORA_ADMIN_ARA
Son:
    application.EnableEvents = ev
End Sub

'---------- birimler, iller ----------
' Birim: ad, renk, canli stok sayfalari, duzen (1 malzeme sutunlari, 2 musteri/yatirim/ilce sutunlari), kisa ad
Private Function ADM_BIRIMLER() As Variant
    ADM_BIRIMLER = Array( _
        Array("Malzeme Stok Birimi", RGB(0, 125, 110), Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK"), 1, "MALZEME"), _
        Array("Yatýrým Birimi", RGB(20, 103, 173), Array("Y_CDP_YATIRIM_STOK", "Y_EPP_YATIRIM_STOK", "Y_CDP_YÜKLENÝCÝ_STOK", "Y_EPP_YÜKLENÝCÝ_STOK"), 2, "YATIRIM"), _
        Array("Müþteri Birimi", RGB(32, 133, 158), Array("M_CDP_STOK", "M_ÝLÇE_STOK", "M_YÜKLENÝCÝ_STOK"), 2, "MÜÞTERÝ"), _
        Array("Ýlçe Yönetim Birimi", RGB(79, 70, 160), Array("I_CDP_ÝLÇELER_STOK", "I_EPP_ÝLÇELER_STOK"), 2, "ÝLÇE"))
End Function

Private Function ADM_ILLER() As Variant
    ADM_ILLER = Array("ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
End Function

Private Function ADM_IL_KOD(ByVal metin As String) As String
    metin = AKFIX_NORMAL(Replace(Replace(metin, "Ýl Müdürlüðü", ""), "ÝL MÜDÜRLÜÐÜ", ""))
    Select Case metin
        Case "ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE": ADM_IL_KOD = metin
        Case Else: ADM_IL_KOD = vbNullString
    End Select
End Function

' Birimin ekrandaki ili (canli stok sayfasinin A5 basligindan).
Private Function ADM_CANLI_IL(ByVal u As Long) As String
    Dim kx As Variant, sh As Worksheet, bb As Variant
    bb = ADM_BIRIMLER()
    For Each kx In bb(u)(2)
        Set sh = ADM_WS(CStr(kx))
        If Not sh Is Nothing Then
            ADM_CANLI_IL = ADM_IL_KOD(ADM_METIN(sh.Range("A5").Value2))
            If Len(ADM_CANLI_IL) > 0 Then Exit Function
        End If
    Next kx
End Function

' Il kayit sayfalari (canli il disindaki iller icin).
Private Function ADM_KAYIT_ADLARI(ByVal u As Long, ByVal il As String) As Variant
    Select Case u
        Case 0: ADM_KAYIT_ADLARI = Array("_AKLIVE_" & il & "_CDP", "_AKLIVE_" & il & "_EPP", "_AKLIVE_" & il & "_ILC", "_AKLIVE_" & il & "_YUK")
        Case 1: ADM_KAYIT_ADLARI = Array("_YA_" & il & "_1", "_YA_" & il & "_2", "_YA_" & il & "_3", "_YA_" & il & "_4")
        Case 2: ADM_KAYIT_ADLARI = Array("_MA_" & il & "_CDP", "_MA_" & il & "_ILCE", "_MA_" & il & "_YUK")
        Case Else: ADM_KAYIT_ADLARI = Array("_IA_" & il)
    End Select
End Function

'---------- veri okuma (bellekte tutulur) ----------
' Kayit: Array(0 tarih, 1 satirlar() metin, 2 satir sayisi, 3 ozet, 4 depolar, 5 isg, 6 kaynak)
'   satir metni: UY|depo|depo adi|malzeme|tanim|parti|miktar|olcu|deger|seri  (vbTab ile)
'   ozet: Array(satir, deger, eksi, fiyatsiz, hurda kalem, hurda TL, cikma kalem)
'   depolar: UY|depo -> Array(ad, satir, deger)
'   isg (yalniz ilce): depo -> Array(ad, zimmet kalem, zimmet adet, sarf kalem, sarf adet)
Private Function ADM_VERI(ByVal u As Long, ByVal il As String) As Variant
    Dim imza As String, canli As Boolean, kx As Variant, sh As Worksheet, son As Long, t As Variant, k As String
    Dim v As Variant, satirlar() As String, n As Long, tarih As Date, ozet(0 To 6) As Double, depolar As Object, isg As Object
    Dim i As Long, cU As Long, cD As Long, cA As Long, cM As Long, cT As Long, cP As Long, cQ As Long, cB As Long, cV As Long, cS As Long, cF As Long
    Dim mal As String, parti As String, x As Double, d As Double, liste As Object, kod As String, kalem As Object, duzen As Long
    Dim bb As Variant, kaynaklar As Variant
    If mAdmOnbellek Is Nothing Then Set mAdmOnbellek = CreateObject("Scripting.Dictionary")
    bb = ADM_BIRIMLER()
    canli = (il = ADM_CANLI_IL(u))
    ' imza: tarih ve satir sayilari
    If canli Then
        For Each kx In bb(u)(2)
            Set sh = ADM_WS(CStr(kx))
            If Not sh Is Nothing Then imza = imza & "|" & CStr(sh.Range("C2").Value2) & "#" & sh.Cells(sh.Rows.Count, IIf(u = 0, 7, 5)).End(xlUp).Row
        Next kx
    Else
        For Each kx In ADM_KAYIT_ADLARI(u, il)
            Set sh = ADM_WS(CStr(kx))
            If Not sh Is Nothing Then imza = imza & "|" & CStr(sh.Range("A1").Value2) & CStr(sh.Range("A3").Value2) & "#" & CStr(sh.Range("D1").Value2) & "#" & CStr(sh.Range("A4").Value2)
        Next kx
    End If
    imza = IIf(canli, "C", "K") & imza
    k = u & "|" & il
    If mAdmOnbellek.Exists(k) Then
        If mAdmOnbellek(k)(7) = imza Then ADM_VERI = mAdmOnbellek(k): Exit Function
    End If
    application.StatusBar = "AKSORA Yönetim Merkezi  •  " & bb(u)(4) & "  •  " & Y_IL_ADI(il) & " okunuyor..."
    Set depolar = CreateObject("Scripting.Dictionary"): depolar.CompareMode = vbTextCompare
    Set isg = CreateObject("Scripting.Dictionary"): isg.CompareMode = vbTextCompare
    Set kalem = CreateObject("Scripting.Dictionary"): kalem.CompareMode = vbTextCompare
    If u = 3 Then Set liste = I_ISG_LISTE()
    ReDim satirlar(1 To 1)
    If canli Then kaynaklar = bb(u)(2) Else kaynaklar = ADM_KAYIT_ADLARI(u, il)
    For Each kx In kaynaklar
        Set sh = ADM_WS(CStr(kx))
        v = Empty: duzen = IIf(u = 0, 1, 2)
        If Not sh Is Nothing Then
            If canli Or u = 0 Then
                ' canli sayfa veya malzeme kaydi (_AKLIVE_): stok sayfasiyla ayni duzen, veri 7. satirdan
                If canli Then
                    t = sh.Range("C2").Value
                Else
                    t = sh.Range("A3").Value
                End If
                If Not IsError(t) Then
                    If IsDate(t) Then If CDate(t) > tarih Then tarih = CDate(t)
                End If
                son = sh.Cells(sh.Rows.Count, IIf(u = 0, 7, 5)).End(xlUp).Row
                If son >= 7 Then v = sh.Range(sh.Cells(7, 1), sh.Cells(son, 14)).Value2
            Else
                t = sh.Range("A1").Value
                If Not IsError(t) Then
                    If IsDate(t) Then If CDate(t) > tarih Then tarih = CDate(t)
                End If
                On Error Resume Next
                v = AKSORA_PAKET_OKU(sh)
                If Err.Number <> 0 Then v = Empty
                On Error GoTo 0
            End If
        End If
        If IsArray(v) Then
            If duzen = 1 Then
                cU = 1: cD = 2: cA = 3: cP = 6: cM = 7: cT = 8: cQ = 9: cB = 10: cV = 11: cS = 0: cF = 0
            Else
                cU = 1: cD = 2: cA = 3: cM = 5: cT = 6: cQ = 7: cB = 8: cP = 9: cS = 10: cV = 11: cF = 14
            End If
            If n + UBound(v, 1) > UBound(satirlar) Then ReDim Preserve satirlar(1 To n + UBound(v, 1) + 1)
            For i = 1 To UBound(v, 1)
                If (i And 2047) = 0 Then AKSORA_BEKLE_NABIZ
                mal = Y_KOD_NORM(v(i, cM))
                If Len(mal) > 0 Then
                    n = n + 1
                    x = I_SAYI(v(i, cQ)): d = I_SAYI(v(i, cV)): parti = I_TEMIZ(v(i, cP))
                    satirlar(n) = I_TEMIZ(v(i, cU)) & vbTab & I_TEMIZ(v(i, cD)) & vbTab & I_TEMIZ(v(i, cA)) & vbTab & mal & vbTab & _
                        I_TEMIZ(v(i, cT)) & vbTab & parti & vbTab & Trim$(Str$(x)) & vbTab & I_TEMIZ(v(i, cB)) & vbTab & Trim$(Str$(d)) & vbTab & _
                        IIf(cS > 0, I_TEMIZ(v(i, IIf(cS > 0, cS, 1))), "")
                    ozet(0) = ozet(0) + 1: ozet(1) = ozet(1) + d
                    If x < 0 Then ozet(2) = ozet(2) + 1
                    If u = 2 And cF > 0 Then
                        If UBound(v, 2) >= cF Then If I_SAYI(v(i, cF)) = 1 Then ozet(3) = ozet(3) + 1
                    End If
                    If InStr(1, parti, "HURDA", vbTextCompare) > 0 Then ozet(4) = ozet(4) + 1: ozet(5) = ozet(5) + d
                    If InStr(1, parti, "CIKMA", vbTextCompare) > 0 Then ozet(6) = ozet(6) + 1
                    kod = I_TEMIZ(v(i, cU)) & "|" & I_TEMIZ(v(i, cD))
                    If Not depolar.Exists(kod) Then depolar(kod) = Array(I_TEMIZ(v(i, cA)), 0#, 0#)
                    t = depolar(kod): t(1) = t(1) + 1: t(2) = t(2) + d: depolar(kod) = t
                    ' Ilce ISG: zimmet listesindeki (Z) ve manuel sarf (M) malzemeler, depoda miktar > 0
                    If u = 3 Then
                        If liste.Exists(mal) And x > 0 Then
                            kod = UCase$(I_TEMIZ(v(i, cD)))
                            If Not isg.Exists(kod) Then isg(kod) = Array(I_TEMIZ(v(i, cA)), 0#, 0#, 0#, 0#)
                            t = isg(kod)
                            If Not kalem.Exists(kod & "|" & I_TEMIZ(v(i, cU)) & "|" & mal) Then
                                kalem(kod & "|" & I_TEMIZ(v(i, cU)) & "|" & mal) = True
                                If liste(mal) = "Z" Then t(1) = t(1) + 1 Else t(3) = t(3) + 1
                            End If
                            If liste(mal) = "Z" Then t(2) = t(2) + x Else t(4) = t(4) + x
                            isg(kod) = t
                        End If
                    End If
                End If
            Next i
        End If
    Next kx
    t = Array(tarih, satirlar, n, ozet, depolar, isg, IIf(canli, "ekran", "kayýt"), imza)
    mAdmOnbellek(k) = t
    ADM_VERI = t
End Function

Private Function ADM_WS(ByVal ad As String) As Worksheet
    On Error Resume Next
    Set ADM_WS = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
End Function

Private Function ADM_SAYI(ByVal v As Variant) As Double
    ADM_SAYI = I_SAYI(v)
End Function

Private Function ADM_METIN(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    ADM_METIN = Trim$(CStr(v))
End Function

Private Function ADM_TL(ByVal x As Double) As String
    ADM_TL = AKSORA_SAYI_METNI(x, "#,##0.00") & " TL"
End Function

' Buyuk tutarlar kisa: 12,4 mn TL
Private Function ADM_TL_KISA(ByVal x As Double) As String
    If Abs(x) >= 1000000# Then
        ADM_TL_KISA = AKSORA_SAYI_METNI(x / 1000000#, "#,##0.0") & " mn TL"
    ElseIf Abs(x) >= 1000# Then
        ADM_TL_KISA = AKSORA_SAYI_METNI(x / 1000#, "#,##0.0") & " bin TL"
    Else
        ADM_TL_KISA = AKSORA_SAYI_METNI(x, "#,##0") & " TL"
    End If
End Function

Private Function ADM_ADET(ByVal x As Double) As String
    If x = Int(x) Then ADM_ADET = AKSORA_SAYI_METNI(x, "#,##0") Else ADM_ADET = AKSORA_SAYI_METNI(x, "#,##0.###")
End Function

' Uyari: Array(seviye 1 KRITIK / 2 UYARI / 3 BILGI, mesaj, kaynak, kategori)
Private Sub ADM_UYAR(ByVal c As Collection, ByVal seviye As Long, ByVal mesaj As String, ByVal kaynak As String, Optional ByVal kategori As String = "")
    If Len(kategori) = 0 Then kategori = kaynak
    c.Add Array(seviye, mesaj, kaynak, kategori)
End Sub

Private Function ADM_GUN(ByVal t As Date) As Long
    If t = 0 Then ADM_GUN = -1 Else ADM_GUN = DateDiff("d", DateValue(t), Date)
End Function

'---------- il secim dugmeleri ----------
Public Sub AKSORA_ADMIN_IL_SEC()
    Dim ws As Worksheet, il As String
    If mAdmMesgul Then Exit Sub
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ADM)
    il = ws.Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    ws.Unprotect ""
    ws.Range("BX3").Value2 = il
    ADM_AC
End Sub

Public Sub AKSORA_ADMIN_ARA_TEMIZLE()
    Dim ws As Worksheet, ev As Boolean
    Set ws = ThisWorkbook.Worksheets(ADM)
    ev = application.EnableEvents: application.EnableEvents = False
    On Error Resume Next
    ws.Unprotect ""
    ws.Range(ADM_ARA_HUCRE).MergeArea.ClearContents
    ws.Range("BX4").ClearContents
    On Error GoTo 0
    application.EnableEvents = ev
    AKSORA_ADMIN_ARA
End Sub

Private Sub ADM_IL_DUGMELERI(ByVal ws As Worksheet, ByVal secIl As String)
    Dim iller As Variant, i As Long, s As Shape, x As Double, w As Double, kod As String, secili As Boolean
    iller = Array("", "ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
    w = ws.Columns("C").Width * 1.05
    x = ws.Range("C4").Left
    For i = 0 To 5
        kod = CStr(iller(i)): secili = (kod = secIl)
        Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x + i * (w + 6), ws.Range("C4").Top + 4, w, ws.Rows(4).Height - 8)
        With s
            .Name = "ADM_IL_" & i
            .AlternativeText = kod
            .Adjustments.Item(1) = 0.3
            .Fill.ForeColor.RGB = IIf(secili, RGB(12, 45, 79), vbWhite)
            .line.ForeColor.RGB = RGB(20, 103, 173): .line.Weight = IIf(secili, 1.5, 1)
            .Placement = xlFreeFloating
            .OnAction = BIRIM_MAKRO("AKSORA_ADMIN_IL_SEC")
            With .TextFrame2
                .MarginLeft = 2: .MarginRight = 2: .MarginTop = 0: .MarginBottom = 0
                .VerticalAnchor = msoAnchorMiddle
                .TextRange.Text = IIf(Len(kod) = 0, "TÜMÜ", UCase$(Y_IL_ADI(kod)))
                .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 10: .TextRange.Font.bold = msoTrue
                .TextRange.Font.Fill.ForeColor.RGB = IIf(secili, vbWhite, RGB(12, 45, 79))
                .TextRange.ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
    Next i
End Sub

'---------- demontaj kontrolu (ilce gecmisinden, her il icin) ----------
' Ilce biriminin STOK/DEMONTAJ ekraniyla ayni kural: yeni = son guncelleme, eski = onceki gunun son guncellemesi.
' Azalan (demontaj disi) her malzeme icin ayni depo + UY + malzemede CIKMA/BDSZ girisi aranir.
' Donus: Array(eskiT, yeniT, depolar, eksikler)
'   depolar: depo -> Array(ad, azalan kalem, girilen kalem, eksik kalem, eksik adet)
'   eksikler: Collection of Array(il, depo, depo adi, malzeme, tanim, eksik adet, olcu, durum)
Private Function ADM_DEMONTAJ(ByVal il As String) As Variant
    Dim tarihler As Collection, eskiT As String, yeniT As String, i As Long, dE As Object, dY As Object, anahtar As Object
    Dim grup As Object, k As Variant, p() As String, t As Variant, g As Variant, depolar As Object, eksik As Collection
    Dim mEs As Double, mYe As Double, gk As String, ad As String, imza As String, o As Variant, bf As Double, vEs As Double, isgL As Object
    If mAdmOnbellek Is Nothing Then Set mAdmOnbellek = CreateObject("Scripting.Dictionary")
    Set depolar = CreateObject("Scripting.Dictionary"): depolar.CompareMode = vbTextCompare
    Set eksik = New Collection
    Set tarihler = I_GECMIS_TARIHLER(il)
    If tarihler.Count > 0 Then
        yeniT = Format$(tarihler(1), I_TF)
        For i = 2 To tarihler.Count
            If Int(CDbl(tarihler(i))) < Int(CDbl(tarihler(1))) Then eskiT = Format$(tarihler(i), I_TF): Exit For
        Next i
    End If
    If Len(eskiT) = 0 Then ADM_DEMONTAJ = Array(eskiT, yeniT, depolar, eksik): Exit Function
    imza = "DEM5|" & il & "|" & eskiT & "|" & yeniT
    If mAdmOnbellek.Exists(imza) Then ADM_DEMONTAJ = mAdmOnbellek(imza): Exit Function
    application.StatusBar = "AKSORA Yönetim Merkezi  •  " & Y_IL_ADI(il) & " demontaj kontrolü..."
    Set dE = I_GECMIS_OKU(il, eskiT): Set dY = I_GECMIS_OKU(il, yeniT)
    Set anahtar = CreateObject("Scripting.Dictionary"): anahtar.CompareMode = vbTextCompare
    Set grup = CreateObject("Scripting.Dictionary"): grup.CompareMode = vbTextCompare
    ' ISG malzemeleri demontaj kontrolune girmez (zimmet / manuel sarf ile cikar).
    Set isgL = I_ISG_LISTE()
    For Each k In dE.keys
        If Len(I_DEM_DISI(ADM_KOD3(k), isgL)) = 0 Then anahtar(k) = True
    Next k
    For Each k In dY.keys
        If Len(I_DEM_DISI(ADM_KOD3(k), isgL)) = 0 Then anahtar(k) = True
    Next k
    For Each k In anahtar.keys
        mEs = 0: mYe = 0
        If dE.Exists(k) Then mEs = dE(k)(3)
        If dY.Exists(k) Then mYe = dY(k)(3)
        p = Split(CStr(k), "|")
        If UBound(p) >= 3 Then
            gk = p(0) & "|" & p(1) & "|" & p(2)
            If Not grup.Exists(gk) Then grup(gk) = Array(0#, 0#)
            g = grup(gk)
            If I_DEMONTAJ_MI(p(3)) Then
                If mYe > mEs Then g(1) = g(1) + (mYe - mEs)
            ElseIf mYe < mEs Then
                g(0) = g(0) + (mEs - mYe)
            End If
            grup(gk) = g
        End If
    Next k
    i = 0
    For Each k In anahtar.keys
        i = i + 1: If (i And 2047) = 0 Then AKSORA_BEKLE_NABIZ
        mEs = 0: mYe = 0: vEs = 0
        If dE.Exists(k) Then mEs = dE(k)(3): vEs = dE(k)(4): t = dE(k)
        If dY.Exists(k) Then mYe = dY(k)(3): t = dY(k)
        p = Split(CStr(k), "|")
        If UBound(p) >= 3 Then
            If mYe + 0.0005 < mEs And Not I_DEMONTAJ_MI(p(3)) Then
                ad = CStr(t(0))
                If mEs > 0 Then bf = vEs / mEs Else bf = 0
                If Not depolar.Exists(p(0)) Then depolar(p(0)) = Array(ad, 0&, 0&, 0&, 0#, 0#)
                o = depolar(p(0))
                o(1) = o(1) + 1
                g = grup(p(0) & "|" & p(1) & "|" & p(2))
                If g(1) <= 0.0005 Then
                    o(3) = o(3) + 1: o(4) = o(4) + (mEs - mYe): o(5) = o(5) + (mEs - mYe) * bf
                    eksik.Add Array(il, p(0), ad, p(2), CStr(t(1)), mEs - mYe, CStr(t(2)), "GÝRÝÞ YOK", (mEs - mYe) * bf)
                ElseIf g(1) + 0.0005 < g(0) Then
                    gk = "#" & p(0) & "|" & p(1) & "|" & p(2)
                    If Not grup.Exists(gk) Then
                        grup(gk) = True
                        o(3) = o(3) + 1: o(4) = o(4) + (g(0) - g(1)): o(5) = o(5) + (g(0) - g(1)) * bf
                        eksik.Add Array(il, p(0), ad, p(2), CStr(t(1)), g(0) - g(1), CStr(t(2)), "EKSÝK GÝRÝÞ", (g(0) - g(1)) * bf)
                    End If
                Else
                    o(2) = o(2) + 1
                End If
                depolar(p(0)) = o
            End If
        End If
    Next k
    I_ANAHTAR_SIRALA depolar
    t = Array(eskiT, yeniT, depolar, eksik)
    mAdmOnbellek(imza) = t
    ADM_DEMONTAJ = t
End Function

'---------- rapor ----------
Private Sub ADM_KUR(ByVal ws As Worksheet)
    Dim b As Variant, iller As Variant, secIl As String, u As Long, p As Long, il As String, veri As Variant, i As Long, j As Long, r As Long
    Dim ad(0 To 3) As String, deger(0 To 3) As Double, satir(0 To 3) As Double, depoSay(0 To 3) As Long
    Dim fiyatsiz(0 To 3) As Double, hurda(0 To 3) As Double, hurdaTl(0 To 3) As Double, enEski(0 To 3) As Date, enYeni(0 To 3) As Date
    Dim ilDeger(0 To 4, 0 To 3) As Double, ilTarih(0 To 4, 0 To 3) As Date, ilIsgZ(0 To 4) As Double, ilIsgS(0 To 4) As Double
    Dim ilIsgZK(0 To 4) As Double, ilIsgSK(0 To 4) As Double, ilHurda(0 To 4) As Double, ilEksi(0 To 4) As Double, ilTekil(0 To 4) As Double
    Dim ilDemK(0 To 4) As Double, ilDemA(0 To 4) As Double, ilDemAz(0 To 4) As Double, ilDemG(0 To 4) As Double, ilDemT(0 To 4) As String
    Dim ilDemTl(0 To 4) As Double, demToplamTl As Double, enler As Variant
    Dim tekil As Object, tekilSec As Object, k As Variant, kx As Variant, t As Variant, uy As Collection, tekilToplam As Double, tumToplam As Double
    Dim isgSatir As Collection, dem As Variant, demEksik As Collection, sh As Worksheet, gun As Long, secili As Boolean
    Dim krit As Long, uyari As Long, bilgi As Long, skor As Long, ozet As String, oncelik As String, n As Long, renk As Long
    Dim sira() As Variant, satirNo As Long, s As Shape, malz As Object, parca() As String, x As Double, d As Double, ilMetin As String
    Dim tAsim As Long, tDikkat As Long, tGirilmedi As Long, t45 As Long, ter As Long, terIl As String, enBirim As Long, isgZ As Double
    Dim kategori As Object, aramaMetni As String, demToplamK As Double, demToplamA As Double
    Dim aylik As Object, ayCur As Long
    Set uy = New Collection: Set isgSatir = New Collection: Set demEksik = New Collection
    Set tekil = CreateObject("Scripting.Dictionary"): tekil.CompareMode = vbTextCompare
    Set tekilSec = CreateObject("Scripting.Dictionary"): tekilSec.CompareMode = vbTextCompare
    b = ADM_BIRIMLER(): iller = ADM_ILLER()
    secIl = ADM_IL_KOD(CStr(ws.Range("BX3").Value2))
    aramaMetni = CStr(ws.Range("BX4").Value2)
    If Len(secIl) = 0 Then ilMetin = "Tüm iller" Else ilMetin = Y_IL_ADI(secIl)
    '--- birim x il verileri (il karsilastirmasi icin her zaman bes il)
    For u = 0 To 3
        ad(u) = CStr(b(u)(0))
        For p = 0 To 4
            il = CStr(iller(p))
            secili = (Len(secIl) = 0 Or il = secIl)
            ADM_ILERLE 0.05 + 0.6 * (u * 5 + p) / 20, ad(u) & "  •  " & Y_IL_ADI(il) & " okunuyor..."
            veri = ADM_VERI(u, il)
            ilTarih(p, u) = veri(0)
            If veri(2) > 0 Then
                ilDeger(p, u) = veri(3)(1)
                If u = 0 Then ilHurda(p) = veri(3)(4)
                ilEksi(p) = ilEksi(p) + veri(3)(2)
                For Each k In veri(4).keys
                    kx = il & "|" & k
                    If Not tekil.Exists(kx) Then
                        tekil(kx) = Array(u, veri(4)(k)(0), veri(4)(k)(1), veri(4)(k)(2), CStr(b(u)(4)), il, CStr(k))
                    Else
                        t = tekil(kx)
                        If InStr(t(4), CStr(b(u)(4))) = 0 Then t(4) = t(4) & " + " & CStr(b(u)(4))
                        If veri(4)(k)(2) > t(3) Then t(0) = u: t(1) = veri(4)(k)(0): t(2) = veri(4)(k)(1): t(3) = veri(4)(k)(2)
                        tekil(kx) = t
                    End If
                Next k
                If u = 3 Then
                    For Each k In veri(5).keys
                        ilIsgZ(p) = ilIsgZ(p) + veri(5)(k)(2): ilIsgS(p) = ilIsgS(p) + veri(5)(k)(4)
                        ilIsgZK(p) = ilIsgZK(p) + veri(5)(k)(1): ilIsgSK(p) = ilIsgSK(p) + veri(5)(k)(3)
                        If secili Then isgSatir.Add Array(il, CStr(k), veri(5)(k)(0), veri(5)(k)(1), veri(5)(k)(2), veri(5)(k)(3), veri(5)(k)(4))
                    Next k
                End If
            End If
            If secili Then
                If veri(2) > 0 Then
                    satir(u) = satir(u) + veri(3)(0): deger(u) = deger(u) + veri(3)(1)
                    fiyatsiz(u) = fiyatsiz(u) + veri(3)(3)
                    hurda(u) = hurda(u) + veri(3)(4): hurdaTl(u) = hurdaTl(u) + veri(3)(5)
                    depoSay(u) = depoSay(u) + veri(4).Count
                End If
                If veri(0) > 0 Then
                    If enEski(u) = 0 Or veri(0) < enEski(u) Then enEski(u) = veri(0)
                    If veri(0) > enYeni(u) Then enYeni(u) = veri(0)
                End If
                gun = ADM_GUN(CDate(veri(0)))
                If gun < 0 Then
                    ADM_UYAR uy, 3, ad(u) & " • " & Y_IL_ADI(il) & ": henüz güncelleme / kayýt yok.", ad(u), "Güncellik"
                ElseIf gun >= 3 Then
                    ADM_UYAR uy, 1, ad(u) & " • " & Y_IL_ADI(il) & ": veriler " & gun & " gündür güncellenmedi (son " & Format$(veri(0), "dd.mm.yyyy hh:nn") & ").", ad(u), "Güncellik"
                ElseIf gun >= 1 Then
                    ADM_UYAR uy, 2, ad(u) & " • " & Y_IL_ADI(il) & ": son güncelleme " & gun & " gün önce.", ad(u), "Güncellik"
                End If
                If veri(3)(2) > 0 Then ADM_UYAR uy, 1, ad(u) & " • " & Y_IL_ADI(il) & ": " & veri(3)(2) & " stok satýrý EKSÝ miktarda; SAP kayýtlarý kontrol edilmeli.", ad(u), "Eksi stok"
                If veri(3)(3) > 0 Then ADM_UYAR uy, 2, ad(u) & " • " & Y_IL_ADI(il) & ": " & AKSORA_SAYI_METNI(veri(3)(3), "#,##0") & " stok satýrýnýn fiyatý yok; deðer eksik hesaplanýyor.", ad(u), "Fiyat"
            End If
        Next p
    Next u
    For Each k In tekil.keys
        tumToplam = tumToplam + tekil(k)(3)
        For p = 0 To 4
            If tekil(k)(5) = iller(p) Then ilTekil(p) = ilTekil(p) + tekil(k)(3)
        Next p
        If Len(secIl) = 0 Or tekil(k)(5) = secIl Then tekilSec(k) = tekil(k): tekilToplam = tekilToplam + tekil(k)(3)
    Next k
    If hurda(0) > 0 Then ADM_UYAR uy, 3, "Malzeme • " & ilMetin & ": HURDA partide " & AKSORA_SAYI_METNI(hurda(0), "#,##0") & " kalem (" & ADM_TL(hurdaTl(0)) & ") bekliyor; hurda satýþ / imha planlanabilir.", ad(0), "Hurda"
    '--- demontaj (her il, ilce gecmisinden)
    For p = 0 To 4
        il = CStr(iller(p))
        ADM_ILERLE 0.66 + 0.04 * p, Y_IL_ADI(il) & "  •  demontaj kontrolü..."
        dem = ADM_DEMONTAJ(il)
        ilDemT(p) = IIf(Len(dem(0)) > 0, Left$(dem(0), 16) & "  >  " & Left$(dem(1), 16), "")
        For Each k In dem(2).keys
            ilDemAz(p) = ilDemAz(p) + dem(2)(k)(1): ilDemG(p) = ilDemG(p) + dem(2)(k)(2)
            ilDemK(p) = ilDemK(p) + dem(2)(k)(3): ilDemA(p) = ilDemA(p) + dem(2)(k)(4): ilDemTl(p) = ilDemTl(p) + dem(2)(k)(5)
        Next k
        If Len(secIl) = 0 Or il = secIl Then
            For Each t In dem(3)
                demEksik.Add t
            Next t
            demToplamK = demToplamK + ilDemK(p): demToplamA = demToplamA + ilDemA(p): demToplamTl = demToplamTl + ilDemTl(p)
            If ilDemK(p) > 0 Then
                ' en cok eksik kalemli 3 ilce
                ReDim sira(1 To dem(2).Count)
                n = 0
                For Each k In dem(2).keys
                    n = n + 1: sira(n) = Array(dem(2)(k)(0), dem(2)(k)(3))
                Next k
                t = vbNullString
                For i = 1 To application.Min(3, n)
                    For j = i + 1 To n
                        If sira(j)(1) > sira(i)(1) Then enler = sira(i): sira(i) = sira(j): sira(j) = enler
                    Next j
                    If sira(i)(1) > 0 Then t = t & IIf(i > 1, ", ", "") & sira(i)(0) & " (" & sira(i)(1) & " kalem)"
                Next i
                ADM_UYAR uy, IIf(ilDemK(p) >= 10, 1, 2), Y_IL_ADI(il) & ": azalan " & ilDemK(p) & " kalem / " & ADM_ADET(ilDemA(p)) & _
                    " adet malzemenin (" & ADM_TL_KISA(ilDemTl(p)) & ") DEMONTAJ GÝRÝÞÝ YOK! En çok: " & t & ".", "Ýlçe | Demontaj", "Demontaj"
            End If
        End If
    Next p
    '--- aylik demontaj defteri (her il)
    Set aylik = CreateObject("Scripting.Dictionary")
    For p = 0 To 4
        il = CStr(iller(p))
        ADM_ILERLE 0.83 + 0.01 * p, Y_IL_ADI(il) & "  •  aylýk demontaj kayýtlarý..."
        I_DEFTER_GUNCELLE il
        aylik.Add il, I_DEFTER_OKU(il)
        For Each k In aylik(il).keys
            If CLng(k) > ayCur Then ayCur = CLng(k)
        Next k
    Next p
    '--- ISG (il bazinda tek uyari)
    For p = 0 To 4
        If Len(secIl) = 0 Or iller(p) = secIl Then
            isgZ = isgZ + ilIsgZ(p)
            If ilIsgZ(p) > 0 Then
                ReDim sira(1 To application.Max(1, isgSatir.Count))
                n = 0
                For Each kx In isgSatir
                    If kx(0) = iller(p) Then n = n + 1: sira(n) = Array(kx(2), kx(4))
                Next kx
                t = vbNullString
                For i = 1 To application.Min(3, n)
                    For j = i + 1 To n
                        If sira(j)(1) > sira(i)(1) Then enler = sira(i): sira(i) = sira(j): sira(j) = enler
                    Next j
                    If sira(i)(1) > 0 Then t = t & IIf(i > 1, ", ", "") & sira(i)(0) & " (" & ADM_ADET(sira(i)(1)) & ")"
                Next i
                ADM_UYAR uy, IIf(ilIsgZ(p) >= 300, 1, 2), Y_IL_ADI(CStr(iller(p))) & ": ilçe depolarýnda " & ADM_ADET(ilIsgZ(p)) & " adet (" & ilIsgZK(p) & _
                    " kalem) ÝSG malzemesi ZÝMMETLENMEMÝÞ!" & IIf(Len(t) > 0, " En yüksek: " & t & ".", ""), "Ýlçe | ÝSG", "ÝSG"
            End If
            If ilIsgS(p) > 0 Then ADM_UYAR uy, 3, Y_IL_ADI(CStr(iller(p))) & ": " & ADM_ADET(ilIsgS(p)) & " adet (" & ilIsgSK(p) & " kalem) ÝSG malzemesi manuel sarf bekliyor.", "Ýlçe | ÝSG", "ÝSG sarf"
        End If
    Next p
    '--- yatirim teminat (ekrandaki il)
    terIl = ADM_CANLI_IL(1)
    Set sh = ADM_WS("Y_TEMÝNAT")
    If Not sh Is Nothing And (Len(secIl) = 0 Or terIl = secIl) Then
        r = 10
        Do While r < 200
            If Len(ADM_METIN(sh.Cells(r, 1).Value2)) = 0 Then Exit Do
            t = UCase$(Replace(Replace(ADM_METIN(sh.Cells(r, 9).Value2), "Þ", "S"), "Ý", "I"))
            If InStr(t, "ASIM") > 0 Then
                tAsim = tAsim + 1
                ADM_UYAR uy, 1, "Teminat aþýmý • " & Y_IL_ADI(terIl) & ": " & ADM_METIN(sh.Cells(r, 2).Value2) & " (" & ADM_METIN(sh.Cells(r, 1).Value2) & ") teminat tutarýný aþmýþ.", "Yatýrým | Teminat", "Teminat"
            ElseIf InStr(t, "DIKKAT") > 0 Then
                tDikkat = tDikkat + 1
                ADM_UYAR uy, 2, "Teminat sýnýrýna yakýn • " & Y_IL_ADI(terIl) & ": " & ADM_METIN(sh.Cells(r, 2).Value2) & " (" & ADM_METIN(sh.Cells(r, 1).Value2) & ").", "Yatýrým | Teminat", "Teminat"
            ElseIf InStr(t, "GIRILMEDI") > 0 Then
                tGirilmedi = tGirilmedi + 1
            End If
            If ADM_SAYI(sh.Cells(r, 10).Value2) > 0 Then
                t45 = t45 + 1
                ADM_UYAR uy, 2, Y_IL_ADI(terIl) & " • " & ADM_METIN(sh.Cells(r, 2).Value2) & ": sahada 45 günü geçen " & ADM_ADET(ADM_SAYI(sh.Cells(r, 10).Value2)) & " kalem; hakediþ bekliyor.", "Yatýrým | Teminat", "Teminat"
            End If
            ter = ter + 1
            r = r + 1
        Loop
        If tGirilmedi > 0 Then ADM_UYAR uy, 3, Y_IL_ADI(terIl) & ": " & tGirilmedi & " yüklenici deposunda teminat tutarý girilmemiþ; teminat riski ölçülemiyor.", "Yatýrým | Teminat", "Teminat"
    End If
    '--- malzeme gun bazli karsilastirma (ekrandaki il)
    Set sh = ADM_WS("STOK_KARÞILAÞTIRMA")
    If Not sh Is Nothing And (Len(secIl) = 0 Or ADM_CANLI_IL(0) = secIl) Then
        If ADM_SAYI(sh.Range("J10").Value2) > 0 Then ADM_UYAR uy, 2, "Malzeme • " & Y_IL_ADI(ADM_CANLI_IL(0)) & ": son karþýlaþtýrmada (" & ADM_METIN(sh.Range("B6").Value2) & " - " & _
            ADM_METIN(sh.Range("I6").Value2) & ") " & ADM_SAYI(sh.Range("J10").Value2) & " malzemenin stoðu sýfýrlandý.", "Malzeme | Karþýlaþtýrma", "Karþýlaþtýrma"
        d = ADM_SAYI(sh.Range("G12").Value2)
        If d < 0 Then ADM_UYAR uy, 3, "Malzeme • " & Y_IL_ADI(ADM_CANLI_IL(0)) & ": son karþýlaþtýrmada stok deðeri " & ADM_TL(-d) & " azaldý (" & _
            ADM_SAYI(sh.Range("A10").Value2) & " azalan, " & ADM_SAYI(sh.Range("D10").Value2) & " artan kalem).", "Malzeme | Karþýlaþtýrma", "Karþýlaþtýrma"
    End If
    If Len(ADM_AD_OKU("_AKSORA_ADMIN_SON_TOPLU")) > 0 Then ADM_UYAR uy, 3, "Son tek tuþ güncelleme: " & ADM_AD_OKU("_AKSORA_ADMIN_SON_TOPLU"), "Yönetim", "Yönetim"
    ' Son basarili giristen once yapilan hatali sifre denemeleri
    parca = Split(ADM_AD_OKU("_AKSORA_ADMIN_LOG"), ";")
    n = 0
    For i = 1 To UBound(parca)
        If Right$(parca(i), 6) = "HATALI" Then
            n = n + 1
        ElseIf Right$(parca(i), 5) = "GÝRÝÞ" Then
            Exit For
        End If
    Next i
    If n > 0 Then ADM_UYAR uy, 2, "Güvenlik: bu giriþten önce " & n & " hatalý yönetici þifresi denemesi yapýldý (" & Split(parca(1), "|")(0) & ", " & Split(parca(1), "|")(1) & ").", "Yönetim | Güvenlik", "Güvenlik"
    '--- skor: her kategori bir kez (kritik -15, uyari -6)
    Set kategori = CreateObject("Scripting.Dictionary")
    For Each kx In uy
        Select Case kx(0)
            Case 1: krit = krit + 1
            Case 2: uyari = uyari + 1
            Case Else: bilgi = bilgi + 1
        End Select
        If Not kategori.Exists(kx(3)) Then
            kategori(kx(3)) = kx(0)
        ElseIf kx(0) < kategori(kx(3)) Then
            kategori(kx(3)) = kx(0)
        End If
    Next kx
    skor = 100
    For Each k In kategori.keys
        If kategori(k) = 1 Then
            skor = skor - 15
        ElseIf kategori(k) = 2 Then
            skor = skor - 6
        End If
    Next k
    If skor < 0 Then skor = 0
    For u = 1 To 3
        If deger(u) > deger(enBirim) Then enBirim = u
    Next u
    ozet = ilMetin & " için tekil stok deðeri " & ADM_TL_KISA(tekilToplam) & " (birimler arasý ortak depolar bir kez). En büyük birim " & ad(enBirim) & " (" & ADM_TL_KISA(deger(enBirim)) & ")."
    If demToplamK > 0 Then ozet = ozet & " Demontaj giriþi yapýlmayan " & demToplamK & " kalem / " & ADM_ADET(demToplamA) & " adet (" & ADM_TL_KISA(demToplamTl) & ") malzeme var."
    If isgZ > 0 Then ozet = ozet & " Ýlçelerde " & ADM_ADET(isgZ) & " adet ÝSG malzemesi zimmet bekliyor."
    n = 0
    For j = 1 To 2
        For Each kx In uy
            If kx(0) = j And n < 3 Then n = n + 1: oncelik = oncelik & vbLf & "  " & n & ")  " & kx(1)
        Next kx
    Next j
    If Len(oncelik) > 0 Then ozet = ozet & vbLf & "Öncelikli aksiyonlar:" & oncelik Else ozet = ozet & " Kritik veya dikkat gerektiren konu yok."
    '================== CIZIM ==================
    ADM_ILERLE 0.9, "Rapor çiziliyor..."
    application.ScreenUpdating = False
    ws.Unprotect ""
    ws.AutoFilterMode = False
    ' Bekleme karti (ADMY_) silinmez; donma boyunca ekranda kalir, en sonda AKSORA_BEKLE_KAPAT kaldirir.
    For i = ws.Shapes.Count To 1 Step -1
        If Left$(ws.Shapes(i).Name, 5) <> "ADMY_" Then ws.Shapes(i).Delete
    Next i
    ws.Cells.Clear
    ws.Cells.Validation.Delete
    ws.Cells.FormatConditions.Delete
    ws.Rows.Hidden = False
    ws.Cells.Font.Name = "Aptos": ws.Cells.Font.Size = 10: ws.Cells.Font.Color = RGB(30, 41, 59)
    ws.Cells.Interior.Color = RGB(244, 248, 252)
    ws.Cells.VerticalAlignment = xlCenter
    ws.Columns("A").ColumnWidth = 2: ws.Columns("N").ColumnWidth = 2
    ws.Columns("B:M").ColumnWidth = 14
    ' Malzeme tanimi sutunu (G) genis: arama tablosu birlesik hucresiz ve filtrelenebilir olur.
    ws.Columns("G").ColumnWidth = 30
    ws.Rows.RowHeight = 20
    ws.Tab.Color = RGB(12, 45, 79)
    ws.Range("BX1").Value2 = "AKSORA_ADMIN_V4": ws.Range("BX3").Value2 = secIl: ws.Range("BX4").Value2 = aramaMetni
    '--- baslik
    ws.Rows("1:2").RowHeight = 26: ws.Rows(3).RowHeight = 20
    ADM_BLOK ws, "A1:N2", "   AKSORA  |  YÖNETÝM MERKEZÝ", RGB(12, 45, 79), vbWhite, 22, True, xlLeft
    ADM_BLOK ws, "A3:N3", "      Çoruh EDAÞ  •  " & ilMetin & "  •  Rapor: " & Format$(Now, "dd.mm.yyyy hh:nn") & _
        "  •  Tekil stok deðeri: " & ADM_TL(tekilToplam), RGB(23, 66, 107), RGB(198, 240, 97), 10, True, xlLeft
    ADM_DUGME ws, "ADM_TUM", "TÜM BÝRÝMLERÝ GÜNCELLE", ws.Range("H1").Left, RGB(0, 125, 110), "AKSORA_ADMIN_TUMUNU_GUNCELLE", 1.6
    ADM_DUGME ws, "ADM_YENILE", "RAPORU YENÝLE", ws.Range("J1").Left - 8, RGB(20, 103, 173), "AKSORA_ADMIN_YENILE", 1
    ADM_DUGME ws, "ADM_CIKTI", "ÇIKTI AL", ws.Range("K1").Left, RGB(15, 118, 110), "AKSORA_ADMIN_CIKTI", 1
    ADM_DUGME ws, "ADM_CIKIS", "BÝRÝM SEÇÝMÝNE DÖN", ws.Range("L1").Left + 8, RGB(30, 41, 59), "AKSORA_ADMIN_CIKIS", 1.4
    '--- il secimi (dugmeler)
    ws.Rows(4).RowHeight = 34: ws.Rows(5).RowHeight = 8
    ws.Range("A4:N4").Interior.Color = RGB(226, 236, 247)
    ADM_BLOK ws, "B4", "ÝL:", RGB(226, 236, 247), RGB(12, 45, 79), 11, True, xlRight
    ADM_IL_DUGMELERI ws, secIl
    ADM_BLOK ws, "J4:M4", "Seçili: " & ilMetin & "   •   her birimin ekrandaki ili canlý, diðer iller il kayýtlarýndan okunur", RGB(226, 236, 247), RGB(71, 85, 105), 8, False, xlRight
    ws.Range("J4").Font.Italic = True
    '--- arama (en ustte)
    ADM_BOLUM ws, 6, "B6:M6", "MALZEME / SERÝ NUMARASI ARAMA   •   " & ilMetin & "  •  tüm birimler"
    ws.Rows(7).RowHeight = 30
    ADM_BLOK ws, "B7", "ARANAN:", RGB(244, 248, 252), RGB(12, 45, 79), 10, True, xlRight
    ADM_BLOK ws, "C7:H7", aramaMetni, vbWhite, RGB(12, 45, 79), 12, True, xlLeft
    With ws.Range("C7:H7")
        .Locked = False
        .Borders.LineStyle = xlContinuous: .Borders.Color = RGB(20, 103, 173): .Borders.Weight = xlMedium
        .IndentLevel = 1
    End With
    ADM_DUGME ws, "ADM_ARA", "ARA", ws.Range("I7").Left + 6, RGB(12, 45, 79), "AKSORA_ADMIN_ARA", 0.9, ws.Range("I7").Top + 3, ws.Range("I7").Height - 6
    ADM_DUGME ws, "ADM_ARA_SIL", "TEMÝZLE", ws.Range("J7").Left - 4, RGB(100, 116, 139), "AKSORA_ADMIN_ARA_TEMIZLE", 0.9, ws.Range("J7").Top + 3, ws.Range("J7").Height - 6
    ADM_BLOK ws, "K7:M7", "Kod, ad veya seri no. Birden çok kelime: hepsi aranýr (örn. kablo 3x95). Enter'a basýn.", RGB(244, 248, 252), RGB(100, 116, 139), 8, False, xlLeft
    ws.Range("K7").Font.Italic = True: ws.Range("K7").WrapText = True
    ws.Rows(ADM_ARA_SATIR & ":" & ADM_ARA_SATIR + ADM_ARA_N + 1).Hidden = True
    r = ADM_ARA_SATIR + ADM_ARA_N + 3
    '--- demontaj kontrolu
    ADM_ILERLE 0.91, "Demontaj tablosu çiziliyor..."
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "DEMONTAJ KONTROLÜ   •   " & ilMetin & "  (azalan malzeme - ÇIKMA/BDSZ giriþi; ÝSG, sayaç, modem ve mühür hariç)"
    r = r + 1: satirNo = r
    If Len(secIl) = 0 Then
        ADM_BASLIKLAR ws, r, 2, Array("Ýl", "Karþýlaþtýrýlan günler", "", "", "Azalan kalem", "Giriþi yapýlan", "Giriþi YOK kalem", "Eksik adet", "Eksik deðer", "Durum", "", "")
        ws.Range(ADM_ADRES(r, 3, r, 5)).Merge: ws.Range(ADM_ADRES(r, 11, r, 13)).Merge
        For p = 0 To 4
            r = r + 1
            ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
            ws.Cells(r, 2).Value2 = Y_IL_ADI(CStr(iller(p))): ws.Cells(r, 2).Font.bold = True
            ws.Range(ADM_ADRES(r, 3, r, 5)).Merge: ws.Cells(r, 3).Value2 = IIf(Len(ilDemT(p)) > 0, ilDemT(p), "en az iki günlük ilçe stoðu gerekir")
            ws.Cells(r, 6).Value2 = ilDemAz(p): ws.Cells(r, 7).Value2 = ilDemG(p): ws.Cells(r, 8).Value2 = ilDemK(p): ws.Cells(r, 9).Value2 = ilDemA(p)
            ws.Cells(r, 10).Value2 = ilDemTl(p): ws.Cells(r, 10).NumberFormat = "#,##0.00 ""TL"""
            ADM_MIKTAR ws.Range(ADM_ADRES(r, 6, r, 9))
            ws.Range(ADM_ADRES(r, 2, r, 10)).HorizontalAlignment = xlCenter
            If Len(ilDemT(p)) = 0 Then
                ADM_BLOK ws, ADM_ADRES(r, 11, r, 13), "KARÞILAÞTIRMA YOK", RGB(241, 245, 249), RGB(100, 116, 139), 9, True, xlCenter
            ElseIf ilDemK(p) > 0 Then
                ADM_BLOK ws, ADM_ADRES(r, 11, r, 13), "DEMONTAJ GÝRÝÞÝ YAPILMALI", RGB(254, 226, 226), RGB(185, 28, 28), 9, True, xlCenter
                ws.Cells(r, 8).Font.Color = RGB(185, 28, 28): ws.Cells(r, 8).Font.bold = True
            Else
                ADM_BLOK ws, ADM_ADRES(r, 11, r, 13), "TAMAM", RGB(220, 252, 231), RGB(21, 128, 61), 9, True, xlCenter
            End If
        Next p
    Else
        dem = ADM_DEMONTAJ(secIl)
        ADM_BASLIKLAR ws, r, 2, Array("Ýlçe deposu", "", "", "Azalan kalem", "Giriþi yapýlan", "Giriþi YOK kalem", "Eksik adet", "Eksik deðer", "Durum", "", "", "")
        ws.Range(ADM_ADRES(r, 2, r, 4)).Merge: ws.Range(ADM_ADRES(r, 10, r, 13)).Merge
        If Len(dem(0)) = 0 Then
            r = r + 1
            ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), Y_IL_ADI(secIl) & " için karþýlaþtýrma yok: ilçe stoðu en az iki farklý günde güncellenmiþ olmalý.", vbWhite, RGB(100, 116, 139), 10, False, xlLeft
        ElseIf dem(2).Count = 0 Then
            r = r + 1
            ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), dem(0) & "  >  " & dem(1) & " arasýnda azalan malzeme yok.", RGB(220, 252, 231), RGB(21, 128, 61), 10, True, xlLeft
        Else
            For Each k In dem(2).keys
                r = r + 1
                ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
                ADM_BLOK ws, ADM_ADRES(r, 2, r, 4), k & "  " & dem(2)(k)(0), vbWhite, RGB(30, 41, 59), 10, False, xlLeft
                ws.Cells(r, 5).Value2 = dem(2)(k)(1): ws.Cells(r, 6).Value2 = dem(2)(k)(2): ws.Cells(r, 7).Value2 = dem(2)(k)(3): ws.Cells(r, 8).Value2 = dem(2)(k)(4)
                ws.Cells(r, 9).Value2 = dem(2)(k)(5): ws.Cells(r, 9).NumberFormat = "#,##0.00 ""TL"""
                ADM_MIKTAR ws.Range(ADM_ADRES(r, 5, r, 8))
                ws.Range(ADM_ADRES(r, 5, r, 9)).HorizontalAlignment = xlCenter
                If dem(2)(k)(3) > 0 Then
                    ADM_BLOK ws, ADM_ADRES(r, 10, r, 13), "DEMONTAJ GÝRÝÞÝ YAPILMALI", RGB(254, 226, 226), RGB(185, 28, 28), 9, True, xlCenter
                    ws.Cells(r, 7).Font.Color = RGB(185, 28, 28): ws.Cells(r, 7).Font.bold = True
                Else
                    ADM_BLOK ws, ADM_ADRES(r, 10, r, 13), "TAMAM", RGB(220, 252, 231), RGB(21, 128, 61), 9, True, xlCenter
                End If
            Next k
            r = r + 1
            ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Karþýlaþtýrýlan: " & dem(0) & "  >  " & dem(1), RGB(244, 248, 252), RGB(100, 116, 139), 8, False, xlLeft
            ws.Cells(r, 2).Font.Italic = True
        End If
    End If
    ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 13)), True
    ' girisi yapilmayan malzemeler (en buyuk eksik adet)
    If demEksik.Count > 0 Then
        r = r + 2: satirNo = r
        ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "DEMONTAJ GÝRÝÞÝ YAPILMAYAN MALZEMELER   •   " & ilMetin & "  (kritik malzemeler önce: kablo, direk, pano, hücre, trafo; ilk 25 / toplam " & demEksik.Count & ")"
        r = r + 1: satirNo = r
        ADM_BASLIKLAR ws, r, 2, Array("Ýl", "Depo", "Ýlçe", "", "Malzeme", "Malzeme tanýmý", "Eksik", "Ölçü", "Eksik deðer", "Durum", "", "")
        ws.Range(ADM_ADRES(r, 4, r, 5)).Merge: ws.Range(ADM_ADRES(r, 11, r, 13)).Merge
        ReDim sira(1 To demEksik.Count)
        i = 0
        For Each t In demEksik
            i = i + 1: sira(i) = t
        Next t
        For i = 1 To application.Min(25, demEksik.Count)
            For j = i + 1 To demEksik.Count
                If ADM_DEM_SKOR(sira(j)) > ADM_DEM_SKOR(sira(i)) Then t = sira(i): sira(i) = sira(j): sira(j) = t
            Next j
        Next i
        For i = 1 To application.Min(25, demEksik.Count)
            r = r + 1
            ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
            ws.Range(ADM_ADRES(r, 2, r, 7)).NumberFormat = "@"
            ws.Cells(r, 2).Value2 = Y_IL_ADI(CStr(sira(i)(0))): ws.Cells(r, 3).Value2 = sira(i)(1)
            ws.Range(ADM_ADRES(r, 4, r, 5)).Merge: ws.Cells(r, 4).Value2 = sira(i)(2)
            ws.Cells(r, 6).Value2 = sira(i)(3)
            ws.Cells(r, 7).Value2 = sira(i)(4): ws.Cells(r, 7).HorizontalAlignment = xlLeft
            ws.Cells(r, 8).Value2 = sira(i)(5): ws.Cells(r, 8).Font.bold = True: ws.Cells(r, 8).Font.Color = RGB(185, 28, 28)
            ADM_MIKTAR ws.Cells(r, 8)
            ws.Cells(r, 9).Value2 = sira(i)(6)
            ws.Cells(r, 10).Value2 = sira(i)(8): ws.Cells(r, 10).NumberFormat = "#,##0.00 ""TL"""
            If sira(i)(7) = "EKSÝK GÝRÝÞ" Then
                ADM_BLOK ws, ADM_ADRES(r, 11, r, 13), ADM_DEM_DURUM(sira(i)), RGB(254, 243, 199), RGB(180, 83, 9), 8, True, xlCenter
                ws.Cells(r, 8).Font.Color = RGB(180, 83, 9)
            Else
                ADM_BLOK ws, ADM_ADRES(r, 11, r, 13), ADM_DEM_DURUM(sira(i)), RGB(254, 226, 226), RGB(185, 28, 28), 8, True, xlCenter
            End If
            ws.Range(ADM_ADRES(r, 2, r, 6)).HorizontalAlignment = xlCenter: ws.Cells(r, 4).HorizontalAlignment = xlLeft
            ws.Range(ADM_ADRES(r, 8, r, 9)).HorizontalAlignment = xlCenter
        Next i
        ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 13)), True
    End If
    '--- aylik demontaj raporu
    ADM_ILERLE 0.93, "Aylýk demontaj raporu çiziliyor..."
    If ayCur > 0 Then ADM_AYLIK_CIZ ws, r, aylik, ayCur, secIl, ilMetin
    '--- ozet + skor
    r = r + 2
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 8), "YÖNETÝCÝ ÖZETÝ  •  " & ilMetin
    ADM_BOLUM ws, r, ADM_ADRES(r, 10, r, 13), "AKSORA SAÐLIK SKORU"
    ws.Range(ADM_ADRES(r + 1, 1, r + 4, 1)).EntireRow.RowHeight = 30
    ADM_BLOK ws, ADM_ADRES(r + 1, 2, r + 4, 8), ozet, vbWhite, RGB(30, 41, 59), 10, False, xlLeft
    ws.Cells(r + 1, 2).WrapText = True: ws.Cells(r + 1, 2).VerticalAlignment = xlTop
    ADM_KENAR ws.Range(ADM_ADRES(r + 1, 2, r + 4, 8))
    If skor >= 80 Then
        renk = RGB(21, 128, 61)
    ElseIf skor >= 55 Then
        renk = RGB(180, 83, 9)
    Else
        renk = RGB(185, 28, 28)
    End If
    ADM_BLOK ws, ADM_ADRES(r + 1, 10, r + 3, 13), skor & " / 100", renk, vbWhite, 30, True, xlCenter
    ADM_BLOK ws, ADM_ADRES(r + 4, 10, r + 4, 13), IIf(skor >= 80, "ÝYÝ", IIf(skor >= 55, "DÝKKAT", "KRÝTÝK")) & "   •   " & kategori.Count & " konu  •  " & krit & " kritik  •  " & uyari & " uyarý", _
        vbWhite, renk, 10, True, xlCenter
    ADM_KENAR ws.Range(ADM_ADRES(r + 1, 10, r + 4, 13))
    r = r + 6
    '--- birim kartlari
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "BÝRÝM ÖZETÝ  •  " & ilMetin
    ws.Rows(r + 4).RowHeight = 30
    For u = 0 To 3
        j = 2 + u * 3
        ADM_BLOK ws, ADM_ADRES(r + 1, j, r + 1, j + 2), ad(u), CLng(b(u)(1)), vbWhite, 11, True, xlLeft
        If Len(secIl) = 0 Then
            t = "5 il  •  güncelleme: " & IIf(enEski(u) > 0, Format$(enEski(u), "dd.mm hh:nn") & " - " & Format$(enYeni(u), "dd.mm hh:nn"), "-")
        Else
            t = ilMetin & "  •  Son güncelleme: " & IIf(enYeni(u) > 0, Format$(enYeni(u), "dd.mm.yyyy hh:nn"), "-")
        End If
        ADM_BLOK ws, ADM_ADRES(r + 2, j, r + 2, j + 2), CStr(t), vbWhite, RGB(71, 85, 105), 9, False, xlLeft
        If enEski(u) = 0 Then
            ADM_BLOK ws, ADM_ADRES(r + 3, j, r + 3, j + 2), "VERÝ YOK", vbWhite, RGB(185, 28, 28), 10, True, xlLeft
        Else
            gun = ADM_GUN(enEski(u))
            ADM_BLOK ws, ADM_ADRES(r + 3, j, r + 3, j + 2), IIf(gun <= 0, "GÜNCEL", IIf(Len(secIl) = 0, "EN ESKÝ ÝL ", "") & gun & " GÜN ÖNCE"), vbWhite, _
                IIf(gun <= 0, RGB(21, 128, 61), IIf(gun < 3, RGB(180, 83, 9), RGB(185, 28, 28))), 10, True, xlLeft
        End If
        ADM_BLOK ws, ADM_ADRES(r + 4, j, r + 4, j + 2), ADM_TL(deger(u)), vbWhite, RGB(12, 45, 79), 16, True, xlLeft
        ADM_BLOK ws, ADM_ADRES(r + 5, j, r + 5, j + 2), AKSORA_SAYI_METNI(satir(u), "#,##0") & " stok satýrý  •  " & depoSay(u) & " depo", vbWhite, RGB(71, 85, 105), 9, False, xlLeft
        Select Case u
            Case 0: t = IIf(hurda(u) > 0, "HURDA partide " & AKSORA_SAYI_METNI(hurda(u), "#,##0") & " kalem", "Hurda bekleyen kalem yok")
            Case 1: t = IIf(ter > 0, Y_IL_ADI(terIl) & ": " & tAsim & " aþým  •  " & tDikkat & " dikkat  •  " & t45 & " depoda 45+ gün", "Teminat verisi yok")
            Case 2: t = IIf(fiyatsiz(u) > 0, AKSORA_SAYI_METNI(fiyatsiz(u), "#,##0") & " fiyatsýz satýr", "Fiyatsýz satýr yok")
            Case Else: t = ADM_ADET(isgZ) & " zimmet bekliyor  •  " & demToplamK & " demontaj eksik"
        End Select
        ADM_BLOK ws, ADM_ADRES(r + 6, j, r + 6, j + 2), CStr(t), RGB(248, 250, 252), RGB(51, 65, 85), 9, True, xlLeft
        ADM_KENAR ws.Range(ADM_ADRES(r + 1, j, r + 6, j + 2))
    Next u
    r = r + 8
    '--- il karsilastirmasi (her zaman bes il)
    ADM_ILERLE 0.95, "Ýl karþýlaþtýrmasý çiziliyor..."
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "ÝL KARÞILAÞTIRMASI   (beþ il • tekil deðer: birimler arasý ortak depolar bir kez)"
    r = r + 1: satirNo = r
    ADM_BASLIKLAR ws, r, 2, Array("Ýl", "Malzeme", "Yatýrým", "Müþteri", "Ýlçe", "Tekil deðer", "Pay", "ÝSG zimmet bekleyen", "Demontaj giriþi yok", "Hurda kalem", "Eksi stok satýrý", "En eski güncelleme")
    For p = 0 To 4
        r = r + 1
        il = CStr(iller(p))
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = IIf(il = secIl, RGB(226, 236, 247), vbWhite)
        ws.Cells(r, 2).Value2 = Y_IL_ADI(il): ws.Cells(r, 2).Font.bold = True
        For u = 0 To 3
            ws.Cells(r, 3 + u).Value2 = ilDeger(p, u)
        Next u
        ws.Cells(r, 7).Value2 = ilTekil(p)
        ws.Cells(r, 8).Value2 = IIf(tumToplam > 0, ilTekil(p) / tumToplam, 0)
        ws.Cells(r, 9).Value2 = ilIsgZ(p): ws.Cells(r, 10).Value2 = ilDemK(p)
        ws.Cells(r, 11).Value2 = ilHurda(p): ws.Cells(r, 12).Value2 = ilEksi(p)
        d = 0
        For u = 0 To 3
            If ilTarih(p, u) > 0 Then
                If d = 0 Or CDbl(ilTarih(p, u)) < d Then d = CDbl(ilTarih(p, u))
            End If
        Next u
        If d > 0 Then
            gun = ADM_GUN(CDate(d))
            ws.Cells(r, 13).Value2 = Format$(CDate(d), "dd.mm.yyyy hh:nn")
            ws.Cells(r, 13).Interior.Color = IIf(gun <= 0, RGB(220, 252, 231), IIf(gun < 3, RGB(254, 243, 199), RGB(254, 226, 226)))
        Else
            ws.Cells(r, 13).Value2 = "veri yok": ws.Cells(r, 13).Interior.Color = RGB(254, 226, 226)
        End If
        For j = 9 To 12
            If ws.Cells(r, j).Value2 > 0 Then ws.Cells(r, j).Font.Color = RGB(185, 28, 28): ws.Cells(r, j).Font.bold = True
        Next j
    Next p
    r = r + 1
    ws.Cells(r, 2).Value2 = "TOPLAM"
    For j = 3 To 12
        If j <> 8 Then ws.Cells(r, j).Formula = "=SUM(" & ADM_ADRES(satirNo + 1, j, r - 1, j) & ")"
    Next j
    ws.Cells(r, 8).Value2 = IIf(tumToplam > 0, 1, 0)
    ws.Range(ADM_ADRES(r, 2, r, 13)).Font.bold = True
    ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = RGB(226, 236, 247)
    ws.Range(ADM_ADRES(satirNo + 1, 3, r, 7)).NumberFormat = "#,##0 ""TL"""
    ws.Range(ADM_ADRES(satirNo + 1, 3, r, 7)).ShrinkToFit = True
    ws.Range(ADM_ADRES(satirNo + 1, 8, r, 8)).NumberFormat = "0.0%"
    ws.Range(ADM_ADRES(satirNo + 1, 9, r, 12)).NumberFormat = "#,##0"
    ws.Range(ADM_ADRES(satirNo + 1, 2, r, 13)).HorizontalAlignment = xlCenter
    ws.Range(ADM_ADRES(satirNo + 1, 3, r, 7)).HorizontalAlignment = xlRight
    ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 13)), True
    On Error Resume Next
    With ws.Range(ADM_ADRES(satirNo + 1, 7, r - 1, 7)).FormatConditions.AddDatabar
        .BarColor.Color = RGB(147, 197, 253)
        .BarFillType = xlDataBarFillSolid
    End With
    On Error GoTo 0
    '--- atil stok (malzeme birimi atil stok kaydindan)
    r = ADM_ATIL_CIZ(ws, r, secIl)
    '--- uyarilar
    r = r + 2
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 10), "YÖNETÝCÝ UYARILARI  •  " & ilMetin
    ADM_BLOK ws, ADM_ADRES(r, 11, r, 13), krit & " kritik  •  " & uyari & " uyarý  •  " & bilgi & " bilgi", RGB(244, 248, 252), RGB(71, 85, 105), 9, True, xlRight
    r = r + 1
    If uy.Count = 0 Then
        ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Kritik bir durum bulunmadý. Birimler güncel ve kontrol altýnda.", RGB(220, 252, 231), RGB(21, 128, 61), 10, True, xlLeft
        r = r + 1
    Else
        For j = 1 To 3
            For Each kx In uy
                If kx(0) = j Then
                    Select Case j
                        Case 1: ADM_BLOK ws, ADM_ADRES(r, 2, r, 2), "KRÝTÝK", RGB(185, 28, 28), vbWhite, 9, True, xlCenter
                        Case 2: ADM_BLOK ws, ADM_ADRES(r, 2, r, 2), "UYARI", RGB(217, 119, 6), vbWhite, 9, True, xlCenter
                        Case Else: ADM_BLOK ws, ADM_ADRES(r, 2, r, 2), "BÝLGÝ", RGB(20, 103, 173), vbWhite, 9, True, xlCenter
                    End Select
                    ADM_BLOK ws, ADM_ADRES(r, 3, r, 11), CStr(kx(1)), vbWhite, RGB(30, 41, 59), 10, (j = 1), xlLeft
                    ws.Cells(r, 3).WrapText = True
                    ADM_BLOK ws, ADM_ADRES(r, 12, r, 13), CStr(kx(2)), vbWhite, RGB(100, 116, 139), 9, False, xlCenter
                    ws.Rows(r).RowHeight = IIf(Len(kx(1)) > 150, 34, 22)
                    ADM_KENAR ws.Range(ADM_ADRES(r, 2, r, 13)), True
                    r = r + 1
                End If
            Next kx
        Next j
    End If
    '--- ISG tablosu + il ozet
    ADM_ILERLE 0.97, "ÝSG ve depo tablolarý çiziliyor..."
    r = r + 1
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 8), IIf(Len(secIl) = 0, "ÝSG ZÝMMET DURUMU  •  il bazýnda", "ÝLÇE ÝSG ZÝMMET DURUMU  •  " & ilMetin)
    ADM_BOLUM ws, r, ADM_ADRES(r, 10, r, 13), "EN ÇOK ZÝMMET BEKLEYEN 5 DEPO"
    r = r + 1: satirNo = r
    ADM_BASLIKLAR ws, r, 2, Array(IIf(Len(secIl) = 0, "Ýl", "Ýlçe deposu"), "", "Zimmet kalem", "Zimmet adet", "Sarf kalem", "Sarf adet", "Durum")
    ws.Range(ADM_ADRES(r, 2, r, 3)).Merge
    If Len(secIl) = 0 Then
        For p = 0 To 4
            r = r + 1
            ADM_BLOK ws, ADM_ADRES(r, 2, r, 3), Y_IL_ADI(CStr(iller(p))), vbWhite, RGB(30, 41, 59), 10, True, xlLeft
            ws.Cells(r, 4).Value2 = ilIsgZK(p): ws.Cells(r, 5).Value2 = ilIsgZ(p): ws.Cells(r, 6).Value2 = ilIsgSK(p): ws.Cells(r, 7).Value2 = ilIsgS(p)
            ADM_ISG_DURUM ws, r, ilIsgZ(p), ilIsgS(p)
        Next p
    Else
        For Each t In isgSatir
            r = r + 1
            ADM_BLOK ws, ADM_ADRES(r, 2, r, 3), t(1) & "  " & t(2), vbWhite, RGB(30, 41, 59), 9, False, xlLeft
            ws.Cells(r, 4).Value2 = t(3): ws.Cells(r, 5).Value2 = t(4): ws.Cells(r, 6).Value2 = t(5): ws.Cells(r, 7).Value2 = t(6)
            ADM_ISG_DURUM ws, r, t(4), t(6)
        Next t
        If isgSatir.Count = 0 Then r = r + 1: ADM_BLOK ws, ADM_ADRES(r, 2, r, 8), "ÝSG verisi yok (bu il için ilçe stok güncellemesi gerekir).", vbWhite, RGB(100, 116, 139), 9, False, xlLeft
    End If
    If r > satirNo Then ADM_MIKTAR ws.Range(ADM_ADRES(satirNo + 1, 4, r, 7))
    ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 8)), True
    ' en cok zimmet bekleyen 5 depo
    n = satirNo
    ADM_BASLIKLAR ws, n, 10, Array("Depo", "", "Zimmet adet", "Kalem")
    ws.Range(ADM_ADRES(n, 10, n, 11)).Merge
    If isgSatir.Count > 0 Then
        ReDim sira(1 To isgSatir.Count)
        i = 0
        For Each t In isgSatir
            i = i + 1: sira(i) = t
        Next t
        For i = 1 To application.Min(5, isgSatir.Count)
            For j = i + 1 To isgSatir.Count
                If sira(j)(4) > sira(i)(4) Then t = sira(i): sira(i) = sira(j): sira(j) = t
            Next j
        Next i
        For i = 1 To application.Min(5, isgSatir.Count)
            If sira(i)(4) > 0 Then
                n = n + 1
                ADM_BLOK ws, ADM_ADRES(n, 10, n, 11), Y_IL_ADI(CStr(sira(i)(0))) & " • " & sira(i)(2), vbWhite, RGB(30, 41, 59), 9, False, xlLeft
                ws.Cells(n, 12).Value2 = sira(i)(4): ws.Cells(n, 13).Value2 = sira(i)(3)
                ws.Range(ADM_ADRES(n, 12, n, 13)).Interior.Color = vbWhite: ws.Range(ADM_ADRES(n, 12, n, 13)).HorizontalAlignment = xlCenter
                ws.Cells(n, 12).Font.Color = RGB(185, 28, 28): ws.Cells(n, 12).Font.bold = True
            End If
        Next i
    End If
    If n = satirNo Then n = n + 1: ADM_BLOK ws, ADM_ADRES(n, 10, n, 13), "Zimmet bekleyen yok.", RGB(220, 252, 231), RGB(21, 128, 61), 9, True, xlLeft
    ADM_KENAR ws.Range(ADM_ADRES(satirNo, 10, n, 13)), True
    If n > r Then r = n
    '--- en degerli 10 depo (tekil)
    r = r + 2
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "EN DEÐERLÝ 10 DEPO   •   " & ilMetin & "  (tekil)"
    r = r + 1: satirNo = r
    ADM_BASLIKLAR ws, r, 2, Array("Ýl", "Birim(ler)", "", "ÜY", "Depo", "Depo adý", "", "Stok satýrý", "Deðer", "", "Pay", "")
    ws.Range(ADM_ADRES(r, 3, r, 4)).Merge: ws.Range(ADM_ADRES(r, 7, r, 8)).Merge: ws.Range(ADM_ADRES(r, 10, r, 11)).Merge: ws.Range(ADM_ADRES(r, 12, r, 13)).Merge
    ReDim sira(1 To application.Max(1, tekilSec.Count))
    i = 0
    For Each k In tekilSec.keys
        i = i + 1: sira(i) = tekilSec(k)
    Next k
    For i = 1 To application.Min(10, tekilSec.Count)
        For j = i + 1 To tekilSec.Count
            If sira(j)(3) > sira(i)(3) Then t = sira(i): sira(i) = sira(j): sira(j) = t
        Next j
    Next i
    For i = 1 To application.Min(10, tekilSec.Count)
        r = r + 1
        parca = Split(sira(i)(6), "|")
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
        ws.Cells(r, 2).Value2 = Y_IL_ADI(CStr(sira(i)(5)))
        ws.Range(ADM_ADRES(r, 3, r, 4)).Merge: ws.Cells(r, 3).Value2 = sira(i)(4)
        ws.Range(ADM_ADRES(r, 5, r, 6)).NumberFormat = "@"
        ws.Cells(r, 5).Value2 = parca(0): ws.Cells(r, 6).Value2 = parca(1)
        ws.Range(ADM_ADRES(r, 7, r, 8)).Merge: ws.Cells(r, 7).Value2 = sira(i)(1): ws.Cells(r, 7).HorizontalAlignment = xlLeft
        ws.Cells(r, 9).Value2 = sira(i)(2): ws.Cells(r, 9).NumberFormat = "#,##0"
        ws.Range(ADM_ADRES(r, 10, r, 11)).Merge
        ws.Cells(r, 10).Value2 = sira(i)(3): ws.Cells(r, 10).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ADM_ADRES(r, 12, r, 13)).Merge
        ws.Cells(r, 12).Value2 = IIf(tekilToplam > 0, sira(i)(3) / tekilToplam, 0): ws.Cells(r, 12).NumberFormat = "0.0%"
        ws.Range(ADM_ADRES(r, 2, r, 6)).HorizontalAlignment = xlCenter
    Next i
    If r > satirNo Then
        ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 13)), True
        On Error Resume Next
        With ws.Range(ADM_ADRES(satirNo + 1, 12, r, 12)).FormatConditions.AddDatabar
            .BarColor.Color = RGB(147, 197, 253)
            .BarFillType = xlDataBarFillSolid
        End With
        On Error GoTo 0
    End If
    '--- en degerli 10 malzeme (tekil depolardan)
    ADM_ILERLE 0.98, "Malzeme tablolarý çiziliyor..."
    r = r + 2
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "EN DEÐERLÝ 10 MALZEME   •   " & ilMetin & "  (tekil depolar, tüm partiler)"
    r = r + 1: satirNo = r
    ADM_BASLIKLAR ws, r, 2, Array("Malzeme", "Malzeme tanýmý", "", "", "", "Miktar", "Ölçü", "Depo sayýsý", "Deðer", "", "Pay", "")
    ws.Range(ADM_ADRES(r, 3, r, 6)).Merge: ws.Range(ADM_ADRES(r, 10, r, 11)).Merge: ws.Range(ADM_ADRES(r, 12, r, 13)).Merge
    Set malz = CreateObject("Scripting.Dictionary"): malz.CompareMode = vbTextCompare
    For u = 0 To 3
        For p = 0 To 4
            il = CStr(iller(p))
            If Len(secIl) = 0 Or il = secIl Then
                veri = ADM_VERI(u, il)
                For i = 1 To veri(2)
                    If (i And 2047) = 0 Then AKSORA_BEKLE_NABIZ
                    parca = Split(veri(1)(i), vbTab)
                    kx = il & "|" & parca(0) & "|" & parca(1)
                    If tekil.Exists(kx) Then
                        If tekil(kx)(0) = u Then
                            If Not malz.Exists(parca(3)) Then malz(parca(3)) = Array(parca(4), 0#, parca(7), 0#, CreateObject("Scripting.Dictionary"))
                            t = malz(parca(3))
                            t(1) = t(1) + Val(parca(6)): t(3) = t(3) + Val(parca(8))
                            t(4)(kx) = True
                            malz(parca(3)) = t
                        End If
                    End If
                Next i
            End If
        Next p
    Next u
    ReDim sira(1 To application.Max(1, malz.Count))
    i = 0
    For Each k In malz.keys
        i = i + 1: sira(i) = Array(CStr(k), malz(k)(0), malz(k)(1), malz(k)(2), malz(k)(3), malz(k)(4).Count)
    Next k
    For i = 1 To application.Min(10, malz.Count)
        For j = i + 1 To malz.Count
            If sira(j)(4) > sira(i)(4) Then t = sira(i): sira(i) = sira(j): sira(j) = t
        Next j
    Next i
    For i = 1 To application.Min(10, malz.Count)
        r = r + 1
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = vbWhite
        ws.Cells(r, 2).NumberFormat = "@": ws.Cells(r, 2).Value2 = sira(i)(0)
        ws.Range(ADM_ADRES(r, 3, r, 6)).Merge: ws.Cells(r, 3).Value2 = sira(i)(1): ws.Cells(r, 3).HorizontalAlignment = xlLeft
        ws.Cells(r, 7).Value2 = sira(i)(2): ADM_MIKTAR ws.Cells(r, 7)
        ws.Cells(r, 8).Value2 = sira(i)(3)
        ws.Cells(r, 9).Value2 = sira(i)(5)
        ws.Range(ADM_ADRES(r, 10, r, 11)).Merge
        ws.Cells(r, 10).Value2 = sira(i)(4): ws.Cells(r, 10).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ADM_ADRES(r, 12, r, 13)).Merge
        ws.Cells(r, 12).Value2 = IIf(tekilToplam > 0, sira(i)(4) / tekilToplam, 0): ws.Cells(r, 12).NumberFormat = "0.0%"
        ws.Cells(r, 2).HorizontalAlignment = xlCenter
        ws.Range(ADM_ADRES(r, 8, r, 9)).HorizontalAlignment = xlCenter
    Next i
    If r > satirNo Then
        ADM_KENAR ws.Range(ADM_ADRES(satirNo, 2, r, 13)), True
        On Error Resume Next
        With ws.Range(ADM_ADRES(satirNo + 1, 12, r, 12)).FormatConditions.AddDatabar
            .BarColor.Color = RGB(134, 239, 172)
            .BarFillType = xlDataBarFillSolid
        End With
        On Error GoTo 0
    End If
    r = r + 2
    ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Bu sayfa yalnýzca yöneticiler içindir. Birim kartlarý birimin kendi görünümüdür; tekil deðerde birimler arasý ortak depolar bir kez sayýlýr. " & _
        "Demontaj kontrolü ilçe biriminin günlük kayýtlarýndan hesaplanýr. Tam güncel rapor için TÜM BÝRÝMLERÝ GÜNCELLE kullanýn.", RGB(244, 248, 252), RGB(100, 116, 139), 8, False, xlLeft
    ws.Range(ADM_ADRES(r, 2, r, 2)).Font.Italic = True
    parca = Split(ADM_AD_OKU("_AKSORA_ADMIN_LOG"), ";")
    If UBound(parca) >= 0 Then
        t = vbNullString
        For i = 0 To application.Min(4, UBound(parca))
            t = t & IIf(i > 0, "    •    ", "") & Replace(parca(i), "|", "  ")
        Next i
        r = r + 1
        ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Son yönetici giriþleri:  " & t, RGB(244, 248, 252), RGB(100, 116, 139), 8, False, xlLeft
    End If
    ws.Range("BX2").Value2 = r
    On Error Resume Next
    For Each s In ws.Shapes
        s.DrawingObject.PrintObject = (Left$(s.Name, 7) = "ADM_IL_")
    Next s
    On Error GoTo 0
    ws.ScrollArea = "A1:N" & (r + 2)
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    ' Onceki arama yeni secimle tekrar calistirilir.
    If Len(aramaMetni) > 0 Then AKSORA_ADMIN_ARA
    application.StatusBar = False
End Sub

' Yonetici giris kaydi (son 15): "tarih|kullanici|GIRIS/HATALI" ; ile ayrik, gizli ad _AKSORA_ADMIN_LOG.
Private Sub ADM_LOG_YAZ(ByVal sonuc As String)
    Dim eski As String, yeni As String, p() As String, i As Long, kullanici As String
    On Error Resume Next
    kullanici = Environ$("USERNAME")
    If Len(kullanici) = 0 Then kullanici = application.UserName
    eski = ADM_AD_OKU("_AKSORA_ADMIN_LOG")
    yeni = Format$(Now, "dd.mm.yyyy hh:nn") & "|" & Replace(Replace(kullanici, "|", " "), ";", " ") & "|" & sonuc
    If Len(eski) > 0 Then
        p = Split(eski, ";")
        For i = 0 To application.Min(13, UBound(p))
            If Len(yeni) + Len(p(i)) + 1 > 250 Then Exit For
            yeni = yeni & ";" & p(i)
        Next i
    End If
    ThisWorkbook.names.Add Name:="_AKSORA_ADMIN_LOG", RefersTo:="=""" & Replace(yeni, """", "'") & """", Visible:=False
    On Error GoTo 0
End Sub

' Gecmis anahtarindan (depo|UY|malzeme|degerleme) malzeme kodu.
Private Function ADM_KOD3(ByVal k As Variant) As String
    Dim p() As String
    p = Split(CStr(k), "|")
    If UBound(p) >= 2 Then ADM_KOD3 = Trim$(p(2))
End Function

' Miktar bicimi: binlik ".", kusurat yalniz varsa "," (tam sayida sonda virgul kalmaz).
Private Sub ADM_MIKTAR(ByVal alan As Range)
    Dim c As Range
    For Each c In alan.Columns
        AKSORA_V13_MIKTAR_BICIMI c
    Next c
End Sub

' Depo turu: 1 ana depo (D..), 2 ilce (K.., PB..), 3 yuklenici / diger
Private Function ADM_DEPO_TUR(ByVal depo As String) As Long
    If AKSORA_IL_DEPO_GECERLI(depo, "ILCE", "") Then ADM_DEPO_TUR = 2: Exit Function
    Select Case UCase$(Left$(Trim$(depo), 1))
        Case "D": ADM_DEPO_TUR = 1
        Case "K": ADM_DEPO_TUR = 2
        Case Else: ADM_DEPO_TUR = 3
    End Select
End Function

Private Sub ADM_ISG_DURUM(ByVal ws As Worksheet, ByVal r As Long, ByVal z As Double, ByVal s As Double)
    ws.Range(ADM_ADRES(r, 4, r, 7)).HorizontalAlignment = xlCenter
    ws.Range(ADM_ADRES(r, 2, r, 7)).Interior.Color = vbWhite
    If z > 0 Then
        ws.Cells(r, 5).Font.Color = RGB(185, 28, 28): ws.Cells(r, 5).Font.bold = True
        ADM_BLOK ws, ADM_ADRES(r, 8, r, 8), "ZÝMMET YAPILMALI", RGB(254, 226, 226), RGB(185, 28, 28), 8, True, xlCenter
    ElseIf s > 0 Then
        ADM_BLOK ws, ADM_ADRES(r, 8, r, 8), "SARF BEKLÝYOR", RGB(224, 242, 254), RGB(20, 103, 173), 8, True, xlCenter
    Else
        ADM_BLOK ws, ADM_ADRES(r, 8, r, 8), "TEMÝZ", RGB(220, 252, 231), RGB(21, 128, 61), 8, True, xlCenter
    End If
End Sub

'---------- arama (secili il veya tum iller, tum birimler; birden cok kelime: hepsi) ----------
' Ayni il+UY+depo+malzeme birden cok birimde varsa tek satir ("MALZEME + YATIRIM"), miktar bir kez.
' Siralama: ana depolar (D..), ilce depolari (K..), yukleniciler; her grupta miktar buyukten kucuge. Tablo filtrelenebilir.
Public Sub AKSORA_ADMIN_ARA()
    Dim ws As Worksheet, ara As String, b As Variant, iller As Variant, secIl As String, u As Long, p As Long, il As String, veri As Variant
    Dim i As Long, r As Long, n As Long, m As Long, parca() As String, seriEsles As Boolean, k As String, t As Variant, kx As Variant
    Dim grup As Object, depolar As Object, birimler As Object, olcu As Object, ilSet As Object, kk As Variant
    Dim satirlar As Long, toplamMk As Double, toplamTl As Double, ev As Boolean, kelime() As String, w As Long, uy As Boolean
    Dim alanMetin As String, ilMk As Object, dagilim As String, tur As Long, turMk(1 To 3) As Double, turSay(1 To 3) As Long
    Dim anah() As String, idx() As Long, cikti() As Variant, bas As Long, ilkS As Long, sonS As Long, satir1 As Long
    Dim sc As Boolean, ac As Boolean, desen() As String, araDesen As String
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(ADM)
    ara = Trim$(CStr(ws.Range(ADM_ARA_HUCRE).Value2))
    ev = application.EnableEvents: sc = application.ScreenUpdating
    application.EnableEvents = False: application.ScreenUpdating = False
    ' Rapor kurulurken (ADM_KUR) kart zaten acik; tek basina aramada kart acilir.
    If Len(ara) >= 2 And Not mAdmYukleme And Not mAdmMesgul Then ac = AKSORA_BEKLE_AC("AKSORA  |  YÖNETÝM MERKEZÝ", "Aranýyor: " & ara, 0.2)
    application.Cursor = xlWait
    ws.Unprotect ""
    ws.AutoFilterMode = False
    ws.Range("BX4").Value2 = ara
    With ws.Range(ADM_ADRES(ADM_ARA_SATIR, 2, ADM_ARA_SATIR + ADM_ARA_N + 1, 13))
        .UnMerge: .Clear: .Interior.Color = RGB(244, 248, 252)
        .Font.Name = "Aptos": .Font.Size = 10: .VerticalAlignment = xlCenter
    End With
    ws.Rows(ADM_ARA_SATIR & ":" & ADM_ARA_SATIR + ADM_ARA_N + 1).Hidden = True
    If Len(ara) < 2 Then GoTo Bitir
    Do While InStr(ara, "  ") > 0: ara = Replace(ara, "  ", " "): Loop
    kelime = Split(ara, " ")
    ' Turkce harfler (i/I/Ý/ý, þ/s ...) ve buyuk/kucuk harf ayni sayilir.
    ReDim desen(0 To UBound(kelime))
    For w = 0 To UBound(kelime): desen(w) = AKSORA_ARAMA_DESEN(CStr(kelime(w))): Next w
    araDesen = AKSORA_ARAMA_DESEN(ara)
    secIl = ADM_IL_KOD(CStr(ws.Range("BX3").Value2))
    Set grup = CreateObject("Scripting.Dictionary"): grup.CompareMode = vbTextCompare
    Set depolar = CreateObject("Scripting.Dictionary"): Set birimler = CreateObject("Scripting.Dictionary")
    Set olcu = CreateObject("Scripting.Dictionary"): Set ilSet = CreateObject("Scripting.Dictionary"): Set ilMk = CreateObject("Scripting.Dictionary")
    b = ADM_BIRIMLER(): iller = ADM_ILLER()
    For u = 0 To 3
        For p = 0 To 4
            il = CStr(iller(p))
            If Len(secIl) = 0 Or il = secIl Then
                veri = ADM_VERI(u, il)
                For i = 1 To veri(2)
                    If (i And 2047) = 0 Then AKSORA_BEKLE_NABIZ
                    If veri(1)(i) Like desen(0) Then
                        parca = Split(veri(1)(i), vbTab)
                        seriEsles = (UBound(kelime) = 0 And parca(9) Like araDesen)
                        alanMetin = parca(3) & " " & parca(4)
                        uy = True
                        For w = 0 To UBound(kelime)
                            If Not (alanMetin Like desen(w)) Then uy = False: Exit For
                        Next w
                        If Val(parca(6)) = 0 Then uy = False: seriEsles = False
                        If seriEsles Or uy Then
                            k = il & "|" & parca(0) & "|" & parca(1) & "|" & parca(3) & IIf(seriEsles, "|" & parca(9), "")
                            If Not grup.Exists(k) Then
                                tur = ADM_DEPO_TUR(parca(1))
                                grup(k) = Array(CStr(b(u)(4)), Y_IL_ADI(il), parca(1), parca(2), parca(3), parca(4), _
                                    IIf(seriEsles, parca(9), parca(5)), 0&, 0#, parca(7), 0#, u, tur)
                            End If
                            t = grup(k)
                            If t(11) = u Then
                                ' depo sahibi birim: miktar ve deger bir kez toplanir
                                t(7) = t(7) + 1: t(8) = t(8) + Val(parca(6)): t(10) = t(10) + Val(parca(8))
                                satirlar = satirlar + 1
                                olcu(parca(7)) = True: ilSet(il) = True
                                ilMk(Y_IL_ADI(il)) = ilMk(Y_IL_ADI(il)) + Val(parca(6))
                                toplamMk = toplamMk + Val(parca(6)): toplamTl = toplamTl + Val(parca(8))
                                turMk(t(12)) = turMk(t(12)) + Val(parca(6))
                                depolar(il & "|" & parca(0) & "|" & parca(1)) = t(12)
                            ElseIf InStr(t(0), CStr(b(u)(4))) = 0 Then
                                t(0) = t(0) & " + " & CStr(b(u)(4))
                            End If
                            birimler(CStr(b(u)(4))) = True
                            grup(k) = t
                        End If
                    End If
                Next i
            End If
        Next p
    Next u
    r = ADM_ARA_SATIR
    ws.Rows(r).Hidden = False
    If grup.Count = 0 Then
        ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "“" & ara & "” için " & IIf(Len(secIl) = 0, "beþ ilde", Y_IL_ADI(secIl) & " ilinde") & " hiçbir birimde kayýt bulunamadý.", _
            RGB(254, 243, 199), RGB(146, 64, 14), 10, True, xlLeft
        GoTo Bitir
    End If
    For Each kx In depolar.keys
        turSay(depolar(kx)) = turSay(depolar(kx)) + 1
    Next kx
    kk = olcu.keys
    If olcu.Count = 1 And ilSet.Count > 1 Then
        For Each kx In ilMk.keys
            dagilim = dagilim & "  " & kx & " " & ADM_ADET(ilMk(kx))
        Next kx
        dagilim = vbLf & "Ýl daðýlýmý:" & dagilim
    End If
    t = "“" & ara & "”:  " & ilSet.Count & " il  •  " & birimler.Count & " birim  •  " & depolar.Count & " depo  •  " & _
        AKSORA_SAYI_METNI(satirlar, "#,##0") & " satýr" & IIf(olcu.Count = 1, "  •  toplam " & ADM_ADET(toplamMk) & " " & CStr(kk(0)), "") & "  •  " & ADM_TL(toplamTl) & _
        vbLf & "ANA DEPO: " & turSay(1) & " depo" & IIf(olcu.Count = 1, " / " & ADM_ADET(turMk(1)), "") & _
        "     ÝLÇE: " & turSay(2) & " depo" & IIf(olcu.Count = 1, " / " & ADM_ADET(turMk(2)), "") & _
        "     YÜKLENÝCÝ: " & turSay(3) & " depo" & IIf(olcu.Count = 1, " / " & ADM_ADET(turMk(3)), "") & dagilim & _
        IIf(grup.Count > ADM_ARA_N - 2, vbLf & "(ilk " & ADM_ARA_N - 2 & " / " & grup.Count & " satýr gösteriliyor)", "")
    ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), CStr(t), RGB(220, 252, 231), RGB(21, 128, 61), 10, True, xlLeft
    ws.Cells(r, 2).WrapText = True: ws.Cells(r, 2).VerticalAlignment = xlTop
    ws.Rows(r).RowHeight = 16 * (2 + IIf(Len(dagilim) > 0, 1, 0) + IIf(grup.Count > ADM_ARA_N - 2, 1, 0)) + 6
    ' siralama: depo turu, sonra miktar (buyukten kucuge)
    n = grup.Count
    ReDim anah(1 To n): ReDim idx(1 To n)
    i = 0
    For Each kx In grup.keys
        i = i + 1: idx(i) = i
        t = grup(kx)
        anah(i) = CStr(t(12)) & Right$(String$(18, "0") & Format$(Int((1E+15 - t(8)) * 100), "0"), 18) & kx
    Next kx
    Y_SIRALA anah, idx, 1, n
    m = application.Min(n, ADM_ARA_N - 2)
    ReDim cikti(1 To m, 1 To 12)
    For i = 1 To m
        ' anahtarin sonundaki grup anahtari degil, sira numarasi kullanilir
        t = grup(Mid$(anah(i), 20))
        cikti(i, 1) = t(0): cikti(i, 2) = t(1): cikti(i, 3) = t(2): cikti(i, 4) = t(3)
        cikti(i, 5) = t(4): cikti(i, 6) = t(5): cikti(i, 7) = t(6): cikti(i, 8) = t(7)
        cikti(i, 9) = t(8): cikti(i, 10) = t(9): cikti(i, 11) = t(10)
        cikti(i, 12) = Choose(t(12), "ANA DEPO", "ÝLÇE", "YÜKLENÝCÝ")
    Next i
    r = r + 1: bas = r
    ws.Rows(r & ":" & r + m).Hidden = False
    ADM_BASLIKLAR ws, r, 2, Array("Birim", "Ýl", "Depo", "Depo adý", "Malzeme", "Malzeme tanýmý", "Seri no / Parti", "Satýr", "Miktar", "Ölçü", "Deðer", "Depo türü")
    satir1 = r + 1
    ws.Range(ADM_ADRES(satir1, 2, satir1 + m - 1, 8)).NumberFormat = "@"
    ws.Range(ADM_ADRES(satir1, 11, satir1 + m - 1, 11)).NumberFormat = "@"
    ws.Range(ADM_ADRES(satir1, 13, satir1 + m - 1, 13)).NumberFormat = "@"
    ws.Range(ADM_ADRES(satir1, 2, satir1 + m - 1, 13)).Value2 = cikti
    r = satir1 + m - 1
    With ws.Range(ADM_ADRES(satir1, 2, r, 13))
        .Font.Name = "Aptos": .Font.Size = 10: .Font.Color = RGB(30, 41, 59)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
    End With
    ws.Range(ADM_ADRES(satir1, 5, r, 5)).HorizontalAlignment = xlLeft
    ws.Range(ADM_ADRES(satir1, 7, r, 7)).HorizontalAlignment = xlLeft
    ws.Range(ADM_ADRES(satir1, 10, r, 10)).HorizontalAlignment = xlRight
    ws.Range(ADM_ADRES(satir1, 12, r, 12)).HorizontalAlignment = xlRight
    ws.Range(ADM_ADRES(satir1, 10, r, 10)).Font.bold = True
    ws.Range(ADM_ADRES(satir1, 9, r, 9)).NumberFormat = "#,##0"
    ADM_MIKTAR ws.Range(ADM_ADRES(satir1, 10, r, 10))
    ws.Range(ADM_ADRES(satir1, 12, r, 12)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ADM_ADRES(satir1, 2, r, 2)).ShrinkToFit = True
    ws.Range(ADM_ADRES(satir1, 5, r, 5)).ShrinkToFit = True
    ws.Range(ADM_ADRES(satir1, 12, r, 12)).ShrinkToFit = True
    ' depo turune gore renk (siralama tur bazinda oldugundan bloklar ardisik)
    ilkS = satir1
    For i = 1 To m
        If i = m Then
            sonS = satir1 + i - 1
        ElseIf cikti(i + 1, 12) <> cikti(i, 12) Then
            sonS = satir1 + i - 1
        Else
            sonS = 0
        End If
        If sonS > 0 Then
            Select Case cikti(i, 12)
                Case "ANA DEPO": ws.Range(ADM_ADRES(ilkS, 2, sonS, 13)).Interior.Color = RGB(219, 234, 254)
                Case "ÝLÇE": ws.Range(ADM_ADRES(ilkS, 2, sonS, 13)).Interior.Color = RGB(237, 233, 254)
                Case Else: ws.Range(ADM_ADRES(ilkS, 2, sonS, 13)).Interior.Color = RGB(254, 243, 199)
            End Select
            Select Case cikti(i, 12)
                Case "ANA DEPO": ws.Range(ADM_ADRES(ilkS, 13, sonS, 13)).Font.Color = RGB(29, 78, 216)
                Case "ÝLÇE": ws.Range(ADM_ADRES(ilkS, 13, sonS, 13)).Font.Color = RGB(91, 33, 182)
                Case Else: ws.Range(ADM_ADRES(ilkS, 13, sonS, 13)).Font.Color = RGB(146, 64, 14)
            End Select
            ws.Range(ADM_ADRES(ilkS, 13, sonS, 13)).Font.bold = True
            ilkS = sonS + 1
        End If
    Next i
    ADM_KENAR ws.Range(ADM_ADRES(bas, 2, r, 13)), True
    ' baslik satirindan filtre (her sutun)
    ws.Range(ADM_ADRES(bas, 2, r, 13)).AutoFilter
Bitir:
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
    If ac Then AKSORA_BEKLE_KAPAT
    ' Rapor kurulurken cagrildiysa ekran ve imlec ADM_AC'ye birakilir (yarim sayfa gorunmez).
    If Not mAdmMesgul Then application.StatusBar = False: application.Cursor = xlDefault
    application.EnableEvents = ev: application.ScreenUpdating = IIf(mAdmMesgul, sc, True)
    Exit Sub
Hata:
    t = Err.Description
    On Error Resume Next
    If ac Then AKSORA_BEKLE_KAPAT
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    If Not mAdmMesgul Then application.StatusBar = False: application.Cursor = xlDefault
    application.EnableEvents = ev: application.ScreenUpdating = IIf(mAdmMesgul, sc, True)
    If Not mAdmMesgul Then MsgBox "Arama yapýlamadý: " & t, vbExclamation, "AKSORA | ADMÝN"
End Sub

'---------- tek tus: tum birimler ----------
Public Sub AKSORA_ADMIN_TUMUNU_GUNCELLE()
    Dim bas As Double, sure As Double, sonuc As String, ws As Worksheet, sh As Worksheet
    If AKSORA_V11_MESGUL Then Exit Sub
    If MsgBox("Tüm birimler (Malzeme, Müþteri, Yatýrým, Ýlçe) beþ il için SAP'den sýrayla güncellenecek." & vbCrLf & vbCrLf & _
        "SAP oturumu açýk olmalýdýr; iþlem uzun sürebilir. Ýlerleme bu ekranda kartta izlenir (birim, il ve geçen süre); " & _
        "sonunda yönetim merkezi yenilenir." & vbCrLf & vbCrLf & "Devam edilsin mi?", vbYesNo + vbQuestion, "AKSORA | TÜM BÝRÝMLERÝ GÜNCELLE") <> vbYes Then Exit Sub
    ' SAP'ye giris yapilmamissa hic baslanmaz (birimler tek tek hata verip uzamasin).
    Dim sapMesaj As String, sapYok As String
    If SAP_SESSION_AL("CDP", sapMesaj) Is Nothing Then sapYok = sapYok & vbCrLf & "CDP: " & sapMesaj
    sapMesaj = vbNullString
    If SAP_SESSION_AL("EPP", sapMesaj) Is Nothing Then sapYok = sapYok & vbCrLf & "EPP: " & sapMesaj
    If Len(sapYok) > 0 Then
        MsgBox "SAP baðlantýsý bulunamadý; güncelleme baþlatýlmadý." & vbCrLf & sapYok & vbCrLf & vbCrLf & _
            "SAP'de CDP ve EPP sistemlerine giriþ yapýp tekrar deneyin.", vbExclamation, "AKSORA | TÜM BÝRÝMLERÝ GÜNCELLE"
        Exit Sub
    End If
    AKSORA_ADMIN_SONUC = vbNullString
    bas = Timer: mAdmTopluBas = bas: mAdmTopluSon = 0
    ' Kart birimler calismadan once acilir; ekran yonetim merkezinde sabit kalir (birim panellerine gidilmez).
    AKSORA_ADMIN_IPTAL = False
    mAdmTopluKart = AKSORA_BEKLE_AC("AKSORA  |  TÜM BÝRÝMLER GÜNCELLENÝYOR", "SAP baðlantýsý hazýrlanýyor...", 0.01, True)
    AKSORA_ADMIN_TOPLU = True
    ADM_EKRAN_KILIT True
    ' Calisirken Excel'e tiklama / yazma / baska dosyaya gecme engellenir (kart ve SAP pencereleri etkilenmez).
    application.Interactive = False
    ' Yatirim ve Ilce, Malzemenin ayni SAP okumasini kullanir (hazirsa); Musteri ayri rapor (IQ09) oldugundan sonda.
    ADM_BIRIM_CALISTIR 0, "MALZEME", "AKSORA_BUTUN_ILLERI_GUNCELLE"
    AKSORA_ORTAK_AKTIF = AKSORA_ORTAK_HAZIR()
    If AKSORA_ORTAK_AKTIF Then AKSORA_ADMIN_SONUC = AKSORA_ADMIN_SONUC & "(Yatýrým ve Ýlçe, Malzemenin SAP okumasýyla güncellendi; ek sorgu yapýlmadý.)" & vbCrLf
    ADM_BIRIM_CALISTIR 1, "YATIRIM", "AKSORA_YATIRIM_BUTUN_ILLER"
    ADM_BIRIM_CALISTIR 2, "ÝLÇE", "AKSORA_ILCE_BUTUN_ILLER"
    AKSORA_ORTAK_BIRAK
    ADM_BIRIM_CALISTIR 3, "MÜÞTERÝ", "AKSORA_MUSTERI_BUTUN_ILLERI_GUNCELLE"
    On Error Resume Next
    application.Interactive = True
    AKSORA_ARKA_PLAN_BITIR
    AKSORA_ORTAK_BIRAK
    AKSORA_ADMIN_TOPLU = False
    ADM_EKRAN_KILIT False
    If mAdmTopluKart Then AKSORA_BEKLE_KAPAT
    mAdmTopluKart = False
    sure = Timer - bas: If sure < 0 Then sure = sure + 86400#
    sonuc = Format$(Now, "dd.mm.yyyy hh:nn") & "  •  süre " & AKSORA_SURE_METNI(sure)
    ThisWorkbook.names.Add Name:="_AKSORA_ADMIN_SON_TOPLU", RefersTo:="=""" & Replace(sonuc, """", """""") & """", Visible:=False
    ' Birimlerin actigi paneller gizlenir; yalniz yonetim merkezi gorunur.
    application.ScreenUpdating = False
    Set ws = ThisWorkbook.Worksheets(ADM)
    ws.Visible = xlSheetVisible
    ws.Select Replace:=True
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name <> ws.Name And sh.Visible = xlSheetVisible Then sh.Visible = xlSheetVeryHidden
    Next sh
    On Error GoTo 0
    ADM_AC
    If AKSORA_ADMIN_IPTAL Then
        AKSORA_ADMIN_IPTAL = False
        MsgBox "Güncelleme durduruldu (" & sonuc & ")." & vbCrLf & "Çalýþan il tamamlandý, kalanlar atlandý." & vbCrLf & vbCrLf & AKSORA_ADMIN_SONUC, vbExclamation, "AKSORA | TÜM BÝRÝMLER"
    Else
        MsgBox "Tüm birimler güncellendi (" & sonuc & ")." & vbCrLf & vbCrLf & AKSORA_ADMIN_SONUC, vbInformation, "AKSORA | TÜM BÝRÝMLER"
    End If
End Sub

Private Sub ADM_BIRIM_CALISTIR(ByVal sira As Long, ByVal ad As String, ByVal makro As String)
    On Error Resume Next
    If AKSORA_ADMIN_IPTAL Then AKSORA_ADMIN_SONUC = AKSORA_ADMIN_SONUC & ad & ": durduruldu (çalýþtýrýlmadý)." & vbCrLf: Exit Sub
    mAdmTopluSira = sira: mAdmTopluAd = ad
    Set mAdmTopluIl = CreateObject("Scripting.Dictionary")
    AKSORA_ADMIN_DURUM "güncelleme baþlýyor..."
    Err.Clear
    application.Run "'" & Replace(ThisWorkbook.Name, "'", "''") & "'!" & makro
    If Err.Number <> 0 Then AKSORA_ADMIN_SONUC = AKSORA_ADMIN_SONUC & ad & ": HATA - " & Err.Description & vbCrLf
    Err.Clear
    ' Birim calisirken ekran kilidi baska pencereye gecmis olabilir; yeniden uygulanir.
    ADM_EKRAN_KILIT True
    AKSORA_ADMIN_DURUM "tamamlandý."
    DoEvents
    On Error GoTo 0
End Sub

' Tum birimler guncellenirken kart: birim, il ve son islem; cubuk birim (4) ve il (5) ilerledikce dolar.
' Birimlerin durum yazan yordamlari (CORE_DURUM, Y_DURUM, I_DURUM, musteri / malzeme toplu durum) burayi cagirir.
Public Sub AKSORA_ADMIN_DURUM(ByVal metin As String)
    Dim sure As Double, u As String, k As Variant, y As Double, s As String, sn As Long
    If Not mAdmTopluKart Or Not mAdmYukleme Then Exit Sub
    If Len(metin) = 0 And Abs(Timer - mAdmTopluSon) < 1 Then Exit Sub
    If Len(metin) > 0 And metin = mAdmTopluMetin And Abs(Timer - mAdmTopluSon) < 0.5 Then Exit Sub
    On Error Resume Next
    If Len(metin) > 0 Then
        mAdmTopluMetin = metin
        u = Y_NORM(metin)
        If Not mAdmTopluIl Is Nothing Then
            For Each k In Array("ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
                If InStr(u, CStr(k)) > 0 Then mAdmTopluIl(CStr(k)) = True
            Next k
        End If
    End If
    mAdmTopluSon = Timer
    sure = Timer - mAdmTopluBas: If sure < 0 Then sure = sure + 86400#
    sn = CLng(Fix(sure))
    If mAdmTopluIl Is Nothing Then y = 0 Else y = mAdmTopluIl.Count / 5
    If metin = "tamamlandý." Then y = 1
    y = (mAdmTopluSira + 0.05 + 0.9 * y) / 4
    s = mAdmTopluAd & "  •  " & mAdmTopluMetin
    If AKSORA_ADMIN_IPTAL Then s = "Durduruluyor: çalýþan il bitince iþlem duracak..."
    If Len(s) > 72 Then s = Left$(s, 69) & "..."
    frmSplash.lblDurum.Caption = s & vbLf & "Birim " & (mAdmTopluSira + 1) & " / 4   •   Geçen süre: " & (sn \ 60) & " dk " & Format$(sn Mod 60, "00") & " sn"
    frmSplash.lblBar.Width = application.Max(1, frmSplash.lblBarArka.Width * application.Min(1, y))
    frmSplash.Repaint
    On Error GoTo 0
End Sub

' Excel penceresi cizimi durdurulur (birim sayfalari arkada degisir, ekranda yonetim merkezi kalir); kart ayri pencere oldugundan cizilir.
Private Sub ADM_EKRAN_KILIT(ByVal kilit As Boolean)
    On Error Resume Next
    ' Pencere kilidi kullanilmaz (ekran griye donup SAP/Excel uyarilarini gizleyebiliyordu);
    ' ekranin yonetim merkezinde kalmasini birimlerdeki AKSORA_ADMIN_TOPLU korumalari saglar.
    If kilit Then
        ' (kilit acilmaz; yalniz eski surumden kalan kilit kaldirilir)
    ElseIf mAdmKilit Then
        ADMW_LockWindowUpdate 0
        mAdmKilit = False
    End If
    On Error GoTo 0
End Sub

'---------- yardimcilar ----------
Private Function ADM_AD_OKU(ByVal ad As String) As String
    Dim v As Variant
    On Error Resume Next
    v = application.Evaluate(ThisWorkbook.names(ad).RefersTo)
    If Not IsError(v) Then ADM_AD_OKU = CStr(v)
    On Error GoTo 0
End Function

' Atil stok ozeti (ATIL STOK sayfasiyla ayni hesap). SAP okumasi yoksa bolum cizilmez. Donus: son satir.
Private Function ADM_ATIL_CIZ(ByVal ws As Worksheet, ByVal r As Long, ByVal secIl As String) As Long
    Dim c As Collection, bilgi As Variant, x As Variant, iller As Variant, p As Long, j As Long, ilk As Long, n As Long
    Dim v(0 To 4, 0 To 8) As Double, kapsam(0 To 4) As Double, ok As String, t As Double, enAz As Double, enCok As Double, t2 As Long
    ADM_ATIL_CIZ = r
    If ADM_WS(ATIL_KAYIT) Is Nothing Then Exit Function
    On Error GoTo Hata
    ADM_ILERLE 0.96, "Atýl stok özeti hesaplanýyor..."
    Set c = ATIL_HESAPLA(bilgi)
    If bilgi(0) <= 0 And bilgi(1) <= 0 Then Exit Function
    iller = ADM_ILLER()
    For Each x In c
        For p = 0 To 4
            If x(0) = iller(p) Then Exit For
        Next p
        If p <= 4 Then
            kapsam(p) = kapsam(p) + x(8)
            If x(10) >= 1 Then
                ' v: 0/1 atil kalem/deger, 2/3 3 ay, 4/5 6 ay, 6 transfer onerisi, 7 yuksek oncelik, 8 stogu dusen
                For t2 = 1 To 3
                    If x(20)(t2) > 0.0000001 Then j = (3 - t2) * 2: v(p, j) = v(p, j) + 1: v(p, j + 1) = v(p, j + 1) + x(19)(t2)
                Next t2
                If Len(x(13)) > 0 Then
                    v(p, 6) = v(p, 6) + 1
                    If x(18) = "YÜKSEK" Then v(p, 7) = v(p, 7) + 1
                End If
                If x(16) Then v(p, 8) = v(p, 8) + 1
            End If
        End If
    Next x
    ok = "SAP okumasý: CDP " & IIf(bilgi(0) > 0, Format$(bilgi(0), "dd.mm.yyyy"), "-") & "  •  EPP " & IIf(bilgi(1) > 0, Format$(bilgi(1), "dd.mm.yyyy"), "-")
    r = r + 2
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "ATIL STOK   (24 aydan eski stok • D, K ve yüklenici depolarý • " & ok & ")"
    r = r + 1: ilk = r
    ADM_BASLIKLAR ws, r, 2, Array("Ýl", "ATIL kalem", "ATIL deðeri", "3 AY kalem", "3 AY deðeri", "6 AY kalem", "6 AY deðeri", _
        "Atýl oraný", "Transfer önerisi", "Yüksek öncelik", "Stoðu düþen", "Deðerlendirme")
    enAz = -1: enCok = -1
    For p = 0 To 4
        If kapsam(p) > 0 Then t = v(p, 1) / kapsam(p) Else t = 0
        If enAz < 0 Or t < enAz Then enAz = t
        If t > enCok Then enCok = t
    Next p
    For p = 0 To 4
        r = r + 1
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = IIf(CStr(iller(p)) = secIl, RGB(226, 236, 247), vbWhite)
        ws.Cells(r, 2).Value2 = Y_IL_ADI(CStr(iller(p))): ws.Cells(r, 2).Font.bold = True
        For j = 0 To 5
            ws.Cells(r, 3 + j).Value2 = IIf(j Mod 2 = 1, Round(v(p, j), 2), v(p, j))
        Next j
        If kapsam(p) > 0 Then t = v(p, 1) / kapsam(p) Else t = 0
        ws.Cells(r, 9).Value2 = t
        ws.Cells(r, 10).Value2 = v(p, 6): ws.Cells(r, 11).Value2 = v(p, 7): ws.Cells(r, 12).Value2 = v(p, 8)
        If enCok > enAz + 0.0005 And t = enAz Then
            ws.Cells(r, 13).Value2 = ChrW(9733) & " EN ÝYÝ"
        ElseIf enCok > enAz + 0.0005 And t = enCok Then
            ws.Cells(r, 13).Value2 = ChrW(9660) & " EN KÖTÜ"
        Else
            ws.Cells(r, 13).Value2 = IIf(t < 0.05, "ÝYÝ", IIf(t < 0.15, "ORTA", "ZAYIF"))
        End If
    Next p
    r = r + 1
    ws.Cells(r, 2).Value2 = "TOPLAM"
    For j = 3 To 12
        If j <> 9 Then ws.Cells(r, j).Formula = "=SUM(" & ADM_ADRES(ilk + 1, j, r - 1, j) & ")"
    Next j
    t = 0
    For p = 0 To 4: t = t + kapsam(p): Next p
    If t > 0 Then ws.Cells(r, 9).Value2 = (v(0, 1) + v(1, 1) + v(2, 1) + v(3, 1) + v(4, 1)) / t Else ws.Cells(r, 9).Value2 = 0
    ws.Range(ADM_ADRES(r, 2, r, 13)).Font.bold = True
    ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = RGB(226, 236, 247)
    ws.Range(ADM_ADRES(ilk + 1, 2, r, 13)).HorizontalAlignment = xlCenter
    ws.Range(ADM_ADRES(ilk + 1, 4, r, 4)).NumberFormat = "#,##0 ""TL"""
    ws.Range(ADM_ADRES(ilk + 1, 6, r, 6)).NumberFormat = "#,##0 ""TL"""
    ws.Range(ADM_ADRES(ilk + 1, 8, r, 8)).NumberFormat = "#,##0 ""TL"""
    ws.Range(ADM_ADRES(ilk + 1, 9, r, 9)).NumberFormat = "0.0%"
    ws.Range(ADM_ADRES(ilk + 1, 3, r, 4)).Font.Color = RGB(185, 28, 28)
    ws.Range(ADM_ADRES(ilk + 1, 4, r, 8)).ShrinkToFit = True
    I_DURUM_RENK ws.Range(ADM_ADRES(ilk + 1, 13, r - 1, 13))
    ADM_KENAR ws.Range(ADM_ADRES(ilk, 2, r, 13)), True
    ADM_ATIL_CIZ = r
    Exit Function
Hata:
    On Error Resume Next
    r = r + 1
    ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Atýl stok özeti hesaplanamadý: " & Err.Description, RGB(254, 243, 199), RGB(146, 64, 14), 9, False, xlLeft
    ADM_ATIL_CIZ = r
End Function

Private Function ADM_ADRES(ByVal r1 As Long, ByVal c1 As Long, ByVal r2 As Long, ByVal c2 As Long) As String
    ADM_ADRES = Chr$(64 + c1) & r1 & ":" & Chr$(64 + c2) & r2
End Function

Private Sub ADM_BLOK(ByVal ws As Worksheet, ByVal adres As String, ByVal metin As String, ByVal dolgu As Long, _
    ByVal yazi As Long, ByVal boyut As Single, ByVal kalin As Boolean, ByVal hiz As Long)
    With ws.Range(adres)
        If .Cells.Count > 1 Then .Merge
        .Interior.Color = dolgu
        .Font.Color = yazi: .Font.Size = boyut: .Font.bold = kalin: .Font.Name = "Aptos"
        .HorizontalAlignment = hiz: .VerticalAlignment = xlCenter
        .NumberFormat = "@"
        .Cells(1, 1).Value2 = metin
        If hiz = xlLeft Then .IndentLevel = 1
    End With
End Sub

Private Sub ADM_BOLUM(ByVal ws As Worksheet, ByVal r As Long, ByVal adres As String, ByVal metin As String)
    ws.Rows(r).RowHeight = 26
    ADM_BLOK ws, adres, metin, RGB(244, 248, 252), RGB(12, 45, 79), 12, True, xlLeft
    ws.Range(adres).IndentLevel = 0
    With ws.Range(adres).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous: .Weight = xlMedium: .Color = RGB(34, 185, 191)
    End With
End Sub

Private Sub ADM_BASLIKLAR(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long, ByVal basliklar As Variant)
    Dim i As Long
    For i = 0 To UBound(basliklar)
        ws.Cells(r, c + i).Value2 = basliklar(i)
    Next i
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + UBound(basliklar)))
        .Interior.Color = RGB(26, 75, 120): .Font.Color = vbWhite: .Font.bold = True: .Font.Size = 9
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    ws.Rows(r).RowHeight = 24
End Sub

Private Sub ADM_KENAR(ByVal alan As Range, Optional ByVal ic As Boolean = False)
    With alan
        .Borders(xlEdgeLeft).LineStyle = xlContinuous: .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeTop).LineStyle = xlContinuous: .Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).Color = RGB(203, 213, 225): .Borders(xlEdgeRight).Color = RGB(203, 213, 225)
        .Borders(xlEdgeTop).Color = RGB(203, 213, 225): .Borders(xlEdgeBottom).Color = RGB(203, 213, 225)
        If ic And .Rows.Count > 1 Then
            .Borders(xlInsideHorizontal).LineStyle = xlContinuous
            .Borders(xlInsideHorizontal).Color = RGB(226, 232, 240)
        End If
    End With
End Sub

Private Sub ADM_DUGME(ByVal ws As Worksheet, ByVal ad As String, ByVal metin As String, ByVal x As Double, _
    ByVal renk As Long, ByVal makro As String, ByVal genislik As Double, Optional ByVal y As Double = -1, Optional ByVal h As Double = -1)
    Dim s As Shape, w As Double
    w = ws.Columns("B").Width * genislik
    If y < 0 Then y = ws.Rows(1).Top + 12
    If h < 0 Then h = 28
    Set s = ws.Shapes.AddShape(msoShapeRoundedRectangle, x, y, w, h)
    With s
        .Name = ad
        .Fill.ForeColor.RGB = renk: .line.ForeColor.RGB = RGB(116, 192, 226): .line.Weight = 1
        .Adjustments.Item(1) = 0.18
        .Placement = xlFreeFloating
        .OnAction = BIRIM_MAKRO(makro)
        With .TextFrame2
            .MarginLeft = 4: .MarginRight = 4: .MarginTop = 0: .MarginBottom = 0
            .VerticalAnchor = msoAnchorMiddle: .WordWrap = msoTrue
            .TextRange.Text = metin
            .TextRange.Font.Name = "Aptos": .TextRange.Font.Size = 9: .TextRange.Font.bold = msoTrue
            .TextRange.Font.Fill.ForeColor.RGB = vbWhite
            .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        End With
    End With
End Sub
'==========================================================
' ATIL STOK (Malzeme Stok Birimi)
' Atil (uretim yeri bazinda): eldeki stok FIFO ile girislere dagitilir (en yeni giris once kalir);
' 24 aydan eski giristen kalan kisim ATIL, 21-24 ay "3 AY KALDI", 18-21 ay "6 AY KALDI".
' Giris: 101/105/501/561 ve uretim yerleri arasi transferin alan tarafi (iptalleriyle). Ayni uretim yerindeki
' depo transferi (311) yasi sifirlamaz; yuklenici sarfi stogu azaltir, kalan eski stok atil kalir.
' Kapsam: illerin CDP ve EPP D001-D999, K001-K999 ve yuklenici depolari (uretim yeri toplami);
'         HURDA ve CIKMA/BDSZ partiler haric.
' SAP (MB51, /ATIL duzeni) arada bir okunur; sayac SAP'siz isler. Sonraki stok verilerinde artis yeni giris katmani olur.
' _ATIL (ATIL2): 1. satir etiket / okuma tarihleri (D1 CDP, E1 EPP), 2. satir basliklar, 3. satirdan:
'   sistem | UY | malzeme | girisler (yyyymmdd=miktar;...) | 24-60 ay onceki son giris | son 6 ay sarf | son sarf |
'   SAP okumasindaki stok | son karsilastirilan stok | son karsilastirilan stok verisinin tarihi
'==========================================================
Private Function ATIL_IL_BUYUK(ByVal il As String) As String
    Select Case il
        Case "ARTVIN": ATIL_IL_BUYUK = "ARTVÝN"
        Case "GUMUSHANE": ATIL_IL_BUYUK = "GÜMÜÞHANE"
        Case "GIRESUN": ATIL_IL_BUYUK = "GÝRESUN"
        Case "RIZE": ATIL_IL_BUYUK = "RÝZE"
        Case Else: ATIL_IL_BUYUK = il
    End Select
End Function

Private Function ATIL_SISTEM(ByVal uy As String) As String
    Select Case Left$(uy, 3)
        Case "715": ATIL_SISTEM = "CDP"
        Case "705": ATIL_SISTEM = "EPP"
    End Select
End Function

Private Function ATIL_TESIS_IL(ByVal uy As String) As String
    Dim il As Variant
    For Each il In ADM_ILLER()
        If AKSORA_IL_TESIS(CStr(il), "CDP") = uy Or AKSORA_IL_TESIS(CStr(il), "EPP") = uy Then ATIL_TESIS_IL = CStr(il): Exit Function
    Next il
End Function

' D001-D999 ve K001-K999
Private Function ATIL_DEPO_MI(ByVal depo As String) As Boolean
    depo = UCase$(Trim$(depo))
    If Len(depo) <> 4 Then Exit Function
    If Left$(depo, 1) <> "D" And Left$(depo, 1) <> "K" Then Exit Function
    If Not (Mid$(depo, 2) Like "###") Then Exit Function
    ATIL_DEPO_MI = (Mid$(depo, 2) <> "000")
End Function

Private Function ATIL_PARTI_DISI(ByVal parti As String) As Boolean
    Dim t As String
    t = Y_NORM(parti)
    ATIL_PARTI_DISI = (InStr(t, "HURDA") > 0 Or (InStr(t, "CIKMA") > 0 And (InStr(t, "BDSZ") > 0 Or InStr(t, "BEDELSIZ") > 0)))
End Function

Private Function ATIL_KATEGORI(ByVal mg As String, ByVal tn As String) As String
    mg = UCase$(Trim$(mg)): tn = Y_NORM(tn)
    If Left$(tn, 6) = "IPTAL." Then tn = Trim$(Mid$(tn, 7))
    Select Case True
        Case Left$(mg, 5) = "M0316": ATIL_KATEGORI = "KABLO"
        Case Left$(mg, 5) = "M0308": ATIL_KATEGORI = "DÝREK"
        Case Left$(mg, 5) = "M0327": ATIL_KATEGORI = "PANO"
        Case Left$(mg, 5) = "M0314": ATIL_KATEGORI = "HÜCRE"
        Case Left$(mg, 5) = "M0340": ATIL_KATEGORI = "TRAFO"
        Case Len(mg) > 0: ATIL_KATEGORI = "DÝÐER"
        Case Left$(tn, 6) = "KABLO ": ATIL_KATEGORI = "KABLO"
        Case Left$(tn, 6) = "DIREK ": ATIL_KATEGORI = "DÝREK"
        Case Left$(tn, 5) = "PANO ": ATIL_KATEGORI = "PANO"
        Case Left$(tn, 6) = "HUCRE ": ATIL_KATEGORI = "HÜCRE"
        Case Left$(tn, 6) = "TRAFO ": ATIL_KATEGORI = "TRAFO"
        Case Else: ATIL_KATEGORI = "DÝÐER"
    End Select
End Function

Private Function ATIL_GECEN(ByVal bas As Double) As Double
    ATIL_GECEN = Timer - bas
    If ATIL_GECEN < 0 Then ATIL_GECEN = ATIL_GECEN + 86400#
End Function

Private Function ATIL_HUCRE_SAYI(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then ATIL_HUCRE_SAYI = CDbl(v)
End Function

'---------- ATIL STOK: stok verisi ----------
' Ilin kapsamdaki stogu (D/K depolari ve yuklenici depolari; malzeme birimi stok sayfasi ya da il kaydi):
'   UY|malzeme -> Array(0 miktar, 1 deger, 2 birim, 3 tanim, 4 kategori, 5 depolar: depo -> Array(ad, miktar, deger))
' tarih: stok verisinin tarihi (0: veri yok)
Private Function ATIL_STOK_OKU(ByVal il As String, ByRef tarih As Date) As Object
    Dim d As Object, gorulen As Object, kaynaklar As Variant, kx As Variant, sh As Worksheet, canli As Boolean, t As Variant
    Dim v As Variant, son As Long, i As Long, uy As String, depo As String, mal As String, k As String, x As Double, dg As Double
    Dim o As Variant, dp As Object, ka As String, sira As Long, p As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set gorulen = CreateObject("Scripting.Dictionary"): gorulen.CompareMode = vbTextCompare
    tarih = 0
    canli = (il = ADM_CANLI_IL(0))
    If canli Then
        kaynaklar = Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
    Else
        kaynaklar = Array("_AKLIVE_" & il & "_CDP", "_AKLIVE_" & il & "_EPP", "_AKLIVE_" & il & "_ILC", "_AKLIVE_" & il & "_YUK")
    End If
    For Each kx In kaynaklar
        sira = sira + 1
        Set sh = ADM_WS(CStr(kx))
        If Not sh Is Nothing Then
            If canli Then t = sh.Range("C2").Value Else t = sh.Range("A3").Value
            If Not IsError(t) Then
                If IsDate(t) Then If CDate(t) > tarih Then tarih = CDate(t)
            End If
            son = sh.Cells(sh.Rows.Count, 7).End(xlUp).Row
            v = Empty
            If son >= 7 Then v = sh.Range(sh.Cells(7, 1), sh.Cells(son, 14)).Value2
            If IsArray(v) Then
                For i = 1 To UBound(v, 1)
                    If (i And 2047) = 0 Then AKSORA_BEKLE_NABIZ
                    mal = Y_KOD_NORM(v(i, 7))
                    depo = UCase$(I_TEMIZ(v(i, 2)))
                    uy = I_TEMIZ(v(i, 1))
                    If Len(mal) > 0 And (ATIL_DEPO_MI(depo) Or (sira = 4 And Len(depo) > 0)) And Len(ATIL_SISTEM(uy)) > 0 Then
                        If ATIL_TESIS_IL(uy) = il And Not ATIL_PARTI_DISI(I_TEMIZ(v(i, 6))) Then
                            If sira = 4 And Not mAtilYuk Is Nothing Then
                                If Not mAtilYuk.Exists(uy) Then mAtilYuk.Add uy, CreateObject("Scripting.Dictionary")
                                mAtilYuk(uy)(depo) = True
                            End If
                            x = I_SAYI(v(i, 9))
                            ' Ayni satir iki sayfada varsa (eski kayit duzeni) bir kez sayilir.
                            ka = uy & "|" & depo & "|" & mal & "|" & UCase$(I_TEMIZ(v(i, 6)))
                            If gorulen.Exists(ka) Then
                                If gorulen(ka) <> sira Then x = 0
                            Else
                                gorulen.Add ka, sira
                            End If
                            If x > 0 Then
                                dg = I_SAYI(v(i, 11))
                                k = uy & "|" & mal
                                If Not d.Exists(k) Then
                                    Set dp = CreateObject("Scripting.Dictionary"): dp.CompareMode = vbTextCompare
                                    d.Add k, Array(0#, 0#, I_TEMIZ(v(i, 10)), I_TEMIZ(v(i, 8)), ATIL_KATEGORI(I_TEMIZ(v(i, 4)), I_TEMIZ(v(i, 8))), dp)
                                End If
                                o = d(k): o(0) = o(0) + x: o(1) = o(1) + dg: d(k) = o
                                Set dp = o(5)
                                If Not dp.Exists(depo) Then dp.Add depo, Array(I_TEMIZ(v(i, 3)), 0#, 0#)
                                p = dp(depo): p(1) = p(1) + x: p(2) = p(2) + dg: dp(depo) = p
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next kx
    Set ATIL_STOK_OKU = d
End Function

' il -> Array(stok sozlugu, stok tarihi)
Private Function ATIL_TUM_STOK() As Object
    Dim d As Object, il As Variant, t As Date, s As Object
    Set d = CreateObject("Scripting.Dictionary")
    Set mAtilYuk = CreateObject("Scripting.Dictionary"): mAtilYuk.CompareMode = vbTextCompare
    For Each il In ADM_ILLER()
        AKSORA_BEKLE_NABIZ
        Set s = ATIL_STOK_OKU(CStr(il), t)
        d.Add CStr(il), Array(s, t)
    Next il
    Set ATIL_TUM_STOK = d
End Function

'---------- ATIL STOK: kayit ----------
' UY|malzeme -> Array(0 sistem, 1 girisler, 2 eski giris, 3 son 6 ay sarf, 4 son sarf, 5 SAP stogu, 6 son stok, 7 son stok tarihi)
' okuma: Array(CDP okuma tarihi, EPP okuma tarihi) (0: okunmadi). Eski kural kaydi (ATIL1) okunmaz.
Private Function ATIL_KAYIT_OKU(ByRef okuma As Variant) As Object
    Dim d As Object, ws As Worksheet, son As Long, v As Variant, r As Long, k As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    okuma = Array(0#, 0#)
    mAtilEski = False
    Set ATIL_KAYIT_OKU = d
    Set ws = ADM_WS(ATIL_KAYIT)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Range("A1").Value2) = "ATIL1" Then mAtilEski = True
    If CStr(ws.Range("A1").Value2) <> "ATIL2" Then Exit Function
    okuma = Array(ATIL_HUCRE_SAYI(ws.Range("D1").Value2), ATIL_HUCRE_SAYI(ws.Range("E1").Value2))
    son = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    If son < 3 Then Exit Function
    v = ws.Range("A3:J" & son).Value2
    For r = 1 To UBound(v, 1)
        k = CStr(v(r, 2)) & "|" & CStr(v(r, 3))
        If Len(CStr(v(r, 2))) > 0 And Len(CStr(v(r, 3))) > 0 Then
            d(k) = Array(CStr(v(r, 1)), CStr(v(r, 4)), ATIL_HUCRE_SAYI(v(r, 5)), ATIL_HUCRE_SAYI(v(r, 6)), ATIL_HUCRE_SAYI(v(r, 7)), _
                ATIL_HUCRE_SAYI(v(r, 8)), ATIL_HUCRE_SAYI(v(r, 9)), ATIL_HUCRE_SAYI(v(r, 10)))
        End If
    Next r
End Function

' Yeni stok verisi gelen illerde: stok onceki miktardan fazlaysa fark yeni giris sayilir (stok tarihiyle katman eklenir);
' azalis FIFO ile en eski stoktan duser. SAP okumasindan sonra stoga giren yeni malzemeler de eklenir. Donus: azalan kalem.
Private Function ATIL_KARSILASTIR(ByVal tum As Object) As Long
    Dim ws As Worksheet, son As Long, v As Variant, r As Long, il As Variant, uy As String, k As Variant, s As Object
    Dim cq As Double, td As Double, degisti As Boolean, n As Long, var As Object, yeni As Collection, a() As Variant, c As Long
    Dim ok0 As Double, ok1 As Double, sis As String, x As Variant
    Set ws = ADM_WS(ATIL_KAYIT)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Range("A1").Value2) <> "ATIL2" Then Exit Function
    ok0 = ATIL_HUCRE_SAYI(ws.Range("D1").Value2): ok1 = ATIL_HUCRE_SAYI(ws.Range("E1").Value2)
    son = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    Set var = CreateObject("Scripting.Dictionary"): var.CompareMode = vbTextCompare
    If son >= 3 Then
        v = ws.Range("A3:J" & son).Value2
        For r = 1 To UBound(v, 1)
            uy = CStr(v(r, 2))
            var(uy & "|" & CStr(v(r, 3))) = True
            il = ATIL_TESIS_IL(uy)
            If Len(il) > 0 And Len(CStr(v(r, 3))) > 0 Then
                If tum.Exists(CStr(il)) Then
                    td = CDbl(tum(CStr(il))(1))
                    If td > 0 And td > ATIL_HUCRE_SAYI(v(r, 10)) + 0.00001 Then
                        Set s = tum(CStr(il))(0)
                        cq = 0
                        If s.Exists(uy & "|" & CStr(v(r, 3))) Then cq = s(uy & "|" & CStr(v(r, 3)))(0)
                        If cq > ATIL_HUCRE_SAYI(v(r, 9)) + 0.0005 Then
                            v(r, 4) = ATIL_KATMAN_EKLE(CStr(v(r, 4)), td, cq - ATIL_HUCRE_SAYI(v(r, 9)))
                        ElseIf cq < ATIL_HUCRE_SAYI(v(r, 9)) - 0.0005 Then
                            n = n + 1
                        End If
                        v(r, 9) = cq: v(r, 10) = td
                        degisti = True
                    End If
                End If
            End If
        Next r
    End If
    ' SAP okumasindan sonra stoga giren malzemeler: tamami yeni giris.
    Set yeni = New Collection
    For Each il In ADM_ILLER()
        If tum.Exists(CStr(il)) Then
            td = CDbl(tum(CStr(il))(1))
            Set s = tum(CStr(il))(0)
            For Each k In s.keys
                uy = Left$(CStr(k), InStr(CStr(k), "|") - 1)
                sis = ATIL_SISTEM(uy)
                If Not var.Exists(CStr(k)) And td > 0 And ((sis = "CDP" And ok0 > 0) Or (sis = "EPP" And ok1 > 0)) Then
                    yeni.Add Array(sis, uy, Mid$(CStr(k), InStr(CStr(k), "|") + 1), ATIL_KATMAN_EKLE(vbNullString, td, s(k)(0)), 0#, 0#, 0#, 0#, s(k)(0), td)
                End If
            Next k
        End If
    Next il
    If degisti Then ws.Range("A3:J" & son).Value2 = v
    If yeni.Count > 0 Then
        ReDim a(1 To yeni.Count, 1 To 10)
        r = 0
        For Each x In yeni
            r = r + 1
            For c = 1 To 10: a(r, c) = x(c - 1): Next c
        Next x
        ws.Range("A" & application.Max(3, son + 1)).Resize(yeni.Count, 10).Value2 = a
        ws.Range("C1").Value2 = application.Max(3, son + 1) + yeni.Count - 3
    End If
    If degisti Or yeni.Count > 0 Then AKSORA_SAP_TESHIS_YAZ "Atýl stok karþýlaþtýrma: " & n & " kalemde stok düþtü, " & yeni.Count & " yeni kalem"
    ATIL_KARSILASTIR = n
End Function

' Malzeme stoklari guncellenince cagrilir (kayit yoksa hicbir sey yapmaz).
Public Sub AKSORA_ATIL_STOK_KONTROL()
    On Error Resume Next
    If ADM_WS(ATIL_KAYIT) Is Nothing Then Exit Sub
    ATIL_KARSILASTIR ATIL_TUM_STOK()
    On Error GoTo 0
End Sub

Private Sub ATIL_KAYDET(ByVal sonuc As Object, ByVal tum As Object, ByVal oturumlar As Object)
    Dim ws As Worksheet, son As Long, eski As Variant, r As Long, c As Long, n As Long, a() As Variant, il As Variant, sis As Variant
    Dim uy As String, s As Object, k As Variant, o As Variant, td As Double, yazildi As Object, kap As Long, ok As Variant
    Set ws = Y_KAYIT_SAYFA(ATIL_KAYIT)
    ok = Array(0#, 0#)
    If CStr(ws.Range("A1").Value2) = "ATIL2" Then ok = Array(ATIL_HUCRE_SAYI(ws.Range("D1").Value2), ATIL_HUCRE_SAYI(ws.Range("E1").Value2))
    son = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    kap = application.Max(son, 3) + sonuc.Count
    For Each il In ADM_ILLER()
        kap = kap + tum(CStr(il))(0).Count
    Next il
    ReDim a(1 To kap, 1 To 10)
    ' Bu okumada sorgulanmayan sistemin satirlari korunur (yalniz yeni kural kaydi).
    If son >= 3 And CStr(ws.Range("A1").Value2) = "ATIL2" Then
        eski = ws.Range("A3:J" & son).Value2
        For r = 1 To UBound(eski, 1)
            If Len(CStr(eski(r, 1))) > 0 And Not oturumlar.Exists(CStr(eski(r, 1))) Then
                n = n + 1
                For c = 1 To 10: a(n, c) = eski(r, c): Next c
            End If
        Next r
    End If
    Set yazildi = CreateObject("Scripting.Dictionary"): yazildi.CompareMode = vbTextCompare
    For Each sis In oturumlar.keys
        For Each il In ADM_ILLER()
            uy = AKSORA_IL_TESIS(CStr(il), CStr(sis))
            Set s = tum(CStr(il))(0): td = CDbl(tum(CStr(il))(1))
            For Each k In s.keys
                If Left$(CStr(k), Len(uy) + 1) = uy & "|" Then
                    n = n + 1
                    a(n, 1) = CStr(sis): a(n, 2) = uy: a(n, 3) = Mid$(CStr(k), Len(uy) + 2)
                    If sonuc.Exists(CStr(k)) Then
                        o = sonuc(CStr(k))
                        a(n, 4) = CStr(o(0)): a(n, 5) = CDbl(o(1)): a(n, 6) = CDbl(o(2)): a(n, 7) = CDbl(o(3))
                    Else
                        a(n, 4) = vbNullString: a(n, 5) = 0#: a(n, 6) = 0#: a(n, 7) = 0#
                    End If
                    a(n, 8) = s(k)(0): a(n, 9) = s(k)(0): a(n, 10) = td
                    yazildi(CStr(k)) = True
                End If
            Next k
            ' Stogu olmayan ama son 6 ayda sarf eden uretim yeri (transfer onerisi icin talep).
            For Each k In sonuc.keys
                If Left$(CStr(k), Len(uy) + 1) = uy & "|" And Not yazildi.Exists(CStr(k)) Then
                    o = sonuc(CStr(k))
                    n = n + 1
                    a(n, 1) = CStr(sis): a(n, 2) = uy: a(n, 3) = Mid$(CStr(k), Len(uy) + 2)
                    a(n, 4) = CStr(o(0)): a(n, 5) = CDbl(o(1)): a(n, 6) = CDbl(o(2)): a(n, 7) = CDbl(o(3))
                    a(n, 8) = 0#: a(n, 9) = 0#: a(n, 10) = td
                    yazildi(CStr(k)) = True
                End If
            Next k
        Next il
        If CStr(sis) = "CDP" Then ok(0) = CDbl(Now) Else ok(1) = CDbl(Now)
    Next sis
    ws.Cells.Clear
    ws.Range("A:D").NumberFormat = "@"
    ws.Range("A1").Value2 = "ATIL2": ws.Range("B1").Value2 = CDbl(Now): ws.Range("C1").Value2 = n
    ws.Range("D1").NumberFormat = "General": ws.Range("D1").Value2 = ok(0): ws.Range("E1").Value2 = ok(1)
    ws.Range("A2:J2").Value2 = Array("Sistem", "ÜY", "Malzeme", "Giriþler (24 ay)", "Eski giriþ", "Son 6 ay sarf", "Son sarf", _
        "SAP stoðu", "Son stok", "Son stok tarihi")
    If n > 0 Then ws.Range("A3").Resize(n, 10).Value2 = a
End Sub

'---------- ATIL STOK: SAP (MB51) ----------
Private Sub ATIL_PANO_YAZ(ByVal metin As String)
    Dim h As LongPtr, p As LongPtr, boyut As LongPtr, i As Long, acik As Boolean
    application.CutCopyMode = False
    For i = 1 To 40
        If ATIL_OpenClipboard(0) <> 0 Then acik = True: Exit For
        ATIL_Sleep 50
    Next i
    If Not acik Then Err.Raise vbObjectError + 2330, , "Pano açýlamadý (baþka bir program kullanýyor)."
    ATIL_EmptyClipboard
    boyut = (Len(metin) + 1) * 2
    h = ATIL_GlobalAlloc(&H2, boyut)
    If h = 0 Then ATIL_CloseClipboard: Err.Raise vbObjectError + 2331, , "Pano için bellek ayrýlamadý."
    p = ATIL_GlobalLock(h)
    If p = 0 Then ATIL_GlobalFree h: ATIL_CloseClipboard: Err.Raise vbObjectError + 2332, , "Pano belleði kilitlenemedi."
    ATIL_CopyMemory p, StrPtr(metin), boyut
    ATIL_GlobalUnlock h
    If ATIL_SetClipboardData(13, h) = 0 Then
        ATIL_GlobalFree h: ATIL_CloseClipboard
        Err.Raise vbObjectError + 2333, , "Malzeme listesi panoya yazýlamadý."
    End If
    ATIL_CloseClipboard
End Sub

Private Sub ATIL_SURE_YAZ(Optional ByVal zorla As Boolean = False)
    Dim sn As Long
    If Not mAtilKart Then Exit Sub
    If Not zorla And Abs(Timer - mAtilSon) < 1 Then Exit Sub
    mAtilSon = Timer
    sn = CLng(Fix(ATIL_GECEN(mAtilBas)))
    On Error Resume Next
    If AKSORA_ADMIN_IPTAL Then mAtilMetin = "Durduruluyor: çalýþan SAP sorgusu bitince duracak..."
    frmSplash.lblDurum.Caption = mAtilMetin & vbLf & "Geçen süre: " & (sn \ 60) & " dk " & Format$(sn Mod 60, "00") & " sn"
    frmSplash.Repaint
    On Error GoTo 0
End Sub

Private Sub ATIL_ADIM(ByVal y As Double, ByVal metin As String)
    application.StatusBar = "AKSORA | Atýl stok | " & metin
    If Len(metin) > 72 Then metin = Left$(metin, 69) & "..."
    If AKSORA_ADMIN_IPTAL Then metin = "Durduruluyor: çalýþan SAP sorgusu bitince duracak..."
    mAtilMetin = metin
    AKSORA_BEKLE_ILERLE y, metin
    ATIL_SURE_YAZ True
End Sub

Private Sub ATIL_BEKLE(ByVal ses As Object)
    Dim b As Double
    b = Timer
    Do While ses.Busy
        If ATIL_GECEN(b) > 1800 Then Err.Raise vbObjectError + 2334, , "SAP 30 dakika içinde yanýt vermedi."
        ATIL_SURE_YAZ
        DoEvents: ATIL_Sleep 50
    Loop
End Sub

Private Function ATIL_SBAR(ByVal ses As Object) As String
    On Error Resume Next
    ATIL_SBAR = ses.FindById("wnd[0]/sbar").Text & " [" & ses.FindById("wnd[0]/sbar").MessageType & "]"
    On Error GoTo 0
End Function

Private Sub ATIL_METIN_TOPLA(ByVal kok As Object, ByRef metin As String, ByVal derinlik As Long)
    Dim i As Long, c As Object, t As String, adet As Long
    If derinlik > 5 Then Exit Sub
    On Error Resume Next
    adet = kok.Children.Count
    For i = 0 To adet - 1
        Set c = Nothing
        Set c = kok.Children(i)
        If Not c Is Nothing Then
            t = vbNullString: t = Trim$(c.Text)
            If Len(t) > 0 And c.Type <> "GuiButton" Then metin = metin & " " & t
            If c.ContainerType Then ATIL_METIN_TOPLA c, metin, derinlik + 1
        End If
    Next i
    On Error GoTo 0
End Sub

' "Kayit yok" bildirimi mi? (yetki ve diger hatalar bos sonuc sayilmaz)
Private Function ATIL_BOS_MESAJ(ByVal t As String) As Boolean
    t = " " & Y_NORM(t) & " "
    If InStr(t, "YETKI") > 0 Or InStr(t, "AUTHORI") > 0 Then Exit Function
    If InStr(t, "BELGE") = 0 And InStr(t, "DOCUMENT") = 0 And InStr(t, "KAYIT") = 0 And InStr(t, "VERI") = 0 And InStr(t, "DATA") = 0 Then Exit Function
    ATIL_BOS_MESAJ = (InStr(t, " YOK") > 0 Or InStr(t, "BULUNAMADI") > 0 Or InStr(t, "BULUNMADI") > 0 Or InStr(t, "NOT FOUND") > 0 Or _
        InStr(t, " NO ") > 0 Or InStr(t, "MEVCUT DEGIL") > 0)
End Function

Private Sub ATIL_COKLU_SIL(ByVal ses As Object)
    Dim sil As Object, tip As String
    Set sil = ses.FindById("wnd[1]/tbar[0]/btn[16]", False)
    If sil Is Nothing Then Err.Raise vbObjectError + 2335, , "SAP çoklu seçim penceresi açýlmadý."
    tip = Y_NORM(sil.Tooltip)
    If InStr(tip, "DELETE") = 0 And InStr(tip, "SIL") = 0 Then _
        Err.Raise vbObjectError + 2336, , "SAP çoklu seçim temizleme düðmesi doðrulanamadý: " & sil.Tooltip
    sil.Press
    ATIL_BEKLE ses
End Sub

Private Function ATIL_TARIH(ByVal t As String) As Date
    Dim g As Long, a As Long, y As Long
    On Error GoTo Yok
    t = Trim$(t)
    If t Like "##.##.####" Or t Like "##/##/####" Or t Like "##-##-####" Then
        g = CLng(Left$(t, 2)): a = CLng(Mid$(t, 4, 2)): y = CLng(Mid$(t, 7, 4))
    ElseIf t Like "####-##-##" Or t Like "####.##.##" Or t Like "####/##/##" Then
        y = CLng(Left$(t, 4)): a = CLng(Mid$(t, 6, 2)): g = CLng(Mid$(t, 9, 2))
    ElseIf t Like "########" Then
        y = CLng(Left$(t, 4)): a = CLng(Mid$(t, 5, 2)): g = CLng(Right$(t, 2))
    Else
        Exit Function
    End If
    If y < 1990 Or a < 1 Or a > 12 Or g < 1 Or g > 31 Then Exit Function
    ATIL_TARIH = DateSerial(y, a, g)
    If Month(ATIL_TARIH) <> a Or ATIL_TARIH > Date + 1 Then ATIL_TARIH = 0
Yok:
End Function

Private Function ATIL_SAYI_DENE(ByVal v As Variant, ByRef x As Double) As Boolean
    Dim t As String, eksi As Boolean
    x = 0
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString And IsNumeric(v) Then x = CDbl(v): ATIL_SAYI_DENE = True: Exit Function
    t = Replace(Replace(Trim$(CStr(v)), ChrW(160), ""), " ", "")
    If Len(t) = 0 Then Exit Function
    If Right$(t, 1) = "-" Then eksi = True: t = Left$(t, Len(t) - 1)
    If Left$(t, 1) = "-" Then eksi = True: t = Mid$(t, 2)
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[0-9]{1,3}([.][0-9]{3})+$"
    If InStr(t, ",") = 0 Then
        If re.Test(t) Then t = Replace(t, ".", "")
    End If
    If InStr(t, ",") > 0 Then
        If InStrRev(t, ".") > InStrRev(t, ",") Then
            t = Replace(t, ",", "")
        Else
            t = Replace(Replace(t, ".", ""), ",", ".")
        End If
    End If
    If Len(t) = 0 Or t Like "*[!0-9.]*" Or (Len(t) - Len(Replace(t, ".", ""))) > 1 Then Exit Function
    x = Val(t)
    If eksi Then x = -x
    ATIL_SAYI_DENE = True
End Function

' Atil kapsami: D001-D999, K001-K999 ve uretim yerinin yuklenici depolari (yuklenici stok sayfasindan; M### / BO## kaliplari).
Private Function ATIL_KAPSAM_MI(ByVal uy As String, ByVal depo As String) As Boolean
    depo = UCase$(Trim$(depo))
    If Len(depo) = 0 Then Exit Function
    If ATIL_DEPO_MI(depo) Then ATIL_KAPSAM_MI = True: Exit Function
    If depo Like "M###" Or depo Like "BO##" Then ATIL_KAPSAM_MI = True: Exit Function
    If Not mAtilYuk Is Nothing Then
        If mAtilYuk.Exists(uy) Then ATIL_KAPSAM_MI = mAtilYuk(uy).Exists(depo)
    End If
End Function

' MB51 hareket turleri: girisler (iptalleriyle) ve sarflar (iptalleriyle).
Private Function ATIL_GIRIS_TUR() As String()
    ATIL_GIRIS_TUR = Split("101,102,105,106,501,502,561,562,301,302,303,304,305,306,315,316", ",")
End Function

Private Function ATIL_SARF_TUR() As String()
    ATIL_SARF_TUR = Split("201,202,221,222,241,242,261,262,281,282,291,292", ",")
End Function

' Satirin uretim yerine giris katkisi: 101/105/501/561 ve iptalleri isaretiyle; uretim yerleri arasi
' transferde yalniz alan taraf (301/303/305/315 arti, iptalleri 302/304/306/316 eksi).
Private Function ATIL_GIRIS_KATKI(ByVal itu As String, ByVal q As Double) As Double
    Select Case Trim$(itu)
        Case "101", "102", "105", "106", "501", "502", "561", "562": ATIL_GIRIS_KATKI = q
        Case "301", "303", "305", "315": If q > 0 Then ATIL_GIRIS_KATKI = q
        Case "302", "304", "306", "316": If q < 0 Then ATIL_GIRIS_KATKI = q
    End Select
End Function

Private Function ATIL_DS(ByVal d As Object, ByVal k As String) As Double
    If d.Exists(k) Then ATIL_DS = ATIL_HUCRE_SAYI(d(k))
End Function

' Gun gun net girisler "yyyymmdd=miktar;..." (iptal gunu kendinden onceki girislerden duser).
Private Function ATIL_KATMAN_METIN(ByVal katman As Object, ByVal mal As String) As String
    Dim gun As Object, k As Variant, n As Long, a() As String, q() As Double, i As Long, j As Long, ts As String, tq As Double, d As Double, t As String
    If Not katman.Exists(mal) Then Exit Function
    Set gun = katman(mal)
    n = gun.Count
    If n = 0 Then Exit Function
    ReDim a(1 To n): ReDim q(1 To n)
    For Each k In gun.keys
        i = i + 1: a(i) = CStr(k): q(i) = CDbl(gun(k))
    Next k
    For i = 2 To n
        ts = a(i): tq = q(i): j = i - 1
        Do While j >= 1
            If a(j) <= ts Then Exit Do
            a(j + 1) = a(j): q(j + 1) = q(j): j = j - 1
        Loop
        a(j + 1) = ts: q(j + 1) = tq
    Next i
    For i = 1 To n
        If q(i) < 0 Then
            d = -q(i): q(i) = 0
            For j = i - 1 To 1 Step -1
                If d <= 0 Then Exit For
                If q(j) > 0 Then
                    If q(j) >= d Then
                        q(j) = q(j) - d: d = 0
                    Else
                        d = d - q(j): q(j) = 0
                    End If
                End If
            Next j
        End If
    Next i
    For i = 1 To n
        If q(i) > 0.0000001 Then t = t & IIf(Len(t) > 0, ";", "") & a(i) & "=" & Trim$(Str$(Round(q(i), 6)))
    Next i
    ATIL_KATMAN_METIN = t
End Function

' Eldeki stok FIFO ile girislere dagitilir (en yeni giris once kalir). 24 aydan eski giristen kalan (ya da
' girisi bulunmayan) kisim ATIL; 21-24 ay 3 AY KALDI; 18-21 ay 6 AY KALDI. q(1) 6 ay, q(2) 3 ay, q(3) atil miktari.
Private Function ATIL_YAS(ByVal stok As Double, ByVal girisler As String, ByVal eski As Double, ByRef q() As Double, _
    ByRef gunMetin As String, ByRef girisMetin As String) As Long
    Dim p As Variant, i As Long, j As Long, n As Long, t() As Double, m() As Double, tt As Double, mm As Double, k As Long
    Dim kalan As Double, al As Double, s24 As Double, s21 As Double, s18 As Double, enYeniAtil As Double, enEskiRisk As Double
    ReDim q(0 To 3)
    gunMetin = vbNullString: girisMetin = vbNullString
    s24 = CDbl(DateAdd("m", -24, Date)): s21 = CDbl(DateAdd("m", -21, Date)): s18 = CDbl(DateAdd("m", -18, Date))
    If Len(girisler) > 0 Then
        p = Split(girisler, ";")
        ReDim t(0 To UBound(p)): ReDim m(0 To UBound(p))
        For i = 0 To UBound(p)
            k = InStr(p(i), "=")
            If k = 9 Then
                t(n) = CDbl(DateSerial(Val(Left$(p(i), 4)), Val(Mid$(p(i), 5, 2)), Val(Mid$(p(i), 7, 2))))
                m(n) = Val(Mid$(p(i), k + 1))
                n = n + 1
            End If
        Next i
    End If
    ' en yeni giris once
    For i = 1 To n - 1
        tt = t(i): mm = m(i): j = i - 1
        Do While j >= 0
            If t(j) >= tt Then Exit Do
            t(j + 1) = t(j): m(j + 1) = m(j): j = j - 1
        Loop
        t(j + 1) = tt: m(j + 1) = mm
    Next i
    kalan = stok
    For i = 0 To n - 1
        If kalan <= 0.0000001 Then Exit For
        If m(i) > 0 Then
            al = m(i)
            If al > kalan Then al = kalan
            kalan = kalan - al
            If t(i) < s24 Then
                q(3) = q(3) + al
                If t(i) > enYeniAtil Then enYeniAtil = t(i)
            ElseIf t(i) < s21 Then
                q(2) = q(2) + al
                If enEskiRisk = 0 Or t(i) < enEskiRisk Then enEskiRisk = t(i)
            ElseIf t(i) < s18 Then
                q(1) = q(1) + al
                If enEskiRisk = 0 Or t(i) < enEskiRisk Then enEskiRisk = t(i)
            End If
        End If
    Next i
    If kalan > 0.0000001 Then
        q(3) = q(3) + kalan
        If eski > enYeniAtil Then enYeniAtil = eski
    End If
    If q(3) > 0.0000001 Then
        ATIL_YAS = 3
        If enYeniAtil > 0 Then
            gunMetin = Format$(application.Max(0, Date - DateAdd("m", 24, CDate(Int(enYeniAtil)))), "#,##0") & " gündür atýl"
            girisMetin = Format$(CDate(Int(enYeniAtil)), "dd.mm.yyyy")
        Else
            gunMetin = "5+ yýldýr giriþ yok": girisMetin = "5 yýldan eski"
        End If
    ElseIf q(2) > 0.0000001 Or q(1) > 0.0000001 Then
        If q(2) > 0.0000001 Then ATIL_YAS = 2 Else ATIL_YAS = 1
        gunMetin = Format$(DateAdd("m", 24, CDate(Int(enEskiRisk))) - Date, "#,##0") & " gün kaldý"
        girisMetin = Format$(CDate(Int(enEskiRisk)), "dd.mm.yyyy")
    End If
End Function

Private Function ATIL_KATMAN_EKLE(ByVal girisler As String, ByVal tarih As Double, ByVal miktar As Double) As String
    ATIL_KATMAN_EKLE = girisler & IIf(Len(girisler) > 0, ";", "") & Format$(CDate(Int(tarih)), "yyyymmdd") & "=" & Trim$(Str$(Round(miktar, 6)))
End Function

' Coklu secim penceresi: temizlenir, degerler panodan yuklenir, ilk deger kontrol edilir.
Private Sub ATIL_COKLU_PANO(ByVal ses As Object, ByVal alan As String, ByRef degerler() As String)
    Dim b As Object, t As String, ilk As String
    ATIL_PANO_YAZ Join(degerler, vbCrLf)
    ses.FindById("wnd[0]/usr/btn%_" & alan & "_%_APP_%-VALU_PUSH").Press: ATIL_BEKLE ses
    ATIL_COKLU_SIL ses
    On Error Resume Next
    ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    On Error GoTo 0
    Set b = ses.FindById("wnd[1]/tbar[0]/btn[24]", False)
    If b Is Nothing Then Err.Raise vbObjectError + 2343, , "SAP 'Panodan yükle' düðmesi bulunamadý."
    t = Y_NORM(b.Tooltip)
    If InStr(t, "PANO") = 0 And InStr(t, "CLIPBOARD") = 0 Then Err.Raise vbObjectError + 2344, , "SAP 'Panodan yükle' düðmesi doðrulanamadý: " & b.Tooltip
    b.Press: ATIL_BEKLE ses
    On Error Resume Next
    ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    ilk = ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text
    On Error GoTo 0
    If Y_KOD_NORM(ilk) <> Y_KOD_NORM(degerler(LBound(degerler))) Then _
        Err.Raise vbObjectError + 2345, , alan & " listesi SAP'ye aktarýlamadý (ilk deðer '" & ilk & "', beklenen '" & degerler(LBound(degerler)) & "')."
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press: ATIL_BEKLE ses
End Sub

' MB51 listesi satirlari: Array(malzeme, kayit tarihi, miktar (isaretli), hareket turu, depo, degerleme turu).
' Ayni degerli bos hucreler (ALV siralama gosterimi) bir onceki satirdan tamamlanir; ara toplam ve baska uretim yeri satirlari alinmaz.
Private Sub ATIL_MB51_COZ(ByRef v As Variant, ByVal uy As String, ByVal sat As Collection, ByVal etiket As String)
    Dim r As Long, c As Long, h As Long, cM As Long, cT As Long, cQ As Long, cU As Long, cI As Long, cD As Long, cV As Long
    Dim a As String, t As String, x As Double, son As Long, n As Long, dolu As Long, enIyi As Long
    Dim ham As String, mal As String, tar As Date, uyx As String, itu As String, depo As String, dgr As String, q As Double
    Dim sMal As String, sTar As Date, sUy As String, sItu As String, sDepo As String, sDgr As String, atla As Long, okunan As Long
    For r = 1 To application.Min(UBound(v, 1), 40)
        For c = 1 To UBound(v, 2)
            If Y_NORM(CStr(v(r, c))) = "MALZEME" Or Y_NORM(CStr(v(r, c))) = "MATERIAL" Then h = r: cM = c: Exit For
        Next c
        If h > 0 Then Exit For
    Next r
    If h = 0 Then Err.Raise vbObjectError + 2340, , "MB51 listesinde 'Malzeme' sütunu bulunamadý (/ATIL düzenini kontrol edin)."
    a = vbNullString
    For c = 1 To UBound(v, 2): a = a & " | " & CStr(v(h, c)): Next c
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | baþlýk (satýr " & h & ")" & a
    For c = 1 To UBound(v, 2)
        a = Y_NORM(CStr(v(h, c)))
        If cT = 0 And (InStr(a, "KAYIT") > 0 Or InStr(a, "KYT") > 0 Or InStr(a, "POSTING") > 0) Then cT = c
        If cQ = 0 And (a = "MIKTAR" Or a = "QUANTITY") Then cQ = c
        If cU = 0 And (InStr(a, "URETIM") > 0 Or a = "PLANT" Or a = "UY") Then cU = c
        If cI = 0 And (a = "ITU" Or a = "HTU" Or a = "BWART" Or InStr(a, "HAREKET") > 0 Or InStr(a, "HRK") > 0 Or InStr(a, "MOVEMENT") > 0 Or a = "MVT") Then cI = c
        If cD = 0 And (a = "DPYR" Or InStr(a, "DEPO") > 0 Or a = "SLOC" Or InStr(a, "STOR") > 0) Then cD = c
        If cV = 0 And (InStr(a, "DGRL") > 0 Or InStr(a, "DEGERLEME") > 0 Or InStr(a, "VALUATION") > 0) Then cV = c
    Next c
    If cQ = 0 Then
        For c = 1 To UBound(v, 2)
            a = Y_NORM(CStr(v(h, c)))
            If c <> cM And c <> cT And (InStr(a, "MIK") > 0 Or InStr(a, "MKT") > 0 Or InStr(a, "QUANT") > 0) Then cQ = c: Exit For
        Next c
    End If
    son = application.Min(UBound(v, 1), h + 300)
    If cT = 0 Then
        enIyi = 0
        For c = 1 To UBound(v, 2)
            If c <> cM Then
                n = 0: dolu = 0
                For r = h + 1 To son
                    t = Trim$(CStr(v(r, c)))
                    If Len(t) > 0 Then
                        dolu = dolu + 1
                        If ATIL_TARIH(t) > 0 Then n = n + 1
                    End If
                Next r
                If n > enIyi And n * 10 >= dolu * 9 Then enIyi = n: cT = c
            End If
        Next c
    End If
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | sütunlar: malzeme " & cM & ", tarih " & cT & ", miktar " & cQ & ", ÜY " & cU & ", ÝTü " & cI & ", depo " & cD & ", dðrl " & cV
    If cT = 0 Or cQ = 0 Or cI = 0 Or cD = 0 Then Err.Raise vbObjectError + 2341, , _
        "MB51 /ATIL düzeninde kayýt tarihi, miktar, hareket türü (ÝTü) ve depo yeri sütunlarý gerekli. Baþlýk:" & Left$(a, 200)
    For r = h + 1 To UBound(v, 1)
        If (r And 4095) = 0 Then AKSORA_BEKLE_NABIZ: ATIL_SURE_YAZ
        ham = Trim$(CStr(v(r, cM)))
        If Y_NORM(ham) = "MALZEME" Or Y_NORM(ham) = "MATERIAL" Then GoTo Sonraki
        t = Trim$(CStr(v(r, cT)))
        If Left$(ham, 1) = "*" Or Left$(t, 1) = "*" Then atla = atla + 1: GoTo Sonraki
        If Len(ham) = 0 And Len(t) = 0 And Len(Trim$(CStr(v(r, cQ)))) = 0 Then GoTo Sonraki
        mal = Y_KOD_NORM(ham)
        If Len(mal) = 0 Then mal = sMal
        If Len(t) = 0 Then tar = sTar Else tar = ATIL_TARIH(t)
        uyx = vbNullString
        If cU > 0 Then uyx = Trim$(CStr(v(r, cU)))
        If Len(uyx) = 0 Then uyx = sUy
        itu = Trim$(CStr(v(r, cI)))
        If Len(itu) = 0 Then itu = sItu
        depo = UCase$(Trim$(CStr(v(r, cD))))
        If Len(depo) = 0 Then depo = sDepo
        dgr = vbNullString
        If cV > 0 Then dgr = Trim$(CStr(v(r, cV)))
        If Len(dgr) = 0 And Len(ham) = 0 Then dgr = sDgr
        If Len(mal) = 0 Or tar = 0 Or Len(itu) = 0 Then atla = atla + 1: GoTo Sonraki
        sMal = mal: sTar = tar: sUy = uyx: sItu = itu: sDepo = depo: sDgr = dgr
        If Len(uyx) > 0 And uyx <> uy Then atla = atla + 1: GoTo Sonraki
        q = 0
        If Not ATIL_SAYI_DENE(v(r, cQ), q) Then atla = atla + 1: GoTo Sonraki
        sat.Add Array(mal, CDbl(tar), q, itu, depo, dgr)
        okunan = okunan + 1
Sonraki:
    Next r
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | okunan " & okunan & " satýr, atlanan " & atla
End Sub

' MB51 (/ATIL duzeni): uretim yeri, hareket turleri ve malzemeler panodan, tarih araligi; depo filtresi yok
' (satirlar okunurken atil kapsamina gore suzulur). Donus: satirlar (ATIL_MB51_COZ).
Private Function ATIL_MB51(ByVal ses As Object, ByVal uy As String, ByRef mallar() As String, ByVal bas As Date, ByVal son As Date, _
    ByVal etiket As String, ByRef turler() As String) As Collection
    Dim sat As Collection, alv As Object, v As Variant, n As Long, pop As Object, t As String, b0 As Double, varsayilan As Boolean
    Set sat = New Collection
    Set ATIL_MB51 = sat
    n = UBound(mallar) - LBound(mallar) + 1
    If n <= 0 Then Exit Function
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | " & n & " malzeme, " & Format$(bas, "dd.mm.yyyy") & " - " & Format$(son, "dd.mm.yyyy") & " | ÝTü " & Join(turler, ",")
    ses.StartTransaction "MB51": ATIL_BEKLE ses
    If ses.FindById("wnd[0]/usr/ctxtWERKS-LOW", False) Is Nothing Then _
        Err.Raise vbObjectError + 2342, , "MB51 seçim ekraný açýlmadý: " & ATIL_SBAR(ses)
    ATIL_COKLU_PANO ses, "MATNR", mallar
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = uy
    On Error Resume Next
    ses.FindById("wnd[0]/usr/ctxtWERKS-HIGH").Text = vbNullString
    On Error GoTo 0
    ' Depo: tum depolar (onceki secim temizlenir)
    ses.FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press: ATIL_BEKLE ses
    ATIL_COKLU_SIL ses
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press: ATIL_BEKLE ses
    On Error Resume Next
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString
    On Error GoTo 0
    ATIL_COKLU_PANO ses, "BWART", turler
    ses.FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = Format$(bas, "dd.mm.yyyy")
    ses.FindById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = Format$(son, "dd.mm.yyyy")
    ses.FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/ATIL"
Calistir:
    b0 = Timer
    ses.FindById("wnd[0]/tbar[1]/btn[8]").Press
    ATIL_BEKLE ses
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | sorgu " & AKSORA_SAYI_METNI(ATIL_GECEN(b0), "0.0") & " sn | " & ATIL_SBAR(ses)
    Set pop = ses.FindById("wnd[1]", False)
    If Not pop Is Nothing Then
        t = vbNullString
        On Error Resume Next
        t = pop.Text
        On Error GoTo 0
        ATIL_METIN_TOPLA pop, t, 0
        AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | SAP penceresi: " & t
        If ATIL_BOS_MESAJ(t) Then
            pop.SendVKey 0: ATIL_BEKLE ses
            Exit Function
        End If
        Err.Raise vbObjectError + 2346, , "SAP penceresi: " & Trim$(t)
    End If
    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
    If alv Is Nothing And Not ses.FindById("wnd[0]/usr/ctxtWERKS-LOW", False) Is Nothing Then
        t = ATIL_SBAR(ses)
        If ATIL_BOS_MESAJ(t) Then Exit Function
        ' /ATIL duzeni bu liste gosteriminde yoksa varsayilan duzen denenir (gerekli sutunlar basliktan bulunur).
        If InStr(Y_NORM(t), "DUZEN") > 0 And Not varsayilan Then
            AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | /ATIL düzeni kullanýlamadý (" & t & "); varsayýlan düzenle deneniyor"
            varsayilan = True
            ses.FindById("wnd[0]/usr/ctxtALV_DEF").Text = vbNullString
            GoTo Calistir
        End If
        If InStr(Y_NORM(t), "DUZEN") > 0 Then Err.Raise vbObjectError + 2347, , _
            "Bu sistemde MB51 için /ATIL düzeni yok (SAP: " & t & "). MB51'de listeyi açýp sütunlarý ayarlayýn ve /ATIL adýyla kaydedin."
        Err.Raise vbObjectError + 2347, , "MB51 sonuç vermedi: " & t
    End If
    v = AKSORA_ALV_METIN_AL(ses, alv, "ATIL_" & uy)
    If Not IsArray(v) Then Err.Raise vbObjectError + 2348, , "MB51 listesi okunamadý (yerel dosya). Ayrýntý: AKSORA_SAP_TESHIS.txt"
    ATIL_MB51_COZ v, uy, sat, etiket
End Function

Private Function ATIL_DIZI(ByVal d As Object) As String()
    Dim a() As String, k As Variant, i As Long
    If d.Count = 0 Then Err.Raise vbObjectError + 2349, , "Sorgulanacak malzeme listesi boþ."
    ReDim a(0 To d.Count - 1)
    For Each k In d.keys
        a(i) = CStr(k): i = i + 1
    Next k
    ATIL_DIZI = a
End Function

Public Sub AKSORA_ATIL_SAP_OKU()
    Dim tum As Object, il As Variant, sis As Variant, ses As Object, msg As String, oturumlar As Object
    Dim sonuc As Object, adim As String, eksik As String, atlanan As String, bugun As Date
    Dim adet As Long, toplam As Long, no As Long, metin As String, ev As Boolean, sure As Double, ozet As String
    Dim tamam As Object, hatalar As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ADMIN_TOPLU Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error GoTo Hata
    adim = "Stok verisi okunuyor"
    Set tum = ATIL_TUM_STOK()
    For Each il In ADM_ILLER()
        If tum(CStr(il))(1) = 0 Then eksik = eksik & ", " & Y_IL_ADI(CStr(il))
    Next il
    If Len(eksik) > 0 Then
        MsgBox "Þu illerin Malzeme stok verisi yok: " & Mid$(eksik, 3) & vbCrLf & _
            "Önce Malzeme biriminde bütün illerin stoklarýný güncelleyin.", vbExclamation, "AKSORA | ATIL STOK"
        Exit Sub
    End If
    adim = "SAP oturumlarý"
    Set oturumlar = CreateObject("Scripting.Dictionary")
    For Each sis In Array("CDP", "EPP")
        msg = vbNullString
        Set ses = Nothing
        Set ses = SAP_SESSION_AL(CStr(sis), msg)
        If ses Is Nothing Then atlanan = atlanan & ", " & sis Else oturumlar.Add CStr(sis), ses
    Next sis
    If oturumlar.Count = 0 Then
        MsgBox "SAP CDP ve EPP oturumu bulunamadý." & vbCrLf & "SAP'ye giriþ yapýn, sonra tekrar deneyin.", vbExclamation, "AKSORA | ATIL STOK"
        Exit Sub
    End If
    If MsgBox("SAP'den atýl stok verisi okunacak (MB51)." & vbCrLf & vbCrLf & _
        "Sistem: " & Join(oturumlar.keys, " + ") & _
        IIf(Len(atlanan) > 0, vbCrLf & "Açýk oturum yok: " & Mid$(atlanan, 3) & " (eski kaydý korunur)", "") & vbCrLf & _
        "Süre: üretim yeri baþýna yaklaþýk 1-4 dakika." & vbCrLf & vbCrLf & "Baþlasýn mý?", vbYesNo + vbQuestion, "AKSORA | ATIL STOK") <> vbYes Then Exit Sub
    AKSORA_ADMIN_IPTAL = False
    AKSORA_ATIL_CALISIYOR = True
    mAtilBas = Timer: mAtilSon = 0
    mAtilKart = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "SAP'den okunuyor...", 0.02, True)
    AKSORA_SAP_OLE_SESSIZ True
    ev = application.EnableEvents
    application.EnableEvents = False
    bugun = Date
    Set sonuc = CreateObject("Scripting.Dictionary"): sonuc.CompareMode = vbTextCompare
    toplam = oturumlar.Count * 15
    Set tamam = CreateObject("Scripting.Dictionary")
    For Each sis In oturumlar.keys
        If AKSORA_ADMIN_IPTAL Then
            hatalar = hatalar & vbCrLf & "• " & sis & ": baþlatýlmadý (durduruldu)"
        Else
            metin = ATIL_SISTEM_OKU(oturumlar(sis), CStr(sis), tum, sonuc, bugun, adet, toplam)
            If Len(metin) = 0 Then tamam.Add CStr(sis), oturumlar(sis) Else hatalar = hatalar & vbCrLf & "• " & sis & ": " & metin
            metin = vbNullString
        End If
    Next sis
    If tamam.Count = 0 Then adim = "SAP okumasý": Err.Raise vbObjectError + 2351, , Mid$(hatalar, 3)
    adim = "Kayýt yazýlýyor"
    ATIL_ADIM 0.96, "Kayýt yazýlýyor..."
    ATIL_KAYDET sonuc, tum, tamam
    sure = ATIL_GECEN(mAtilBas)
    AKSORA_SAP_TESHIS_YAZ "Atýl stok SAP okumasý tamamlandý: " & AKSORA_SAYI_METNI(sure, "0") & " sn"
    adim = "Sayfa hazýrlanýyor"
    ATIL_ADIM 0.98, "Sayfa hazýrlanýyor..."
    GoTo Bitir
Hata:
    no = Err.Number: metin = Err.Description
    Resume Bitir
Bitir:
    On Error Resume Next
    If Not oturumlar Is Nothing And AKSORA_ATIL_CALISIYOR Then
        For Each sis In oturumlar.keys
            Set ses = oturumlar(sis)
            If Not ses.FindById("wnd[1]", False) Is Nothing Then ses.FindById("wnd[1]/tbar[0]/btn[12]").Press
            ' SAP hafizasindaki malzeme (MAT) bosaltilir; yoksa sonraki MB52 stok okumalari tek malzemeyle calisir.
            ses.StartTransaction "MB51"
            AKSORA_SAP_MALZEME_TEMIZLE ses
            ses.FindById("wnd[0]").SendVKey 0
            ses.StartTransaction "SESSION_MANAGER"
        Next sis
    End If
    AKSORA_SAP_OLE_SESSIZ False
    AKSORA_ARKA_PLAN_BITIR
    If AKSORA_ATIL_CALISIYOR Then application.EnableEvents = ev
    application.StatusBar = False
    If no = 0 Then
        ATIL_GOSTER "TUM"
        ozet = ATIL_OZET_METNI()
    End If
    If mAtilKart Then AKSORA_BEKLE_KAPAT
    mAtilKart = False
    AKSORA_ATIL_CALISIYOR = False
    AKSORA_ADMIN_IPTAL = False
    On Error GoTo 0
    If no <> 0 Then
        MsgBox "Atýl stok SAP okumasý tamamlanamadý." & vbCrLf & vbCrLf & "Adým: " & adim & vbCrLf & "Hata: " & metin & vbCrLf & vbCrLf & _
            "Önceki atýl stok kaydý deðiþtirilmedi.", vbExclamation, "AKSORA | ATIL STOK"
    ElseIf Len(hatalar) > 0 Then
        MsgBox "Atýl stok SAP okumasý kýsmen tamamlandý." & vbCrLf & "Okunan ve kaydedilen: " & Join(tamam.keys, " + ") & vbCrLf & _
            "Süre: " & AKSORA_SURE_METNI(sure) & vbCrLf & vbCrLf & "Okunamayan (eski kaydý korunur):" & hatalar & vbCrLf & vbCrLf & ozet, vbExclamation, "AKSORA | ATIL STOK"
    ElseIf adet > 0 Then
        MsgBox "Atýl stok SAP okumasý tamamlandý." & vbCrLf & "Süre: " & AKSORA_SURE_METNI(sure) & vbCrLf & vbCrLf & ozet, vbInformation, "AKSORA | ATIL STOK"
    End If
End Sub

'---------- ATIL STOK: hesap ----------

Private Function ATIL_DURUM_ADI(ByVal durum As Long) As String
    Select Case durum
        Case 3: ATIL_DURUM_ADI = "ATIL"
        Case 2: ATIL_DURUM_ADI = "3 AY KALDI"
        Case 1: ATIL_DURUM_ADI = "6 AY KALDI"
        Case 0: ATIL_DURUM_ADI = "NORMAL"
        Case Else: ATIL_DURUM_ADI = "SORGULANMADI"
    End Select
End Function

' Kalemler: Array(0 il, 1 sistem, 2 UY, 3 malzeme, 4 tanim, 5 kategori, 6 stok, 7 birim, 8 stok degeri, 9 riskli girisin tarihi,
'   10 durum (-1 sorgulanmadi, 0 normal, 1 6 ay, 2 3 ay, 3 atil), 11 gun metni, 12 stok degisimi, 13 oneri ili, 14 oneri nedeni,
'   15 depolar, 16 stogu dustu mu, 17 gonderilecek, 18 oncelik, 19 deger (1 6 ay, 2 3 ay, 3 atil), 20 miktar (ayni sira),
'   21 riskli miktar, 22 riskli deger)
' bilgi: Array(okuma CDP, okuma EPP, sorgulanmayan, stogu dusen, stok tarihleri, eski kural kaydi var mi)
Private Function ATIL_HESAPLA(ByRef bilgi As Variant) As Collection
    Dim tum As Object, kayit As Object, okuma As Variant, c As Collection, il As Variant, s As Object, k As Variant
    Dim uy As String, mal As String, sis As String, o As Variant, kr As Variant, durum As Long, q() As Double
    Dim gunMetin As String, girisMetin As String, degisim As String, qq As Variant, quy As String, enIyi As Double
    Dim oIl As String, oMetin As String, qs As Double, sorgusuz As Long, dusen As Long, tarihler As String, ok As Double, s2 As Object
    Dim sarf6 As Double, oSarf As Double, oStok As Double, oMik As Double, oOnc As String, birim As String, bd As Double
    Dim risk As Double, dustu As Boolean, tv As Variant, tq As Variant
    Set c = New Collection
    Set tum = ATIL_TUM_STOK()
    ATIL_KARSILASTIR tum
    Set kayit = ATIL_KAYIT_OKU(okuma)
    For Each il In ADM_ILLER()
        If tum(CStr(il))(1) > 0 Then
            tarihler = tarihler & "  " & Y_IL_ADI(CStr(il)) & " " & Format$(tum(CStr(il))(1), "dd.mm.yyyy")
        Else
            tarihler = tarihler & "  " & Y_IL_ADI(CStr(il)) & " -"
        End If
        Set s = tum(CStr(il))(0)
        For Each k In s.keys
            AKSORA_BEKLE_NABIZ
            uy = Left$(CStr(k), InStr(CStr(k), "|") - 1): mal = Mid$(CStr(k), InStr(CStr(k), "|") + 1)
            sis = ATIL_SISTEM(uy)
            o = s(k)
            If sis = "CDP" Then ok = okuma(0) Else ok = okuma(1)
            oIl = vbNullString: oMetin = vbNullString: degisim = vbNullString: gunMetin = vbNullString: girisMetin = vbNullString
            oSarf = 0: oStok = 0: oMik = 0: oOnc = vbNullString: birim = CStr(o(2)): risk = 0: dustu = False
            tv = Array(0#, 0#, 0#, 0#): tq = Array(0#, 0#, 0#, 0#)
            If o(0) > 0 Then bd = o(1) / o(0) Else bd = 0
            If kayit.Exists(CStr(k)) And ok > 0 Then
                kr = kayit(CStr(k))
                durum = ATIL_YAS(o(0), CStr(kr(1)), kr(2), q, gunMetin, girisMetin)
                tq = Array(0#, q(1), q(2), q(3))
                tv = Array(0#, q(1) * bd, q(2) * bd, q(3) * bd)
                risk = q(1) + q(2) + q(3)
                If o(0) > kr(5) + 0.0005 Then
                    degisim = "ARTTI"
                ElseIf o(0) < kr(5) - 0.0005 Then
                    degisim = "DÜÞTÜ": dustu = True: dusen = dusen + 1
                Else
                    degisim = "DEÐÝÞMEDÝ"
                End If
                ' Transfer onerisi (ayni sistem): son 6 ayda bu malzemeyi kullanan iller icinden eksigi en buyuk olan
                ' (6 aylik kullanim - stogu). YUKSEK: stogu 6 aylik kullanimindan az; DUSUK: stogu yetiyor.
                If risk > 0.0000001 Then
                    For Each qq In ADM_ILLER()
                        If CStr(qq) <> CStr(il) Then
                            quy = AKSORA_IL_TESIS(CStr(qq), sis)
                            If kayit.Exists(quy & "|" & mal) Then
                                sarf6 = kayit(quy & "|" & mal)(3)
                                If sarf6 > 0 Then
                                    qs = 0
                                    Set s2 = tum(CStr(qq))(0)
                                    If s2.Exists(quy & "|" & mal) Then qs = s2(quy & "|" & mal)(0)
                                    If Len(oIl) = 0 Or sarf6 - qs > enIyi Or (sarf6 - qs = enIyi And sarf6 > oSarf) Then
                                        enIyi = sarf6 - qs: oIl = CStr(qq): oSarf = sarf6: oStok = qs
                                    End If
                                End If
                            End If
                        End If
                    Next qq
                    If Len(oIl) > 0 Then
                        If oSarf > oStok + 0.0005 Then
                            oOnc = "YÜKSEK": oMik = application.Min(risk, oSarf - oStok)
                        Else
                            oOnc = "DÜÞÜK": oMik = application.Min(risk, oSarf)
                        End If
                        oMetin = "Son 6 ayda " & ADM_ADET(oSarf) & " " & birim & " kullandý; deposunda " & _
                            IIf(oStok > 0, ADM_ADET(oStok) & " " & birim & " var.", "hiç yok.")
                    End If
                End If
            Else
                durum = -1: sorgusuz = sorgusuz + 1
            End If
            c.Add Array(CStr(il), sis, uy, mal, CStr(o(3)), CStr(o(4)), o(0), birim, o(1), girisMetin, durum, gunMetin, degisim, oIl, oMetin, o(5), _
                dustu, oMik, oOnc, tv, tq, risk, risk * bd)
        Next k
    Next il
    bilgi = Array(okuma(0), okuma(1), sorgusuz, dusen, Trim$(tarihler), mAtilEski)
    Set ATIL_HESAPLA = c
End Function

Private Function ATIL_OZET_METNI() As String
    Dim c As Collection, bilgi As Variant, x As Variant, t(0 To 3) As Double, a(0 To 3) As Long, n As Long, j As Long
    Set c = ATIL_HESAPLA(bilgi)
    For Each x In c
        If x(10) >= 1 Then
            For j = 1 To 3
                If x(20)(j) > 0.0000001 Then t(j) = t(j) + x(19)(j): a(j) = a(j) + 1
            Next j
            If Len(x(13)) > 0 Then n = n + 1
        End If
    Next x
    ATIL_OZET_METNI = "ATIL: " & a(3) & " kalem, " & ADM_TL(t(3)) & vbCrLf & _
        "3 AY KALDI: " & a(2) & " kalem, " & ADM_TL(t(2)) & vbCrLf & _
        "6 AY KALDI: " & a(1) & " kalem, " & ADM_TL(t(1)) & vbCrLf & _
        "Transfer önerisi: " & n & " kalem"
End Function

'---------- ATIL STOK: sayfa ----------
Public Sub AKSORA_ATIL_AC()
    Dim ac As Boolean, no As Long, metin As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Atýl stok sayfasý hazýrlanýyor...")
    On Error GoTo Son
    ATIL_GOSTER vbNullString
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Atýl stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_ATIL_IL_SEC()
    Dim ac As Boolean, no As Long, metin As String, il As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error Resume Next
    il = ThisWorkbook.Worksheets(ATIL_SAYFA).Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(il) = 0 Then il = "TUM"
    ac = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Atýl stok sayfasý hazýrlanýyor...")
    On Error GoTo Son
    ATIL_GOSTER il
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Atýl stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Private Sub ATIL_GOSTER(ByVal il As String, Optional ByVal gorunum As String = "")
    Dim ws As Worksheet, ev As Boolean, hesap As XlCalculation, no As Long, metin As String
    ev = application.EnableEvents: hesap = application.Calculation
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    Set ws = ATIL_SAYFA_KUR()
    If Len(gorunum) > 0 Then ws.Unprotect "": ws.Range("BX21").Value2 = gorunum
    If Len(il) = 0 Then il = CStr(ws.Range("BX20").Value2)
    If Len(il) = 0 Then il = "TUM"
    ATIL_CIZ ws, il
    ws.Visible = xlSheetVisible
    mYonlendirme = True
    ws.Select Replace:=True
    mYonlendirme = False
    AKSORA_YATIRIM_GORUNUM ws
Cikis:
    mYonlendirme = False
    application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = True
    If no <> 0 Then Err.Raise no, "ATIL_GOSTER", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_ATIL_GORUNUM()
    Dim ac As Boolean, no As Long, metin As String, gor As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error Resume Next
    gor = ThisWorkbook.Worksheets(ATIL_SAYFA).Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(gor) = 0 Then gor = "OZET"
    ac = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Atýl stok sayfasý hazýrlanýyor...")
    On Error GoTo Son
    ATIL_GOSTER vbNullString, gor
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Atýl stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

' Transfer onerileri PDF: gecici cikti sayfasinda ozet + tablo (kurum yazicisina uygun acik gri / siyah),
' filtre uygulanmissa yalniz gorunen satirlar. Yatay, tek sayfa genisligi, baslik satiri her sayfada.
Public Sub AKSORA_ATIL_TRANSFER_PDF()
    Dim ws As Worksheet, cs As Worksheet, onceki As Object, adres As String, yol As String, il As String, ilAd As String, metin As String
    Dim v As Variant, a() As Variant, kol As Variant, hedef As Object, k As Variant, r As Long, n As Long, c As Long, ilk As Long
    Dim yuksek As Long, toplam As Double, ozet As String, ev As Boolean, sc As Boolean, da As Boolean, no As Long
    Dim anah() As String, idx() As Long, adlar() As String, kart As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(ATIL_SAYFA)
    adres = CStr(ws.Range("BX22").Value2)
    If Len(adres) = 0 Or CStr(ws.Range("BX21").Value2) <> "TRANSFER" Then
        MsgBox "Önce TRANSFER ÖNERÝLERÝ sekmesini açýn.", vbInformation, "AKSORA | ATIL STOK": Exit Sub
    End If
    il = CStr(ws.Range("BX20").Value2)
    If il = "TUM" Or Len(il) = 0 Then ilAd = "Tüm Ýller" Else ilAd = Y_IL_ADI(il)
    ' Gorunen satirlar: Nereden, Malzeme, Tanim, Depo, Riskli miktar, Riskli deger, Durum, Nereye, Gonderilecek, Oncelik, Neden (Sistem cikmaz).
    v = ws.Range(adres).Value2
    ilk = ws.Range(adres).Row
    kol = Array(1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    ReDim a(1 To UBound(v, 1), 1 To 11)
    Set hedef = CreateObject("Scripting.Dictionary"): hedef.CompareMode = vbTextCompare
    For r = 2 To UBound(v, 1)
        If Not ws.Rows(ilk + r - 1).Hidden Then
            n = n + 1
            For c = 0 To 10: a(n, c + 1) = v(r, kol(c)): Next c
            toplam = toplam + ATIL_HUCRE_SAYI(v(r, 7))
            If CStr(v(r, 11)) = "YÜKSEK" Then yuksek = yuksek + 1
            hedef(CStr(v(r, 9))) = hedef(CStr(v(r, 9))) + 1
        End If
    Next r
    If n = 0 Then MsgBox "Yazdýrýlacak satýr yok (filtreyi kontrol edin).", vbInformation, "AKSORA | ATIL STOK": Exit Sub
    kart = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "PDF hazýrlanýyor (" & n & " satýr)...", 0.3)
    ' Nereye ozeti (cok olandan aza)
    ReDim anah(1 To hedef.Count): ReDim idx(1 To hedef.Count): ReDim adlar(1 To hedef.Count)
    r = 0
    For Each k In hedef.keys
        r = r + 1: idx(r) = r: adlar(r) = CStr(k)
        anah(r) = Format$(99999 - hedef(k), "00000") & CStr(k)
    Next k
    Y_SIRALA anah, idx, 1, hedef.Count
    For r = 1 To hedef.Count
        ozet = ozet & IIf(r > 1, "   •   ", "") & adlar(idx(r)) & ": " & hedef(adlar(idx(r))) & " kalem"
    Next r
    ev = application.EnableEvents: sc = application.ScreenUpdating: da = application.DisplayAlerts
    application.EnableEvents = False: application.ScreenUpdating = False: application.DisplayAlerts = False
    Set onceki = ActiveSheet
    Set cs = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    With cs.Cells.Font
        .Name = "Aptos": .Size = 9: .Color = vbBlack
    End With
    cs.Range("A1").Value2 = "ATIL STOK TRANSFER ÖNERÝLERÝ   |   " & ilAd
    cs.Range("A1").Font.Size = 14: cs.Range("A1").Font.bold = True
    cs.Range("A2").Value2 = "Çoruh EDAÞ  •  Malzeme Stok Birimi  •  " & Format$(Now, "dd.mm.yyyy hh:nn") & _
        "  •  Atýl: 24 aydan eski stok (D, K ve yüklenici depolarý)  •  Öneri: ayný sistemde son 6 ayda bu malzemeyi kullanan il"
    cs.Range("A2").Font.Color = RGB(64, 64, 64)
    cs.Range("A3").Value2 = "ÖZET:  " & n & " öneri   •   Toplam riskli deðer: " & ADM_TL(toplam) & "   •   YÜKSEK öncelik: " & yuksek & "   •   DÜÞÜK öncelik: " & (n - yuksek)
    cs.Range("A3").Font.bold = True: cs.Range("A3").Font.Size = 10
    cs.Range("A4").Value2 = "NEREYE:  " & ozet
    cs.Range("A5").Value2 = "YÜKSEK: alacak ilin deposu 6 aylýk kullanýmýna yetmiyor; gönderilecek = eksik kadar.   DÜÞÜK: alacak ilin stoðu yetiyor."
    cs.Range("A5").Font.Italic = True: cs.Range("A5").Font.Color = RGB(64, 64, 64)
    cs.Range("A7:K7").Value2 = Array("Nereden", "Malzeme", "Taným", "Depo", "Riskli miktar", "Riskli deðer", "Durum", "Nereye", "Gönderilecek", "Öncelik", "Neden")
    With cs.Range("A7:K7")
        .Font.bold = True: .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    cs.Rows(7).RowHeight = 24
    cs.Range(cs.Cells(8, 2), cs.Cells(7 + n, 2)).NumberFormat = "@"
    cs.Range("A8").Resize(n, 11).Value2 = a
    With cs.Range(cs.Cells(7, 1), cs.Cells(7 + n, 11))
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(166, 166, 166)
        .VerticalAlignment = xlCenter
    End With
    cs.Range(cs.Rows(8), cs.Rows(7 + n)).RowHeight = 15
    cs.Range(cs.Cells(8, 4), cs.Cells(7 + n, 4)).WrapText = True
    ' Kesirsiz miktarda virgul gosterilmez (ekrandaki tablolarla ayni bicim).
    AKSORA_V13_MIKTAR_BICIMI cs.Range(cs.Cells(8, 5), cs.Cells(7 + n, 5))
    cs.Range(cs.Cells(8, 6), cs.Cells(7 + n, 6)).NumberFormat = "#,##0.00 ""TL"""
    cs.Range(cs.Cells(8, 1), cs.Cells(7 + n, 2)).HorizontalAlignment = xlCenter
    cs.Range(cs.Cells(8, 4), cs.Cells(7 + n, 4)).HorizontalAlignment = xlCenter
    cs.Range(cs.Cells(8, 7), cs.Cells(7 + n, 10)).HorizontalAlignment = xlCenter
    cs.Range(cs.Cells(8, 7), cs.Cells(7 + n, 10)).Font.bold = True
    cs.Range(cs.Cells(7, 6), cs.Cells(7 + n, 6)).HorizontalAlignment = xlRight
    kol = Array(9, 11, 32, 20, 9, 13, 10, 10, 11, 8, 41)
    For c = 0 To 10: cs.Columns(c + 1).ColumnWidth = kol(c): Next c
    For r = 8 To 7 + n
        If Len(CStr(cs.Cells(r, 4).Value2)) > 20 Then cs.Rows(r).AutoFit
    Next r
    yol = AKSORA_KAYIT_KLASORU() & application.PathSeparator & "ATIL_TRANSFER_" & IIf(il = "TUM" Or Len(il) = 0, "TUM_ILLER", il) & "_" & Format$(Now, "yyyymmdd_hhnn") & ".pdf"
    On Error Resume Next
    With cs.PageSetup
        .PrintArea = cs.Range(cs.Cells(1, 1), cs.Cells(7 + n, 11)).Address
        .PrintTitleRows = "$7:$7"
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .LeftMargin = application.InchesToPoints(0.3): .RightMargin = application.InchesToPoints(0.3)
        .TopMargin = application.InchesToPoints(0.4): .BottomMargin = application.InchesToPoints(0.5)
        .CenterFooter = "Sayfa &P / &N"
        .LeftFooter = "AKSORA  •  Atýl Stok Transfer Önerileri  •  " & ilAd
    End With
    On Error GoTo Hata
    cs.ExportAsFixedFormat Type:=xlTypePDF, Filename:=yol, Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
Cikis:
    On Error Resume Next
    If Not cs Is Nothing Then cs.Delete
    If Not onceki Is Nothing Then onceki.Activate
    If kart Then AKSORA_BEKLE_KAPAT
    application.DisplayAlerts = da: application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If no <> 0 Then MsgBox "PDF oluþturulamadý: " & metin, vbExclamation, "AKSORA | ATIL STOK"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

' Malzemenin bulundugu tum depolar ve miktarlari: "D001 (3), K002 (5)".
Private Function ATIL_DEPO_METIN(ByVal dp As Object) As String
    Dim k As Variant, n As Long
    For Each k In dp.keys
        n = n + 1
        ATIL_DEPO_METIN = ATIL_DEPO_METIN & IIf(n > 1, ", ", "") & CStr(k) & " (" & ADM_ADET(CDbl(dp(k)(1))) & ")"
    Next k
End Function

Public Function AKSORA_ATIL_SAYFA(ByVal sh As Object) As Boolean
    If Not TypeOf sh Is Worksheet Then Exit Function
    If sh.Name <> ATIL_SAYFA Then Exit Function
    AKSORA_ATIL_SAYFA = (CStr(sh.Range("BX1").Value2) = ATIL_ETIKET)
End Function

' Sayfa musteri rapor sayfasinin basligiyla kurulur (veri tasinmaz).
Private Function ATIL_SAYFA_KUR() As Worksheet
    Dim ws As Worksheet, kaynak As Worksheet, gor As XlSheetVisibility, onceki As Collection, sh As Object, varmi As Boolean
    Dim da As Boolean
    Set ws = ADM_WS(ATIL_SAYFA)
    If Not ws Is Nothing Then
        If CStr(ws.Range("BX1").Value2) <> ATIL_ETIKET Then Err.Raise vbObjectError + 2360, , "Ayný isimde baþka bir sayfa var: " & ATIL_SAYFA
        Set ATIL_SAYFA_KUR = ws
        Exit Function
    End If
    Set kaynak = ADM_WS("M_CDP_RAPOR")
    If kaynak Is Nothing Then Err.Raise vbObjectError + 2361, , "Sayfa þablonu (M_CDP_RAPOR) bulunamadý."
    da = application.DisplayAlerts: application.DisplayAlerts = False
    gor = kaynak.Visible
    kaynak.Visible = xlSheetVisible
    Set onceki = New Collection
    For Each sh In ThisWorkbook.Sheets
        onceki.Add sh.Name, sh.Name
    Next sh
    kaynak.Copy after:=kaynak
    For Each sh In ThisWorkbook.Sheets
        varmi = False
        On Error Resume Next
        varmi = (Len(onceki(sh.Name)) > 0)
        On Error GoTo 0
        If Not varmi Then Set ws = sh: Exit For
    Next sh
    kaynak.Visible = gor
    application.DisplayAlerts = da
    If ws Is Nothing Then Err.Raise vbObjectError + 2362, , "Kopya sayfa bulunamadý."
    ws.Name = ATIL_SAYFA
    ws.Unprotect ""
    ws.Range("BX1:BX40").ClearContents
    ws.Range("BX1").Value2 = ATIL_ETIKET
    On Error Resume Next
    ws.Move after:=ThisWorkbook.Worksheets("YÜKLENÝCÝ_RAPOR")
    On Error GoTo 0
    Set ATIL_SAYFA_KUR = ws
End Function

Private Sub ATIL_DURUM_RENK(ByVal r As Range)
    Y_DURUM_KF r, "ATIL", RGB(220, 38, 38), RGB(255, 255, 255)
    Y_DURUM_KF r, "3 AY KALDI", RGB(254, 215, 170), RGB(154, 52, 18)
    Y_DURUM_KF r, "6 AY KALDI", RGB(254, 243, 199), RGB(146, 64, 14)
    r.HorizontalAlignment = xlCenter
End Sub

Private Sub ATIL_CIZ(ByVal ws As Worksheet, ByVal ilSec As String)
    Dim c As Collection, bilgi As Variant, x As Variant, w As Double, s As Shape, y As Double, kw As Double, bw As Double
    Dim i As Long, j As Long, r As Long, a1 As Long, n As Long, k As Variant, o As Variant, t As Variant, il As Variant
    Dim kpiT(0 To 3) As Double, kpiA(0 To 3) As Long, oneriA As Long, oneriT As Double, kats As Variant, katD As Object
    Dim karne As Object, kapsam As Object, dp As Object, dk As Variant, anah() As String, idx() As Long, satir() As Variant
    Dim liste As Collection, sw As Double, enAz As Double, enCok As Double, ayrim As Boolean
    Dim gec As Boolean, okMetin As String, kod As String, tablo() As Variant
    Dim gor As String, yuksek As Long, gAd As Variant, gKod As Variant, bw2 As Double, t2 As Long, pay As Double
    ws.Unprotect ""
    Y_SAYFA_SIFIRLA ws
    ' Sablondan gelen AKSORA amblemi (filigran) bu sayfada kullanilmaz.
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(ws.Shapes(i).Name, "FILIGRAN") > 0 Then ws.Shapes(i).Delete
    Next i
    Set s = Y_SEKIL(ws, "btnAnaPanelDon")
    If Not s Is Nothing Then s.AlternativeText = vbNullString: s.OnAction = BIRIM_MAKRO("ANA_PANELE_DON")
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | ATIL STOK  |  " & IIf(ilSec = "TUM", "Tüm Ýller", Y_IL_ADI(ilSec))
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    sw = application.Max(170, w - 1285)
    Y_KOLONLAR ws, Array(90, 60, 95, 230, 140, 95, 115, 105, 125, 105, 125, sw)
    ws.Range("BX20").Value2 = ilSec
    gor = CStr(ws.Range("BX21").Value2)
    If gor <> "TRANSFER" And gor <> "LISTE" Then gor = "OZET"
    Set c = ATIL_HESAPLA(bilgi)
    ' kapsam degeri (atil orani icin) ve karne
    Set kapsam = CreateObject("Scripting.Dictionary"): kapsam.CompareMode = vbTextCompare
    Set karne = CreateObject("Scripting.Dictionary"): karne.CompareMode = vbTextCompare
    Set katD = CreateObject("Scripting.Dictionary"): katD.CompareMode = vbTextCompare
    Set liste = New Collection
    For Each x In c
        If ilSec = "TUM" Or x(0) = ilSec Then
            ' karne anahtari: tum illerde il, tek ilde depo
            If ilSec = "TUM" Then
                kod = x(0)
                kapsam(kod) = kapsam(kod) + x(8)
            Else
                Set dp = x(15)
                For Each dk In dp.keys
                    kapsam(x(1) & "|" & dk) = kapsam(x(1) & "|" & dk) + dp(dk)(2)
                    If Not karne.Exists(x(1) & "|" & dk) Then karne.Add x(1) & "|" & dk, Array(CStr(dp(dk)(0)), 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                Next dk
            End If
            If x(10) >= 1 Then
                ' Bir malzemenin stogu birden fazla basamakta olabilir (ornek: bir kismi ATIL, bir kismi 3 AY):
                ' her basamak kendi miktar ve degeriyle sayilir.
                For t2 = 1 To 3
                    If x(20)(t2) > 0.0000001 Then kpiT(t2) = kpiT(t2) + x(19)(t2): kpiA(t2) = kpiA(t2) + 1
                Next t2
                If Len(x(13)) > 0 Then oneriA = oneriA + 1: oneriT = oneriT + x(22)
                If x(18) = "YÜKSEK" Then yuksek = yuksek + 1
                liste.Add x
                ' kategori: Array(atil kalem, atil deger, 3ay kalem, 3ay deger, 6ay kalem, 6ay deger, stogu dusen, riskli kalem)
                If Not katD.Exists(x(5)) Then katD.Add x(5), Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                o = katD(x(5))
                For t2 = 1 To 3
                    If x(20)(t2) > 0.0000001 Then j = (3 - t2) * 2: o(j) = o(j) + 1: o(j + 1) = o(j + 1) + x(19)(t2)
                Next t2
                If x(16) Then o(6) = o(6) + 1
                o(7) = o(7) + 1
                katD(x(5)) = o
                ' karne: Array(ad, atil kalem, atil deger, 3ay kalem, 3ay deger, 6ay kalem, 6ay deger, stogu dusen, riskli kalem)
                If ilSec = "TUM" Then
                    If Not karne.Exists(kod) Then karne.Add kod, Array(Y_IL_ADI(kod), 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                    o = karne(kod)
                    For t2 = 1 To 3
                        If x(20)(t2) > 0.0000001 Then j = (3 - t2) * 2: o(1 + j) = o(1 + j) + 1: o(2 + j) = o(2 + j) + x(19)(t2)
                    Next t2
                    If x(16) Then o(7) = o(7) + 1
                    o(8) = o(8) + 1
                    karne(kod) = o
                Else
                    ' Depo karnesi: riskli deger depolara stok payiyla dagitilir.
                    Set dp = x(15)
                    For Each dk In dp.keys
                        o = karne(x(1) & "|" & dk)
                        If x(6) > 0 Then pay = dp(dk)(1) / x(6) Else pay = 0
                        For t2 = 1 To 3
                            If x(20)(t2) > 0.0000001 Then j = (3 - t2) * 2: o(1 + j) = o(1 + j) + 1: o(2 + j) = o(2 + j) + x(19)(t2) * pay
                        Next t2
                        If x(16) Then o(7) = o(7) + 1
                        o(8) = o(8) + 1
                        karne(x(1) & "|" & dk) = o
                    Next dk
                End If
            End If
        End If
    Next x
    If ilSec = "TUM" Then
        For Each il In ADM_ILLER()
            If Not karne.Exists(CStr(il)) Then karne.Add CStr(il), Array(Y_IL_ADI(CStr(il)), 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        Next il
    End If
    gec = (bilgi(0) > 0 Or bilgi(1) > 0)

    '--- KPI
    ws.Range("BZ21").Value2 = ADM_TL_KISA(kpiT(3)) & "  •  " & kpiA(3) & " kalem"
    ws.Range("BZ22").Value2 = ADM_TL_KISA(kpiT(2)) & "  •  " & kpiA(2) & " kalem"
    ws.Range("BZ23").Value2 = ADM_TL_KISA(kpiT(1)) & "  •  " & kpiA(1) & " kalem"
    ws.Range("BZ24").Value2 = oneriA & " kalem  •  " & yuksek & " yüksek öncelik"
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10: ws.Rows(8).RowHeight = 30: ws.Rows(9).RowHeight = 20: ws.Rows(10).RowHeight = 30: ws.Rows(11).RowHeight = 8
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "ATIL (24 AYDAN ESKÝ STOK)", "$BZ$21", RGB(220, 38, 38)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "3 AY KALDI (21-24 AY)", "$BZ$22", RGB(234, 88, 12)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "6 AY KALDI (18-21 AY)", "$BZ$23", RGB(202, 138, 4)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "TRANSFER ÖNERÝSÝ (AYNI SÝSTEM)", "$BZ$24", RGB(20, 103, 173)
    For i = 1 To 4
        Set s = Y_SEKIL(ws, "Y_KPI_DEGER_" & i)
        If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 15
    Next i

    '--- il secimi ve islemler
    y = ws.Rows(8).Top + 1
    bw = application.Min(118, (w - 2 * 150 - 7 * 6) / 6)
    i = 0
    For Each il In Array("TUM", "ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
        Set s = Y_DUGME(ws, "ATIL_IL_" & i, IIf(CStr(il) = "TUM", "TÜM ÝLLER", ATIL_IL_BUYUK(CStr(il))), i * (bw + 6), y, bw, 28, _
            IIf(CStr(il) = ilSec, RGB(0, 125, 110), RGB(148, 163, 184)), "AKSORA_ATIL_IL_SEC")
        s.AlternativeText = CStr(il)
        i = i + 1
    Next il
    Y_DUGME ws, "ATIL_YENILE", "YENÝLE", w - 300 - 6, y, 150, 28, RGB(20, 103, 173), "AKSORA_ATIL_AC"
    Y_DUGME ws, "ATIL_SAP", "SAP'DEN OKU (MB51)", w - 150, y, 150, 28, RGB(12, 45, 79), "AKSORA_ATIL_SAP_OKU"
    If gec Then
        okMetin = "Atýl: eldeki stoðun 24 aydan eski giriþe ait kýsmý  •  D, K ve yüklenici depolarý   |   SAP okumasý: CDP " & _
            IIf(bilgi(0) > 0, Format$(bilgi(0), "dd.mm.yyyy"), "-") & "  •  EPP " & IIf(bilgi(1) > 0, Format$(bilgi(1), "dd.mm.yyyy"), "-") & _
            "   |   Stok verisi: " & bilgi(4)
        If bilgi(3) > 0 Then okMetin = okMetin & "   |   Stoðu düþen: " & bilgi(3) & " kalem"
        If bilgi(2) > 0 Then okMetin = okMetin & "   |   SAP okumasýndan sonra gelen " & bilgi(2) & " kalem sonraki okumada deðerlendirilecek"
    ElseIf bilgi(5) Then
        okMetin = "Atýl kuralý deðiþti (giriþ yaþý + yüklenici depolarý). Saðdaki 'SAP'DEN OKU (MB51)' ile yeniden okuyun (SAP'de CDP ve EPP oturumu açýk olmalý)."
    Else
        okMetin = "Henüz SAP okumasý yok. Saðdaki 'SAP'DEN OKU (MB51)' düðmesine basýn (SAP'de CDP ve EPP oturumu açýk olmalý)."
    End If
    ws.Range("A9").Value2 = okMetin
    With ws.Range("A9").Font
        .Name = "Aptos": .Size = 9: .Italic = True: .Color = IIf(gec, RGB(100, 116, 139), RGB(185, 28, 28)): .bold = Not gec
    End With
    ws.Range("A9").VerticalAlignment = xlCenter
    If Not gec Then GoTo Son

    '--- gorunum sekmeleri: ozet / transfer / liste (tek bolum gosterilir, sayfa kisa kalir)
    gAd = Array("ÖZET  •  kategori ve karne", "TRANSFER ÖNERÝLERÝ  •  " & oneriA & " kalem", "ATIL LÝSTESÝ  •  " & liste.Count & " kalem")
    gKod = Array("OZET", "TRANSFER", "LISTE")
    bw2 = (w - 2 * 8) / 3
    y = ws.Rows(10).Top + 1
    For i = 0 To 2
        Set s = Y_DUGME(ws, "ATIL_GOR_" & i, CStr(gAd(i)), i * (bw2 + 8), y, bw2, 28, _
            IIf(CStr(gKod(i)) = gor, RGB(12, 45, 79), RGB(226, 232, 240)), "AKSORA_ATIL_GORUNUM")
        s.AlternativeText = CStr(gKod(i))
        s.TextFrame2.TextRange.Font.Size = 11
        If CStr(gKod(i)) <> gor Then s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(12, 45, 79)
    Next i
    If gor <> "OZET" Then GoTo Ayrinti

    '--- kategori ozeti
    r = 12
    Y_BOLUM ws, r, "KATEGORÝ ÖZETÝ", "Atýl ve atýla yaklaþan malzemelerin kategorilere daðýlýmý.", 12
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array("Kategori", "Kalem", "ATIL kalem", "ATIL deðeri", "3 AY kalem", _
        "3 AY deðeri", "6 AY kalem", "6 AY deðeri", "Toplam risk deðeri", "Risk payý", "Stoðu düþen", "Yorum")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    kats = Array("KABLO", "DÝREK", "PANO", "HÜCRE", "TRAFO", "DÝÐER")
    enCok = kpiT(1) + kpiT(2) + kpiT(3)
    For i = 0 To 5
        r = r + 1
        ws.Cells(r, 1).Value2 = kats(i)
        If katD.Exists(kats(i)) Then o = katD(kats(i)) Else o = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        ws.Cells(r, 2).Value2 = o(7)
        ws.Cells(r, 3).Value2 = o(0): ws.Cells(r, 4).Value2 = Round(o(1), 2)
        ws.Cells(r, 5).Value2 = o(2): ws.Cells(r, 6).Value2 = Round(o(3), 2)
        ws.Cells(r, 7).Value2 = o(4): ws.Cells(r, 8).Value2 = Round(o(5), 2)
        ws.Cells(r, 9).Value2 = Round(o(1) + o(3) + o(5), 2)
        If enCok > 0 Then ws.Cells(r, 10).Value2 = (o(1) + o(3) + o(5)) / enCok Else ws.Cells(r, 10).Value2 = 0
        ws.Cells(r, 11).Value2 = o(6)
        If o(7) = 0 Then
            ws.Cells(r, 12).Value2 = "Riskli malzeme yok"
        ElseIf o(1) >= o(3) And o(1) >= o(5) Then
            ws.Cells(r, 12).Value2 = "Aðýrlýk atýlda: transfer / ihtiyaç kontrolü"
        Else
            ws.Cells(r, 12).Value2 = "Atýla yaklaþýyor: kullaným planlayýn"
        End If
    Next i
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12))
    ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 22
    r = r + 1
    ws.Cells(r, 1).Value2 = "TOPLAM"
    For j = 2 To 9
        ws.Cells(r, j).Formula = "=SUM(" & ws.Range(ws.Cells(a1, j), ws.Cells(r - 1, j)).Address(False, False) & ")"
    Next j
    ws.Cells(r, 10).Value2 = IIf(enCok > 0, 1, 0)
    ws.Cells(r, 11).Formula = "=SUM(" & ws.Range(ws.Cells(a1, 11), ws.Cells(r - 1, 11)).Address(False, False) & ")"
    Y_TOPLAM_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12))
    ws.Rows(r).RowHeight = 22
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).Font.bold = True
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).IndentLevel = 1
    ws.Range(ws.Cells(a1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 8), ws.Cells(r, 9)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(a1, 3), ws.Cells(r, 4)).Font.Color = RGB(185, 28, 28)
    I_FIRE_ISI ws.Range(ws.Cells(a1, 9), ws.Cells(r - 1, 9))

    '--- il / depo karnesi
    r = r + 2
    If ilSec = "TUM" Then
        Y_BOLUM ws, r, "ÝL KARNESÝ", "Atýl oraný = atýl deðer / kapsamdaki (D, K ve yüklenici depolarý) stok deðeri. Sýra: atýl oraný düþükten yükseðe.", 12
    Else
        Y_BOLUM ws, r, "DEPO KARNESÝ  –  " & Y_IL_ADI(ilSec), "Riskli malzemelerin depolara daðýlýmý. Sýra: toplam risk deðeri yüksekten düþüðe.", 12
    End If
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array(IIf(ilSec = "TUM", "Ýl", "Depo"), IIf(ilSec = "TUM", "Kalem", "Sistem"), _
        "ATIL kalem", "ATIL deðeri", "3 AY kalem", "3 AY deðeri", "6 AY kalem", "6 AY deðeri", "Toplam risk deðeri", "Atýl oraný", _
        "Stoðu düþen", IIf(ilSec = "TUM", "Deðerlendirme", "Depo adý"))
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    n = karne.Count
    If n > 0 Then
        ReDim anah(1 To n): ReDim idx(1 To n): ReDim satir(1 To n)
        i = 0
        For Each k In karne.keys
            i = i + 1: idx(i) = i: satir(i) = k
            o = karne(k)
            If ilSec = "TUM" Then
                If kapsam(CStr(k)) > 0 Then t = o(2) / kapsam(CStr(k)) Else t = 0
                anah(i) = Format$(Round(t * 1000000, 0), "000000000000") & CStr(k)
            Else
                anah(i) = Format$(999999999999# - Round(o(2) + o(4) + o(6), 0), "000000000000") & CStr(k)
            End If
        Next k
        Y_SIRALA anah, idx, 1, n
        enAz = -1: enCok = -1
        For i = 1 To n
            k = satir(idx(i)): o = karne(k)
            r = r + 1
            If kapsam(CStr(k)) > 0 Then t = o(2) / kapsam(CStr(k)) Else t = 0
            If ilSec = "TUM" Then
                ws.Cells(r, 1).Value2 = o(0)
                ws.Cells(r, 2).Value2 = o(8)
                If enAz < 0 Then enAz = t
                enCok = t
            Else
                ws.Cells(r, 1).NumberFormat = "@": ws.Cells(r, 1).Value2 = Mid$(CStr(k), InStr(CStr(k), "|") + 1)
                ws.Cells(r, 2).Value2 = Left$(CStr(k), InStr(CStr(k), "|") - 1)
                ws.Cells(r, 12).Value2 = o(0)
            End If
            ws.Cells(r, 3).Value2 = o(1): ws.Cells(r, 4).Value2 = Round(o(2), 2)
            ws.Cells(r, 5).Value2 = o(3): ws.Cells(r, 6).Value2 = Round(o(4), 2)
            ws.Cells(r, 7).Value2 = o(5): ws.Cells(r, 8).Value2 = Round(o(6), 2)
            ws.Cells(r, 9).Value2 = Round(o(2) + o(4) + o(6), 2)
            ws.Cells(r, 10).Value2 = t
            ws.Cells(r, 11).Value2 = o(7)
        Next i
        If ilSec = "TUM" Then
            ayrim = (enCok > enAz + 0.0005)
            For i = 1 To n
                t = ws.Cells(a1 + i - 1, 10).Value2
                If i = 1 And ayrim Then
                    ws.Cells(a1 + i - 1, 12).Value2 = ChrW(9733) & " EN ÝYÝ"
                ElseIf i = n And ayrim Then
                    ws.Cells(a1 + i - 1, 12).Value2 = ChrW(9660) & " EN KÖTÜ"
                ElseIf t < 0.05 Then
                    ws.Cells(a1 + i - 1, 12).Value2 = "ÝYÝ"
                ElseIf t < 0.15 Then
                    ws.Cells(a1 + i - 1, 12).Value2 = "ORTA"
                Else
                    ws.Cells(a1 + i - 1, 12).Value2 = "ZAYIF"
                End If
            Next i
        End If
        Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12))
        ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 22
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).Font.bold = True
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ws.Cells(a1, 8), ws.Cells(r, 9)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(a1, 3), ws.Cells(r, 4)).Font.Color = RGB(185, 28, 28)
        I_FIRE_ISI ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9))
        If ilSec = "TUM" Then I_DURUM_RENK ws.Range(ws.Cells(a1, 12), ws.Cells(r, 12))
    End If

    GoTo Son

Ayrinti:
    n = liste.Count
    If gor = "TRANSFER" Then
        ' Siralama: oncelik (YUKSEK once), durum (ATIL once), deger.
        j = 0
        For Each x In liste
            If Len(x(13)) > 0 Then j = j + 1
        Next x
        r = 12
        Y_BOLUM ws, r, "TRANSFER ÖNERÝLERÝ (" & j & ")", "Ayný sistemde (CDP-CDP, EPP-EPP). YÜKSEK: alacak ilin deposu 6 aylýk kullanýmýna yetmiyor.", 10
        ws.Rows(r).RowHeight = 30
        r = r + 1
        Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array("Nereden", "Sistem", "Malzeme", "Taným", "Depo", _
            "Riskli miktar", "Riskli deðer", "Durum", "Nereye", "Gönderilecek", "Öncelik", "Neden")
        ws.Rows(r).RowHeight = 30
        a1 = r
        If j = 0 Then
            r = r + 1
            ws.Cells(r, 1).Value2 = "Öneri yok: riskli malzemeleri ayný sistemde son 6 ayda kullanan baþka il bulunmadý."
            ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(100, 116, 139)
        Else
            ReDim anah(1 To j): ReDim idx(1 To j): ReDim satir(1 To j)
            i = 0
            For Each x In liste
                If Len(x(13)) > 0 Then
                    i = i + 1: idx(i) = i: satir(i) = x
                    anah(i) = IIf(x(18) = "YÜKSEK", "0", "1") & CStr(9 - x(10)) & Format$(999999999999# - Round(x(22), 0), "000000000000") & x(0) & x(3)
                End If
            Next x
            Y_SIRALA anah, idx, 1, j
            ReDim tablo(1 To j, 1 To 12)
            For i = 1 To j
                x = satir(idx(i))
                tablo(i, 1) = Y_IL_ADI(CStr(x(0))): tablo(i, 2) = x(1): tablo(i, 3) = CStr(x(3)): tablo(i, 4) = x(4)
                tablo(i, 5) = ATIL_DEPO_METIN(x(15))
                tablo(i, 6) = x(21): tablo(i, 7) = Round(x(22), 2)
                tablo(i, 8) = ATIL_DURUM_ADI(CLng(x(10)))
                tablo(i, 9) = Y_IL_ADI(CStr(x(13)))
                tablo(i, 10) = ADM_ADET(CDbl(x(17))) & " " & x(7)
                tablo(i, 11) = x(18)
                tablo(i, 12) = x(14)
            Next i
            ATIL_TABLO_YAZ ws, a1 + 1, tablo
            r = a1 + j
            Y_TABLO_GOVDE ws.Range(ws.Cells(a1 + 1, 1), ws.Cells(r, 12))
            ws.Range(ws.Rows(a1 + 1), ws.Rows(r)).RowHeight = 20
            ws.Range(ws.Cells(a1 + 1, 1), ws.Cells(r, 1)).Font.bold = True
            ws.Range(ws.Cells(a1 + 1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(a1 + 1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
            ' Depo: tum depolar gorunur; uzun listede satir yuksekligi metne gore acilir.
            ws.Range(ws.Cells(a1 + 1, 5), ws.Cells(r, 5)).WrapText = True
            For i = a1 + 1 To r
                If Len(CStr(ws.Cells(i, 5).Value2)) > 24 Then
                    ws.Rows(i).AutoFit
                    If ws.Rows(i).RowHeight < 20 Then ws.Rows(i).RowHeight = 20
                End If
            Next i
            ws.Range(ws.Cells(a1 + 1, 9), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(a1 + 1, 4), ws.Cells(r, 4)).WrapText = False
            ws.Range(ws.Cells(a1 + 1, 12), ws.Cells(r, 12)).WrapText = False
            AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1 + 1, 6), ws.Cells(r, 6))
            ws.Range(ws.Cells(a1 + 1, 7), ws.Cells(r, 7)).NumberFormat = "#,##0.00 ""TL"""
            ATIL_DURUM_RENK ws.Range(ws.Cells(a1 + 1, 8), ws.Cells(r, 8))
            With ws.Range(ws.Cells(a1 + 1, 9), ws.Cells(r, 10))
                .Font.bold = True: .Font.Color = RGB(20, 103, 173)
            End With
            Y_DURUM_KF ws.Range(ws.Cells(a1 + 1, 11), ws.Cells(r, 11)), "YÜKSEK", RGB(220, 252, 231), RGB(21, 128, 61)
            Y_DURUM_KF ws.Range(ws.Cells(a1 + 1, 11), ws.Cells(r, 11)), "DÜÞÜK", RGB(241, 245, 249), RGB(100, 116, 139)
            ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12)).AutoFilter
            ' PDF cikti: basliktan son satira (filtrede gizlenen satirlar basilmaz).
            ws.Range("BX22").Value2 = ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12)).Address
            Set s = Y_DUGME(ws, "ATIL_PDF", "PDF ÇIKTI", w - 150, ws.Rows(12).Top + 1, 150, 28, RGB(15, 118, 110), "AKSORA_ATIL_TRANSFER_PDF")
        End If
    Else
        ' tum riskli malzemeler
        r = 12
        Y_BOLUM ws, r, "ATIL VE ATILA YAKLAÞAN MALZEMELER (" & n & ")", "Riskli miktar: stoðun 18 aydan eski giriþlere ait kýsmý. Sýra: durum (ATIL önce), sonra riskli deðer. Baþlýklardan filtrelenebilir.", 12
        ws.Rows(r).RowHeight = 30
        r = r + 1
        Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array("Ýl", "Sistem", "Malzeme", "Taným", "Kategori", _
            "Stok", "Riskli miktar", "Riskli deðer", "Riskli giriþ", "Gün sayacý", "Durum", "Transfer önerisi")
        ws.Rows(r).RowHeight = 30
        a1 = r
        If n = 0 Then
            r = r + 1
            ws.Cells(r, 1).Value2 = "Atýl ya da atýla yaklaþan malzeme yok."
            ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(100, 116, 139)
        Else
            ReDim anah(1 To n): ReDim idx(1 To n): ReDim satir(1 To n)
            i = 0
            For Each x In liste
                i = i + 1: idx(i) = i: satir(i) = x
                anah(i) = CStr(9 - x(10)) & Format$(999999999999# - Round(x(22), 0), "000000000000") & x(0) & x(3)
            Next x
            Y_SIRALA anah, idx, 1, n
            ReDim tablo(1 To n, 1 To 12)
            For i = 1 To n
                x = satir(idx(i))
                ATIL_SATIR tablo, i, x
                If Len(x(13)) > 0 Then tablo(i, 12) = ChrW(8594) & " " & Y_IL_ADI(CStr(x(13))) & "  (" & x(18) & ")"
            Next i
            ATIL_TABLO_YAZ ws, a1 + 1, tablo
            r = a1 + n
            ATIL_TABLO_BICIM ws, a1 + 1, r
            ws.Range(ws.Cells(a1 + 1, 11), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(a1 + 1, 12), ws.Cells(r, 12)).Font.Color = RGB(20, 103, 173)
            ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12)).AutoFilter
        End If
    End If
Son:
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Sub ATIL_SATIR(ByRef tablo() As Variant, ByVal r As Long, ByVal x As Variant)
    tablo(r, 1) = Y_IL_ADI(CStr(x(0)))
    tablo(r, 2) = x(1)
    tablo(r, 3) = CStr(x(3))
    tablo(r, 4) = x(4)
    tablo(r, 5) = x(5)
    tablo(r, 6) = x(6)
    tablo(r, 7) = x(21)
    tablo(r, 8) = Round(x(22), 2)
    tablo(r, 9) = CStr(x(9))
    tablo(r, 10) = x(11)
    tablo(r, 11) = ATIL_DURUM_ADI(CLng(x(10)))
End Sub

' Malzeme kodu ve tarih metin olarak yazilir (bastaki sifirlar / tarih donusumu olmaz).
Private Sub ATIL_TABLO_YAZ(ByVal ws As Worksheet, ByVal a1 As Long, ByRef tablo() As Variant)
    Dim m As Long
    m = UBound(tablo, 1)
    ws.Range(ws.Cells(a1, 3), ws.Cells(a1 + m - 1, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(a1, 9), ws.Cells(a1 + m - 1, 9)).NumberFormat = "@"
    ws.Cells(a1, 1).Resize(m, 12).Value2 = tablo
End Sub

Private Sub ATIL_TABLO_BICIM(ByVal ws As Worksheet, ByVal a1 As Long, ByVal r As Long)
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12))
    ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 20
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).Font.bold = True
    ws.Range(ws.Cells(a1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 9), ws.Cells(r, 10)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).WrapText = False
    ws.Range(ws.Cells(a1, 12), ws.Cells(r, 12)).WrapText = False
    AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6))
    AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7))
    ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).Font.bold = True
    ws.Range(ws.Cells(a1, 8), ws.Cells(r, 8)).NumberFormat = "#,##0.00 ""TL"""
    ATIL_DURUM_RENK ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11))
End Sub

' Bir sistemin bes uretim yeri. Her uretim yerinde: 1) son 24 ayin girisleri (kendi stogu), 2) stogu bu girislerden
' fazla olanlarda 24-60 ay onceki son giris, 3) son 6 ay sarfi (sistemin tum malzemeleri; transfer onerisi icin).
' Basarida bos, hatada "adim | hata" doner (diger sistem etkilenmez).
Private Function ATIL_SISTEM_OKU(ByVal ses As Object, ByVal sis As String, ByVal tum As Object, ByVal sonuc As Object, _
    ByVal bugun As Date, ByRef adet As Long, ByVal toplam As Long) As String
    Dim il As Variant, s As Object, k As Variant, uy As String, adim As String, birlesim As Object, kendi As Object, kalan As Object
    Dim mallar() As String, sat As Collection, x As Variant, katman As Object, gun As Object, eski As Object, sarf As Object, son6 As Object
    Dim g As Double, gk As String, toplamR As Double, y As Variant, tGiris() As String, tSarf() As String
    On Error GoTo Hata
    tGiris = ATIL_GIRIS_TUR(): tSarf = ATIL_SARF_TUR()
    Set birlesim = CreateObject("Scripting.Dictionary"): birlesim.CompareMode = vbTextCompare
    For Each il In ADM_ILLER()
        uy = AKSORA_IL_TESIS(CStr(il), sis)
        Set s = tum(CStr(il))(0)
        For Each k In s.keys
            If Left$(CStr(k), Len(uy) + 1) = uy & "|" Then birlesim(Mid$(CStr(k), Len(uy) + 2)) = True
        Next k
    Next il
    For Each il In ADM_ILLER()
        uy = AKSORA_IL_TESIS(CStr(il), sis)
        Set s = tum(CStr(il))(0)
        Set kendi = CreateObject("Scripting.Dictionary"): kendi.CompareMode = vbTextCompare
        For Each k In s.keys
            If Left$(CStr(k), Len(uy) + 1) = uy & "|" Then kendi(Mid$(CStr(k), Len(uy) + 2)) = True
        Next k
        Set katman = CreateObject("Scripting.Dictionary"): katman.CompareMode = vbTextCompare
        Set eski = CreateObject("Scripting.Dictionary"): eski.CompareMode = vbTextCompare
        Set sarf = CreateObject("Scripting.Dictionary"): sarf.CompareMode = vbTextCompare
        Set son6 = CreateObject("Scripting.Dictionary"): son6.CompareMode = vbTextCompare
        ' 1) son 24 ay girisleri
        If AKSORA_ADMIN_IPTAL Then Err.Raise vbObjectError + 2350, , "Kullanýcý durdurdu."
        If kendi.Count > 0 Then
            adim = sis & " " & uy & " (" & Y_IL_ADI(CStr(il)) & ") 1/3: giriþler 24 ay, " & kendi.Count & " malzeme"
            ATIL_ADIM 0.03 + 0.92 * adet / toplam, adim
            mallar = ATIL_DIZI(kendi)
            Set sat = ATIL_MB51(ses, uy, mallar, DateAdd("m", -24, bugun), bugun, sis & " " & uy & " 1/3", tGiris)
            For Each x In sat
                If ATIL_KAPSAM_MI(uy, CStr(x(4))) And Not ATIL_PARTI_DISI(CStr(x(5))) Then
                    g = ATIL_GIRIS_KATKI(CStr(x(3)), CDbl(x(2)))
                    If g <> 0 Then
                        If Not katman.Exists(CStr(x(0))) Then katman.Add CStr(x(0)), CreateObject("Scripting.Dictionary")
                        Set gun = katman(CStr(x(0)))
                        gk = Format$(CDate(x(1)), "yyyymmdd")
                        gun(gk) = ATIL_DS(gun, gk) + g
                    End If
                End If
            Next x
        End If
        adet = adet + 1
        ' 2) stogu 24 ay girislerinden fazla olanlar: 24-60 ay onceki son giris
        Set kalan = CreateObject("Scripting.Dictionary"): kalan.CompareMode = vbTextCompare
        For Each k In kendi.keys
            toplamR = 0
            If katman.Exists(CStr(k)) Then
                For Each y In katman(CStr(k)).items
                    toplamR = toplamR + CDbl(y)
                Next y
            End If
            If s(uy & "|" & CStr(k))(0) > toplamR + 0.0005 Then kalan(CStr(k)) = True
        Next k
        If kalan.Count > 0 Then
            If AKSORA_ADMIN_IPTAL Then Err.Raise vbObjectError + 2350, , "Kullanýcý durdurdu."
            adim = sis & " " & uy & " (" & Y_IL_ADI(CStr(il)) & ") 2/3: eski giriþler 2-5 yýl, " & kalan.Count & " malzeme"
            ATIL_ADIM 0.03 + 0.92 * adet / toplam, adim
            mallar = ATIL_DIZI(kalan)
            Set sat = ATIL_MB51(ses, uy, mallar, DateAdd("m", -60, bugun), DateAdd("m", -24, bugun) - 1, sis & " " & uy & " 2/3", tGiris)
            For Each x In sat
                If ATIL_KAPSAM_MI(uy, CStr(x(4))) And Not ATIL_PARTI_DISI(CStr(x(5))) Then
                    If ATIL_GIRIS_KATKI(CStr(x(3)), CDbl(x(2))) > 0 Then
                        If CDbl(x(1)) > ATIL_DS(eski, CStr(x(0))) Then eski(CStr(x(0))) = CDbl(x(1))
                    End If
                End If
            Next x
        End If
        adet = adet + 1
        ' 3) son 6 ay sarfi (yuklenici dahil; diger illerin talebi)
        If birlesim.Count > 0 Then
            If AKSORA_ADMIN_IPTAL Then Err.Raise vbObjectError + 2350, , "Kullanýcý durdurdu."
            adim = sis & " " & uy & " (" & Y_IL_ADI(CStr(il)) & ") 3/3: sarf 6 ay, " & birlesim.Count & " malzeme"
            ATIL_ADIM 0.03 + 0.92 * adet / toplam, adim
            mallar = ATIL_DIZI(birlesim)
            Set sat = ATIL_MB51(ses, uy, mallar, DateAdd("m", -6, bugun), bugun, sis & " " & uy & " 3/3", tSarf)
            For Each x In sat
                If ATIL_KAPSAM_MI(uy, CStr(x(4))) Then
                    sarf(CStr(x(0))) = ATIL_DS(sarf, CStr(x(0))) - CDbl(x(2))
                    If CDbl(x(2)) < 0 And CDbl(x(1)) > ATIL_DS(son6, CStr(x(0))) Then son6(CStr(x(0))) = CDbl(x(1))
                End If
            Next x
        End If
        adet = adet + 1
        For Each k In kendi.keys
            sonuc(uy & "|" & CStr(k)) = Array(ATIL_KATMAN_METIN(katman, CStr(k)), ATIL_DS(eski, CStr(k)), ATIL_DS(sarf, CStr(k)), ATIL_DS(son6, CStr(k)))
        Next k
        For Each k In sarf.keys
            If Not kendi.Exists(CStr(k)) And ATIL_DS(sarf, CStr(k)) > 0 Then _
                sonuc(uy & "|" & CStr(k)) = Array(vbNullString, 0#, ATIL_DS(sarf, CStr(k)), ATIL_DS(son6, CStr(k)))
        Next k
    Next il
    Exit Function
Hata:
    ATIL_SISTEM_OKU = adim & " | " & Err.Description
End Function
