Attribute VB_Name = "modAKSORA_KORUMA"
Option Explicit

' Sadece belirtilen sayfalardaki resimleri kilitler.
' Sayfa korumasi, resimlerin elle tasinmasini engeller.
' NOT: Koruma altinda sekil/grafik ekleyen makrolar ek uyarlama isteyebilir.
Public Sub AKSORA_RESIMLERI_GERCEKTEN_KORU()
    Dim isimler As Variant, ad As Variant
    Dim ws As Worksheet, sh As Shape
    Dim kilit As Long, sonuc As String
    Dim eskiEkran As Boolean

    isimler = Array("ANA PANEL", "CDP_STOK", "CDP_RAPOR", _
                   "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK")
    eskiEkran = application.ScreenUpdating
    application.ScreenUpdating = False
    On Error GoTo Hata

    For Each ad In isimler
        Set ws = AKSORA_KorumaSayfaBul(CStr(ad))
        If ws Is Nothing Then
            sonuc = sonuc & vbCrLf & CStr(ad) & ": bulunamadi"
        ElseIf ws.ProtectContents Or ws.ProtectDrawingObjects Then
            sonuc = sonuc & vbCrLf & ws.Name & ": zaten korumali; degistirilmedi"
        Else
            kilit = 0
            For Each sh In ws.Shapes
                If sh.Type = msoPicture Or sh.Type = msoLinkedPicture Then
                    sh.Locked = True
                    sh.Placement = xlFreeFloating
                    kilit = kilit + 1
                Else
                    ' Mevcut butonlar tiklanabilir kalir.
                    ' Grafikleri VBA ile degistiren kodlarin yine de test edilmesi gerekir.
                    On Error Resume Next
                    sh.Locked = False
                    Err.Clear
                    On Error GoTo Hata
                End If
            Next sh
            ws.Protect Password:="", DrawingObjects:=True, Contents:=False, _
                       Scenarios:=False, UserInterfaceOnly:=True, _
                       AllowFiltering:=True
            sonuc = sonuc & vbCrLf & ws.Name & ": " & kilit & " resim korundu"
        End If
        Set ws = Nothing
    Next ad

    application.ScreenUpdating = eskiEkran
    MsgBox "Resim koruma islemi:" & sonuc & vbCrLf & vbCrLf & _
           "SAP guncellemesi ve grafik yenilemeyi bir kez test edin.", _
           vbInformation, "AKSORA"
    Exit Sub
Hata:
    application.ScreenUpdating = eskiEkran
    MsgBox "Islem durdu: " & Err.Description & vbCrLf & _
           "Son islenen sayfalari kontrol edin.", vbExclamation, "AKSORA"
End Sub

Private Function AKSORA_KorumaSayfaBul(ByVal ad As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If AKSORA_KorumaAdNorm(ws.Name) = AKSORA_KorumaAdNorm(ad) Then
            Set AKSORA_KorumaSayfaBul = ws
            Exit Function
        End If
    Next ws
End Function

Private Function AKSORA_KorumaAdNorm(ByVal s As String) As String
    s = Replace(s, ChrW(304), "I")
    s = Replace(s, ChrW(305), "I")
    s = Replace(s, "i", "I")
    s = Replace(s, "ü", "U"): s = Replace(s, "Ü", "U")
    s = Replace(s, "þ", "S"): s = Replace(s, "Þ", "S")
    s = Replace(s, "ð", "G"): s = Replace(s, "Ð", "G")
    s = Replace(s, "ö", "O"): s = Replace(s, "Ö", "O")
    s = Replace(s, "ç", "C"): s = Replace(s, "Ç", "C")
    AKSORA_KorumaAdNorm = UCase$(s)
End Function


Public Sub AKSORA_KORUMA_KUR()
    On Error GoTo Hata
    AKSORA_KAPAK_KORU
    MsgBox "AKSORA kapak tiklama korumasi aktif.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    MsgBox "Koruma hatasi (" & Err.Number & "): " & Err.Description, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_KAPAK_KORU()
    Dim ws As Worksheet, s As Shape, i As Long
    Dim oldEvents As Boolean, oldScreen As Boolean
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets("KAPAK")
    oldEvents = application.EnableEvents
    oldScreen = application.ScreenUpdating
    application.EnableEvents = False
    application.ScreenUpdating = False

    If ws.ProtectContents Or ws.ProtectDrawingObjects Then ws.Unprotect

    ' Gorsel ve dekoratif nesneler kilitli, islevsel butonlar makro calistirir.
    For i = 1 To ws.Shapes.Count
        Set s = ws.Shapes(i)
        s.Locked = True
    Next i

    ' UserInterfaceOnly: VBA'nin V8 tasarim nesnelerini duzenlemesine izin verir.
    ws.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, _
               UserInterfaceOnly:=True, AllowFormattingCells:=False
    ws.EnableSelection = xlNoSelection
    ws.ScrollArea = "A1:AY49"

    If ws Is ActiveSheet Then
        ActiveWindow.DisplayGridlines = False
        ActiveWindow.DisplayHeadings = False
        ActiveWindow.DisplayWorkbookTabs = False
        ActiveWindow.DisplayHorizontalScrollBar = False
        ActiveWindow.DisplayVerticalScrollBar = False
        ActiveWindow.ScrollRow = 1
        ActiveWindow.ScrollColumn = 1
    End If

Cikis:
    application.EnableEvents = oldEvents
    application.ScreenUpdating = oldScreen
    Exit Sub
Hata:
    application.EnableEvents = True
    application.ScreenUpdating = True
    MsgBox "Kapak koruma hatasi (" & Err.Number & "): " & Err.Description, vbExclamation, "AKSORA"
End Sub

' Gerektiginde tasarim duzenlemek icin.
Public Sub AKSORA_KAPAK_KORUMAYI_KALDIR()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("KAPAK")
    ws.Unprotect
    ws.EnableSelection = xlNoRestrictions
    MsgBox "Kapak korumasi kaldirildi.", vbInformation, "AKSORA"
End Sub


