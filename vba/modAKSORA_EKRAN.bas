Attribute VB_Name = "modAKSORA_EKRAN"
Option Explicit

Private mSigdirmaSuruyor As Boolean
Private mOtomatikKapali As Boolean
Private mSonHata As String
Private mZoomHafiza As Object

' Yalnizca aktif sayfanin zoom'u degisir. Veri ve SAP koduna dokunulmaz.
Public Function AKSORA_PENCERE_KUCUK(Optional ByVal wn As Window) As Boolean
    ' Excel veya pencere simge durumundayken pencere ozellikleri ayarlanamaz (hata 1004).
    On Error GoTo Kucuk
    If Not application.Visible Then GoTo Kucuk
    If application.WindowState = xlMinimized Then GoTo Kucuk
    If Not wn Is Nothing Then
        If Not wn.Visible Then GoTo Kucuk
        If wn.WindowState = xlMinimized Then GoTo Kucuk
    End If
    Exit Function
Kucuk:
    AKSORA_PENCERE_KUCUK = True
End Function

Public Sub AKSORA_HATA_GOSTER(ByVal islem As String, ByVal no As Long, ByVal metin As String)
    On Error Resume Next
    application.EnableEvents = True
    application.ScreenUpdating = True
    application.Cursor = xlDefault
    application.StatusBar = False
    If Len(metin) = 0 Then metin = "Beklenmeyen hata (" & no & ")."
    MsgBox islem & " tamamlanamadý." & vbCrLf & vbCrLf & metin, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_EKRAN_SIGDIR(ByVal ws As Worksheet)
    Dim wn As Window, alan As Range, eskiSecim As Range
    Dim adres As String, ad As String
    Dim olay As Boolean, ekran As Boolean, eskiZoom As Variant
    Dim eskiSecilebilir As XlEnableSelection, satir As Long, sutun As Long
    Dim durumAlindi As Boolean, hataMetni As String
    Dim kapakKilidi As Boolean, yukseklik As Double, yeniZoom As Long, anahtar As String
    If mSigdirmaSuruyor Or mOtomatikKapali Then Exit Sub
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    If ws.Visible <> xlSheetVisible Then Exit Sub
    ad = AKFIX_NORMAL(ws.Name)
    Select Case ad
        Case "KAPAK": adres = "A1:AE39"
        Case "ANA PANEL": adres = "A1:AB30"
        Case "MALZEME_ARA": adres = "A1:N31"
        Case "STOK_KARSILASTIRMA": adres = "A1:P29"
        Case "CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK"
            ' Tek satir secimi: binlerce kayit zoom'u kucultmez.
            ' O sutunu mevcut ANA PANEL butonunun sag kenarini da kapsar.
            If AKSORA_TASARIM_AKTIF() Then
                adres = "A1:L1"
            Else
                adres = "A1:V1"
            End If
        Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR"
            ' Grafikler ve donus butonu dahil genislige sigdir.
            adres = "A1:T1"
        Case Else: Exit Sub
    End Select
    mSonHata = vbNullString
    On Error GoTo Hata
    Set wn = ActiveWindow
    If wn Is Nothing Then Exit Sub
    If AKSORA_PENCERE_KUCUK(wn) Then Exit Sub
    olay = application.EnableEvents
    ekran = application.ScreenUpdating
    eskiZoom = wn.Zoom
    eskiSecilebilir = ws.EnableSelection
    satir = wn.ScrollRow: sutun = wn.ScrollColumn
    Set eskiSecim = wn.RangeSelection
    durumAlindi = True
    mSigdirmaSuruyor = True
    application.EnableEvents = False
    application.ScreenUpdating = False
    ' Korumayi kaldirmadan yalnizca gecici secim kisitini gevset.
    If ad = "KAPAK" Then
        kapakKilidi = ws.ProtectContents Or ws.ProtectDrawingObjects
        If kapakKilidi Then ws.Unprotect Password:=""
    End If
    ws.EnableSelection = xlNoRestrictions
    Select Case ad
        Case "CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR"
            AKSORA_RAPOR_AMBLEM_HIZALA ws.Name
    End Select
    AKSORA_TASARIM_SAYFA ws
    ' Ayni pencere boyutunda sayfanin zoom'u bir kez hesaplanir; sonraki gecislerde
    ' yeniden secim/zoom yapilmaz (gecis aninda ve titremesiz olur).
    If ad <> "KAPAK" Then
        anahtar = ws.Name & "|" & Format$(wn.UsableWidth, "0") & "|" & Format$(wn.UsableHeight, "0")
        If mZoomHafiza Is Nothing Then Set mZoomHafiza = CreateObject("Scripting.Dictionary")
        If mZoomHafiza.Exists(anahtar) Then
            If wn.Zoom <> mZoomHafiza(anahtar) Then wn.Zoom = mZoomHafiza(anahtar)
            GoTo Cikis
        End If
    End If
    Set alan = ws.Range(adres)
    alan.Select
    wn.Zoom = True
    If ad = "KAPAK" Then
        ' Genislik sigsa bile il secim panelinin yuksekligini kontrol et.
        yukseklik = wn.VisibleRange.Height - wn.VisibleRange.Rows(wn.VisibleRange.Rows.Count).Height
        If yukseklik > 6 Then yukseklik = yukseklik - 6
        If yukseklik > 0 And yukseklik < alan.Height Then
            yeniZoom = Int(CDbl(wn.Zoom) * yukseklik / alan.Height)
            If yeniZoom < 10 Then yeniZoom = 10
            wn.Zoom = yeniZoom
        End If
    End If
    ' Son sutun ve buton kenarlari icin kucuk bir pay.
    If CLng(wn.Zoom) > 12 Then wn.Zoom = CLng(wn.Zoom) - 2
    If Len(anahtar) > 0 Then mZoomHafiza(anahtar) = CLng(wn.Zoom)
    If ad = "KAPAK" Then AKSORA_KAPAK_GENISLIK_UYUM ws, wn
Cikis:
    On Error Resume Next
    If durumAlindi Then
        If Len(hataMetni) > 0 Then wn.Zoom = eskiZoom
        If Not eskiSecim Is Nothing Then eskiSecim.Select
        If ad = "KAPAK" Then
            wn.ScrollRow = 1: wn.ScrollColumn = 1
        Else
            wn.ScrollRow = satir: wn.ScrollColumn = sutun
        End If
        Select Case ad
            Case "CDP_STOK", "EPP_STOK", "ILCE_STOK", "YUKLENICI_STOK"
                If Len(hataMetni) = 0 Then ws.Range("H2").Select
        End Select
        If kapakKilidi Then
            ws.Protect Password:="", DrawingObjects:=True, Contents:=True, _
                Scenarios:=True, UserInterfaceOnly:=True, AllowFormattingCells:=False
        End If
        ws.EnableSelection = eskiSecilebilir
        application.EnableEvents = olay
        application.ScreenUpdating = ekran
    End If
    mSigdirmaSuruyor = False
    If Len(hataMetni) = 0 Then AKSORA_TASARIM_PENCERE ws
    On Error GoTo 0
    If Len(hataMetni) > 0 Then
        mSonHata = ws.Name & ": " & hataMetni
        Debug.Print "AKSORA ekran uyumu | " & mSonHata
    End If
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_EKRANA_SIGDIR()
    Dim ws As Worksheet
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    mOtomatikKapali = False
    mSonHata = vbNullString
    Set ws = ActiveSheet
    AKSORA_EKRAN_SIGDIR ws
    If Len(mSonHata) > 0 Then
        MsgBox "Ekran uyumu uygulanamadi; onceki zoom korundu." & vbCrLf & mSonHata, vbExclamation, "AKSORA"
    End If
End Sub

Public Sub AKSORA_OTOMATIK_ZOOM_KAPAT()
    mOtomatikKapali = True
    MsgBox "Otomatik zoom bu oturum icin kapatildi. Excel'in zoom ayarini kullanabilirsiniz.", vbInformation, "AKSORA"
End Sub

Public Sub AKSORA_RAPOR_GORSEL_ONAR()
    Dim ad As Variant, olay As Boolean, ekran As Boolean, hataMetni As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    olay = application.EnableEvents
    ekran = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False
    application.ScreenUpdating = False
    For Each ad In Array("CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ÝLÇE_RAPORLAR", "YÜKLENÝCÝ_RAPOR")
        AKSORA_RAPOR_AMBLEM_HIZALA CStr(ad)
    Next ad
Cikis:
    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    If Len(hataMetni) > 0 Then
        MsgBox "Gorsel onarim tamamlanamadi: " & hataMetni, vbExclamation, "AKSORA"
    Else
        AKSORA_EKRANA_SIGDIR
        MsgBox "Bes raporun amblemi duzenlendi. Kapak ve rapor sayfalarini kontrol edin.", vbInformation, "AKSORA"
    End If
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Cikis
End Sub

Public Sub AKSORA_KAPAK_SIGDIR_TEST()
    Dim ws As Worksheet, onceki As Variant
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If ActiveSheet.Name <> "KAPAK" Then
        MsgBox "Once KAPAK sayfasini acin.", vbInformation, "AKSORA"
        Exit Sub
    End If
    Set ws = ActiveSheet
    onceki = ActiveWindow.Zoom
    AKSORA_EKRANA_SIGDIR
    If Len(mSonHata) = 0 Then
        MsgBox "Kapak zoom: " & onceki & "% -> " & ActiveWindow.Zoom & "%" & vbCrLf & _
               "Il butonlari ve panele giris tam gorunuyor mu kontrol edin.", vbInformation, "AKSORA"
    End If
End Sub


