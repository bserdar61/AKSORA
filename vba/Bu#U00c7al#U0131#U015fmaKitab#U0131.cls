VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "BuÇalýþmaKitabý"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()

    Dim adim As String
    Dim bas As Double
    Dim Hata As String

    bas = Timer

    On Error GoTo Hata

    ' Yukleme ekrani ilk is olarak gelir.
    AKSORA_V13_ACILIS = True
    application.EnableEvents = False
    application.ScreenUpdating = False

    adim = "Yükleme ekraný"

    SPLASH_GOSTER

    AKSORA_SAYI_AYIRACLARI
    SPLASH_DURUM 0.12, "Baþlangýç ayarlarý hazýrlanýyor..."

    adim = "Makro korumalarý"

    AKSORA_HAFIF_KORUMALARI_YENILE
    AKSORA_SAYI_BICIMLERINI_YENILE
    AKSORA_MUSTERI_ARAMA_KPI_YENILE

    adim = "Ýl butonlarý"

    AKSORA_V13_BUTONLARI_BAGLA

    SPLASH_DURUM 0.4, "Ýl seçimi hazýrlanýyor..."

    adim = "Kapak"

    AKSORA_KAPAK_KILITLE

    ThisWorkbook.Worksheets("KAPAK").Activate

  AKSORA_KAPAK_KORU
    AKSORA_V8_KAPAK_GORUNUM

    SPLASH_DURUM 0.7, _
        "Güncelleme tarihleri kontrol ediliyor..."

    adim = "Kapak tarihi"

    AKSORA_V10_TARIH_GOSTER

    adim = "Uygulama görünümü"

    UYGULAMA_MODU_AC

    AKSORA_V8_KAPAK_GORUNUM

    adim = "Fiyat listesi"
    AKSORA_MUSTERI_FIYAT_KONTROL

    adim = "Kayýt bakýmý"
    SPLASH_DURUM 0.85, "Kayýtlar kontrol ediliyor..."
    AKSORA_ARSIV_BAKIM

    ' AKSORA_BIRIM_HOOK_OPEN
    AKSORA_BIRIM_ACILIS
    SPLASH_DURUM 1, "AKSORA hazýr"


Cikis:

    On Error Resume Next

    AKSORA_V13_ACILIS = False

    application.EnableEvents = True
    application.ScreenUpdating = True

    SPLASH_KAPAT

    AKSORA_V13_SURE_KAYDET _
        "Açýlýþ", _
        bas

    On Error GoTo 0


    If Len(Hata) > 0 Then

        MsgBox _
            Hata, _
            vbExclamation, _
            "AKSORA"

    End If

    Exit Sub


Hata:

    Hata = _
        "Açýlýþ tamamlanamadý: " & _
        adim & _
        vbCrLf & _
        Err.Description

    Resume Cikis

End Sub


Private Sub Workbook_Activate()

    If AKSORA_V13_ACILIS _
       Or AKSORA_V13_GECIS Then

        Exit Sub

    End If


    On Error Resume Next

    If AKSORA_PENCERE_KUCUK(ThisWorkbook.Windows(1)) Then Exit Sub


    UYGULAMA_MODU_AC
    ' AKSORA_BIRIM_HOOK_ACTIVATE
    If AKSORA_BIRIM_AKTIF_GORUNUM() Then Exit Sub


    If Not AKSORA_GIRIS_ACIK Then

        AKSORA_KAPAK_KORU

        AKSORA_V8_KAPAK_GORUNUM

    End If

End Sub


Private Sub Workbook_Deactivate()

    On Error Resume Next

    Sayfa12.STOK_KARS_FILTRE_TAKIP_DURDUR

    UYGULAMA_MODU_KAPAT

End Sub


'Yatirim teminat / hakedis girisleri aninda kaydedilir.
Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    If AKSORA_V13_ACILIS Then Exit Sub
    If Left$(Sh.Name, 2) <> "Y_" And Left$(Sh.Name, 2) <> "I_" Then Exit Sub
    On Error Resume Next
    If Left$(Sh.Name, 2) = "I_" Then AKSORA_ILCE_DEGISIM Sh, Target Else AKSORA_YATIRIM_DEGISIM Sh, Target
End Sub

' Dosya her acilista normal gorunumde gelir (sayfa duzeni / %10 yakinlastirma kalmaz).
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Dim wn As Window
    On Error Resume Next
    For Each wn In ThisWorkbook.Windows
        If wn.View <> xlNormalView Then wn.View = xlNormalView
        If wn.Zoom < 40 Then wn.Zoom = 100
    Next wn
End Sub

Private Sub Workbook_BeforeClose( _
    Cancel As Boolean)

    On Error Resume Next

    Sayfa12.STOK_KARS_FILTRE_TAKIP_DURDUR

    UYGULAMA_MODU_KAPAT

End Sub


Private Sub Workbook_SheetActivate( _
    ByVal sh As Object)

    Dim ev As Boolean
    Dim sc As Boolean
    Dim Hata As String


    ' AKSORA_BIRIM_HOOK_SHEET
    If AKSORA_BIRIM_OLAY(sh) Then Exit Sub
    If AKSORA_V13_ACILIS _
       Or AKSORA_V13_GECIS Then

        Exit Sub

    End If


    If AKSORA_V11_MESGUL Then
        Exit Sub
    End If


    If Not TypeOf sh Is Worksheet Then
        Exit Sub
    End If


    ev = application.EnableEvents
    sc = application.ScreenUpdating


    On Error GoTo Hata


    application.EnableEvents = False


    If AKSORA_GIRIS_ACIK Then

        AKSORA_V13_GORUNUM sh

    Else

        ThisWorkbook.Worksheets("KAPAK").Activate

        AKSORA_KAPAK_KORU

        AKSORA_V8_KAPAK_GORUNUM

    End If


Cikis:

    application.EnableEvents = ev

    application.ScreenUpdating = sc


    If Len(Hata) > 0 Then

        MsgBox _
            Hata, _
            vbExclamation, _
            "AKSORA"

    End If

    Exit Sub


Hata:

    Hata = _
        "Sayfa görünümü: " & _
        Err.Description

    Resume Cikis

End Sub


Private Sub Workbook_WindowResize(ByVal wn As Window)
    If AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Or AKSORA_V11_MESGUL Then Exit Sub
    If AKSORA_PENCERE_KUCUK(wn) Then Exit Sub
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    ' AKSORA_BIRIM_HOOK_RESIZE
    If AKSORA_BIRIM_AKTIF_GORUNUM() Then Exit Sub
    If Not AKSORA_GIRIS_ACIK And ActiveSheet.Name <> "KAPAK" Then Exit Sub
    AKSORA_EKRAN_SIGDIR ActiveSheet
End Sub

Private Sub Workbook_WindowActivate(ByVal wn As Window)
    If AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Or AKSORA_V11_MESGUL Then Exit Sub
    If AKSORA_PENCERE_KUCUK(wn) Then Exit Sub
    On Error Resume Next
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not TypeOf ActiveSheet Is Worksheet Then Exit Sub
    ' AKSORA_BIRIM_HOOK_WINDOW
    If AKSORA_BIRIM_AKTIF_GORUNUM() Then Exit Sub
    If Not AKSORA_GIRIS_ACIK And ActiveSheet.Name <> "KAPAK" Then Exit Sub
    AKSORA_EKRAN_SIGDIR ActiveSheet
End Sub






Private Sub Workbook_SheetCalculate(ByVal sh As Object)
    If Not TypeOf sh Is Worksheet Then Exit Sub
    If sh.Name <> "M_MALZEME_ARA" And sh.Name <> "Y_MALZEME_ARA" And sh.Name <> "I_MALZEME_ARA" Then Exit Sub
    If Not application.EnableEvents Then Exit Sub
    On Error GoTo Hata
    AKSORA_MUSTERI_ARAMA_KPI_HESAPLA sh
    Exit Sub
Hata:
    MsgBox "Arama özeti güncellenemedi: " & Err.Description, vbExclamation, "AKSORA"
End Sub

