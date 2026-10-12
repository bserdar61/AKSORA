Attribute VB_Name = "modAKSORA_IL_ILK_YUKLE"
Option Explicit

Private Sub AKIL_VERI_DOGRULA( _
    ByRef veri As Variant, _
    ByVal satir As Long, _
    ByVal sutun As Long, _
    ByVal il As String, _
    ByVal sistem As String, _
    ByVal tur As String)

    Dim i As Long
    Dim tesis As String, depo As String
    Dim birim As String

    tesis = AKSORA_IL_TESIS(il, sistem)

    If Not IsArray(veri) Then GoTo HATALI
    If satir <= 0 Or satir > 250000 Then GoTo HATALI
    If sutun <> 12 Then GoTo HATALI

    If LBound(veri, 1) <> 1 Then GoTo HATALI
    If UBound(veri, 1) <> satir Then GoTo HATALI
    If LBound(veri, 2) <> 1 Then GoTo HATALI
    If UBound(veri, 2) <> 12 Then GoTo HATALI

    For i = 1 To satir
        If IsError(veri(i, 1)) Then GoTo HATALI
        If Trim$(CStr(veri(i, 1))) <> tesis Then GoTo HATALI

        If IsError(veri(i, 2)) Then GoTo HATALI
        depo = Trim$(CStr(veri(i, 2)))

        If Not AKSORA_IL_DEPO_GECERLI(depo, tur, il) Then _
            GoTo HATALI

        If IsError(veri(i, 7)) Then GoTo HATALI
        If Len(Trim$(CStr(veri(i, 7)))) = 0 Then GoTo HATALI

        If IsError(veri(i, 10)) Then GoTo HATALI
        birim = Trim$(CStr(veri(i, 10)))
        If Len(birim) = 0 Then GoTo HATALI

        If IsError(veri(i, 9)) Then GoTo HATALI
        If IsError(veri(i, 11)) Then GoTo HATALI
        If IsError(veri(i, 12)) Then GoTo HATALI

        If Not IsNumeric(veri(i, 9)) Then GoTo HATALI
        If Not IsNumeric(veri(i, 11)) Then GoTo HATALI
        If Not IsNumeric(veri(i, 12)) Then GoTo HATALI
    Next i

    Exit Sub

HATALI:
    Err.Raise vbObjectError + 800, "AKIL_VERI_DOGRULA", _
              il & " / " & sistem & " / " & tur & _
              " verisi doðrulanamadý." & vbCrLf & _
              "Kontrol satýrý: " & i & vbCrLf & _
              "Mevcut stoklar deðiþtirilmedi."
End Sub

Public Function AKIL_VERI_BIRLESTIR( _
    ByRef birinci As Variant, _
    ByRef ikinci As Variant) As Variant

    Dim sonuc() As Variant
    Dim n1 As Long, n2 As Long
    Dim i As Long, j As Long

    n1 = UBound(birinci, 1)
    n2 = UBound(ikinci, 1)

    ReDim sonuc(1 To n1 + n2, 1 To 12)

    For i = 1 To n1
        For j = 1 To 12
            sonuc(i, j) = birinci(i, j)
        Next j
    Next i

    For i = 1 To n2
        For j = 1 To 12
            sonuc(n1 + i, j) = ikinci(i, j)
        Next j
    Next i

    AKIL_VERI_BIRLESTIR = sonuc
End Function

Public Function AKIL_SAP_VERI_AL( _
    ByVal adim As Long, _
    ByVal il As String) As Variant

    Dim veri As Variant
    Dim satir As Long, sutun As Long
    Dim sure As Double, hataMetni As String
    Dim basarili As Boolean
    Dim sistem As String, tur As String
    Dim olcumBas As Double, kontrolBas As Double, okumaSure As Double

    olcumBas = Timer
    Select Case adim
        Case 1
            sistem = "CDP": tur = "NORMAL"
            basarili = CDP_V2_30_BLOK_AL( _
                veri, satir, sutun, sure, hataMetni)

        Case 2
            sistem = "EPP": tur = "NORMAL"
            basarili = EPP_V2_EXPORT_AL_OZEL( _
                veri, satir, sutun, sure, hataMetni)

        Case 3
            sistem = "CDP": tur = "ILCE"
            basarili = ILCE_CDP_BLOK_AL( _
                veri, satir, sutun, sure, hataMetni)

        Case 4
            sistem = "EPP": tur = "ILCE"
            basarili = ILCE_EPP_BLOK_AL( _
                veri, satir, sutun, sure, hataMetni)

        Case 5
            sistem = "CDP": tur = "YUKLENICI"
            basarili = YUK_CDP_BLOK_AL( _
                veri, satir, sutun, sure, hataMetni)

        Case 6
            sistem = "EPP": tur = "YUKLENICI"
            basarili = YUK_EPP_BLOK_AL( _
                veri, satir, sutun, sure, hataMetni)

        Case Else
            Err.Raise vbObjectError + 801, _
                      "AKIL_SAP_VERI_AL", "Geçersiz iþlem adýmý."
    End Select

    okumaSure = Timer - olcumBas
    If okumaSure < 0 Then okumaSure = okumaSure + 86400#
    If Not basarili Then
        Err.Raise vbObjectError + 802, "AKIL_SAP_VERI_AL", _
                  sistem & " / " & tur & _
                  " stoklarý alýnamadý." & vbCrLf & hataMetni
    End If

    If AKSORA_AKTIF_IL() <> il Then
        Err.Raise vbObjectError + 803, "AKIL_SAP_VERI_AL", _
                  "Ýþlem sýrasýnda il seçimi deðiþti."
    End If

    kontrolBas = Timer
    AKIL_VERI_DOGRULA veri, satir, sutun, il, sistem, tur
    olcumBas = Timer - kontrolBas
    If olcumBas < 0 Then olcumBas = olcumBas + 86400#
    AKSORA_SAP_OLCUM_KAYDET il, adim, satir, okumaSure, olcumBas
    AKIL_SAP_VERI_AL = veri
End Function

Public Sub AKSORA_IL_SAP_ON_KONTROL()
    Dim il As String, yuklu As String
    Dim veri As Variant
    Dim i As Long, j As Long
    Dim toplam As Double
    Dim sonuc As String, hataMetni As String
    Static calisiyor As Boolean

    If calisiyor Then Exit Sub
    On Error GoTo Hata

    il = AKSORA_AKTIF_IL()
    yuklu = AKSORA_CANLI_STOK_IL()

    If MsgBox( _
        "SAP veri kontrolü yapýlacak." & vbCrLf & _
        "Kontrol edilecek il: " & il & vbCrLf & _
        "Mevcut stoklarýn ili: " & yuklu & vbCrLf & vbCrLf & _
        "Stok, rapor, arþiv ve geçmiþ sayfalarýna veri yazýlmaz." & _
        vbCrLf & "Devam edilsin mi?", _
        vbYesNo + vbQuestion, "AKSORA") <> vbYes Then Exit Sub

    calisiyor = True

    For i = 1 To 6
        veri = AKIL_SAP_VERI_AL(i, il)
        toplam = 0

        For j = 1 To UBound(veri, 1)
            toplam = toplam + CDbl(veri(j, 11))
        Next j

        sonuc = sonuc & "Adým " & i & ": " & _
                UBound(veri, 1) & " satýr / " & _
                AKSORA_SAYI_METNI(toplam, "#,##0.00") & " TL" & vbCrLf

        veri = Empty
    Next i

    calisiyor = False

    MsgBox il & " SAP veri kontrolü tamamlandý." & vbCrLf & vbCrLf & _
           "1–2: Normal CDP / EPP" & vbCrLf & _
           "3–4: Ýlçe CDP / EPP" & vbCrLf & _
           "5–6: Yüklenici CDP / EPP" & vbCrLf & vbCrLf & _
           sonuc & vbCrLf & _
           "Mevcut Excel verileri deðiþtirilmedi.", _
           vbInformation, "AKSORA"
    Exit Sub

Hata:
    hataMetni = Err.Description
    calisiyor = False

    MsgBox "SAP veri kontrolü durduruldu." & vbCrLf & _
           hataMetni & vbCrLf & vbCrLf & _
           "Mevcut Excel verileri deðiþtirilmedi.", _
           vbExclamation, "AKSORA"
End Sub
Public Sub AKSORA_IL_ILK_ARSIV_YUKLE()
    Dim il As String, yuklu As String
    Dim tipler As Variant
    Dim ar(1 To 5) As Worksheet
    Dim eski(1 To 5) As Variant
    Dim ham(1 To 6) As Variant
    Dim veri(1 To 5) As Variant
    Dim tarihler(1 To 6) As Date
    Dim i As Long, j As Long, n As Long
    Dim toplam As Double, ad As String
    Dim ekran As Boolean, olay As Boolean
    Dim ayarDegisti As Boolean, yazmaBasladi As Boolean
    Dim hataMetni As String, geriHata As String
    Static calisiyor As Boolean

    If calisiyor Then Exit Sub
    On Error GoTo Hata

    il = AKSORA_AKTIF_IL()
    yuklu = AKSORA_CANLI_STOK_IL()

    If il = yuklu Then
        MsgBox "Bu ilin stoklarý zaten yüklü: " & il, _
               vbInformation, "AKSORA"
        Exit Sub
    End If

    tipler = Array("CDP", "EPP", "TUM", "ILC", "YUK")

    For i = 1 To 5
        ad = "_AKLIVE_" & il & "_" & tipler(i - 1)
        Set ar(i) = ThisWorkbook.Worksheets(ad)

        If ar(i).ProtectContents Then
            Err.Raise vbObjectError + 810, , _
                      "Hedef arþiv korumalý: " & ad
        End If

        If CStr(ar(i).Range("A1").Value2) <> _
           "AKSORA_CANLI_ARSIV_V1" Or _
           CStr(ar(i).Range("B1").Value2) <> il Or _
           CStr(ar(i).Range("C1").Value2) <> _
           CStr(tipler(i - 1)) Then

            Err.Raise vbObjectError + 811, , _
                      "Hedef arþiv kimliði uyuþmuyor: " & ad
        End If

        For j = 1 To 12
            If ar(i).Cells(ar(i).Rows.Count, j).End(xlUp).Row >= 7 Then
                Err.Raise vbObjectError + 812, , _
                          "Hedef arþiv boþ deðil: " & ad
            End If
        Next j

        If application.CountA(ar(i).Range("A3:A5")) > 0 Then
            Err.Raise vbObjectError + 813, , _
                      "Hedef arþivde önceki kayýt bilgileri var: " & ad
        End If

        eski(i) = ar(i).Range("A1:L6").Value2
    Next i

    If MsgBox( _
        il & " için ilk SAP stoklarý alýnacak." & vbCrLf & _
        "Altý veri grubu doðrulanýnca beþ il arþivine yazýlacak." & _
        vbCrLf & vbCrLf & _
        yuklu & " stoklarý ve raporlarý deðiþtirilmeyecek." & _
        vbCrLf & "Devam edilsin mi?", _
        vbYesNo + vbQuestion, "AKSORA") <> vbYes Then Exit Sub

    calisiyor = True

    For i = 1 To 6
        ham(i) = AKIL_SAP_VERI_AL(i, il)
        tarihler(i) = Now
    Next i

    veri(1) = ham(1)
    veri(2) = ham(2)
    veri(3) = AKIL_VERI_BIRLESTIR(ham(1), ham(2))
    veri(4) = AKIL_VERI_BIRLESTIR(ham(3), ham(4))
    veri(5) = AKIL_VERI_BIRLESTIR(ham(5), ham(6))

    If AKSORA_AKTIF_IL() <> il Or _
       AKSORA_CANLI_STOK_IL() <> yuklu Then

        Err.Raise vbObjectError + 814, , _
                  "Ýþlem sýrasýnda il veya mevcut stoklar deðiþti."
    End If

    ekran = application.ScreenUpdating
    olay = application.EnableEvents
    ayarDegisti = True

    application.ScreenUpdating = False
    application.EnableEvents = False
    yazmaBasladi = True

    For i = 1 To 5
        n = UBound(veri(i), 1)
        toplam = 0

        For j = 1 To n
            toplam = toplam + CDbl(veri(i)(j, 11))
        Next j

        ar(i).Range("A7").Resize(n, 12).Value2 = veri(i)

        Select Case i
            Case 1: ar(i).Range("A3").Value = tarihler(1)
            Case 2: ar(i).Range("A3").Value = tarihler(2)
            Case 3: ar(i).Range("A3").Value = tarihler(2)
            Case 4: ar(i).Range("A3").Value = tarihler(4)
            Case 5: ar(i).Range("A3").Value = tarihler(6)
        End Select

        ar(i).Range("A4").Value = n
        ar(i).Range("A5").Value = toplam
        ar(i).Range("D1").Value = Now
    Next i

    yazmaBasladi = False
    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    ayarDegisti = False
    calisiyor = False

    MsgBox il & " için beþ ilk stok arþivi hazýr." & vbCrLf & _
           yuklu & " stoklarý ve raporlarý korundu." & vbCrLf & _
           "Ýl geçiþi henüz yapýlmadý.", _
           vbInformation, "AKSORA"
    Exit Sub

Hata:
    hataMetni = Err.Description
    On Error Resume Next

    If yazmaBasladi Then
        For i = 1 To 5
            Err.Clear
            n = UBound(veri(i), 1)
            ar(i).Range("A7").Resize(n, 12).ClearContents
            ar(i).Range("A1:L6").Value2 = eski(i)

            If Err.Number <> 0 Then
                geriHata = geriHata & vbCrLf & _
                           "Arþiv geri alma hatasý: " & ar(i).Name
            End If
        Next i
    End If

    If ayarDegisti Then
        application.EnableEvents = olay
        application.ScreenUpdating = ekran
    End If

    calisiyor = False
    On Error GoTo 0

    MsgBox "Ýlk arþiv yüklemesi durduruldu." & vbCrLf & _
           hataMetni & geriHata & vbCrLf & vbCrLf & _
           "Canlý stok ve rapor sayfalarýna veri yazýlmadý.", _
           vbExclamation, "AKSORA"
End Sub

