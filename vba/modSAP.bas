Attribute VB_Name = "modSAP"
Option Explicit

Private mSAPDetay As Object

'==========================================================
' AKSORA
' modSAP
'
' MERKEZÝ SAP BAÐLANTI MOTORU
'
' Sistemler:
' CDP -> SystemName = CDP
' EPP -> SystemName = EPP
'
' Amaç:
' - SAP GUI açýk mý kontrol etmek
' - Tüm açýk SAP baðlantýlarýný taramak
' - Ýstenen sistemi (CDP / EPP) özellikle bulmak
' - Yanlýþ SAP oturumunun kullanýlmasýný engellemek
' - Kullanýcý bilgisi göstermemek
'
' Bu modül þu aþamada:
' - MB52 çalýþtýrmaz
' - SAP ekranýna veri yazmaz
' - Excel stok verisine dokunmaz
'==========================================================

'==========================================================
' HEDEF SAP SESSION AL
'
' Örnek:
'
' Set session = SAP_SESSION_AL("CDP", hata)
' Set session = SAP_SESSION_AL("EPP", hata)
'
' HedefSistem boþ býrakýlýrsa kullanýlabilir ilk
' SAP oturumunu döndürür.
'==========================================================
Private mCDPSession As Object
Private mEPPSession As Object

Public Sub SAP_OTURUM_ONBELLEK_TEMIZLE()
    Set mCDPSession = Nothing
    Set mEPPSession = Nothing
End Sub

Private Function SAP_ONBELLEK_GECERLI( _
    ByVal Oturum As Object, _
    ByVal BeklenenSistem As String) As Boolean

    Dim GercekSistem As String
    On Error GoTo Cikis
    If Oturum Is Nothing Then Exit Function
    GercekSistem = UCase$(Trim$(CStr(Oturum.Info.SystemName)))
    If GercekSistem <> BeklenenSistem Then Exit Function
    SAP_ONBELLEK_GECERLI = True
Cikis:
End Function

Public Function SAP_SESSION_AL( _
Optional ByVal HedefSistem As String = "", _
Optional ByRef hataMesaji As String = "") As Object

Dim SapGuiAuto As Object
Dim SAPApp As Object
Dim SAPConnection As Object
Dim SAPSession As Object

Dim BaglantiSayisi As Long
Dim OturumSayisi As Long

Dim SistemAdi As String

Dim i As Long
Dim j As Long

On Error GoTo Hata

hataMesaji = vbNullString
HedefSistem = UCase$(Trim$(HedefSistem))

' Ayný SAP oturumunu yeniden kullan; kopmuþsa normal taramaya dön.
Select Case HedefSistem
    Case "CDP"
        If SAP_ONBELLEK_GECERLI(mCDPSession, "CDP") Then
            Set SAP_SESSION_AL = mCDPSession
            Exit Function
        End If
        Set mCDPSession = Nothing
    Case "EPP"
        If SAP_ONBELLEK_GECERLI(mEPPSession, "EPP") Then
            Set SAP_SESSION_AL = mEPPSession
            Exit Function
        End If
        Set mEPPSession = Nothing
End Select

'======================================================
' 1. SAP GUI NESNESÝNÝ AL
'======================================================

Set SapGuiAuto = Nothing

On Error Resume Next

Err.Clear
Set SapGuiAuto = GetObject("SAPGUI")

If Err.Number <> 0 Then
    Err.Clear
    Set SapGuiAuto = Nothing
End If

On Error GoTo Hata

If SapGuiAuto Is Nothing Then

    hataMesaji = "SAP GUI açýk deðil."

    Set SAP_SESSION_AL = Nothing
    Exit Function

End If

'======================================================
' 2. SAP SCRIPTING ENGINE
'======================================================

Set SAPApp = Nothing

On Error Resume Next

Err.Clear
Set SAPApp = SapGuiAuto.GetScriptingEngine

If Err.Number <> 0 Then

    Err.Clear
    Set SAPApp = Nothing

End If

On Error GoTo Hata

If SAPApp Is Nothing Then

    hataMesaji = _
        "SAP GUI açýk ancak Scripting Engine'e baðlanýlamadý."

    Set SAP_SESSION_AL = Nothing
    Exit Function

End If

'======================================================
' 3. BAÐLANTI SAYISINI AL
'======================================================

BaglantiSayisi = 0

On Error Resume Next

Err.Clear
BaglantiSayisi = SAPApp.Children.Count

If Err.Number <> 0 Then

    Err.Clear
    BaglantiSayisi = 0

End If

On Error GoTo Hata

If BaglantiSayisi <= 0 Then

    hataMesaji = _
        "SAP GUI açýk ancak aktif SAP baðlantýsý bulunamadý."

    Set SAP_SESSION_AL = Nothing
    Exit Function

End If

'======================================================
' 4. TÜM BAÐLANTILARI VE OTURUMLARI TARA
'======================================================

For i = 0 To BaglantiSayisi - 1

    Set SAPConnection = Nothing

    On Error Resume Next

    Err.Clear

    Set SAPConnection = _
        SAPApp.Children(CLng(i))

    If Err.Number <> 0 Then

        Err.Clear
        Set SAPConnection = Nothing

    End If

    On Error GoTo Hata

    If Not SAPConnection Is Nothing Then

        OturumSayisi = 0

        On Error Resume Next

        Err.Clear
        OturumSayisi = _
            SAPConnection.Children.Count

        If Err.Number <> 0 Then

            Err.Clear
            OturumSayisi = 0

        End If

        On Error GoTo Hata

        If OturumSayisi > 0 Then

            For j = 0 To OturumSayisi - 1

                Set SAPSession = Nothing

                On Error Resume Next

                Err.Clear

                Set SAPSession = _
                    SAPConnection.Children(CLng(j))

                If Err.Number <> 0 Then

                    Err.Clear
                    Set SAPSession = Nothing

                End If

                On Error GoTo Hata

                If Not SAPSession Is Nothing Then

                    SistemAdi = vbNullString

                    On Error Resume Next

                    Err.Clear
                    SistemAdi = _
                        CStr(SAPSession.Info.SystemName)

                    If Err.Number <> 0 Then

                        Err.Clear
                        SistemAdi = vbNullString

                    End If

                    On Error GoTo Hata

                    SistemAdi = _
                        UCase$(Trim$(SistemAdi))

                    '----------------------------------
                    ' Hedef belirtilmediyse ilk
                    ' kullanýlabilir Session'ý döndür.
                    '----------------------------------
                    If Len(HedefSistem) = 0 Then

                        Set SAP_SESSION_AL = SAPSession
                        Exit Function

                    End If

                    '----------------------------------
                    ' Ýstenen sistem bulundu.
                    '----------------------------------
                    If SistemAdi = HedefSistem Then

                        If HedefSistem = "CDP" Then Set mCDPSession = SAPSession
                        If HedefSistem = "EPP" Then Set mEPPSession = SAPSession
                        Set SAP_SESSION_AL = SAPSession
                        Exit Function

                    End If

                End If

            Next j

        End If

    End If

Next i

'======================================================
' 5. HEDEF BULUNAMADI
'======================================================

If Len(HedefSistem) > 0 Then

    hataMesaji = _
        HedefSistem & _
        " SAP sistemi açýk deðil veya aktif oturum bulunamadý."

Else

    hataMesaji = _
        "Kullanýlabilir aktif SAP oturumu bulunamadý."

End If

Set SAP_SESSION_AL = Nothing

Exit Function

Hata:

hataMesaji = _
    "SAP baðlantýsý kontrol edilirken hata oluþtu." & _
    vbCrLf & _
    "Hata No: " & Err.Number & _
    vbCrLf & _
    "Açýklama: " & Err.Description

Set SAP_SESSION_AL = Nothing

End Function

'==========================================================
' BELÝRLÝ BÝR SAP SÝSTEMÝ AÇIK MI?
'
' Örnek:
' SAP_SISTEM_ACIK_MI("CDP")
' SAP_SISTEM_ACIK_MI("EPP")
'==========================================================
Public Function SAP_SISTEM_ACIK_MI( _
ByVal HedefSistem As String) As Boolean

Dim SAPSession As Object
Dim hataMesaji As String

On Error GoTo Hata

Set SAPSession = _
    SAP_SESSION_AL(HedefSistem, hataMesaji)

SAP_SISTEM_ACIK_MI = _
    Not SAPSession Is Nothing

Exit Function

Hata:

SAP_SISTEM_ACIK_MI = False

End Function

'==========================================================
' HERHANGÝ BÝR SAP BAÐLANTISI VAR MI?
'==========================================================
Public Function SAP_BAGLANTI_VAR_MI() As Boolean

Dim SAPSession As Object
Dim hataMesaji As String

On Error GoTo Hata

Set SAPSession = _
    SAP_SESSION_AL("", hataMesaji)

SAP_BAGLANTI_VAR_MI = _
    Not SAPSession Is Nothing

Exit Function

Hata:

SAP_BAGLANTI_VAR_MI = False

End Function

'==========================================================
' SAP SÝSTEM TESTÝ
'
' HedefSistem:
' CDP veya EPP
'==========================================================
Public Sub SAP_SISTEM_TEST( _
ByVal HedefSistem As String)

Dim SAPSession As Object
Dim hataMesaji As String

Dim SistemAdi As String
Dim ClientNo As String
Dim Transaction As String

Dim GecenSure As Double

On Error GoTo Hata

HedefSistem = _
    UCase$(Trim$(HedefSistem))

If HedefSistem <> "CDP" And _
   HedefSistem <> "EPP" Then

    MsgBox _
        "Geçersiz SAP sistem adý." & _
        vbCrLf & vbCrLf & _
        "Geçerli deðerler: CDP veya EPP", _
        vbExclamation, _
        "AKSORA"

    Exit Sub

End If

CORE_ISLEM_BASLAT _
    HedefSistem & " SAP baðlantý kontrolü"

CORE_DURUM _
    HedefSistem & _
    " SAP oturumu aranýyor..."

Set SAPSession = _
    SAP_SESSION_AL(HedefSistem, hataMesaji)

If SAPSession Is Nothing Then

    GecenSure = CORE_ISLEM_BITIR()

    MsgBox _
        HedefSistem & _
        " SAP baðlantýsý bulunamadý." & _
        vbCrLf & vbCrLf & _
        hataMesaji, _
        vbExclamation, _
        "AKSORA - SAP"

    Exit Sub

End If

SistemAdi = vbNullString
ClientNo = vbNullString
Transaction = vbNullString

On Error Resume Next

SistemAdi = _
    CStr(SAPSession.Info.SystemName)

ClientNo = _
    CStr(SAPSession.Info.Client)

Transaction = _
    CStr(SAPSession.Info.Transaction)

On Error GoTo Hata

GecenSure = CORE_ISLEM_BITIR()

MsgBox _
    HedefSistem & _
    " SAP baðlantýsý baþarýyla bulundu." & _
    vbCrLf & vbCrLf & _
    "Sistem     : " & SistemAdi & vbCrLf & _
    "Client     : " & ClientNo & vbCrLf & _
    "Ýþlem      : " & Transaction & vbCrLf & _
    "Kontrol    : " & _
    AKSORA_SAYI_METNI(GecenSure, "0.00") & " sn", _
    vbInformation, _
    "AKSORA - " & HedefSistem

Exit Sub

Hata:

hataMesaji = Err.Description

GecenSure = CORE_ISLEM_BITIR()

MsgBox _
    HedefSistem & _
    " SAP baðlantý testi sýrasýnda hata oluþtu." & _
    vbCrLf & vbCrLf & _
    "Hata No: " & Err.Number & _
    vbCrLf & _
    "Açýklama: " & hataMesaji, _
    vbCritical, _
    "AKSORA"

End Sub

'==========================================================
' CDP TEST
'
' Alt + F8 listesinde görünmesi için parametresiz
' test prosedürü.
'==========================================================
Public Function SAP_DURUM_METNI( _
ByVal HedefSistem As String) As String

If SAP_SISTEM_ACIK_MI(HedefSistem) Then

    SAP_DURUM_METNI = _
        UCase$(Trim$(HedefSistem)) & " BAÐLI"

Else

    SAP_DURUM_METNI = _
        UCase$(Trim$(HedefSistem)) & " BAÐLI DEÐÝL"

End If

End Function


Public Function AKSORA_SAP_BLOK_BOYU(ByVal alv As Object) As Long
    Dim gorunen As Long
    AKSORA_SAP_BLOK_BOYU = 30
    On Error GoTo Cikis
    gorunen = CLng(alv.VisibleRowCount)
    If gorunen < 1 Then Exit Function
    If gorunen > 60 Then gorunen = 60
    AKSORA_SAP_BLOK_BOYU = gorunen
Cikis:
End Function

Public Sub AKSORA_SAP_OLCUM_KAYDET(ByVal il As String, ByVal adim As Long, _
    ByVal satir As Long, ByVal ToplamSure As Double, ByVal kontrolSure As Double)
    On Error Resume Next
    If mSAPDetay Is Nothing Then Set mSAPDetay = CreateObject("Scripting.Dictionary")
    mSAPDetay(il & "|" & CStr(adim)) = Array(satir, ToplamSure, kontrolSure)
    On Error GoTo 0
End Sub

Public Sub AKSORA_SAP_HIZ_OLC()
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Then Exit Sub
    Set mSAPDetay = CreateObject("Scripting.Dictionary")
    AKSORA_BUTUN_ILLERI_GUNCELLE
    AKSORA_SAP_HIZ_RAPORU
End Sub

Public Sub AKSORA_SAP_HIZ_RAPORU()
    Dim il As Variant, etiket As Variant, veri As Variant, adim As Long
    Dim n As Long, adet As Long, sure As Double, kontrol As Double, metin As String
    If mSAPDetay Is Nothing Then
        MsgBox "Önce AKSORA_SAP_HIZ_OLC komutuyla ölçümlü güncelleme yapýn.", vbInformation, "AKSORA | SAP Ölçümü"
        Exit Sub
    End If
    If mSAPDetay.Count = 0 Then
        MsgBox "Ölçülen baþarýlý SAP adýmý yok. Güncelleme iptal edilmiþ veya baþarýsýz olmuþ olabilir.", vbInformation, "AKSORA | SAP Ölçümü"
        Exit Sub
    End If
    etiket = Array("Normal CDP", "Normal EPP", "Ýlçe CDP", "Ýlçe EPP", "Yüklenici CDP", "Yüklenici EPP")
    metin = "Baþarýlý adýmlarýn toplamý (5 il):" & vbCrLf
    Debug.Print "ÝL | ADIM | SATIR | SAP OKUYUCU sn | EK VERÝ KONTROLÜ sn"
    For adim = 1 To 6
        n = 0: adet = 0: sure = 0: kontrol = 0
        For Each il In Array("ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
            If mSAPDetay.Exists(CStr(il) & "|" & CStr(adim)) Then
                veri = mSAPDetay(CStr(il) & "|" & CStr(adim))
                n = n + CLng(veri(0)): sure = sure + CDbl(veri(1))
                kontrol = kontrol + CDbl(veri(2)): adet = adet + 1
                Debug.Print CStr(il) & " | " & CStr(etiket(adim - 1)) & " | " & veri(0) & _
                    " | " & AKSORA_SAYI_METNI(veri(1), "0.000") & " | " & AKSORA_SAYI_METNI(veri(2), "0.000")
            End If
        Next il
        metin = metin & vbCrLf & CStr(etiket(adim - 1)) & ": " & AKSORA_SAYI_METNI(sure, "0.1") & _
            " sn / " & AKSORA_SAYI_METNI(n, "#,##0") & " satýr (" & adet & "/5)"
    Next adim
    metin = metin & vbCrLf & vbCrLf & "Ayrýntýlar Ctrl+G penceresinde. Eksik adýmlar toplama dahil deðil."
    MsgBox metin, vbInformation, "AKSORA | SAP Adým Süreleri"
End Sub

