Attribute VB_Name = "modAKSORA_CEKIRDEK"
Option Explicit

'==========================================================
' AKSORA
' modCORE
'
' Sistemin ortak altyapýsý:
' - Excel performans yönetimi
' - Güvenli baþlatma / bitirme
' - EnableEvents korumasý
' - Ýþlem süresi ölçümü
' - Excel durumunu eski haline getirme
'==========================================================

Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mDisplayAlerts As Boolean
Private mCalculation As XlCalculation
Private mStatusBar As Variant

Private mBaslangicZamani As Double
Private mAktifIslem As String
Private mSistemAktif As Boolean

'==========================================================
' ÝÞLEMÝ GÜVENLÝ ÞEKÝLDE BAÞLAT
'==========================================================
Public Sub CORE_ISLEM_BASLAT(Optional ByVal IslemAdi As String = "")

On Error GoTo Hata

' Zaten CORE tarafýndan baþlatýlmýþ bir iþlem varsa
' Excel'in mevcut durumunu tekrar kaydetme.
If mSistemAktif Then Exit Sub

' Excel'in mevcut durumunu hafýzaya al
mScreenUpdating = application.ScreenUpdating
mEnableEvents = application.EnableEvents
mDisplayAlerts = application.DisplayAlerts
mCalculation = application.Calculation
mStatusBar = application.StatusBar

mAktifIslem = IslemAdi
mBaslangicZamani = Timer
mSistemAktif = True

' Performans modu
application.ScreenUpdating = False
application.EnableEvents = False
application.DisplayAlerts = False
application.Calculation = xlCalculationManual

If Len(IslemAdi) > 0 Then
    application.StatusBar = IslemAdi & " iþlemi devam ediyor..."
Else
    application.StatusBar = "Ýþlem devam ediyor..."
End If

Exit Sub

Hata:

' Baþlatma sýrasýnda hata oluþursa Excel'i
' güvenli varsayýlan duruma getir.
On Error Resume Next

application.ScreenUpdating = True
application.EnableEvents = True
application.DisplayAlerts = True
application.Calculation = xlCalculationAutomatic
application.StatusBar = False

mSistemAktif = False

On Error GoTo 0

MsgBox _
    "CORE sistemi baþlatýlýrken hata oluþtu." & _
    vbCrLf & vbCrLf & _
    "Hata: " & Err.Description & _
    vbCrLf & _
    "Hata No: " & Err.Number, _
    vbCritical, _
    "AKSORA"

End Sub

'==========================================================
' ÝÞLEMÝ GÜVENLÝ ÞEKÝLDE BÝTÝR
'==========================================================
Public Function CORE_ISLEM_BITIR() As Double

Dim GecenSure As Double

On Error Resume Next

If Not mSistemAktif Then
    CORE_ISLEM_BITIR = 0
    Exit Function
End If

GecenSure = CORE_GECEN_SURE()

' Excel'i iþlemden önceki durumuna döndür
If ANA_SESSIZ_MOD Then
    application.ScreenUpdating = False
Else
    application.ScreenUpdating = mScreenUpdating
End If
application.EnableEvents = mEnableEvents
application.DisplayAlerts = mDisplayAlerts
application.Calculation = mCalculation
application.StatusBar = mStatusBar

mSistemAktif = False
mAktifIslem = vbNullString

CORE_ISLEM_BITIR = GecenSure

On Error GoTo 0

End Function

'==========================================================
' ACÝL DURUMDA EXCEL'Ý NORMALE DÖNDÜR
'
' VBA çalýþmasý yarýda kesilirse veya olaylar kapalý
' kalýrsa Alt+F8 üzerinden bunu çalýþtýrabilirsin.
'==========================================================
Public Sub CORE_EXCEL_KURTAR()

On Error Resume Next

application.ScreenUpdating = True
application.EnableEvents = True
application.DisplayAlerts = True
application.Calculation = xlCalculationAutomatic
application.StatusBar = False
application.Cursor = xlDefault

mSistemAktif = False
mAktifIslem = vbNullString

On Error GoTo 0

MsgBox _
    "Excel güvenli çalýþma durumuna getirildi." & _
    vbCrLf & vbCrLf & _
    "Ekran güncelleme : AÇIK" & vbCrLf & _
    "Olaylar          : AÇIK" & vbCrLf & _
    "Uyarýlar         : AÇIK" & vbCrLf & _
    "Hesaplama        : OTOMATÝK", _
    vbInformation, _
    "AKSORA"

End Sub

'==========================================================
' GEÇEN SÜRE
'==========================================================
Public Function CORE_GECEN_SURE() As Double

Dim bitis As Double

If Not mSistemAktif Then
    CORE_GECEN_SURE = 0
    Exit Function
End If

bitis = Timer

' Gece 00:00 geçiþine karþý koruma
If bitis >= mBaslangicZamani Then
    CORE_GECEN_SURE = bitis - mBaslangicZamani
Else
    CORE_GECEN_SURE = (86400# - mBaslangicZamani) + bitis
End If

End Function

'==========================================================
' DURUM ÇUBUÐUNDA MESAJ GÖSTER
'==========================================================
Public Sub CORE_DURUM(ByVal mesaj As String)
Static sonYazim As Double
On Error Resume Next
application.StatusBar = mesaj
AKSORA_ADMIN_DURUM mesaj
' Panelden baslatilan guncellemede anlik durum seritte de gorunur.
If AKSORA_PANEL_CANLI Then
    If Abs(Timer - sonYazim) >= 0.5 Then
        sonYazim = Timer
        AKSORA_PANEL_SERIT "Güncelleniyor  •  " & mesaj
    End If
End If
On Error GoTo 0

End Sub

'==========================================================
' CORE AKTÝF MÝ?
'==========================================================
Public Function CORE_AKTIF_MI() As Boolean

CORE_AKTIF_MI = mSistemAktif

End Function

'==========================================================
' AKTÝF ÝÞLEM ADI
'==========================================================
Public Function CORE_AKTIF_ISLEM() As String

CORE_AKTIF_ISLEM = mAktifIslem

End Function

