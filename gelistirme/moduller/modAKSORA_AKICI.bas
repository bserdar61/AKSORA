Attribute VB_Name = "modAKSORA_AKICI"
Option Explicit

'==========================================================
' AKSORA v3.1  |  modAKSORA_AKICI
'
' Bekleme kartinin (frmSplash) akici gorunmesi.
'
' Sorun: kart yalniz VBA kodu "ilerle" dedigi anda cizilir. SAP sorgusu surerken
' (10-20 sn) VBA beklemededir; cubuk ve gecen sure donar, ekran takilmis gorunur.
'
' Cozum: kart acikken bir Windows zamanlayicisi 40 ms'de bir kartta
'   - cubugun uzerinde saga sola kayan bir serit gezdirir (sistem calisiyor),
'   - "Gecen sure" yazisini saniye saniye gunceller.
' Zamanlayici Excel'in is parcaciginda, Windows mesaj dagittikca calisir; hesap yapmaz,
' SAP'ye ya da sayfalara dokunmaz, yalniz karti cizer.
'
' Baglanti noktalari (kurulum betigi ekler):
'   AKSORA_BEKLE_AC     : AKSORA_AKICI_BASLAT
'   AKSORA_BEKLE_KAPAT  : AKSORA_AKICI_DURDUR
'
' Not: kod durdurulup VBA duzenleyicisinde hata ayiklamaya gecilecekse once
' Alt+F8 > AKSORA_AKICI_DURDUR calistirin (zamanlayici acikken kesme kipi Excel'i kilitleyebilir).
'==========================================================

Private Declare PtrSafe Function AKC_SetTimer Lib "user32" Alias "SetTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr, _
    ByVal uElapse As Long, ByVal lpTimerFunc As LongPtr) As LongPtr
Private Declare PtrSafe Function AKC_KillTimer Lib "user32" Alias "KillTimer" (ByVal hWnd As LongPtr, ByVal nIDEvent As LongPtr) As Long

Private Const AKC_ARALIK As Long = 40           ' ms
Private Const AKC_TUR As Double = 2.4           ' serit bir gidis-donusu kac saniyede tamamlar
Private Const AKC_SERIT_ORAN As Double = 0.18   ' seridin cubuga orani
Private Const AKC_SURE_ETIKET As String = "Geçen süre: "

Private mZaman As LongPtr
Private mBas As Double
Private mAktif As Boolean
Private mIcinde As Boolean
Private mSonSn As Long

' Kart acildiktan hemen sonra cagrilir.
Public Sub AKSORA_AKICI_BASLAT()
    On Error Resume Next
    AKSORA_AKICI_DURDUR
    If AKC_KART() Is Nothing Then Exit Sub
    mBas = Timer: mSonSn = -1: mIcinde = False
    AKC_SERIT_HAZIRLA
    mAktif = True
    mZaman = AKC_SetTimer(0, 0, AKC_ARALIK, AddressOf AKSORA_AKICI_ZAMAN)
    If mZaman = 0 Then mAktif = False
End Sub

' Kart kapatilmadan hemen once cagrilir (birden cok kez cagrilabilir).
Public Sub AKSORA_AKICI_DURDUR()
    On Error Resume Next
    mAktif = False
    If mZaman <> 0 Then AKC_KillTimer 0, mZaman
    mZaman = 0
End Sub

' Zamanlayici geri cagrisi. Asla hata firlatmaz; kart yoksa ya da durum sifirlandiysa kendini kapatir.
Public Sub AKSORA_AKICI_ZAMAN(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal idEvent As LongPtr, ByVal dwTime As Long)
    On Error Resume Next
    If Not mAktif Then
        AKC_KillTimer 0, idEvent
        Exit Sub
    End If
    If mIcinde Then Exit Sub
    mIcinde = True
    AKC_CIZ
    mIcinde = False
End Sub

' Seridin cubuk uzerindeki konumu (0 = sol, 1 = sag): yumusak gidis-donus.
Public Function AKSORA_AKICI_KONUM(ByVal gecen As Double) As Double
    Dim faz As Double
    If gecen < 0 Then gecen = 0
    faz = gecen / AKC_TUR
    faz = faz - Int(faz)
    AKSORA_AKICI_KONUM = 0.5 - 0.5 * Cos(faz * 6.28318530717959)
End Function

' Yuklu kart (frmSplash); yuklu degilse Nothing. (Ada dogrudan basvurmak formu yeniden yuklerdi.)
Private Function AKC_KART() As Object
    Dim f As Object
    On Error Resume Next
    For Each f In VBA.UserForms
        If f.Name = "frmSplash" Then
            Set AKC_KART = f
            Exit Function
        End If
    Next f
End Function

Private Sub AKC_SERIT_HAZIRLA()
    Dim frm As Object, iz As Object, c As Object
    On Error Resume Next
    Set frm = AKC_KART()
    If frm Is Nothing Then Exit Sub
    Set iz = frm.Controls("lblBarArka")
    If iz Is Nothing Then Exit Sub
    Set c = frm.Controls("lblAkis")
    If c Is Nothing Then
        Err.Clear
        Set c = frm.Controls.Add("Forms.Label.1", "lblAkis", True)
    End If
    If c Is Nothing Then Exit Sub
    With c
        .Caption = vbNullString
        .BackStyle = fmBackStyleOpaque
        .BackColor = RGB(94, 234, 212)
        .BorderStyle = fmBorderStyleNone
        .Height = iz.Height
        .Top = iz.Top
        .Width = iz.Width * AKC_SERIT_ORAN
        .Left = iz.Left
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub AKC_CIZ()
    Dim frm As Object, iz As Object, c As Object, gecen As Double, sn As Long, t As String, p As Long
    On Error Resume Next
    Set frm = AKC_KART()
    If frm Is Nothing Then
        AKSORA_AKICI_DURDUR
        Exit Sub
    End If
    gecen = Timer - mBas
    If gecen < 0 Then gecen = gecen + 86400#
    Set iz = frm.Controls("lblBarArka")
    Set c = frm.Controls("lblAkis")
    If Not (iz Is Nothing Or c Is Nothing) Then
        c.Left = iz.Left + (iz.Width - c.Width) * AKSORA_AKICI_KONUM(gecen)
    End If
    ' "Gecen sure" satiri (tek tus guncelleme karti) saniyede bir yenilenir.
    sn = CLng(Fix(gecen))
    If sn <> mSonSn Then
        mSonSn = sn
        t = frm.Controls("lblDurum").Caption
        p = InStr(1, t, AKC_SURE_ETIKET)
        If p > 0 Then
            frm.Controls("lblDurum").Caption = Left$(t, p + Len(AKC_SURE_ETIKET) - 1) & (sn \ 60) & " dk " & Format$(sn Mod 60, "00") & " sn"
        End If
    End If
    frm.Repaint
End Sub
