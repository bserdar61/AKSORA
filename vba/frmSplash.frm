VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSplash 
   ClientHeight    =   4032
   ClientLeft      =   144
   ClientTop       =   408
   ClientWidth     =   7824
   OleObjectBlob   =   "frmSplash.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long

    Private Declare PtrSafe Function BringWindowToTop Lib "user32" ( _
        ByVal hwnd As LongPtr) As Long

    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As LongPtr, _
        ByVal hWndInsertAfter As LongPtr, _
        ByVal x As Long, _
        ByVal y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
#Else
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As Long

    Private Declare Function SetForegroundWindow Lib "user32" ( _
        ByVal hwnd As Long) As Long

    Private Declare Function BringWindowToTop Lib "user32" ( _
        ByVal hwnd As Long) As Long

    Private Declare Function SetWindowPos Lib "user32" ( _
        ByVal hwnd As Long, _
        ByVal hWndInsertAfter As Long, _
        ByVal X As Long, _
        ByVal Y As Long, _
        ByVal cx As Long, _
        ByVal cy As Long, _
        ByVal wFlags As Long) As Long
#End If

Private Const HWND_TOPMOST As Long = -1
Private Const HWND_NOTOPMOST As Long = -2
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_SHOWWINDOW As Long = &H40

Public Sub SplashBaslat()

    SplashOrtala

    Me.lblAltBaslik.Caption = "Stok ve Verimlilik, AKSORA!"
    Me.lblBar.Width = 1
    Me.lblDurum.Caption = "Sistem hazýrlanýyor..."

    Me.Repaint
    DoEvents

    SplashOneGetir

End Sub

Public Sub SplashIlerle(ByVal Yuzde As Double, ByVal mesaj As String)

    If Yuzde < 0 Then Yuzde = 0
    If Yuzde > 1 Then Yuzde = 1

    Me.lblDurum.Caption = mesaj
    Me.lblBar.Width = Me.lblBarArka.Width * Yuzde

    Me.Repaint
    DoEvents

    SplashOneGetir

End Sub

Private Sub UserForm_Initialize()

    On Error Resume Next
    Me.BackColor = RGB(255, 255, 255)

    Me.lblBaslik.ForeColor = RGB(31, 78, 121)
    Me.lblAltBaslik.ForeColor = RGB(110, 125, 140)
    Me.lblDurum.ForeColor = RGB(31, 78, 121)
    Me.lblImza.ForeColor = RGB(110, 125, 140)

    Me.imgLogo.BackColor = RGB(255, 255, 255)

    Me.lblBarArka.BackColor = RGB(224, 228, 234)
    Me.lblBar.BackColor = RGB(47, 144, 217)

    Me.lblBarArka.Width = 260
    Me.lblBar.Width = 1
AKSORA_SPLASH_GORSEL_UYGULA Me
    SplashOrtala

End Sub

' Tum birimler guncellemesinde DURDUR: calisan il bitince islem durur.
Private Sub lblImza_Click()
    If Not AKSORA_ADMIN_TOPLU Then Exit Sub
    If AKSORA_ADMIN_IPTAL Then Exit Sub
    AKSORA_ADMIN_IPTAL = True
    Me.lblImza.Caption = "DURDURULUYOR..."
    Me.lblImza.BackColor = RGB(148, 163, 184)
    Me.lblDurum.Caption = "Durduruluyor: çalýþan il bitince iþlem duracak..."
    Me.Repaint
End Sub

' Kart Alt+F4 ile kapatilamaz; islem bitince kod kapatir.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

Private Sub SplashOrtala()

    Dim merkez As Double

    merkez = Me.InsideWidth / 2

    Me.imgLogo.Left = merkez - (Me.imgLogo.Width / 2)

    Me.lblBaslik.Left = merkez - (Me.lblBaslik.Width / 2)
    Me.lblAltBaslik.Left = merkez - (Me.lblAltBaslik.Width / 2)
    Me.lblDurum.Left = merkez - (Me.lblDurum.Width / 2)

    Me.lblBarArka.Left = merkez - (Me.lblBarArka.Width / 2)
    Me.lblBar.Left = Me.lblBarArka.Left

    Me.lblImza.Left = merkez - (Me.lblImza.Width / 2)
    Me.lblImza.Width = 300
Me.lblImza.Left = (Me.InsideWidth - Me.lblImza.Width) / 2

End Sub

Private Sub SplashOneGetir()

#If VBA7 Then
    Dim hSplash As LongPtr
#Else
    Dim hSplash As Long
#End If

    On Error Resume Next

    hSplash = FindWindow("ThunderDFrame", Me.Caption)

    If hSplash = 0 Then
        hSplash = FindWindow("ThunderXFrame", Me.Caption)
    End If

    If hSplash <> 0 Then

        SetWindowPos hSplash, HWND_TOPMOST, 0, 0, 0, 0, _
                     SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW

        BringWindowToTop hSplash
        SetForegroundWindow hSplash

        SetWindowPos hSplash, HWND_NOTOPMOST, 0, 0, 0, 0, _
                     SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW

    End If

    On Error GoTo 0

End Sub

