Attribute VB_Name = "frmSplash"
Attribute VB_Base = "0{401A6CCB-5AF7-46B5-A7FD-8360CD4A00B0}{69966854-3B19-4E11-9F79-FD4669BB5FDD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" ( _
        ByVal lpClassName As String, _
        ByVal lpWindowName As String) As LongPtr

    Private Declare PtrSafe Function SetForegroundWindow Lib "user32" ( _
        ByVal hWnd As LongPtr) As Long

    Private Declare PtrSafe Function BringWindowToTop Lib "user32" ( _
        ByVal hWnd As LongPtr) As Long

    Private Declare PtrSafe Function SetWindowPos Lib "user32" ( _
        ByVal hWnd As LongPtr, _
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

' Stok arama kipi: kart basliksiz yazi kutusu olur, her harfte tablo suzulur.
Public AraKipi As Boolean
Private WithEvents mAraTb As MSForms.TextBox
Attribute mAraTb.VB_VarHelpID = -1
Private mAraYer As Object

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
    If Not (AKSORA_ADMIN_TOPLU Or AKSORA_ATIL_CALISIYOR) Then Exit Sub
    If AKSORA_ADMIN_IPTAL Then Exit Sub
    AKSORA_ADMIN_IPTAL = True
    Me.lblImza.Caption = "DURDURULUYOR..."
    Me.lblImza.BackColor = RGB(148, 163, 184)
    If AKSORA_ATIL_CALISIYOR Then
        Me.lblDurum.Caption = "Durduruluyor: çalýþan SAP sorgusu bitince duracak..."
    Else
        Me.lblDurum.Caption = "Durduruluyor: çalýþan il bitince iþlem duracak..."
    End If
    Me.Repaint
End Sub

' ARKA PLANA AL: Excel simge durumuna alinir, guncelleme surer; bitince pencere geri acilir.
Private Sub lblAltBaslik_Click()
    If Not (AKSORA_ADMIN_TOPLU Or AKSORA_ATIL_CALISIYOR Or AKSORA_PANEL_KART_ACIK()) Then Exit Sub
    ' Tiklama aninda gorunur geri bildirim (pencere arkaya gidince baska pencere yoksa degisiklik fark edilmiyordu).
    If AKSORA_ARKA_PLAN Then
        Me.lblAltBaslik.Caption = "ARKA PLANA AL"
        Me.lblAltBaslik.BackColor = RGB(21, 101, 192)
        Me.Repaint
        AKSORA_ONE_GETIR
    Else
        Me.lblAltBaslik.Caption = "ARKA PLANDA  •  ÖNE GETÝR"
        Me.lblAltBaslik.BackColor = RGB(21, 128, 61)
        Me.lblDurum.Caption = "Arka planda çalýþýyor: baþka programlarla çalýþabilirsiniz." & vbLf & _
            "Ýþ bitince görev çubuðunda AKSORA yanýp söner."
        Me.Repaint
        AKSORA_ARKA_PLANA_AL
    End If
End Sub

' Kart Alt+F4 ile kapatilamaz; islem bitince kod kapatir.
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu And Not AraKipi Then Cancel = True
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

' Kart, sayfadaki arama kutusunun ICINE (simgeden sonra) beyaz ve kenarsiz oturur: kutunun gorunusu degismez,
' yalniz imlec yanip soner. Bos iken ayni gri ornek yazi gorunur, yazinca kaybolur.
Public Sub AraKur(ByVal metin As String)
    Dim c As Object
    On Error Resume Next
    AraKipi = True
    For Each c In Me.Controls
        c.Visible = False
    Next c
    Me.BackColor = RGB(255, 255, 255)
    Me.BorderStyle = fmBorderStyleNone
    Me.Caption = "AKSORA ARA"
    Set mAraYer = Me.Controls.Add("Forms.Label.1", "lblAraYer", True)
    With mAraYer
        .Caption = "Malzeme ara: kod ya da ad (" & ChrW(246) & "r. 6533, 50 kva trafo)"
        .BackStyle = fmBackStyleTransparent: .ForeColor = RGB(148, 163, 184)
        .Font.Name = "Aptos": .WordWrap = False
    End With
    Set mAraTb = Me.Controls.Add("Forms.TextBox.1", "tbAra", True)
    With mAraTb
        .BorderStyle = fmBorderStyleNone
        .SpecialEffect = fmSpecialEffectFlat
        .BackStyle = fmBackStyleTransparent
        .ForeColor = RGB(15, 23, 42)
        .Font.Name = "Aptos"
        .Font.bold = False
        .Text = metin
    End With
End Sub

Public Sub AraBoyut(ByVal fs As Single)
    Dim w As Single, h As Single
    On Error Resume Next
    w = Me.InsideWidth: h = Me.InsideHeight
    With mAraTb
        .Font.Size = fs
        .Height = application.Min(h, fs * 1.75)
        .Left = 0: .top = (h - .Height) / 2
        .Width = w
    End With
    With mAraYer
        .Font.Size = fs
        .Left = 3: .Width = w - 3: .Height = mAraTb.Height: .top = mAraTb.top + 2
        .Visible = (Len(mAraTb.Text) = 0)
    End With
    mAraTb.ZOrder 0
    mAraTb.Visible = True
    mAraTb.SelStart = Len(mAraTb.Text)
    mAraTb.SetFocus
End Sub

Private Sub mAraTb_Change()
    On Error Resume Next
    mAraYer.Visible = (Len(mAraTb.Text) = 0)
    On Error GoTo 0
    If AraKipi Then AKSORA_STOK_ARA_YAZ CStr(mAraTb.Text)
End Sub

Private Sub mAraTb_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Or KeyCode = 27 Or KeyCode = 9 Then
        KeyCode = 0
        AKSORA_ARA_KART_KAPAT
    End If
End Sub
