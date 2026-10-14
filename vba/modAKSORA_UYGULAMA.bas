Attribute VB_Name = "modAKSORA_UYGULAMA"
Option Explicit

'==========================================================
' AKSORA - UYGULAMA / EXCEL GORUNUMU
'==========================================================

Private mRibbonGorunur As Boolean
Private mFormulaBarGorunur As Boolean
Private mStatusBarGorunur As Boolean
Private mHeadingsGorunur As Boolean
Private mGridlinesGorunur As Boolean
Private mYatayCubukGorunur As Boolean
Private mDikeyCubukGorunur As Boolean
Private mSekmelerGorunur As Boolean
Private mDurumKaydedildi As Boolean


'==========================================================
' AKSORA - SPLASH
'==========================================================

Private mExcelGizli As Boolean
Private mSplashKapandi As Boolean
Private mSplashBas As Double
Private mSplashYuzde As Double


'==========================================================
' UYGULAMA MODUNU AC
'==========================================================
Public Sub UYGULAMA_MODU_AC()

    AKSORA_SAYI_AYIRACLARI

    On Error Resume Next


    If Not mDurumKaydedildi Then

        mRibbonGorunur = _
            application.CommandBars("Ribbon").Visible

        mFormulaBarGorunur = _
            application.DisplayFormulaBar

        mStatusBarGorunur = _
            application.DisplayStatusBar


        With ThisWorkbook.Windows(1)

            mHeadingsGorunur = _
                .DisplayHeadings

            mGridlinesGorunur = _
                .DisplayGridlines

            mYatayCubukGorunur = _
                .DisplayHorizontalScrollBar

            mDikeyCubukGorunur = _
                .DisplayVerticalScrollBar

            mSekmelerGorunur = _
                .DisplayWorkbookTabs

        End With


        mDurumKaydedildi = True

    End If


    application.ExecuteExcel4Macro _
        "SHOW.TOOLBAR(""Ribbon"",False)"


    application.DisplayFormulaBar = False


    application.StatusBar = False

    application.DisplayStatusBar = True


    With ThisWorkbook.Windows(1)

        .DisplayWorkbookTabs = True

        .DisplayHeadings = False

        .DisplayGridlines = False

        .DisplayHorizontalScrollBar = True

        .DisplayVerticalScrollBar = True

    End With

End Sub


'==========================================================
' UYGULAMA MODUNU KAPAT
'==========================================================
Public Sub UYGULAMA_MODU_KAPAT()

    On Error Resume Next


    If ANA_SESSIZ_MOD Then Exit Sub

    application.UseSystemSeparators = True
    application.MoveAfterReturn = True


    If mDurumKaydedildi Then

        application.ExecuteExcel4Macro _
            "SHOW.TOOLBAR(""Ribbon""," & _
            IIf(mRibbonGorunur, "True", "False") & ")"


        application.DisplayFormulaBar = _
            mFormulaBarGorunur


        application.StatusBar = False

        application.DisplayStatusBar = True


        With ThisWorkbook.Windows(1)

            .DisplayHeadings = _
                mHeadingsGorunur

            .DisplayGridlines = _
                mGridlinesGorunur

            .DisplayHorizontalScrollBar = _
                mYatayCubukGorunur

            .DisplayVerticalScrollBar = _
                mDikeyCubukGorunur

            .DisplayWorkbookTabs = _
                mSekmelerGorunur

        End With

    End If


    mDurumKaydedildi = False

End Sub


'==========================================================
' SPLASH GOSTER
'==========================================================
Public Sub SPLASH_GOSTER()

    mSplashKapandi = False

    mExcelGizli = False

    mSplashBas = Timer: mSplashYuzde = 0


    On Error GoTo Hata


    application.Visible = False

    mExcelGizli = True


    Load frmSplash

    frmSplash.Show vbModeless

    frmSplash.SplashBaslat

    SPLASH_ISARET


    Exit Sub


Hata:

    application.Visible = True

    mExcelGizli = False

End Sub


'==========================================================
' SPLASH DURUM
'==========================================================
Public Sub SPLASH_DURUM( _
    ByVal Yuzde As Double, _
    ByVal mesaj As String)

    On Error Resume Next


    If mSplashKapandi Then Exit Sub

    ' Cubuk hedefe adim adim dolar (yukleme ekrani goz onunde ilerler).
    SPLASH_AKIT Yuzde, mesaj


    On Error GoTo 0

End Sub


'==========================================================
' SPLASH KAPAT
'==========================================================
Private Sub SPLASH_AKIT(ByVal hedef As Double, ByVal mesaj As String)
    Dim i As Long, adim As Long, bas As Double, y As Double, t As Double
    On Error Resume Next
    If hedef > 1 Then hedef = 1
    If hedef < mSplashYuzde Then hedef = mSplashYuzde
    adim = 12
    bas = mSplashYuzde
    For i = 1 To adim
        y = bas + (hedef - bas) * i / adim
        frmSplash.SplashIlerle y, mesaj
        t = Timer
        Do While Timer - t < 0.025 And Timer >= t
            DoEvents
        Loop
    Next i
    mSplashYuzde = hedef
    On Error GoTo 0
End Sub

' Masaustu baslaticisi (AKSORA_BASLAT.hta) bu isareti gorunce kendi bekleme ekranini kapatir.
Private Sub SPLASH_ISARET()
    Dim f As Integer
    On Error Resume Next
    f = FreeFile
    Open Environ$("TEMP") & "\AKSORA_ACILIS.flag" For Output As #f
    Print #f, "HAZIR"
    Close #f
    On Error GoTo 0
End Sub

Public Sub SPLASH_KAPAT()

    Dim gecen As Double

    If mSplashKapandi Then Exit Sub

    ' Yukleme ekrani en az ~2 sn gorunur ve %100'e dolarak kapanir.
    On Error Resume Next
    SPLASH_AKIT 1, "AKSORA hazýr"
    gecen = Timer - mSplashBas
    If gecen < 0 Then gecen = gecen + 86400#
    Do While gecen < 2
        DoEvents
        gecen = Timer - mSplashBas
        If gecen < 0 Then gecen = gecen + 86400#
    Loop
    On Error GoTo 0


    mSplashKapandi = True


    On Error Resume Next


    application.Visible = True

    application.ScreenUpdating = True


    mExcelGizli = False


    Unload frmSplash


    On Error GoTo 0

End Sub

