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


    On Error GoTo Hata


    application.Visible = False

    mExcelGizli = True


    Load frmSplash

    frmSplash.Show vbModeless

    frmSplash.SplashBaslat


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


    frmSplash.SplashIlerle _
        Yuzde, _
        mesaj


    On Error GoTo 0

End Sub


'==========================================================
' SPLASH KAPAT
'==========================================================
Public Sub SPLASH_KAPAT()

    If mSplashKapandi Then Exit Sub


    mSplashKapandi = True


    On Error Resume Next


    application.Visible = True

    application.ScreenUpdating = True


    mExcelGizli = False


    Unload frmSplash


    On Error GoTo 0

End Sub

