Attribute VB_Name = "Sayfa76"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim q As String
    If Intersect(Target, Me.Range("B6:J6")) Is Nothing Then Exit Sub
    If AKSORA_V11_MESGUL Then Exit Sub
    application.ScreenUpdating = False
    On Error Resume Next
    q = Trim$(CStr(Me.Range("B6").Value2))
    On Error GoTo 0
    If Len(q) = 0 Then
        AKSORA_MUSTERI_ARAMA_TEMIZLE
    Else
        AKSORA_MUSTERI_YEREL_ARA
    End If
    On Error Resume Next
    application.ScreenUpdating = False
    application.MoveAfterReturn = False
    Me.Range("B6").Select
    ActiveWindow.ScrollRow = 1: ActiveWindow.ScrollColumn = 1
    application.ScreenUpdating = True
End Sub

Private Sub Worksheet_Calculate()
    If AKSORA_V11_MESGUL Or AKSORA_V13_ACILIS Then Exit Sub
    On Error Resume Next
    AKSORA_MUSTERI_BANT Me
End Sub
