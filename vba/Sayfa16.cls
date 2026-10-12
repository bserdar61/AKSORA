VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private mOzetCalisiyor As Boolean
Private mSonGorunen As String

Private Sub Worksheet_Activate()
    If AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Or AKSORA_V11_MESGUL Then Exit Sub
    On Error GoTo Hata
    AKSORA_V13_GORUNUM Me
    Exit Sub
Hata:
    MsgBox "Arama ekraný açýlamadý: " & Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim olay As Boolean
    Dim hataMetni As String

    If Intersect(Target, Me.Range("B6").MergeArea) Is Nothing Then Exit Sub

    olay = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False

    If Len(Trim$(CStr(Me.Range("B6").Value2))) < 2 Then
        AKSORA_ARAMA_SONUCLARI_TEMIZLE
    Else
        AKSORA_IL_ARAMA_YAP
    End If

    application.EnableEvents = olay
    Exit Sub

Hata:
    hataMetni = Err.Description
    application.EnableEvents = olay
    MsgBox "Arama ekraný: " & hataMetni, vbExclamation, "AKSORA"
End Sub


Private Sub Worksheet_Calculate()
    Dim son As Long, gorunen As Range, imza As String
    Dim hataMetni As String
    If AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Or AKSORA_V11_MESGUL Then Exit Sub
    If Not application.EnableEvents Or mOzetCalisiyor Then Exit Sub
    If Not (ActiveSheet Is Me) Then Exit Sub
    On Error GoTo Hata
    son = Me.Cells(Me.Rows.Count, "D").End(xlUp).Row
    imza = "BOS"
    If son >= 13 Then
        Set gorunen = AKSORA_V13_GORUNEN_SATIRLAR(Me, son)
        imza = CStr(son) & "|"
        If Not gorunen Is Nothing Then imza = imza & gorunen.Address
    End If
    ' Ayný görünür satýrlar için hesaplama olayýný yeniden çalýþtýrma.
    If imza = mSonGorunen Then Exit Sub
    mSonGorunen = imza
    mOzetCalisiyor = True
    AKSORA_V13_ARAMA_OZET Me
    mOzetCalisiyor = False
    Exit Sub
Hata:
    hataMetni = Err.Description
    mOzetCalisiyor = False
    mSonGorunen = ""
    MsgBox "Arama özeti güncellenemedi: " & hataMetni, vbExclamation, "AKSORA"
End Sub


