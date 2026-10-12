VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Activate()
    If AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Or AKSORA_V11_MESGUL Then Exit Sub
    Dim olay As Boolean
    olay = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    AKFIX_STOK_LISTE Me
    AKFIX_STOK_SECIM Me
Cikis:
    application.EnableEvents = olay
    Exit Sub
Hata:
    MsgBox "Depo listesi: " & Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Sub Worksheet_Change(ByVal Target As Range)
    Dim olay As Boolean
    If Intersect(Target, Me.Range("H2").MergeArea) Is Nothing Then Exit Sub
    olay = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    AKFIX_STOK_SECIM Me
Cikis:
    application.EnableEvents = olay
    Exit Sub
Hata:
    MsgBox "Depo filtresi: " & Err.Description, vbExclamation, "AKSORA"
    Resume Cikis
End Sub

Private Sub Worksheet_Calculate()
    AKFIX_FILTRE_KONTROL Me
End Sub


