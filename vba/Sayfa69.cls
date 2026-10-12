VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa69"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Calculate()
    If AKSORA_V11_MESGUL Or AKSORA_V13_ACILIS Then Exit Sub
    On Error Resume Next
    AKSORA_MUSTERI_BANT Me
End Sub
