VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo Cikis

    With ThisWorkbook.Windows(1)
        .DisplayWorkbookTabs = True
        .DisplayGridlines = False
    End With

Cikis:
End Sub

