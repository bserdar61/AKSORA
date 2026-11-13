Attribute VB_Name = "Sayfa11"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Option Explicit

Private Sub Worksheet_Activate()
    On Error GoTo Cikis

    With ThisWorkbook.Windows(1)
        .DisplayWorkbookTabs = True
        .DisplayGridlines = False
    End With

Cikis:
End Sub

