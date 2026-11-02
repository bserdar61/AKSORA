Attribute VB_Name = "Sayfa1"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
    If AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Or AKSORA_V11_MESGUL Then Exit Sub
    AKFIX_NORMAL_STOK_HAZIRLA Me
    AKFIX_EN_USTE Me
End Sub

Private Sub Worksheet_Calculate()
    AKFIX_FILTRE_KONTROL Me
End Sub


