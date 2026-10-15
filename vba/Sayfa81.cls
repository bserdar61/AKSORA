Attribute VB_Name = "Sayfa81"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Calculate()
    ' Filtre sonrasi istenmeyen sayfa gecisine karsi: son hesaplanan stok sayfasi ve zamani.
    If ActiveSheet Is Me Then AKSORA_MUSTERI_STOK_HESAP_SAYFA = Me.Name: AKSORA_MUSTERI_STOK_HESAP_ZAMAN = Timer
    If AKSORA_V11_MESGUL Or AKSORA_V13_ACILIS Then Exit Sub
    On Error Resume Next
    AKSORA_MUSTERI_BANT Me
End Sub
