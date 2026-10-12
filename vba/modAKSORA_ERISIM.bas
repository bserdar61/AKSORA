Attribute VB_Name = "modAKSORA_ERISIM"
Option Explicit

Private Const BTN_AD As String = "btnAnaPanelDon"



Public Sub ANA_PANELE_DON()
    On Error GoTo Hata
    AKSORA_V13_GIT "ANA PANEL"
    Exit Sub
Hata:
    MsgBox "Ana panel açýlamadý." & vbCrLf & Err.Description, vbExclamation, "AKSORA"
End Sub


Public Sub AKSORA_ANA_PANEL_YAZILARINI_ESITLE()

    Dim ws As Worksheet, sh As Shape
    Dim metin As String
    Dim koruma As Boolean, icerik As Boolean, senaryo As Boolean
    Dim filtre As Boolean, sirala As Boolean
    Dim hataNo As Long, hataMetni As String

    On Error GoTo Hata

    For Each ws In ThisWorkbook.Worksheets

        koruma = False

        For Each sh In ws.Shapes

            metin = vbNullString

            On Error Resume Next
            metin = UCase$(Trim$(sh.TextFrame2.TextRange.Text))
            On Error GoTo Hata

            If metin = "ANA" Then

                If Not koruma Then
                    If ws.ProtectContents Or ws.ProtectDrawingObjects Or _
                       ws.ProtectScenarios Then

                        icerik = ws.ProtectContents
                        senaryo = ws.ProtectScenarios
                        filtre = ws.Protection.AllowFiltering
                        sirala = ws.Protection.AllowSorting

                        ws.Unprotect Password:=""
                        koruma = True
                    End If
                End If

                With sh
                    .TextFrame2.TextRange.Text = "ANA PANEL"
                    .TextFrame2.TextRange.Font.Size = 8
                    .TextFrame2.TextRange.Font.bold = msoTrue
                    .TextFrame2.TextRange.ParagraphFormat.Alignment = _
                        msoAlignCenter

                    If .Width < 76 Then .Width = 76
                End With

            End If

        Next sh

        If koruma Then
            ws.Protect Password:="", DrawingObjects:=True, _
                Contents:=icerik, Scenarios:=senaryo, _
                UserInterfaceOnly:=True, _
                AllowFiltering:=filtre, AllowSorting:=sirala
        End If

    Next ws

    Exit Sub

Hata:
    hataNo = Err.Number
    hataMetni = Err.Description

    On Error Resume Next
    If koruma Then
        ws.Protect Password:="", DrawingObjects:=True, _
            Contents:=icerik, Scenarios:=senaryo, _
            UserInterfaceOnly:=True, _
            AllowFiltering:=filtre, AllowSorting:=sirala
    End If
    On Error GoTo 0

    Err.Raise hataNo, "AKSORA_ANA_PANEL_YAZILARINI_ESITLE", hataMetni

End Sub
Public Sub STOK_KARS_PANEL_AC()

    On Error GoTo Hata

    AKSORA_V12_GIT "STOK_KARÞILAÞTIRMA"

    Exit Sub

Hata:

    MsgBox _
        "Karþýlaþtýrma ekraný açýlamadý: " & Err.Description, _
        vbExclamation, _
        "AKSORA"

End Sub

Public Sub PANEL_GIT_STOK_KARSILASTIRMA()

    STOK_KARS_PANEL_AC

End Sub


Public Sub PANEL_GIT_MALZEME_ARA()

    On Error GoTo Hata

    AKSORA_V12_GIT "MALZEME_ARA"

    Exit Sub

Hata:

    MsgBox _
        "Arama ekraný açýlamadý: " & Err.Description, _
        vbExclamation, _
        "AKSORA"

End Sub
