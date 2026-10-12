Attribute VB_Name = "modAKSORA_BASLANGIC"
Option Explicit

Private Function AKSORA_YENI_TEMA_SAYFA(ByVal ws As Worksheet) As Boolean
    Dim tag As String
    tag = CStr(ws.Range("BX1").Value2)
    AKSORA_YENI_TEMA_SAYFA = (tag = "AKSORA_MUSTERI_TEMA_V1" Or tag = "AKSORA_BIRIM_V1")
End Function

Public Sub AKSORA_HAFIF_KORUMALARI_YENILE()
    Dim ws As Worksheet, izin(1 To 11) As Boolean
    Dim cizim As Boolean, icerik As Boolean, senaryo As Boolean, secim As Long
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 1) <> "_" And Not AKSORA_YENI_TEMA_SAYFA(ws) Then
            cizim = ws.ProtectDrawingObjects: icerik = ws.ProtectContents
            senaryo = ws.ProtectScenarios: secim = ws.EnableSelection
            If cizim Or icerik Or senaryo Then
                With ws.Protection
                    izin(1) = .AllowFiltering: izin(2) = .AllowSorting
                    izin(3) = .AllowFormattingCells: izin(4) = .AllowFormattingColumns
                    izin(5) = .AllowFormattingRows: izin(6) = .AllowInsertingColumns
                    izin(7) = .AllowInsertingRows: izin(8) = .AllowInsertingHyperlinks
                    izin(9) = .AllowDeletingColumns: izin(10) = .AllowDeletingRows
                    izin(11) = .AllowUsingPivotTables
                End With
                ws.Unprotect Password:=""
                ws.Protect Password:="", DrawingObjects:=cizim, Contents:=icerik, _
                    Scenarios:=senaryo, UserInterfaceOnly:=True, _
                    AllowFiltering:=izin(1), AllowSorting:=izin(2), _
                    AllowFormattingCells:=izin(3), AllowFormattingColumns:=izin(4), _
                    AllowFormattingRows:=izin(5), AllowInsertingColumns:=izin(6), _
                    AllowInsertingRows:=izin(7), AllowInsertingHyperlinks:=izin(8), _
                    AllowDeletingColumns:=izin(9), AllowDeletingRows:=izin(10), _
                    AllowUsingPivotTables:=izin(11)
                ws.EnableSelection = secim
            End If
        End If
    Next ws
End Sub
