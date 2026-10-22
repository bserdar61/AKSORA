Attribute VB_Name = "modMALZEME_ARA"
Option Explicit

'==========================================================
' AKSORA
' modMALZEME_ARA
'
' MALZEME ARAMA MERKEZÝ
'
' - Normal + Ýlçe + Yüklenici
' - Kod / açýklama arama
' - Enter ile arama
' - Filtreye göre canlý KPI
' - ANA PANEL butonu O5
'==========================================================

Private Const ARA_SAYFA As String = "MALZEME_ARA"
Private Const BASLIK_SATIRI As Long = 12
Private Const ILK_SONUC_SATIRI As Long = 13

'==========================================================
' SAYFAYI KUR
'==========================================================
Public Sub MALZEME_ARA_KUR()

    Dim ws As Worksheet

    On Error GoTo Hata

    application.ScreenUpdating = False
    application.EnableEvents = False

    Set ws = MALZEME_ARA_SAYFA_AL()

    If ws.Visible <> xlSheetVisible Then
        ws.Visible = xlSheetVisible
    End If

    MALZEME_ARA_TEMIZLE ws
    MALZEME_ARA_ZEMIN ws
    MALZEME_ARA_BASLIK ws
    MALZEME_ARA_GIRIS_ALANI ws
    MALZEME_ARA_KPI_ALANI ws
    MALZEME_ARA_TABLO_BASLIK ws
    MALZEME_ARA_BUTONLAR ws

    ws.Activate

    ActiveWindow.DisplayGridlines = False
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1

    ws.Range("B6").Select

Cikis:
    On Error Resume Next
    AKSORA_V11_BICIM_SAYFA ThisWorkbook.Worksheets("MALZEME_ARA")
    On Error GoTo 0

    application.EnableEvents = True
    application.ScreenUpdating = True

    Exit Sub

Hata:

    MsgBox _
        "MALZEME ARAMA MERKEZÝ oluþturulamadý." & _
        vbCrLf & vbCrLf & _
        "Hata No: " & Err.Number & vbCrLf & _
        "Açýklama: " & Err.Description, _
        vbCritical, _
        "AKSORA"

    Resume Cikis

End Sub

'==========================================================
' MALZEME ARA
'==========================================================
Public Sub MALZEME_ARA()
    AKSORA_IL_ARAMA_YAP
End Sub

'==========================================================
' KAYNAK TARA
'==========================================================
Public Sub MALZEME_ARA_SONUCLARI_TEMIZLE( _
Optional ByVal ws As Worksheet)

    Dim sonSatir As Long

    If ws Is Nothing Then

        If Not MALZEME_ARA_SAYFA_VAR_MI(ARA_SAYFA) Then
            Exit Sub
        End If

        Set ws = _
            ThisWorkbook.Worksheets(ARA_SAYFA)

    End If

    '------------------------------------------------------
    ' FÝLTRELERÝ TEMÝZLE
    '------------------------------------------------------
    On Error Resume Next

    If ws.FilterMode Then
        ws.ShowAllData
    End If

    On Error GoTo 0

    '------------------------------------------------------
    ' "BULUNAMADI" MERGE'Ý VARSA KALDIR
    '------------------------------------------------------
    If ws.Range("A13").MergeCells Then

        ws.Range("A13").MergeArea.UnMerge

    End If

    '------------------------------------------------------
    ' SON SATIR
    '------------------------------------------------------
    sonSatir = _
        AKSORA_SON_SATIR(ws, "A")

    If sonSatir < ILK_SONUC_SATIRI Then
        sonSatir = ILK_SONUC_SATIRI
    End If

    '------------------------------------------------------
    ' ESKÝ SONUÇLAR
    '------------------------------------------------------
    ws.Range( _
        "A" & ILK_SONUC_SATIRI & _
        ":J" & sonSatir).ClearContents

    With ws.Range( _
        "A" & ILK_SONUC_SATIRI & _
        ":J2000")

        .Interior.Color = RGB(255, 255, 255)

    End With

    '------------------------------------------------------
    ' KPI MERGE ALANLARINI GÜVENLÝ TEMÝZLE
    '------------------------------------------------------
    ws.Range("A10").MergeArea.ClearContents
    ws.Range("C10").MergeArea.ClearContents
    ws.Range("F10").MergeArea.ClearContents
    ws.Range("I10").MergeArea.ClearContents

End Sub

'==========================================================
' SAYFAYI AL
'==========================================================
Private Function MALZEME_ARA_SAYFA_AL() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(ARA_SAYFA)

    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = _
            ThisWorkbook.Worksheets.Add( _
                after:=ThisWorkbook.Worksheets( _
                    ThisWorkbook.Worksheets.Count))

        ws.Name = ARA_SAYFA

    End If

    Set MALZEME_ARA_SAYFA_AL = ws

End Function

'==========================================================
' SAYFA VAR MI?
'==========================================================
Private Function MALZEME_ARA_SAYFA_VAR_MI( _
ByVal sayfaAdi As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next

    Set ws = _
        ThisWorkbook.Worksheets(sayfaAdi)

    MALZEME_ARA_SAYFA_VAR_MI = _
        Not ws Is Nothing

    On Error GoTo 0

End Function

'==========================================================
' SAYFAYI TEMÝZLE
'==========================================================
Private Sub MALZEME_ARA_TEMIZLE( _
ByVal ws As Worksheet)

    Dim i As Long

    On Error Resume Next

    If ws.FilterMode Then
        ws.ShowAllData
    End If

    ws.AutoFilterMode = False

    On Error GoTo 0

    ws.Cells.UnMerge
    ws.Cells.Clear

    On Error Resume Next

    For i = ws.Shapes.Count To 1 Step -1

        ws.Shapes(i).Delete

    Next i

    On Error GoTo 0

End Sub

'==========================================================
' ZEMÝN
'==========================================================
Private Sub MALZEME_ARA_ZEMIN( _
ByVal ws As Worksheet)

    With ws

        .Cells.Font.Name = "Aptos"
        .Cells.Font.Size = 10

        .Range("A1:J2000").Interior.Color = _
            RGB(245, 247, 250)

        .Columns("A").ColumnWidth = 10
        .Columns("B").ColumnWidth = 11
        .Columns("C").ColumnWidth = 25
        .Columns("D").ColumnWidth = 17
        .Columns("E").ColumnWidth = 40
        .Columns("F").ColumnWidth = 14
        .Columns("G").ColumnWidth = 14
        .Columns("H").ColumnWidth = 10
        .Columns("I").ColumnWidth = 18
        .Columns("J").ColumnWidth = 15

        .Columns("O").ColumnWidth = 11

        .Rows("1:3").RowHeight = 25
        .Rows("4:11").RowHeight = 24
        .Rows("12").RowHeight = 30

    End With

End Sub

'==========================================================
' BAÞLIK
'==========================================================
Private Sub MALZEME_ARA_BASLIK( _
ByVal ws As Worksheet)

    With ws.Range("A1:J2")

        .Merge

        .Value = _
            "MALZEME ARAMA MERKEZÝ"

        .Interior.Color = _
            RGB(20, 55, 95)

        .Font.Color = _
            RGB(255, 255, 255)

        .Font.bold = True
        .Font.Size = 20

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Color = RGB(0, 0, 0)

    End With

    With ws.Range("A3:J3")

        .Merge

        .Value = _
            "Normal + Ýlçe + Yüklenici Stoklarý"

        .Interior.Color = _
            RGB(20, 55, 95)

        .Font.Color = _
            RGB(215, 225, 238)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub

'==========================================================
' ARAMA GÝRÝÞ ALANI
'==========================================================
Private Sub MALZEME_ARA_GIRIS_ALANI( _
ByVal ws As Worksheet)

    With ws.Range("A5:J5")

        .Merge

        .Value = _
            "MALZEME KODU VEYA MALZEME AÇIKLAMASI"

        .Font.bold = True
        .Font.Color = RGB(75, 85, 95)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    With ws.Range("B6:H6")

        .Merge

        .Interior.Color = RGB(255, 255, 255)

        .Borders.LineStyle = xlContinuous
        .Borders.Color = RGB(185, 195, 205)

        .Font.bold = True
        .Font.Size = 11

        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter

    End With

    With ws.Range("A7:J7")

        .Merge

        .Value = _
            "Malzeme kodunun tamamýný, bir bölümünü veya açýklamadan bir kelime yazabilirsiniz."

        .Font.Size = 8
        .Font.Color = RGB(125, 135, 145)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub

'==========================================================
' KPI ALANI
'==========================================================
Private Sub MALZEME_ARA_KPI_ALANI( _
ByVal ws As Worksheet)

    MALZEME_ARA_KPI _
        ws, _
        "A9:B9", _
        "A10:B10", _
        "BULUNAN SATIR"

    MALZEME_ARA_KPI _
        ws, _
        "C9:E9", _
        "C10:E10", _
        "BULUNAN DEPO"

    MALZEME_ARA_KPI _
        ws, _
        "F9:H9", _
        "F10:H10", _
        "TOPLAM MÝKTAR"

    MALZEME_ARA_KPI _
        ws, _
        "I9:J9", _
        "I10:J10", _
        "TOPLAM DEÐER"

End Sub

'==========================================================
' KPI KUTUSU
'==========================================================
Private Sub MALZEME_ARA_KPI( _
ByVal ws As Worksheet, _
ByVal BaslikAlan As String, _
ByVal DegerAlan As String, _
ByVal baslik As String)

    With ws.Range(BaslikAlan)

        .Merge
        .Value = baslik

        .Interior.Color = RGB(225, 230, 236)

        .Font.bold = True
        .Font.Size = 8
        .Font.Color = RGB(75, 85, 95)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

    With ws.Range(DegerAlan)

        .Merge
        .Value = 0

        .Interior.Color = RGB(255, 255, 255)

        .Font.bold = True
        .Font.Size = 12
        .Font.Color = RGB(20, 55, 95)

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

    End With

End Sub

'==========================================================
' TABLO BAÞLIKLARI
'==========================================================
Private Sub MALZEME_ARA_TABLO_BASLIK( _
ByVal ws As Worksheet)

    Dim basliklar As Variant
    Dim i As Long

    basliklar = Array( _
        "ÜY", _
        "DEPO", _
        "DEPO TANIMI", _
        "MALZEME", _
        "MALZEME AÇIKLAMASI", _
        "PARTÝ", _
        "MÝKTAR", _
        "BÝRÝM", _
        "DEÐER", _
        "KAYNAK")

    For i = 0 To UBound(basliklar)

        ws.Cells(12, i + 1).Value = _
            basliklar(i)

    Next i

    With ws.Range("A12:J12")

        .Interior.Color = RGB(20, 55, 95)

        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True

        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter

        .AutoFilter

    End With

End Sub

'==========================================================
' BUTONLAR
'==========================================================
Private Sub MALZEME_ARA_BUTONLAR( _
ByVal ws As Worksheet)

    Dim hedef As Range
    Dim shp As Shape

    '------------------------------------------------------
    ' ARA
    '------------------------------------------------------
    Set hedef = ws.Range("I6:J6")

    Set shp = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            hedef.Left, _
            hedef.Top, _
            hedef.Width, _
            hedef.Height)

    With shp

        .Name = "btnMalzemeAra"

        .TextFrame2.TextRange.Text = "ARA"

        .Fill.ForeColor.RGB = RGB(45, 95, 145)
        .line.ForeColor.RGB = RGB(45, 95, 145)

        .TextFrame2.TextRange.Font.Name = "Aptos"
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.bold = msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.MarginLeft = 1
        .TextFrame2.MarginRight = 1
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        .OnAction = _
            "'" & ThisWorkbook.Name & "'!MALZEME_ARA"

        .Placement = xlMoveAndSize

    End With

    '------------------------------------------------------
    ' ANA PANEL - O5
    '------------------------------------------------------
    Set hedef = ws.Range("O5")

    Set shp = _
        ws.Shapes.AddShape( _
            msoShapeRoundedRectangle, _
            hedef.Left, _
            hedef.Top, _
            hedef.Width, _
            hedef.Height)

    With shp

        .Name = "btnMalzemeAnaPanel"

        .TextFrame2.TextRange.Text = "ANA PANEL"

        .Fill.ForeColor.RGB = _
            RGB(20, 55, 95)

        .line.ForeColor.RGB = _
            RGB(20, 55, 95)

        .TextFrame2.TextRange.Font.Name = _
            "Aptos"

        .TextFrame2.TextRange.Font.Size = 7
        .TextFrame2.TextRange.Font.bold = msoTrue

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(255, 255, 255)

        .TextFrame2.VerticalAnchor = _
            msoAnchorMiddle

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        .OnAction = _
            "'" & ThisWorkbook.Name & "'!ANA_PANELE_DON"

        .Placement = xlMoveAndSize

    End With

End Sub

'==========================================================
' SONUÇ TABLOSU FORMAT
'==========================================================
