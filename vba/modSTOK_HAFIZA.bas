Attribute VB_Name = "modSTOK_HAFIZA"
Option Explicit

Private Const KAYNAK_SAYFA As String = "STOKLAR"
Private Const SAKLAMA_GUN As Long = 180
Private mHafizaIl As String

Public Sub STOK_HAFIZA_KAYDET(Optional ByVal arsivIl As String = "")

    Dim wsKaynak As Worksheet
    Dim wsIndex As Worksheet
    Dim wsDelta As Worksheet
    Dim wsCurrent As Worksheet

    Dim yeni As Object
    Dim eski As Object
    Dim TumAnahtarlar As Object

    Dim SnapshotID As String
    Dim SnapshotZamani As Date

    Dim sonSatir As Long
    Dim i As Long
    Dim DeltaSatir As Long
    Dim IndexSatir As Long
    Dim DegisimSayisi As Long

    Dim anahtar As Variant
    Dim YeniSatir As Variant
    Dim EskiSatir As Variant
    Dim durum As String
    Dim deltaBuffer() As Variant, deltaCikti() As Variant
    Dim deltaKayit As Variant, deltaIlk As Long, c As Long, d As Long

    On Error GoTo Hata

    If Len(arsivIl) > 0 Then
        mHafizaIl = AKFIX_NORMAL(arsivIl)
        Set wsKaynak = ThisWorkbook.Worksheets("_AKLIVE_" & mHafizaIl & "_TUM")
    Else
        mHafizaIl = ""
        Set wsKaynak = ThisWorkbook.Worksheets(KAYNAK_SAYFA)
    End If

    STOK_HAFIZA_ALTYAPI_KUR

    Set wsIndex = ThisWorkbook.Worksheets(HAFIZA_INDEX)
    Set wsDelta = ThisWorkbook.Worksheets(HAFIZA_DELTA)
    Set wsCurrent = ThisWorkbook.Worksheets(HAFIZA_CURRENT)

    sonSatir = STOK_HAFIZA_SON_SATIR(wsKaynak, "A")

    If sonSatir < 7 Then
        Err.Raise vbObjectError + 7101, _
                  "STOK_HAFIZA_KAYDET", _
                  "STOKLAR sayfasýnda hafýzaya alýnacak veri bulunamadý."
    End If

    SnapshotZamani = Now
    SnapshotID = "STK-" & Format(SnapshotZamani, "yyyymmdd-hhnnss")

    Set yeni = STOK_HAFIZA_SOZLUK_OLUSTUR(wsKaynak, 7, sonSatir)

    If yeni.Count = 0 Then
        Err.Raise vbObjectError + 7102, _
                  "STOK_HAFIZA_KAYDET", _
                  "Yeni stok sözlüðü oluþturulamadý."
    End If

    If STOK_HAFIZA_CURRENT_VAR_MI(wsCurrent) Then

        Set eski = STOK_HAFIZA_SOZLUK_OLUSTUR( _
                    wsCurrent, _
                    2, _
                    STOK_HAFIZA_SON_SATIR(wsCurrent, "A"))

        Set TumAnahtarlar = CreateObject("Scripting.Dictionary")
        TumAnahtarlar.CompareMode = vbTextCompare

        For Each anahtar In eski.keys
            If Not TumAnahtarlar.Exists(CStr(anahtar)) Then
                TumAnahtarlar.Add CStr(anahtar), True
            End If
        Next anahtar

        For Each anahtar In yeni.keys
            If Not TumAnahtarlar.Exists(CStr(anahtar)) Then
                TumAnahtarlar.Add CStr(anahtar), True
            End If
        Next anahtar

        DeltaSatir = STOK_HAFIZA_SON_SATIR(wsDelta, "A") + 1
        If DeltaSatir < 2 Then DeltaSatir = 2
        deltaIlk = DeltaSatir
        ReDim deltaBuffer(1 To TumAnahtarlar.Count, 1 To 15)

        For Each anahtar In TumAnahtarlar.keys

            EskiSatir = Empty
            YeniSatir = Empty
            durum = vbNullString

            If eski.Exists(CStr(anahtar)) Then
                EskiSatir = eski(CStr(anahtar))
            End If

            If yeni.Exists(CStr(anahtar)) Then
                YeniSatir = yeni(CStr(anahtar))
            End If

            If Not eski.Exists(CStr(anahtar)) And yeni.Exists(CStr(anahtar)) Then

                durum = "YENÝ"

            ElseIf eski.Exists(CStr(anahtar)) And Not yeni.Exists(CStr(anahtar)) Then

                durum = "SIFIRLANDI"

            ElseIf eski.Exists(CStr(anahtar)) And yeni.Exists(CStr(anahtar)) Then

                If STOK_HAFIZA_DEGISIK_MI(EskiSatir, YeniSatir) Then

                    If STOK_HAFIZA_DOUBLE(YeniSatir(9)) > STOK_HAFIZA_DOUBLE(EskiSatir(9)) Then
                        durum = "ARTTI"
                    ElseIf STOK_HAFIZA_DOUBLE(YeniSatir(9)) < STOK_HAFIZA_DOUBLE(EskiSatir(9)) Then
                        durum = "AZALDI"
                    Else
                        durum = "DEÐÝÞTÝ"
                    End If

                End If

            End If

            If Len(durum) > 0 Then

                deltaKayit = STOK_HAFIZA_DELTA_DIZISI(SnapshotID, SnapshotZamani, _
                    CStr(anahtar), EskiSatir, YeniSatir, durum)
                For c = 1 To 15
                    deltaBuffer(DegisimSayisi + 1, c) = deltaKayit(c)
                Next c

                DeltaSatir = DeltaSatir + 1
                DegisimSayisi = DegisimSayisi + 1

            End If

        Next anahtar
        If DegisimSayisi > 0 Then
            ReDim deltaCikti(1 To DegisimSayisi, 1 To 15)
            For d = 1 To DegisimSayisi
                For c = 1 To 15: deltaCikti(d, c) = deltaBuffer(d, c): Next c
            Next d
            wsDelta.Cells(deltaIlk, 1).Resize(DegisimSayisi, 15).Value = deltaCikti
            wsDelta.Cells(deltaIlk, 2).Resize(DegisimSayisi, 1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
            wsDelta.Cells(deltaIlk, 12).Resize(DegisimSayisi, 4).NumberFormat = "#,##0.00"
        End If

    Else

        DegisimSayisi = 0

    End If

    STOK_HAFIZA_CURRENT_YAZ wsCurrent, wsKaynak, sonSatir, SnapshotID, SnapshotZamani

    IndexSatir = STOK_HAFIZA_SON_SATIR(wsIndex, "A") + 1
    If IndexSatir < 2 Then IndexSatir = 2

    wsIndex.Cells(IndexSatir, 1).Value = SnapshotID
    wsIndex.Cells(IndexSatir, 2).Value = SnapshotZamani
    wsIndex.Cells(IndexSatir, 3).Value = yeni.Count
    wsIndex.Cells(IndexSatir, 4).Value = DegisimSayisi
    If Len(mHafizaIl) > 0 Then
        wsIndex.Cells(IndexSatir, 5).Value = wsKaynak.Range("A5").Value2
    Else
        wsIndex.Cells(IndexSatir, 5).Value = wsKaynak.Range("C4").Value2
    End If

    wsIndex.Cells(IndexSatir, 2).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsIndex.Cells(IndexSatir, 5).NumberFormat = "#,##0.00"

    STOK_HAFIZA_ESKI_KAYITLARI_TEMIZLE
    mHafizaIl = ""

    Exit Sub

Hata:
    mHafizaIl = ""

    Err.Raise Err.Number, _
              "STOK_HAFIZA_KAYDET", _
              Err.Description

End Sub


Public Sub STOK_HAFIZA_TEMIZ_BASLAT()

    Dim wsKaynak As Worksheet
    Dim wsIndex As Worksheet
    Dim wsDelta As Worksheet
    Dim wsCurrent As Worksheet

    Dim sonSatir As Long
    Dim SnapshotID As String
    Dim SnapshotZamani As Date
    Dim IndexSatir As Long
    Dim SatirSayisi As Long

    On Error GoTo Hata

    application.ScreenUpdating = False
    application.EnableEvents = False

    Set wsKaynak = ThisWorkbook.Worksheets(KAYNAK_SAYFA)

    STOK_HAFIZA_ALTYAPI_KUR

    Set wsIndex = ThisWorkbook.Worksheets(HAFIZA_INDEX)
    Set wsDelta = ThisWorkbook.Worksheets(HAFIZA_DELTA)
    Set wsCurrent = ThisWorkbook.Worksheets(HAFIZA_CURRENT)

    sonSatir = STOK_HAFIZA_SON_SATIR(wsKaynak, "A")

    If sonSatir < 7 Then
        Err.Raise vbObjectError + 7191, _
                  "STOK_HAFIZA_TEMIZ_BASLAT", _
                  "STOKLAR sayfasýnda baþlangýç stoku bulunamadý."
    End If

    wsIndex.Rows("2:" & wsIndex.Rows.Count).ClearContents
    wsDelta.Rows("2:" & wsDelta.Rows.Count).ClearContents
    wsCurrent.Range("A2:O" & wsCurrent.Rows.Count).ClearContents

    SnapshotZamani = Now
    SnapshotID = "STK-" & Format(SnapshotZamani, "yyyymmdd-hhnnss")

    SatirSayisi = sonSatir - 6

    wsCurrent.Range("A2").Resize(SatirSayisi, 12).Value = _
        wsKaynak.Range("A7:L" & sonSatir).Value2

    wsCurrent.Range("N2").Value = SnapshotID
    wsCurrent.Range("O2").Value = SnapshotZamani
    wsCurrent.Range("O2").NumberFormat = "dd.mm.yyyy hh:mm:ss"

    IndexSatir = 2

    wsIndex.Cells(IndexSatir, 1).Value = SnapshotID
    wsIndex.Cells(IndexSatir, 2).Value = SnapshotZamani
    wsIndex.Cells(IndexSatir, 3).Value = _
        STOK_HAFIZA_SOZLUK_OLUSTUR(wsKaynak, 7, sonSatir).Count
    wsIndex.Cells(IndexSatir, 4).Value = 0
    wsIndex.Cells(IndexSatir, 5).Value = _
        STOK_HAFIZA_DOUBLE(wsKaynak.Range("C4").Value2)

    wsIndex.Cells(IndexSatir, 2).NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsIndex.Cells(IndexSatir, 5).NumberFormat = "#,##0.00"

    wsIndex.Visible = xlSheetVeryHidden
    wsDelta.Visible = xlSheetVeryHidden
    wsCurrent.Visible = xlSheetVeryHidden

    application.EnableEvents = True
    application.ScreenUpdating = True

    MsgBox _
        "STOK HAFIZA temiz baþlangýç tamamlandý." & vbCrLf & vbCrLf & _
        "Baþlangýç stok satýrý : " & _
        STOK_HAFIZA_SOZLUK_OLUSTUR(wsKaynak, 7, sonSatir).Count & vbCrLf & _
        "Baþlangýç tarihi      : " & _
        Format(SnapshotZamani, "dd.mm.yyyy hh:mm:ss"), _
        vbInformation, _
        "AKSORA - STOK HAFIZA"

    Exit Sub

Hata:

    application.EnableEvents = True
    application.ScreenUpdating = True

    MsgBox _
        "Temiz baþlangýç yapýlamadý." & vbCrLf & vbCrLf & _
        Err.Description, _
        vbCritical, _
        "AKSORA - STOK HAFIZA"

End Sub

Public Sub STOK_HAFIZA_ALTYAPI_KUR()

    Dim wsIndex As Worksheet
    Dim wsDelta As Worksheet
    Dim wsCurrent As Worksheet

    Set wsIndex = STOK_HAFIZA_SAYFA_AL(HAFIZA_INDEX)
    Set wsDelta = STOK_HAFIZA_SAYFA_AL(HAFIZA_DELTA)
    Set wsCurrent = STOK_HAFIZA_SAYFA_AL(HAFIZA_CURRENT)

    With wsIndex
        .Cells(1, 1).Value = "SNAPSHOT_ID"
        .Cells(1, 2).Value = "TARÝH"
        .Cells(1, 3).Value = "STOK_SATIRI"
        .Cells(1, 4).Value = "DEÐÝÞÝM_SAYISI"
        .Cells(1, 5).Value = "TOPLAM_DEÐER"
        .Rows(1).Font.bold = True
        .Visible = xlSheetVeryHidden
    End With

    With wsDelta
        .Cells(1, 1).Value = "SNAPSHOT_ID"
        .Cells(1, 2).Value = "TARÝH"
        .Cells(1, 3).Value = "DURUM"
        .Cells(1, 4).Value = "ANAHTAR"
        .Cells(1, 5).Value = "ÜY"
        .Cells(1, 6).Value = "DEPO"
        .Cells(1, 7).Value = "DEPO_TANIMI"
        .Cells(1, 8).Value = "MALZEME"
        .Cells(1, 9).Value = "MALZEME_AÇIKLAMASI"
        .Cells(1, 10).Value = "PARTÝ"
        .Cells(1, 11).Value = "BÝRÝM"
        .Cells(1, 12).Value = "ESKÝ_MÝKTAR"
        .Cells(1, 13).Value = "YENÝ_MÝKTAR"
        .Cells(1, 14).Value = "ESKÝ_DEÐER"
        .Cells(1, 15).Value = "YENÝ_DEÐER"
        .Rows(1).Font.bold = True
        .Visible = xlSheetVeryHidden
    End With

    With wsCurrent
        .Cells(1, 1).Value = "ÜY"
        .Cells(1, 2).Value = "DpYr"
        .Cells(1, 3).Value = "Depo yeri tanýmý"
        .Cells(1, 4).Value = "Mal grubu"
        .Cells(1, 5).Value = "Mal grubu tanýmý"
        .Cells(1, 6).Value = "Parti"
        .Cells(1, 7).Value = "Malzeme"
        .Cells(1, 8).Value = "Malzeme kýsa metni"
        .Cells(1, 9).Value = "Tahditsiz"
        .Cells(1, 10).Value = "Birim"
        .Cells(1, 11).Value = "Deðer"
        .Cells(1, 12).Value = "Para birimi"
        .Cells(1, 14).Value = "SNAPSHOT_ID"
        .Cells(1, 15).Value = "TARÝH"
        .Rows(1).Font.bold = True
        .Visible = xlSheetVeryHidden
    End With

End Sub

Private Function STOK_HAFIZA_SAYFA_AL(ByVal sayfaAdi As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sayfaAdi)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
                    after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sayfaAdi
    End If

    Set STOK_HAFIZA_SAYFA_AL = ws

End Function

Private Function STOK_HAFIZA_CURRENT_VAR_MI(ByVal ws As Worksheet) As Boolean

    STOK_HAFIZA_CURRENT_VAR_MI = _
        (STOK_HAFIZA_SON_SATIR(ws, "A") >= 2)

End Function

Private Function STOK_HAFIZA_SOZLUK_OLUSTUR( _
ByVal ws As Worksheet, _
ByVal IlkSatir As Long, _
ByVal sonSatir As Long) As Object

    Dim d As Object
    Dim veri As Variant, satir As Variant, mevcut As Variant
    Dim i As Long, anahtar As String

    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare

    If sonSatir < IlkSatir Then
        Set STOK_HAFIZA_SOZLUK_OLUSTUR = d
        Exit Function
    End If

    veri = ws.Range("A" & IlkSatir & ":L" & sonSatir).Value2

    For i = 1 To UBound(veri, 1)

        anahtar = STOK_HAFIZA_ANAHTAR( _
            veri(i, 1), veri(i, 2), veri(i, 7), veri(i, 6))

        If Len(anahtar) > 0 Then
            satir = STOK_HAFIZA_SATIR_DIZISI(veri, i)

            If d.Exists(anahtar) Then
                mevcut = d(anahtar)

                If StrComp(Trim$(CStr(mevcut(10))), _
                           Trim$(CStr(satir(10))), vbTextCompare) <> 0 Then
                    Err.Raise vbObjectError + 7110, _
                        "STOK_HAFIZA_SOZLUK_OLUSTUR", _
                        "Ayni stok anahtarinda farkli birimler: " & anahtar
                End If

                mevcut(9) = STOK_HAFIZA_DOUBLE(mevcut(9)) + _
                            STOK_HAFIZA_DOUBLE(satir(9))

                mevcut(11) = CDec(mevcut(11)) + CDec(satir(11))

                If mevcut(9) <> 0 Then
                    mevcut(12) = mevcut(11) / mevcut(9)
                Else
                    mevcut(12) = 0
                End If

                d(anahtar) = mevcut
            Else
                d.Add anahtar, satir
            End If
        End If
    Next i

    Set STOK_HAFIZA_SOZLUK_OLUSTUR = d

End Function

Private Function STOK_HAFIZA_SATIR_DIZISI( _
ByRef veri As Variant, _
ByVal satir As Long) As Variant

    Dim x(1 To 12) As Variant
    Dim j As Long

    For j = 1 To 12
        x(j) = veri(satir, j)
    Next j

    STOK_HAFIZA_SATIR_DIZISI = x

End Function

Private Function STOK_HAFIZA_ANAHTAR( _
ByVal uy As Variant, _
ByVal depo As Variant, _
ByVal malzeme As Variant, _
ByVal parti As Variant) As String

    Dim sUY As String
    Dim sDepo As String
    Dim sMalzeme As String
    Dim sParti As String

    sUY = Trim$(CStr(uy))
    sDepo = UCase$(Trim$(CStr(depo)))
    sMalzeme = Trim$(CStr(malzeme))
    sParti = UCase$(Trim$(CStr(parti)))

    If Len(sUY) = 0 Or Len(sDepo) = 0 Or Len(sMalzeme) = 0 Then
        STOK_HAFIZA_ANAHTAR = vbNullString
    Else
        STOK_HAFIZA_ANAHTAR = _
            sUY & "|" & sDepo & "|" & sMalzeme & "|" & sParti
    End If

End Function

Private Function STOK_HAFIZA_DEGISIK_MI( _
ByVal EskiSatir As Variant, _
ByVal YeniSatir As Variant) As Boolean

    If Abs(STOK_HAFIZA_DOUBLE(EskiSatir(9)) - STOK_HAFIZA_DOUBLE(YeniSatir(9))) > 0.0000001 Then
        STOK_HAFIZA_DEGISIK_MI = True
        Exit Function
    End If

    If CDec(STOK_HAFIZA_DOUBLE(EskiSatir(11))) <> CDec(STOK_HAFIZA_DOUBLE(YeniSatir(11))) Then
        STOK_HAFIZA_DEGISIK_MI = True
        Exit Function
    End If

    If StrComp(Trim$(CStr(EskiSatir(7))), Trim$(CStr(YeniSatir(7))), vbTextCompare) <> 0 Then
        STOK_HAFIZA_DEGISIK_MI = True
        Exit Function
    End If

End Function

Private Function STOK_HAFIZA_DELTA_DIZISI( _
ByVal SnapshotID As String, _
ByVal SnapshotZamani As Date, _
ByVal anahtar As String, _
ByVal EskiSatir As Variant, _
ByVal YeniSatir As Variant, _
ByVal durum As String) As Variant
    Dim kayit As Variant, x(1 To 15) As Variant
    If IsEmpty(YeniSatir) Then kayit = EskiSatir Else kayit = YeniSatir
    x(1) = SnapshotID: x(2) = SnapshotZamani: x(3) = durum: x(4) = anahtar
    x(5) = kayit(1): x(6) = kayit(2): x(7) = kayit(3)
    x(8) = kayit(7): x(9) = kayit(8): x(10) = kayit(6): x(11) = kayit(10)
    x(12) = 0: x(13) = 0: x(14) = 0: x(15) = 0
    If Not IsEmpty(EskiSatir) Then
        x(12) = STOK_HAFIZA_DOUBLE(EskiSatir(9))
        x(14) = STOK_HAFIZA_DOUBLE(EskiSatir(11))
    End If
    If Not IsEmpty(YeniSatir) Then
        x(13) = STOK_HAFIZA_DOUBLE(YeniSatir(9))
        x(15) = STOK_HAFIZA_DOUBLE(YeniSatir(11))
    End If
    STOK_HAFIZA_DELTA_DIZISI = x
End Function

Private Sub STOK_HAFIZA_CURRENT_YAZ( _
ByVal wsCurrent As Worksheet, _
ByVal wsKaynak As Worksheet, _
ByVal sonSatir As Long, _
ByVal SnapshotID As String, _
ByVal SnapshotZamani As Date)

    Dim SatirSayisi As Long
    Dim sonHucre As Range

    ' Clear all previous content in A:L, including stray cells and formulas.
    ' Formatting and N2/O2 metadata are preserved as in the existing routine.
    Set sonHucre = wsCurrent.Range("A:L").Find(What:="*", _
        after:=wsCurrent.Range("A1"), LookIn:=xlFormulas, LookAt:=xlPart, _
        SearchOrder:=xlByRows, SearchDirection:=xlPrevious, _
        MatchCase:=False, SearchFormat:=False)
    If Not sonHucre Is Nothing Then
        If sonHucre.Row >= 2 Then
            wsCurrent.Range("A2:L" & sonHucre.Row).ClearContents
        End If
    End If

    SatirSayisi = sonSatir - 6

    If SatirSayisi > 0 Then
        wsCurrent.Range("A2").Resize(SatirSayisi, 12).Value = _
            wsKaynak.Range("A7:L" & sonSatir).Value2
    End If

    wsCurrent.Range("N2").Value = SnapshotID
    wsCurrent.Range("O2").Value = SnapshotZamani
    wsCurrent.Range("O2").NumberFormat = "dd.mm.yyyy hh:mm:ss"

    wsCurrent.Visible = xlSheetVeryHidden

End Sub

Private Sub STOK_HAFIZA_ESKI_KAYITLARI_TEMIZLE()
    Dim wsIndex As Worksheet, wsDelta As Worksheet, KesimTarihi As Date
    KesimTarihi = DateAdd("d", -SAKLAMA_GUN, Now)
    Set wsIndex = ThisWorkbook.Worksheets(HAFIZA_INDEX)
    Set wsDelta = ThisWorkbook.Worksheets(HAFIZA_DELTA)
    STOK_HAFIZA_ESKI_SATIR_BLOKLARI wsDelta, KesimTarihi
    STOK_HAFIZA_ESKI_SATIR_BLOKLARI wsIndex, KesimTarihi
    wsIndex.Visible = xlSheetVeryHidden
    wsDelta.Visible = xlSheetVeryHidden
End Sub

Private Sub STOK_HAFIZA_ESKI_SATIR_BLOKLARI(ByVal ws As Worksheet, ByVal kesim As Date)
    Dim son As Long, i As Long, BlokSon As Long, eski As Boolean, tarihler As Variant
    son = STOK_HAFIZA_SON_SATIR(ws, "A")
    If son < 2 Then Exit Sub
    ' Value, not Value2: date-typed variants match the old Cells(...).Value check.
    tarihler = ws.Range("A1:B" & son).Value
    For i = son To 1 Step -1
        eski = False
        If i >= 2 Then
            If IsDate(tarihler(i, 2)) Then eski = (CDate(tarihler(i, 2)) < kesim)
        End If
        If eski Then
            If BlokSon = 0 Then BlokSon = i
        ElseIf BlokSon > 0 Then
            ws.Rows(CStr(i + 1) & ":" & CStr(BlokSon)).Delete
            BlokSon = 0
        End If
    Next i
End Sub

Private Function STOK_HAFIZA_SON_SATIR( _
ByVal ws As Worksheet, _
ByVal sutun As String) As Long

    Dim son As Long

    son = ws.Cells(ws.Rows.Count, sutun).End(xlUp).Row

    If son = 1 And Len(Trim$(CStr(ws.Cells(1, sutun).Value2))) = 0 Then
        STOK_HAFIZA_SON_SATIR = 0
    Else
        STOK_HAFIZA_SON_SATIR = son
    End If

End Function

Private Function STOK_HAFIZA_DOUBLE(ByVal deger As Variant) As Double

    If IsNumeric(deger) Then
        STOK_HAFIZA_DOUBLE = CDbl(deger)
    Else
        STOK_HAFIZA_DOUBLE = 0
    End If

End Function
Private Function HAFIZA_INDEX() As String
    If Len(mHafizaIl) > 0 Then
        HAFIZA_INDEX = AKFIX_HAFIZA_ADI(mHafizaIl, "INDEX")
    Else
        HAFIZA_INDEX = AKSORA_HAFIZA_ADI("INDEX")
    End If
End Function

Private Function HAFIZA_DELTA() As String
    If Len(mHafizaIl) > 0 Then
        HAFIZA_DELTA = AKFIX_HAFIZA_ADI(mHafizaIl, "DELTA")
    Else
        HAFIZA_DELTA = AKSORA_HAFIZA_ADI("DELTA")
    End If
End Function

Private Function HAFIZA_CURRENT() As String
    If Len(mHafizaIl) > 0 Then
        HAFIZA_CURRENT = AKFIX_HAFIZA_ADI(mHafizaIl, "CURRENT")
    Else
        HAFIZA_CURRENT = AKSORA_HAFIZA_ADI("CURRENT")
    End If
End Function

