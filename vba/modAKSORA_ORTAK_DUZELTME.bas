Attribute VB_Name = "modAKSORA_ORTAK_DUZELTME"
Option Explicit

Private mImza As Object
Private mListeIl As Object
Private mKarsIller As Object
Public AKFIX_KARS_GECMIS_FARK As Double

Public Function AKFIX_NORMAL(ByVal metin As String) As String
    metin = UCase$(Trim$(metin))
    metin = Replace(metin, "Ý", "I"): metin = Replace(metin, "Ü", "U")
    metin = Replace(metin, "Þ", "S"): metin = Replace(metin, "Ð", "G")
    metin = Replace(metin, "Ç", "C"): metin = Replace(metin, "Ö", "O")
    AKFIX_NORMAL = metin
End Function

Public Function AKFIX_HAFIZA_ADI(ByVal il As String, ByVal tur As String) As String
    il = AKFIX_NORMAL(il)
    Select Case il
        Case "ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN"
        Case Else: Err.Raise vbObjectError + 900, , "Geçersiz geçmiþ ili: " & il
    End Select
    Select Case tur
        Case "INDEX", "DELTA", "CURRENT"
        Case Else: Err.Raise vbObjectError + 901, , "Geçersiz geçmiþ türü."
    End Select
    AKFIX_HAFIZA_ADI = "_STOK_HAFIZA_" & tur
    If il <> "ARTVIN" Then AKFIX_HAFIZA_ADI = AKFIX_HAFIZA_ADI & "_" & il
End Function

Private Function AKFIX_WS(ByVal ad As String) As Worksheet
    On Error Resume Next
    Set AKFIX_WS = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
End Function

Public Sub AKFIX_EN_USTE(ByVal ws As Worksheet)
    If Not (ActiveWorkbook Is ThisWorkbook) Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    ActiveWindow.ScrollRow = 1
    ActiveWindow.ScrollColumn = 1
End Sub

Private Function AKFIX_IMZA(ByVal ws As Worksheet) As String
    Dim f As Filter, i As Long, v As Variant, k As Variant, t As String
    t = ws.Name
    If Not ws.AutoFilterMode Then AKFIX_IMZA = t: Exit Function
    t = t & "|" & ws.AutoFilter.Range.Address
    For i = 1 To ws.AutoFilter.Filters.Count
        Set f = ws.AutoFilter.Filters(i)
        If f.On Then
            t = t & "|" & i & ":" & f.Operator
            v = Empty
            On Error Resume Next
            v = f.Criteria1
            On Error GoTo 0
            If IsArray(v) Then
                For Each k In v: t = t & ";" & CStr(k): Next k
            Else
                t = t & ";" & CStr(v)
            End If
            v = Empty
            On Error Resume Next
            v = f.Criteria2
            On Error GoTo 0
            If IsArray(v) Then
                For Each k In v: t = t & ";" & CStr(k): Next k
            ElseIf Not IsEmpty(v) Then
                t = t & ";" & CStr(v)
            End If
        End If
    Next i
    AKFIX_IMZA = t
End Function

Public Sub AKFIX_FILTRE_KONTROL(ByVal ws As Worksheet)
    Dim t As String
    If Not application.EnableEvents Then Exit Sub
    If Not (ActiveSheet Is ws) Then Exit Sub
    If mImza Is Nothing Then Set mImza = CreateObject("Scripting.Dictionary")
    t = AKFIX_IMZA(ws)
    If mImza.Exists(ws.Name) Then
        If mImza(ws.Name) <> t Then AKFIX_EN_USTE ws
    End If
    mImza(ws.Name) = t
End Sub

Public Sub AKFIX_STOK_LISTE(ByVal ws As Worksheet, Optional ByVal sifirla As Boolean = False)
    Dim d As Object, r As Long, son As Long, eskiSon As Long
    Dim veri As Variant, k As Variant, metin As String, sec As String, il As String
    Dim olay As Boolean, ad As String, oncekiIl As String
    Dim hataNo As Long, hataMetni As String
    olay = application.EnableEvents
    On Error GoTo Hata
    application.EnableEvents = False
    il = AKSORA_CANLI_STOK_IL()
    If mListeIl Is Nothing Then Set mListeIl = CreateObject("Scripting.Dictionary")
    If mListeIl.Exists(ws.Name) Then oncekiIl = mListeIl(ws.Name)
    If oncekiIl <> il Then sifirla = True
    son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    If son < 7 Then son = 7
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    veri = ws.Range("C7:C" & son).Value2
    If son = 7 Then
        metin = Trim$(CStr(ws.Range("C7").Value2))
        If Len(metin) > 0 Then d(metin) = True
    Else
        For r = 1 To UBound(veri, 1)
            metin = Trim$(CStr(veri(r, 1)))
            If Len(metin) > 0 Then d(metin) = True
        Next r
    End If
    sec = Trim$(CStr(ws.Range("H2").Value2))
    eskiSon = ws.Cells(ws.Rows.Count, "N").End(xlUp).Row
    If eskiSon < 2 Then eskiSon = 2
    ws.Range("N2:N" & eskiSon).ClearContents
    ws.Range("N2").Value = "TÜMÜ"
    r = 3
    For Each k In d.keys
        ws.Cells(r, "N").Value = CStr(k)
        r = r + 1
    Next k
    If r > 4 Then ws.Range("N3:N" & r - 1).Sort Key1:=ws.Range("N3"), Order1:=xlAscending, Header:=xlNo
    ad = "_AKFIX_DEPO_" & ws.CodeName
    ThisWorkbook.names.Add Name:=ad, RefersTo:="='" & Replace(ws.Name, "'", "''") & "'!$N$2:$N$" & r - 1
    With ws.Range("H2").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=" & ad
        .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True
    End With
    If sifirla Or Not d.Exists(sec) Then sec = "TÜMÜ"
    ws.Range("H2").Value = sec
    If AKFIX_NORMAL(ws.Name) = "ILCE_STOK" Then
        ws.Range("G2").Value = "ÝLÇELER"
    Else
        ws.Range("G2").Value = "DEPOLAR"
    End If
    ws.Columns("N").Hidden = True
    If sifirla Then
        If ws.FilterMode Then ws.ShowAllData
    End If
    If ws.AutoFilterMode Then
        If ws.AutoFilter.Range.Address <> ws.Range("A6:L" & son).Address Then ws.AutoFilterMode = False
    End If
    If Not ws.AutoFilterMode Then ws.Range("A6:L" & son).AutoFilter
    ws.Range("AZ1").Formula = "=SUBTOTAL(103,G7:G" & son & ")"
    ws.Columns("AZ").Hidden = True
    AKFIX_EK_OKLARI_KALDIR ws
    AKFIX_BOS_SERITLERI_TEMIZLE ws
    mListeIl(ws.Name) = il
    If mImza Is Nothing Then Set mImza = CreateObject("Scripting.Dictionary")
    mImza(ws.Name) = AKFIX_IMZA(ws)
Cikis:
    application.EnableEvents = olay
    Exit Sub
Hata:
    hataNo = Err.Number: hataMetni = Err.Description
    application.EnableEvents = olay
    Err.Raise hataNo, "AKFIX_STOK_LISTE", hataMetni
End Sub

Public Sub AKFIX_STOK_SECIM(ByVal ws As Worksheet)
    Dim sec As String, son As Long
    sec = Trim$(CStr(ws.Range("H2").Value2))
    son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    If son < 7 Then son = 7
    If Not ws.AutoFilterMode Then ws.Range("A6:L" & son).AutoFilter
    If Len(sec) = 0 Or AKFIX_NORMAL(sec) = "TUMU" Then
        If ws.FilterMode Then ws.ShowAllData
    Else
        sec = Replace(sec, "~", "~~")
        sec = Replace(sec, "*", "~*"): sec = Replace(sec, "?", "~?")
        ws.Range("A6:L" & son).AutoFilter Field:=3, Criteria1:="=" & sec
    End If
    If mImza Is Nothing Then Set mImza = CreateObject("Scripting.Dictionary")
    mImza(ws.Name) = AKFIX_IMZA(ws)
    AKFIX_KPI_FORMULLERI ws
    AKFIX_BOS_SERITLERI_TEMIZLE ws
    ws.Calculate
    AKFIX_EN_USTE ws
End Sub

Public Sub AKFIX_STOKLARI_YENILE()
    Dim ad As Variant, ws As Worksheet
    For Each ad In Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
        Set ws = ThisWorkbook.Worksheets(CStr(ad))
        AKFIX_STOK_LISTE ws, True
    Next ad
    Set ws = ThisWorkbook.Worksheets("STOKLAR")
    AKFIX_NORMAL_STOK_HAZIRLA ws
End Sub

Public Sub AKFIX_NORMAL_STOK_HAZIRLA(ByVal ws As Worksheet)
    Dim son As Long
    son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    If son < 7 Then son = 7
    If ws.AutoFilterMode Then
        If ws.AutoFilter.Range.Address <> ws.Range("A6:L" & son).Address Then ws.AutoFilterMode = False
    End If
    If Not ws.AutoFilterMode Then ws.Range("A6:L" & son).AutoFilter
    ws.Range("AZ1").Formula = "=SUBTOTAL(103,G7:G" & son & ")"
    ws.Columns("AZ").Hidden = True
    If mImza Is Nothing Then Set mImza = CreateObject("Scripting.Dictionary")
    mImza(ws.Name) = AKFIX_IMZA(ws)
    AKFIX_BOS_SERITLERI_TEMIZLE ws
End Sub

Public Function AKFIX_KARS_IL() As String
    AKFIX_KARS_IL = AKFIX_NORMAL(CStr(ThisWorkbook.Worksheets("STOK_KARÞILAÞTIRMA").Range("AA2").Value2))
    Select Case AKFIX_KARS_IL
        Case "ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN", "TUMU"
        Case Else: AKFIX_KARS_IL = AKSORA_CANLI_STOK_IL()
    End Select
End Function

Private Function AKFIX_GECMIS_VAR(ByVal il As String) As Boolean
    Dim ws As Worksheet
    Set ws = AKFIX_WS(AKFIX_HAFIZA_ADI(il, "INDEX"))
    If ws Is Nothing Then Exit Function
    AKFIX_GECMIS_VAR = (ws.Cells(ws.Rows.Count, "A").End(xlUp).Row >= 2)
End Function

Public Sub AKFIX_KARS_TARIHLER()
    Dim w As Worksheet, idx As Worksheet, d As Object
    Dim il As Variant, iller As Variant, sec As String
    Dim r As Long, son As Long, tarih As Double, ilk As Double, ortakIlk As Double
    Dim k As Variant, n As Long, alt As String
    Dim hataNo As Long, hataMetni As String
    On Error GoTo Hata
    Set w = ThisWorkbook.Worksheets("STOK_KARÞILAÞTIRMA")
    AKFIX_KARS_GECMIS_FARK = 0
    sec = AKFIX_KARS_IL()
    If sec = "TUMU" Then
        iller = Array("ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
    Else
        iller = Array(sec)
    End If
    Set mKarsIller = CreateObject("Scripting.Dictionary")
    Set d = CreateObject("Scripting.Dictionary")
    For Each il In iller
        If AKFIX_GECMIS_VAR(CStr(il)) Then
            mKarsIller(CStr(il)) = True
            Set idx = ThisWorkbook.Worksheets(AKFIX_HAFIZA_ADI(CStr(il), "INDEX"))
            son = idx.Cells(idx.Rows.Count, "A").End(xlUp).Row
            ilk = 1E+20
            For r = 2 To son
                If IsNumeric(idx.Cells(r, "B").Value2) Then
                    tarih = CDbl(idx.Cells(r, "B").Value2)
                    If tarih > 0 Then
                        If tarih < ilk Then ilk = tarih
                        d(CStr(tarih)) = tarih
                    End If
                End If
            Next r
            If ilk < 1E+20 And ilk > ortakIlk Then ortakIlk = ilk
        End If
    Next il
    If w.FilterMode Then w.ShowAllData
    son = w.Cells(w.Rows.Count, "A").End(xlUp).Row
    If son >= 16 Then w.Range("A16:M" & son).ClearContents
    For Each k In Array("A10", "D10", "G10", "J10", "A12", "D12", "G12", "J12")
        w.Range(CStr(k)).Value = 0
    Next k
    w.Range("B6").MergeArea.ClearContents: w.Range("I6").MergeArea.ClearContents
    w.Range("B7").MergeArea.ClearContents: w.Range("I7").MergeArea.ClearContents
    w.Range("B6").Validation.Delete: w.Range("I6").Validation.Delete
    son = w.Cells(w.Rows.Count, "X").End(xlUp).Row
    If son < 2 Then son = 2
    w.Range("X1:Y" & son).ClearContents
    w.Range("X1").Value = "GÖRÜNEN": w.Range("Y1").Value = "SNAPSHOT_ID"
    n = 2
    For Each k In d.keys
        tarih = CDbl(d(k))
        If sec <> "TUMU" Or tarih >= ortakIlk - 0.00000001 Then
            w.Cells(n, "X").Value2 = tarih
            n = n + 1
        End If
    Next k
    If n > 3 Then w.Range("X2:X" & n - 1).Sort Key1:=w.Range("X2"), Order1:=xlAscending, Header:=xlNo
    For r = 2 To n - 1
        w.Cells(r, "X").Value = Format$(CDate(w.Cells(r, "X").Value2), "dd.mm.yyyy hh:nn:ss")
        w.Cells(r, "Y").Value = "AKFIX-" & CStr(r)
    Next r
    If n > 2 Then
        ThisWorkbook.names.Add Name:="_AKFIX_KARS_TARIHLER", RefersTo:="='" & w.Name & "'!$X$2:$X$" & n - 1
        For Each k In Array("B6", "I6")
            With w.Range(CStr(k)).Validation
                .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=_AKFIX_KARS_TARIHLER"
                .IgnoreBlank = True: .InCellDropdown = True: .ShowError = True
            End With
        Next k
        w.Range("B6").Value = w.Cells(application.Max(2, n - 2), "X").Value
        w.Range("I6").Value = w.Cells(n - 1, "X").Value
    End If
    If sec = "TUMU" Then
        alt = "TÜMÜ | "
        For Each k In mKarsIller.keys: alt = alt & CStr(k) & " / ": Next k
        If mKarsIller.Count > 0 Then alt = Left$(alt, Len(alt) - 3)
        alt = alt & " | Seçilen zamana kadar kaydedilmiþ son stoklar"
    Else
        alt = sec & " | CDP " & AKSORA_IL_TESIS(sec, "CDP") & " + EPP " & AKSORA_IL_TESIS(sec, "EPP") & " | Normal Stok Hafýzasý"
    End If
    w.Range("A3").Value = alt
    If n = 2 Then w.Range("B7").Value = "Bu il için geçmiþ kayýt yok."
    If n = 3 Then w.Range("B7").Value = "Ýkinci geçmiþ kayýt oluþtuðunda karþýlaþtýrýlabilir."
    w.Columns("X:Y").Hidden = True
    AKSORA_V11_KARS_UYARI w
    AKFIX_EN_USTE w
Cikis:
    Exit Sub
Hata:
    hataNo = Err.Number: hataMetni = Err.Description
    Err.Raise hataNo, "AKFIX_KARS_TARIHLER", hataMetni
End Sub

Public Function AKFIX_KARS_TARIH(ByVal id As String) As Date
    Dim w As Worksheet, r As Long, son As Long
    Set w = ThisWorkbook.Worksheets("STOK_KARÞILAÞTIRMA")
    son = w.Cells(w.Rows.Count, "Y").End(xlUp).Row
    For r = 2 To son
        If CStr(w.Cells(r, "Y").Value2) = id Then
            AKFIX_KARS_TARIH = CDate(w.Cells(r, "X").Value)
            Exit Function
        End If
    Next r
    Err.Raise vbObjectError + 902, , "Karþýlaþtýrma tarihi bulunamadý."
End Function

Private Function AKFIX_INDEX_SATIR(ByVal il As String, ByVal tarih As Date) As Long
    Dim idx As Worksheet, r As Long, son As Long, t As Double, enSon As Double
    Set idx = ThisWorkbook.Worksheets(AKFIX_HAFIZA_ADI(il, "INDEX"))
    son = idx.Cells(idx.Rows.Count, "A").End(xlUp).Row
    For r = 2 To son
        If IsNumeric(idx.Cells(r, "B").Value2) Then
            t = CDbl(idx.Cells(r, "B").Value2)
            If t <= CDbl(tarih) + 0.00000001 And t >= enSon Then
                enSon = t: AKFIX_INDEX_SATIR = r
            End If
        End If
    Next r
    If AKFIX_INDEX_SATIR = 0 Then Err.Raise vbObjectError + 903, , il & " için bu tarihte geçmiþ kayýt yok."
End Function

Public Function AKFIX_KARS_TOPLAM(ByVal id As String) As Double
    Dim tarih As Date, il As Variant, idx As Worksheet, r As Long
    tarih = AKFIX_KARS_TARIH(id)
    If mKarsIller Is Nothing Then Err.Raise vbObjectError + 904, , "Önce karþýlaþtýrýlacak ili seçin."
    For Each il In mKarsIller.keys
        Set idx = ThisWorkbook.Worksheets(AKFIX_HAFIZA_ADI(CStr(il), "INDEX"))
        r = AKFIX_INDEX_SATIR(CStr(il), tarih)
        AKFIX_KARS_TOPLAM = AKFIX_KARS_TOPLAM + CDbl(idx.Cells(r, "E").Value2)
    Next il
End Function

Public Function AKFIX_KARS_DURUM(ByVal id As String) As Object
    Dim d As Object, il As Variant, ws As Worksheet, idx As Worksheet
    Dim veri As Variant, kayit As Variant, mevcut As Variant, anahtar As String
    Dim r As Long, j As Long, son As Long, ix As Long, tarih As Date, hedef As Double
    Dim durum As String, tesis As String, cdp As String, epp As String
    If mKarsIller Is Nothing Then Err.Raise vbObjectError + 904, , "Önce karþýlaþtýrýlacak ili seçin."
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    tarih = AKFIX_KARS_TARIH(id)
    For Each il In mKarsIller.keys
        cdp = AKSORA_IL_TESIS(CStr(il), "CDP"): epp = AKSORA_IL_TESIS(CStr(il), "EPP")
        Set idx = ThisWorkbook.Worksheets(AKFIX_HAFIZA_ADI(CStr(il), "INDEX"))
        ix = AKFIX_INDEX_SATIR(CStr(il), tarih)
        hedef = CDbl(idx.Cells(ix, "B").Value2)
        Set ws = ThisWorkbook.Worksheets(AKFIX_HAFIZA_ADI(CStr(il), "CURRENT"))
        son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If son < 2 Then Err.Raise vbObjectError + 905, , CStr(il) & " geçmiþ stok verisi boþ."
        veri = ws.Range("A2:L" & son).Value2
        For r = 1 To UBound(veri, 1)
            If Len(Trim$(CStr(veri(r, 7)))) > 0 Then
                tesis = Trim$(CStr(veri(r, 1)))
                If tesis <> cdp And tesis <> epp Then Err.Raise vbObjectError + 906, , CStr(il) & " geçmiþinde farklý ilin tesisi var."
                anahtar = tesis & "|" & UCase$(Trim$(CStr(veri(r, 2)))) & "|" & Trim$(CStr(veri(r, 7))) & "|" & UCase$(Trim$(CStr(veri(r, 6))))
                ReDim kayit(1 To 12)
                For j = 1 To 12: kayit(j) = veri(r, j): Next j
                If d.Exists(anahtar) Then
                    mevcut = d(anahtar)
                    If StrComp(CStr(mevcut(10)), CStr(kayit(10)), vbTextCompare) <> 0 Then Err.Raise vbObjectError + 907, , "Geçmiþ stok birimleri uyuþmuyor."
                    mevcut(9) = CDbl(mevcut(9)) + CDbl(kayit(9))
                    mevcut(11) = CDec(mevcut(11)) + CDec(kayit(11))
                    d(anahtar) = mevcut
                Else
                    d.Add anahtar, kayit
                End If
            End If
        Next r
        Set ws = ThisWorkbook.Worksheets(AKFIX_HAFIZA_ADI(CStr(il), "DELTA"))
        son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        If son >= 2 Then
            veri = ws.Range("A2:N" & son).Value2
            For r = UBound(veri, 1) To 1 Step -1
                If IsNumeric(veri(r, 2)) Then
                    If CDbl(veri(r, 2)) > hedef + 0.00000001 Then
                        anahtar = CStr(veri(r, 4))
                        durum = AKFIX_NORMAL(CStr(veri(r, 3)))
                        If durum = "YENI" Then
                            If d.Exists(anahtar) Then d.Remove anahtar
                        Else
                            ReDim kayit(1 To 12)
                            kayit(1) = veri(r, 5): kayit(2) = veri(r, 6): kayit(3) = veri(r, 7)
                            kayit(4) = "": kayit(5) = "": kayit(6) = veri(r, 10)
                            kayit(7) = veri(r, 8): kayit(8) = veri(r, 9)
                            kayit(9) = veri(r, 12): kayit(10) = veri(r, 11)
                            kayit(11) = veri(r, 14): kayit(12) = ""
                            d(anahtar) = kayit
                        End If
                    End If
                End If
            Next r
        End If
    Next il
    AKFIX_KARS_DOGRULA d, AKFIX_KARS_TOPLAM(id)
    Set AKFIX_KARS_DURUM = d
End Function

Private Sub AKFIX_KARS_DOGRULA(ByVal d As Object, ByVal beklenen As Double)
    Dim k As Variant, kayit As Variant, toplam As Double
    For Each k In d.keys
        kayit = d(k)
        toplam = toplam + CDbl(kayit(11))
    Next k
    If Abs(toplam - beklenen) >= 0.005 Then
        If Abs(toplam - beklenen) > AKFIX_KARS_GECMIS_FARK Then AKFIX_KARS_GECMIS_FARK = Abs(toplam - beklenen)
    End If
    If Abs(toplam - beklenen) > 0.05 Then
        Err.Raise vbObjectError + 908, , _
            "Bu tarihin geçmiþ kayýtlarý toplam stok deðeriyle uyuþmuyor. " & _
            "Hatalý bir karþýlaþtýrma gösterilmedi."
    End If
End Sub

Public Sub AKFIX_KUR()
    Dim olay As Boolean, ekran As Boolean, ad As Variant
    olay = application.EnableEvents: ekran = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    AKFIX_STOKLARI_YENILE
    For Each ad In Array("CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ÝLÇE_RAPORLAR", "YÜKLENÝCÝ_RAPOR")
        RAPOR_GORUNUM_TEK CStr(ad)
    Next ad
    AKFIX_KARS_TARIHLER
    application.EnableEvents = olay: application.ScreenUpdating = ekran
    MsgBox "Filtreler, rapor biçimleri ve ortak karþýlaþtýrma yenilendi.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    application.EnableEvents = olay: application.ScreenUpdating = ekran
    MsgBox "Düzeltme kurulumu: " & Err.Description, vbExclamation, "AKSORA"
End Sub

Private Sub AKFIX_EK_OKLARI_KALDIR(ByVal ws As Worksheet)
    Dim ad As Variant, s As Shape
    For Each ad In Array("AKSORA_KALICI_DEPO", "AKSORA_MAVI_FILTRE_OK")
        Set s = Nothing
        On Error Resume Next
        Set s = ws.Shapes(CStr(ad))
        On Error GoTo 0
        If Not s Is Nothing Then s.Delete
    Next ad
End Sub

Private Sub AKFIX_KPI_FORMULLERI(ByVal ws As Worksheet)
    Dim son As Long, f As String
    son = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row
    If son < 7 Then son = 7
    f = "=SUBTOTAL(103,A7:A" & son & ")"
    If ws.Range("C3").Formula <> f Then ws.Range("C3").Formula = f
    f = "=SUBTOTAL(109,K7:K" & son & ")"
    If ws.Range("C4").Formula <> f Then ws.Range("C4").Formula = f
End Sub

Private Sub AKFIX_BOS_SERITLERI_TEMIZLE(ByVal ws As Worksheet)
    Dim son As Long, bitis As Long, dolu As Range, alan As Range
    ' Find gizli/filtrelenmis satirlari da gorur. Gorunen son satir kullanilmaz.
    Set dolu = ws.Range("A:L").Find(What:="*", after:=ws.Range("A1"), _
        LookIn:=xlFormulas, LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)
    son = 6
    If Not dolu Is Nothing Then
        If dolu.Row > son Then son = dolu.Row
    End If
    bitis = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
    If bitis <= son Then Exit Sub
    Set alan = ws.Range("A" & (son + 1) & ":L" & bitis)
    ' Veri, formul, sayi bicimi, satir gizliligi ve sutun genisligi korunur.
    alan.Interior.Pattern = xlNone
    alan.Borders.LineStyle = xlNone
End Sub

Public Sub AKSORA_STOK_FILTRE_SADELESTIR()
    Dim ad As Variant, ws As Worksheet, ekran As Boolean, olay As Boolean
    Dim hataMetni As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Then Exit Sub
    ekran = application.ScreenUpdating
    olay = application.EnableEvents
    On Error GoTo Hata
    application.ScreenUpdating = False
    application.EnableEvents = False
    For Each ad In Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK", "STOKLAR")
        Set ws = ThisWorkbook.Worksheets(CStr(ad))
        AKFIX_EK_OKLARI_KALDIR ws
        AKFIX_BOS_SERITLERI_TEMIZLE ws
        AKFIX_KPI_FORMULLERI ws
        ws.Calculate
    Next ad
    application.EnableEvents = True
    application.ScreenUpdating = ekran
    MsgBox "Ek oklar kaldirildi. Yerlesik depo listesi korundu; tablo disindaki bos seritler temizlendi.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    hataMetni = Err.Description
    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    MsgBox "Duzenleme tamamlanamadi: " & hataMetni, vbExclamation, "AKSORA"
End Sub

' Onceki kurulum adlari artik ek kontrol olusturmaz.
Public Sub AKSORA_MAVI_FILTRE_ONAR()
    AKSORA_STOK_FILTRE_SADELESTIR
End Sub

Public Sub AKSORA_KALICI_FILTRE_KUR()
    AKSORA_STOK_FILTRE_SADELESTIR
End Sub


