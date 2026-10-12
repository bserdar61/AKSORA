Attribute VB_Name = "modAKSORA_MUSTERI_FIYAT"
Option Explicit

'==========================================================
' AKSORA - MUSTERI BIRIMI FIYAT LISTESI
'
' MB52 yerine: Malzeme + Parti (degerleme turu) bazli
' birim fiyat listesi gizli "_FIYAT_LISTESI" sayfasinda tutulur.
' Deger = Miktar x Birim fiyat (SAP guncellemesinde hesaplanir).
'
' Kullanim: AKSORA_MUSTERI_FIYAT_LISTESI_YUKLE makrosu ile
' SAP'tan alinan EXPORT dosyasi secilir. Liste her yuklemede
' bastan olusturulur (ayni malzeme+parti tek satira iner).
'
' Cakisma kurali (ayni malzeme+parti icin birden cok fiyat):
'  - 0 fiyatli satirlar, pozitif fiyat varsa yok sayilir
'  - Pozitif fiyatlar birbirine 0,02 TL veya %0,1 icindeyse
'    (kurus yuvarlama farki) TEK sayilir; en sik gecen fiyat alinir
'  - Aksi halde BELIRSIZ: fiyat uygulanmaz, satir "Fiyat belirsiz"
'    olarak isaretlenir. Duzeltmek icin sayfada C sutununa fiyat,
'    D sutununa TEK yazilabilir (sonraki yuklemede silinir).
'==========================================================

Private Const FIYAT_SAYFA As String = "_FIYAT_LISTESI"
Private Const TOL_MUTLAK As Double = 0.02
Private Const TOL_ORAN As Double = 0.001

'----------------------------------------------------------
' Eslesme anahtari: MALZEME|PARTI (bastaki sifirlar atilir)
'----------------------------------------------------------
Public Function MF_Anahtar(ByVal malzeme As Variant, ByVal parti As Variant) As String
    Dim m As String, p As String
    If Not IsError(malzeme) Then m = Trim$(CStr(malzeme))
    If Not IsError(parti) Then p = CStr(parti)
    If Len(m) > 0 Then
        If Not (m Like "*[!0-9]*") Then
            Do While Len(m) > 1 And Left$(m, 1) = "0": m = Mid$(m, 2): Loop
        End If
    End If
    MF_Anahtar = AKFIX_NORMAL(m) & "|" & AKFIX_NORMAL(p)
End Function

Private Function MF_Kolon(ByVal ws As Worksheet, ByVal adaylar As String) As Long
    Dim c As Long, t As String
    For c = 1 To 40
        t = vbNullString
        If Not IsError(ws.Cells(1, c).Value2) Then t = CStr(ws.Cells(1, c).Value2)
        t = AKFIX_NORMAL(t)
        t = Replace(Replace(Replace(t, " ", ""), ".", ""), "/", "")
        If Len(t) > 0 Then
            If InStr(1, "|" & adaylar & "|", "|" & t & "|") > 0 Then MF_Kolon = c: Exit Function
        End If
    Next c
End Function

Private Function MF_DosyaTarihi(ByVal yol As String) As Date
    Dim ad As String, i As Long, s As String, ay As Long, gun As Long, saat As Long
    ad = Mid$(yol, InStrRev(yol, "\") + 1)
    For i = 1 To Len(ad) - 13
        s = Mid$(ad, i, 14)
        If s Like "##############" Then
            ay = CLng(Mid$(s, 5, 2)): gun = CLng(Mid$(s, 7, 2)): saat = CLng(Mid$(s, 9, 2))
            If ay >= 1 And ay <= 12 And gun >= 1 And gun <= 31 And saat <= 23 Then
                MF_DosyaTarihi = DateSerial(CLng(Left$(s, 4)), ay, gun) + _
                    TimeSerial(saat, CLng(Mid$(s, 11, 2)), CLng(Mid$(s, 13, 2)))
                Exit Function
            End If
        End If
    Next i
    MF_DosyaTarihi = FileDateTime(yol)
End Function

'----------------------------------------------------------
' SAP guncellemesinin kullandigi fiyat sozlugu
' Anahtar -> Array(fiyat As Double, durum As String)
' durum: TEK / BELIRSIZ / SIFIR
'----------------------------------------------------------
Public Function MF_FIYAT_SOZLUGU(ByRef tarih As Date, ByRef Hata As String) As Object
    Dim ws As Worksheet, v As Variant, r As Long, son As Long
    Dim d As Object, k As String, durum As String
    Hata = vbNullString
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(FIYAT_SAYFA)
    On Error GoTo 0
    If ws Is Nothing Then
        Hata = "Fiyat listesi yüklenmemiþ. Önce AKSORA_MUSTERI_FIYAT_LISTESI_YUKLE makrosunu çalýþtýrýp fiyat dosyasýný seçin."
        Exit Function
    End If
    son = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If son < 2 Then
        Hata = "Fiyat listesi boþ. AKSORA_MUSTERI_FIYAT_LISTESI_YUKLE makrosuyla yeniden yükleyin."
        Exit Function
    End If
    If IsDate(ws.Range("J1").Value2) Or IsNumeric(ws.Range("J1").Value2) Then tarih = CDate(ws.Range("J1").Value2)
    v = ws.Range("A2:D" & son).Value2
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For r = 1 To UBound(v, 1)
        If Not IsError(v(r, 1)) Then
            If Len(Trim$(CStr(v(r, 1)))) > 0 Then
                k = MF_Anahtar(v(r, 1), v(r, 2))
                durum = vbNullString
                If Not IsError(v(r, 4)) Then durum = AKFIX_NORMAL(CStr(v(r, 4)))
                If Not d.Exists(k) Then
                    If durum = "TEK" And Not IsError(v(r, 3)) And VarType(v(r, 3)) <> vbString And IsNumeric(v(r, 3)) Then
                        If CDbl(v(r, 3)) > 0 Then
                            d.Add k, Array(CDbl(v(r, 3)), "TEK")
                        Else
                            d.Add k, Array(0#, "SIFIR")
                        End If
                    ElseIf durum = "TEK" Or durum = "SIFIR" Then
                        d.Add k, Array(0#, "SIFIR")
                    Else
                        d.Add k, Array(0#, "BELIRSIZ")
                    End If
                End If
            End If
        End If
    Next r
    Set MF_FIYAT_SOZLUGU = d
End Function

'----------------------------------------------------------
' FIYAT DOSYASINI YUKLE
'----------------------------------------------------------
Public Sub AKSORA_MUSTERI_FIYAT_LISTESI_YUKLE()
    Dim dosya As Variant, wbK As Workbook, wsK As Worksheet, wsF As Worksheet, aktif As Object
    Dim v As Variant, cM As Long, cP As Long, cF As Long, cT As Long, cG As Long, maxc As Long
    Dim son As Long, r As Long, d As Object, pd As Object, kayit As Variant, k As Variant, fk As Variant
    Dim m As String, f As Double, ham As Variant
    Dim bos As Long, gecersiz As Long, okunan As Long
    Dim ev As Boolean, sc As Boolean, al As Boolean
    Dim lo As Double, hi As Double, secili As Double, enCok As Long, poz As Long, sifirVar As Boolean, liste As String
    Dim out() As Variant, n As Long, i As Long, gecis As Long, durum As String, nt As String
    Dim nTek As Long, nBel As Long, nSif As Long, ozet As String, tarih As Date, ilk As Long

    dosya = application.GetOpenFilename("Excel dosyalarý (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , "Birim fiyat dosyasýný seçin")
    If VarType(dosya) = vbBoolean Then Exit Sub

    ev = application.EnableEvents: sc = application.ScreenUpdating: al = application.DisplayAlerts
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False: application.DisplayAlerts = False
    Set aktif = ActiveSheet

    Set wbK = Workbooks.Open(CStr(dosya), ReadOnly:=True, UpdateLinks:=0)
    Set wsK = wbK.Worksheets(1)
    cM = MF_Kolon(wsK, "MALZEME|MATNR")
    cP = MF_Kolon(wsK, "PARTI|DEGERLEMETURU|DEGERLTURU|BWTAR")
    cF = MF_Kolon(wsK, "BIRIMFIYAT|BIRIMFIYATTL|FIYAT")
    cT = MF_Kolon(wsK, "MALZEMEKISAMETNI|MALZEMEACIKLAMASI|MAKTX|TANIM")
    cG = MF_Kolon(wsK, "MALGRUBU|MATKL")
    If cM = 0 Or cF = 0 Or cP = 0 Then
        Err.Raise vbObjectError + 2001, , "Fiyat dosyasýnda 'Malzeme', 'Parti' ve 'Birim Fiyat' baþlýklarý (1. satýr) bulunamadý."
    End If
    maxc = cM: If cP > maxc Then maxc = cP
    If cF > maxc Then maxc = cF
    If cT > maxc Then maxc = cT
    If cG > maxc Then maxc = cG
    son = wsK.Cells(wsK.Rows.Count, cM).End(xlUp).Row
    If son < 2 Then Err.Raise vbObjectError + 2002, , "Fiyat dosyasýnda veri satýrý yok."
    v = wsK.Range(wsK.Cells(2, 1), wsK.Cells(son, maxc)).Value2
    tarih = MF_DosyaTarihi(CStr(dosya))

    Set d = CreateObject("Scripting.Dictionary")
    For r = 1 To UBound(v, 1)
        m = vbNullString
        If Not IsError(v(r, cM)) Then m = Trim$(CStr(v(r, cM)))
        If Len(m) = 0 Then
            bos = bos + 1
        ElseIf IsError(v(r, cF)) Then
            gecersiz = gecersiz + 1
        ElseIf IsEmpty(v(r, cF)) Or VarType(v(r, cF)) = vbString Or Not IsNumeric(v(r, cF)) Then
            gecersiz = gecersiz + 1
        ElseIf CDbl(v(r, cF)) < 0 Then
            gecersiz = gecersiz + 1
        Else
            okunan = okunan + 1
            f = Round(CDbl(v(r, cF)), 2)
            k = MF_Anahtar(m, v(r, cP))
            If Not d.Exists(k) Then
                Set pd = CreateObject("Scripting.Dictionary")
                ham = vbNullString: If cT > 0 Then If Not IsError(v(r, cT)) Then ham = CStr(v(r, cT))
                nt = vbNullString: If cG > 0 Then If Not IsError(v(r, cG)) Then nt = CStr(v(r, cG))
                If IsError(v(r, cP)) Then
                    d.Add k, Array(m, vbNullString, ham, nt, pd)
                Else
                    d.Add k, Array(m, Trim$(CStr(v(r, cP))), ham, nt, pd)
                End If
            End If
            kayit = d(k)
            Set pd = kayit(4)
            If pd.Exists(f) Then pd(f) = pd(f) + 1 Else pd.Add f, 1
        End If
    Next r
    If d.Count = 0 Then Err.Raise vbObjectError + 2003, , "Fiyat dosyasýndan geçerli satýr okunamadý."

    ReDim out(1 To d.Count, 1 To 7)
    For gecis = 1 To 2
        For Each k In d.keys
            kayit = d(k): Set pd = kayit(4)
            poz = 0: lo = 0: hi = 0: enCok = 0: secili = 0: liste = vbNullString: sifirVar = False
            For Each fk In pd.keys
                f = CDbl(fk)
                If f > 0 Then
                    If poz = 0 Then lo = f: hi = f
                    If f < lo Then lo = f
                    If f > hi Then hi = f
                    poz = poz + 1
                    If pd(fk) > enCok Or (pd(fk) = enCok And f > secili) Then enCok = pd(fk): secili = f
                    If Len(liste) > 0 Then liste = liste & "; "
                    liste = liste & Format$(f, "0.00")
                Else
                    sifirVar = True
                End If
            Next fk
            nt = vbNullString
            If poz = 0 Then
                durum = "SIFIR": secili = 0: nt = "Listede yalnýzca 0 fiyatý var"
            ElseIf (hi - lo) <= TOL_MUTLAK Or (hi - lo) / hi <= TOL_ORAN Then
                durum = "TEK"
                If poz > 1 Then nt = "Kuruþ farklarý birleþtirildi: " & liste
                If sifirVar Then nt = nt & IIf(Len(nt) > 0, " | ", "") & "0 fiyatlý satýrlar yok sayýldý"
            Else
                ' Ayni malzeme + parti icin farkli fiyatlar: listede en sik gecen fiyat alinir
                ' (esitlikte yuksek olan). YENI ve CIKMA/BDSZ partileri ayri tutulur.
                durum = "TEK"
                nt = "Farklý fiyatlar (" & liste & "); en sýk geçen fiyat alýndý"
            End If
            If (gecis = 1 And durum <> "TEK") Or (gecis = 2 And durum = "TEK") Then
                i = i + 1
                out(i, 1) = kayit(0): out(i, 2) = kayit(1)
                If durum = "TEK" Then out(i, 3) = secili Else out(i, 3) = Empty
                out(i, 4) = durum: out(i, 5) = kayit(2): out(i, 6) = kayit(3): out(i, 7) = nt
                If durum = "TEK" Then
                    nTek = nTek + 1
                ElseIf durum = "BELIRSIZ" Then
                    nBel = nBel + 1
                Else
                    nSif = nSif + 1
                End If
            End If
        Next k
    Next gecis

    wbK.Close SaveChanges:=False
    Set wbK = Nothing

    On Error Resume Next
    Set wsF = ThisWorkbook.Worksheets(FIYAT_SAYFA)
    On Error GoTo Hata
    If wsF Is Nothing Then
        Set wsF = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsF.Name = FIYAT_SAYFA
    End If
    wsF.Visible = xlSheetVisible
    wsF.Cells.Clear
    wsF.Range("A:B").NumberFormat = "@"
    wsF.Range("A1:G1").Value2 = Array("Malzeme", "Parti", "Birim Fiyat", "Durum", "Malzeme kýsa metni", "Mal grubu", "Not")
    wsF.Range("A2").Resize(d.Count, 7).Value2 = out
    wsF.Range("C2:C" & d.Count + 1).NumberFormat = "#,##0.00 ""TL"""
    wsF.Range("I1").Value2 = "Liste tarihi": wsF.Range("J1").Value2 = CDbl(tarih)
    wsF.Range("J1").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsF.Range("I2").Value2 = "Kaynak dosya": wsF.Range("J2").Value2 = Mid$(CStr(dosya), InStrRev(CStr(dosya), "\") + 1)
    wsF.Range("I3").Value2 = "Yükleme zamaný": wsF.Range("J3").Value2 = CDbl(Now)
    wsF.Range("J3").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsF.Columns("A:J").AutoFit
    wsF.Visible = xlSheetHidden
    On Error Resume Next
    aktif.Activate
    On Error GoTo 0

    ozet = "Fiyat listesi yüklendi." & vbCrLf & vbCrLf & _
        "Kaynak satýr: " & okunan & " | Benzersiz malzeme+parti: " & d.Count & vbCrLf & _
        "Kullanýlabilir (TEK): " & nTek & vbCrLf & _
        "Çeliþkili (BELIRSIZ): " & nBel & vbCrLf & _
        "Yalnýzca 0 fiyatlý (SIFIR): " & nSif
    If gecersiz > 0 Then ozet = ozet & vbCrLf & "Okunamayan fiyat satýrý: " & gecersiz
    ozet = ozet & vbCrLf & "Liste tarihi: " & Format$(tarih, "dd.mm.yyyy hh:nn")
    If nBel + nSif > 0 Then
        ozet = ozet & vbCrLf & vbCrLf & "BELIRSIZ ve SIFIR kayýtlar için SAP güncellemesinde fiyat uygulanmaz; ilgili satýrlar 'Fiyat belirsiz' olarak iþaretlenir." & _
            vbCrLf & "Ayrýntý: _FIYAT_LISTESI sayfasý (gizli, sað týk > Göster)."
    End If
    application.EnableEvents = ev: application.ScreenUpdating = sc: application.DisplayAlerts = al
    AKSORA_MUSTERI_FIYATLARI_UYGULA False
    AKSORA_MUSTERI_PANEL_YENILE
    MsgBox ozet, vbInformation, "AKSORA | Müþteri fiyat listesi"
    Exit Sub

Hata:
    ozet = "Fiyat listesi yüklenemedi: " & Err.Description
    On Error Resume Next
    If Not wbK Is Nothing Then wbK.Close SaveChanges:=False
    application.EnableEvents = ev: application.ScreenUpdating = sc: application.DisplayAlerts = al
    MsgBox ozet, vbExclamation, "AKSORA | Müþteri fiyat listesi"
End Sub
