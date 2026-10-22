Attribute VB_Name = "modKS_SAP"
Option Explicit

' Yýllýk SAP satýnalma listesi (ME2N biçimi; "ÖRNEK.xlsx" / "TEDARÝKÇÝ ÝSÝM.xlsx") aktarýmý.
' Baþlýklar adla bulunur. Ayný SAS belgesi + kalem + malzeme ikinci kez aktarýlmaz; aktarým tekrar çalýþtýrýlabilir.
' SAP listesinde mal giriþ tarihi ve garanti süresi yoktur: garanti MALZEMELER / SEVKÝYATLAR sayfasýnda elle girilir.

Private Function KS_Norm(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    KS_Norm = UCase$(Application.WorksheetFunction.Trim(CStr(v)))
End Function

Private Function KS_SapKolon(ByVal baslik As Variant, ByVal ad As String) As Long
    Dim j As Long
    For j = LBound(baslik, 2) To UBound(baslik, 2)
        If KS_Norm(baslik(1, j)) = KS_Norm(ad) Then
            KS_SapKolon = j
            Exit Function
        End If
    Next j
End Function

Public Sub KS_SAP_AKTAR()
    Dim dosya As Variant, wb As Workbook, ws As Worksheet, v As Variant, bas As Variant, adim As String
    Dim cT As Long, cM As Long, cK As Long, cMik As Long, cGir As Long, cBr As Long, cYer As Long
    Dim cBel As Long, cKal As Long, cSat As Long, cFiy As Long, cPb As Long, cSil As Long, cTed As Long
    Dim baslangic As Date, d As Date, s As String, i As Long, j As Long, gelisSor As VbMsgBoxResult
    Dim dYer As Object, dYeniM As Object, dYeniYer As Object, mSat() As Variant, mN As Long, sSat() As Variant, sN As Long
    Dim g(0 To 11) As Variant, satir As Variant, hata As String, uyari As String, x As Double, mik As Double
    Dim nTarih As Long, nBolge As Long, nSil As Long, nVar As Long, nHata As Long, ilkHata As String
    Dim kod As String, pb As String, lo As ListObject, yerEkle As Boolean, sevkci As String, k As Variant, liste As String
    Dim gecis As Long, toplamSatir As Long, sonSatir As Long
    On Error GoTo Hata

    adim = "Dosya seçimi"
    dosya = Application.GetOpenFilename("Excel Dosyalarý (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , _
        "Yýllýk SAP satýnalma listesini seçin")
    If VarType(dosya) = vbBoolean Then Exit Sub
    s = InputBox("Hangi teslimat tarihinden itibaren aktarýlsýn? (gg.aa.yyyy)", KS_AD, "01.01." & Year(Date))
    If Len(s) = 0 Then Exit Sub
    If Not KS_Tarih(s, baslangic) Then
        MsgBox "Tarih anlaþýlamadý: " & s, vbExclamation, KS_AD
        Exit Sub
    End If
    gelisSor = MsgBox("SAP listesinde mal giriþ (gerçek geliþ) tarihi bulunmaz." & vbCrLf & vbCrLf & _
        "Girilen miktarý sipariþ miktarýna ulaþmýþ ve teslimat tarihi geçmiþ kalemlerde, geliþ tarihi olarak " & _
        "SAP teslimat tarihi yazýlsýn mý?" & vbCrLf & vbCrLf & _
        "EVET: garanti hemen iþler, baþlangýç yaklaþýktýr (notta belirtilir)." & vbCrLf & _
        "HAYIR: geliþ tarihi boþ kalýr; SEVKÝYATLAR sayfasýnda elle girersiniz.", vbQuestion + vbYesNoCancel, KS_AD)
    If gelisSor = vbCancel Then Exit Sub

    adim = "Yedek"
    KS_YedekKopya "SAP_ONCESI"
    KS_Basla
    Set dYer = KS_Sozluk(T_YER)
    Set dYeniM = CreateObject("Scripting.Dictionary")
    Set dYeniYer = CreateObject("Scripting.Dictionary")
    adim = "SAP dosyasýný açma"
    Set wb = Workbooks.Open(Filename:=CStr(dosya), ReadOnly:=True, UpdateLinks:=0)
    For Each ws In wb.Worksheets
        toplamSatir = toplamSatir + ws.UsedRange.Row + ws.UsedRange.Rows.Count
    Next ws
    ReDim mSat(1 To toplamSatir, 1 To M_KOLON)
    ReDim sSat(1 To toplamSatir, 1 To S_KOLON)
    KS_BaglamYukle

    ' 0: tanýmsýz üretim yerleri, 1: yeni malzemeler, 2: sevkiyatlar
    For gecis = 0 To 2
        If gecis = 1 And dYeniYer.Count > 0 Then
            liste = ""
            For Each k In dYeniYer.Keys
                If Len(liste) < 400 Then liste = liste & vbCrLf & "  - " & dYeniYer(k)
            Next k
            KS_Bitir True
            yerEkle = (MsgBox("Dosyada TANIMLAR'da olmayan " & dYeniYer.Count & " üretim yeri var:" & liste & vbCrLf & vbCrLf & _
                "EVET: bunlar sevk yeri olarak eklenir ve kalemleri aktarýlýr." & vbCrLf & _
                "HAYIR: bu üretim yerlerinin kalemleri atlanýr (bölge dýþý).", vbQuestion + vbYesNo, KS_AD) = vbYes)
            KS_Basla
        End If
        If gecis = 2 Then
            If mN > 0 Then
                adim = "Yeni malzemeleri yazma"
                Set lo = KS_LOZ(T_MALZ)
                KS_Ac lo.Parent
                KS_TopluEkle lo, mSat, mN
                KS_Kapat lo.Parent
            End If
            KS_BaglamYukle
            KS_OtoTanimAyarla True
        End If
        For Each ws In wb.Worksheets
            adim = "Sayfa okuma: " & ws.Name
            If Application.WorksheetFunction.CountA(ws.Rows(1)) > 3 Then
                bas = ws.Range(ws.Cells(1, 1), ws.Cells(1, ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column)).Value
                cSat = KS_SapKolon(bas, "Satýcý/teslimat yapan ürt.yeri")
                cTed = KS_SapKolon(bas, "Tedarikçi adý")
                cM = KS_SapKolon(bas, "Malzeme")
                cT = KS_SapKolon(bas, "Teslimat tarihi")
                cYer = KS_SapKolon(bas, "Üretim Yeri Adý")
                If (cSat > 0 Or cTed > 0) And cM > 0 And cT > 0 And cYer > 0 Then
                    cK = KS_SapKolon(bas, "Kýsa metin")
                    cMik = KS_SapKolon(bas, "Miktar (depo ÖB)")
                    cGir = KS_SapKolon(bas, "Girilen miktar")
                    cBr = KS_SapKolon(bas, "SAS ölçü birimi")
                    cBel = KS_SapKolon(bas, "Satýnalma belgesi")
                    cKal = KS_SapKolon(bas, "Kalem")
                    cFiy = KS_SapKolon(bas, "Net fiyat")
                    cPb = KS_SapKolon(bas, "Para birimi")
                    cSil = KS_SapKolon(bas, "Silme göstergesi")
                    sonSatir = ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1
                    If sonSatir < 2 Then sonSatir = 2
                    v = ws.Range(ws.Cells(1, 1), ws.Cells(sonSatir, UBound(bas, 2))).Value
                    For i = 2 To UBound(v, 1)
                        If gecis = 0 Then
                            If Len(KS_Anahtar(v(i, cYer))) > 0 And Not dYer.Exists(KS_Anahtar(v(i, cYer))) Then
                                dYeniYer(KS_Anahtar(v(i, cYer))) = KS_Metin(v(i, cYer))
                            End If
                        ElseIf KS_SapSatirUygun(v, i, cT, cYer, cSil, baslangic, dYer, yerEkle, gecis = 2, nTarih, nBolge, nSil) Then
                            kod = KS_Metin(v(i, cM))
                            pb = ""
                            If cPb > 0 Then pb = UCase$(KS_Metin(v(i, cPb)))
                            If Len(pb) = 0 Then pb = "TRY"
                            If gecis = 1 Then
                                If Len(kod) > 0 And Not KS_MalzemeVar(kod) And Not dYeniM.Exists(KS_Anahtar(kod)) Then
                                    mN = mN + 1
                                    dYeniM(KS_Anahtar(kod)) = mN
                                    mSat(mN, M_KOD) = kod
                                    If cK > 0 Then mSat(mN, M_AD) = KS_Metin(v(i, cK))
                                    If cBr > 0 Then mSat(mN, M_BIRIM) = KS_Metin(v(i, cBr))
                                    mSat(mN, M_URETICI) = KS_SapSevkci(v, i, cSat, cTed)
                                    If cFiy > 0 Then mSat(mN, M_FIYAT) = v(i, cFiy)
                                    mSat(mN, M_PB) = pb
                                    mSat(mN, M_KAYIT) = Date
                                    mSat(mN, M_NOT) = "SAP aktarýmý - garanti süresini girin"
                                End If
                            Else
                                adim = ws.Name & " satýr " & i
                                sevkci = KS_SapSevkci(v, i, cSat, cTed)
                                g(0) = sevkci
                                g(1) = KS_Metin(v(i, cYer))
                                g(2) = kod
                                If cMik > 0 Then g(3) = v(i, cMik) Else g(3) = Empty
                                If cFiy > 0 Then g(4) = v(i, cFiy) Else g(4) = Empty
                                g(5) = pb
                                g(6) = v(i, cT)
                                g(7) = Empty
                                If gelisSor = vbYes And cGir > 0 And cMik > 0 Then
                                    If KS_Sayi(v(i, cGir), x) And KS_Sayi(v(i, cMik), mik) And KS_Tarih(v(i, cT), d) Then
                                        If x >= mik And mik > 0 And d <= Date Then g(7) = d
                                    End If
                                End If
                                g(8) = Empty
                                g(9) = "SAS " & KS_Metin(v(i, cBel)) & "/" & KS_Metin(v(i, cKal))
                                g(10) = Empty
                                g(11) = "SAP | girilen " & KS_Metin(v(i, cGir)) & " / " & KS_Metin(v(i, cMik))
                                If Not IsEmpty(g(7)) Then g(11) = g(11) & " | geliþ = SAP teslimat tarihi (yaklaþýk)"
                                If KS_SevkHazirla(g, 0, satir, hata, uyari) Then
                                    If Len(uyari) > 0 Then
                                        nVar = nVar + 1
                                        KS_SevkSayacGeri
                                    Else
                                        sN = sN + 1
                                        For j = 1 To S_KOLON
                                            sSat(sN, j) = satir(1, j)
                                        Next j
                                    End If
                                Else
                                    nHata = nHata + 1
                                    If Len(ilkHata) = 0 Then ilkHata = ws.Name & " satýr " & i & ": " & hata
                                End If
                            End If
                        End If
                    Next i
                End If
            End If
        Next ws
    Next gecis

    adim = "Sevkiyatlarý yazma"
    wb.Close SaveChanges:=False
    Set wb = Nothing
    If sN > 0 Then
        Set lo = KS_LOZ(T_SEVK)
        KS_Ac lo.Parent
        KS_TopluEkle lo, sSat, sN
        KS_Kapat lo.Parent
    End If
    KS_OtoTanimAyarla False
    KS_HesaplaTumu
    KS_Bitir True
    KS_Log "SAP AKTAR", CStr(dosya), sN & " sevkiyat, " & mN & " yeni malzeme, " & nVar & " zaten vardý, " & nHata & " hatalý"
    MsgBox "Yýllýk liste aktarýmý tamamlandý." & vbCrLf & vbCrLf & _
        "Yeni sevkiyat: " & sN & vbCrLf & _
        "Yeni malzeme kartý: " & mN & vbCrLf & _
        "Daha önce aktarýlmýþ (atlandý): " & nVar & vbCrLf & _
        "Baþlangýç tarihinden önce: " & nTarih & vbCrLf & _
        "Atlanan üretim yeri: " & nBolge & vbCrLf & _
        "Silme göstergeli: " & nSil & vbCrLf & _
        "Hatalý: " & nHata & IIf(Len(ilkHata) > 0, vbCrLf & "Ýlk hata: " & ilkHata, "") & vbCrLf & vbCrLf & _
        "Sýradaki adým: MALZEMELER > GARANTÝSÝ GÝRÝLMEMÝÞ ile garanti sürelerini girin." & vbCrLf & _
        "Döviz kalemleri için YÖNETÝM sayfasýndaki kur tablosunu doldurun.", _
        IIf(nHata > 0, vbExclamation, vbInformation), KS_AD
    Exit Sub
Hata:
    KS_OtoTanimAyarla False
    KS_Bitir True
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "SAP aktarýmý durdu." & vbCrLf & "Adým: " & adim & vbCrLf & "Hata: " & Err.Description, vbCritical, KS_AD
End Sub

' Sevkçi deðeri SAP biçiminde: "1300001416 LUNA ELEKTRÝK ELEKTRONÝK". Satýcý sütunu boþsa Tedarikçi adý kullanýlýr.
Private Function KS_SapSevkci(ByRef v As Variant, ByVal i As Long, ByVal cSat As Long, ByVal cTed As Long) As String
    If cSat > 0 Then KS_SapSevkci = Application.WorksheetFunction.Trim(KS_Metin(v(i, cSat)))
    If Len(KS_SapSevkci) = 0 And cTed > 0 Then KS_SapSevkci = Application.WorksheetFunction.Trim(KS_Metin(v(i, cTed)))
End Function

Private Function KS_SapSatirUygun(ByRef v As Variant, ByVal i As Long, ByVal cT As Long, ByVal cYer As Long, _
        ByVal cSil As Long, ByVal baslangic As Date, ByVal dYer As Object, ByVal yerEkle As Boolean, ByVal say As Boolean, _
        ByRef nTarih As Long, ByRef nBolge As Long, ByRef nSil As Long) As Boolean
    Dim d As Date
    If cSil > 0 Then
        If Len(KS_Metin(v(i, cSil))) > 0 Then
            If say Then nSil = nSil + 1
            Exit Function
        End If
    End If
    If Not KS_Tarih(v(i, cT), d) Then
        If say Then nTarih = nTarih + 1
        Exit Function
    End If
    If d < baslangic Then
        If say Then nTarih = nTarih + 1
        Exit Function
    End If
    If Not dYer.Exists(KS_Anahtar(v(i, cYer))) And Not yerEkle Then
        If say Then nBolge = nBolge + 1
        Exit Function
    End If
    KS_SapSatirUygun = True
End Function
