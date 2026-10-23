Attribute VB_Name = "modKS_SAP"
Option Explicit

Private mAcikAl As Boolean     ' baþlangýçtan önce teslim tarihli ama teslimatý bitmemiþ sipariþler de alýnsýn
Private mDisAl As Boolean      ' bölge dýþý (Fýrat) kalemler de alýnsýn
Private mCAcik As Long         ' "Teslimatý yapýlacak (miktar)" sütunu (sayfa baþýna)

' Yýllýk SAP satýnalma listesi (ME2N biçimi; "ÖRNEK.xlsx" / "TEDARÝKÇÝ ÝSÝM.xlsx") aktarýmý.
' Baþlýklar adla bulunur. Ayný SAS belgesi + kalem + malzeme + teslimat tarihi ikinci kez aktarýlmaz;
' aktarým tekrar çalýþtýrýlabilir. Çoruh dýþý üretim yerleri (Fýrat bölgesi) sorulur, varsayýlan atlanýr.
' Baþlangýç tarihinden önce teslim tarihli ama teslimatý bitmemiþ (devreden) sipariþler de istenirse alýnýr.
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
    Dim nAcikOnce As Long, nDisSay As Long, soru As String
    Dim gecis As Long, toplamSatir As Long, sonSatir As Long, cUy As Long, dKodYer As Object, dDis As Object, yerAd As String
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
    KS_SuzgecTemizle T_SEVK
    KS_SuzgecTemizle T_MALZ
    KS_Ac KS_WS(SH_TANIM)
    KS_YerTamamla KS_LOZ(T_YER)
    Set dYer = KS_Sozluk(T_YER)
    Set dKodYer = KS_UretimYerleri()
    Set dDis = KS_BolgeDisiYerler()
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

    ' 0: tanýmsýz üretim yerleri ve sayýmlar, 1: yeni malzemeler, 2: sevkiyatlar
    mAcikAl = False
    mDisAl = False
    For gecis = 0 To 2
        If gecis = 1 And nAcikOnce > 0 Then
            KS_Bitir True
            mAcikAl = (MsgBox("Teslim tarihi " & Format$(baslangic, "dd.mm.yyyy") & " öncesinde olan ama teslimatý henüz " & _
                "tamamlanmamýþ " & nAcikOnce & " açýk sipariþ kalemi var (önceki yýllardan devreden sipariþler)." & vbCrLf & vbCrLf & _
                "EVET: bunlar da aktarýlýr (önerilen; kapsam tam olur)." & vbCrLf & _
                "HAYIR: yalnýz baþlangýç tarihinden sonrakiler aktarýlýr.", vbQuestion + vbYesNo, KS_AD) = vbYes)
            KS_Basla
        End If
        If gecis = 1 And nDisSay > 0 Then
            KS_Bitir True
            mDisAl = (MsgBox("Dosyada Çoruh bölgesi dýþýna (Fýrat bölgesi: FEDAÞ, Fýrat Aksa vb.) ait " & nDisSay & _
                " kalem var." & vbCrLf & vbCrLf & _
                "HAYIR: bölge dýþý kalemler atlanýr (önerilen)." & vbCrLf & _
                "EVET: bunlar da aktarýlýr.", vbQuestion + vbYesNo + vbDefaultButton2, KS_AD) = vbYes)
            KS_Basla
        End If
        If gecis = 1 And dYeniYer.Count > 0 Then
            liste = ""
            For Each k In dYeniYer.Keys
                If Len(liste) < 400 Then liste = liste & vbCrLf & "  - " & dYeniYer(k)
            Next k
            KS_Bitir True
            yerEkle = (MsgBox("Dosyada TANIMLAR'da olmayan " & dYeniYer.Count & " üretim yeri var:" & liste & vbCrLf & vbCrLf & _
                "Çoruh bölgesindeyse EVET: sevk yeri olarak eklenir ve kalemleri aktarýlýr." & vbCrLf & _
                "Bölge dýþýysa HAYIR: bu üretim yerlerinin kalemleri atlanýr." & vbCrLf & vbCrLf & _
                "(Bilinen Fýrat bölgesi üretim yerleri bu listede yer almaz; onlar için ayrýca soruldu.)", _
                vbQuestion + vbYesNo + vbDefaultButton2, KS_AD) = vbYes)
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
                cUy = KS_SapKolon(bas, "Üretim yeri")
                mCAcik = KS_SapKolon(bas, "Teslimatý yapýlacak (miktar)")
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
                        yerAd = KS_SapYer(v, i, cUy, cYer, dKodYer)
                        If gecis = 0 Then
                            If Len(KS_Anahtar(yerAd)) > 0 And Not dYer.Exists(KS_Anahtar(yerAd)) And Not KS_SapDisi(v, i, cUy) Then
                                dYeniYer(KS_Anahtar(yerAd)) = yerAd
                            End If
                            KS_SapSay v, i, cT, cSil, baslangic, KS_SapDisi(v, i, cUy) Or dDis.Exists(KS_Anahtar(yerAd)), _
                                nAcikOnce, nDisSay
                        ElseIf KS_SapSatirUygun(v, i, cT, yerAd, cSil, baslangic, dYer, dDis, KS_SapDisi(v, i, cUy), yerEkle, _
                                gecis = 2, nTarih, nBolge, nSil) Then
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
                                g(1) = yerAd
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
    KS_Kapat KS_WS(SH_TANIM)
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
    KS_KurGuncelle True
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

' Sevk yeri: üretim yeri kodu tanýmlýysa TANIMLAR'daki adý, deðilse dosyadaki üretim yeri adý.
Private Function KS_SapYer(ByRef v As Variant, ByVal i As Long, ByVal cUy As Long, ByVal cYer As Long, _
        ByVal dKodYer As Object) As String
    If cUy > 0 Then
        If dKodYer.Exists(KS_Anahtar(v(i, cUy))) Then
            KS_SapYer = dKodYer(KS_Anahtar(v(i, cUy)))
            Exit Function
        End If
    End If
    KS_SapYer = Application.WorksheetFunction.Trim(KS_Metin(v(i, cYer)))
End Function

Private Function KS_SapDisi(ByRef v As Variant, ByVal i As Long, ByVal cUy As Long) As Boolean
    If cUy > 0 Then KS_SapDisi = KS_BolgeDisiKod(v(i, cUy))
End Function

' Teslimatý tamamlanmamýþ kalem: "Teslimatý yapýlacak (miktar)" > 0.
Private Function KS_SapAcik(ByRef v As Variant, ByVal i As Long) As Boolean
    Dim x As Double
    If mCAcik > 0 Then
        If KS_Sayi(v(i, mCAcik), x) Then KS_SapAcik = (x > 0)
    End If
End Function

' Ýlk geçiþte soru için sayým: baþlangýçtan önceki açýk sipariþler (bölge içi) ve bölge dýþý kalemler.
Private Sub KS_SapSay(ByRef v As Variant, ByVal i As Long, ByVal cT As Long, ByVal cSil As Long, ByVal baslangic As Date, _
        ByVal disi As Boolean, ByRef nAcikOnce As Long, ByRef nDisSay As Long)
    Dim d As Date
    If cSil > 0 Then
        If Len(KS_Metin(v(i, cSil))) > 0 Then Exit Sub
    End If
    If Not KS_Tarih(v(i, cT), d) Then Exit Sub
    If disi Then
        If d >= baslangic Or KS_SapAcik(v, i) Then nDisSay = nDisSay + 1
    ElseIf d < baslangic And KS_SapAcik(v, i) Then
        nAcikOnce = nAcikOnce + 1
    End If
End Sub

Private Function KS_SapSatirUygun(ByRef v As Variant, ByVal i As Long, ByVal cT As Long, ByVal yerAd As String, _
        ByVal cSil As Long, ByVal baslangic As Date, ByVal dYer As Object, ByVal dDis As Object, ByVal disKod As Boolean, _
        ByVal yerEkle As Boolean, ByVal say As Boolean, _
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
        If Not (mAcikAl And KS_SapAcik(v, i)) Then
            If say Then nTarih = nTarih + 1
            Exit Function
        End If
    End If
    If (disKod Or dDis.Exists(KS_Anahtar(yerAd))) And Not mDisAl Then
        If say Then nBolge = nBolge + 1
        Exit Function
    End If
    If Not dYer.Exists(KS_Anahtar(yerAd)) And Not yerEkle Then
        If say Then nBolge = nBolge + 1
        Exit Function
    End If
    KS_SapSatirUygun = True
End Function

' ---------------------------------------------------------------- bölge, kur ve taným bakýmý

' Çoruh bölgesi illeri. Sevk yeri tablosunda Ýl'i bu listede olmayan (ve boþ olmayan) yerler bölge dýþýdýr.
Public Function KS_BolgeIli(ByVal il As Variant) As Boolean
    Dim x As Variant
    For Each x In Array("Artvin", "Rize", "Trabzon", "Giresun", "Gümüþhane")
        If StrComp(Trim$(KS_Metin(il)), CStr(x), vbTextCompare) = 0 Then
            KS_BolgeIli = True
            Exit Function
        End If
    Next x
End Function

' Bilinen Fýrat bölgesi üretim yerleri (FEDAÞ 720x, Fýrat Aksa 710x): tabloda olmasalar da aktarýlmaz.
Public Function KS_BolgeDisiKod(ByVal kod As Variant) As Boolean
    KS_BolgeDisiKod = (InStr(1, "|7101|7102|7103|7104|7201|7202|7203|7204|", "|" & KS_Metin(kod) & "|") > 0)
End Function

' Sevk yeri tablosunu tamamlar: bilinen yerlerin boþ üretim yeri kodu / il hücrelerini doldurur,
' eksik Çoruh yerlerini ekler. Kullanýcýnýn yazdýðý deðerlere dokunmaz.
Public Sub KS_YerTamamla(ByVal lo As ListObject)
    Dim t As Variant, a As Variant, i As Long, j As Long, bulundu As Boolean, yeni() As Variant, n As Long
    t = Array( _
        Array("ÇEDAÞ TRABZON", "7152", "Trabzon", 1), Array("ÇEDAÞ RÝZE", "7151", "Rize", 1), _
        Array("ÇEDAÞ ARTVÝN", "7155", "Artvin", 1), Array("ÇEDAÞ GÝRESUN", "7154", "Giresun", 1), _
        Array("ÇEDAÞ GÜMÜÞHANE", "7153", "Gümüþhane", 1), _
        Array("Çoruh Aksa- Rize ÜY", "7051", "Rize", 1), Array("Çoruh Aksa- Trabzon ÜY", "7052", "Trabzon", 1), _
        Array("Çoruh Aksa- Gümüþhane ÜY", "7053", "Gümüþhane", 1), Array("Çoruh Aksa- Giresun ÜY", "7054", "Giresun", 1), _
        Array("Çoruh Aksa- Artvin ÜY", "7055", "Artvin", 1), _
        Array("Fýrat Aksa- Bingöl ÜY", "7101", "Bingöl", 0), Array("Fýrat Aksa- Elâzýð ÜY", "7102", "Elazýð", 0), _
        Array("Fýrat Aksa- Malatya ÜY", "7103", "Malatya", 0), Array("Fýrat Aksa- Tunceli ÜY", "7104", "Tunceli", 0), _
        Array("FEDAÞ BÝNGÖL", "7201", "Bingöl", 0), Array("FEDAÞ ELAZIÐ", "7202", "Elazýð", 0), _
        Array("FEDAÞ MALATYA", "7203", "Malatya", 0), Array("FEDAÞ TUNCELÝ", "7204", "Tunceli", 0))
    If lo Is Nothing Then Exit Sub
    KS_Basla
    a = KS_Veri(lo)
    ReDim yeni(1 To UBound(t) + 1, 1 To lo.ListColumns.Count)
    For j = 0 To UBound(t)
        bulundu = False
        If IsArray(a) Then
            For i = 1 To UBound(a, 1)
                If StrComp(Trim$(KS_Metin(a(i, 1))), CStr(t(j)(0)), vbTextCompare) = 0 Then
                    bulundu = True
                    If KS_Bos(a(i, 2)) Then lo.DataBodyRange.Cells(i, 2).Value = t(j)(1)
                    If KS_Bos(a(i, 3)) Then lo.DataBodyRange.Cells(i, 3).Value = t(j)(2)
                    Exit For
                End If
            Next i
        End If
        If Not bulundu And t(j)(3) = 1 Then
            n = n + 1
            yeni(n, 1) = t(j)(0)
            yeni(n, 2) = t(j)(1)
            yeni(n, 3) = t(j)(2)
        End If
    Next j
    If n > 0 Then KS_TopluEkle lo, yeni, n
    KS_Bitir
End Sub

' Bölge dýþý sevk yerleri (anahtar: büyük harf ad).
Public Function KS_BolgeDisiYerler() As Object
    Dim d As Object, a As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    a = KS_Veri(KS_LOZ(T_YER))
    If IsArray(a) Then
        For i = 1 To UBound(a, 1)
            If Len(KS_Anahtar(a(i, 1))) > 0 Then
                If KS_BolgeDisiKod(a(i, 2)) Or (Not KS_Bos(a(i, 3)) And Not KS_BolgeIli(a(i, 3))) Then
                    d(KS_Anahtar(a(i, 1))) = KS_Metin(a(i, 1))
                End If
            End If
        Next i
    End If
    Set KS_BolgeDisiYerler = d
End Function

' Üretim yeri kodu -> sevk yeri adý.
Public Function KS_UretimYerleri() As Object
    Dim d As Object, a As Variant, i As Long
    Set d = CreateObject("Scripting.Dictionary")
    a = KS_Veri(KS_LOZ(T_YER))
    If IsArray(a) Then
        For i = 1 To UBound(a, 1)
            If Len(KS_Anahtar(a(i, 2))) > 0 And Len(KS_Anahtar(a(i, 1))) > 0 Then
                If Not d.Exists(KS_Anahtar(a(i, 2))) Then d(KS_Anahtar(a(i, 2))) = KS_Metin(a(i, 1))
            End If
        Next i
    End If
    Set KS_UretimYerleri = d
End Function

Public Sub KS_BOLGE_DISI_TEMIZLE()
    KS_BolgeDisiTemizle False
End Sub

' Bölge dýþý sevk yerlerine ait sevkiyatlarý ve yalnýz onlar için açýlmýþ SAP malzeme kartlarýný siler.
' Arýzaya baðlý sevkiyata dokunulmaz. Önce yedek alýnýr.
Public Sub KS_BolgeDisiTemizle(ByVal otomatik As Boolean)
    Dim dDis As Object, loS As ListObject, loM As ListObject, loA As ListObject, arS As Variant, aM As Variant, aA As Variant
    Dim dBagli As Object, dKod As Object, dSay As Object, tutS() As Boolean, tutM() As Boolean
    Dim i As Long, nS As Long, nM As Long, kalS As Long, kalM As Long, nBagli As Long, k As Variant, liste As String
    On Error GoTo Hata
    KS_YerTamamla KS_LOZ(T_YER)
    Set dDis = KS_BolgeDisiYerler()
    Set loS = KS_LOZ(T_SEVK)
    Set loM = KS_LOZ(T_MALZ)
    Set loA = KS_LOZ(T_ARIZA)
    arS = KS_Veri(loS)
    aM = KS_Veri(loM)
    aA = KS_Veri(loA)
    Set dBagli = CreateObject("Scripting.Dictionary")
    Set dSay = CreateObject("Scripting.Dictionary")
    Set dKod = CreateObject("Scripting.Dictionary")
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            dBagli(KS_Anahtar(aA(i, A_SEVK))) = True
            dKod(KS_Anahtar(aA(i, A_KOD))) = True
        Next i
    End If
    If IsArray(arS) Then
        ReDim tutS(1 To UBound(arS, 1))
        For i = 1 To UBound(arS, 1)
            tutS(i) = True
            If dDis.Exists(KS_Anahtar(arS(i, S_YER))) Then
                If dBagli.Exists(KS_Anahtar(arS(i, S_NO))) Then
                    nBagli = nBagli + 1
                Else
                    tutS(i) = False
                    nS = nS + 1
                    dSay(KS_Metin(arS(i, S_YER))) = dSay(KS_Metin(arS(i, S_YER))) + 1
                End If
            End If
            If tutS(i) Then
                kalS = kalS + 1
                dKod(KS_Anahtar(arS(i, S_KOD))) = True
            End If
        Next i
    End If
    If IsArray(aM) Then
        ReDim tutM(1 To UBound(aM, 1))
        For i = 1 To UBound(aM, 1)
            tutM(i) = True
            If nS > 0 And Not dKod.Exists(KS_Anahtar(aM(i, M_KOD))) And Left$(KS_Metin(aM(i, M_NOT)), 13) = "SAP aktarýmý " Then
                tutM(i) = False
                nM = nM + 1
            Else
                kalM = kalM + 1
            End If
        Next i
    End If
    If nS = 0 Then
        If Not otomatik Then MsgBox "Bölge dýþý (Çoruh dýþý) sevk yerine ait sevkiyat bulunmadý.", vbInformation, KS_AD
        Exit Sub
    End If
    For Each k In dSay.Keys
        liste = liste & vbCrLf & "  - " & k & ": " & dSay(k)
    Next k
    If MsgBox("Çoruh bölgesi dýþýndaki sevk yerlerine ait " & nS & " sevkiyat bulundu:" & liste & vbCrLf & vbCrLf & _
        "Bu sevkiyatlar ve yalnýz onlar için açýlmýþ " & nM & " malzeme kartý silinsin mi?" & vbCrLf & _
        IIf(nBagli > 0, "(Arýzaya baðlý " & nBagli & " sevkiyat korunur.)" & vbCrLf, "") & _
        "Önce dosyanýn yedeði alýnýr. Sonraki yýllýk aktarýmlarda bu yerler kendiliðinden atlanýr.", _
        vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    KS_YedekKopya "BOLGE_TEMIZLIK_ONCESI"
    KS_Basla
    KS_Ac loS.Parent
    KS_TabloSikistir loS, arS, tutS, kalS
    KS_Kapat loS.Parent
    If nM > 0 Then
        KS_Ac loM.Parent
        KS_TabloSikistir loM, aM, tutM, kalM
        KS_Kapat loM.Parent
    End If
    KS_Bitir True
    KS_Log "BÖLGE DIÞI SÝL", "", nS & " sevkiyat, " & nM & " malzeme kartý silindi"
    KS_PanelYenile False
    MsgBox nS & " bölge dýþý sevkiyat ve " & nM & " malzeme kartý silindi." & vbCrLf & _
        "Kalan sevkiyat: " & kalS, vbInformation, KS_AD
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Bölge dýþý temizlik durdu: " & Err.Description, vbExclamation, KS_AD
End Sub

' Tabloyu yalnýz tut(i) = True satýrlarla yeniden yazar (satýr satýr silmekten çok hýzlý).
Private Sub KS_TabloSikistir(ByVal lo As ListObject, ByRef a As Variant, ByRef tut() As Boolean, ByVal n As Long)
    Dim w() As Variant, i As Long, j As Long, k As Long, c As Long, eski As Long, ws As Worksheet, bas As Range
    c = lo.ListColumns.Count
    eski = UBound(a, 1)
    Set ws = lo.Parent
    Set bas = lo.HeaderRowRange.Cells(1, 1)
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    If n > 0 Then
        ReDim w(1 To n, 1 To c)
        For i = 1 To eski
            If tut(i) Then
                k = k + 1
                For j = 1 To c
                    w(k, j) = a(i, j)
                Next j
            End If
        Next i
        ws.Cells(bas.Row + 1, bas.Column).Resize(n, c).Value = w
        lo.Resize ws.Range(bas, ws.Cells(bas.Row + n, bas.Column + c - 1))
    Else
        lo.Resize ws.Range(bas, ws.Cells(bas.Row + 1, bas.Column + c - 1))
        lo.DataBodyRange.ClearContents
        n = 1
    End If
    If eski > n Then ws.Range(ws.Cells(bas.Row + n + 1, bas.Column), ws.Cells(bas.Row + eski, bas.Column + c - 1)).Clear
End Sub

' Kuru tanýmsýz döviz birimleri (verideki para birimlerine göre), virgülle.
Public Function KS_KurEksik() As String
    Dim a As Variant, i As Long, d As Object, pb As String
    Set d = CreateObject("Scripting.Dictionary")
    KS_ParamYukle
    a = KS_Veri(KS_LOZ(T_SEVK))
    If IsArray(a) Then
        For i = 1 To UBound(a, 1)
            pb = KS_Anahtar(a(i, S_PB))
            If Len(pb) > 0 And Not d.Exists(pb) Then
                If KS_Kur(pb) <= 0 Then d(pb) = True
            End If
        Next i
    End If
    If d.Count > 0 Then KS_KurEksik = Join(d.Keys, ", ")
End Function

Public Sub KS_KUR_GUNCELLE()
    KS_KurGuncelle False
End Sub

' TCMB günlük kurlarý (döviz satýþ) kur tablosuna yazýlýr, tüm TL tutarlar yeniden hesaplanýr.
Public Sub KS_KurGuncelle(ByVal sor As Boolean)
    Dim http As Object, xml As String, lo As ListObject, a As Variant, i As Long, x As Double
    Dim n As Long, tarih As String, liste As String, eksik As String
    On Error GoTo Hata
    If sor Then
        eksik = KS_KurEksik()
        If Len(eksik) = 0 Then Exit Sub
        If MsgBox("Kuru girilmemiþ para birimleri var: " & eksik & vbCrLf & _
            "Bu kalemlerin TL tutarý hesaplanamýyor." & vbCrLf & vbCrLf & _
            "TCMB'den bugünkü döviz satýþ kurlarý alýnsýn mý?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    End If
    Set http = CreateObject("MSXML2.XMLHTTP")
    http.Open "GET", "https://www.tcmb.gov.tr/kurlar/today.xml", False
    http.send
    If http.Status <> 200 Then Err.Raise vbObjectError + 513, , "TCMB yanýtý: " & http.Status
    xml = http.responseText
    tarih = KS_XmlOznitelik(xml, "Tarih=""")
    Set lo = KS_LOZ(T_KUR)
    a = KS_Veri(lo)
    If Not IsArray(a) Then Exit Sub
    KS_Ac lo.Parent
    For i = 1 To UBound(a, 1)
        Select Case KS_Anahtar(a(i, 1))
            Case "", "TRY", "TL"
            Case Else
                x = KS_TcmbKur(xml, KS_Anahtar(a(i, 1)))
                If x > 0 Then
                    lo.DataBodyRange.Cells(i, 2).Value = x
                    n = n + 1
                    liste = liste & vbCrLf & "  " & KS_Anahtar(a(i, 1)) & " = " & Format$(x, "0.0000")
                End If
        End Select
    Next i
    KS_Kapat lo.Parent
    If n = 0 Then Err.Raise vbObjectError + 514, , "Yanýtta kur bulunamadý."
    KS_Log "KUR GÜNCELLE", "TCMB " & tarih, Replace(Mid$(liste, 3), vbCrLf, " |")
    KS_PanelYenile False
    MsgBox "TCMB kurlarý (" & tarih & ") yazýldý:" & liste & vbCrLf & vbCrLf & _
        "TL tutarlar yeniden hesaplandý. Kurlar YÖNETÝM sayfasýndan elle de deðiþtirilebilir.", vbInformation, KS_AD
    Exit Sub
Hata:
    MsgBox "TCMB kurlarý alýnamadý: " & Err.Description & vbCrLf & vbCrLf & _
        "Ýnternet / proxy eriþimi yoksa kurlarý YÖNETÝM sayfasýndaki kur tablosuna elle yazýn.", vbExclamation, KS_AD
End Sub

Private Function KS_TcmbKur(ByVal xml As String, ByVal pb As String) As Double
    Dim p As Long, q As Long, s As String
    p = InStr(1, xml, "CurrencyCode=""" & pb & """", vbTextCompare)
    If p = 0 Then Exit Function
    p = InStr(p, xml, "<ForexSelling>", vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len("<ForexSelling>")
    q = InStr(p, xml, "</ForexSelling>", vbTextCompare)
    If q <= p Then Exit Function
    s = Trim$(Mid$(xml, p, q - p))
    If Len(s) > 0 Then KS_TcmbKur = Val(s)
End Function

Private Function KS_XmlOznitelik(ByVal xml As String, ByVal bas As String) As String
    Dim p As Long, q As Long
    p = InStr(1, xml, bas, vbTextCompare)
    If p = 0 Then Exit Function
    p = p + Len(bas)
    q = InStr(p, xml, """")
    If q > p Then KS_XmlOznitelik = Mid$(xml, p, q - p)
End Function

Public Sub KS_SEVK_ONAR()
    KS_SevkOnar False
End Sub

' Sevkiyat tablosu bütünlük onarýmý: ayný Sevk No ikinci kez geçen ya da hata deðeri (#N/A vb.) içeren
' satýrlar bozuk sayýlýr ve silinir; ilk (özgün) kayýt korunur. Önce yedek alýnýr.
Public Sub KS_SevkOnar(ByVal otomatik As Boolean)
    Dim lo As ListObject, a As Variant, i As Long, j As Long, d As Object, tut() As Boolean
    Dim n As Long, bozuk As Long, k As String, hatali As Boolean, bos As Boolean, korumali As Boolean
    On Error GoTo Hata
    Set lo = KS_LOZ(T_SEVK)
    korumali = lo.Parent.ProtectContents
    KS_Ac lo.Parent
    KS_SuzgecKaldir lo
    a = KS_Veri(lo)
    If Not IsArray(a) Then GoTo Bitti
    ReDim tut(1 To UBound(a, 1))
    Set d = CreateObject("Scripting.Dictionary")
    For i = 1 To UBound(a, 1)
        hatali = False
        bos = True
        For j = 1 To UBound(a, 2)
            If IsError(a(i, j)) Then
                hatali = True
            ElseIf Not KS_Bos(a(i, j)) Then
                bos = False
            End If
        Next j
        k = ""
        If Not IsError(a(i, S_NO)) Then k = KS_Anahtar(a(i, S_NO))
        If bos Then
            tut(i) = False
        ElseIf hatali Or Len(k) = 0 Or d.Exists(k) Then
            tut(i) = False
            bozuk = bozuk + 1
        Else
            tut(i) = True
            n = n + 1
            d(k) = True
        End If
    Next i
    If bozuk = 0 Then
        If Not otomatik Then MsgBox "Sevkiyat tablosunda bozuk satýr bulunmadý (" & n & " kayýt saðlam).", vbInformation, KS_AD
        GoTo Bitti
    End If
    KS_YedekKopya "SEVK_ONARIM_ONCESI"
    KS_Basla
    KS_TabloSikistir lo, a, tut, n
    KS_Bitir
    If korumali Then KS_Kapat lo.Parent
    KS_Log "SEVKÝYAT ONARIM", "", bozuk & " bozuk satýr silindi, " & n & " kayýt saðlam"
    KS_PanelYenile False
    If MsgBox(bozuk & " bozuk sevkiyat satýrý bulundu ve silindi; " & n & " kayýt saðlam." & vbCrLf & vbCrLf & _
        "Neden: liste süzülmüþken (ör. GARANTÝSÝ BÝTECEK) yapýlan yýllýk aktarýmda yeni satýrlar karýþmýþtý. " & _
        "Bu sürümde aktarým ve hesaplama öncesi süzgeç kendiliðinden kaldýrýlýyor." & vbCrLf & vbCrLf & _
        "Silinen kalemleri geri almak için yýllýk liste þimdi yeniden yüklensin mi? " & _
        "(Var olan kayýtlar atlanýr, yalnýz eksikler eklenir.)", vbQuestion + vbYesNo, KS_AD) = vbYes Then
        KS_SAP_AKTAR
    End If
    Exit Sub
Bitti:
    If korumali Then KS_Kapat lo.Parent
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Sevkiyat onarýmý durdu: " & Err.Description, vbExclamation, KS_AD
End Sub
