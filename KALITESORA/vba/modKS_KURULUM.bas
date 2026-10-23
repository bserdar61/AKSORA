Attribute VB_Name = "modKS_KURULUM"
Option Explicit

' KALÝTESORA kurulumu. Boþ bir .xlsm dosyasýnda tüm sayfalarý, tablolarý, formu ve paneli oluþturur.
' Tekrar çalýþtýrýldýðýnda veri tablolarýna dokunmaz; yalnýz eksikleri tamamlar ve tasarýmý yeniler.

Private Const ETIKET As String = "_KS_KURULU"
Public Const KS_TASARIM_NO As Long = 23      ' tasarým sürümü; açýlýþta eskiyse kendiliðinden yenilenir

Private mSessiz As Boolean

Public Function KS_Kurulu() As Boolean
    KS_Kurulu = KS_AdVar(ETIKET)
End Function

Public Function KS_KuruluTasarim() As Long
    Dim s As String
    On Error Resume Next
    s = Mid$(ThisWorkbook.Names(ETIKET).RefersTo, 2)
    If IsNumeric(s) Then KS_KuruluTasarim = CLng(s)
    On Error GoTo 0
End Function

' Açýlýþta eski tasarým bulunursa onaysýz onarým (yedek yine alýnýr).
Public Sub KS_TasarimGuncelle()
    mSessiz = True
    KALITESORA_KUR
    mSessiz = False
End Sub

Public Sub KALITESORA_KUR()
    Dim adim As String, yedek As String, ilk As Boolean, ad As Variant, ws As Worksheet
    Dim onceki As Long, gorselYok As Boolean
    On Error GoTo Hata

    If Len(ThisWorkbook.Path) = 0 Or LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        MsgBox "Önce dosyayý 'Makro Ýçerebilen Excel Çalýþma Kitabý (.xlsm)' olarak kaydedin, " & _
            "sonra kurulumu çalýþtýrýn.", vbExclamation, KS_AD
        Exit Sub
    End If
    If ThisWorkbook.ReadOnly Then
        MsgBox "Dosya salt okunur açýk. Yazýlabilir kopyada çalýþtýrýn.", vbExclamation, KS_AD
        Exit Sub
    End If
    gorselYok = (Len(KS_GorselYol("KALITESORA_BANNER.png")) = 0)

    ilk = Not KS_Kurulu()
    If ilk Then
        If MsgBox(KS_AD & " bu dosyaya kurulacak." & vbCrLf & vbCrLf & _
            "Sayfalar, veri tablolarý, kayýt formu, panel ve raporlar oluþturulacak." & _
            IIf(gorselYok, vbCrLf & vbCrLf & "UYARI: " & KS_KLASOR & " klasörü dosyanýn yanýnda bulunamadý; " & _
                "tasarým görselsiz kurulacak.", "") & vbCrLf & vbCrLf & _
            "Devam edilsin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    Else
        If Not mSessiz Then
            If MsgBox(KS_AD & " zaten kurulu." & vbCrLf & vbCrLf & _
                "Onarým modunda: kayýtlar korunur, eksik parçalar tamamlanýr, tüm sayfalarýn tasarýmý yenilenir." & _
                vbCrLf & "Önce dosyanýn yedeði alýnacak. Devam edilsin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
        End If
        adim = "Kurulum öncesi yedek"
        yedek = KS_YedekKopya("ONARIM_ONCESI")
    End If

    KS_Basla
    adim = "Sayfalarý hazýrlama"
    KS_Durum adim
    For Each ad In KS_Sayfalar()
        Set ws = KS_WS(CStr(ad))
        If Not ws Is Nothing Then KS_Ac ws
    Next ad
    KS_SayfalariHazirla ilk

    adim = "Tablo stili"
    KS_TabloStiliKur

    adim = "TANIMLAR sayfasý"
    KS_Durum adim
    KS_TanimlarKur

    adim = "Veri tablolarý"
    KS_Durum adim
    KS_VeriSayfalariKur

    adim = "Adlandýrýlmýþ listeler"
    KS_AdlariKur

    adim = "RAPORLAR sayfasý"
    KS_Durum adim
    KS_RaporKur

    adim = "KAYIT formu"
    KS_Durum adim
    KS_KayitFormKur

    adim = "KARNE sayfasý"
    KS_KarneKur

    adim = "TOPLU AKTAR sayfasý"
    KS_AktarKur

    adim = "PANEL sayfasý"
    KS_Durum adim
    KS_PanelKur

    adim = "Kurulum kaydý"
    KS_AdTanimla ETIKET, "=" & KS_TASARIM_NO
    ThisWorkbook.Names(ETIKET).Visible = False
    If Not KS_AdVar("_KS_KORUMA") Then
        KS_AdTanimla "_KS_KORUMA", "=1"
        ThisWorkbook.Names("_KS_KORUMA").Visible = False
    End If

    adim = "Hesaplama ve panel"
    KS_Durum adim
    KS_HesaplaTumu
    KS_PanelYenile False

    adim = "Sayfa korumasý"
    For Each ad In KS_Sayfalar()
        KS_Kapat KS_WS(CStr(ad))
    Next ad

    KS_Bitir True
    KS_Log IIf(ilk, "KURULUM", "ONARIM"), "v" & KS_SURUM, IIf(Len(yedek) > 0, "Yedek: " & yedek, "")
    KS_WS(SH_PANEL).Activate
    ActiveWindow.ScrollRow = 1
    KS_UygulamaModuUygula
    KS_EkranSigdir KS_WS(SH_PANEL)
    If mSessiz Then
        MsgBox KS_AD & " tasarýmý güncellendi (v" & KS_SURUM & ")." & vbCrLf & _
            "Kayýtlarýnýz korundu. Önceki dosyanýn yedeði: " & vbCrLf & yedek, vbInformation, KS_AD
        Exit Sub
    End If
    MsgBox KS_AD & " " & IIf(ilk, "kuruldu.", "onarýldý.") & vbCrLf & vbCrLf & _
        "Baþlangýç:" & vbCrLf & _
        "1) SEVKÝYATLAR > YILLIK EXCEL YÜKLE ile yýl baþý listesini alýn." & vbCrLf & _
        "2) MALZEMELER sayfasýnda garanti sürelerini girin." & vbCrLf & _
        "3) Arýzalarý KAYIT > ARIZA sekmesinden girin." & vbCrLf & vbCrLf & _
        "Sayfa koruma þifresi: " & KS_SIFRE & _
        IIf(Len(yedek) > 0, vbCrLf & "Onarým öncesi yedek: " & yedek, ""), vbInformation, KS_AD
    Exit Sub

Hata:
    onceki = Err.Number
    KS_Bitir True
    MsgBox "Kurulum tamamlanamadý." & vbCrLf & _
        "Adým: " & adim & vbCrLf & _
        "Hata: " & onceki & " / " & Err.Description & _
        IIf(Len(yedek) > 0, vbCrLf & "Yedek: " & yedek, ""), vbCritical, KS_AD
End Sub

' ---------------------------------------------------------------- sayfalar

Private Sub KS_SayfalariHazirla(ByVal ilk As Boolean)
    Dim ad As Variant, ws As Worksheet, onceki As Worksheet, i As Long, uyari As Boolean
    For Each ad In KS_Sayfalar()
        Set ws = KS_WS(CStr(ad))
        If ws Is Nothing Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            ws.Name = CStr(ad)
        End If
        If onceki Is Nothing Then
            ws.Move Before:=ThisWorkbook.Worksheets(1)
        Else
            ws.Move After:=onceki
        End If
        Set onceki = KS_WS(CStr(ad))
        Select Case CStr(ad)
            Case SH_PANEL: onceki.Tab.Color = C_MAVI()
            Case SH_KAYIT: onceki.Tab.Color = C_YESIL2()
            Case SH_SEVK, SH_ARIZA, SH_MALZ: onceki.Tab.Color = C_LAC()
            Case SH_KARNE, SH_RAPOR: onceki.Tab.Color = C_TEAL()
            Case Else: onceki.Tab.Color = C_GRI()
        End Select
    Next ad
    If ilk Then
        uyari = Application.DisplayAlerts
        Application.DisplayAlerts = False
        For i = ThisWorkbook.Worksheets.Count To 1 Step -1
            Set ws = ThisWorkbook.Worksheets(i)
            If Not KS_BizimSayfa(ws.Name) Then
                If Application.WorksheetFunction.CountA(ws.Cells) = 0 And ws.Shapes.Count = 0 Then ws.Delete
            End If
        Next i
        Application.DisplayAlerts = uyari
    End If
End Sub

' Tabloyu oluþturur ya da mevcut tablonun baþlýklarýný doðrular. yeni = True ise tablo yeni kuruldu.
Private Function KS_TabloKur(ByVal ws As Worksheet, ByVal ad As String, ByVal sol As Range, _
        ByVal basliklar As Variant, ByVal genislik As Variant, ByVal bicim As Variant, _
        ByRef yeni As Boolean) As ListObject
    Dim lo As ListObject, n As Long, i As Long
    n = UBound(basliklar) - LBound(basliklar) + 1
    For i = 0 To n - 1
        ws.Columns(sol.Column + i).ColumnWidth = genislik(LBound(genislik) + i)
        ws.Range(ws.Cells(sol.Row + 1, sol.Column + i), ws.Cells(ws.Rows.Count, sol.Column + i)).NumberFormat = _
            bicim(LBound(bicim) + i)
    Next i
    Set lo = KS_LO(ad)
    yeni = lo Is Nothing
    If yeni Then
        sol.Resize(1, n).Value = basliklar
        Set lo = ws.ListObjects.Add(xlSrcRange, sol.Resize(2, n), , xlYes)
        lo.Name = ad
        For i = lo.ListRows.Count To 1 Step -1
            If Application.WorksheetFunction.CountA(lo.ListRows(i).Range) = 0 Then lo.ListRows(i).Delete
        Next i
    Else
        If lo.ListColumns.Count <> n Then
            Err.Raise vbObjectError + 7101, KS_AD, ad & " tablosunun sütun sayýsý beklenenden farklý (" & _
                lo.ListColumns.Count & " / " & n & "). Kurulum veriye dokunmadan durduruldu."
        End If
        For i = 1 To n
            If StrComp(lo.HeaderRowRange.Cells(1, i).Value, basliklar(LBound(basliklar) + i - 1), vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 7102, KS_AD, ad & " tablosunda " & i & ". sütun baþlýðý farklý: '" & _
                    lo.HeaderRowRange.Cells(1, i).Value & "'. Kurulum veriye dokunmadan durduruldu."
            End If
        Next i
    End If
    lo.TableStyle = KS_STIL
    lo.ShowTableStyleRowStripes = True
    With lo.HeaderRowRange
        .WrapText = True
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Font.Bold = True
        .Font.Size = 9
        .Interior.Pattern = xlNone
    End With
    lo.HeaderRowRange.EntireRow.RowHeight = 36
    Set KS_TabloKur = lo
End Function

' Hesaplanan sütunlarýn baþlýðý koyu turkuaz, elle girilebilen sütunlarýn baþlýðý amber.
Private Sub KS_BaslikRenk(ByVal lo As ListObject, ByVal kolonlar As Variant, ByVal renk As Long)
    Dim k As Variant
    ' Tüm baþlýklar tek renk (tablo stili). Eski sürümün renkli baþlýklarý temizlenir.
    For Each k In kolonlar
        With lo.HeaderRowRange.Cells(1, CLng(k))
            .Interior.Pattern = xlNone
            .Font.Color = vbWhite
        End With
    Next k
End Sub

' Elle düzenlenebilir sütunlar: kilit açýk, açýk sarý zemin.
Private Sub KS_DuzenlenebilirSutun(ByVal lo As ListObject, ByVal kolonlar As Variant)
    Dim k As Variant
    For Each k In kolonlar
        With lo.HeaderRowRange.Cells(1, CLng(k))
            .Interior.Pattern = xlNone
            .Font.Color = vbWhite
        End With
    Next k
    KS_KilitAyarla lo, kolonlar
End Sub

' Elle yazýlan tablo sütunlarýna giriþ doðrulamasý (tarih / sayý).
Private Sub KS_SutunDogrulama(ByVal lo As ListObject, ByVal kolon As Long, ByVal tur As String)
    Dim ws As Worksheet, c As Long, r As Range
    Set ws = lo.Parent
    c = lo.HeaderRowRange.Cells(1, kolon).Column
    Set r = ws.Range(ws.Cells(lo.HeaderRowRange.Row + 1, c), ws.Cells(lo.HeaderRowRange.Row + 100000, c))
    If tur = "tarih" Then KS_TarihDogrulama r Else KS_SayiDogrulama r
End Sub

Private Sub KS_VeriSayfaGenel(ByVal ws As Worksheet, ByVal ilkVeri As Long)
    ws.Cells.Font.Name = KS_FONT
    ws.Cells.Font.Size = 9.5
    ws.Columns(1).ColumnWidth = 2
    ws.Rows(ilkVeri & ":" & ws.Rows.Count).RowHeight = 19
    ws.Cells.VerticalAlignment = xlCenter
End Sub

' v2.1 ve öncesinde tablo 3. satýrdaydý; grup bandý için bir satýr aþaðý alýnýr (veri korunur).
Private Sub KS_TabloAsagiAl(ByVal ad As String)
    Dim lo As ListObject
    Set lo = KS_LO(ad)
    If lo Is Nothing Then Exit Sub
    If lo.HeaderRowRange.Row = 3 Then lo.Parent.Rows(3).Insert
End Sub

' Baþlýk gruplarý: üstte etiketli bant, baþlýk hücreleri grubun tonunda; gruplar arasý ince beyaz ayraç.
' gruplar: Array(Array("ETÝKET", ilkSütun, sonSütun, renk), ...)
Private Sub KS_GrupBandi(ByVal lo As ListObject, ByVal gruplar As Variant)
    Dim ws As Worksheet, g As Variant, r As Long, c0 As Long, bant As Range, bas As Range
    Set ws = lo.Parent
    r = lo.HeaderRowRange.Row - 1
    c0 = lo.HeaderRowRange.Column - 1
    With ws.Range(ws.Cells(r, c0 + 1), ws.Cells(r, c0 + lo.ListColumns.Count))
        .UnMerge
        .Clear
        .Font.Name = KS_FONT_KALIN
    End With
    ws.Rows(r).RowHeight = 20
    For Each g In gruplar
        Set bant = ws.Range(ws.Cells(r, c0 + g(1)), ws.Cells(r, c0 + g(2)))
        Set bas = ws.Range(ws.Cells(r + 1, c0 + g(1)), ws.Cells(r + 1, c0 + g(2)))
        bant.Merge
        With bant
            .Value = g(0)
            .Interior.Color = g(3)
            .Font.Color = RGB(205, 228, 245)
            .Font.Bold = True
            .Font.Size = 8
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        With bas
            .Interior.Color = g(3)
            .Font.Color = vbWhite
        End With
        With ws.Range(ws.Cells(r, c0 + g(1)), ws.Cells(r + 1, c0 + g(1))).Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Color = vbWhite
            .Weight = xlMedium
        End With
        With bant.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Color = RGB(0, 180, 216)
            .Weight = xlThin
        End With
    Next g
End Sub

' Özet görünüm: ikincil sütunlar gizlenir, tablo banner geniþliðine oturur.
Public Sub KS_OzetGorunum(ByVal lo As ListObject, ByVal gizli As Variant, ByVal ozet As Boolean)
    Dim k As Variant, ws As Worksheet
    Set ws = lo.Parent
    For Each k In gizli
        ws.Columns(lo.HeaderRowRange.Cells(1, CLng(k)).Column).Hidden = ozet
    Next k
End Sub

Public Function KS_GizliSutunlar(ByVal tablo As String) As Variant
    Select Case tablo
        Case T_SEVK: KS_GizliSutunlar = Array(S_FIYAT, S_PB, S_GECIKME, S_IRS, S_SERI, S_ARIZA, S_NOT, S_ZAMAN, S_KULL)
        Case T_ARIZA: KS_GizliSutunlar = Array(A_SEVKCI, A_SEVK, A_GARBIT, A_KAPKAYNAK, A_ADET, A_TIP, A_BILDIRIM, _
            A_MAIL, A_MAILDOSYA, A_DONUS, A_DEVAM, A_COZULDU, A_COZUM, A_COZUMGUN, A_SEKIL, A_KARS, A_ONCEKI, _
            A_TEKRAR, A_NOT, A_ZAMAN, A_KULL)
        Case Else: KS_GizliSutunlar = Array(M_NOT, M_KAYIT)
    End Select
End Function

' Düðme: özet / tüm sütunlar arasýnda geçiþ.
Public Sub KS_SUTUN_GORUNUM()
    Dim lo As ListObject, ws As Worksheet, korumali As Boolean, ozet As Boolean, g As Variant, s As Shape
    On Error GoTo Hata
    Set ws = ActiveSheet
    Select Case ws.Name
        Case SH_SEVK: Set lo = KS_LOZ(T_SEVK)
        Case SH_ARIZA: Set lo = KS_LOZ(T_ARIZA)
        Case SH_MALZ: Set lo = KS_LOZ(T_MALZ)
        Case Else: Exit Sub
    End Select
    g = KS_GizliSutunlar(lo.Name)
    ozet = Not ws.Columns(lo.HeaderRowRange.Cells(1, CLng(g(0))).Column).Hidden
    korumali = ws.ProtectContents
    KS_Ac ws
    KS_OzetGorunum lo, g, ozet
    If korumali Then KS_Kapat ws
    On Error Resume Next
    Set s = ws.Shapes("ksB_Gorunum")
    If Not s Is Nothing Then s.TextFrame2.TextRange.Text = IIf(ozet, "TÜM SÜTUNLAR", "ÖZET GÖRÜNÜM")
    On Error GoTo 0
    ActiveWindow.ScrollColumn = 1
    Exit Sub
Hata:
    MsgBox "Görünüm deðiþtirilemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Private Sub KS_GorunumDugmesi(ByVal ws As Worksheet)
    Dim x As Double
    x = ws.Cells(1, 2).Left + KS_GENISLIK - 112
    KS_Buton ws, "ksB_Gorunum", "TÜM SÜTUNLAR", x, ws.Rows(2).Top + 9, 112, 26, "KS_SUTUN_GORUNUM", "cizgi"
End Sub

Public Sub KS_VeriSayfaUst(ByVal ws As Worksheet, ByVal baslik As String, ByVal aciklama As String, ByVal dugmeler As Variant)
    Dim alt As Double, sol As Double, i As Long, n As Long, gen As Double, x As Double
    KS_SekilSil ws, "ksB_"
    alt = KS_SayfaCerceve(ws, False)
    ws.Rows(1).RowHeight = alt + 6
    ws.Rows(2).RowHeight = 44
    sol = ws.Cells(1, 2).Left
    x = sol + KS_GENISLIK
    If IsArray(dugmeler) Then
        n = (UBound(dugmeler) + 1) \ 3
        gen = 112
        x = sol + KS_GENISLIK - n * gen - (n - 1) * 6
    End If
    KS_SayfaBaslik ws, ws.Rows(2).Top + 4, baslik, aciklama, x - sol - 12
    If IsArray(dugmeler) Then
        For i = 0 To n - 1
            If Len(dugmeler(i * 3)) > 0 Then
                KS_Buton ws, "ksB_" & i, CStr(dugmeler(i * 3)), x + i * (gen + 6), ws.Rows(2).Top + 9, gen, 26, _
                    CStr(dugmeler(i * 3 + 1)), CStr(dugmeler(i * 3 + 2))
            End If
        Next i
    End If
End Sub

Private Sub KS_VeriSayfalariKur()
    Dim ws As Worksheet, lo As ListObject, yeni As Boolean, tl As String, tr As String
    Dim gLac As Long, gCelik As Long, gTeal As Long, gYesil As Long, gGri As Long, gMor As Long
    tl = "#,##0.00"
    tr = KS_FmtTarih()
    gLac = RGB(12, 45, 79)
    gCelik = RGB(31, 72, 122)
    gTeal = RGB(14, 92, 108)
    gYesil = RGB(20, 92, 68)
    gGri = RGB(78, 92, 112)
    gMor = RGB(82, 62, 112)

    ' MALZEMELER
    Set ws = KS_WS(SH_MALZ)
    KS_TabloAsagiAl T_MALZ
    KS_VeriSayfaGenel ws, 5
    Set lo = KS_TabloKur(ws, T_MALZ, ws.Range("B4"), _
        Array("Malzeme Kodu", "Malzeme Adý", "Kategori", "Birim", "Üretici", "Birim Fiyat", "Para Birimi", _
            "Garanti (Ay)", "Sevk Edilen", "Arýzalý", "Arýza Oraný", "Son Arýza", "Not", "Kayýt Tarihi"), _
        Array(15, 44, 16, 8, 40, 12, 8, 9, 10, 9, 9, 11, 30, 12), _
        Array("@", "@", "@", "@", "@", tl, "@", "0", "#,##0", "#,##0", "0.0%", tr, "@", tr), yeni)
    KS_BaslikRenk lo, Array(M_SEVKMIK, M_ARIZAMIK, M_ORAN, M_SONARIZA, M_KAYIT), 0
    KS_DuzenlenebilirSutun lo, KS_DuzenMalz()
    KS_SutunDogrulama lo, M_GARANTI, "sayi"
    KS_GrupBandi lo, Array(Array("MALZEME TANIMI", 1, 5, gLac), Array("FÝYAT VE GARANTÝ", 6, 8, gYesil), _
        Array("KALÝTE PERFORMANSI", 9, 12, gTeal), Array("DÝÐER", 13, 14, gGri))
    KS_VeriSayfaUst ws, "MALZEMELER", "Kategori, birim, garanti ve not doðrudan yazýlabilir; garanti boþ sevkiyatlara uygulanýr.", _
        Array("+ YENÝ MALZEME", "KS_GIT_KAYIT_MALZ", "ana", "GARANTÝSÝ GÝRÝLMEMÝÞ", "KS_F_MALZ_GARYOK", "turuncu", _
            "ARIZA VERENLER", "KS_F_MALZ_ARIZALI", "teal", "TÜMÜNÜ GÖSTER", "KS_F_TEMIZLE", "cizgi", _
            "ARA", "KS_ARA", "koyu", "", "", "")
    KS_GorunumDugmesi ws
    KS_OzetGorunum lo, KS_GizliSutunlar(T_MALZ), True
    KS_Pencere ws, 90, 4

    ' SEVKÝYATLAR
    Set ws = KS_WS(SH_SEVK)
    KS_TabloAsagiAl T_SEVK
    KS_VeriSayfaGenel ws, 5
    Set lo = KS_TabloKur(ws, T_SEVK, ws.Range("B4"), _
        Array("Sevk No", "Sevkçi", "Sevk Yeri", "Malzeme Kodu", "Malzeme Adý", "Adet", "Birim Fiyat", _
            "Para Birimi", "Tutar (TL)", "Planlanan Teslim", "Gerçek Geliþ", "Gecikme (Gün)", "Teslim Durumu", _
            "Garanti (Ay)", "Garanti Bitiþ", "Garanti Durumu", "Belge / SAS No", "Seri / Lot No", "Arýzalý", _
            "Not", "Kayýt Zamaný", "Kaydeden"), _
        Array(11.5, 36, 16, 12.5, 36, 7, 12, 8, 13, 11, 11, 9, 14, 8.5, 11, 18, 18, 14, 8, 30, 15, 11), _
        Array("@", "@", "@", "@", "@", "General", tl, "@", tl, tr, tr, "0", "@", "0", tr, "@", "@", "@", _
            "General", "@", "dd.mm.yyyy hh:mm", "@"), yeni)
    KS_BaslikRenk lo, Array(S_NO, S_AD, S_TUTAR, S_GECIKME, S_TDURUM, S_GARBIT, S_GDURUM, S_ARIZA, S_ZAMAN, S_KULL), 0
    KS_DuzenlenebilirSutun lo, KS_DuzenSevk()
    KS_SutunDogrulama lo, S_GELIS, "tarih"
    KS_SutunDogrulama lo, S_GARAY, "sayi"
    KS_DurumRenkleri lo, S_TDURUM, Array("ZAMANINDA", "GECÝKMELÝ", "GECÝKÝYOR")
    KS_DurumRenkleri lo, S_GDURUM, Array("DEVAM", "BÝTTÝ", "BÝTÝYOR")
    KS_GrupBandi lo, Array(Array("SEVKÝYAT", 1, 5, gLac), Array("MÝKTAR / TUTAR", 6, 9, gCelik), _
        Array("TESLÝMAT", 10, 13, gTeal), Array("GARANTÝ", 14, 16, gYesil), Array("BELGE / DÝÐER", 17, 22, gGri))
    KS_VeriSayfaUst ws, "SEVKÝYATLAR", "Gerçek geliþ, garanti, belge, seri ve not doðrudan yazýlabilir.", _
        Array("YILLIK EXCEL YÜKLE", "KS_SAP_AKTAR", "yesil", "GARANTÝSÝ BÝTECEK", "KS_F_SEVK_GARBITIYOR", "turuncu", _
            "GECÝKENLER", "KS_F_SEVK_GECIKEN", "turuncu", "GARANTÝSÝ GÝRÝLMEMÝÞ", "KS_F_SEVK_GARYOK", "teal", _
            "TÜMÜNÜ GÖSTER", "KS_F_TEMIZLE", "cizgi", "ARA", "KS_ARA", "koyu", "", "", "")
    KS_GorunumDugmesi ws
    KS_OzetGorunum lo, KS_GizliSutunlar(T_SEVK), True
    KS_Pencere ws, 90, 4

    ' ARIZALAR
    Set ws = KS_WS(SH_ARIZA)
    KS_TabloAsagiAl T_ARIZA
    KS_VeriSayfaGenel ws, 5
    Set lo = KS_TabloKur(ws, T_ARIZA, ws.Range("B4"), _
        Array("Arýza No", "Arýza Tarihi", "Malzeme Kodu", "Malzeme Adý", "Seri No", "Üretici", "Sevkçi", _
            "Sevk Yeri", "Baðlý Sevkiyat", "Garanti Bitiþ", "Garanti Devam Ediyor mu?", "Garanti Bilgisi", _
            "Arýzalý Adet", "Arýza Tipi", "Arýza Sebebi / Açýklama", "Üreticiye Bildirim", "Mail Yazýþmasý", _
            "Mail Dosyasý", "Üretici Dönüþ", "Dönüþ Süresi (Gün)", "Arýza Devam Ediyor mu?", "Çözüldü mü?", _
            "Durum", "Çözüm Tarihi", "Çözüm Süresi (Gün)", "Çözüm Þekli", "Tamir Maliyeti (TL)", _
            "Maliyeti Karþýlayan", "Önceki Arýza", "Tekrar", "Not", "Kayýt Zamaný", "Kaydeden"), _
        Array(11, 10, 12.5, 30, 12, 32, 34, 15, 12, 11, 10, 11, 8, 15, 38, 11, 40, 30, 11, 9, 11, 10, 15, 11, _
            9, 12, 12, 13, 11, 7, 26, 15, 11), _
        Array("@", tr, "@", "@", "@", "@", "@", "@", "@", tr, "@", "@", "General", "@", "@", tr, "@", "@", tr, _
            "0", "@", "@", "@", tr, "0", "@", tl, "@", tr, "0", "@", "dd.mm.yyyy hh:mm", "@"), yeni)
    KS_BaslikRenk lo, Array(A_NO, A_AD, A_URETICI, A_SEVKCI, A_YER, A_SEVK, A_GARBIT, A_KAPKAYNAK, A_DONUSGUN, _
        A_DURUM, A_COZUMGUN, A_ONCEKI, A_TEKRAR, A_ZAMAN, A_KULL), 0
    KS_DurumRenkleri lo, A_KAPSAM, Array("EVET", "HAYIR", "BELÝRSÝZ")
    KS_DurumRenkleri lo, A_DURUM, Array("ÇÖZÜLDÜ", "ÇÖZÜLEMEDÝ", "DEVAM")
    KS_DurumRenkleri lo, A_DURUM, Array("ÇÖZÜLDÜ", "ÇÖZÜLEMEDÝ", "ÜRETÝCÝDE"), True
    KS_GrupBandi lo, Array(Array("ARIZA", 1, 5, gLac), Array("KAYNAK", 6, 9, gCelik), Array("GARANTÝ", 10, 12, gYesil), _
        Array("ARIZA DETAYI", 13, 15, gTeal), Array("ÜRETÝCÝ SÜRECÝ", 16, 20, gMor), Array("SONUÇ", 21, 28, gCelik), _
        Array("GEÇMÝÞ / DÝÐER", 29, 33, gGri))
    KS_VeriSayfaUst ws, "ARIZALAR", "Düzeltme için KAYIT > ARIZA sekmesinde arýza numarasýný yazýn.", _
        Array("+ YENÝ ARIZA", "KS_GIT_KAYIT_ARIZA", "ana", "DEVAM EDENLER", "KS_F_ARIZA_ACIK", "turuncu", _
            "ÜRETÝCÝDE", "KS_F_ARIZA_URETICIDE", "turuncu", "ÇÖZÜLENLER", "KS_F_ARIZA_COZULEN", "yesil", _
            "TÜMÜNÜ GÖSTER", "KS_F_TEMIZLE", "cizgi", "ARA", "KS_ARA", "koyu", "", "", "")
    KS_GorunumDugmesi ws
    KS_OzetGorunum lo, KS_GizliSutunlar(T_ARIZA), True
    KS_Pencere ws, 90, 4

    ' GÜNLÜK
    Set ws = KS_WS(SH_LOG)
    KS_VeriSayfaGenel ws, 4
    Set lo = KS_TabloKur(ws, T_LOG, ws.Range("B3"), _
        Array("Zaman", "Kullanýcý", "Ýþlem", "Kayýt No", "Detay"), _
        Array(18, 16, 22, 16, 100), _
        Array("dd.mm.yyyy hh:mm:ss", "@", "@", "@", "@"), yeni)
    KS_VeriSayfaUst ws, "ÝÞLEM GÜNLÜÐÜ", "Her kayýt, güncelleme, aktarým ve yedek burada iz býrakýr. Silinmez.", _
        Array("TANIMLAR", "KS_GIT_TANIM", "cizgi", "TOPLU AKTAR", "KS_GIT_AKTAR", "cizgi")
    KS_Pencere ws, 90, 3
End Sub

' Durum sütununa koþullu renk: iyi = yeþil, kötü = kýrmýzý, uyarý = turuncu (metnin baþýna göre).
Private Sub KS_DurumRenkleri(ByVal lo As ListObject, ByVal kolon As Long, ByVal deger As Variant, _
        Optional ByVal ekle As Boolean = False)
    Dim r As Range, i As Long, renk As Variant, fc As Object
    Set r = lo.ListColumns(kolon).Range
    Set r = r.Offset(1).Resize(100000)
    If Not ekle Then r.FormatConditions.Delete
    renk = Array(C_YESIL(), C_KIRMIZI(), C_TURUNCU())
    For i = 0 To 2
        If Len(deger(i)) > 0 And Not (ekle And i < 2) Then
            Set fc = r.FormatConditions.Add(Type:=xlTextString, String:=CStr(deger(i)), TextOperator:=xlBeginsWith)
            fc.Font.Color = renk(i)
            fc.Font.Bold = True
        End If
    Next i
End Sub

' ---------------------------------------------------------------- TANIMLAR (YÖNETÝM)

Private Sub KS_TanimlarKur()
    Dim ws As Worksheet, lo As ListObject, yeni As Boolean, i As Long, v As Variant, alt As Double, sol As Double
    Set ws = KS_WS(SH_TANIM)
    ws.Cells.Font.Name = KS_FONT
    ws.Cells.Font.Size = 9.5
    ws.Columns(1).ColumnWidth = 2

    ' Sevkçi listesinde SAP biçimi (kod + ad) anahtardýr; ayný ad farklý kodla gelebilir.
    Set lo = KS_TabloKur(ws, T_SEVKCI, ws.Range("B3"), _
        Array("Sevkçi (SAP)", "Tedarikçi Kodu", "Tedarikçi Adý", "Tür", "Sevk Yerleri", "Not"), _
        Array(44, 13, 34, 15, 30, 40), Array("@", "@", "@", "@", "@", "@"), yeni)
    If yeni Then KS_TedarikciEkle
    ws.Columns("H").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_YER, ws.Range("I3"), Array("Sevk Yeri", "Üretim Yeri", "Ýl"), _
        Array(22, 11, 12), Array("@", "@", "@"), yeni)
    If yeni Then
        KS_Tohum lo, Array("ÇEDAÞ TRABZON", "ÇEDAÞ RÝZE", "ÇEDAÞ ARTVÝN", "ÇEDAÞ GÝRESUN", "ÇEDAÞ GÜMÜÞHANE"), _
            Array("7152", "7151", "7155", "7154", "7153")
        lo.ListColumns(3).DataBodyRange.Value = Application.Transpose( _
            Array("Trabzon", "Rize", "Artvin", "Giresun", "Gümüþhane"))
    End If
    ws.Columns("L").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_KAT, ws.Range("M3"), Array("Kategori"), Array(20), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Kablo", "Trafo", "Sayaç", "Kesici / Ayýrýcý", "Direk / Travers", _
        "Ýzolatör", "Parafudr", "Pano / Hücre", "Redresör / Akü", "Aydýnlatma", "ÝSG Malzemesi", "Hizmet", "Diðer")
    ws.Columns("N").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_BIRIM, ws.Range("O3"), Array("Birim"), Array(12), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("ADT", "M", "KG", "SET", "TAKIM", "AY", "GÜN")
    ws.Columns("P").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_TIP, ws.Range("Q3"), Array("Arýza Tipi"), Array(22), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Elektriksel", "Mekanik", "Ýzolasyon", "Üretim Hatasý", _
        "Montaj Kaynaklý", "Nakliye Hasarý", "Yazýlým / Haberleþme", "Diðer")
    ws.Columns("R").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_COZUM, ws.Range("S3"), Array("Çözüm Þekli"), Array(18), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Tamir", "Deðiþim", "Ýade", "Hurda", "Beklemede")
    ws.Columns("T").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_KARS, ws.Range("U3"), Array("Maliyeti Karþýlayan"), Array(20), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Üretici", "Kurum", "Yüklenici", "Paylaþýmlý")
    ws.Columns("V").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_KUR, ws.Range("W3"), Array("Para Birimi", "Kur (TL)"), Array(12, 12), _
        Array("@", "#,##0.0000"), yeni)
    If yeni Then
        KS_Tohum lo, Array("TRY", "USD", "EUR", "GBP"), Array(1, Empty, Empty, Empty)
    End If
    ws.Columns("Y").ColumnWidth = 2

    ' Parametreler
    ws.Columns("Z").ColumnWidth = 46
    ws.Columns("AA").ColumnWidth = 15
    ws.Columns("AB").ColumnWidth = 2
    KS_TabloBaslik ws.Range("Z3:AA3"), Array("PARAMETRE", "DEÐER")
    v = Array( _
        Array("Garanti bitiþ uyarýsý (gün)", "pGarUyari", 60, "0"), _
        Array("Kalite skoru: arýza oraný katsayýsý (puan / %1)", "pKatAriza", 5, "0.0"), _
        Array("Kalite skoru: gecikme katsayýsý (puan / gün)", "pKatGecikme", 1, "0.0"), _
        Array("Kalite skoru: üretici dönüþ katsayýsý (puan / gün)", "pKatDonus", 0.5, "0.0"), _
        Array("Yedek hatýrlatma aralýðý (gün)", "pYedekGun", 7, "0"), _
        Array("Son yedek", "pSonYedek", Empty, "dd.mm.yyyy hh:mm"))
    For i = 0 To UBound(v)
        ws.Cells(4 + i, "Z").Value = v(i)(0)
        KS_AdHucre CStr(v(i)(1)), ws.Cells(4 + i, "AA")
        If KS_Bos(ws.Cells(4 + i, "AA").Value) And Not IsEmpty(v(i)(2)) Then ws.Cells(4 + i, "AA").Value = v(i)(2)
        ws.Cells(4 + i, "AA").NumberFormat = v(i)(3)
        ws.Cells(4 + i, "AA").HorizontalAlignment = xlCenter
    Next i
    KS_TabloGovde ws.Range("Z4:AA" & 4 + UBound(v))

    ' Açýlýr liste kaynaklarý
    ws.Columns("AC").ColumnWidth = 10
    ws.Columns("AD").ColumnWidth = 12
    ws.Columns("AE").ColumnWidth = 12
    KS_TabloBaslik ws.Range("AC3:AE3"), Array("YIL", "EVET / HAYIR", "ÇÖZÜLDÜ MÜ")
    ws.Range("AD4:AD5").Value = Application.Transpose(Array("EVET", "HAYIR"))
    ws.Range("AE4:AE6").Value = Application.Transpose(Array("EVET", "HAYIR", "KISMEN"))
    KS_TabloGovde ws.Range("AD4:AE6")
    KS_YilListesiGuncelle

    ' Üst bölüm ve yönetim araçlarý
    KS_VeriSayfaUst ws, "YÖNETÝM  /  TANIMLAR", "Açýlýr listelerin kaynaðý, döviz kurlarý ve parametreler. Tablolara satýr ekleyebilirsiniz.", _
        Array("TOPLU AKTAR", "KS_GIT_AKTAR", "koyu", "ÝÞLEM GÜNLÜÐÜ", "KS_GIT_LOG", "koyu", _
            "YEDEK AL", "KS_YEDEK_AL", "ana")
    KS_SekilSil ws, "ksArac_"
    KS_Yazi ws, "ksArac_B", "YÖNETÝM ARAÇLARI", ws.Range("Z12").Left, ws.Range("Z12").Top, 300, 18, 10.5, True, _
        C_LAC(), msoAlignLeft
    v = Array("TEDARÝKÇÝ LÝSTESÝNÝ YÜKLE", "KS_TEDARIKCI_YUKLE", "koyu", _
        "TÜMÜNÜ YENÝDEN HESAPLA", "KS_YENIDEN_HESAPLA", "koyu", _
        "SAYFA KORUMASI AÇ / KAPAT", "KS_KORUMA_DEGISTIR", "koyu", _
        "TASARIMI YENÝLE (ONARIM)", "KALITESORA_KUR", "teal", _
        "UYGULAMA / EXCEL GÖRÜNÜMÜ", "KS_UYGULAMA_MODU_DEGISTIR", "teal", _
        "ÖRNEK VERÝ YÜKLE (deneme)", "KS_ORNEK_VERI_YUKLE", "cizgi", _
        "ÖRNEK VERÝYÝ SÝL", "KS_ORNEK_VERI_SIL", "cizgi", _
        "VBA KODUNU DIÞA AKTAR", "KS_VBA_DISARI_AKTAR", "gri")
    For i = 0 To UBound(v) \ 3
        KS_Buton ws, "ksArac_" & i, CStr(v(i * 3)), ws.Range("Z13").Left, ws.Range("Z12").Top + 24 + i * 30, _
            ws.Range("Z13").Width + ws.Range("AA13").Width, 24, CStr(v(i * 3 + 1)), CStr(v(i * 3 + 2))
    Next i
    KS_Pencere ws, 90, 3
End Sub

Private Sub KS_Tohum(ByVal lo As ListObject, ByVal ilk As Variant, Optional ByVal ikinci As Variant)
    Dim v() As Variant, i As Long, n As Long
    n = UBound(ilk) - LBound(ilk) + 1
    ReDim v(1 To n, 1 To lo.ListColumns.Count)
    For i = 1 To n
        v(i, 1) = ilk(LBound(ilk) + i - 1)
        If Not IsMissing(ikinci) Then v(i, 2) = ikinci(LBound(ikinci) + i - 1)
    Next i
    KS_TopluEkle lo, v, n
End Sub

' Yýl listesi: TÜMÜ + verideki ilk yýldan bu yýla (en az son 5 yýl).
Public Sub KS_YilListesiGuncelle()
    Dim ws As Worksheet, ilkYil As Long, y As Long, i As Long, a As Variant, d As Date, lo As ListObject
    Dim korumali As Boolean
    Set ws = KS_WS(SH_TANIM)
    If ws Is Nothing Then Exit Sub
    korumali = ws.ProtectContents
    KS_Ac ws
    ilkYil = Year(Date) - 4
    Set lo = KS_LO(T_SEVK)
    If Not lo Is Nothing Then
        a = KS_Veri(lo)
        If IsArray(a) Then
            For i = 1 To UBound(a, 1)
                If KS_Tarih(a(i, S_PLAN), d) Then If Year(d) < ilkYil Then ilkYil = Year(d)
                If KS_Tarih(a(i, S_GELIS), d) Then If Year(d) < ilkYil Then ilkYil = Year(d)
            Next i
        End If
    End If
    If ilkYil < Year(Date) - 25 Then ilkYil = Year(Date) - 25
    ws.Range("AC4:AC40").ClearContents
    ws.Range("AC4").Value = "TÜMÜ"
    i = 5
    For y = Year(Date) + 1 To ilkYil Step -1
        ws.Cells(i, "AC").Value = y
        i = i + 1
    Next y
    KS_TabloGovde ws.Range("AC4:AC" & i - 1)
    KS_AdTanimla "lstYil", "='" & ws.Name & "'!$AC$4:$AC$" & i - 1
    KS_AdTanimla "lstYilSayi", "='" & ws.Name & "'!$AC$5:$AC$" & i - 1
    If korumali Then KS_Kapat ws
End Sub

' ---------------------------------------------------------------- adlar

Private Sub KS_AdlariKur()
    KS_AdTanimla "lstMalzeme", "=INDEX(" & T_MALZ & ",0,1)"
    KS_AdTanimla "lstMalzemeAd", "=INDEX(" & T_MALZ & ",0,2)"
    KS_AdTanimla "lstSevk", "=INDEX(" & T_SEVK & ",0,1)"
    KS_AdTanimla "lstAriza", "=INDEX(" & T_ARIZA & ",0,1)"
    KS_AdTanimla "lstSevkci", "=INDEX(" & T_SEVKCI & ",0,1)"
    KS_AdTanimla "lstYer", "=INDEX(" & T_YER & ",0,1)"
    KS_AdTanimla "lstKategori", "=INDEX(" & T_KAT & ",0,1)"
    KS_AdTanimla "lstBirim", "=INDEX(" & T_BIRIM & ",0,1)"
    KS_AdTanimla "lstTip", "=INDEX(" & T_TIP & ",0,1)"
    KS_AdTanimla "lstCozum", "=INDEX(" & T_COZUM & ",0,1)"
    KS_AdTanimla "lstKarsilayan", "=INDEX(" & T_KARS & ",0,1)"
    KS_AdTanimla "lstPB", "=INDEX(" & T_KUR & ",0,1)"
    KS_AdTanimla "lstEvetHayir", "='" & SH_TANIM & "'!$AD$4:$AD$5"
    KS_AdTanimla "lstCozuldu", "='" & SH_TANIM & "'!$AE$4:$AE$6"
End Sub

' Açýlýr liste ekler. sert = False ise listede olmayan deðere uyarýyla izin verilir.
Public Sub KS_Liste(ByVal r As Range, ByVal ad As String, Optional ByVal sert As Boolean = True)
    r.Validation.Delete
    r.Validation.Add Type:=xlValidateList, AlertStyle:=IIf(sert, xlValidAlertStop, xlValidAlertInformation), _
        Formula1:="=" & ad
    r.Validation.IgnoreBlank = True
    r.Validation.InCellDropdown = True
    r.Validation.ShowError = True
    r.Validation.ErrorTitle = KS_AD
    r.Validation.ErrorMessage = "Listede olmayan bir deðer girdiniz. Yeni deðerler TANIMLAR sayfasýndan eklenir."
End Sub

Public Sub KS_TarihDogrulama(ByVal r As Range)
    r.Validation.Delete
    r.Validation.Add Type:=xlValidateDate, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, _
        Formula1:=CStr(CLng(DateSerial(1990, 1, 1))), Formula2:=CStr(CLng(DateSerial(2100, 12, 31)))
    r.Validation.ErrorTitle = KS_AD
    r.Validation.ErrorMessage = "Geçerli bir tarih girin (gg.aa.yyyy)."
End Sub

Public Sub KS_SayiDogrulama(ByVal r As Range)
    r.Validation.Delete
    r.Validation.Add Type:=xlValidateDecimal, AlertStyle:=xlValidAlertStop, Operator:=xlGreaterEqual, _
        Formula1:="0"
    r.Validation.ErrorTitle = KS_AD
    r.Validation.ErrorMessage = "Sýfýr veya pozitif bir sayý girin."
End Sub
