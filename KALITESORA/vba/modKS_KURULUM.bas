Attribute VB_Name = "modKS_KURULUM"
Option Explicit

' KALİTESORA kurulumu. Boş bir .xlsm dosyasında tüm sayfaları, tabloları, formu ve paneli oluşturur.
' Tekrar çalıştırıldığında veri tablolarına dokunmaz; yalnız eksikleri tamamlar ve tasarımı yeniler.

Private Const ETIKET As String = "_KS_KURULU"

Public Function KS_Kurulu() As Boolean
    KS_Kurulu = KS_AdVar(ETIKET)
End Function

Public Sub KALITESORA_KUR()
    Dim adim As String, yedek As String, ilk As Boolean, ad As Variant, ws As Worksheet
    Dim onceki As Long
    On Error GoTo Hata

    If Len(ThisWorkbook.Path) = 0 Or LCase$(Right$(ThisWorkbook.Name, 5)) <> ".xlsm" Then
        MsgBox "Önce dosyayı 'Makro İçerebilen Excel Çalışma Kitabı (.xlsm)' olarak kaydedin, " & _
            "sonra kurulumu çalıştırın.", vbExclamation, KS_AD
        Exit Sub
    End If
    If ThisWorkbook.ReadOnly Then
        MsgBox "Dosya salt okunur açık. Yazılabilir kopyada çalıştırın.", vbExclamation, KS_AD
        Exit Sub
    End If

    ilk = Not KS_Kurulu()
    If ilk Then
        If MsgBox(KS_AD & " bu dosyaya kurulacak." & vbCrLf & vbCrLf & _
            "11 sayfa, veri tabloları, kayıt formu, panel ve raporlar oluşturulacak." & vbCrLf & _
            "Devam edilsin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    Else
        If MsgBox(KS_AD & " zaten kurulu." & vbCrLf & vbCrLf & _
            "Onarım modunda: veri tablolarındaki kayıtlar korunur, eksik parçalar tamamlanır, " & _
            "form, panel ve rapor sayfalarının tasarımı yeniden çizilir." & vbCrLf & _
            "Önce dosyanın yedeği alınacak. Devam edilsin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
        adim = "Kurulum öncesi yedek"
        yedek = KS_YedekKopya("ONARIM_ONCESI")
    End If

    KS_Basla
    adim = "Sayfaları hazırlama"
    KS_Durum adim
    For Each ad In KS_Sayfalar()
        Set ws = KS_WS(CStr(ad))
        If Not ws Is Nothing Then KS_Ac ws
    Next ad
    KS_SayfalariHazirla ilk

    adim = "Tablo stili"
    KS_TabloStiliKur

    adim = "TANIMLAR sayfası"
    KS_Durum adim
    KS_TanimlarKur

    adim = "Veri tabloları"
    KS_Durum adim
    KS_VeriSayfalariKur

    adim = "Adlandırılmış listeler"
    KS_AdlariKur

    adim = "5 YILLIK RAPOR sayfası"
    KS_Durum adim
    KS_RaporKur

    adim = "KAYIT formu"
    KS_Durum adim
    KS_KayitFormKur

    adim = "MALZEME KARNESİ sayfası"
    KS_KarneKur

    adim = "PERFORMANS sayfası"
    KS_PerfKur

    adim = "TOPLU AKTAR sayfası"
    KS_AktarKur

    adim = "PANEL sayfası"
    KS_Durum adim
    KS_PanelKur

    adim = "Gezinme menüsü"
    For Each ad In KS_Sayfalar()
        KS_NavKur KS_WS(CStr(ad))
    Next ad

    adim = "Kurulum kaydı"
    KS_AdTanimla ETIKET, "=1"
    ThisWorkbook.Names(ETIKET).Visible = False
    If Not KS_AdVar("_KS_KORUMA") Then
        KS_AdTanimla "_KS_KORUMA", "=1"
        ThisWorkbook.Names("_KS_KORUMA").Visible = False
    End If

    adim = "Hesaplama ve panel"
    KS_Durum adim
    KS_HesaplaTumu
    KS_PanelYenile False

    adim = "Sayfa koruması"
    For Each ad In KS_Sayfalar()
        KS_Kapat KS_WS(CStr(ad))
    Next ad

    KS_Bitir True
    KS_Log IIf(ilk, "KURULUM", "ONARIM"), "v" & KS_SURUM, IIf(Len(yedek) > 0, "Yedek: " & yedek, "")
    KS_WS(SH_PANEL).Activate
    MsgBox KS_AD & " " & IIf(ilk, "kuruldu.", "onarıldı.") & vbCrLf & vbCrLf & _
        "Başlangıç sırası:" & vbCrLf & _
        "1) TANIMLAR: sevkçi / üretici ve sevk yeri listelerini doldurun." & vbCrLf & _
        "2) KAYIT: önce malzemeyi, sonra sevkiyatı, sonra arızayı kaydedin." & vbCrLf & _
        "3) PANEL / RAPORLAR: YENİLE ile güncel tabloyu görün." & vbCrLf & vbCrLf & _
        "Sayfa koruma şifresi: " & KS_SIFRE & _
        IIf(Len(yedek) > 0, vbCrLf & "Onarım öncesi yedek: " & yedek, ""), vbInformation, KS_AD
    Exit Sub

Hata:
    onceki = Err.Number
    KS_Bitir True
    MsgBox "Kurulum tamamlanamadı." & vbCrLf & _
        "Adım: " & adim & vbCrLf & _
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
            Case SH_PANEL: onceki.Tab.Color = C_CAMGOBEGI()
            Case SH_KAYIT: onceki.Tab.Color = C_YESIL()
            Case SH_SEVK, SH_ARIZA, SH_MALZ: onceki.Tab.Color = C_LAC()
            Case SH_KARNE, SH_PERF, SH_RAPOR: onceki.Tab.Color = C_TURKUAZ()
            Case Else: onceki.Tab.Color = C_GRI()
        End Select
    Next ad
    ' İlk kurulumda yalnız tamamen boş varsayılan sayfalar silinir.
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

' Tabloyu oluşturur ya da mevcut tablonun başlıklarını doğrular. yeni = True ise tablo yeni kuruldu.
Private Function KS_TabloKur(ByVal ws As Worksheet, ByVal ad As String, ByVal sol As Range, _
        ByVal basliklar As Variant, ByVal genislik As Variant, ByVal bicim As Variant, _
        ByRef yeni As Boolean) As ListObject
    Dim lo As ListObject, n As Long, i As Long, rng As Range
    n = UBound(basliklar) - LBound(basliklar) + 1
    For i = 0 To n - 1
        ws.Columns(sol.Column + i).ColumnWidth = genislik(LBound(genislik) + i)
        ws.Columns(sol.Column + i).NumberFormat = bicim(LBound(bicim) + i)
    Next i
    Set lo = KS_LO(ad)
    yeni = lo Is Nothing
    If yeni Then
        Set rng = sol.Resize(1, n)
        rng.Value = basliklar
        Set lo = ws.ListObjects.Add(xlSrcRange, sol.Resize(2, n), , xlYes)
        lo.Name = ad
        For i = lo.ListRows.Count To 1 Step -1
            If Application.WorksheetFunction.CountA(lo.ListRows(i).Range) = 0 Then lo.ListRows(i).Delete
        Next i
    Else
        If lo.ListColumns.Count <> n Then
            Err.Raise vbObjectError + 7101, KS_AD, ad & " tablosunun sütun sayısı beklenenden farklı (" & _
                lo.ListColumns.Count & " / " & n & "). Kurulum veriye dokunmadan durduruldu."
        End If
        For i = 1 To n
            If StrComp(lo.HeaderRowRange.Cells(1, i).Value, basliklar(LBound(basliklar) + i - 1), vbTextCompare) <> 0 Then
                Err.Raise vbObjectError + 7102, KS_AD, ad & " tablosunda " & i & ". sütun başlığı farklı: '" & _
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
    lo.HeaderRowRange.EntireRow.RowHeight = 34
    Set KS_TabloKur = lo
End Function

' Sistem tarafından hesaplanan sütunların başlığı turkuaz.
Private Sub KS_HesapBaslik(ByVal lo As ListObject, ByVal kolonlar As Variant)
    Dim k As Variant
    For Each k In kolonlar
        With lo.HeaderRowRange.Cells(1, CLng(k))
            .Interior.Color = C_TURKUAZ()
            .Font.Color = vbWhite
        End With
    Next k
End Sub

Private Sub KS_VeriSayfaGenel(ByVal ws As Worksheet)
    ws.Cells.Font.Name = KS_FONT
    ws.Cells.Font.Size = 9
    ws.Columns(1).ColumnWidth = 2
    ws.Rows("7:" & ws.Rows.Count).RowHeight = 17
End Sub

Private Sub KS_VeriSayfalariKur()
    Dim ws As Worksheet, lo As ListObject, yeni As Boolean, tl As String, tr As String
    tl = KS_FmtTL()
    tr = KS_FmtTarih()

    Set ws = KS_WS(SH_MALZ)
    KS_VeriSayfaGenel ws
    Set lo = KS_TabloKur(ws, T_MALZ, ws.Range("B6"), _
        Array("Malzeme Kodu", "Malzeme Adı", "Kategori", "Birim", "Üretici", "Birim Fiyat (TL)", _
            "Garanti (Ay)", "Kayıt Tarihi", "Not", "Toplam Sevk Miktarı", "Toplam Arızalı Miktar", _
            "Arıza Oranı", "Son Arıza Tarihi"), _
        Array(16, 34, 16, 9, 22, 15, 10, 12, 28, 13, 13, 11, 13), _
        Array("@", "@", "@", "@", "@", tl, "0", tr, "@", "General", "General", "0.0%", tr), yeni)
    KS_HesapBaslik lo, Array(M_KAYIT, M_SEVKMIK, M_ARIZAMIK, M_ORAN, M_SONARIZA)
    KS_Bant ws, M_KOLON + 2, "MALZEME ANA KAYITLARI", _
        "Kayıt ve güncelleme KAYIT sayfasından yapılır. Turkuaz başlıklı sütunları sistem hesaplar."
    KS_Pencere ws, 90, 6, True

    Set ws = KS_WS(SH_SEVK)
    KS_VeriSayfaGenel ws
    Set lo = KS_TabloKur(ws, T_SEVK, ws.Range("B6"), _
        Array("Sevk No", "Sevkçi", "Sevk Yeri", "Malzeme Kodu", "Malzeme Adı", "Adet / Miktar", _
            "Birim Fiyat (TL)", "Toplam Tutar (TL)", "Planlanan Teslim", "Gerçek Geliş", "Gecikme (Gün)", _
            "Teslim Durumu", "Garanti (Ay)", "Garanti Bitiş", "Garanti Durumu", "İrsaliye / Belge No", _
            "Seri / Lot No", "Arızalı Miktar", "Not", "Kayıt Zamanı", "Kaydeden"), _
        Array(13, 22, 16, 15, 30, 10, 14, 15, 12, 12, 9, 14, 9, 12, 18, 16, 14, 10, 26, 16, 12), _
        Array("@", "@", "@", "@", "@", "General", tl, tl, tr, tr, "0", "@", "0", tr, "@", "@", "@", _
            "General", "@", "dd.mm.yyyy hh:mm", "@"), yeni)
    KS_HesapBaslik lo, Array(S_NO, S_AD, S_TUTAR, S_GECIKME, S_TDURUM, S_GARAY, S_GARBIT, S_GDURUM, _
        S_ARIZA, S_ZAMAN, S_KULL)
    KS_Bant ws, S_KOLON + 2, "SEVKİYAT KAYITLARI", _
        "Garanti, teslim gerçek geliş tarihinden başlar. Turkuaz başlıklı sütunları sistem hesaplar."
    KS_DurumRenkleri lo, S_TDURUM, Array("ZAMANINDA", "GECİKMELİ", "GECİKİYOR")
    KS_DurumRenkleri lo, S_GDURUM, Array("DEVAM", "BİTTİ", "BİTİYOR")
    KS_Pencere ws, 90, 6, True

    Set ws = KS_WS(SH_ARIZA)
    KS_VeriSayfaGenel ws
    Set lo = KS_TabloKur(ws, T_ARIZA, ws.Range("B6"), _
        Array("Arıza No", "Sevk No", "Seri No", "Malzeme Kodu", "Malzeme Adı", "Sevkçi", "Üretici", _
            "Sevk Yeri", "Arıza Tarihi", "Arızalı Adet", "Arıza Tipi", "Arıza Açıklaması", "Garanti Bitiş", _
            "Garanti Kapsamı", "Üreticiye Bildirim", "Üretici Dönüş", "Dönüş Süresi (Gün)", "Çözüm Tarihi", _
            "Çözüm Süresi (Gün)", "Çözüm Şekli", "Tamir Maliyeti (TL)", "Maliyeti Karşılayan", _
            "Önceki Arıza Tarihi", "Tekrar Sayısı", "Durum", "Not", "Kayıt Zamanı", "Kaydeden"), _
        Array(13, 13, 13, 15, 28, 20, 20, 16, 12, 9, 16, 32, 12, 11, 12, 12, 10, 12, 10, 13, 14, 14, _
            12, 8, 14, 24, 16, 12), _
        Array("@", "@", "@", "@", "@", "@", "@", "@", tr, "General", "@", "@", tr, "@", tr, tr, "0", _
            tr, "0", "@", tl, "@", tr, "0", "@", "@", "dd.mm.yyyy hh:mm", "@"), yeni)
    KS_HesapBaslik lo, Array(A_NO, A_KOD, A_AD, A_SEVKCI, A_URETICI, A_YER, A_GARBIT, A_KAPSAM, _
        A_DONUSGUN, A_COZUMGUN, A_ONCEKI, A_TEKRAR, A_DURUM, A_ZAMAN, A_KULL)
    KS_Bant ws, A_KOLON + 2, "ARIZA KAYITLARI", _
        "Her arıza bir sevkiyata bağlıdır. Önceki arıza: aynı seri no, seri yoksa aynı sevkiyat."
    KS_DurumRenkleri lo, A_KAPSAM, Array("EVET", "HAYIR", "")
    KS_DurumRenkleri lo, A_DURUM, Array("KAPANDI", "AÇIK", "ÜRETİCİDE")
    KS_Pencere ws, 90, 6, True

    Set ws = KS_WS(SH_LOG)
    KS_VeriSayfaGenel ws
    Set lo = KS_TabloKur(ws, T_LOG, ws.Range("B6"), _
        Array("Zaman", "Kullanıcı", "İşlem", "Kayıt No", "Detay"), _
        Array(18, 16, 20, 16, 90), _
        Array("dd.mm.yyyy hh:mm:ss", "@", "@", "@", "@"), yeni)
    KS_Bant ws, 8, "İŞLEM GÜNLÜĞÜ", "Her kayıt, güncelleme, aktarım ve yedek burada iz bırakır. Silinmez."
    KS_Pencere ws, 90, 6, True
End Sub

' Durum sütununa koşullu renk: iyi = yeşil, kötü = kırmızı, uyarı = turuncu (metnin başına göre).
' Metin koşulu kullanılır; formül içermediği için Türkçe / İngilizce Excel'de aynı çalışır.
Private Sub KS_DurumRenkleri(ByVal lo As ListObject, ByVal kolon As Long, ByVal deger As Variant)
    Dim r As Range, i As Long, renk As Variant, fc As Object
    Set r = lo.ListColumns(kolon).Range
    Set r = r.Offset(1).Resize(100000)
    r.FormatConditions.Delete
    renk = Array(C_YESIL(), C_KIRMIZI(), C_TURUNCU())
    For i = 0 To 2
        If Len(deger(i)) > 0 Then
            Set fc = r.FormatConditions.Add(Type:=xlTextString, String:=CStr(deger(i)), TextOperator:=xlBeginsWith)
            fc.Font.Color = renk(i)
            fc.Font.Bold = True
        End If
    Next i
End Sub

' ---------------------------------------------------------------- TANIMLAR

Private Sub KS_TanimlarKur()
    Dim ws As Worksheet, lo As ListObject, yeni As Boolean, i As Long, v As Variant
    Set ws = KS_WS(SH_TANIM)
    ws.Cells.Font.Name = KS_FONT
    ws.Cells.Font.Size = 9
    ws.Columns(1).ColumnWidth = 2

    ' Sevkçi listesinde SAP biçimi (kod + ad) anahtardır; aynı ad farklı kodla gelebilir.
    Set lo = KS_TabloKur(ws, T_SEVKCI, ws.Range("B6"), _
        Array("Sevkçi (SAP)", "Tedarikçi Kodu", "Tedarikçi Adı", "Tür", "Sevk Yerleri", "Not"), _
        Array(44, 13, 34, 15, 30, 40), Array("@", "@", "@", "@", "@", "@"), yeni)
    If yeni Then KS_TedarikciEkle
    ws.Columns("H").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_YER, ws.Range("I6"), Array("Sevk Yeri", "Üretim Yeri", "İl"), _
        Array(20, 11, 12), Array("@", "@", "@"), yeni)
    If yeni Then
        KS_Tohum lo, Array("ÇEDAŞ TRABZON", "ÇEDAŞ RİZE", "ÇEDAŞ ARTVİN", "ÇEDAŞ GİRESUN", "ÇEDAŞ GÜMÜŞHANE"), _
            Array("7152", "7151", "7155", "7154", "7153")
        lo.ListColumns(3).DataBodyRange.Value = Application.Transpose( _
            Array("Trabzon", "Rize", "Artvin", "Giresun", "Gümüşhane"))
    End If
    ws.Columns("L").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_KAT, ws.Range("M6"), Array("Kategori"), Array(20), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Kablo", "Trafo", "Sayaç", "Kesici / Ayırıcı", "Direk / Travers", _
        "İzolatör", "Parafudr", "Pano / Hücre", "Aydınlatma", "Diğer")
    ws.Columns("N").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_BIRIM, ws.Range("O6"), Array("Birim"), Array(12), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Adet", "Metre", "Kg", "Takım", "Set")
    ws.Columns("P").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_TIP, ws.Range("Q6"), Array("Arıza Tipi"), Array(22), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Elektriksel", "Mekanik", "İzolasyon", "Üretim Hatası", _
        "Montaj Kaynaklı", "Nakliye Hasarı", "Yazılım / Haberleşme", "Diğer")
    ws.Columns("R").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_COZUM, ws.Range("S6"), Array("Çözüm Şekli"), Array(18), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Tamir", "Değişim", "İade", "Hurda", "Beklemede")
    ws.Columns("T").ColumnWidth = 2
    Set lo = KS_TabloKur(ws, T_KARS, ws.Range("U6"), Array("Maliyeti Karşılayan"), Array(20), Array("@"), yeni)
    If yeni Then KS_Tohum lo, Array("Üretici", "Kurum", "Yüklenici", "Paylaşımlı")
    ws.Columns("V").ColumnWidth = 2

    ' Parametreler
    ws.Columns("W").ColumnWidth = 42
    ws.Columns("X").ColumnWidth = 14
    ws.Columns("Y").ColumnWidth = 2
    KS_TabloBaslik ws.Range("W6:X6"), Array("PARAMETRE", "DEĞER")
    v = Array( _
        Array("Garanti bitiş uyarısı (gün)", "pGarUyari", 60, "0"), _
        Array("Kalite skoru: arıza oranı katsayısı (puan / %1)", "pKatAriza", 5, "0.0"), _
        Array("Kalite skoru: gecikme katsayısı (puan / gün)", "pKatGecikme", 1, "0.0"), _
        Array("Kalite skoru: üretici dönüş katsayısı (puan / gün)", "pKatDonus", 0.5, "0.0"), _
        Array("Yedek hatırlatma aralığı (gün)", "pYedekGun", 7, "0"), _
        Array("Son yedek", "pSonYedek", Empty, "dd.mm.yyyy hh:mm"))
    For i = 0 To UBound(v)
        ws.Cells(7 + i, "W").Value = v(i)(0)
        KS_AdHucre CStr(v(i)(1)), ws.Cells(7 + i, "X")
        If KS_Bos(ws.Cells(7 + i, "X").Value) And Not IsEmpty(v(i)(2)) Then ws.Cells(7 + i, "X").Value = v(i)(2)
        ws.Cells(7 + i, "X").NumberFormat = v(i)(3)
        ws.Cells(7 + i, "X").Locked = (v(i)(1) = "pSonYedek")
        ws.Cells(7 + i, "X").Interior.Color = IIf(v(i)(1) = "pSonYedek", C_OTO(), C_GIRIS())
        ws.Cells(7 + i, "X").HorizontalAlignment = xlCenter
    Next i
    KS_TabloGovde ws.Range("W7:W" & 6 + UBound(v) + 1)

    ' Açılır liste kaynakları
    ws.Columns("Z").ColumnWidth = 10
    ws.Columns("AA").ColumnWidth = 12
    ws.Columns("AB").ColumnWidth = 12
    KS_TabloBaslik ws.Range("Z6:AB6"), Array("YIL", "GRUPLAMA", "EVET / HAYIR")
    ws.Range("AA7:AA8").Value = Application.Transpose(Array("SEVKÇİ", "ÜRETİCİ"))
    ws.Range("AB7:AB8").Value = Application.Transpose(Array("EVET", "HAYIR"))
    KS_TabloGovde ws.Range("Z7:AB8")
    KS_YilListesiGuncelle

    KS_Bant ws, 29, "TANIMLAR", "Açılır listelerin kaynağı. Tablolara satır ekleyerek listeleri genişletin."
    KS_Pencere ws, 90, 0, False
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

' Yıl listesi: TÜMÜ + verideki ilk yıldan bu yıla (en az son 5 yıl).
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
    ws.Range("Z7:Z40").ClearContents
    ws.Range("Z7").Value = "TÜMÜ"
    i = 8
    For y = Year(Date) + 1 To ilkYil Step -1
        ws.Cells(i, "Z").Value = y
        i = i + 1
    Next y
    KS_TabloGovde ws.Range("Z7:Z" & i - 1)
    KS_AdTanimla "lstYil", "='" & ws.Name & "'!$Z$7:$Z$" & i - 1
    KS_AdTanimla "lstYilSayi", "='" & ws.Name & "'!$Z$8:$Z$" & i - 1
    If korumali Then KS_Kapat ws
End Sub

' ---------------------------------------------------------------- adlar

Private Sub KS_AdlariKur()
    KS_AdTanimla "lstMalzeme", "=INDEX(" & T_MALZ & ",0,1)"
    KS_AdTanimla "lstSevk", "=INDEX(" & T_SEVK & ",0,1)"
    KS_AdTanimla "lstAriza", "=INDEX(" & T_ARIZA & ",0,1)"
    KS_AdTanimla "lstSevkci", "=INDEX(" & T_SEVKCI & ",0,1)"
    KS_AdTanimla "lstYer", "=INDEX(" & T_YER & ",0,1)"
    KS_AdTanimla "lstKategori", "=INDEX(" & T_KAT & ",0,1)"
    KS_AdTanimla "lstBirim", "=INDEX(" & T_BIRIM & ",0,1)"
    KS_AdTanimla "lstTip", "=INDEX(" & T_TIP & ",0,1)"
    KS_AdTanimla "lstCozum", "=INDEX(" & T_COZUM & ",0,1)"
    KS_AdTanimla "lstKarsilayan", "=INDEX(" & T_KARS & ",0,1)"
    KS_AdTanimla "lstGrup", "='" & SH_TANIM & "'!$AA$7:$AA$8"
    KS_AdTanimla "lstEvetHayir", "='" & SH_TANIM & "'!$AB$7:$AB$8"
End Sub

' Açılır liste ekler. sert = False ise listede olmayan değere uyarıyla izin verilir.
Public Sub KS_Liste(ByVal r As Range, ByVal ad As String, Optional ByVal sert As Boolean = True)
    r.Validation.Delete
    r.Validation.Add Type:=xlValidateList, AlertStyle:=IIf(sert, xlValidAlertStop, xlValidAlertInformation), _
        Formula1:="=" & ad
    r.Validation.IgnoreBlank = True
    r.Validation.InCellDropdown = True
    r.Validation.ErrorTitle = KS_AD
    r.Validation.ErrorMessage = "Listeden seçin. Yeni değer için TANIMLAR sayfasına ekleyin."
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
    r.Validation.ErrorMessage = "Sıfır veya pozitif bir sayı girin."
End Sub
