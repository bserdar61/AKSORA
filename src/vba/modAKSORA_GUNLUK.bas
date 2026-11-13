Attribute VB_Name = "modAKSORA_GUNLUK"
Option Explicit

'==========================================================
' AKSORA v3.1  |  modAKSORA_GUNLUK
'
' 1) ISLEM GUNLUGU (denetim izi)
'    Kim, ne zaman, hangi islemi yapti, ne kadar surdu, hata oldu mu?
'    Kayitlar cok gizli "_AKSORA_GUNLUK" sayfasinda tutulur (en fazla 5.000 satir).
'    Yonetim Merkezi > ISLEM GUNLUGU dugmesi kayitlari Excel'e aktarir.
'
' 2) SISTEM SAGLIK KONTROLU
'    Dosya, sayfa, bilesen, SAP ve guvenlik kontrollerini yapar; sonucu Excel'e aktarir.
'    Hicbir veriyi degistirmez (yalniz okur).
'
' 3) AKSORA_GELISTIRME_SINAMA
'    Kurulum betigi kaydetmeden once cagirir: yeni modullerin ve eklenen
'    baglanti satirlarinin derlendigini dogrular.
'
' Kural: gunluk yazimi asla hata firlatmaz, asla mesaj gostermez,
' etkin sayfayi ve olay durumunu degistirmez.
'==========================================================

Public Const AKSORA_GELISTIRME_SURUM As String = "3.1"

Private Const GNL_SAYFA As String = "_AKSORA_GUNLUK"
Private Const GNL_SUTUN As Long = 7
Private Const GNL_AZAMI As Long = 5000
Private Const GNL_AKTAR_AZAMI As Long = 3000

'----------------------------------------------------------
' Cok gizli veri sayfasi: varsa dondurur, yoksa olusturur.
' Olusturma sirasinda olaylar kapatilir (yeni sayfa birim menusune yonlendirmesin)
' ve etkin sayfa eski haline getirilir. Olusturulamazsa Nothing doner.
'----------------------------------------------------------
Public Function AKSORA_GIZLI_SAYFA(ByVal ad As String, ByVal basliklar As Variant) As Worksheet
    Dim ws As Worksheet, onceki As Object, ev As Boolean, su As Boolean, i As Long
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    If Not ws Is Nothing Then
        Set AKSORA_GIZLI_SAYFA = ws
        Exit Function
    End If
    Err.Clear
    ev = Application.EnableEvents: su = Application.ScreenUpdating
    Application.EnableEvents = False: Application.ScreenUpdating = False
    If ActiveWorkbook Is ThisWorkbook Then Set onceki = ActiveSheet
    Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    If Not ws Is Nothing Then
        ws.Name = ad
        If IsArray(basliklar) Then
            For i = LBound(basliklar) To UBound(basliklar)
                ws.Cells(1, i - LBound(basliklar) + 1).Value2 = basliklar(i)
            Next i
        End If
        ws.Rows(1).Font.Bold = True
        ws.Visible = xlSheetVeryHidden
    End If
    If Not onceki Is Nothing Then onceki.Activate
    Application.EnableEvents = ev: Application.ScreenUpdating = su
    Set AKSORA_GIZLI_SAYFA = ws
End Function

Public Function AKSORA_KULLANICI() As String
    On Error Resume Next
    AKSORA_KULLANICI = Environ$("USERNAME")
    If Len(AKSORA_KULLANICI) = 0 Then AKSORA_KULLANICI = Application.UserName
End Function

'----------------------------------------------------------
' Gunluge bir satir ekler.
' tur   : GÜNCELLEME / HATA / YÖNETİCİ / SİSTEM
' sure  : saniye (bilinmiyorsa verilmez)
'----------------------------------------------------------
Public Sub AKSORA_GUNLUK_YAZ(ByVal tur As String, ByVal islem As String, Optional ByVal sure As Double = -1)
    Dim ws As Worksheet, r As Long, v(1 To 1, 1 To 7) As Variant
    On Error Resume Next
    Set ws = AKSORA_GIZLI_SAYFA(GNL_SAYFA, Array("Tarih", "Kullanıcı", "Bilgisayar", "Tür", "İşlem", "Süre (sn)", "Dosya"))
    If ws Is Nothing Then Exit Sub
    r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    If r > GNL_AZAMI + 1 Then
        ws.Rows("2:501").Delete
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
    End If
    v(1, 1) = CDbl(Now)
    v(1, 2) = AKSORA_KULLANICI()
    v(1, 3) = Environ$("COMPUTERNAME")
    v(1, 4) = tur
    v(1, 5) = Left$(Replace(Replace(islem, vbCr, " "), vbLf, " "), 500)
    If sure >= 0 Then v(1, 6) = Round(sure, 2)
    v(1, 7) = ThisWorkbook.Name
    ws.Range(ws.Cells(r, 1), ws.Cells(r, GNL_SUTUN)).Value2 = v
    If r = 2 Then
        ws.Columns(1).NumberFormat = "dd.mm.yyyy hh:mm:ss"
        ws.Columns(5).NumberFormat = "@"
    End If
End Sub

' Son "gun" gun icindeki belirli turdeki kayit sayisi (saglik kontrolu ve yonetim ozeti icin).
Public Function AKSORA_GUNLUK_SAY(ByVal tur As String, ByVal gun As Long) As Long
    Dim ws As Worksheet, son As Long, v As Variant, i As Long, sinir As Double
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GNL_SAYFA)
    On Error GoTo Cikis
    If ws Is Nothing Then Exit Function
    son = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If son < 2 Then Exit Function
    v = ws.Range(ws.Cells(2, 1), ws.Cells(son, 4)).Value2
    sinir = CDbl(Date) - gun
    For i = UBound(v, 1) To 1 Step -1
        If IsNumeric(v(i, 1)) And Not IsEmpty(v(i, 1)) Then
            If CDbl(v(i, 1)) < sinir Then Exit For
            If CStr(v(i, 4)) = tur Then AKSORA_GUNLUK_SAY = AKSORA_GUNLUK_SAY + 1
        End If
    Next i
Cikis:
End Function

'----------------------------------------------------------
' ISLEM GUNLUGU: kayitlar (en yeni ustte) yeni bir Excel kitabina aktarilir, masaustune kaydedilir.
'----------------------------------------------------------
Public Sub AKSORA_GUNLUK_EXCEL()
    Dim ws As Worksheet, son As Long, kaynak As Variant, veri() As Variant, n As Long, i As Long, j As Long
    Dim wb As Workbook, c As Worksheet, ac As Boolean, ev As Boolean, sonuc As String, no As Long, metin As String
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(GNL_SAYFA)
    On Error GoTo 0
    If Not ws Is Nothing Then son = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If son < 2 Then
        MsgBox "Henüz işlem kaydı yok." & vbCrLf & vbCrLf & _
            "Güncellemeler, hatalar ve yönetici girişleri bu sürümden itibaren kayda alınır.", vbInformation, "AKSORA | İŞLEM GÜNLÜĞÜ"
        Exit Sub
    End If
    AKSORA_DONUS_KAYDET True
    ac = AKSORA_BEKLE_AC("AKSORA  |  İŞLEM GÜNLÜĞÜ", "Kayıtlar hazırlanıyor...", 0.2)
    ev = Application.EnableEvents
    On Error GoTo Hata
    Application.EnableEvents = False: Application.ScreenUpdating = False
    kaynak = ws.Range(ws.Cells(2, 1), ws.Cells(son, GNL_SUTUN)).Value2
    n = UBound(kaynak, 1)
    If n > GNL_AKTAR_AZAMI Then n = GNL_AKTAR_AZAMI
    ReDim veri(1 To n, 1 To GNL_SUTUN)
    For i = 1 To n
        j = UBound(kaynak, 1) - i + 1
        If IsNumeric(kaynak(j, 1)) And Not IsEmpty(kaynak(j, 1)) Then
            veri(i, 1) = Format$(CDate(kaynak(j, 1)), "dd.mm.yyyy hh:nn:ss")
        Else
            veri(i, 1) = CStr(kaynak(j, 1))
        End If
        veri(i, 2) = CStr(kaynak(j, 2))
        veri(i, 3) = CStr(kaynak(j, 3))
        veri(i, 4) = CStr(kaynak(j, 4))
        veri(i, 5) = CStr(kaynak(j, 5))
        If IsNumeric(kaynak(j, 6)) And Not IsEmpty(kaynak(j, 6)) Then
            veri(i, 6) = AKSORA_SURE_METNI(CDbl(kaynak(j, 6)))
        Else
            veri(i, 6) = vbNullString
        End If
        veri(i, 7) = CStr(kaynak(j, 7))
    Next i
    AKSORA_DONUS_KAYDET
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set c = wb.Worksheets(1)
    AKSORA_EXCEL_SAYFA c, Array("Tarih", "Kullanıcı", "Bilgisayar", "Tür", "İşlem", "Süre", "Dosya"), veri, n, _
        Array(20, 16, 18, 14, 80, 14, 24), Array("@", "@", "@", "@", "@", "@", "@"), _
        Array(xlCenter, xlCenter, xlCenter, xlCenter, xlLeft, xlCenter, xlLeft), _
        "AKSORA  |  İŞLEM GÜNLÜĞÜ  |  en yeni kayıt üstte", _
        "Güncelleme, hata ve yönetici girişi kayıtları", "İşlem Günlüğü", Empty
    For i = 1 To n
        Select Case CStr(veri(i, 4))
            Case "HATA"
                c.Cells(i + 4, 4).Interior.Color = RGB(254, 226, 226): c.Cells(i + 4, 4).Font.Color = RGB(185, 28, 28)
            Case "YÖNETİCİ"
                c.Cells(i + 4, 4).Interior.Color = RGB(226, 236, 247)
        End Select
    Next i
    wb.Worksheets(1).Activate
    sonuc = AKSORA_EXCEL_BITIR(wb, AKSORA_MASAUSTU() & Application.PathSeparator & _
        AKSORA_DOSYA_ADI("AKSORA İşlem Günlüğü - " & Format$(Date, "dd.mm.yyyy")) & ".xlsx", False)
    Application.EnableEvents = ev: Application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    If Len(sonuc) > 0 Then Application.StatusBar = "Masaüstüne kaydedildi: " & sonuc
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.EnableEvents = ev: Application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    MsgBox "İşlem günlüğü Excel'e aktarılamadı." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

'==========================================================
' SISTEM SAGLIK KONTROLU
'==========================================================
Private Sub SK_EKLE(ByVal c As Collection, ByVal kontrol As String, ByVal sonuc As String, ByVal durum As String, _
    Optional ByVal oneri As String = "")
    c.Add Array(kontrol, sonuc, durum, oneri)
End Sub

Private Function SK_BILESEN(ByVal progId As String) As Boolean
    Dim o As Object
    On Error Resume Next
    Set o = CreateObject(progId)
    SK_BILESEN = Not (o Is Nothing)
    Set o = Nothing
End Function

Private Function SK_BIT() As String
#If Win64 Then
    SK_BIT = "64 bit"
#Else
    SK_BIT = "32 bit"
#End If
End Function

' Kontrolleri calistirir. Donus: Collection of Array(kontrol, sonuc, durum, oneri); durum: TAMAM / UYARI / HATA / BİLGİ.
Public Function AKSORA_SAGLIK_HESAPLA() As Collection
    Dim c As Collection, ws As Worksheet, ad As Variant, eksik As String, n As Long, mb As Double, t As String
    Dim nm As Name, bozuk As Long, o As Object, enAd As String, enHucre As Double, h As Double, kilit As Long
    Dim hataSay As Long
    Set c = New Collection
    On Error Resume Next

    SK_EKLE c, "AKSORA sürümü", "Geliştirme paketi v" & AKSORA_GELISTIRME_SURUM & "  •  dosya: " & ThisWorkbook.Name, "BİLGİ"
    SK_EKLE c, "Excel", "Sürüm " & Application.Version & " (" & SK_BIT() & ")  •  " & Application.OperatingSystem, _
        IIf(Val(Application.Version) >= 16, "TAMAM", "UYARI"), _
        IIf(Val(Application.Version) >= 16, "", "Excel 2016 ve üzeri önerilir; eski sürümlerde görünüm ve yazı tipleri bozulabilir.")

    ' Dosya boyutu ve konumu
    mb = FileLen(ThisWorkbook.FullName) / 1048576#
    If Err.Number <> 0 Then mb = 0: Err.Clear
    If mb > 0 Then
        SK_EKLE c, "Dosya boyutu", Format$(mb, "0.0") & " MB", IIf(mb > 40, "UYARI", "TAMAM"), _
            IIf(mb > 40, "Dosya büyüdükçe açılış ve kaydetme yavaşlar. Eski kayıtların (stok hafızası, arşiv) temizlenmesini planlayın.", "")
    End If
    t = ThisWorkbook.Path
    If LCase$(Left$(t, 4)) = "http" Then
        SK_EKLE c, "Dosya konumu", t, "UYARI", "Dosya OneDrive / SharePoint adresinden açılmış. SAP dışa aktarımları ve PDF kayıtları için yerel klasörden açın."
    Else
        SK_EKLE c, "Dosya konumu", t, "TAMAM"
    End If
    If AKSORA_YAYIN_MODU() Then
        SK_EKLE c, "Çalışma kipi", "YAYIN kopyası (izleme)", "BİLGİ", "Bu kopyada SAP güncellemesi kapalıdır."
    Else
        SK_EKLE c, "Çalışma kipi", "Ana dosya (SAP güncellemesi açık)", "TAMAM"
    End If

    ' Zorunlu sayfalar
    For Each ad In Array("BIRIM SECIM", "KAPAK", "ANA PANEL", "CDP_STOK", "EPP_STOK", "İLÇE_STOK", "YÜKLENİCİ_STOK", _
        "STOK_KARŞILAŞTIRMA", "MALZEME_ARA", "M_ANA PANEL", "Y_ANA PANEL", "I_ANA PANEL", "Y_TEMİNAT", "Y_HAKEDİŞ", "_FIYAT_LISTESI")
        Set ws = Nothing
        Set ws = ThisWorkbook.Worksheets(CStr(ad))
        If ws Is Nothing Then eksik = eksik & IIf(Len(eksik) > 0, ", ", "") & CStr(ad)
        Err.Clear
    Next ad
    If Len(eksik) = 0 Then
        SK_EKLE c, "Zorunlu sayfalar", "Tamamı mevcut", "TAMAM"
    Else
        SK_EKLE c, "Zorunlu sayfalar", "Eksik: " & eksik, "HATA", "Eksik sayfalar silinmiş ya da yeniden adlandırılmış. Son yedekten geri yükleyin."
    End If

    ' Sayfa sayisi ve en buyuk sayfa
    n = ThisWorkbook.Worksheets.Count
    For Each ws In ThisWorkbook.Worksheets
        h = 0
        h = CDbl(ws.UsedRange.Rows.Count) * CDbl(ws.UsedRange.Columns.Count)
        If h > enHucre Then enHucre = h: enAd = ws.Name
        Err.Clear
    Next ws
    SK_EKLE c, "Sayfa sayısı", n & " sayfa", IIf(n > 230, "UYARI", "TAMAM"), _
        IIf(n > 230, "Sayfa sayısı çok yüksek; il / birim kayıt sayfalarının veritabanına taşınması değerlendirilmeli.", "")
    SK_EKLE c, "En büyük sayfa", enAd & "  •  " & Format$(enHucre, "#,##0") & " hücre", IIf(enHucre > 6000000#, "UYARI", "BİLGİ"), _
        IIf(enHucre > 6000000#, "Kullanılan alan çok büyük; boş biçimli satır / sütunlar dosyayı şişiriyor olabilir.", "")

    ' Bozuk adlar
    For Each nm In ThisWorkbook.Names
        If InStr(1, nm.RefersTo, "#REF!", vbTextCompare) > 0 Or InStr(1, nm.RefersTo, "#BAŞV!", vbTextCompare) > 0 Then bozuk = bozuk + 1
        Err.Clear
    Next nm
    SK_EKLE c, "Adlandırılmış aralıklar", ThisWorkbook.Names.Count & " ad, " & bozuk & " bozuk başvuru", IIf(bozuk > 0, "UYARI", "TAMAM"), _
        IIf(bozuk > 0, "Bozuk (#BAŞV!) adlar Ad Yöneticisi'nden temizlenmeli.", "")

    ' Bilesenler
    eksik = vbNullString
    For Each ad In Array("Scripting.Dictionary", "Scripting.FileSystemObject", "WScript.Shell", "ADODB.Stream", "MSXML2.DOMDocument")
        If Not SK_BILESEN(CStr(ad)) Then eksik = eksik & IIf(Len(eksik) > 0, ", ", "") & CStr(ad)
    Next ad
    If Len(eksik) = 0 Then
        SK_EKLE c, "Windows bileşenleri", "Tamamı kullanılabilir", "TAMAM"
    Else
        SK_EKLE c, "Windows bileşenleri", "Kullanılamıyor: " & eksik, "HATA", "Bilgi işlem bu bileşenleri etkinleştirmeli; ilgili işlemler çalışmaz."
    End If
    If SK_BILESEN("VBScript.RegExp") Then
        SK_EKLE c, "VBScript.RegExp", "Kullanılabilir", "BİLGİ", _
            "Microsoft, VBScript'i Windows'tan aşamalı olarak kaldırıyor. Metin arama işlevleri ileride yerleşik VBA RegExp'e taşınmalı."
    Else
        SK_EKLE c, "VBScript.RegExp", "KULLANILAMIYOR", "HATA", "VBScript bu bilgisayarda kapalı; düzenli ifade kullanan aramalar çalışmaz."
    End If

    ' SAP
    If Not AKSORA_YAYIN_MODU() Then
        Set o = Nothing
        Set o = GetObject("SAPGUI")
        Err.Clear
        If o Is Nothing Then
            SK_EKLE c, "SAP GUI", "SAP Logon açık değil ya da betik (scripting) erişimi kapalı", "UYARI", _
                "Güncelleme öncesi SAP'de CDP ve EPP sistemlerine giriş yapın; betik erişimi SAP GUI seçeneklerinden açık olmalı."
        Else
            SK_EKLE c, "SAP GUI", "SAP Logon çalışıyor, betik erişimi açık", "TAMAM"
        End If
        Set o = Nothing
    End If

    ' Guvenlik: VBA projesi kilidi
    kilit = -1
    kilit = ThisWorkbook.VBProject.Protection
    If Err.Number <> 0 Then kilit = -1: Err.Clear
    Select Case kilit
        Case 0
            SK_EKLE c, "VBA projesi", "KİLİTSİZ", "UYARI", _
                "Kod ve içindeki şifreler herkesçe okunabilir. VBA projesine parola verin (Araçlar > VBAProject Özellikleri > Koruma)."
        Case 1
            SK_EKLE c, "VBA projesi", "Parola ile kilitli", "TAMAM"
        Case Else
            SK_EKLE c, "VBA projesi", "Kilit durumu okunamadı (VBA nesne modeline erişim kapalı)", "BİLGİ"
    End Select

    ' Excel durumu
    SK_EKLE c, "Excel durumu", "Olaylar " & IIf(Application.EnableEvents, "açık", "KAPALI") & "  •  Hesaplama " & _
        IIf(Application.Calculation = xlCalculationAutomatic, "otomatik", "elle"), _
        IIf(Application.EnableEvents, "TAMAM", "UYARI"), _
        IIf(Application.EnableEvents, "", "Olaylar kapalı kalmış. Alt+F8 > CORE_EXCEL_KURTAR çalıştırın.")

    ' Gunluk ve KPI
    hataSay = AKSORA_GUNLUK_SAY("HATA", 7)
    SK_EKLE c, "Son 7 günde hata", hataSay & " hata kaydı", IIf(hataSay >= 5, "UYARI", "TAMAM"), _
        IIf(hataSay > 0, "Ayrıntı için Yönetim Merkezi > İŞLEM GÜNLÜĞÜ.", "")
    SK_EKLE c, "Son 7 günde güncelleme", AKSORA_GUNLUK_SAY("GÜNCELLEME", 7) & " güncelleme kaydı", "BİLGİ"
    SK_EKLE c, "Yönetim karnesi geçmişi", AKSORA_KPI_KAYIT_SAYISI() & " kayıt", "BİLGİ", "Her yönetim raporu yenilemesinde günde bir kayıt tutulur."

    On Error GoTo 0
    Set AKSORA_SAGLIK_HESAPLA = c
End Function

' Saglik kontrolu sonucu Excel'e aktarilir (masaustu) ve ozet gunluge yazilir.
Public Sub AKSORA_SAGLIK_KONTROL()
    Dim liste As Collection, veri() As Variant, n As Long, i As Long, x As Variant, wb As Workbook, c As Worksheet
    Dim ac As Boolean, ev As Boolean, sonuc As String, no As Long, metin As String, nHata As Long, nUyari As Long
    AKSORA_DONUS_KAYDET True
    ac = AKSORA_BEKLE_AC("AKSORA  |  SAĞLIK KONTROLÜ", "Sistem kontrol ediliyor...", 0.2)
    ev = Application.EnableEvents
    On Error GoTo Hata
    Application.EnableEvents = False: Application.ScreenUpdating = False
    Set liste = AKSORA_SAGLIK_HESAPLA()
    n = liste.Count
    ReDim veri(1 To n, 1 To 4)
    For Each x In liste
        i = i + 1
        veri(i, 1) = CStr(x(0)): veri(i, 2) = CStr(x(1)): veri(i, 3) = CStr(x(2)): veri(i, 4) = CStr(x(3))
        If CStr(x(2)) = "HATA" Then nHata = nHata + 1
        If CStr(x(2)) = "UYARI" Then nUyari = nUyari + 1
    Next x
    AKSORA_DONUS_KAYDET
    Set wb = Workbooks.Add(xlWBATWorksheet)
    Set c = wb.Worksheets(1)
    AKSORA_EXCEL_SAYFA c, Array("Kontrol", "Sonuç", "Durum", "Öneri"), veri, n, _
        Array(28, 60, 12, 90), Array("@", "@", "@", "@"), Array(xlLeft, xlLeft, xlCenter, xlLeft), _
        "AKSORA  |  SİSTEM SAĞLIK RAPORU  |  " & nHata & " hata, " & nUyari & " uyarı", _
        "Kullanıcı: " & AKSORA_KULLANICI(), "Sağlık Raporu", Empty
    For i = 1 To n
        With c.Cells(i + 4, 3)
            .Font.Bold = True
            Select Case CStr(veri(i, 3))
                Case "TAMAM": .Interior.Color = RGB(220, 252, 231): .Font.Color = RGB(21, 128, 61)
                Case "UYARI": .Interior.Color = RGB(254, 243, 199): .Font.Color = RGB(146, 64, 14)
                Case "HATA": .Interior.Color = RGB(254, 226, 226): .Font.Color = RGB(185, 28, 28)
                Case Else: .Interior.Color = RGB(241, 245, 249): .Font.Color = RGB(71, 85, 105)
            End Select
        End With
        c.Cells(i + 4, 2).WrapText = True: c.Cells(i + 4, 4).WrapText = True
    Next i
    wb.Worksheets(1).Activate
    sonuc = AKSORA_EXCEL_BITIR(wb, AKSORA_MASAUSTU() & Application.PathSeparator & _
        AKSORA_DOSYA_ADI("AKSORA Sağlık Raporu - " & Format$(Date, "dd.mm.yyyy")) & ".xlsx", False)
    Application.EnableEvents = ev: Application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    AKSORA_GUNLUK_YAZ "SİSTEM", "Sağlık kontrolü: " & nHata & " hata, " & nUyari & " uyarı"
    If Len(sonuc) > 0 Then Application.StatusBar = "Masaüstüne kaydedildi: " & sonuc
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    Application.EnableEvents = ev: Application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    MsgBox "Sağlık raporu hazırlanamadı." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

'==========================================================
' KURULUM SINAMASI
' Kurulum betigi kaydetmeden once cagirir. Yeni modullerin ve baglanti satiri eklenen
' modullerin (BIRIM, ANALIZ, CEKIRDEK, EKRAN, SISTEM) derlendigini dogrular; veri degistirmez.
' Donus: "OK" ya da "HATA: ..."
'==========================================================
Public Function AKSORA_GELISTIRME_SINAMA() As String
    Dim x As Variant, s As String, b As Boolean
    On Error GoTo Hata
    s = AKSORA_KPI_SINAMA()
    If s <> "OK" Then AKSORA_GELISTIRME_SINAMA = "HATA: " & s: Exit Function
    s = AKSORA_PANO_SINAMA()                ' modAKSORA_PANO ve modAKSORA_PANO_SABLON derlenir
    If s <> "OK" Then AKSORA_GELISTIRME_SINAMA = "HATA: pano " & s: Exit Function
    x = ADM_ILLER()                         ' modAKSORA_BIRIM derlenir
    If UBound(x) <> 4 Then AKSORA_GELISTIRME_SINAMA = "HATA: il listesi beklenen gibi değil": Exit Function
    s = ATIL_TESIS_IL("0000")               ' modAKSORA_ANALIZ derlenir
    b = CORE_AKTIF_MI()                     ' modAKSORA_CEKIRDEK derlenir
    b = AKSORA_PENCERE_KUCUK()              ' modAKSORA_EKRAN derlenir
    s = AKSORA_SAYI_METNI(1, "0")           ' modAKSORA_SISTEM derlenir
    If Abs(AKSORA_AKICI_KONUM(0)) > 0.0001 Or Abs(AKSORA_AKICI_KONUM(1.2) - 1) > 0.0001 Then    ' modAKSORA_AKICI derlenir
        AKSORA_GELISTIRME_SINAMA = "HATA: akıcı kart konum hesabı": Exit Function
    End If
    If Len(AKSORA_KULLANICI()) = 0 Then AKSORA_GELISTIRME_SINAMA = "HATA: kullanıcı adı okunamadı": Exit Function
    AKSORA_GELISTIRME_SINAMA = "OK"
    Exit Function
Hata:
    AKSORA_GELISTIRME_SINAMA = "HATA: " & Err.Description & " (" & Err.Number & ")"
End Function
