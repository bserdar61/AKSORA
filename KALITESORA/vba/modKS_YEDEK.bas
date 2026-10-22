Attribute VB_Name = "modKS_YEDEK"
Option Explicit

' Yedekleme, açýlýþ, VBA dýþa aktarma ve örnek (deneme) verisi.
' 5+ yýllýk saklama: kayýtlar hiçbir makro tarafýndan kendiliðinden silinmez; yedekler ayrý klasörde birikir.

Public Function KS_YedekKopya(ByVal etiket As String) As String
    Dim klasor As String, yol As String, ws As Worksheet, korumali As Boolean
    If Len(ThisWorkbook.Path) = 0 Then Err.Raise vbObjectError + 7201, KS_AD, "Dosya henüz kaydedilmemiþ."
    If LCase$(Left$(ThisWorkbook.Path, 4)) = "http" Then
        Err.Raise vbObjectError + 7202, KS_AD, "Dosya OneDrive / SharePoint adresinden açýlmýþ. " & _
            "Yedek için dosyayý bilgisayardaki bir klasörden açýn."
    End If
    klasor = ThisWorkbook.Path & Application.PathSeparator & "KALITESORA_YEDEK"
    If Len(Dir$(klasor, vbDirectory)) = 0 Then MkDir klasor
    yol = klasor & Application.PathSeparator & "KALITESORA_" & etiket & "_" & Format$(Now, "yyyymmdd_hhnnss") & ".xlsm"
    ThisWorkbook.SaveCopyAs yol
    Set ws = KS_WS(SH_TANIM)
    If Not ws Is Nothing And KS_AdVar("pSonYedek") Then
        korumali = ws.ProtectContents
        KS_Ac ws
        KS_F("pSonYedek").Value = Now
        If korumali Then KS_Kapat ws
    End If
    KS_YedekKopya = yol
End Function

Public Sub KS_YEDEK_AL()
    Dim yol As String
    On Error GoTo Hata
    yol = KS_YedekKopya("ELLE")
    KS_Log "YEDEK", "", yol
    MsgBox "Yedek alýndý:" & vbCrLf & yol, vbInformation, KS_AD
    Exit Sub
Hata:
    MsgBox "Yedek alýnamadý: " & Err.Description, vbExclamation, KS_AD
End Sub

Public Function KS_YedekNotu() As String
    Dim d As Date, gun As Long
    On Error GoTo Son
    If Not KS_Tarih(KS_FV("pSonYedek"), d) Then
        KS_YedekNotu = "     |     HÝÇ YEDEK ALINMADI  -  YEDEK AL"
        Exit Function
    End If
    gun = CLng(Date - d)
    If gun > KS_ParamSayi("pYedekGun", 7) Then
        KS_YedekNotu = "     |     SON YEDEK " & gun & " GÜN ÖNCE  -  YEDEK AL"
    Else
        KS_YedekNotu = "     |     Son yedek: " & Format$(KS_FV("pSonYedek"), "dd.mm.yyyy hh:nn")
    End If
Son:
End Function

' Workbook_Open içinden çaðrýlýr (bkz. ThisWorkbook_KODU.txt).
Public Sub KS_ACILIS()
    Dim ws As Worksheet, d As Date
    On Error GoTo Son
    ' Hazýr dosya ilk açýldýðýnda kurulum kendiliðinden baþlar.
    If Not KS_Kurulu() Then
        KALITESORA_KUR
        Exit Sub
    End If
    Set ws = KS_WS(SH_PANEL)
    KS_PanelYenile False
    ws.Activate
    ActiveWindow.ScrollRow = 1
    If KS_Tarih(KS_FV("pSonYedek"), d) Then
        If Date - d <= KS_ParamSayi("pYedekGun", 7) Then Exit Sub
    End If
    If MsgBox("Son yedekten bu yana " & KS_ParamSayi("pYedekGun", 7) & " günden fazla geçti " & _
        "(ya da hiç yedek alýnmadý)." & vbCrLf & vbCrLf & "Þimdi yedek alýnsýn mý?", vbQuestion + vbYesNo, KS_AD) = vbYes Then
        KS_YEDEK_AL
    End If
Son:
End Sub

Public Sub KS_VBA_DISARI_AKTAR()
    Dim proje As Object, bilesen As Object, fso As Object, klasor As String, uzanti As String, adet As Long
    On Error GoTo Hata
    Set proje = ThisWorkbook.VBProject
    Set fso = CreateObject("Scripting.FileSystemObject")
    klasor = fso.BuildPath(CreateObject("WScript.Shell").SpecialFolders("Desktop"), _
        "KALITESORA_VBA_" & Format$(Now, "yyyymmdd_hhnnss"))
    fso.CreateFolder klasor
    For Each bilesen In proje.VBComponents
        Select Case bilesen.Type
            Case 1: uzanti = ".bas"
            Case 2, 100: uzanti = ".cls"
            Case 3: uzanti = ".frm"
            Case Else: uzanti = ""
        End Select
        If Len(uzanti) > 0 Then
            bilesen.Export fso.BuildPath(klasor, bilesen.Name & uzanti)
            adet = adet + 1
        End If
    Next bilesen
    MsgBox adet & " VBA bileþeni dýþarý aktarýldý:" & vbCrLf & klasor, vbInformation, KS_AD
    Exit Sub
Hata:
    MsgBox "VBA aktarýlamadý: " & Err.Description & vbCrLf & _
        "Dosya > Seçenekler > Güven Merkezi > Makro Ayarlarý > 'VBA proje nesne modeline eriþime güven' açýk olmalý.", _
        vbExclamation, KS_AD
End Sub

' ================================================================ örnek veri

Private Function KS_Rast(ByVal alt As Long, ByVal ust As Long) As Long
    KS_Rast = alt + Int(Rnd() * (ust - alt + 1))
End Function

Public Sub KS_ORNEK_VERI_YUKLE()
    Dim loM As ListObject, loS As ListObject, loA As ListObject, loC As ListObject
    Dim m() As Variant, s() As Variant, a() As Variant, c() As Variant, yer As Variant, aY As Variant
    Dim ad As Variant, kat As Variant, fiy As Variant, gar As Variant, uret As Variant, tip As Variant
    Dim i As Long, j As Long, n As Long, na As Long, mi As Long, sNo As Long, aNo As Long
    Dim plan As Date, gelis As Date, t As Date, bit As Date, bil As Date, don As Date, ilk As Date
    On Error GoTo Hata
    Set loM = KS_LOZ(T_MALZ): Set loS = KS_LOZ(T_SEVK): Set loA = KS_LOZ(T_ARIZA): Set loC = KS_LOZ(T_SEVKCI)
    If loS.ListRows.Count > 0 Or loA.ListRows.Count > 0 Then
        MsgBox "Örnek veri yalnýz sevkiyat ve arýza tablolarý boþken yüklenir; gerçek veriyle karýþmaz.", _
            vbExclamation, KS_AD
        Exit Sub
    End If
    If MsgBox("Deneme amaçlý kurgusal veri yüklenecek:" & vbCrLf & _
        "8 malzeme, 4 kurgusal üretici, ~160 sevkiyat (son 5 yýl), ~50 arýza." & vbCrLf & vbCrLf & _
        "Tüm satýrlar 'ÖRNEK VERÝ' notuyla iþaretlenir ve ÖRNEK VERÝYÝ SÝL ile tek seferde kaldýrýlýr." & vbCrLf & _
        "Devam edilsin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub

    KS_Basla
    Call Rnd(-1)
    Randomize 2026
    yer = KS_Veri(KS_LOZ(T_YER))
    If Not IsArray(yer) Then Err.Raise vbObjectError + 7301, KS_AD, "TANIMLAR'da sevk yeri yok."
    uret = Array("ÖRNEK ÜRETÝCÝ A", "ÖRNEK ÜRETÝCÝ B", "ÖRNEK ÜRETÝCÝ C", "ÖRNEK ÜRETÝCÝ D")
    ReDim c(1 To 4, 1 To loC.ListColumns.Count)
    For i = 1 To 4
        c(i, 1) = uret(i - 1)
        c(i, 4) = "Malzeme"
        c(i, loC.ListColumns.Count) = KS_ORNEK
    Next i
    ad = Array("Daðýtým Trafosu 250 kVA", "OG Kesici 36 kV", "Elektronik Sayaç Trifaze", "XLPE Kablo 3x150 mm2", _
        "Kompozit Ýzolatör 36 kV", "Parafudr 36 kV", "AG Daðýtým Panosu", "Redresör 26 A")
    kat = Array("Trafo", "Kesici / Ayýrýcý", "Sayaç", "Kablo", "Ýzolatör", "Parafudr", "Pano / Hücre", "Diðer")
    fiy = Array(185000, 240000, 3200, 950, 1450, 6800, 42000, 28000)
    gar = Array(24, 24, 36, 12, 24, 24, 24, 24)
    ReDim m(1 To 8, 1 To M_KOLON)
    For i = 1 To 8
        m(i, M_KOD) = "ORN-" & Format$(i, "000")
        m(i, M_AD) = ad(i - 1)
        m(i, M_KAT) = kat(i - 1)
        m(i, M_BIRIM) = IIf(i = 4, "Metre", "Adet")
        m(i, M_URETICI) = uret((i - 1) Mod 4)
        m(i, M_FIYAT) = fiy(i - 1)
        m(i, M_PB) = "TRY"
        m(i, M_GARANTI) = gar(i - 1)
        m(i, M_KAYIT) = DateSerial(Year(Date) - 4, 1, 2)
        m(i, M_NOT) = KS_ORNEK
    Next i

    sNo = KS_EnBuyukNo(loS, "SVK-")
    aNo = KS_EnBuyukNo(loA, "ARZ-")
    ilk = DateSerial(Year(Date) - 4, 1, 1)
    n = 160
    ReDim s(1 To n, 1 To S_KOLON)
    For i = 1 To n
        mi = KS_Rast(1, 8)
        sNo = sNo + 1
        plan = ilk + KS_Rast(0, CLng(Date - ilk) + 60)
        s(i, S_NO) = KS_NoBicim("SVK-", sNo)
        s(i, S_SEVKCI) = uret((mi - 1) Mod 4)
        s(i, S_YER) = yer(KS_Rast(1, UBound(yer, 1)), 1)
        s(i, S_KOD) = m(mi, M_KOD)
        s(i, S_AD) = m(mi, M_AD)
        If mi = 4 Then s(i, S_ADET) = KS_Rast(5, 60) * 50 Else s(i, S_ADET) = KS_Rast(1, 40)
        s(i, S_FIYAT) = Round(m(mi, M_FIYAT) * (0.9 + Rnd() * 0.25), 2)
        s(i, S_PB) = "TRY"
        s(i, S_PLAN) = plan
        If plan < Date - 5 Then
            gelis = plan + KS_Rast(-6, 12) + IIf((mi - 1) Mod 4 = 2, KS_Rast(0, 25), 0)
            If gelis > Date Then gelis = Date
            s(i, S_GELIS) = gelis
        End If
        s(i, S_GARAY) = m(mi, M_GARANTI)
        s(i, S_IRS) = "ORN-IRS-" & Format$(i, "0000")
        s(i, S_NOT) = KS_ORNEK
        s(i, S_ZAMAN) = plan
        s(i, S_KULL) = "ornek"
    Next i

    ' Arýzalar: teslim alýnmýþ sevkiyatlardan; C üreticisi daha sýk arýza verir.
    tip = Array("Elektriksel", "Mekanik", "Ýzolasyon", "Üretim Hatasý", "Montaj Kaynaklý")
    ReDim a(1 To 60, 1 To A_KOLON)
    For j = 1 To 400
        If na >= 52 Then Exit For
        i = KS_Rast(1, n)
        If Not IsEmpty(s(i, S_GELIS)) Then
            If Rnd() < IIf(s(i, S_SEVKCI) = uret(2), 0.55, 0.18) Then
                gelis = s(i, S_GELIS)
                t = gelis + KS_Rast(10, 900)
                If t <= Date Then
                    na = na + 1
                    aNo = aNo + 1
                    bit = DateAdd("m", s(i, S_GARAY), gelis)
                    a(na, A_NO) = KS_NoBicim("ARZ-", aNo)
                    a(na, A_SEVK) = s(i, S_NO)
                    a(na, A_KOD) = s(i, S_KOD)
                    a(na, A_TARIH) = t
                    a(na, A_ADET) = KS_Rast(1, 3)
                    a(na, A_TIP) = tip(KS_Rast(0, 4))
                    a(na, A_ACIK) = "Örnek arýza kaydý: sahada " & LCase$(tip(KS_Rast(0, 4))) & " arýza tespit edildi."
                    bil = t + KS_Rast(0, 4)
                    If bil <= Date Then
                        a(na, A_BILDIRIM) = bil
                        a(na, A_MAIL) = Format$(bil, "dd.mm.yyyy") & " - Arýza üreticiye bildirildi." & vbLf & _
                            "Üretici: inceleme için malzemenin gönderilmesini istedi."
                    End If
                    a(na, A_DEVAM) = "EVET"
                    don = bil + KS_Rast(3, 45) + IIf(s(i, S_SEVKCI) = uret(1), KS_Rast(10, 40), 0)
                    If don <= Date And Not IsEmpty(a(na, A_BILDIRIM)) Then
                        a(na, A_DONUS) = don
                        If Rnd() < 0.85 And don + 7 <= Date Then
                            a(na, A_COZUM) = don + KS_Rast(0, 7)
                            a(na, A_DEVAM) = "HAYIR"
                            a(na, A_COZULDU) = "EVET"
                            a(na, A_SEKIL) = IIf(t <= bit, "Deðiþim", "Tamir")
                            If t <= bit Then
                                a(na, A_MALIYET) = 0
                                a(na, A_KARS) = "Üretici"
                            Else
                                a(na, A_MALIYET) = Round(s(i, S_FIYAT) * (0.1 + Rnd() * 0.3), 2)
                                a(na, A_KARS) = "Kurum"
                            End If
                        End If
                    End If
                    a(na, A_NOT) = KS_ORNEK
                    a(na, A_ZAMAN) = t
                    a(na, A_KULL) = "ornek"
                End If
            End If
        End If
    Next j

    KS_Ac loC.Parent: KS_TopluEkle loC, c, 4: KS_Kapat loC.Parent
    KS_Ac loM.Parent: KS_TopluEkle loM, m, 8: KS_Kapat loM.Parent
    KS_Ac loS.Parent: KS_TopluEkle loS, s, n: KS_Kapat loS.Parent
    KS_Ac loA.Parent: KS_TopluEkle loA, a, na: KS_Kapat loA.Parent
    KS_HesaplaTumu
    KS_Bitir True
    KS_Log "ÖRNEK VERÝ", "", "8 malzeme, " & n & " sevkiyat, " & na & " arýza yüklendi"
    KS_PanelYenile True
    MsgBox "Örnek veri yüklendi: 8 malzeme, " & n & " sevkiyat, " & na & " arýza." & vbCrLf & _
        "Gerçek veriye geçmeden önce ÖRNEK VERÝYÝ SÝL makrosunu çalýþtýrýn.", vbInformation, KS_AD
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Örnek veri yüklenemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Public Sub KS_ORNEK_VERI_SIL()
    Dim toplam As Long
    On Error GoTo Hata
    If MsgBox("'ÖRNEK VERÝ' notlu tüm satýrlar (arýza, sevkiyat, malzeme, kurgusal üretici) silinecek." & vbCrLf & _
        "Gerçek kayýtlara dokunulmaz. Önce yedek alýnacak. Devam edilsin mi?", vbQuestion + vbYesNo, KS_AD) <> vbYes Then Exit Sub
    KS_YedekKopya "ORNEK_SILME_ONCESI"
    KS_Basla
    toplam = KS_OrnekSil(T_ARIZA, A_NOT) + KS_OrnekSil(T_SEVK, S_NOT) + KS_OrnekSil(T_MALZ, M_NOT)
    toplam = toplam + KS_OrnekSil(T_SEVKCI, KS_LOZ(T_SEVKCI).ListColumns.Count)
    KS_HesaplaTumu
    KS_Bitir True
    KS_Log "ÖRNEK VERÝ SÝL", "", toplam & " satýr silindi"
    KS_PanelYenile False
    MsgBox toplam & " örnek satýr silindi.", vbInformation, KS_AD
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Örnek veri silinemedi: " & Err.Description, vbExclamation, KS_AD
End Sub

Private Function KS_OrnekSil(ByVal tablo As String, ByVal kolon As Long) As Long
    Dim lo As ListObject, a As Variant, i As Long
    Set lo = KS_LOZ(tablo)
    a = KS_Veri(lo)
    If Not IsArray(a) Then Exit Function
    KS_Ac lo.Parent
    For i = UBound(a, 1) To 1 Step -1
        If KS_Metin(a(i, kolon)) = KS_ORNEK Then
            lo.ListRows(i).Delete
            KS_OrnekSil = KS_OrnekSil + 1
        End If
    Next i
    KS_Kapat lo.Parent
End Function
