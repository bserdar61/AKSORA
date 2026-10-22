Attribute VB_Name = "modEPP"
Option Explicit

                '==========================================================
                ' AKSORA
                ' modEPP
                '
                ' EPP NORMAL STOK MOTORU
                '==========================================================

                Private Const EPP_SISTEM As String = "EPP"
                Private Const EPP_DEPO_ILK As String = "D001"
                Private Const EPP_DEPO_SON As String = "D999"

                Private Const EPP_STOK_SAYFA As String = "EPP_STOK"
                Private Const EPP_RAPOR_SAYFA As String = "EPP_RAPOR"

                Public Sub EPP_STOK_GUNCELLE_V2()

                Dim wsStok As Worksheet
                Dim wsRapor As Worksheet

                Dim YeniVeri As Variant
                Dim EskiVeri As Variant

                Dim YeniSatirSayisi As Long
                Dim YeniSutunSayisi As Long

                Dim EskiSonSatir As Long
                Dim EskiSonSutun As Long

                Dim ExportSuresi As Double
                Dim ExcelBaslangic As Double
                Dim ExcelSuresi As Double
                Dim ToplamSure As Double

                Dim hataMesaji As String
                Dim hataNo As Long
                Dim basarili As Boolean

                Dim yazmaBasladi As Boolean
                Dim YedekVar As Boolean
                
                If Not AKSORA_SAP_GUNCELLEME_IZNI() Then Exit Sub

                On Error GoTo Hata

                Set wsStok = _
                ThisWorkbook.Worksheets(EPP_STOK_SAYFA)

                Set wsRapor = _
                ThisWorkbook.Worksheets(EPP_RAPOR_SAYFA)

                CORE_ISLEM_BASLAT _
                "EPP stoklarý güncelleniyor"

                CORE_DURUM _
                "EPP yeni stok verisi SAP'tan alýnýyor..."

                basarili = _
                EPP_V2_EXPORT_AL_OZEL( _
                YeniVeri, _
                YeniSatirSayisi, _
                YeniSutunSayisi, _
                ExportSuresi, _
                hataMesaji)

                If Not basarili Then

                ToplamSure = CORE_ISLEM_BITIR()

                If Not ANA_SESSIZ_MOD Then

                    MsgBox _
                        "EPP stok güncellemesi baþlatýlamadý." & _
                        vbCrLf & vbCrLf & _
                        hataMesaji & _
                        vbCrLf & vbCrLf & _
                        "Mevcut EPP stok verileri deðiþtirilmedi.", _
                        vbCritical, _
                        "AKSORA - EPP"

                End If

                Exit Sub

                End If

                If Not IsArray(YeniVeri) Then

                Err.Raise _
                    vbObjectError + 3201, _
                    "EPP_STOK_GUNCELLE_V2", _
                    "SAP verisi geçerli bir dizi olarak alýnamadý."

                End If

                If YeniSatirSayisi <= 0 Then

                Err.Raise _
                    vbObjectError + 3202, _
                    "EPP_STOK_GUNCELLE_V2", _
                    "SAP'tan stok satýrý gelmedi."

                End If

                If YeniSutunSayisi < 12 Then

                Err.Raise _
                    vbObjectError + 3203, _
                    "EPP_STOK_GUNCELLE_V2", _
                    "EPP export yapýsý beklenen 12 sütundan daha az."

                End If

                CORE_DURUM _
                "Mevcut EPP stok verisi güvenliðe alýnýyor..."

                EskiSonSatir = _
                STOK_SON_SATIR(wsStok)

                EskiSonSutun = _
                STOK_SON_SUTUN(wsStok)

                If EskiSonSatir >= 7 And _
                EskiSonSutun >= 1 Then

                If EskiSonSutun > 12 Then
                    EskiSonSutun = 12
                End If

                EskiVeri = _
                    wsStok.Range( _
                        wsStok.Cells(7, 1), _
                        wsStok.Cells( _
                            EskiSonSatir, _
                            EskiSonSutun)).Value2

                YedekVar = True

                Else

                YedekVar = False

                End If

                ExcelBaslangic = Timer

                CORE_DURUM _
                "EPP stok verileri Excel'e yazýlýyor..."

                yazmaBasladi = True

                EPP_V2_BASLIKLARI_YAZ wsStok

                wsStok.Range( _
                "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

                wsStok.Range("A7").Resize( _
                YeniSatirSayisi, _
                12).Value2 = _
                EPP_V2_ILK_12_SUTUN(YeniVeri)

                CORE_DURUM _
                "EPP stok özeti hesaplanýyor..."

                EPP_V2_STOK_OZET_GUNCELLE _
                wsStok, _
                YeniSatirSayisi

                CORE_DURUM _
                "EPP depo raporu hazýrlanýyor..."

                EPP_V2_RAPOR_OLUSTUR _
                wsStok, _
                wsRapor, _
                YeniSatirSayisi

                EPP_V2_DEPO_LISTESI wsStok

                CORE_DURUM _
                "EPP görünümü düzenleniyor..."

                EPP_V2_GORSEL_DUZEN _
                wsStok, _
                YeniSatirSayisi

    STOK_GORUNUM_TEK "EPP_STOK"
    RAPOR_GORUNUM_TEK "EPP_RAPOR"

                ExcelSuresi = _
                STOK_GECEN_SURE(ExcelBaslangic)

                ToplamSure = _
                CORE_ISLEM_BITIR()

                application.Calculation = xlCalculationAutomatic
                application.CalculateFull

                If Not ANA_SESSIZ_MOD Then

                MsgBox _
                    "EPP stok güncellemesi baþarýyla tamamlandý." & _
                    vbCrLf & vbCrLf & _
                    "Üretim Yeri : " & EPP_UY & vbCrLf & _
                    "Depo Aralýðý: " & _
                    EPP_DEPO_ILK & " - " & EPP_DEPO_SON & vbCrLf & _
                    "Stok Satýrý : " & _
                    AKSORA_SAYI_METNI(YeniSatirSayisi, "#,##0") & vbCrLf & _
                    "SAP + Export : " & _
                    AKSORA_SAYI_METNI(ExportSuresi, "0.00") & " sn" & vbCrLf & _
                    "Excel Ýþleme: " & _
                    AKSORA_SAYI_METNI(ExcelSuresi, "0.00") & " sn" & vbCrLf & _
                    "Toplam Süre : " & _
                    AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn", _
                    vbInformation, _
                    "AKSORA - EPP"

                End If

                Exit Sub

Hata:

                hataNo = Err.Number
                hataMesaji = Err.Description

                On Error Resume Next

                If yazmaBasladi Then

                CORE_DURUM _
                    "Hata oluþtu. Eski EPP stok verisi geri yükleniyor..."

                wsStok.Range( _
                    "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

                If YedekVar Then

                    If IsArray(EskiVeri) Then

                        wsStok.Range("A7").Resize( _
                            UBound(EskiVeri, 1), _
                            UBound(EskiVeri, 2)).Value2 = _
                            EskiVeri

                    End If

                End If

                End If

                On Error GoTo 0

                ToplamSure = _
                CORE_ISLEM_BITIR()

                If Not ANA_SESSIZ_MOD Then

                MsgBox _
                    "EPP stok güncellemesi sýrasýnda hata oluþtu." & _
                    vbCrLf & vbCrLf & _
                    "Hata No : " & hataNo & vbCrLf & _
                    "Açýklama: " & hataMesaji & _
                    vbCrLf & vbCrLf & _
                    "Yeni veri tamamlanamadý." & vbCrLf & _
                    "Eski stok verisi korunmaya çalýþýldý.", _
                    vbCritical, _
                    "AKSORA - EPP"

                End If
                
If ANA_SESSIZ_MOD Then
    Err.Raise hataNo, "EPP_STOK_GUNCELLE_V2", hataMesaji
End If
                End Sub


                Public Function EPP_V2_EXPORT_AL_OZEL( _
                ByRef VeriDizisi As Variant, _
                ByRef veriSatiri As Long, _
                ByRef SutunSayisi As Long, _
                ByRef ExportSuresi As Double, _
                ByRef hataMesaji As String) As Boolean

                Dim SAPSession As Object
                Dim alv As Object
                Dim KolonSirasi As Object
                Dim KolonID(1 To 12) As String
                Dim ham() As Variant
                Dim sonuc() As Variant

                Dim baslangic As Double
                Dim satirSayisi As Long
                Dim r As Long
                Dim c As Long
                Dim n As Long
                Dim uy As String
                Dim depo As String
                Dim malzeme As String
                Dim HizliOkuma As Boolean
                Dim KontrolSatiri As Long
                Dim KontrolSutunu As Long
                Dim KontrolDegeri As Variant
                Dim HizliHataNedeni As String

                On Error GoTo Hata

                EPP_V2_EXPORT_AL_OZEL = False
                veriSatiri = 0
                SutunSayisi = 0
                ExportSuresi = 0
                hataMesaji = vbNullString
                baslangic = Timer
                ' Hizli yol: ilin MB52 listesi tek sorguda metin olarak alinir (olmazsa asagidaki eski yontem).
                If AKHZ_TEK_AL(2, VeriDizisi, veriSatiri, SutunSayisi, hataMesaji) Then EPP_V2_EXPORT_AL_OZEL = True: Exit Function

                Set SAPSession = SAP_SESSION_AL(EPP_SISTEM)

                If SAPSession Is Nothing Then
                    hataMesaji = "EPP SAP oturumu bulunamadý."
                    Exit Function
                End If

                SAPSession.FindById("wnd[0]/tbar[0]/okcd").Text = "/nMB52"
                SAPSession.FindById("wnd[0]").SendVKey 0

                EPP_V2_KISA_BEKLE 0.3

                AKSORA_SAP_MALZEME_TEMIZLE SAPSession
                SAPSession.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = EPP_UY
                SAPSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = EPP_DEPO_ILK
                SAPSession.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = EPP_DEPO_SON
                SAPSession.FindById("wnd[0]").SendVKey 8

                EPP_V2_KISA_BEKLE 0.5

                If Not EPP_V2_PANEL_SEC(SAPSession, hataMesaji) Then
                    Exit Function
                End If

                Set alv = SAPSession.FindById( _
                    "wnd[0]/usr/cntlGRID1/shellcont/shell")

                On Error Resume Next
                alv.PressToolbarButton "&DELETE_FILTER"
                Err.Clear
                On Error GoTo Hata

                EPP_V2_KISA_BEKLE 0.3

                Set KolonSirasi = alv.ColumnOrder

                If KolonSirasi Is Nothing Then
                    hataMesaji = "EPP MB52 sütun sýrasý alýnamadý."
                    Exit Function
                End If

                If CLng(KolonSirasi.Count) < 12 Then
                    hataMesaji = "EPP PANEL düzeninde 12 sütun bulunamadý."
                    Exit Function
                End If

                For c = 1 To 12
                    KolonID(c) = CStr(KolonSirasi.Item(c - 1))
                Next c

                satirSayisi = CLng(alv.RowCount)

                If satirSayisi <= 0 Then
                    hataMesaji = "EPP MB52 sonuç tablosunda stok satýrý bulunamadý."
                    Exit Function
                End If

                ReDim ham(1 To satirSayisi, 1 To 12)

                ' SAP ALV ekran disindaki satirlari bos dondurebilir.
                ' Otuzar satirlik bloklarda kaydir; hatali blokta tek satirlik
                ' orijinal guvenli okumaya don. Hicbir satiri sessizce atlama.
                Dim BlokBas As Long
                Dim BlokSon As Long
                Dim BlokBoyu As Long
                Dim SatirHatasi As Boolean
                Dim tekrar As Long

                BlokBoyu = AKSORA_SAP_BLOK_BOYU(alv)
                For BlokBas = 0 To satirSayisi - 1 Step BlokBoyu
                    BlokSon = BlokBas + BlokBoyu - 1
                    If BlokSon >= satirSayisi Then BlokSon = satirSayisi - 1

                    On Error Resume Next
                    Err.Clear
                    alv.FirstVisibleRow = BlokBas
                    alv.SetCurrentCell BlokBas, KolonID(7)
                    If Err.Number <> 0 Then
                        hataMesaji = "EPP ALV blok kaydirma hatasi: satir " & (BlokBas + 1) & ": " & Err.Description
                        On Error GoTo Hata
                        Exit Function
                    End If
                    On Error GoTo Hata
                    DoEvents
                    EPP_V2_KISA_BEKLE 0.015

                    AKSORA_TOPLU_SURE_TIK
                    For r = BlokBas To BlokSon
                        SatirHatasi = False
                        On Error Resume Next
                        For c = 1 To 12
                            Err.Clear
                            ham(r + 1, c) = alv.GetCellValue(r, KolonID(c))
                            If Err.Number <> 0 Then SatirHatasi = True: Exit For
                        Next c
                        On Error GoTo Hata

                        ' Gorunmeyen satirlar hata vermeden bos gelebiliyor.
                        If Trim$(CStr(ham(r + 1, 1))) <> EPP_UY Or _
                           Len(Trim$(CStr(ham(r + 1, 7)))) = 0 Then SatirHatasi = True

                        If SatirHatasi Then
                            ' Bu satiri orijinal yontemle yeniden oku.
                            For tekrar = 1 To 2
                                On Error Resume Next
                                Err.Clear
                                alv.FirstVisibleRow = r
                                alv.SetCurrentCell r, KolonID(7)
                                SatirHatasi = (Err.Number <> 0)
                                On Error GoTo Hata
                                If Not SatirHatasi Then
                                    DoEvents
                                    EPP_V2_KISA_BEKLE 0.015
                                    On Error Resume Next
                                    For c = 1 To 12
                                        Err.Clear
                                        ham(r + 1, c) = alv.GetCellValue(r, KolonID(c))
                                        If Err.Number <> 0 Then SatirHatasi = True: Exit For
                                    Next c
                                    On Error GoTo Hata
                                    If Trim$(CStr(ham(r + 1, 1))) <> EPP_UY Or _
                                       Len(Trim$(CStr(ham(r + 1, 7)))) = 0 Then SatirHatasi = True
                                End If
                                If Not SatirHatasi Then Exit For
                            Next tekrar
                            If SatirHatasi Then
                                hataMesaji = "EPP ALV satiri dogrulanamadi: " & (r + 1) & _
                                              ". Eksik veri yazilmadi."
                                Exit Function
                            End If
                        End If
                    Next r
                Next BlokBas

                ReDim sonuc(1 To satirSayisi, 1 To 12)
                n = 0

                For r = 1 To satirSayisi

                    uy = Trim$(CStr(ham(r, 1)))
                    depo = UCase$(Trim$(CStr(ham(r, 2))))
                    malzeme = Trim$(CStr(ham(r, 7)))

                    If uy = EPP_UY And _
                       Len(depo) = 4 And _
                       Left$(depo, 1) = "D" And _
                       depo >= EPP_DEPO_ILK And _
                       depo <= EPP_DEPO_SON And _
                       Len(malzeme) > 0 Then

                        n = n + 1

                        For c = 1 To 12
                            sonuc(n, c) = ham(r, c)
                        Next c

                        sonuc(n, 9) = EPP_V2_SAP_SAYI_CEVIR(sonuc(n, 9))
                        sonuc(n, 11) = EPP_V2_SAP_SAYI_CEVIR(sonuc(n, 11))
                        sonuc(n, 12) = EPP_V2_SAP_SAYI_CEVIR(sonuc(n, 12))

                    End If

                Next r

                If n <= 0 Then
                    hataMesaji = "EPP geçerli stok satýrý bulunamadý."
                    Exit Function
                End If

                ReDim Preserve sonuc(1 To satirSayisi, 1 To 12)

                If n < satirSayisi Then
                    Dim temiz() As Variant
                    ReDim temiz(1 To n, 1 To 12)
                    For r = 1 To n
                        For c = 1 To 12
                            temiz(r, c) = sonuc(r, c)
                        Next c
                    Next r
                    VeriDizisi = temiz
                Else
                    VeriDizisi = sonuc
                End If

                veriSatiri = n
                SutunSayisi = 12
                ExportSuresi = EPP_V2_GECEN_SURE(baslangic)

                EPP_V2_EXPORT_AL_OZEL = True
                Exit Function

Hata:
                hataMesaji = _
                    "EPP stok verisi SAP tablosundan alýnamadý." & _
                    vbCrLf & _
                    "Hata No: " & Err.Number & _
                    vbCrLf & _
                    "Açýklama: " & Err.Description

                ExportSuresi = EPP_V2_GECEN_SURE(baslangic)
                EPP_V2_EXPORT_AL_OZEL = False

                End Function

                Private Function EPP_V2_PANEL_SEC( _
                ByVal SAPSession As Object, _
                ByRef hataMesaji As String) As Boolean

                Dim LayoutGrid As Object
                Dim satir As Long
                Dim satirSayisi As Long
                Dim BulunanSatir As Long
                Dim DuzenAdi As String
                Dim DuzenTanimi As String

                On Error GoTo Hata

                EPP_V2_PANEL_SEC = False
                BulunanSatir = -1

                SAPSession.FindById("wnd[0]/tbar[1]/btn[33]").Press

                If Not EPP_V2_PENCERE_BEKLE(SAPSession, "wnd[1]", 10) Then
                    hataMesaji = "EPP MB52 düzen seçme penceresi açýlamadý."
                    Exit Function
                End If

                Set LayoutGrid = SAPSession.FindById( _
                    "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")

                satirSayisi = LayoutGrid.RowCount

                For satir = 0 To satirSayisi - 1

                    DuzenAdi = vbNullString
                    DuzenTanimi = vbNullString

                    On Error Resume Next
                    DuzenAdi = Trim$(CStr(LayoutGrid.GetCellValue(satir, "VARIANT")))
                    DuzenTanimi = Trim$(CStr(LayoutGrid.GetCellValue(satir, "TEXT")))
                    On Error GoTo Hata

                    If StrComp(DuzenAdi, "PANEL", vbTextCompare) = 0 Or _
                       StrComp(DuzenTanimi, "MLZ STOK PANEL", vbTextCompare) = 0 Then
                        BulunanSatir = satir
                        Exit For
                    End If

                Next satir

                If BulunanSatir < 0 Then
                    hataMesaji = "EPP SAP MB52 içinde PANEL düzeni bulunamadý."
                    On Error Resume Next
                    SAPSession.FindById("wnd[1]").SendVKey 12
                    On Error GoTo 0
                    Exit Function
                End If

                LayoutGrid.SetCurrentCell BulunanSatir, "TEXT"
                LayoutGrid.SelectedRows = CStr(BulunanSatir)
                LayoutGrid.ClickCurrentCell

                EPP_V2_KISA_BEKLE 0.3
                EPP_V2_PANEL_SEC = True
                Exit Function

Hata:
                hataMesaji = _
                    "EPP PANEL düzeni seçilemedi." & vbCrLf & _
                    "Hata No: " & Err.Number & vbCrLf & _
                    "Açýklama: " & Err.Description

                On Error Resume Next
                SAPSession.FindById("wnd[1]").SendVKey 12
                On Error GoTo 0

                EPP_V2_PANEL_SEC = False

                End Function

                Private Function EPP_V2_PENCERE_BEKLE( _
                ByVal SAPSession As Object, _
                ByVal NesneID As String, _
                ByVal MaksSaniye As Double) As Boolean

                Dim baslangic As Double
                Dim obj As Object

                baslangic = Timer

                Do
                    Set obj = Nothing
                    On Error Resume Next
                    Set obj = SAPSession.FindById(NesneID)
                    On Error GoTo 0

                    If Not obj Is Nothing Then
                        EPP_V2_PENCERE_BEKLE = True
                        Exit Function
                    End If

                    DoEvents
                    EPP_V2_KISA_BEKLE 0.05

                    If EPP_V2_GECEN_SURE(baslangic) >= MaksSaniye Then Exit Do
                Loop

                EPP_V2_PENCERE_BEKLE = False

                End Function

                Private Function EPP_V2_SAP_SAYI_CEVIR(ByVal deger As Variant) As Double

                Dim s As String
                Dim negatif As Boolean

                On Error GoTo Cikis

                If IsNumeric(deger) Then
                    EPP_V2_SAP_SAYI_CEVIR = CDbl(deger)
                    Exit Function
                End If

                s = Trim$(CStr(deger))

                If Len(s) = 0 Or s = "-" Then Exit Function

                If Left$(s, 1) = "-" Then
                    negatif = True
                    s = Mid$(s, 2)
                ElseIf Right$(s, 1) = "-" Then
                    negatif = True
                    s = Left$(s, Len(s) - 1)
                End If

                s = Replace(s, " ", vbNullString)

                If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
                    If InStrRev(s, ",") > InStrRev(s, ".") Then
                        s = Replace(s, ".", vbNullString)
                        s = Replace(s, ",", application.DecimalSeparator)
                    Else
                        s = Replace(s, ",", vbNullString)
                        s = Replace(s, ".", application.DecimalSeparator)
                    End If
                ElseIf InStr(s, ",") > 0 Then
                    s = Replace(s, ",", application.DecimalSeparator)
                ElseIf InStr(s, ".") > 0 Then
                    If application.DecimalSeparator <> "." Then
                        s = Replace(s, ".", application.DecimalSeparator)
                    End If
                End If

                If IsNumeric(s) Then
                    EPP_V2_SAP_SAYI_CEVIR = CDbl(s)
                    If negatif Then EPP_V2_SAP_SAYI_CEVIR = -EPP_V2_SAP_SAYI_CEVIR
                End If

Cikis:
                End Function

                Private Sub EPP_V2_KISA_BEKLE(ByVal saniye As Double)

                Dim baslangic As Double
                baslangic = Timer

                Do While EPP_V2_GECEN_SURE(baslangic) < saniye
                    DoEvents
                Loop

                End Sub

                Private Function EPP_V2_GECEN_SURE(ByVal baslangic As Double) As Double

                If Timer >= baslangic Then
                    EPP_V2_GECEN_SURE = Timer - baslangic
                Else
                    EPP_V2_GECEN_SURE = (86400# - baslangic) + Timer
                End If

                End Function

                Public Sub EPP_V2_VERIDEN_GUNCELLE(ByRef TumVeri As Variant, ByVal TumSatir As Long)

Dim wsStok As Worksheet
Dim wsRapor As Worksheet
Dim veri() As Variant
Dim TemizVeri() As Variant
Dim i As Long, j As Long, n As Long
Dim depo As String

Set wsStok = ThisWorkbook.Worksheets("EPP_STOK")
Set wsRapor = ThisWorkbook.Worksheets("EPP_RAPOR")

ReDim veri(1 To TumSatir, 1 To 12)
n = 0

For i = 1 To TumSatir
    depo = UCase$(Trim$(CStr(TumVeri(i, 2))))
        If Trim$(CStr(TumVeri(i, 1))) = EPP_UY() And _
       Len(depo) = 4 And Left$(depo, 1) = "D" And _
       depo >= "D001" And depo <= "D999" Then
        n = n + 1
        For j = 1 To 12
            veri(n, j) = TumVeri(i, j)
        Next j
    End If
Next i

If n <= 0 Then Err.Raise vbObjectError + 7101, "EPP_V2_VERIDEN_GUNCELLE", "EPP normal stok verisi bulunamadý."

ReDim TemizVeri(1 To n, 1 To 12)
For i = 1 To n
    For j = 1 To 12
        TemizVeri(i, j) = veri(i, j)
    Next j
Next i

EPP_V2_BASLIKLARI_YAZ wsStok
wsStok.Range("A7:L" & AKSORA_V13_SON(wsStok)).ClearContents
wsStok.Range("A7").Resize(n, 12).Value2 = TemizVeri
EPP_V2_STOK_OZET_GUNCELLE wsStok, n
EPP_V2_RAPOR_OLUSTUR wsStok, wsRapor, n
EPP_V2_DEPO_LISTESI wsStok
EPP_V2_GORSEL_DUZEN wsStok, n

End Sub

Private Function EPP_V2_ILK_12_SUTUN( _
                ByRef kaynak As Variant) As Variant

                Dim sonuc() As Variant
                Dim satir As Long
                Dim sutun As Long
                Dim satirSayisi As Long
                Dim KaynakSutunSayisi As Long

                satirSayisi = _
                UBound(kaynak, 1)

                KaynakSutunSayisi = _
                UBound(kaynak, 2)

                If KaynakSutunSayisi < 12 Then

                Err.Raise _
                    vbObjectError + 3210, _
                    "EPP_V2_ILK_12_SUTUN", _
                    "Kaynak veri 12 sütundan az."

                End If

                ReDim sonuc( _
                1 To satirSayisi, _
                1 To 12)

                For satir = 1 To satirSayisi

                For sutun = 1 To 12

                    sonuc(satir, sutun) = _
                        kaynak(satir, sutun)

                Next sutun

                Next satir

                EPP_V2_ILK_12_SUTUN = sonuc

                End Function

                Private Sub EPP_V2_BASLIKLARI_YAZ( _
                ByVal ws As Worksheet)

                ws.Range("A6:L6").Value = Array( _
                "ÜY", _
                "DpYr", _
                "Depo yeri tanýmý", _
                "Mal grubu", _
                "Mal grubu tanýmý", _
                "Parti", _
                "Malzeme", _
                "Malzeme kýsa metni", _
                "Tahditsiz", _
                "Birim", _
                "Deðer (TL)", _
                "Birim Fiyat (TL)")

                End Sub

                Private Sub EPP_V2_STOK_OZET_GUNCELLE( _
                ByVal ws As Worksheet, _
                ByVal veriSatiri As Long)

                Dim sonSatir As Long
                Dim toplamDeger As Double

                sonSatir = _
                6 + veriSatiri

                toplamDeger = 0

                If veriSatiri > 0 Then

                On Error Resume Next

                toplamDeger = _
                    application.WorksheetFunction.Sum( _
                        ws.Range( _
                            "K7:K" & sonSatir))

                On Error GoTo 0

                End If

                ws.Range("A2").Value = _
                "Son Güncelleme"

                ws.Range("A3").Value = _
                "Toplam Stok Satýrý"

                ws.Range("A4").Value = _
                "Toplam Stok Deðeri"

                On Error Resume Next

                ws.Range("C2:E2").UnMerge
                ws.Range("C3:E3").UnMerge
                ws.Range("C4:E4").UnMerge

                On Error GoTo 0

                ws.Range("C2:E2").Merge
                ws.Range("C3:E3").Merge
                ws.Range("C4:E4").Merge

                ws.Range("C2").Value = Now

                If veriSatiri > 0 Then
                    ws.Range("C3").Formula = _
                        "=SUBTOTAL(103,A7:A" & sonSatir & ")"
                    ws.Range("C4").Formula = _
                        "=SUBTOTAL(109,K7:K" & sonSatir & ")"
                Else
                    ws.Range("C3").Value = 0
                    ws.Range("C4").Value = 0
                End If

                ws.Range("C2").NumberFormat = _
                "dd.mm.yyyy hh:mm:ss"

                ws.Range("C3").NumberFormat = _
                "#,##0"

                ws.Range("C4").NumberFormat = _
                "#,##0.00 ""TL"""

                
    ' Son satýrlar dahil tüm canlý stok alanýný ayný biçime getir
    With ws.Range("A7:L" & sonSatir)
        .Borders.LineStyle = xlNone
        .Font.Name = ws.Range("A7").Font.Name
        .Font.Size = ws.Range("A7").Font.Size
        .Font.bold = False
        .VerticalAlignment = xlCenter
    End With

    Dim rr As Long
    For rr = 7 To sonSatir
        If (rr - 7) Mod 2 = 0 Then
            ws.Range("A" & rr & ":L" & rr).Interior.Color = RGB(242, 248, 252)
        Else
            ws.Range("A" & rr & ":L" & rr).Interior.Color = RGB(255, 255, 255)
        End If
    Next rr

    ws.Range("I7:I" & sonSatir).NumberFormat = "#,##0.###"
    ws.Range("K7:L" & sonSatir).NumberFormat = "#,##0.00 [$TL-tr-TR]"

    ' Nihai kural: stok tablosunda hiçbir kenarlýk olmayacak
    ws.Range("A7:L" & sonSatir).Borders.LineStyle = xlNone

End Sub

                Public Sub EPP_V2_RAPOR_OLUSTUR( _
                ByVal wsStok As Worksheet, _
                ByVal wsRapor As Worksheet, _
                ByVal veriSatiri As Long)

                Dim dict As Object
                Dim veri As Variant
                Dim bilgi As Variant

                Dim i As Long
                Dim HedefSatir As Long
                Dim SonStokSatir As Long

                Dim depo As String
                Dim DepoTanimi As String
                Dim deger As Double

                Dim anahtar As Variant

                Dim ToplamSatir As Long
                Dim toplamDeger As Double
                Dim TasarimSatir As Long

                Set dict = _
                CreateObject("Scripting.Dictionary")

                dict.CompareMode = vbTextCompare

                SonStokSatir = _
                6 + veriSatiri

                If veriSatiri > 0 Then

                veri = _
                    wsStok.Range( _
                        "A7:K" & SonStokSatir).Value2

                For i = 1 To UBound(veri, 1)

                    If Trim$(CStr(veri(i, 1))) = EPP_UY Then

                        depo = _
                            Trim$(CStr(veri(i, 2)))

                        DepoTanimi = _
                            CStr(veri(i, 3))

                        If Len(depo) > 0 Then

                            deger = 0

                            If IsNumeric(veri(i, 11)) Then
                                deger = CDbl(veri(i, 11))
                            End If

                            If dict.Exists(depo) Then

                                bilgi = dict(depo)

                                bilgi(1) = _
                                    CLng(bilgi(1)) + 1

                                bilgi(2) = _
                                    CDbl(bilgi(2)) + deger

                                If Len(CStr(bilgi(0))) = 0 Then
                                    bilgi(0) = DepoTanimi
                                End If

                                dict(depo) = bilgi

                            Else

                                dict.Add _
                                    depo, _
                                    Array( _
                                        DepoTanimi, _
                                        1&, _
                                        deger)

                            End If

                        End If

                    End If

                Next i

                End If

                wsRapor.Range( _
                "A9:D" & AKSORA_V13_SON(wsRapor)).ClearContents

                wsRapor.Range("A8").Value = _
                "Depo yeri"

                wsRapor.Range("B8").Value = _
                "Depo yeri tanýmý"

                wsRapor.Range("C8").Value = _
                "Stok Satýrý"

                wsRapor.Range("D8").Value = _
                "Toplam Stok Deðeri"

                HedefSatir = 9
                ToplamSatir = 0
                toplamDeger = 0

                For Each anahtar In dict.keys

                bilgi = dict(anahtar)

                wsRapor.Cells( _
                    HedefSatir, 1).Value = _
                    anahtar

                wsRapor.Cells( _
                    HedefSatir, 2).Value = _
                    bilgi(0)

                wsRapor.Cells( _
                    HedefSatir, 3).Value = _
                    CLng(bilgi(1))

                wsRapor.Cells( _
                    HedefSatir, 4).Value = _
                    CDbl(bilgi(2))

                ToplamSatir = _
                    ToplamSatir + CLng(bilgi(1))

                toplamDeger = _
                    toplamDeger + CDbl(bilgi(2))

                HedefSatir = _
                    HedefSatir + 1

                Next anahtar

                wsRapor.Cells( _
                HedefSatir, 1).Value = _
                "TOPLAM"

                wsRapor.Cells( _
                HedefSatir, 3).Value = _
                ToplamSatir

                wsRapor.Cells( _
                HedefSatir, 4).Value = _
                toplamDeger

                wsRapor.Range( _
                "A" & HedefSatir & _
                ":D" & HedefSatir).Font.bold = True

                wsRapor.Range("B3").Value = Now
                wsRapor.Range("B4").Value = dict.Count
                wsRapor.Range("B5").Value = toplamDeger

                wsRapor.Range("B3").NumberFormat = _
                "dd.mm.yyyy hh:mm:ss"

                wsRapor.Range("B4").NumberFormat = _
                "#,##0"

                wsRapor.Range("B5").NumberFormat = _
                "#,##0.00 ""TL"""

                If HedefSatir >= 9 Then

                wsRapor.Range( _
                    "C9:C" & HedefSatir).NumberFormat = _
                    "#,##0"

                wsRapor.Range( _
                    "D9:D" & HedefSatir).NumberFormat = _
                    "#,##0.00 ""TL"""

                End If

                On Error Resume Next

                If wsRapor.AutoFilterMode Then
                wsRapor.AutoFilterMode = False
                End If

                On Error GoTo 0

                If HedefSatir > 9 Then
                    wsRapor.Range( _
                        "A8:D" & HedefSatir - 1).AutoFilter
                Else
                    wsRapor.Range("A8:D8").AutoFilter
                End If

                '==========================================================
                ' EPP_RAPOR NÝHAÝ TABLO TASARIMI
                ' CDP_RAPOR ile tamamen ayný standart
                '==========================================================

                wsRapor.Range("A8:D" & HedefSatir).Borders.LineStyle = xlNone

                With wsRapor.Range("A8:D8")
                    .Interior.Color = RGB(31, 78, 121)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                End With

                For TasarimSatir = 9 To HedefSatir - 1
                    With wsRapor.Range("A" & TasarimSatir & ":D" & TasarimSatir)
                        If (TasarimSatir - 9) Mod 2 = 0 Then
                            .Interior.Color = RGB(242, 248, 252)
                        Else
                            .Interior.Color = RGB(255, 255, 255)
                        End If
                        .Font.Color = RGB(0, 0, 0)
                        .Font.bold = False
                    End With
                Next TasarimSatir

                With wsRapor.Range("A" & HedefSatir & ":D" & HedefSatir)
                    .Interior.Color = RGB(221, 235, 247)
                    .Font.Color = RGB(0, 0, 0)
                    .Font.bold = True
                End With

                wsRapor.Range( _
                    "A" & HedefSatir + 1 & _
                    ":D" & AKSORA_V13_SON(wsRapor)).ClearFormats

                wsRapor.Columns("A").ColumnWidth = 14
                wsRapor.Columns("B").ColumnWidth = 32
                wsRapor.Columns("C").ColumnWidth = 16
                wsRapor.Columns("D").ColumnWidth = 22

                End Sub

                Private Sub EPP_V2_DEPO_LISTESI( _
                ByVal ws As Worksheet)

                Dim MevcutSecim As String
                Dim SonVeriSatiri As Long
                Dim SonListeSatiri As Long
                Dim i As Long

                Dim DepoTanimi As String
                Dim dict As Object
                Dim anahtar As Variant

                Dim nm As Name
                Dim BulunanAd As Name
                Dim KisaAd As String
                Dim RefersToMetni As String

                MevcutSecim = Trim$(CStr(ws.Range("H2").Value))

                If ws.Range("G2").MergeCells Then
                    ws.Range("G2").MergeArea.UnMerge
                End If

                If ws.Range("H2").MergeCells Then
                    ws.Range("H2").MergeArea.UnMerge
                End If

                ws.Range("G2").Value = "DEPOLAR"

                ws.Range("N2:N" & AKSORA_V13_SON(ws)).ClearContents

                Set dict = CreateObject("Scripting.Dictionary")
                dict.CompareMode = vbTextCompare

                SonVeriSatiri = 6 + EPP_V2_LONG_DEGER(ws.Range("C3"))

                If SonVeriSatiri >= 7 Then

                    For i = 7 To SonVeriSatiri

                        DepoTanimi = Trim$(CStr(ws.Cells(i, 3).Value2))

                        If Len(DepoTanimi) > 0 Then
                            If Not dict.Exists(DepoTanimi) Then
                                dict.Add DepoTanimi, DepoTanimi
                            End If
                        End If

                    Next i

                End If

                ws.Range("N2").Value = "TÜMÜ"
                SonListeSatiri = 2

                For Each anahtar In dict.keys
                    SonListeSatiri = SonListeSatiri + 1
                    ws.Cells(SonListeSatiri, "N").Value = CStr(anahtar)
                Next anahtar

                Set BulunanAd = Nothing

                For Each nm In ThisWorkbook.names

                    KisaAd = nm.Name

                    If InStrRev(KisaAd, "!") > 0 Then
                        KisaAd = Mid$(KisaAd, InStrRev(KisaAd, "!") + 1)
                    End If

                    KisaAd = Replace(KisaAd, "'", "")

                    If StrComp(KisaAd, "_EPP_DEPO_LISTESI", vbTextCompare) = 0 Then
                        Set BulunanAd = nm
                        Exit For
                    End If

                Next nm

                RefersToMetni = _
                "='" & Replace(ws.Name, "'", "''") & _
                "'!$N$2:$N$" & SonListeSatiri

                If BulunanAd Is Nothing Then
                    ThisWorkbook.names.Add _
                        Name:="_EPP_DEPO_LISTESI", _
                        RefersTo:=RefersToMetni
                Else
                    BulunanAd.RefersTo = RefersToMetni
                End If

                On Error Resume Next
                ws.Range("H2").Validation.Delete
                On Error GoTo 0

                ws.Range("H2").Validation.Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="=_EPP_DEPO_LISTESI"

                With ws.Range("H2").Validation
                    .IgnoreBlank = True
                    .InCellDropdown = True
                    .ShowError = True
                End With

                If StrComp(MevcutSecim, "TÜMÜ", vbTextCompare) = 0 Then
                    ws.Range("H2").Value = "TÜMÜ"
                ElseIf Len(MevcutSecim) > 0 And dict.Exists(MevcutSecim) Then
                    ws.Range("H2").Value = MevcutSecim
                Else
                    ws.Range("H2").Value = "TÜMÜ"
                End If

                ws.Columns("N").Hidden = True

                End Sub

                Private Function EPP_V2_LONG_DEGER( _
                ByVal hucre As Range) As Long

                If IsNumeric(hucre.Value2) Then
                    EPP_V2_LONG_DEGER = CLng(hucre.Value2)
                Else
                    EPP_V2_LONG_DEGER = 0
                End If

                End Function

                Private Sub EPP_V2_GORSEL_DUZEN( _
                ByVal ws As Worksheet, _
                ByVal veriSatiri As Long)

                Dim sonSatir As Long
                Dim TasarimSatir As Long

                sonSatir = _
                6 + veriSatiri

                If veriSatiri <= 0 Then Exit Sub

                ws.Range( _
                "I7:I" & sonSatir).NumberFormat = _
                "#,##0"

                ws.Range( _
                "K7:K" & sonSatir).NumberFormat = _
                "#,##0.00 ""TL"""

                ws.Range( _
                "L7:L" & sonSatir).NumberFormat = _
                "#,##0.00 ""TL"""

                On Error Resume Next

                If ws.AutoFilterMode Then
                ws.AutoFilterMode = False
                End If

                On Error GoTo 0

                ws.Range( _
                "A6:L" & sonSatir).AutoFilter

                If sonSatir < ws.Rows.Count Then
                    ws.Range("A" & sonSatir + 1 & ":L" & AKSORA_V13_SON(ws)).ClearContents
                    ws.Range("A" & sonSatir + 1 & ":L" & AKSORA_V13_SON(ws)).ClearFormats
                End If

                With ws.Range("A6:L6")
                    .Interior.Color = RGB(31, 78, 121)
                    .Font.Color = RGB(255, 255, 255)
                    .Font.bold = True
                    .HorizontalAlignment = xlCenter
                    .VerticalAlignment = xlCenter
                    .Borders.LineStyle = xlNone
                End With

                For TasarimSatir = 7 To sonSatir

                    With ws.Range("A" & TasarimSatir & ":L" & TasarimSatir)

                        If (TasarimSatir - 7) Mod 2 = 0 Then
                            .Interior.Color = RGB(242, 248, 252)
                        Else
                            .Interior.Color = RGB(255, 255, 255)
                        End If

                        .Font.Color = RGB(0, 0, 0)
                        .Font.bold = False
                        .Borders.LineStyle = xlNone

                    End With

                Next TasarimSatir

                ws.Range("A6:L" & sonSatir).Borders.LineStyle = xlNone

                End Sub

Private Function EPP_UY() As String
    EPP_UY = AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "EPP")
End Function


