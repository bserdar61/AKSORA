Attribute VB_Name = "modYUKLENICI"
Option Explicit

    '==========================================================
    ' AKSORA
    ' modYUKLENICI
    '
    ' YÜKLENÝCÝ STOK V2
    '
    ' EDAÞ = 7155
    ' AKSADA = 7055
    '
    ' DEPOLAR:
    ' BO13
    ' M122
    ' M123
    ' M124
    '==========================================================

    Private Const YUK_STOK_SAYFA As String = "YÜKLENÝCÝ_STOK"
    Private Const YUK_RAPOR_SAYFA As String = "YÜKLENÝCÝ_RAPOR"

  

    '==========================================================
    ' ANA GÜNCELLEME
    '==========================================================
    Public Sub YUKLENICI_STOK_GUNCELLE_V2()

    Dim wsStok As Worksheet
    Dim wsRapor As Worksheet

    Dim Veri7155 As Variant
    Dim Veri7055 As Variant
    Dim BirlesikVeri As Variant
    Dim EskiVeri As Variant

    Dim Satir7155 As Long
    Dim Sutun7155 As Long

    Dim Satir7055 As Long
    Dim Sutun7055 As Long

    Dim ToplamSatir As Long

    Dim EskiSonSatir As Long
    Dim EskiSonSutun As Long

    Dim baslangic As Double
    Dim ToplamSure As Double

    Dim yazmaBasladi As Boolean
    Dim YedekVar As Boolean

    Dim hataNo As Long
    Dim hataAciklama As String
    
    If Not AKSORA_SAP_GUNCELLEME_IZNI() Then Exit Sub

    On Error GoTo Hata

    Set wsStok = ThisWorkbook.Worksheets(YUK_STOK_SAYFA)
    Set wsRapor = ThisWorkbook.Worksheets(YUK_RAPOR_SAYFA)

    CORE_ISLEM_BASLAT "Yüklenici stoklarý güncelleniyor"

    baslangic = Timer

    '==========================================================
    ' 1. EDAÞ / 7155
    '==========================================================

    CORE_DURUM "[1/4] 7155 yüklenici stoklarý alýnýyor..."

    Dim CDPSure As Double, EPPSure As Double
    Dim BasariliCDP As Boolean, BasariliEPP As Boolean
    BasariliCDP = YUK_CDP_BLOK_AL(Veri7155, Satir7155, _
        Sutun7155, CDPSure, hataAciklama)
    If Not BasariliCDP Then
        hataAciklama = vbNullString
        BasariliCDP = YUK_V2_SAP_EXPORT_AL(YUK_EDAS_UY, _
            Veri7155, Satir7155, Sutun7155, hataAciklama)
    End If
    If Not BasariliCDP Then

    Err.Raise _
        vbObjectError + 5201, _
        "YUKLENICI_STOK_GUNCELLE_V2", _
        "7155 yüklenici stoklarý alýnamadý." & _
        vbCrLf & hataAciklama

    End If

    If Not IsArray(Veri7155) Then

    Err.Raise _
        vbObjectError + 5202, _
        "YUKLENICI_STOK_GUNCELLE_V2", _
        "7155 verisi geçerli bir dizi deðil."

    End If

    If Satir7155 <= 0 Then

    Err.Raise _
        vbObjectError + 5203, _
        "YUKLENICI_STOK_GUNCELLE_V2", _
        "7155 için yüklenici stok satýrý bulunamadý."

    End If

    '==========================================================
    ' 2. AKSADA / 7055
    '==========================================================

    CORE_DURUM "[2/4] 7055 yüklenici stoklarý alýnýyor..."

    hataAciklama = vbNullString

    BasariliEPP = YUK_EPP_BLOK_AL(Veri7055, Satir7055, _
        Sutun7055, EPPSure, hataAciklama)
    If Not BasariliEPP Then
        hataAciklama = vbNullString
        BasariliEPP = YUK_V2_EPP_ALV_AL(Veri7055, Satir7055, _
            Sutun7055, hataAciklama)
    End If
    If Not BasariliEPP Then

    Err.Raise _
        vbObjectError + 5211, _
        "YUKLENICI_STOK_GUNCELLE_V2", _
        "7055 yüklenici stoklarý alýnamadý." & _
        vbCrLf & hataAciklama

    End If

    If Not IsArray(Veri7055) Then

    Err.Raise _
        vbObjectError + 5212, _
        "YUKLENICI_STOK_GUNCELLE_V2", _
        "7055 verisi geçerli bir dizi deðil."

    End If

    If Satir7055 <= 0 Then

    Err.Raise _
        vbObjectError + 5213, _
        "YUKLENICI_STOK_GUNCELLE_V2", _
        "7055 için yüklenici stok satýrý bulunamadý."

    End If

    '==========================================================
    ' 3. VERÝLERÝ DOÐRULA
    '==========================================================

    YUK_V2_VERI_DOGRULA _
    Veri7155, _
    Satir7155, _
    Sutun7155, _
    YUK_EDAS_UY

    YUK_V2_VERI_DOGRULA _
    Veri7055, _
    Satir7055, _
    Sutun7055, _
    YUK_AKSADA_UY

    '==========================================================
    ' 4. MEVCUT STOK VERÝSÝNÝ YEDEKLE
    '==========================================================

    CORE_DURUM "[3/4] Mevcut yüklenici stoklarý yedekleniyor..."

    EskiSonSatir = YUK_V2_SON_SATIR(wsStok)
    EskiSonSutun = YUK_V2_SON_SUTUN(wsStok)

    If EskiSonSatir >= 7 And EskiSonSutun >= 1 Then

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

    End If

    '==========================================================
    ' 5. RAM'DE BÝRLEÞTÝR
    '==========================================================

    BirlesikVeri = _
    YUK_V2_BIRLESTIR( _
    Veri7155, _
    Satir7155, _
    Sutun7155, _
    Veri7055, _
    Satir7055, _
    Sutun7055)

    ' I = miktar, K = deger, L = birim fiyat.
    Dim fiyatSatiri As Long, miktar As Double
    For fiyatSatiri = 1 To Satir7155 + Satir7055
        miktar = YUK_V2_EPP_SAYI(BirlesikVeri(fiyatSatiri, 9))
        If miktar <> 0 Then
            BirlesikVeri(fiyatSatiri, 12) = _
                YUK_V2_EPP_SAYI(BirlesikVeri(fiyatSatiri, 11)) / miktar
        Else
            BirlesikVeri(fiyatSatiri, 12) = 0#
        End If
    Next fiyatSatiri

    ToplamSatir = Satir7155 + Satir7055

    If ToplamSatir <= 0 Then

    Err.Raise _
        vbObjectError + 5220, _
        "YUKLENICI_STOK_GUNCELLE_V2", _
        "Birleþtirilecek yüklenici stok verisi bulunamadý."

    End If

    '==========================================================
    ' 6. SAYFAYA YAZ
    '==========================================================

    CORE_DURUM "[4/4] Yüklenici stok ve raporlarý hazýrlanýyor..."

    yazmaBasladi = True

    YUK_V2_BASLIKLARI_YAZ wsStok

    wsStok.Range( _
    "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

    wsStok.Range("A7").Resize( _
    ToplamSatir, _
    12).Value2 = BirlesikVeri

    YUK_V2_STOK_OZET _
    wsStok, _
    ToplamSatir

    YUK_V2_RAPOR_OLUSTUR _
    wsStok, _
    wsRapor, _
    ToplamSatir

    YUK_V2_DEPO_LISTESI _
    wsStok

    YUK_V2_STOK_GORSEL _
    wsStok, _
    ToplamSatir

    STOK_GORUNUM_TEK "YÜKLENÝCÝ_STOK"
    RAPOR_GORUNUM_TEK "YÜKLENÝCÝ_RAPOR"

    ToplamSure = CORE_ISLEM_BITIR()

    If Not ANA_SESSIZ_MOD Then

    YUK_V2_STOK_ALT_BOSLUK_TEMIZLE

MsgBox _
    "Yüklenici stok güncellemesi baþarýyla tamamlandý." & _
    vbCrLf & vbCrLf & _
    "7155 : " & AKSORA_SAYI_METNI(Satir7155, "#,##0") & " satýr" & vbCrLf & _
    "7055 : " & AKSORA_SAYI_METNI(Satir7055, "#,##0") & " satýr" & vbCrLf & _
    "Toplam: " & AKSORA_SAYI_METNI(ToplamSatir, "#,##0") & " satýr" & vbCrLf & vbCrLf & _
        "Depolar: " & Join(AKSORA_IL_YUKLENICI(AKSORA_AKTIF_IL()), " / ") & vbCrLf & _
    "Toplam Süre: " & AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn", _
    vbInformation, _
    "AKSORA - YÜKLENÝCÝ"

    End If

    Exit Sub

Hata:

    hataNo = Err.Number
    hataAciklama = Err.Description

    On Error Resume Next

    If yazmaBasladi Then

    wsStok.Range( _
        "A7:L" & AKSORA_V13_SON(wsStok)).ClearContents

    If YedekVar Then

        If IsArray(EskiVeri) Then

            wsStok.Range("A7").Resize( _
                UBound(EskiVeri, 1), _
                UBound(EskiVeri, 2)).Value2 = EskiVeri

        End If

    End If

    YUK_V2_STOK_OZET _
        wsStok, _
        YUK_V2_VERI_SATIRI_SAYISI(wsStok)

    YUK_V2_RAPOR_OLUSTUR _
        wsStok, _
        wsRapor, _
        YUK_V2_VERI_SATIRI_SAYISI(wsStok)

    YUK_V2_DEPO_LISTESI wsStok

    End If

    ToplamSure = CORE_ISLEM_BITIR()

    On Error GoTo 0

    If Not ANA_SESSIZ_MOD Then

    MsgBox _
    "Yüklenici stok güncellemesi sýrasýnda hata oluþtu." & _
    vbCrLf & vbCrLf & _
    "Hata No : " & hataNo & vbCrLf & _
    "Açýklama: " & hataAciklama, _
    vbCritical, _
    "AKSORA - YÜKLENÝCÝ"

    End If

If ANA_SESSIZ_MOD Then
    Err.Raise hataNo, "YUKLENICI_STOK_GUNCELLE_V2", hataAciklama
End If

    End Sub

    '==========================================================
    ' SAP + MB52 + ÇOKLU DEPO + EXPORT
    '==========================================================
    Private Function YUK_V2_SAP_EXPORT_AL( _
    ByVal UretimYeri As String, _
    ByRef SonucVeri As Variant, _
    ByRef satirSayisi As Long, _
    ByRef SutunSayisi As Long, _
    ByRef hataMesaji As String) As Boolean

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim session As Object
    Dim alv As Object
    Dim wbAna As Workbook
    Dim wbExport As Workbook
    Dim wb As Workbook
    Dim wsKaynak As Worksheet
    Dim EskiWB As Object
    Dim sonSatir As Long
    Dim sonSutun As Long
    Dim i As Long
    Dim j As Long
    Dim HamVeri As Variant
    Dim TemizVeri() As Variant
    Dim baslangic As Double
    Dim GecenSure As Double
    Dim PencereVar As Boolean

    On Error GoTo Hata

    YUK_V2_SAP_EXPORT_AL = False
    satirSayisi = 0
    SutunSayisi = 0
    hataMesaji = vbNullString

    Set wbAna = ThisWorkbook
    If UretimYeri = YUK_EDAS_UY Then
        Set session = SAP_SESSION_AL("CDP", hataMesaji)
    ElseIf UretimYeri = YUK_AKSADA_UY Then
        Set session = SAP_SESSION_AL("EPP", hataMesaji)
    Else
        Err.Raise vbObjectError + 5304, "YUK_V2_SAP_EXPORT_AL", _
            "Geçersiz üretim yeri: " & UretimYeri
    End If

    If session Is Nothing Then
        Err.Raise vbObjectError + 5301, "YUK_V2_SAP_EXPORT_AL", _
            UretimYeri & " için uygun SAP oturumu bulunamadý."
    End If

    Set EskiWB = CreateObject("Scripting.Dictionary")
    EskiWB.CompareMode = vbTextCompare

    For Each wb In application.Workbooks
        EskiWB(wb.Name) = True
    Next wb

    session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nMB52"
    session.FindById("wnd[0]").SendVKey 0

    session.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = UretimYeri
    session.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
    session.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString

    YUK_V2_SAP_DEPOLARI_SEC session

    session.FindById("wnd[0]/tbar[1]/btn[8]").Press

    baslangic = Timer
    Do
        DoEvents
        Set alv = Nothing
        On Error Resume Next
        Set alv = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        On Error GoTo Hata
        If Not alv Is Nothing Then Exit Do
        If Timer >= baslangic Then
            GecenSure = Timer - baslangic
        Else
            GecenSure = (86400# - baslangic) + Timer
        End If
        If GecenSure >= 30 Then Exit Do
    Loop

    If alv Is Nothing Then
        Err.Raise vbObjectError + 5302, "YUK_V2_SAP_EXPORT_AL", _
            UretimYeri & " MB52 sonuç ekraný hazýr olmadý."
    End If

    If UretimYeri = YUK_EDAS_UY Then

        session.FindById("wnd[0]").SendVKey 43

    ElseIf UretimYeri = YUK_AKSADA_UY Then

        alv.SetCurrentCell -1, "MAKTX"
        alv.SelectColumn "MAKTX"
        alv.ContextMenu
        alv.SelectContextMenuItem "&XXL"

        baslangic = Timer
        Do
            DoEvents
            PencereVar = False
            On Error Resume Next
            PencereVar = Not session.FindById("wnd[1]") Is Nothing
            On Error GoTo Hata
            If PencereVar Then Exit Do
            If Timer >= baslangic Then
                GecenSure = Timer - baslangic
            Else
                GecenSure = (86400# - baslangic) + Timer
            End If
            If GecenSure >= 10 Then Exit Do
        Loop

        If Not PencereVar Then
            Err.Raise vbObjectError + 5303, "YUK_V2_SAP_EXPORT_AL", _
                "7055 Excel export penceresi açýlmadý."
        End If

        On Error Resume Next
        Err.Clear
        session.FindById("wnd[1]/tbar[0]/btn[20]").Press
        If Err.Number <> 0 Then
            Err.Clear
            session.FindById("wnd[1]/tbar[0]/btn[0]").Press
        End If
        Err.Clear
        On Error GoTo Hata

    Else

        Err.Raise vbObjectError + 5304, "YUK_V2_SAP_EXPORT_AL", _
            "Geçersiz üretim yeri: " & UretimYeri

    End If

    baslangic = Timer
    Do
        DoEvents
        Set wbExport = Nothing

        For Each wb In application.Workbooks
            If Not wb Is wbAna Then
                If Not EskiWB.Exists(wb.Name) Then
                    Set wbExport = wb
                    Exit For
                End If
            End If
        Next wb

        If Not wbExport Is Nothing Then Exit Do

        If Timer >= baslangic Then
            GecenSure = Timer - baslangic
        Else
            GecenSure = (86400# - baslangic) + Timer
        End If

        If GecenSure >= 45 Then Exit Do
    Loop

    If wbExport Is Nothing Then
        Err.Raise vbObjectError + 5305, "YUK_V2_SAP_EXPORT_AL", _
            UretimYeri & " SAP export Excel dosyasý bulunamadý."
    End If

    Set wsKaynak = YUK_V2_EXPORT_SAYFASI_BUL(wbExport)

    If wsKaynak Is Nothing Then
        Err.Raise vbObjectError + 5306, "YUK_V2_SAP_EXPORT_AL", _
            "SAP export çalýþma sayfasý bulunamadý."
    End If

    sonSatir = YUK_V2_SON_SATIR(wsKaynak)
    sonSutun = YUK_V2_SON_SUTUN(wsKaynak)

    If sonSatir < 2 Then
        Err.Raise vbObjectError + 5307, "YUK_V2_SAP_EXPORT_AL", _
            "SAP export dosyasýnda veri satýrý bulunamadý."
    End If

    If sonSutun < 11 Then
        Err.Raise vbObjectError + 5308, "YUK_V2_SAP_EXPORT_AL", _
            "SAP export dosyasýndaki sütun sayýsý yetersiz."
    End If

    HamVeri = wsKaynak.Range( _
        wsKaynak.Cells(2, 1), _
        wsKaynak.Cells(sonSatir, sonSutun)).Value2

    satirSayisi = UBound(HamVeri, 1)
    SutunSayisi = UBound(HamVeri, 2)

    ReDim TemizVeri(1 To satirSayisi, 1 To SutunSayisi)

    For i = 1 To satirSayisi
        For j = 1 To SutunSayisi
            TemizVeri(i, j) = HamVeri(i, j)
        Next j
    Next i

    SonucVeri = TemizVeri
    YUK_V2_SAP_EXPORT_AL = True

    On Error Resume Next
    wbExport.Close SaveChanges:=False
    wbAna.Activate
    On Error GoTo 0

    Exit Function

Hata:

    Dim hataNo As Long
    Dim hataAciklama As String

    hataNo = Err.Number
    hataAciklama = Err.Description

    hataMesaji = _
        "YUK_V2_SAP_EXPORT_AL hatasý." & vbCrLf & _
        "Hata No: " & hataNo & vbCrLf & _
        "Açýklama: " & hataAciklama

    YUK_V2_SAP_EXPORT_AL = False

    On Error Resume Next
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    wbAna.Activate
    On Error GoTo 0

    End Function

    '==========================================================
    ' DOÐRU SAP SESSION BUL
    '
    ' 7155 -> CDP
    ' 7055 -> EPP
    '==========================================================

    Private Sub YUK_V2_SAP_DEPOLARI_SEC(ByVal session As Object)
    Dim depolar As Variant
    Dim tablo As Object
    Dim tabloID As String, hucreID As String
    Dim i As Long, ekranSatiri As Long, konum As Long

    depolar = AKSORA_IL_YUKLENICI(AKSORA_AKTIF_IL())

    tabloID = _
        "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/" & _
        "ssubSCREEN_HEADER:SAPLALDB:3010/" & _
        "tblSAPLALDBSINGLE"

    hucreID = tabloID & "/ctxtRSCSEL_255-SLOW_I"

    session.FindById( _
        "wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press

    DoEvents

    Set tablo = session.FindById(tabloID)
    tablo.VerticalScrollbar.Position = 0

    For i = LBound(depolar) To UBound(depolar)
        Set tablo = session.FindById(tabloID)
        konum = CLng(tablo.VerticalScrollbar.Position)
        ekranSatiri = i - LBound(depolar) - konum

        If ekranSatiri < 0 Or _
           ekranSatiri >= CLng(tablo.VisibleRowCount) Then

            tablo.VerticalScrollbar.Position = i - LBound(depolar)
            DoEvents

            Set tablo = session.FindById(tabloID)
            konum = CLng(tablo.VerticalScrollbar.Position)
            ekranSatiri = i - LBound(depolar) - konum
        End If

        If ekranSatiri < 0 Or _
           ekranSatiri >= CLng(tablo.VisibleRowCount) Then

            Err.Raise vbObjectError + 780, _
                      "YUK_V2_SAP_DEPOLARI_SEC", _
                      "SAP depo seçiminde satýra ulaþýlamadý: " & _
                      CStr(depolar(i))
        End If

        session.FindById( _
            hucreID & "[1," & ekranSatiri & "]").Text = _
            CStr(depolar(i))
    Next i

    session.FindById("wnd[1]/tbar[0]/btn[8]").Press
    DoEvents
End Sub
    '==========================================================
    ' EXPORT SAYFASINI BUL
    '==========================================================
    Private Function YUK_V2_EXPORT_SAYFASI_BUL( _
    ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet
    Dim sonSatir As Long
    Dim sonSutun As Long

    On Error Resume Next

    Set ws = wb.Worksheets("RawData")

    On Error GoTo 0

    If Not ws Is Nothing Then

    Set YUK_V2_EXPORT_SAYFASI_BUL = ws
    Exit Function

    End If

    For Each ws In wb.Worksheets

    sonSatir = YUK_V2_SON_SATIR(ws)
    sonSutun = YUK_V2_SON_SUTUN(ws)

    If sonSatir >= 2 And sonSutun >= 11 Then

        Set YUK_V2_EXPORT_SAYFASI_BUL = ws
        Exit Function

    End If

    Next ws

    End Function

    '==========================================================
    ' VERÝ DOÐRULA
    '==========================================================
    Private Sub YUK_V2_VERI_DOGRULA( _
    ByRef veri As Variant, _
    ByVal satirSayisi As Long, _
    ByVal SutunSayisi As Long, _
    ByVal BeklenenUY As String)

    Dim i As Long

    Dim uy As String
    Dim depo As String

    If SutunSayisi < 11 Then

    Err.Raise _
        vbObjectError + 5401, _
        "YUK_V2_VERI_DOGRULA", _
        BeklenenUY & _
        " verisindeki sütun sayýsý yetersiz."

    End If

    For i = 1 To satirSayisi

    uy = Trim$(CStr(veri(i, 1)))
    depo = UCase$(Trim$(CStr(veri(i, 2))))

    If Len(uy) > 0 Then

        If uy <> BeklenenUY Then

            Err.Raise _
                vbObjectError + 5402, _
                "YUK_V2_VERI_DOGRULA", _
                "Beklenmeyen üretim yeri: " & uy

        End If

    End If

    If Len(depo) > 0 Then

        If Not YUK_V2_DEPO_GECERLI(depo) Then

            Err.Raise _
                vbObjectError + 5403, _
                "YUK_V2_VERI_DOGRULA", _
                "Beklenmeyen yüklenici deposu: " & depo

        End If

    End If

    Next i

    End Sub

    '==========================================================
    ' GEÇERLÝ DEPO
    '==========================================================
    Private Function YUK_V2_DEPO_GECERLI(ByVal depo As String) As Boolean
    YUK_V2_DEPO_GECERLI = AKSORA_IL_DEPO_GECERLI( _
        depo, "YUKLENICI", AKSORA_AKTIF_IL())
End Function

    '==========================================================
    ' 7155 + 7055 BÝRLEÞTÝR
    '==========================================================
    Private Function YUK_V2_BIRLESTIR( _
    ByRef Veri7155 As Variant, _
    ByVal Satir7155 As Long, _
    ByVal Sutun7155 As Long, _
    ByRef Veri7055 As Variant, _
    ByVal Satir7055 As Long, _
    ByVal Sutun7055 As Long) As Variant

    Dim sonuc() As Variant

    Dim i As Long
    Dim j As Long
    Dim hedef As Long

    ReDim sonuc( _
    1 To Satir7155 + Satir7055, _
    1 To 12)

    hedef = 1

    For i = 1 To Satir7155

    For j = 1 To 12

        If j <= Sutun7155 Then
            sonuc(hedef, j) = Veri7155(i, j)
        Else
            sonuc(hedef, j) = vbNullString
        End If

    Next j

    hedef = hedef + 1

    Next i

    For i = 1 To Satir7055

    For j = 1 To 12

        If j <= Sutun7055 Then
            sonuc(hedef, j) = Veri7055(i, j)
        Else
            sonuc(hedef, j) = vbNullString
        End If

    Next j

    hedef = hedef + 1

    Next i

    YUK_V2_BIRLESTIR = sonuc

    End Function

    '==========================================================
    ' BAÞLIKLAR
    '==========================================================
    Private Sub YUK_V2_BASLIKLARI_YAZ( _
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

    '==========================================================
    ' STOK ÖZET
    '==========================================================
    Private Sub YUK_V2_STOK_OZET( _
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

    ' KPI deðerleri AutoFilter sonucuna göre otomatik deðiþir.
    ws.Range("C3").Formula = _
        "=SUBTOTAL(103,A7:A" & sonSatir & ")"

    ws.Range("C4").Formula = _
        "=SUBTOTAL(109,K7:K" & sonSatir & ")"

    ws.Range("C2").NumberFormat = _
    "dd.mm.yyyy hh:mm:ss"

    ws.Range("C3").NumberFormat = _
    "#,##0"

    ws.Range("C4").NumberFormat = _
    "#,##0.00 ""TL"""

    End Sub

    '==========================================================
    ' YÜKLENÝCÝ RAPOR
    '
    ' A = Depo
    ' B = Depo tanýmý
    ' C = 7155
    ' D = 7055
    ' E = Toplam
    '==========================================================
    Public Sub YUK_V2_RAPOR_OLUSTUR( _
    ByVal wsStok As Worksheet, _
    ByVal wsRapor As Worksheet, _
    ByVal veriSatiri As Long)

    Dim dict As Object

    Dim veri As Variant
    Dim bilgi As Variant

    Dim i As Long
    Dim HedefSatir As Long
    Dim TasarimSatir As Long

    Dim uy As String
    Dim depo As String
    Dim DepoTanimi As String

    Dim deger As Double

    Dim Toplam7155 As Double
    Dim Toplam7055 As Double
    Dim GenelToplam As Double

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    If veriSatiri > 0 Then

    veri = _
        wsStok.Range( _
            "A7:K" & 6 + veriSatiri).Value2

    For i = 1 To UBound(veri, 1)

        uy = Trim$(CStr(veri(i, 1)))
        depo = UCase$(Trim$(CStr(veri(i, 2))))
        DepoTanimi = CStr(veri(i, 3))

        deger = 0

        If IsNumeric(veri(i, 11)) Then
            deger = CDbl(veri(i, 11))
        End If

        If YUK_V2_DEPO_GECERLI(depo) Then

            If Not dict.Exists(depo) Then

                dict.Add _
                    depo, _
                    Array( _
                        DepoTanimi, _
                        0#, _
                        0#)

            End If

            bilgi = dict(depo)

            If Len(Trim$(CStr(bilgi(0)))) = 0 Then
                bilgi(0) = DepoTanimi
            End If

            If uy = YUK_EDAS_UY Then

                bilgi(1) = _
                    CDbl(bilgi(1)) + deger

                Toplam7155 = _
                    Toplam7155 + deger

            ElseIf uy = YUK_AKSADA_UY Then

                bilgi(2) = _
                    CDbl(bilgi(2)) + deger

                Toplam7055 = _
                    Toplam7055 + deger

            End If

            dict(depo) = bilgi

        End If

    Next i

    End If

    On Error Resume Next

    If wsRapor.AutoFilterMode Then
    wsRapor.AutoFilterMode = False
    End If

    On Error GoTo 0

    wsRapor.Range( _
    "A8:E" & AKSORA_V13_SON(wsRapor)).ClearContents

    wsRapor.Range("A8").Value = "Depo yeri"
    wsRapor.Range("B8").Value = "Depo yeri tanýmý"
        wsRapor.Range("C8").Value = YUK_EDAS_UY() & " Toplam Stok Deðeri"
    wsRapor.Range("D8").Value = YUK_AKSADA_UY() & " Toplam Stok Deðeri"
    wsRapor.Range("E8").Value = "Toplam Stok Deðeri"

    HedefSatir = 9

    YUK_V2_IL_RAPOR_SATIRLARI_YAZ wsRapor, dict, HedefSatir

    GenelToplam = Toplam7155 + Toplam7055

    wsRapor.Cells(HedefSatir, 1).Value = "TOPLAM"
    wsRapor.Cells(HedefSatir, 3).Value = Toplam7155
    wsRapor.Cells(HedefSatir, 4).Value = Toplam7055
    wsRapor.Cells(HedefSatir, 5).Value = GenelToplam

    wsRapor.Range( _
    "A" & HedefSatir & _
    ":E" & HedefSatir).Font.bold = True

    wsRapor.Range("B3").Value = Now
    wsRapor.Range("B4").Value = dict.Count
    wsRapor.Range("B5").Value = GenelToplam

    wsRapor.Range("B3").NumberFormat = "dd.mm.yyyy hh:mm:ss"
    wsRapor.Range("B4").NumberFormat = "#,##0"
    wsRapor.Range("B5").NumberFormat = "#,##0.00 ""TL"""

    wsRapor.Range( _
    "C9:E" & HedefSatir).NumberFormat = _
    "#,##0.00 ""TL"""

    If HedefSatir > 9 Then

    wsRapor.Range( _
        "A8:E" & HedefSatir - 1).AutoFilter

    Else

    wsRapor.Range("A8:E8").AutoFilter

    End If

    '==========================================================
    ' YÜKLENÝCÝ_RAPOR NÝHAÝ TABLO TASARIMI
    ' CDP / EPP / ÝLÇE ile ayný standart
    '==========================================================

    With wsRapor.Range("A8:E8")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    For TasarimSatir = 9 To HedefSatir - 1

        With wsRapor.Range( _
            "A" & TasarimSatir & _
            ":E" & TasarimSatir)

            If (TasarimSatir - 9) Mod 2 = 0 Then
                .Interior.Color = RGB(242, 248, 252)
            Else
                .Interior.Color = RGB(255, 255, 255)
            End If

            .Font.Color = RGB(0, 0, 0)
            .Font.bold = False
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin

        End With

    Next TasarimSatir

    With wsRapor.Range( _
        "A" & HedefSatir & _
        ":E" & HedefSatir)

        .Interior.Color = RGB(221, 235, 247)
        .Font.Color = RGB(0, 0, 0)
        .Font.bold = True
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin

    End With

    wsRapor.Range( _
        "A" & HedefSatir + 1 & _
        ":E" & AKSORA_V13_SON(wsRapor)).ClearFormats

    wsRapor.Columns("A").ColumnWidth = 12
    wsRapor.Columns("B").ColumnWidth = 28
    wsRapor.Columns("C").ColumnWidth = 20
    wsRapor.Columns("D").ColumnWidth = 20
    wsRapor.Columns("E").ColumnWidth = 20

    wsRapor.Rows("8:" & HedefSatir).RowHeight = 20
    wsRapor.Range("A9:B" & HedefSatir).HorizontalAlignment = xlLeft
    wsRapor.Range("C9:E" & HedefSatir).HorizontalAlignment = xlRight
    wsRapor.Range("A8:E" & HedefSatir).VerticalAlignment = xlCenter
    wsRapor.Range("A9:E" & HedefSatir).Borders.LineStyle = xlNone

    End Sub

    '==========================================================
    ' RAPOR SATIRI
    '==========================================================
    Private Sub YUK_V2_RAPOR_SATIR_YAZ( _
    ByVal ws As Worksheet, _
    ByVal dict As Object, _
    ByVal depo As String, _
    ByRef HedefSatir As Long)

    Dim bilgi As Variant

    If Not dict.Exists(depo) Then Exit Sub

    bilgi = dict(depo)

    ws.Cells(HedefSatir, 1).Value = depo
    ws.Cells(HedefSatir, 2).Value = bilgi(0)
    ws.Cells(HedefSatir, 3).Value = CDbl(bilgi(1))
    ws.Cells(HedefSatir, 4).Value = CDbl(bilgi(2))

    ws.Cells(HedefSatir, 5).Value = _
    CDbl(bilgi(1)) + CDbl(bilgi(2))

    HedefSatir = HedefSatir + 1

    End Sub

    '==========================================================
    ' H2 DEPO LÝSTESÝ
    '==========================================================
Private Sub YUK_V2_DEPO_LISTESI(ByVal ws As Worksheet)
    AKFIX_STOK_LISTE ws
End Sub

    '==========================================================
    ' STOK GÖRSEL
    '==========================================================
    Private Sub YUK_V2_STOK_GORSEL( _
    ByVal ws As Worksheet, _
    ByVal veriSatiri As Long)

    Dim sonSatir As Long
    Dim TasarimSatir As Long

    If veriSatiri <= 0 Then Exit Sub

    sonSatir = 6 + veriSatiri

    ws.Range( _
    "I7:I" & sonSatir).NumberFormat = _
    "#,##0"

    ws.Range( _
    "K7:L" & sonSatir).NumberFormat = _
    "#,##0.00 ""TL"""

    On Error Resume Next

    If ws.AutoFilterMode Then
    ws.AutoFilterMode = False
    End If

    On Error GoTo 0

    ws.Range( _
    "A6:L" & sonSatir).AutoFilter

    With ws.Range("A6:L6")
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = RGB(255, 255, 255)
        .Font.bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ws.Rows("6:" & sonSatir).RowHeight = 20
    ws.Range("A7:G" & sonSatir).HorizontalAlignment = xlLeft
    ws.Range("H7:H" & sonSatir).HorizontalAlignment = xlLeft
    ws.Range("I7:I" & sonSatir).HorizontalAlignment = xlRight
    ws.Range("J7:J" & sonSatir).HorizontalAlignment = xlCenter
    ws.Range("K7:K" & sonSatir).HorizontalAlignment = xlRight
    ws.Range("L7:L" & sonSatir).HorizontalAlignment = xlRight
    ws.Range("A7:L" & sonSatir).VerticalAlignment = xlCenter

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

    End Sub

    '==========================================================
    ' VERÝ SATIRI SAYISI
    '==========================================================
    Private Function YUK_V2_VERI_SATIRI_SAYISI( _
    ByVal ws As Worksheet) As Long

    Dim sonSatir As Long

    sonSatir = YUK_V2_SON_SATIR(ws)

    If sonSatir < 7 Then

    YUK_V2_VERI_SATIRI_SAYISI = 0

    Else

    YUK_V2_VERI_SATIRI_SAYISI = _
        sonSatir - 6

    End If

    End Function

    '==========================================================
    ' SON SATIR
    '==========================================================
    Private Function YUK_V2_SON_SATIR( _
    ByVal ws As Worksheet) As Long

    Dim sonHucre As Range

    On Error Resume Next

    Set sonHucre = _
    ws.Cells.Find( _
    What:="*", _
    after:=ws.Cells(1, 1), _
    LookAt:=xlPart, _
    LookIn:=xlFormulas, _
    SearchOrder:=xlByRows, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False)

    On Error GoTo 0

    If sonHucre Is Nothing Then

    YUK_V2_SON_SATIR = 0

    Else

    YUK_V2_SON_SATIR = sonHucre.Row

    End If

    End Function

    '==========================================================
    ' SON SÜTUN
    '==========================================================
    Private Function YUK_V2_SON_SUTUN( _
    ByVal ws As Worksheet) As Long

    Dim sonHucre As Range

    On Error Resume Next

    Set sonHucre = _
    ws.Cells.Find( _
    What:="*", _
    after:=ws.Cells(1, 1), _
    LookAt:=xlPart, _
    LookIn:=xlFormulas, _
    SearchOrder:=xlByColumns, _
    SearchDirection:=xlPrevious, _
    MatchCase:=False)

    On Error GoTo 0

    If sonHucre Is Nothing Then

    YUK_V2_SON_SUTUN = 0

    Else

    YUK_V2_SON_SUTUN = sonHucre.Column

    End If

    End Function

Private Function YUK_V2_EPP_ALV_AL( _
ByRef SonucVeri As Variant, _
ByRef satirSayisi As Long, _
ByRef SutunSayisi As Long, _
ByRef hataMesaji As String) As Boolean

Dim SapGuiAuto As Object, SAPApp As Object, session As Object, alv As Object
Dim KolonSirasi As Object
Dim KolonID(1 To 12) As String
Dim ham() As Variant, temiz() As Variant
Dim r As Long, c As Long, n As Long, toplam As Long
Dim uy As String, depo As String, malzeme As String
Dim baslangic As Double, gecen As Double

On Error GoTo Hata
YUK_V2_EPP_ALV_AL = False
satirSayisi = 0
SutunSayisi = 0
hataMesaji = vbNullString

Set session = SAP_SESSION_AL("EPP", hataMesaji)

If session Is Nothing Then
    hataMesaji = "7055 için uygun SAP oturumu bulunamadý."
    Exit Function
End If

session.FindById("wnd[0]/tbar[0]/okcd").Text = "/nMB52"
session.FindById("wnd[0]").SendVKey 0
YUK_V2_EPP_BEKLE 0.3

session.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = YUK_AKSADA_UY
session.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
session.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString

YUK_V2_SAP_DEPOLARI_SEC session
session.FindById("wnd[0]/tbar[1]/btn[8]").Press

baslangic = Timer
Do
    DoEvents
    Set alv = Nothing
    On Error Resume Next
    Set alv = session.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    On Error GoTo Hata
    If Not alv Is Nothing Then Exit Do
    gecen = YUK_V2_EPP_GECEN(baslangic)
    If gecen >= 30 Then Exit Do
Loop

If alv Is Nothing Then
    hataMesaji = "7055 MB52 sonuç ekraný hazýr olmadý."
    Exit Function
End If

On Error Resume Next
alv.PressToolbarButton "&DELETE_FILTER"
Err.Clear
On Error GoTo Hata
YUK_V2_EPP_BEKLE 0.3

Set KolonSirasi = alv.ColumnOrder
If KolonSirasi Is Nothing Then
    hataMesaji = "7055 MB52 sütun sýrasý alýnamadý."
    Exit Function
End If
If CLng(KolonSirasi.Count) < 12 Then
    hataMesaji = "7055 PANEL düzeninde 12 sütun bulunamadý."
    Exit Function
End If

For c = 1 To 12
    KolonID(c) = CStr(KolonSirasi.Item(c - 1))
Next c

toplam = CLng(alv.RowCount)
If toplam <= 0 Then
    hataMesaji = "7055 yüklenici stok satýrý bulunamadý."
    Exit Function
End If

ReDim ham(1 To toplam, 1 To 12)

For r = 0 To toplam - 1
    On Error Resume Next
    alv.FirstVisibleRow = r
    alv.SetCurrentCell r, KolonID(7)
    On Error GoTo Hata
    DoEvents
    YUK_V2_EPP_BEKLE 0.015

    For c = 1 To 12
        ham(r + 1, c) = alv.GetCellValue(r, KolonID(c))
    Next c
Next r

ReDim temiz(1 To toplam, 1 To 12)

For r = 1 To toplam
    uy = Trim$(CStr(ham(r, 1)))
    depo = UCase$(Trim$(CStr(ham(r, 2))))
    malzeme = Trim$(CStr(ham(r, 7)))

    If uy = YUK_AKSADA_UY And Len(depo) > 0 And Len(malzeme) > 0 Then
        n = n + 1
        For c = 1 To 12
            temiz(n, c) = ham(r, c)
        Next c
        temiz(n, 9) = YUK_V2_EPP_SAYI(temiz(n, 9))
        temiz(n, 11) = YUK_V2_EPP_SAYI(temiz(n, 11))
        temiz(n, 12) = YUK_V2_EPP_SAYI(temiz(n, 12))
    End If
Next r

If n <= 0 Then
    hataMesaji = "7055 geçerli yüklenici stok satýrý bulunamadý."
    Exit Function
End If

Dim cikti() As Variant
ReDim cikti(1 To n, 1 To 12)
For r = 1 To n
    For c = 1 To 12
        cikti(r, c) = temiz(r, c)
    Next c
Next r

SonucVeri = cikti
satirSayisi = n
SutunSayisi = 12
YUK_V2_EPP_ALV_AL = True
Exit Function

Hata:
hataMesaji = "7055 yüklenici stoklarý SAP tablosundan alýnamadý." & vbCrLf & Err.Description
YUK_V2_EPP_ALV_AL = False
End Function

Private Function YUK_V2_EPP_SAYI(ByVal deger As Variant) As Double
Dim s As String, neg As Boolean
If IsNumeric(deger) Then YUK_V2_EPP_SAYI = CDbl(deger): Exit Function
s = Trim$(CStr(deger))
If Len(s) = 0 Or s = "-" Then Exit Function
If Left$(s, 1) = "-" Then neg = True: s = Mid$(s, 2)
If Right$(s, 1) = "-" Then neg = True: s = Left$(s, Len(s) - 1)
s = Replace(s, " ", "")
If InStr(s, ",") > 0 And InStr(s, ".") > 0 Then
    If InStrRev(s, ",") > InStrRev(s, ".") Then
        s = Replace(s, ".", "")
        s = Replace(s, ",", application.DecimalSeparator)
    Else
        s = Replace(s, ",", "")
        s = Replace(s, ".", application.DecimalSeparator)
    End If
ElseIf InStr(s, ",") > 0 Then
    s = Replace(s, ",", application.DecimalSeparator)
ElseIf InStr(s, ".") > 0 And application.DecimalSeparator <> "." Then
    s = Replace(s, ".", application.DecimalSeparator)
End If
If IsNumeric(s) Then
    YUK_V2_EPP_SAYI = CDbl(s)
    If neg Then YUK_V2_EPP_SAYI = -YUK_V2_EPP_SAYI
End If
End Function

Private Sub YUK_V2_EPP_BEKLE(ByVal saniye As Double)
Dim t As Double
t = Timer
Do While YUK_V2_EPP_GECEN(t) < saniye
    DoEvents
Loop
End Sub

Private Function YUK_V2_EPP_GECEN(ByVal baslangic As Double) As Double
If Timer >= baslangic Then
    YUK_V2_EPP_GECEN = Timer - baslangic
Else
    YUK_V2_EPP_GECEN = (86400# - baslangic) + Timer
End If
End Function

Private Sub YUK_V2_STOK_ALT_BOSLUK_TEMIZLE()
Dim ws As Worksheet
Dim sonSatir As Long

Set ws = ThisWorkbook.Worksheets("YÜKLENÝCÝ_STOK")
sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

If sonSatir >= 7 Then
    ws.Range("A7:L" & sonSatir).Borders.LineStyle = xlNone
    If sonSatir < ws.Rows.Count Then
        ws.Range("A" & sonSatir + 1 & ":L" & AKSORA_V13_SON(ws)).ClearFormats
    End If
End If
End Sub





' 30'lu SAP ALV okuma: yalnizca gecerli stok satirlari aktarilir.
Public Function YUK_CDP_BLOK_AL(ByRef veri As Variant, _
    ByRef satir As Long, ByRef sutun As Long, _
    ByRef sure As Double, ByRef Hata As String) As Boolean
    Dim blok As Long
    Dim ses As Object, alv As Object, duzen As Object, kolon As Object
    Dim ids(1 To 12) As String, ham() As Variant, temiz() As Variant
    Dim i As Long, j As Long, b As Long, son As Long, n As Long
    Dim secili As Long, tekrar As Long, stok As Long
    Dim ad As String, tanim As String, depo As String
    Dim sorun As Boolean, ozet As Boolean, bitti As Boolean
    Dim t As Double, neden As String, uy As String
    On Error GoTo Hata
    satir = 0: sutun = 0: sure = 0: Hata = vbNullString
        uy = YUK_EDAS_UY()
    t = Timer
    ' Hizli yol: ilin MB52 listesi tek sorguda metin olarak alinir (olmazsa asagidaki eski yontem).
    If AKHZ_TEK_AL(5, veri, satir, sutun, Hata) Then YUK_CDP_BLOK_AL = True: Exit Function
    Set ses = SAP_SESSION_AL("CDP", Hata)
    If ses Is Nothing Then Err.Raise vbObjectError + 870, , "SAP oturumu bulunamadi: " & Hata
    ses.StartTransaction "MB52"
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = uy
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString
    YUK_V2_SAP_DEPOLARI_SEC ses
    ses.FindById("wnd[0]").SendVKey 8
    YUK_CDP_BLOK_BEKLE 0.4
    ses.FindById("wnd[0]/tbar[1]/btn[33]").Press
    YUK_CDP_BLOK_BEKLE 0.3
    Set duzen = ses.FindById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    secili = -1
    For i = 0 To CLng(duzen.RowCount) - 1
        ad = Trim$(CStr(duzen.GetCellValue(i, "VARIANT")))
        tanim = Trim$(CStr(duzen.GetCellValue(i, "TEXT")))
        If StrComp(ad, "PANEL", vbTextCompare) = 0 Or _
           StrComp(tanim, "MLZ STOK PANEL", vbTextCompare) = 0 Then
            secili = i: Exit For
        End If
    Next i
    If secili < 0 Then Err.Raise vbObjectError + 871, , "PANEL duzeni bulunamadi."
    duzen.SetCurrentCell secili, "TEXT"
    duzen.SelectedRows = CStr(secili)
    duzen.ClickCurrentCell
    YUK_CDP_BLOK_BEKLE 0.3
    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    On Error Resume Next
    alv.PressToolbarButton "&DELETE_FILTER"
    Err.Clear
    On Error GoTo Hata
    Set kolon = alv.ColumnOrder
    If CLng(kolon.Count) < 12 Then Err.Raise vbObjectError + 872, , "PANEL 12 sutun icermiyor."
    For j = 1 To 12
        ids(j) = CStr(kolon.Item(j - 1))
    Next j
    n = CLng(alv.RowCount)
    If n <= 0 Then Err.Raise vbObjectError + 873, , "SAP sonucu bos."
    ReDim ham(1 To n, 1 To 12)
    ReDim temiz(1 To n, 1 To 12)
    blok = AKSORA_SAP_BLOK_BOYU(alv)
    For b = 0 To n - 1 Step blok
        If bitti Then Exit For
        son = b + blok - 1
        If son >= n Then son = n - 1
        On Error Resume Next
        Err.Clear
        alv.FirstVisibleRow = b
        sorun = (Err.Number <> 0)
        On Error GoTo Hata
        If sorun Then Err.Raise vbObjectError + 874, , "Blok kaydirilamadi: " & (b + 1)
        DoEvents
        YUK_CDP_BLOK_BEKLE 0.015
        AKSORA_TOPLU_SURE_TIK
        For i = b To son
            sorun = False: ozet = False
            On Error Resume Next
            For j = 1 To 12
                Err.Clear
                If j = 12 Then
                    ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                    ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                    If Trim$(CStr(ham(i + 1, 1))) = uy And _
                       Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                        ham(i + 1, j) = 0#
                    Else
                        ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                    End If
                Else
                    ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                End If
                If Err.Number <> 0 Then sorun = True: Exit For
            Next j
            On Error GoTo Hata
            If Not sorun Then
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
                ozet = YUK_CDP_OZET_MI(ham, i + 1, uy)
                If ozet Then bitti = True: Exit For
            End If
            For tekrar = 1 To 2
                sorun = False
                On Error Resume Next
                Err.Clear
                alv.FirstVisibleRow = i
                alv.SetCurrentCell i, ids(7)
                sorun = (Err.Number <> 0)
                On Error GoTo Hata
                If Not sorun Then
                    DoEvents
                    YUK_CDP_BLOK_BEKLE 0.015
                    On Error Resume Next
                    For j = 1 To 12
                        Err.Clear
                        If j = 12 Then
                            ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                            ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                            If Trim$(CStr(ham(i + 1, 1))) = uy And _
                               Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                                ham(i + 1, j) = 0#
                            Else
                                ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                            End If
                        Else
                            ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                        End If
                        If Err.Number <> 0 Then sorun = True: Exit For
                    Next j
                    On Error GoTo Hata
                    If Not sorun Then
                        If Trim$(CStr(ham(i + 1, 1))) = uy And _
                           Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then Exit For
                        If YUK_CDP_OZET_MI(ham, i + 1, uy) Then Exit For
                    End If
                End If
            Next tekrar
            If Not sorun Then
                If YUK_CDP_OZET_MI(ham, i + 1, uy) Then bitti = True: Exit For
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
            End If
            ' Ozet satiri SAP tarafindan secilemiyorsa son kisimda dur.
            If i >= n - 10 And stok > 0 And _
               Len(Trim$(CStr(ham(i + 1, 1)))) = 0 Then
                bitti = True: Exit For
            End If
            Err.Raise vbObjectError + 875, , "Satir dogrulanamadi: " & (i + 1)
KAYDET:
            depo = UCase$(Trim$(CStr(ham(i + 1, 2))))
            If Not YUK_V2_DEPO_GECERLI(depo) Then
                Err.Raise vbObjectError + 876, , "Beklenmeyen depo: " & depo
            End If
            stok = stok + 1
            For j = 1 To 12
                temiz(stok, j) = ham(i + 1, j)
            Next j
            temiz(stok, 9) = YUK_V2_EPP_SAYI(temiz(stok, 9))
            temiz(stok, 11) = YUK_V2_EPP_SAYI(temiz(stok, 11))
            If IsNumeric(temiz(stok, 9)) And CDbl(temiz(stok, 9)) <> 0 Then
                temiz(stok, 12) = CDbl(temiz(stok, 11)) / CDbl(temiz(stok, 9))
            Else
                temiz(stok, 12) = 0#
            End If
        Next i
    Next b
    If stok <= 0 Then Err.Raise vbObjectError + 877, , "Gecerli stok satiri yok."
    ReDim ham(1 To stok, 1 To 12)
    For i = 1 To stok
        For j = 1 To 12
            ham(i, j) = temiz(i, j)
        Next j
    Next i
    veri = ham: satir = stok: sutun = 12
    sure = YUK_CDP_BLOK_SURE(t)
    YUK_CDP_BLOK_AL = True
    Exit Function
Hata:
    neden = Err.Description
    Hata = "30'lu blok okuma: " & neden
    sure = YUK_CDP_BLOK_SURE(t)
    YUK_CDP_BLOK_AL = False
End Function

Private Function YUK_CDP_OZET_MI(ByRef ham() As Variant, _
    ByVal r As Long, ByVal uy As String) As Boolean
    Dim c As Long
    If Trim$(CStr(ham(r, 1))) = uy And _
       Len(Trim$(CStr(ham(r, 7)))) > 0 Then Exit Function
    If Len(Trim$(CStr(ham(r, 1)))) = 0 And _
       Len(Trim$(CStr(ham(r, 7)))) = 0 Then
        For c = 8 To 12
            If Len(Trim$(CStr(ham(r, c)))) > 0 Then
                YUK_CDP_OZET_MI = True: Exit Function
            End If
        Next c
    End If
End Function

Private Sub YUK_CDP_BLOK_BEKLE(ByVal saniye As Double)
    Dim t As Double
    t = Timer
    Do While YUK_CDP_BLOK_SURE(t) < saniye
        DoEvents
    Loop
End Sub

Private Function YUK_CDP_BLOK_SURE(ByVal t As Double) As Double
    If Timer >= t Then
        YUK_CDP_BLOK_SURE = Timer - t
    Else
        YUK_CDP_BLOK_SURE = 86400# - t + Timer
    End If
End Function


' 30'lu SAP ALV okuma: yalnizca gecerli stok satirlari aktarilir.
Public Function YUK_EPP_BLOK_AL(ByRef veri As Variant, _
    ByRef satir As Long, ByRef sutun As Long, _
    ByRef sure As Double, ByRef Hata As String) As Boolean
    Dim blok As Long
    Dim ses As Object, alv As Object, duzen As Object, kolon As Object
    Dim ids(1 To 12) As String, ham() As Variant, temiz() As Variant
    Dim i As Long, j As Long, b As Long, son As Long, n As Long
    Dim secili As Long, tekrar As Long, stok As Long
    Dim ad As String, tanim As String, depo As String
    Dim sorun As Boolean, ozet As Boolean, bitti As Boolean
    Dim t As Double, neden As String, uy As String
    On Error GoTo Hata
    satir = 0: sutun = 0: sure = 0: Hata = vbNullString
        uy = YUK_AKSADA_UY()
    t = Timer
    ' Hizli yol: ilin MB52 listesi tek sorguda metin olarak alinir (olmazsa asagidaki eski yontem).
    If AKHZ_TEK_AL(6, veri, satir, sutun, Hata) Then YUK_EPP_BLOK_AL = True: Exit Function
    Set ses = SAP_SESSION_AL("EPP", Hata)
    If ses Is Nothing Then Err.Raise vbObjectError + 870, , "SAP oturumu bulunamadi: " & Hata
    ses.StartTransaction "MB52"
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = uy
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString
    YUK_V2_SAP_DEPOLARI_SEC ses
    ses.FindById("wnd[0]").SendVKey 8
    YUK_EPP_BLOK_BEKLE 0.4
    ses.FindById("wnd[0]/tbar[1]/btn[33]").Press
    YUK_EPP_BLOK_BEKLE 0.3
    Set duzen = ses.FindById("wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")
    secili = -1
    For i = 0 To CLng(duzen.RowCount) - 1
        ad = Trim$(CStr(duzen.GetCellValue(i, "VARIANT")))
        tanim = Trim$(CStr(duzen.GetCellValue(i, "TEXT")))
        If StrComp(ad, "PANEL", vbTextCompare) = 0 Or _
           StrComp(tanim, "MLZ STOK PANEL", vbTextCompare) = 0 Then
            secili = i: Exit For
        End If
    Next i
    If secili < 0 Then Err.Raise vbObjectError + 871, , "PANEL duzeni bulunamadi."
    duzen.SetCurrentCell secili, "TEXT"
    duzen.SelectedRows = CStr(secili)
    duzen.ClickCurrentCell
    YUK_EPP_BLOK_BEKLE 0.3
    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    On Error Resume Next
    alv.PressToolbarButton "&DELETE_FILTER"
    Err.Clear
    On Error GoTo Hata
    Set kolon = alv.ColumnOrder
    If CLng(kolon.Count) < 12 Then Err.Raise vbObjectError + 872, , "PANEL 12 sutun icermiyor."
    For j = 1 To 12
        ids(j) = CStr(kolon.Item(j - 1))
    Next j
    n = CLng(alv.RowCount)
    If n <= 0 Then Err.Raise vbObjectError + 873, , "SAP sonucu bos."
    ReDim ham(1 To n, 1 To 12)
    ReDim temiz(1 To n, 1 To 12)
    blok = AKSORA_SAP_BLOK_BOYU(alv)
    For b = 0 To n - 1 Step blok
        If bitti Then Exit For
        son = b + blok - 1
        If son >= n Then son = n - 1
        On Error Resume Next
        Err.Clear
        alv.FirstVisibleRow = b
        sorun = (Err.Number <> 0)
        On Error GoTo Hata
        If sorun Then Err.Raise vbObjectError + 874, , "Blok kaydirilamadi: " & (b + 1)
        DoEvents
        YUK_EPP_BLOK_BEKLE 0.015
        AKSORA_TOPLU_SURE_TIK
        For i = b To son
            sorun = False: ozet = False
            On Error Resume Next
            For j = 1 To 12
                Err.Clear
                If j = 12 Then
                    ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                    ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                    If Trim$(CStr(ham(i + 1, 1))) = uy And _
                       Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                        ham(i + 1, j) = 0#
                    Else
                        ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                    End If
                Else
                    ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                End If
                If Err.Number <> 0 Then sorun = True: Exit For
            Next j
            On Error GoTo Hata
            If Not sorun Then
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
                ozet = YUK_EPP_OZET_MI(ham, i + 1, uy)
                If ozet Then bitti = True: Exit For
            End If
            For tekrar = 1 To 2
                sorun = False
                On Error Resume Next
                Err.Clear
                alv.FirstVisibleRow = i
                alv.SetCurrentCell i, ids(7)
                sorun = (Err.Number <> 0)
                On Error GoTo Hata
                If Not sorun Then
                    DoEvents
                    YUK_EPP_BLOK_BEKLE 0.015
                    On Error Resume Next
                    For j = 1 To 12
                        Err.Clear
                        If j = 12 Then
                            ' Gerçek stokta birim fiyat KAYDET aþamasýnda deðer/miktardan hesaplanýr.
                            ' Özet/þüpheli satýrda 12. alan yine SAP'den okunur.
                            If Trim$(CStr(ham(i + 1, 1))) = uy And _
                               Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then
                                ham(i + 1, j) = 0#
                            Else
                                ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                            End If
                        Else
                            ham(i + 1, j) = alv.GetCellValue(i, ids(j))
                        End If
                        If Err.Number <> 0 Then sorun = True: Exit For
                    Next j
                    On Error GoTo Hata
                    If Not sorun Then
                        If Trim$(CStr(ham(i + 1, 1))) = uy And _
                           Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then Exit For
                        If YUK_EPP_OZET_MI(ham, i + 1, uy) Then Exit For
                    End If
                End If
            Next tekrar
            If Not sorun Then
                If YUK_EPP_OZET_MI(ham, i + 1, uy) Then bitti = True: Exit For
                If Trim$(CStr(ham(i + 1, 1))) = uy And _
                   Len(Trim$(CStr(ham(i + 1, 7)))) > 0 Then GoTo KAYDET
            End If
            ' Ozet satiri SAP tarafindan secilemiyorsa son kisimda dur.
            If i >= n - 10 And stok > 0 And _
               Len(Trim$(CStr(ham(i + 1, 1)))) = 0 Then
                bitti = True: Exit For
            End If
            Err.Raise vbObjectError + 875, , "Satir dogrulanamadi: " & (i + 1)
KAYDET:
            depo = UCase$(Trim$(CStr(ham(i + 1, 2))))
            If Not YUK_V2_DEPO_GECERLI(depo) Then
                Err.Raise vbObjectError + 876, , "Beklenmeyen depo: " & depo
            End If
            stok = stok + 1
            For j = 1 To 12
                temiz(stok, j) = ham(i + 1, j)
            Next j
            temiz(stok, 9) = YUK_V2_EPP_SAYI(temiz(stok, 9))
            temiz(stok, 11) = YUK_V2_EPP_SAYI(temiz(stok, 11))
            If IsNumeric(temiz(stok, 9)) And CDbl(temiz(stok, 9)) <> 0 Then
                temiz(stok, 12) = CDbl(temiz(stok, 11)) / CDbl(temiz(stok, 9))
            Else
                temiz(stok, 12) = 0#
            End If
        Next i
    Next b
    If stok <= 0 Then Err.Raise vbObjectError + 877, , "Gecerli stok satiri yok."
    ReDim ham(1 To stok, 1 To 12)
    For i = 1 To stok
        For j = 1 To 12
            ham(i, j) = temiz(i, j)
        Next j
    Next i
    veri = ham: satir = stok: sutun = 12
    sure = YUK_EPP_BLOK_SURE(t)
    YUK_EPP_BLOK_AL = True
    Exit Function
Hata:
    neden = Err.Description
    Hata = "30'lu blok okuma: " & neden
    sure = YUK_EPP_BLOK_SURE(t)
    YUK_EPP_BLOK_AL = False
End Function

Private Function YUK_EPP_OZET_MI(ByRef ham() As Variant, _
    ByVal r As Long, ByVal uy As String) As Boolean
    Dim c As Long
    If Trim$(CStr(ham(r, 1))) = uy And _
       Len(Trim$(CStr(ham(r, 7)))) > 0 Then Exit Function
    If Len(Trim$(CStr(ham(r, 1)))) = 0 And _
       Len(Trim$(CStr(ham(r, 7)))) = 0 Then
        For c = 8 To 12
            If Len(Trim$(CStr(ham(r, c)))) > 0 Then
                YUK_EPP_OZET_MI = True: Exit Function
            End If
        Next c
    End If
End Function

Private Sub YUK_EPP_BLOK_BEKLE(ByVal saniye As Double)
    Dim t As Double
    t = Timer
    Do While YUK_EPP_BLOK_SURE(t) < saniye
        DoEvents
    Loop
End Sub

Private Function YUK_EPP_BLOK_SURE(ByVal t As Double) As Double
    If Timer >= t Then
        YUK_EPP_BLOK_SURE = Timer - t
    Else
        YUK_EPP_BLOK_SURE = 86400# - t + Timer
    End If
End Function
Private Function YUK_EDAS_UY() As String
    YUK_EDAS_UY = AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "CDP")
End Function

Private Function YUK_AKSADA_UY() As String
    YUK_AKSADA_UY = AKSORA_IL_TESIS(AKSORA_AKTIF_IL(), "EPP")
End Function

Private Sub YUK_V2_IL_RAPOR_SATIRLARI_YAZ( _
    ByVal wsRapor As Worksheet, _
    ByVal dict As Object, _
    ByRef HedefSatir As Long)

    Dim depolar As Variant
    Dim depo As Variant

    depolar = AKSORA_IL_YUKLENICI(AKSORA_AKTIF_IL())

    For Each depo In depolar
        YUK_V2_RAPOR_SATIR_YAZ _
            wsRapor, dict, CStr(depo), HedefSatir
    Next depo
End Sub


