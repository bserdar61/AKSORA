Attribute VB_Name = "modSTOK"
Option Explicit

'==========================================================
' AKSORA
' modSTOK
' ORTAK SAP MB52 -> EXCEL EXPORT -> RAM MOTORU
'==========================================================

Public Function STOK_EXPORT_AL( _
    ByVal SistemAdi As String, _
    ByVal UretimYeri As String, _
    ByVal DepoBaslangic As String, _
    ByVal DepoBitis As String, _
    ByRef VeriDizisi As Variant, _
    ByRef VeriSatiri As Long, _
    ByRef SutunSayisi As Long, _
    ByRef ExportSuresi As Double, _
    ByRef hataMesaji As String) As Boolean

    Dim SAPSession As Object
    Dim wbExport As Workbook
    Dim wsKaynak As Worksheet
    Dim EskiWorkbooklar As Object
    Dim baslangic As Double

    On Error GoTo Hata

    STOK_EXPORT_AL = False
    VeriSatiri = 0
    SutunSayisi = 0
    ExportSuresi = 0
    hataMesaji = vbNullString

    SistemAdi = UCase$(Trim$(SistemAdi))
    UretimYeri = Trim$(UretimYeri)
    DepoBaslangic = UCase$(Trim$(DepoBaslangic))
    DepoBitis = UCase$(Trim$(DepoBitis))

    If SistemAdi <> "CDP" And SistemAdi <> "EPP" Then
        hataMesaji = "Geçersiz SAP sistemi: " & SistemAdi
        Exit Function
    End If

    If Len(UretimYeri) = 0 Then
        hataMesaji = "Üretim yeri boþ býrakýlamaz."
        Exit Function
    End If

    If Len(DepoBaslangic) = 0 Or Len(DepoBitis) = 0 Then
        hataMesaji = "Depo aralýðý boþ býrakýlamaz."
        Exit Function
    End If

    baslangic = Timer
    Set EskiWorkbooklar = STOK_ACIK_WORKBOOKLARI_AL()

    CORE_DURUM SistemAdi & " SAP oturumu aranýyor..."
    Set SAPSession = SAP_SESSION_AL(SistemAdi, hataMesaji)

    If SAPSession Is Nothing Then
        If Len(hataMesaji) = 0 Then hataMesaji = SistemAdi & " SAP oturumu bulunamadý."
        Exit Function
    End If

    CORE_DURUM SistemAdi & " üzerinde MB52 açýlýyor..."
    SAPSession.StartTransaction "MB52"

    CORE_DURUM SistemAdi & " MB52 kriterleri hazýrlanýyor..."
    SAPSession.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = UretimYeri

    If DepoBaslangic = "*" And DepoBitis = "*" Then
        SAPSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
        SAPSession.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString
    Else
        SAPSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = DepoBaslangic
        SAPSession.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = DepoBitis
    End If

    If SistemAdi = "EPP" Then
        SAPSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").SetFocus
        SAPSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").caretPosition = Len(DepoBaslangic)
        SAPSession.FindById("wnd[0]").SendVKey 0
    End If

    CORE_DURUM SistemAdi & " MB52 raporu çalýþtýrýlýyor..."
    SAPSession.FindById("wnd[0]").SendVKey 8

    If Not STOK_ALV_BEKLE(SAPSession, 30) Then
        hataMesaji = "MB52 sonuç tablosu 30 saniye içinde hazýr olmadý."
        Exit Function
    End If

    CORE_DURUM SistemAdi & " PANEL düzeni seçiliyor..."
    If Not STOK_PANEL_DUZEN_SEC(SAPSession, hataMesaji) Then Exit Function

    If Not STOK_ALV_BEKLE(SAPSession, 10) Then
        hataMesaji = "PANEL düzeni seçildikten sonra MB52 tablosu hazýr olmadý."
        Exit Function
    End If

    CORE_DURUM SistemAdi & " Excel export hazýrlanýyor..."

    If SistemAdi = "CDP" Then
        If Not STOK_CDP_EXPORT_BASLAT(SAPSession, hataMesaji) Then Exit Function
    Else
        If Not STOK_EPP_EXPORT_BASLAT(SAPSession, hataMesaji) Then Exit Function
    End If

    Set wbExport = STOK_YENI_WORKBOOK_BEKLE(EskiWorkbooklar, 45)

    If wbExport Is Nothing Then
        hataMesaji = SistemAdi & " SAP export Excel dosyasý bulunamadý."
        Exit Function
    End If

    Set wsKaynak = STOK_EXPORT_SAYFASI_BUL(wbExport)

    If wsKaynak Is Nothing Then
        hataMesaji = SistemAdi & " SAP export çalýþma sayfasý bulunamadý."
        GoTo TEMIZ_HATA
    End If

    If Not STOK_EXPORT_VERI_AL( _
        wsKaynak, UretimYeri, DepoBaslangic, DepoBitis, _
        VeriDizisi, VeriSatiri, SutunSayisi, hataMesaji) Then
        GoTo TEMIZ_HATA
    End If

    ExportSuresi = STOK_GECEN_SURE(baslangic)

    On Error Resume Next
    wbExport.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0

    STOK_EXPORT_AL = True
    Exit Function

TEMIZ_HATA:
    On Error Resume Next
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0
    ExportSuresi = STOK_GECEN_SURE(baslangic)
    STOK_EXPORT_AL = False
    Exit Function

Hata:
    hataMesaji = "STOK_EXPORT_AL hatasý." & vbCrLf & _
                  "Hata No: " & Err.Number & vbCrLf & _
                  "Açýklama: " & Err.Description
    ExportSuresi = STOK_GECEN_SURE(baslangic)

    On Error Resume Next
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0

    STOK_EXPORT_AL = False
End Function

Public Function STOK_EXPORT_AL_COKLU_DEPO( _
    ByVal SistemAdi As String, _
    ByVal UretimYeri As String, _
    ByVal depolar As Variant, _
    ByRef VeriDizisi As Variant, _
    ByRef VeriSatiri As Long, _
    ByRef SutunSayisi As Long, _
    ByRef ExportSuresi As Double, _
    ByRef hataMesaji As String) As Boolean

    Dim SAPSession As Object
    Dim wbExport As Workbook
    Dim wsKaynak As Worksheet
    Dim EskiWorkbooklar As Object
    Dim baslangic As Double

    On Error GoTo Hata

    STOK_EXPORT_AL_COKLU_DEPO = False
    VeriSatiri = 0
    SutunSayisi = 0
    ExportSuresi = 0
    hataMesaji = vbNullString

    SistemAdi = UCase$(Trim$(SistemAdi))
    UretimYeri = Trim$(UretimYeri)

    If SistemAdi <> "CDP" And SistemAdi <> "EPP" Then
        hataMesaji = "Geçersiz SAP sistemi: " & SistemAdi
        Exit Function
    End If

    If Len(UretimYeri) = 0 Then
        hataMesaji = "Üretim yeri boþ býrakýlamaz."
        Exit Function
    End If

    If Not IsArray(depolar) Then
        hataMesaji = "Depo listesi geçerli deðil."
        Exit Function
    End If

    baslangic = Timer
    Set EskiWorkbooklar = STOK_ACIK_WORKBOOKLARI_AL()

    CORE_DURUM SistemAdi & " SAP oturumu aranýyor..."
    Set SAPSession = SAP_SESSION_AL(SistemAdi, hataMesaji)

    If SAPSession Is Nothing Then
        If Len(hataMesaji) = 0 Then hataMesaji = SistemAdi & " SAP oturumu bulunamadý."
        Exit Function
    End If

    CORE_DURUM SistemAdi & " üzerinde MB52 açýlýyor..."
    SAPSession.StartTransaction "MB52"

    SAPSession.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = UretimYeri
    SAPSession.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
    SAPSession.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString

    If Not STOK_COKLU_DEPO_SEC(SAPSession, depolar, hataMesaji) Then Exit Function

    SAPSession.FindById("wnd[0]").SendVKey 8

    If Not STOK_ALV_BEKLE(SAPSession, 30) Then
        hataMesaji = "MB52 sonuç tablosu 30 saniye içinde hazýr olmadý."
        Exit Function
    End If

    If Not STOK_PANEL_DUZEN_SEC(SAPSession, hataMesaji) Then Exit Function

    If Not STOK_ALV_BEKLE(SAPSession, 10) Then
        hataMesaji = "PANEL düzeni seçildikten sonra MB52 tablosu hazýr olmadý."
        Exit Function
    End If

    CORE_DURUM SistemAdi & " Excel export hazýrlanýyor..."

    If SistemAdi = "CDP" Then
        If Not STOK_CDP_EXPORT_BASLAT(SAPSession, hataMesaji) Then Exit Function
    Else
        If Not STOK_EPP_EXPORT_BASLAT(SAPSession, hataMesaji) Then Exit Function
    End If

    Set wbExport = STOK_YENI_WORKBOOK_BEKLE(EskiWorkbooklar, 45)

    If wbExport Is Nothing Then
        hataMesaji = SistemAdi & " SAP export Excel dosyasý bulunamadý."
        Exit Function
    End If

    Set wsKaynak = STOK_EXPORT_SAYFASI_BUL(wbExport)

    If wsKaynak Is Nothing Then
        hataMesaji = SistemAdi & " SAP export çalýþma sayfasý bulunamadý."
        GoTo TEMIZ_HATA
    End If

    If Not STOK_EXPORT_VERI_AL_COKLU( _
        wsKaynak, UretimYeri, depolar, _
        VeriDizisi, VeriSatiri, SutunSayisi, hataMesaji) Then
        GoTo TEMIZ_HATA
    End If

    ExportSuresi = STOK_GECEN_SURE(baslangic)

    On Error Resume Next
    wbExport.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0

    STOK_EXPORT_AL_COKLU_DEPO = True
    Exit Function

TEMIZ_HATA:
    On Error Resume Next
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0
    ExportSuresi = STOK_GECEN_SURE(baslangic)
    STOK_EXPORT_AL_COKLU_DEPO = False
    Exit Function

Hata:
    hataMesaji = "STOK_EXPORT_AL_COKLU_DEPO hatasý." & vbCrLf & _
                  "Hata No: " & Err.Number & vbCrLf & _
                  "Açýklama: " & Err.Description
    ExportSuresi = STOK_GECEN_SURE(baslangic)

    On Error Resume Next
    If Not wbExport Is Nothing Then wbExport.Close SaveChanges:=False
    ThisWorkbook.Activate
    On Error GoTo 0

    STOK_EXPORT_AL_COKLU_DEPO = False
End Function

Private Function STOK_EXPORT_VERI_AL( _
    ByVal ws As Worksheet, _
    ByVal BeklenenUY As String, _
    ByVal DepoBaslangic As String, _
    ByVal DepoBitis As String, _
    ByRef sonuc As Variant, _
    ByRef SatirSayisi As Long, _
    ByRef SutunSayisi As Long, _
    ByRef hataMesaji As String) As Boolean

    Dim ham As Variant
    Dim temiz() As Variant
    Dim sonSatir As Long
    Dim sonSutun As Long
    Dim i As Long, j As Long, n As Long
    Dim uy As String, depo As String, malzeme As String

    On Error GoTo Hata

    STOK_EXPORT_VERI_AL = False
    SatirSayisi = 0
    SutunSayisi = 0

    sonSatir = STOK_SON_SATIR(ws)
    sonSutun = STOK_SON_SUTUN(ws)

    If sonSatir < 2 Or sonSutun < 11 Then
        hataMesaji = "SAP export dosyasýnda geçerli veri bulunamadý."
        Exit Function
    End If

    ham = ws.Range(ws.Cells(2, 1), ws.Cells(sonSatir, sonSutun)).Value2
    ReDim temiz(1 To UBound(ham, 1), 1 To 12)

    For i = 1 To UBound(ham, 1)

        uy = Trim$(CStr(ham(i, 1)))
        depo = UCase$(Trim$(CStr(ham(i, 2))))

        If UBound(ham, 2) >= 7 Then
            malzeme = Trim$(CStr(ham(i, 7)))
        Else
            malzeme = vbNullString
        End If

        If STOK_GECERLI_VERI_SATIRI(uy, depo, malzeme, BeklenenUY) Then

            If STOK_DEPO_ARALIKTA(depo, DepoBaslangic, DepoBitis) Then
                n = n + 1

                For j = 1 To 12
                    If j <= UBound(ham, 2) Then
                        temiz(n, j) = ham(i, j)
                    Else
                        temiz(n, j) = vbNullString
                    End If
                Next j

                temiz(n, 9) = STOK_SAP_SAYI_CEVIR(temiz(n, 9))
                temiz(n, 11) = STOK_SAP_SAYI_CEVIR(temiz(n, 11))
                temiz(n, 12) = STOK_SAP_SAYI_CEVIR(temiz(n, 12))
            End If

        End If
    Next i

    If n = 0 Then
        hataMesaji = BeklenenUY & " için geçerli stok satýrý bulunamadý."
        Exit Function
    End If

    sonuc = STOK_DIZI_KIRP(temiz, n, 12)
    SatirSayisi = n
    SutunSayisi = 12
    STOK_EXPORT_VERI_AL = True
    Exit Function

Hata:
    hataMesaji = "SAP export verisi okunamadý." & vbCrLf & Err.Description
    STOK_EXPORT_VERI_AL = False
End Function

Private Function STOK_EXPORT_VERI_AL_COKLU( _
    ByVal ws As Worksheet, _
    ByVal BeklenenUY As String, _
    ByVal depolar As Variant, _
    ByRef sonuc As Variant, _
    ByRef SatirSayisi As Long, _
    ByRef SutunSayisi As Long, _
    ByRef hataMesaji As String) As Boolean

    Dim ham As Variant
    Dim temiz() As Variant
    Dim sonSatir As Long
    Dim sonSutun As Long
    Dim i As Long, j As Long, n As Long
    Dim uy As String, depo As String, malzeme As String

    On Error GoTo Hata

    STOK_EXPORT_VERI_AL_COKLU = False
    SatirSayisi = 0
    SutunSayisi = 0

    sonSatir = STOK_SON_SATIR(ws)
    sonSutun = STOK_SON_SUTUN(ws)

    If sonSatir < 2 Or sonSutun < 11 Then
        hataMesaji = "SAP export dosyasýnda geçerli veri bulunamadý."
        Exit Function
    End If

    ham = ws.Range(ws.Cells(2, 1), ws.Cells(sonSatir, sonSutun)).Value2
    ReDim temiz(1 To UBound(ham, 1), 1 To 12)

    For i = 1 To UBound(ham, 1)

        uy = Trim$(CStr(ham(i, 1)))
        depo = UCase$(Trim$(CStr(ham(i, 2))))

        If UBound(ham, 2) >= 7 Then
            malzeme = Trim$(CStr(ham(i, 7)))
        Else
            malzeme = vbNullString
        End If

        If STOK_GECERLI_VERI_SATIRI(uy, depo, malzeme, BeklenenUY) Then
            If STOK_DEPO_LISTEDE(depo, depolar) Then

                n = n + 1

                For j = 1 To 12
                    If j <= UBound(ham, 2) Then
                        temiz(n, j) = ham(i, j)
                    Else
                        temiz(n, j) = vbNullString
                    End If
                Next j

                temiz(n, 9) = STOK_SAP_SAYI_CEVIR(temiz(n, 9))
                temiz(n, 11) = STOK_SAP_SAYI_CEVIR(temiz(n, 11))
                temiz(n, 12) = STOK_SAP_SAYI_CEVIR(temiz(n, 12))
            End If
        End If
    Next i

    If n = 0 Then
        hataMesaji = BeklenenUY & " için geçerli stok satýrý bulunamadý."
        Exit Function
    End If

    sonuc = STOK_DIZI_KIRP(temiz, n, 12)
    SatirSayisi = n
    SutunSayisi = 12
    STOK_EXPORT_VERI_AL_COKLU = True
    Exit Function

Hata:
    hataMesaji = "SAP export verisi okunamadý." & vbCrLf & Err.Description
    STOK_EXPORT_VERI_AL_COKLU = False
End Function

Private Function STOK_GECERLI_VERI_SATIRI( _
    ByVal uy As String, _
    ByVal depo As String, _
    ByVal malzeme As String, _
    ByVal BeklenenUY As String) As Boolean

    STOK_GECERLI_VERI_SATIRI = False

    If StrComp(Trim$(uy), Trim$(BeklenenUY), vbTextCompare) <> 0 Then Exit Function
    If Len(Trim$(depo)) = 0 Then Exit Function
    If Len(Trim$(malzeme)) = 0 Then Exit Function

    STOK_GECERLI_VERI_SATIRI = True
End Function

Private Function STOK_DEPO_ARALIKTA( _
    ByVal depo As String, _
    ByVal baslangic As String, _
    ByVal bitis As String) As Boolean

    depo = UCase$(Trim$(depo))
    baslangic = UCase$(Trim$(baslangic))
    bitis = UCase$(Trim$(bitis))

    If baslangic = "*" And bitis = "*" Then
        STOK_DEPO_ARALIKTA = True
    Else
        STOK_DEPO_ARALIKTA = (StrComp(depo, baslangic, vbTextCompare) >= 0 And _
                              StrComp(depo, bitis, vbTextCompare) <= 0)
    End If
End Function

Private Function STOK_DEPO_LISTEDE(ByVal depo As String, ByVal depolar As Variant) As Boolean
    Dim i As Long

    For i = LBound(depolar) To UBound(depolar)
        If StrComp(Trim$(depo), Trim$(CStr(depolar(i))), vbTextCompare) = 0 Then
            STOK_DEPO_LISTEDE = True
            Exit Function
        End If
    Next i
End Function

Private Function STOK_DIZI_KIRP( _
    ByRef kaynak As Variant, _
    ByVal satir As Long, _
    ByVal sutun As Long) As Variant

    Dim sonuc() As Variant
    Dim i As Long, j As Long

    ReDim sonuc(1 To satir, 1 To sutun)

    For i = 1 To satir
        For j = 1 To sutun
            sonuc(i, j) = kaynak(i, j)
        Next j
    Next i

    STOK_DIZI_KIRP = sonuc
End Function

Private Function STOK_COKLU_DEPO_SEC( _
    ByVal SAPSession As Object, _
    ByVal depolar As Variant, _
    ByRef hataMesaji As String) As Boolean

    Dim temelID As String
    Dim i As Long
    Dim satir As Long

    On Error GoTo Hata

    STOK_COKLU_DEPO_SEC = False

    SAPSession.FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press

    If Not STOK_SAP_PENCERE_BEKLE(SAPSession, "wnd[1]", 10) Then
        hataMesaji = "Depo çoklu seçim penceresi açýlamadý."
        Exit Function
    End If

    temelID = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/" & _
              "ssubSCREEN_HEADER:SAPLALDB:3010/" & _
              "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I"

    satir = 0

    For i = LBound(depolar) To UBound(depolar)
        If Len(Trim$(CStr(depolar(i)))) > 0 Then
            SAPSession.FindById(temelID & "[1," & satir & "]").Text = _
                UCase$(Trim$(CStr(depolar(i))))
            satir = satir + 1
        End If
    Next i

    SAPSession.FindById("wnd[1]/tbar[0]/btn[8]").Press

    STOK_COKLU_DEPO_SEC = True
    Exit Function

Hata:
    hataMesaji = "Depolar MB52 çoklu seçimine yazýlamadý." & vbCrLf & _
                  "Hata No: " & Err.Number & vbCrLf & _
                  "Açýklama: " & Err.Description
End Function

Private Function STOK_PANEL_DUZEN_SEC( _
    ByVal SAPSession As Object, _
    ByRef hataMesaji As String) As Boolean

    Dim LayoutGrid As Object
    Dim satir As Long
    Dim SatirSayisi As Long
    Dim DuzenAdi As String
    Dim DuzenTanimi As String
    Dim BulunanSatir As Long

    On Error GoTo Hata

    STOK_PANEL_DUZEN_SEC = False
    BulunanSatir = -1

    SAPSession.FindById("wnd[0]/tbar[1]/btn[33]").Press

    If Not STOK_SAP_PENCERE_BEKLE(SAPSession, "wnd[1]", 10) Then
        hataMesaji = "MB52 düzen seçme penceresi açýlamadý."
        Exit Function
    End If

    Set LayoutGrid = SAPSession.FindById( _
        "wnd[1]/usr/subSUB_CONFIGURATION:SAPLSALV_CUL_LAYOUT_CHOOSE:0500/cntlD500_CONTAINER/shellcont/shell")

    SatirSayisi = LayoutGrid.RowCount

    For satir = 0 To SatirSayisi - 1

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
        hataMesaji = "SAP MB52 içinde PANEL düzeni bulunamadý."
        On Error Resume Next
        SAPSession.FindById("wnd[1]").SendVKey 12
        On Error GoTo 0
        Exit Function
    End If

    LayoutGrid.SetCurrentCell BulunanSatir, "TEXT"
    LayoutGrid.SelectedRows = CStr(BulunanSatir)
    LayoutGrid.ClickCurrentCell

    STOK_KISA_BEKLE 0.3

    STOK_PANEL_DUZEN_SEC = True
    Exit Function

Hata:
    hataMesaji = "PANEL düzeni otomatik seçilemedi." & vbCrLf & _
                  "Hata No: " & Err.Number & vbCrLf & _
                  "Açýklama: " & Err.Description
    On Error Resume Next
    SAPSession.FindById("wnd[1]").SendVKey 12
    On Error GoTo 0
End Function

Private Function STOK_CDP_EXPORT_BASLAT( _
    ByVal SAPSession As Object, _
    ByRef hataMesaji As String) As Boolean

    On Error GoTo Hata

    STOK_CDP_EXPORT_BASLAT = False
    SAPSession.FindById("wnd[0]").SendVKey 43
    STOK_CDP_EXPORT_BASLAT = True
    Exit Function

Hata:
    hataMesaji = "CDP Excel export baþlatýlamadý." & vbCrLf & Err.Description
End Function

Private Function STOK_EPP_EXPORT_BASLAT( _
    ByVal SAPSession As Object, _
    ByRef hataMesaji As String) As Boolean

    Dim alv As Object

    On Error GoTo Hata

    STOK_EPP_EXPORT_BASLAT = False

    Set alv = SAPSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")

    alv.SetCurrentCell -1, "MAKTX"
    alv.SelectColumn "MAKTX"
    alv.ContextMenu
    alv.SelectContextMenuItem "&XXL"

    If Not STOK_SAP_PENCERE_BEKLE(SAPSession, "wnd[1]", 10) Then
        hataMesaji = "EPP Excel export penceresi açýlmadý."
        Exit Function
    End If

    On Error Resume Next
    Err.Clear
    SAPSession.FindById("wnd[1]/tbar[0]/btn[20]").Press

    If Err.Number <> 0 Then
        Err.Clear
        SAPSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    End If
    Err.Clear
    On Error GoTo Hata

    STOK_KISA_BEKLE 0.2

    On Error Resume Next
    SAPSession.FindById("wnd[1]/tbar[0]/btn[0]").Press
    Err.Clear
    On Error GoTo Hata

    STOK_EPP_EXPORT_BASLAT = True
    Exit Function

Hata:
    hataMesaji = "EPP Excel export baþlatýlamadý." & vbCrLf & Err.Description
End Function

Private Function STOK_ACIK_WORKBOOKLARI_AL() As Object

    Dim dict As Object
    Dim wb As Workbook

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    For Each wb In application.Workbooks
        dict(wb.Name) = True
    Next wb

    Set STOK_ACIK_WORKBOOKLARI_AL = dict
End Function

Private Function STOK_YENI_WORKBOOK_BEKLE( _
    ByVal EskiWorkbooklar As Object, _
    Optional ByVal TimeoutSaniye As Double = 45) As Workbook

    Dim baslangic As Double
    Dim wb As Workbook

    baslangic = Timer

    Do
        DoEvents

        For Each wb In application.Workbooks
            If Not wb Is ThisWorkbook Then
                If Not EskiWorkbooklar.Exists(wb.Name) Then
                    Set STOK_YENI_WORKBOOK_BEKLE = wb
                    Exit Function
                End If
            End If
        Next wb

        STOK_KISA_BEKLE 0.1

        If STOK_TIMEOUT_DOLDU(baslangic, TimeoutSaniye) Then Exit Do
    Loop

    Set STOK_YENI_WORKBOOK_BEKLE = Nothing
End Function

Private Function STOK_EXPORT_SAYFASI_BUL(ByVal wb As Workbook) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets("RawData")
    On Error GoTo 0

    If Not ws Is Nothing Then
        If STOK_SON_SATIR(ws) >= 2 And STOK_SON_SUTUN(ws) >= 11 Then
            Set STOK_EXPORT_SAYFASI_BUL = ws
            Exit Function
        End If
    End If

    Set ws = Nothing

    For Each ws In wb.Worksheets
        If STOK_SON_SATIR(ws) >= 2 And STOK_SON_SUTUN(ws) >= 11 Then
            Set STOK_EXPORT_SAYFASI_BUL = ws
            Exit Function
        End If
    Next ws
End Function

Private Function STOK_ALV_BEKLE( _
    ByVal SAPSession As Object, _
    Optional ByVal TimeoutSaniye As Double = 30) As Boolean

    Dim baslangic As Double
    Dim alv As Object

    baslangic = Timer

    Do
        DoEvents
        Set alv = Nothing

        On Error Resume Next
        Set alv = SAPSession.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
        On Error GoTo 0

        If Not alv Is Nothing Then
            STOK_ALV_BEKLE = True
            Exit Function
        End If

        STOK_KISA_BEKLE 0.1

        If STOK_TIMEOUT_DOLDU(baslangic, TimeoutSaniye) Then Exit Do
    Loop
End Function

Private Function STOK_SAP_PENCERE_BEKLE( _
    ByVal SAPSession As Object, _
    ByVal PencereID As String, _
    Optional ByVal TimeoutSaniye As Double = 10) As Boolean

    Dim baslangic As Double
    Dim obj As Object

    baslangic = Timer

    Do
        DoEvents
        Set obj = Nothing

        On Error Resume Next
        Set obj = SAPSession.FindById(PencereID)
        On Error GoTo 0

        If Not obj Is Nothing Then
            STOK_SAP_PENCERE_BEKLE = True
            Exit Function
        End If

        STOK_KISA_BEKLE 0.1

        If STOK_TIMEOUT_DOLDU(baslangic, TimeoutSaniye) Then Exit Do
    Loop
End Function

Private Function STOK_SAP_SAYI_CEVIR(ByVal deger As Variant) As Double

    Dim metin As String
    Dim negatif As Boolean

    On Error GoTo Cikis

    If IsNumeric(deger) Then
        STOK_SAP_SAYI_CEVIR = CDbl(deger)
        Exit Function
    End If

    metin = Trim$(CStr(deger))
    If Len(metin) = 0 Then Exit Function

    metin = Replace(metin, " ", "")
    metin = Replace(metin, ChrW(160), "")

    If Left$(metin, 1) = "-" Then
        negatif = True
        metin = Mid$(metin, 2)
    ElseIf Right$(metin, 1) = "-" Then
        negatif = True
        metin = Left$(metin, Len(metin) - 1)
    End If

    If InStr(metin, ",") > 0 Then
        metin = Replace(metin, ".", "")
        metin = Replace(metin, ",", ".")
    End If

    STOK_SAP_SAYI_CEVIR = Val(metin)

    If negatif Then STOK_SAP_SAYI_CEVIR = -STOK_SAP_SAYI_CEVIR
    Exit Function

Cikis:
    STOK_SAP_SAYI_CEVIR = 0
End Function

Public Function STOK_SON_SATIR(ByVal ws As Worksheet) As Long

    Dim sonHucre As Range

    On Error Resume Next
    Set sonHucre = ws.Cells.Find( _
        What:="*", _
        after:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If sonHucre Is Nothing Then
        STOK_SON_SATIR = 0
    Else
        STOK_SON_SATIR = sonHucre.Row
    End If
End Function

Public Function STOK_SON_SUTUN(ByVal ws As Worksheet) As Long

    Dim sonHucre As Range

    On Error Resume Next
    Set sonHucre = ws.Cells.Find( _
        What:="*", _
        after:=ws.Cells(1, 1), _
        LookIn:=xlFormulas, _
        LookAt:=xlPart, _
        SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, _
        MatchCase:=False)
    On Error GoTo 0

    If sonHucre Is Nothing Then
        STOK_SON_SUTUN = 0
    Else
        STOK_SON_SUTUN = sonHucre.Column
    End If
End Function

Private Sub STOK_KISA_BEKLE(ByVal saniye As Double)

    Dim baslangic As Double

    baslangic = Timer

    Do
        DoEvents
        If STOK_TIMEOUT_DOLDU(baslangic, saniye) Then Exit Do
    Loop
End Sub

Private Function STOK_TIMEOUT_DOLDU( _
    ByVal baslangic As Double, _
    ByVal TimeoutSaniye As Double) As Boolean

    STOK_TIMEOUT_DOLDU = (STOK_GECEN_SURE(baslangic) >= TimeoutSaniye)
End Function

Public Function STOK_GECEN_SURE(ByVal baslangic As Double) As Double

    Dim Simdi As Double

    Simdi = Timer

    If Simdi >= baslangic Then
        STOK_GECEN_SURE = Simdi - baslangic
    Else
        STOK_GECEN_SURE = (86400# - baslangic) + Simdi
    End If
End Function

Private Sub STOK_MOTOR_TEST( _
    ByVal SistemAdi As String, _
    ByVal UretimYeri As String, _
    ByVal DepoBaslangic As String, _
    ByVal DepoBitis As String)

    Dim veri As Variant
    Dim SatirSayisi As Long
    Dim SutunSayisi As Long
    Dim ExportSuresi As Double
    Dim ToplamSure As Double
    Dim hataMesaji As String
    Dim basarili As Boolean

    On Error GoTo Hata

    CORE_ISLEM_BASLAT SistemAdi & " ortak stok motoru testi"

    basarili = STOK_EXPORT_AL( _
        SistemAdi, UretimYeri, DepoBaslangic, DepoBitis, _
        veri, SatirSayisi, SutunSayisi, ExportSuresi, hataMesaji)

    ToplamSure = CORE_ISLEM_BITIR()

    If Not basarili Then
        MsgBox SistemAdi & " ortak stok motoru testi baþarýsýz." & vbCrLf & vbCrLf & _
               hataMesaji, vbCritical, "AKSORA"
        Exit Sub
    End If

    MsgBox SistemAdi & " ortak stok motoru baþarýyla çalýþtý." & vbCrLf & vbCrLf & _
           "Üretim Yeri : " & UretimYeri & vbCrLf & _
           "Depo Aralýðý: " & DepoBaslangic & " - " & DepoBitis & vbCrLf & _
           "Veri Satýrý : " & AKSORA_SAYI_METNI(SatirSayisi, "#,##0") & vbCrLf & _
           "Sütun       : " & AKSORA_SAYI_METNI(SutunSayisi, "#,##0") & vbCrLf & _
           "Export Süre : " & AKSORA_SAYI_METNI(ExportSuresi, "0.00") & " sn" & vbCrLf & _
           "Toplam Süre : " & AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn", _
           vbInformation, "AKSORA - " & SistemAdi
    Exit Sub

Hata:
    hataMesaji = Err.Description
    ToplamSure = CORE_ISLEM_BITIR()

    MsgBox SistemAdi & " ortak stok motoru testinde hata oluþtu." & vbCrLf & vbCrLf & _
           "Hata No: " & Err.Number & vbCrLf & _
           "Açýklama: " & hataMesaji, _
           vbCritical, "AKSORA"
End Sub

