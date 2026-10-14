Attribute VB_Name = "modAKSORA_GUNCELLEME"
Option Explicit

Public ANA_SESSIZ_MOD As Boolean

Public Sub ANA_TUM_STOKLARI_GUNCELLE_V2()

Dim baslangic As Double
Dim ToplamSure As Double
Dim NormalSatir As Long, IlceSatir As Long, YukSatir As Long
Dim NormalDeger As Double, IlceDeger As Double, YukDeger As Double
Dim GenelSatir As Long, GenelDeger As Double
Dim hataNo As Long, hataAciklama As String
Dim t0 As Double, tNormal As Double, tIlce As Double, tYuk As Double, tKars As Double, tPanel As Double, dokum As String
If Not AKSORA_SAP_GUNCELLEME_IZNI() Then Exit Sub
On Error GoTo Hata

baslangic = Timer
ANA_SESSIZ_MOD = True

ThisWorkbook.Worksheets("ANA PANEL").Activate
PANEL_CANLI_DURUM "SAP baðlantýsý hazýrlanýyor..."

ANA_TUM_SAYFA_KONTROL "CDP_STOK"
ANA_TUM_SAYFA_KONTROL "CDP_RAPOR"
ANA_TUM_SAYFA_KONTROL "EPP_STOK"
ANA_TUM_SAYFA_KONTROL "EPP_RAPOR"
ANA_TUM_SAYFA_KONTROL "STOKLAR"
ANA_TUM_SAYFA_KONTROL "RAPORLAR"
ANA_TUM_SAYFA_KONTROL "ÝLÇE_STOK"
ANA_TUM_SAYFA_KONTROL "ÝLÇE_RAPORLAR"
ANA_TUM_SAYFA_KONTROL "YÜKLENÝCÝ_STOK"
ANA_TUM_SAYFA_KONTROL "YÜKLENÝCÝ_RAPOR"

t0 = Timer
CORE_DURUM "[1/3] Normal stok sistemi güncelleniyor..."
PANEL_CANLI_DURUM "1/3 | CDP ve EPP normal stoklarý SAP'den alýnýyor..."
ANA_NORMAL_STOK_GUNCELLE_V2

tNormal = ANA_TUM_GECEN_SURE(t0): t0 = Timer
NormalSatir = ANA_TUM_LONG(ThisWorkbook.Worksheets("STOKLAR").Range("C3"))
NormalDeger = ANA_TUM_DOUBLE(ThisWorkbook.Worksheets("STOKLAR").Range("C4"))
If NormalSatir <= 0 Then Err.Raise vbObjectError + 6002, "ANA_TUM_STOKLARI_GUNCELLE_V2", "Normal stok güncellemesi sonrasýnda veri bulunamadý."

CORE_DURUM "[2/3] Ýlçe stok sistemi güncelleniyor..."
PANEL_CANLI_DURUM "2/3 | Ýlçe stoklarý SAP'den alýnýyor..."
ILCE_STOK_GUNCELLE_V2

tIlce = ANA_TUM_GECEN_SURE(t0): t0 = Timer
IlceSatir = ANA_TUM_LONG(ThisWorkbook.Worksheets("ÝLÇE_STOK").Range("C3"))
IlceDeger = ANA_TUM_DOUBLE(ThisWorkbook.Worksheets("ÝLÇE_STOK").Range("C4"))
If IlceSatir <= 0 Then Err.Raise vbObjectError + 6011, "ANA_TUM_STOKLARI_GUNCELLE_V2", "Ýlçe stok güncellemesi sonrasýnda veri bulunamadý."
ANA_TUM_ILCE_RAPOR_DOGRULA IlceDeger

CORE_DURUM "[3/3] Yüklenici stok sistemi güncelleniyor..."
PANEL_CANLI_DURUM "3/3 | Yüklenici stoklarý SAP'den alýnýyor..."
YUKLENICI_STOK_GUNCELLE_V2

tYuk = ANA_TUM_GECEN_SURE(t0): t0 = Timer
YukSatir = ANA_TUM_LONG(ThisWorkbook.Worksheets("YÜKLENÝCÝ_STOK").Range("C3"))
YukDeger = ANA_TUM_DOUBLE(ThisWorkbook.Worksheets("YÜKLENÝCÝ_STOK").Range("C4"))
If YukSatir <= 0 Then Err.Raise vbObjectError + 6021, "ANA_TUM_STOKLARI_GUNCELLE_V2", "Yüklenici stok güncellemesi sonrasýnda veri bulunamadý."
ANA_TUM_YUK_RAPOR_DOGRULA YukDeger

application.Calculation = xlCalculationAutomatic
DoEvents


CORE_DURUM "Stok karþýlaþtýrma tarihleri güncelleniyor..."
PANEL_CANLI_DURUM "Stok karþýlaþtýrma tarihleri güncelleniyor..."
t0 = Timer
STOK_KARS_TARIHLERI_YUKLE
tKars = ANA_TUM_GECEN_SURE(t0)


GenelSatir = NormalSatir + IlceSatir + YukSatir
GenelDeger = NormalDeger + IlceDeger + YukDeger
ToplamSure = ANA_TUM_GECEN_SURE(baslangic)

ANA_SESSIZ_MOD = False
CORE_DURUM vbNullString
AKSORA_PANEL_SURE_KAYDET baslangic
t0 = Timer
AKSORA_PANEL_CANLI = False
AKSORA_V11_BICIMLER
AKSORA_V11_ARAMA_TARIH
AKSORA_V10_TARIH_GOSTER
' Sonuc penceresi acilmadan once guncel ana panel ekranda gosterilir.
AKSORA_ANA_PANEL_GOSTER
tPanel = ANA_TUM_GECEN_SURE(t0)
dokum = vbCrLf & "Süre dökümü" & vbCrLf & _
" Normal (CDP+EPP) : " & AKSORA_SAYI_METNI(tNormal, "0.0") & " sn" & vbCrLf & _
" Ýlçe             : " & AKSORA_SAYI_METNI(tIlce, "0.0") & " sn" & vbCrLf & _
" Yüklenici        : " & AKSORA_SAYI_METNI(tYuk, "0.0") & " sn" & vbCrLf & _
" Karþýlaþtýrma    : " & AKSORA_SAYI_METNI(tKars, "0.0") & " sn" & vbCrLf & _
" Panel ve ekran   : " & AKSORA_SAYI_METNI(tPanel, "0.0") & " sn"

MsgBox _
"TÜM STOK GÜNCELLEMELERÝ TAMAMLANDI." & vbCrLf & vbCrLf & _
"NORMAL STOK" & vbCrLf & _
" Satýr : " & AKSORA_SAYI_METNI(NormalSatir, "#,##0") & vbCrLf & _
" Deðer : " & AKSORA_SAYI_METNI(NormalDeger, "#,##0.00") & " TL" & vbCrLf & vbCrLf & _
"ÝLÇE STOK" & vbCrLf & _
" Satýr : " & AKSORA_SAYI_METNI(IlceSatir, "#,##0") & vbCrLf & _
" Deðer : " & AKSORA_SAYI_METNI(IlceDeger, "#,##0.00") & " TL" & vbCrLf & vbCrLf & _
"YÜKLENÝCÝ STOK" & vbCrLf & _
" Satýr : " & AKSORA_SAYI_METNI(YukSatir, "#,##0") & vbCrLf & _
" Deðer : " & AKSORA_SAYI_METNI(YukDeger, "#,##0.00") & " TL" & vbCrLf & vbCrLf & _
"GENEL" & vbCrLf & _
" Satýr : " & AKSORA_SAYI_METNI(GenelSatir, "#,##0") & vbCrLf & _
" Deðer : " & AKSORA_SAYI_METNI(GenelDeger, "#,##0.00") & " TL" & vbCrLf & vbCrLf & _
"Kontroller : BAÞARILI" & vbCrLf & _
"Toplam Süre: " & AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn" & vbCrLf & dokum, _
vbInformation, "AKSORA - TÜM STOKLAR"

Exit Sub

Hata:

    hataNo = Err.Number
    hataAciklama = Err.Description
    ANA_SESSIZ_MOD = False

    On Error Resume Next

    CORE_DURUM vbNullString
    PANEL_YENILE
    PANEL_CANLI_DURUM "GÜNCELLEME TAMAMLANAMADI — KONTROL GEREKÝYOR"

    With ThisWorkbook.Worksheets("ANA PANEL").Range("B6:AB6")
        .Interior.Color = RGB(255, 218, 218)
        .Font.Color = RGB(165, 37, 37)
        .Font.bold = True
    End With

    On Error GoTo 0

    MsgBox _
        "TÜM STOK GÜNCELLEME ZÝNCÝRÝ TAMAMLANAMADI." & _
        vbCrLf & vbCrLf & _
        "Hata No: " & hataNo & vbCrLf & _
        "Açýklama: " & hataAciklama & vbCrLf & vbCrLf & _
        "Sonraki güncelleme adýmlarý durduruldu." & vbCrLf & _
        "Önceki adýmlarda tamamlanan güncellemeler korunur.", _
        vbCritical, "AKSORA"

End Sub

Private Sub ANA_TUM_SAYFA_KONTROL(ByVal sayfaAdi As String)
Dim ws As Worksheet
Set ws = Nothing
On Error Resume Next
Set ws = ThisWorkbook.Worksheets(sayfaAdi)
On Error GoTo 0
If ws Is Nothing Then Err.Raise vbObjectError + 6050, "ANA_TUM_SAYFA_KONTROL", sayfaAdi & " sayfasý bulunamadý."
End Sub

Private Sub ANA_TUM_ILCE_RAPOR_DOGRULA(ByVal BeklenenDeger As Double)
Dim ws As Worksheet, sonSatir As Long, i As Long, toplamSatiri As Long, RaporDeger As Double
Set ws = ThisWorkbook.Worksheets("ÝLÇE_RAPORLAR")
sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
For i = 9 To sonSatir
    If UCase$(Trim$(CStr(ws.Cells(i, 1).Value2))) = "TOPLAM" Then toplamSatiri = i: Exit For
Next i
If toplamSatiri = 0 Then Err.Raise vbObjectError + 6060, "ANA_TUM_ILCE_RAPOR_DOGRULA", "ÝLÇE_RAPORLAR sayfasýnda TOPLAM satýrý bulunamadý."
RaporDeger = ANA_TUM_DOUBLE(ws.Cells(toplamSatiri, 5))
If Abs(RaporDeger - BeklenenDeger) > 0.02 Then Err.Raise vbObjectError + 6061, "ANA_TUM_ILCE_RAPOR_DOGRULA", "ÝLÇE_RAPORLAR toplam deðeri ÝLÇE_STOK ile uyuþmuyor."
End Sub

Private Sub ANA_TUM_YUK_RAPOR_DOGRULA(ByVal BeklenenDeger As Double)
Dim ws As Worksheet, sonSatir As Long, i As Long, toplamSatiri As Long, RaporDeger As Double
Set ws = ThisWorkbook.Worksheets("YÜKLENÝCÝ_RAPOR")
sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
For i = 9 To sonSatir
    If UCase$(Trim$(CStr(ws.Cells(i, 1).Value2))) = "TOPLAM" Then toplamSatiri = i: Exit For
Next i
If toplamSatiri = 0 Then Err.Raise vbObjectError + 6070, "ANA_TUM_YUK_RAPOR_DOGRULA", "YÜKLENÝCÝ_RAPOR sayfasýnda TOPLAM satýrý bulunamadý."
RaporDeger = ANA_TUM_DOUBLE(ws.Cells(toplamSatiri, 5))
If Abs(RaporDeger - BeklenenDeger) > 0.02 Then Err.Raise vbObjectError + 6071, "ANA_TUM_YUK_RAPOR_DOGRULA", "YÜKLENÝCÝ_RAPOR toplam deðeri YÜKLENÝCÝ_STOK ile uyuþmuyor."
End Sub

Private Function ANA_TUM_LONG(ByVal hucre As Range) As Long
If IsNumeric(hucre.Value2) Then ANA_TUM_LONG = CLng(hucre.Value2) Else ANA_TUM_LONG = 0
End Function

Private Function ANA_TUM_DOUBLE(ByVal hucre As Range) As Double
If IsNumeric(hucre.Value2) Then ANA_TUM_DOUBLE = CDbl(hucre.Value2) Else ANA_TUM_DOUBLE = 0
End Function

Private Function ANA_TUM_GECEN_SURE(ByVal baslangic As Double) As Double
Dim Simdi As Double
Simdi = Timer
If Simdi >= baslangic Then ANA_TUM_GECEN_SURE = Simdi - baslangic Else ANA_TUM_GECEN_SURE = (86400# - baslangic) + Simdi
End Function


Public Sub TUM_STOK_SISTEMINI_GUNCELLE()
    On Error GoTo Hata
    STOKLAR_RAPORLAR_BIRLESTIR_V2
    Exit Sub
Hata:
    AKSORA_HATA_GOSTER "Stok sistemi güncelleme", Err.Number, Err.Description
End Sub

Public Sub ANA_NORMAL_STOK_GUNCELLE_V2()

Dim baslangic As Double
Dim ToplamSure As Double

Dim CDPEskiZaman As Double
Dim EPPEskiZaman As Double

Dim CDPYeniZaman As Double
Dim EPPYeniZaman As Double

Dim CDPSatir As Long
Dim EPPSatir As Long
Dim BirlesikSatir As Long

Dim CDPDeger As Double
Dim EPPDeger As Double
Dim BirlesikDeger As Double

Dim hataNo As Long
Dim hataAciklama As String

If Not AKSORA_SAP_GUNCELLEME_IZNI() Then Exit Sub

On Error GoTo Hata

baslangic = Timer

CORE_DURUM _
"AKSORA normal stok güncellemesi hazýrlanýyor..."

If Not ANA_SAYFA_VAR_MI("CDP_STOK") Then
Err.Raise vbObjectError + 3401, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"CDP_STOK sayfasý bulunamadý."
End If

If Not ANA_SAYFA_VAR_MI("EPP_STOK") Then
Err.Raise vbObjectError + 3402, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"EPP_STOK sayfasý bulunamadý."
End If

If Not ANA_SAYFA_VAR_MI("STOKLAR") Then
Err.Raise vbObjectError + 3403, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"STOKLAR sayfasý bulunamadý."
End If

If Not ANA_SAYFA_VAR_MI("RAPORLAR") Then
Err.Raise vbObjectError + 3404, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"RAPORLAR sayfasý bulunamadý."
End If

CDPEskiZaman = _
ANA_ZAMAN_DEGERI( _
ThisWorkbook.Worksheets("CDP_STOK").Range("C2"))

EPPEskiZaman = _
ANA_ZAMAN_DEGERI( _
ThisWorkbook.Worksheets("EPP_STOK").Range("C2"))

CORE_DURUM _
"1/3 - CDP stoklarý güncelleniyor..."

CDP_STOK_GUNCELLE_V2

CDPYeniZaman = _
ANA_ZAMAN_DEGERI( _
ThisWorkbook.Worksheets("CDP_STOK").Range("C2"))

If CDPYeniZaman <= CDPEskiZaman Then
Err.Raise vbObjectError + 3410, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"CDP güncellemesinin baþarýyla tamamlandýðý doðrulanamadý."
End If

CDPSatir = _
ANA_LONG_DEGER( _
ThisWorkbook.Worksheets("CDP_STOK").Range("C3"))

CDPDeger = _
ANA_DOUBLE_DEGER( _
ThisWorkbook.Worksheets("CDP_STOK").Range("C4"))

If CDPSatir <= 0 Then
Err.Raise vbObjectError + 3411, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"CDP güncellemesi sonrasýnda stok satýrý bulunamadý."
End If

CORE_DURUM _
"2/3 - EPP stoklarý güncelleniyor..."

EPP_STOK_GUNCELLE_V2

EPPYeniZaman = _
ANA_ZAMAN_DEGERI( _
ThisWorkbook.Worksheets("EPP_STOK").Range("C2"))

If EPPYeniZaman <= EPPEskiZaman Then
Err.Raise vbObjectError + 3420, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"EPP güncellemesinin baþarýyla tamamlandýðý doðrulanamadý."
End If

EPPSatir = _
ANA_LONG_DEGER( _
ThisWorkbook.Worksheets("EPP_STOK").Range("C3"))

EPPDeger = _
ANA_DOUBLE_DEGER( _
ThisWorkbook.Worksheets("EPP_STOK").Range("C4"))

If EPPSatir <= 0 Then
Err.Raise vbObjectError + 3421, _
"ANA_NORMAL_STOK_GUNCELLE_V2", _
"EPP güncellemesi sonrasýnda stok satýrý bulunamadý."
End If

CORE_DURUM _
"3/3 - STOKLAR ve RAPORLAR birleþtiriliyor..."

STOKLAR_RAPORLAR_BIRLESTIR_V2

BirlesikSatir = _
ANA_LONG_DEGER( _
ThisWorkbook.Worksheets("STOKLAR").Range("C3"))

BirlesikDeger = _
ANA_DOUBLE_DEGER( _
ThisWorkbook.Worksheets("STOKLAR").Range("C4"))

If BirlesikSatir <> CDPSatir + EPPSatir Then

Err.Raise vbObjectError + 3430, _
          "ANA_NORMAL_STOK_GUNCELLE_V2", _
          "Birleþik stok satýrý doðrulanamadý." & _
          vbCrLf & _
          "CDP + EPP : " & _
          AKSORA_SAYI_METNI(CDPSatir + EPPSatir, "#,##0") & _
          vbCrLf & _
          "STOKLAR   : " & _
          AKSORA_SAYI_METNI(BirlesikSatir, "#,##0")

End If

If Abs(BirlesikDeger - _
(CDPDeger + EPPDeger)) > 0.02 Then

Err.Raise vbObjectError + 3431, _
          "ANA_NORMAL_STOK_GUNCELLE_V2", _
          "Birleþik stok deðeri doðrulanamadý." & _
          vbCrLf & _
          "CDP + EPP : " & _
          AKSORA_SAYI_METNI(CDPDeger + EPPDeger, "#,##0.00") & _
          " TL" & vbCrLf & _
          "STOKLAR   : " & _
          AKSORA_SAYI_METNI(BirlesikDeger, "#,##0.00") & _
          " TL"

End If

ANA_RAPORLAR_DOGRULA _
BirlesikSatir, _
BirlesikDeger

STOK_GORUNUM_TEK "STOKLAR"
RAPOR_GORUNUM_TEK "RAPORLAR"

STOK_HAFIZA_KAYDET

STOK_KARS_TARIHLERI_YUKLE

ToplamSure = _
ANA_GECEN_SURE(baslangic)

CORE_DURUM vbNullString

If Not ANA_SESSIZ_MOD Then

MsgBox _
    "NORMAL STOK V2 zinciri baþarýyla tamamlandý." & _
    vbCrLf & vbCrLf & _
    "CDP" & vbCrLf & _
    "  Satýr : " & AKSORA_SAYI_METNI(CDPSatir, "#,##0") & vbCrLf & _
    "  Deðer : " & AKSORA_SAYI_METNI(CDPDeger, "#,##0.00") & " TL" & _
    vbCrLf & vbCrLf & _
    "EPP" & vbCrLf & _
    "  Satýr : " & AKSORA_SAYI_METNI(EPPSatir, "#,##0") & vbCrLf & _
    "  Deðer : " & AKSORA_SAYI_METNI(EPPDeger, "#,##0.00") & " TL" & _
    vbCrLf & vbCrLf & _
    "BÝRLEÞÝK" & vbCrLf & _
    "  Satýr : " & AKSORA_SAYI_METNI(BirlesikSatir, "#,##0") & vbCrLf & _
    "  Deðer : " & AKSORA_SAYI_METNI(BirlesikDeger, "#,##0.00") & " TL" & _
    vbCrLf & vbCrLf & _
    "RAPORLAR kontrolü : BAÞARILI" & vbCrLf & _
    "Toplam Süre       : " & _
    AKSORA_SAYI_METNI(ToplamSure, "0.00") & " sn", _
    vbInformation, _
    "AKSORA - NORMAL STOK"

End If

Exit Sub

Hata:

hataNo = Err.Number
hataAciklama = Err.Description

CORE_DURUM vbNullString

If Not ANA_SESSIZ_MOD Then

MsgBox _
    "NORMAL STOK V2 zinciri tamamlanamadý." & _
    vbCrLf & vbCrLf & _
    "Hata No : " & hataNo & vbCrLf & _
    "Açýklama: " & hataAciklama & _
    vbCrLf & vbCrLf & _
    "Zincir durduruldu.", _
    vbCritical, _
    "AKSORA - NORMAL STOK"

End If

If ANA_SESSIZ_MOD Then
    Err.Raise hataNo, "ANA_NORMAL_STOK_GUNCELLE_V2", hataAciklama
End If

End Sub

'==========================================================
' SADECE MEVCUT SONUÇLARI TEST ET
'==========================================================
Public Sub ANA_SONUC_TEST_V2()

Dim CDPSatir As Long
Dim EPPSatir As Long
Dim BirlesikSatir As Long

Dim CDPDeger As Double
Dim EPPDeger As Double
Dim BirlesikDeger As Double

Dim ToplamSatir As Long

Dim hataNo As Long
Dim hataAciklama As String

On Error GoTo Hata

If Not ANA_SAYFA_VAR_MI("CDP_STOK") Or _
Not ANA_SAYFA_VAR_MI("EPP_STOK") Or _
Not ANA_SAYFA_VAR_MI("STOKLAR") Or _
Not ANA_SAYFA_VAR_MI("RAPORLAR") Then

Err.Raise vbObjectError + 3450, _
          "ANA_SONUC_TEST_V2", _
          "Test için gerekli sayfalardan biri bulunamadý."

End If

CDPSatir = _
ANA_LONG_DEGER( _
ThisWorkbook.Worksheets("CDP_STOK").Range("C3"))

EPPSatir = _
ANA_LONG_DEGER( _
ThisWorkbook.Worksheets("EPP_STOK").Range("C3"))

BirlesikSatir = _
ANA_LONG_DEGER( _
ThisWorkbook.Worksheets("STOKLAR").Range("C3"))

CDPDeger = _
ANA_DOUBLE_DEGER( _
ThisWorkbook.Worksheets("CDP_STOK").Range("C4"))

EPPDeger = _
ANA_DOUBLE_DEGER( _
ThisWorkbook.Worksheets("EPP_STOK").Range("C4"))

BirlesikDeger = _
ANA_DOUBLE_DEGER( _
ThisWorkbook.Worksheets("STOKLAR").Range("C4"))

ToplamSatir = _
CDPSatir + EPPSatir

If BirlesikSatir <> ToplamSatir Then

Err.Raise vbObjectError + 3451, _
          "ANA_SONUC_TEST_V2", _
          "STOKLAR satýr toplamý CDP + EPP ile uyuþmuyor."

End If

If Abs(BirlesikDeger - _
(CDPDeger + EPPDeger)) > 0.02 Then

Err.Raise vbObjectError + 3452, _
          "ANA_SONUC_TEST_V2", _
          "STOKLAR toplam deðeri CDP + EPP ile uyuþmuyor."

End If

ANA_RAPORLAR_DOGRULA _
BirlesikSatir, _
BirlesikDeger

MsgBox _
"V2 sonuç doðrulama testi BAÞARILI." & _
vbCrLf & vbCrLf & _
"CDP Satýrý : " & _
AKSORA_SAYI_METNI(CDPSatir, "#,##0") & vbCrLf & _
"EPP Satýrý : " & _
AKSORA_SAYI_METNI(EPPSatir, "#,##0") & vbCrLf & _
"Birleþik Satýr : " & _
AKSORA_SAYI_METNI(BirlesikSatir, "#,##0") & vbCrLf & _
"Toplam Deðer : " & _
AKSORA_SAYI_METNI(BirlesikDeger, "#,##0.00") & " TL" & _
vbCrLf & vbCrLf & _
"STOKLAR : DOÐRULANDI" & vbCrLf & _
"RAPORLAR : DOÐRULANDI" & _
vbCrLf & vbCrLf & _
"SAP baðlantýsý kullanýlmadý.", _
vbInformation, _
"AKSORA - TEST"

Exit Sub

Hata:

hataNo = Err.Number
hataAciklama = Err.Description

MsgBox _
"V2 sonuç doðrulama testi BAÞARISIZ." & _
vbCrLf & vbCrLf & _
"Hata No : " & hataNo & vbCrLf & _
"Açýklama: " & hataAciklama, _
vbCritical, _
"AKSORA - TEST"

End Sub

Private Sub ANA_RAPORLAR_DOGRULA( _
ByVal BeklenenSatir As Long, _
ByVal BeklenenDeger As Double)

Dim ws As Worksheet

Dim sonSatir As Long
Dim i As Long

Dim toplamSatiri As Long

Dim RaporSatir As Long
Dim RaporDeger As Double

Set ws = _
ThisWorkbook.Worksheets("RAPORLAR")

sonSatir = _
ws.Cells( _
ws.Rows.Count, _
"A").End(xlUp).Row

If sonSatir < 9 Then

Err.Raise vbObjectError + 3440, _
          "ANA_RAPORLAR_DOGRULA", _
          "RAPORLAR sayfasýnda sonuç bulunamadý."

End If

toplamSatiri = 0

For i = 9 To sonSatir

If UCase$( _
    Trim$( _
        CStr(ws.Cells(i, 1).Value2))) = _
   "TOPLAM" Then

    toplamSatiri = i
    Exit For

End If

Next i

If toplamSatiri = 0 Then

Err.Raise vbObjectError + 3441, _
          "ANA_RAPORLAR_DOGRULA", _
          "RAPORLAR sayfasýnda TOPLAM satýrý bulunamadý."

End If

RaporSatir = _
ANA_LONG_DEGER( _
ws.Cells(toplamSatiri, 3))

RaporDeger = _
ANA_DOUBLE_DEGER( _
ws.Cells(toplamSatiri, 4))

If RaporSatir <> BeklenenSatir Then

Err.Raise vbObjectError + 3442, _
          "ANA_RAPORLAR_DOGRULA", _
          "RAPORLAR toplam stok satýrý STOKLAR ile uyuþmuyor." & _
          vbCrLf & _
          "RAPORLAR : " & _
          AKSORA_SAYI_METNI(RaporSatir, "#,##0") & _
          vbCrLf & _
          "STOKLAR  : " & _
          AKSORA_SAYI_METNI(BeklenenSatir, "#,##0")

End If

If Abs( _
RaporDeger - BeklenenDeger) > 0.02 Then

Err.Raise vbObjectError + 3443, _
          "ANA_RAPORLAR_DOGRULA", _
          "RAPORLAR toplam stok deðeri STOKLAR ile uyuþmuyor." & _
          vbCrLf & _
          "RAPORLAR : " & _
          AKSORA_SAYI_METNI(RaporDeger, "#,##0.00") & _
          " TL" & vbCrLf & _
          "STOKLAR  : " & _
          AKSORA_SAYI_METNI(BeklenenDeger, "#,##0.00") & _
          " TL"

End If

End Sub

Private Function ANA_SAYFA_VAR_MI( _
ByVal sayfaAdi As String) As Boolean

Dim ws As Worksheet

Set ws = Nothing

On Error Resume Next

Set ws = _
ThisWorkbook.Worksheets(sayfaAdi)

On Error GoTo 0

ANA_SAYFA_VAR_MI = _
Not ws Is Nothing

End Function

Private Function ANA_LONG_DEGER( _
ByVal hucre As Range) As Long

If IsNumeric(hucre.Value2) Then

ANA_LONG_DEGER = _
    CLng(hucre.Value2)

Else

ANA_LONG_DEGER = 0

End If

End Function

Private Function ANA_DOUBLE_DEGER( _
ByVal hucre As Range) As Double

If IsNumeric(hucre.Value2) Then

ANA_DOUBLE_DEGER = _
    CDbl(hucre.Value2)

Else

ANA_DOUBLE_DEGER = 0

End If

End Function

Private Function ANA_ZAMAN_DEGERI( _
ByVal hucre As Range) As Double

If IsDate(hucre.Value) Then

ANA_ZAMAN_DEGERI = _
    CDbl(CDate(hucre.Value))

ElseIf IsNumeric(hucre.Value2) Then

ANA_ZAMAN_DEGERI = _
    CDbl(hucre.Value2)

Else

ANA_ZAMAN_DEGERI = 0

End If

End Function

Private Function ANA_GECEN_SURE( _
ByVal baslangic As Double) As Double

Dim Simdi As Double

Simdi = Timer

If Simdi >= baslangic Then

ANA_GECEN_SURE = _
    Simdi - baslangic

Else

ANA_GECEN_SURE = _
    (86400# - baslangic) + Simdi

End If

End Function







