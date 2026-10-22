Attribute VB_Name = "modKS_TEDARIKCI"
Option Explicit

' Çoruh bölgesine 01.01.2026 ve sonrasý teslimat tarihli SAP satýnalma kalemlerinden çýkarýlan
' tedarikçi listesi (kaynak: TEDARÝKÇÝ ÝSÝM.xlsx, "cdp" sayfasý, 47 tedarikçi kodu).
' Ad, SAP'taki yazýmdýr (SAP 35 karakterde keser). Tür, SAP malzeme koduna göre tahmindir:
' 3xxx = hizmet, 1xxx / 7xxx = malzeme. Kontrol edip TANIMLAR sayfasýnda düzeltin.

Private Sub E(ByRef v As Variant, ByRef n As Long, ByVal kod As String, ByVal ad As String, _
        ByVal tur As String, ByVal yer As String, ByVal notu As String)
    n = n + 1
    v(n, 1) = kod & " " & ad
    v(n, 2) = kod
    v(n, 3) = ad
    v(n, 4) = tur
    v(n, 5) = yer
    v(n, 6) = notu
End Sub

Public Function KS_TedarikciVeri(ByRef n As Long) As Variant
    Dim v() As Variant
    ReDim v(1 To 47, 1 To 6)
    n = 0
    E v, n, "1200007200", "FIRAT ELEKTRÝK DAÐITIM A.Þ.", "Kurum Ýçi", "TRABZON", "Grup þirketi / kurum içi transfer"
    E v, n, "1300000429", "AKILCI BILIÞIM ÇÖZÜMLERI VE DAN. A.", "Hizmet", "TRABZON", ""
    E v, n, "1300000622", "GEPA ELEKTRONÝK SANAYÝ VE TÝCARET L", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300000637", "BOYDEM YAPI VE ELEKTRÝK LTD. ÞTÝ.", "Malzeme", "GÜMÜÞHANE, GÝRESUN", ""
    E v, n, "1300000645", "MEKAP DERÝ VE AYAKKABI SANAYÝ TÝC.A", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300000667", "ALTINSOY ENERJÝ MAKÝNE ÝNÞAAT TURÝZ", "Malzeme", "ARTVÝN, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300000704", "REFORM ALTYAPI BILISIM VE TEKN.HIZM", "Hizmet", "TRABZON", "Ayný firma farklý tedarikçi kodu ile de var"
    E v, n, "1300000706", "PÝRAMÝT ÝÞ GÜVENLÝÐÝ MALZEMELERÝ SA", "Malzeme", "GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300000769", "BTC BÝLÝÞÝM HÝZMETLERÝ A.Þ.", "Hizmet", "TRABZON", ""
    E v, n, "1300000892", "NAR SÝSTEM TEKNOLOJÝ ANONÝM ÞÝRKETÝ", "Hizmet", "TRABZON", ""
    E v, n, "1300001214", "BAYLAN ÖLÇÜ ALETLERÝ SAN. VE TÝC.A.", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001244", "DAKKÝK PLASTÝK METAL MAKÝNA VE KALI", "Malzeme", "GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001251", "MAVÝ ALP BÝLGÝ TEKNOLOJÝLERÝ TÝC.LT", "Malzeme", "TRABZON", ""
    E v, n, "1300001281", "ÝÞTEKS ÝÞ ELBÝSELERÝ TEKS ÝNÞ PROM", "Malzeme", "GÜMÜÞHANE", "Ayný firma farklý tedarikçi kodu ile de var"
    E v, n, "1300001282", "MUHAMMED SELMAN KELES", "Malzeme", "TRABZON", ""
    E v, n, "1300001290", "ENTES ELEKTRONÝK CÝHAZLAR AÞ", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001292", "MAPSAN MATBAACILIK LTD.ÞTÝ.", "Malzeme", "TRABZON", ""
    E v, n, "1300001309", "CEMDAÐ AYDINLATMA SAN VE TÝC.A .Þ", "Malzeme", "ARTVÝN, TRABZON", ""
    E v, n, "1300001354", "ÇELÝKLER DEMÝR DÝREK ÝMALAT VE TÝC.", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001363", "ALTERNATÝF BÝLGÝSAYAR YAZ.", "Hizmet", "TRABZON", ""
    E v, n, "1300001389", "HALÝÇ TEKNÝK MAK.SAN.VE DIÞ TÝC.LTD", "Malzeme", "TRABZON", ""
    E v, n, "1300001408", "AVANSAS OFÝS MALZEMELERÝ TÝC.A.Þ.", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001416", "LUNA ELEKTRÝK ELEKTRONÝK", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001425", "FEDERAL ELEKTRÝK YATIRIM VE TÝC. A.", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001440", "DKO KALIP - MEHMET DELÝKEÞOÐLU", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001450", "GRUP ÝMAJ AYDINLAT SAN.TÝC. A.Þ.", "Malzeme", "ARTVÝN, GÜMÜÞHANE, RÝZE, TRABZON", ""
    E v, n, "1300001752", "EAM MÜHENDÝSLÝK VE ELEKTRÝK SÝSTEML", "Hizmet", "TRABZON", ""
    E v, n, "1300001787", "AKTÝF ENERJÝ ÝNÞAAT SAN.TÝC.LTD ÞTÝ", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", ""
    E v, n, "1300001814", "TÜRMAK MAKÝNA SANAYÝ VE TÝCARET A.Þ", "Malzeme", "TRABZON", ""
    E v, n, "1300001893", "ER ELEKTRONÝK SAN.VE TÝC.A.Þ.", "Malzeme", "GÝRESUN", ""
    E v, n, "1300001978", "FONT ELEKTRÝK TÝCARET LTD.ÞTÝ.", "Malzeme", "TRABZON", ""
    E v, n, "1300001983", "NEXA ENERJÝ ELK.ÝTH.ÝHR.SAN.VE TÝC.", "Malzeme", "TRABZON", ""
    E v, n, "1300002105", "GANTEK TEKNOLOJÝ BÝLÝÞÝM ÇÖZÜMLERÝ", "Hizmet", "TRABZON", ""
    E v, n, "1300002328", "EKÝP TEMÝZLÝK VE HÝJYEN ÜRÜNLERÝ TE", "Malzeme", "TRABZON", ""
    E v, n, "1300002610", "TESTONE TEKNOLOJÝ ÇÖZÜML.SAN.VE TÝC", "Malzeme", "TRABZON", ""
    E v, n, "1300002778", "HTMLTEK YAZILIM SÝSTEMLERÝ DANIÞMAN", "Hizmet", "TRABZON", ""
    E v, n, "1300003021", "MÝNA GALVANÝZ SAN.TÝC.A.Þ.", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", "SAP'ta farklý yazým: MÝNA GALVANÝZ SAN.VE TÝC. A.Þ."
    E v, n, "1300003115", "MERFA ENERJI ELEKTRIK SAN. VE TIC.", "Malzeme", "TRABZON", ""
    E v, n, "1300003210", "ATLI OTOMATÝV NAKLÝYAT SAN.TÝC.LTD.", "Malzeme", "GÝRESUN", ""
    E v, n, "1300003265", "ETÝEN ENERJÝ SÝSTEMLERÝ SAN.TÝC.LTD", "Malzeme", "ARTVÝN", ""
    E v, n, "1300003330", "TUÞLA TEKNOLOJI HIZMETLERI A.Þ.", "Hizmet", "TRABZON", ""
    E v, n, "1300003503", "TENSE ELEKTRÝK ELEKTRONÝK SAN. TÝC.", "Malzeme", "GÜMÜÞHANE", ""
    E v, n, "1300003520", "ALTERNANS ELEKTRÝK GÜÇ SÝSTEMLERÝ M", "Hizmet", "TRABZON", ""
    E v, n, "1300003523", "VERÝ BÝLGÝ MERKEZÝ BÝLÝÞÝM HÝZ. A.Þ", "Hizmet", "TRABZON", "Ayný firma farklý tedarikçi kodu ile de var"
    E v, n, "1300009683", "REFORM ALTYAPI BILISIM VE TEKN.HIZM", "Hizmet", "TRABZON", "Ayný firma farklý tedarikçi kodu ile de var"
    E v, n, "1300023370", "ÝÞTEKS ÝÞ ELBÝSELERÝ TEKSTÝL ÝNÞ.", "Malzeme", "ARTVÝN, GÜMÜÞHANE, GÝRESUN, RÝZE, TRABZON", "Ayný firma farklý tedarikçi kodu ile de var"
    E v, n, "1300026317", "VERÝ BÝLGÝ MERKEZÝ BÝLÝÞÝM HÝZ. A.Þ", "Hizmet", "TRABZON", "Ayný firma farklý tedarikçi kodu ile de var"
    KS_TedarikciVeri = v
End Function

' Listede olup tabloda olmayan tedarikçileri (koda göre) TANIMLAR sayfasýna ekler. Mevcutlara dokunmaz.
Public Function KS_TedarikciEkle() As Long
    Dim lo As ListObject, v As Variant, n As Long, i As Long, j As Long, d As Object
    Dim yeni() As Variant, m As Long, korumali As Boolean
    Set lo = KS_LOZ(T_SEVKCI)
    Set d = KS_Sozluk(T_SEVKCI, 2)
    v = KS_TedarikciVeri(n)
    ReDim yeni(1 To n, 1 To lo.ListColumns.Count)
    For i = 1 To n
        If Not d.Exists(KS_Anahtar(v(i, 2))) Then
            m = m + 1
            For j = 1 To 6
                yeni(m, j) = v(i, j)
            Next j
        End If
    Next i
    If m > 0 Then
        korumali = lo.Parent.ProtectContents
        KS_Ac lo.Parent
        KS_TopluEkle lo, yeni, m
        If korumali Then KS_Kapat lo.Parent
    End If
    KS_TedarikciEkle = m
End Function

Public Sub KS_TEDARIKCI_YUKLE()
    Dim m As Long
    On Error GoTo Hata
    m = KS_TedarikciEkle()
    KS_Log "TEDARÝKÇÝ YÜKLE", "", m & " tedarikçi eklendi"
    MsgBox m & " tedarikçi TANIMLAR sayfasýna eklendi." & vbCrLf & _
        "Kodu zaten kayýtlý olanlar deðiþtirilmedi.", vbInformation, KS_AD
    Exit Sub
Hata:
    MsgBox "Tedarikçi listesi yüklenemedi: " & Err.Description, vbExclamation, KS_AD
End Sub
