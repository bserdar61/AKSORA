Attribute VB_Name = "modKS_TEDARIKCI"
Option Explicit

' Çoruh bölgesine 01.01.2026 ve sonrası teslimat tarihli SAP satınalma kalemlerinden çıkarılan
' tedarikçi listesi (kaynak: TEDARİKÇİ İSİM.xlsx, "cdp" sayfası, 47 tedarikçi kodu).
' Ad, SAP'taki yazımdır (SAP 35 karakterde keser). Tür, SAP malzeme koduna göre tahmindir:
' 3xxx = hizmet, 1xxx / 7xxx = malzeme. Kontrol edip TANIMLAR sayfasında düzeltin.

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
    E v, n, "1200007200", "FIRAT ELEKTRİK DAĞITIM A.Ş.", "Kurum İçi", "TRABZON", "Grup şirketi / kurum içi transfer"
    E v, n, "1300000429", "AKILCI BILIŞIM ÇÖZÜMLERI VE DAN. A.", "Hizmet", "TRABZON", ""
    E v, n, "1300000622", "GEPA ELEKTRONİK SANAYİ VE TİCARET L", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300000637", "BOYDEM YAPI VE ELEKTRİK LTD. ŞTİ.", "Malzeme", "GÜMÜŞHANE, GİRESUN", ""
    E v, n, "1300000645", "MEKAP DERİ VE AYAKKABI SANAYİ TİC.A", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300000667", "ALTINSOY ENERJİ MAKİNE İNŞAAT TURİZ", "Malzeme", "ARTVİN, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300000704", "REFORM ALTYAPI BILISIM VE TEKN.HIZM", "Hizmet", "TRABZON", "Aynı firma farklı tedarikçi kodu ile de var"
    E v, n, "1300000706", "PİRAMİT İŞ GÜVENLİĞİ MALZEMELERİ SA", "Malzeme", "GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300000769", "BTC BİLİŞİM HİZMETLERİ A.Ş.", "Hizmet", "TRABZON", ""
    E v, n, "1300000892", "NAR SİSTEM TEKNOLOJİ ANONİM ŞİRKETİ", "Hizmet", "TRABZON", ""
    E v, n, "1300001214", "BAYLAN ÖLÇÜ ALETLERİ SAN. VE TİC.A.", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001244", "DAKKİK PLASTİK METAL MAKİNA VE KALI", "Malzeme", "GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001251", "MAVİ ALP BİLGİ TEKNOLOJİLERİ TİC.LT", "Malzeme", "TRABZON", ""
    E v, n, "1300001281", "İŞTEKS İŞ ELBİSELERİ TEKS İNŞ PROM", "Malzeme", "GÜMÜŞHANE", "Aynı firma farklı tedarikçi kodu ile de var"
    E v, n, "1300001282", "MUHAMMED SELMAN KELES", "Malzeme", "TRABZON", ""
    E v, n, "1300001290", "ENTES ELEKTRONİK CİHAZLAR AŞ", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001292", "MAPSAN MATBAACILIK LTD.ŞTİ.", "Malzeme", "TRABZON", ""
    E v, n, "1300001309", "CEMDAĞ AYDINLATMA SAN VE TİC.A .Ş", "Malzeme", "ARTVİN, TRABZON", ""
    E v, n, "1300001354", "ÇELİKLER DEMİR DİREK İMALAT VE TİC.", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001363", "ALTERNATİF BİLGİSAYAR YAZ.", "Hizmet", "TRABZON", ""
    E v, n, "1300001389", "HALİÇ TEKNİK MAK.SAN.VE DIŞ TİC.LTD", "Malzeme", "TRABZON", ""
    E v, n, "1300001408", "AVANSAS OFİS MALZEMELERİ TİC.A.Ş.", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001416", "LUNA ELEKTRİK ELEKTRONİK", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001425", "FEDERAL ELEKTRİK YATIRIM VE TİC. A.", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001440", "DKO KALIP - MEHMET DELİKEŞOĞLU", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001450", "GRUP İMAJ AYDINLAT SAN.TİC. A.Ş.", "Malzeme", "ARTVİN, GÜMÜŞHANE, RİZE, TRABZON", ""
    E v, n, "1300001752", "EAM MÜHENDİSLİK VE ELEKTRİK SİSTEML", "Hizmet", "TRABZON", ""
    E v, n, "1300001787", "AKTİF ENERJİ İNŞAAT SAN.TİC.LTD ŞTİ", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", ""
    E v, n, "1300001814", "TÜRMAK MAKİNA SANAYİ VE TİCARET A.Ş", "Malzeme", "TRABZON", ""
    E v, n, "1300001893", "ER ELEKTRONİK SAN.VE TİC.A.Ş.", "Malzeme", "GİRESUN", ""
    E v, n, "1300001978", "FONT ELEKTRİK TİCARET LTD.ŞTİ.", "Malzeme", "TRABZON", ""
    E v, n, "1300001983", "NEXA ENERJİ ELK.İTH.İHR.SAN.VE TİC.", "Malzeme", "TRABZON", ""
    E v, n, "1300002105", "GANTEK TEKNOLOJİ BİLİŞİM ÇÖZÜMLERİ", "Hizmet", "TRABZON", ""
    E v, n, "1300002328", "EKİP TEMİZLİK VE HİJYEN ÜRÜNLERİ TE", "Malzeme", "TRABZON", ""
    E v, n, "1300002610", "TESTONE TEKNOLOJİ ÇÖZÜML.SAN.VE TİC", "Malzeme", "TRABZON", ""
    E v, n, "1300002778", "HTMLTEK YAZILIM SİSTEMLERİ DANIŞMAN", "Hizmet", "TRABZON", ""
    E v, n, "1300003021", "MİNA GALVANİZ SAN.TİC.A.Ş.", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", "SAP'ta farklı yazım: MİNA GALVANİZ SAN.VE TİC. A.Ş."
    E v, n, "1300003115", "MERFA ENERJI ELEKTRIK SAN. VE TIC.", "Malzeme", "TRABZON", ""
    E v, n, "1300003210", "ATLI OTOMATİV NAKLİYAT SAN.TİC.LTD.", "Malzeme", "GİRESUN", ""
    E v, n, "1300003265", "ETİEN ENERJİ SİSTEMLERİ SAN.TİC.LTD", "Malzeme", "ARTVİN", ""
    E v, n, "1300003330", "TUŞLA TEKNOLOJI HIZMETLERI A.Ş.", "Hizmet", "TRABZON", ""
    E v, n, "1300003503", "TENSE ELEKTRİK ELEKTRONİK SAN. TİC.", "Malzeme", "GÜMÜŞHANE", ""
    E v, n, "1300003520", "ALTERNANS ELEKTRİK GÜÇ SİSTEMLERİ M", "Hizmet", "TRABZON", ""
    E v, n, "1300003523", "VERİ BİLGİ MERKEZİ BİLİŞİM HİZ. A.Ş", "Hizmet", "TRABZON", "Aynı firma farklı tedarikçi kodu ile de var"
    E v, n, "1300009683", "REFORM ALTYAPI BILISIM VE TEKN.HIZM", "Hizmet", "TRABZON", "Aynı firma farklı tedarikçi kodu ile de var"
    E v, n, "1300023370", "İŞTEKS İŞ ELBİSELERİ TEKSTİL İNŞ.", "Malzeme", "ARTVİN, GÜMÜŞHANE, GİRESUN, RİZE, TRABZON", "Aynı firma farklı tedarikçi kodu ile de var"
    E v, n, "1300026317", "VERİ BİLGİ MERKEZİ BİLİŞİM HİZ. A.Ş", "Hizmet", "TRABZON", "Aynı firma farklı tedarikçi kodu ile de var"
    KS_TedarikciVeri = v
End Function

' Listede olup tabloda olmayan tedarikçileri (koda göre) TANIMLAR sayfasına ekler. Mevcutlara dokunmaz.
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
    KS_Log "TEDARİKÇİ YÜKLE", "", m & " tedarikçi eklendi"
    MsgBox m & " tedarikçi TANIMLAR sayfasına eklendi." & vbCrLf & _
        "Kodu zaten kayıtlı olanlar değiştirilmedi.", vbInformation, KS_AD
    Exit Sub
Hata:
    MsgBox "Tedarikçi listesi yüklenemedi: " & Err.Description, vbExclamation, KS_AD
End Sub
