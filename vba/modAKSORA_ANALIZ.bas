Attribute VB_Name = "modAKSORA_ANALIZ"
Option Explicit

' Malzeme Stok birimi analizleri: ATIL STOK, SATINALMA KONTROLU, KRITIK STOK (modAKSORA_BIRIM'den ayrildi; buyuk modul riski).
' Ortak yardimcilar (Y_*, ADM_*, I_* ) modAKSORA_BIRIM'de Public.
Private Const ADM As String = "ADMIN PANEL"
Private Const SA_KAYIT As String = "_SATINALMA"
Private Const SA_SAYFA As String = "SATINALMA_KONTROL"
Private Const SA_ETIKET As String = "AKSORA_SATINALMA_V1"
Private Const KS_SAYFA As String = "KRITIK_STOK"
Private Const KS_ETIKET As String = "AKSORA_KRITIK_V1"
Private Const ATIL_SAYFA As String = "ATIL_STOK"
Private Const ATIL_KAYIT As String = "_ATIL"
Private Const ATIL_ETIKET As String = "AKSORA_ATIL_V1"
Private Declare PtrSafe Function ADMW_CoRegisterMessageFilter Lib "ole32" Alias "CoRegisterMessageFilter" (ByVal lFilterIn As LongPtr, ByRef lPreviousFilter As LongPtr) As Long
' Analiz onbellegi: ayni islem icinde (sayfa / Excel / Yonetim Merkezi) stok ve atil hesabi bir kez yapilir.
' Yalniz AN_ONB_AC .. AN_ONB_KAPAT arasinda ve en fazla 120 sn gecerli (hata ile acik kalsa da bayat veri kullanilmaz).
Private mOnbDerinlik As Long
Private mOnbZaman As Double
Private mOnbTum As Object
Private mOnbHes As Collection
Private mOnbBilgi As Variant
Private Const AN_ESKI_GUN As Long = 7
Private Const TR_SAYFA As String = "TRANSFER_LISTESI"
Private mAdmIsler As Collection
Private Const TR_ETIKET As String = "AKSORA_TRANSFER_V1"
Private Declare PtrSafe Function ATIL_OpenClipboard Lib "user32" Alias "OpenClipboard" (ByVal hWnd As LongPtr) As Long
Private Declare PtrSafe Function ATIL_EmptyClipboard Lib "user32" Alias "EmptyClipboard" () As Long
Private Declare PtrSafe Function ATIL_CloseClipboard Lib "user32" Alias "CloseClipboard" () As Long
Private Declare PtrSafe Function ATIL_SetClipboardData Lib "user32" Alias "SetClipboardData" (ByVal wFormat As Long, ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function ATIL_GlobalAlloc Lib "kernel32" Alias "GlobalAlloc" (ByVal wFlags As Long, ByVal dwBytes As LongPtr) As LongPtr
Private Declare PtrSafe Function ATIL_GlobalLock Lib "kernel32" Alias "GlobalLock" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Function ATIL_GlobalUnlock Lib "kernel32" Alias "GlobalUnlock" (ByVal hMem As LongPtr) As Long
Private Declare PtrSafe Function ATIL_GlobalFree Lib "kernel32" Alias "GlobalFree" (ByVal hMem As LongPtr) As LongPtr
Private Declare PtrSafe Sub ATIL_CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByVal hedef As LongPtr, ByVal kaynak As LongPtr, ByVal adet As LongPtr)
Private Declare PtrSafe Sub ATIL_Sleep Lib "kernel32" Alias "Sleep" (ByVal ms As Long)
Private mAtilKart As Boolean
Private mAtilBas As Double
Private mAtilSon As Double
Private mAtilMetin As String
Private mAtilYuk As Object
Private mAtilEski As Boolean

'==========================================================
' ATIL STOK (Malzeme Stok Birimi)
' Atil (uretim yeri bazinda): eldeki stok FIFO ile girislere dagitilir (en yeni giris once kalir);
' 24 aydan eski giristen kalan kisim ATIL, 21-24 ay "3 AY KALDI", 18-21 ay "6 AY KALDI".
' Giris: 101/105/501/561 ve uretim yerleri arasi transferin alan tarafi (iptalleriyle). Ayni uretim yerindeki
' depo transferi (311) yasi sifirlamaz; yuklenici sarfi stogu azaltir, kalan eski stok atil kalir.
' Kapsam: illerin CDP ve EPP D001-D999, K001-K999 ve yuklenici depolari (uretim yeri toplami);
'         HURDA ve CIKMA/BDSZ partiler haric.
' SAP (MB51, /ATIL duzeni) arada bir okunur; sayac SAP'siz isler. Sonraki stok verilerinde artis yeni giris katmani olur.
' _ATIL (ATIL2): 1. satir etiket / okuma tarihleri (D1 CDP, E1 EPP), 2. satir basliklar, 3. satirdan:
'   sistem | UY | malzeme | girisler (yyyymmdd=miktar;...) | 24-60 ay onceki son giris | son 6 ay sarf | son sarf |
'   SAP okumasindaki stok | son karsilastirilan stok | son karsilastirilan stok verisinin tarihi
'==========================================================
Private Function ATIL_IL_BUYUK(ByVal il As String) As String
    Select Case il
        Case "ARTVIN": ATIL_IL_BUYUK = "ARTVÝN"
        Case "GUMUSHANE": ATIL_IL_BUYUK = "GÜMÜÞHANE"
        Case "GIRESUN": ATIL_IL_BUYUK = "GÝRESUN"
        Case "RIZE": ATIL_IL_BUYUK = "RÝZE"
        Case Else: ATIL_IL_BUYUK = il
    End Select
End Function

Private Function ATIL_SISTEM(ByVal uy As String) As String
    Select Case Left$(uy, 3)
        Case "715": ATIL_SISTEM = "CDP"
        Case "705": ATIL_SISTEM = "EPP"
    End Select
End Function

Public Function ATIL_TESIS_IL(ByVal uy As String) As String
    Dim il As Variant
    For Each il In ADM_ILLER()
        If AKSORA_IL_TESIS(CStr(il), "CDP") = uy Or AKSORA_IL_TESIS(CStr(il), "EPP") = uy Then ATIL_TESIS_IL = CStr(il): Exit Function
    Next il
End Function

' D001-D999 ve K001-K999
Private Function ATIL_DEPO_MI(ByVal depo As String) As Boolean
    depo = UCase$(Trim$(depo))
    If Len(depo) <> 4 Then Exit Function
    If Left$(depo, 1) <> "D" And Left$(depo, 1) <> "K" Then Exit Function
    If Not (Mid$(depo, 2) Like "###") Then Exit Function
    ATIL_DEPO_MI = (Mid$(depo, 2) <> "000")
End Function

Private Function ATIL_PARTI_DISI(ByVal parti As String) As Boolean
    Dim t As String
    t = Y_NORM(parti)
    ATIL_PARTI_DISI = (InStr(t, "HURDA") > 0 Or (InStr(t, "CIKMA") > 0 And (InStr(t, "BDSZ") > 0 Or InStr(t, "BEDELSIZ") > 0)))
End Function

Private Function ATIL_KATEGORI(ByVal mg As String, ByVal tn As String) As String
    mg = UCase$(Trim$(mg)): tn = Y_NORM(tn)
    If Left$(tn, 6) = "IPTAL." Then tn = Trim$(Mid$(tn, 7))
    Select Case True
        Case Left$(mg, 5) = "M0316": ATIL_KATEGORI = "KABLO"
        Case Left$(mg, 5) = "M0308": ATIL_KATEGORI = "DÝREK"
        Case Left$(mg, 5) = "M0327": ATIL_KATEGORI = "PANO"
        Case Left$(mg, 5) = "M0314": ATIL_KATEGORI = "HÜCRE"
        Case Left$(mg, 5) = "M0340": ATIL_KATEGORI = "TRAFO"
        Case Len(mg) > 0: ATIL_KATEGORI = "DÝÐER"
        Case Left$(tn, 6) = "KABLO ": ATIL_KATEGORI = "KABLO"
        Case Left$(tn, 6) = "DIREK ": ATIL_KATEGORI = "DÝREK"
        Case Left$(tn, 5) = "PANO ": ATIL_KATEGORI = "PANO"
        Case Left$(tn, 6) = "HUCRE ": ATIL_KATEGORI = "HÜCRE"
        Case Left$(tn, 6) = "TRAFO ": ATIL_KATEGORI = "TRAFO"
        Case Else: ATIL_KATEGORI = "DÝÐER"
    End Select
End Function

Private Function ATIL_GECEN(ByVal bas As Double) As Double
    ATIL_GECEN = Timer - bas
    If ATIL_GECEN < 0 Then ATIL_GECEN = ATIL_GECEN + 86400#
End Function

Private Function ATIL_HUCRE_SAYI(ByVal v As Variant) As Double
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then ATIL_HUCRE_SAYI = CDbl(v)
End Function

'---------- ATIL STOK: stok verisi ----------
' Ilin kapsamdaki stogu (D/K depolari ve yuklenici depolari; malzeme birimi stok sayfasi ya da il kaydi):
'   UY|malzeme -> Array(0 miktar, 1 deger, 2 birim, 3 tanim, 4 kategori, 5 depolar: depo -> Array(ad, miktar, deger))
' tarih: stok verisinin tarihi (0: veri yok)
Private Function ATIL_STOK_OKU(ByVal il As String, ByRef tarih As Date) As Object
    Dim d As Object, gorulen As Object, kaynaklar As Variant, kx As Variant, sh As Worksheet, canli As Boolean, t As Variant
    Dim v As Variant, son As Long, i As Long, uy As String, depo As String, mal As String, k As String, x As Double, dg As Double
    Dim o As Variant, dp As Object, ka As String, sira As Long, p As Variant, tMin As Date
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set gorulen = CreateObject("Scripting.Dictionary"): gorulen.CompareMode = vbTextCompare
    tarih = 0
    canli = (il = ADM_CANLI_IL(0))
    If canli Then
        kaynaklar = Array("CDP_STOK", "EPP_STOK", "ÝLÇE_STOK", "YÜKLENÝCÝ_STOK")
    Else
        kaynaklar = Array("_AKLIVE_" & il & "_CDP", "_AKLIVE_" & il & "_EPP", "_AKLIVE_" & il & "_ILC", "_AKLIVE_" & il & "_YUK")
    End If
    For Each kx In kaynaklar
        sira = sira + 1
        Set sh = ADM_WS(CStr(kx))
        If Not sh Is Nothing Then
            If canli Then t = sh.Range("C2").Value Else t = sh.Range("A3").Value
            son = AKSORA_SON_SATIR(sh, 7)
            If Not IsError(t) Then
                If IsDate(t) Then
                    If CDate(t) > tarih Then tarih = CDate(t)
                    If son >= 7 Then If tMin = 0 Or CDate(t) < tMin Then tMin = CDate(t)
                End If
            End If
            v = Empty
            If son >= 7 Then v = sh.Range(sh.Cells(7, 1), sh.Cells(son, 14)).Value2
            If IsArray(v) Then
                For i = 1 To UBound(v, 1)
                    If (i And 2047) = 0 Then AKSORA_BEKLE_NABIZ
                    mal = Y_KOD_NORM(v(i, 7))
                    depo = UCase$(I_TEMIZ(v(i, 2)))
                    uy = I_TEMIZ(v(i, 1))
                    If Len(mal) > 0 And (ATIL_DEPO_MI(depo) Or (sira = 4 And Len(depo) > 0)) And Len(ATIL_SISTEM(uy)) > 0 Then
                        If ATIL_TESIS_IL(uy) = il And Not ATIL_PARTI_DISI(I_TEMIZ(v(i, 6))) Then
                            If sira = 4 And Not mAtilYuk Is Nothing Then
                                If Not mAtilYuk.Exists(uy) Then mAtilYuk.Add uy, CreateObject("Scripting.Dictionary")
                                mAtilYuk(uy)(depo) = True
                            End If
                            x = I_SAYI(v(i, 9))
                            ' Ayni satir iki sayfada varsa (eski kayit duzeni) bir kez sayilir.
                            ka = uy & "|" & depo & "|" & mal & "|" & UCase$(I_TEMIZ(v(i, 6)))
                            If gorulen.Exists(ka) Then
                                If gorulen(ka) <> sira Then x = 0
                            Else
                                gorulen.Add ka, sira
                            End If
                            If x > 0 Then
                                dg = I_SAYI(v(i, 11))
                                k = uy & "|" & mal
                                If Not d.Exists(k) Then
                                    Set dp = CreateObject("Scripting.Dictionary"): dp.CompareMode = vbTextCompare
                                    d.Add k, Array(0#, 0#, I_TEMIZ(v(i, 10)), I_TEMIZ(v(i, 8)), ATIL_KATEGORI(I_TEMIZ(v(i, 4)), I_TEMIZ(v(i, 8))), dp)
                                End If
                                o = d(k): o(0) = o(0) + x: o(1) = o(1) + dg: d(k) = o
                                Set dp = o(5)
                                If Not dp.Exists(depo) Then dp.Add depo, Array(I_TEMIZ(v(i, 3)), 0#, 0#)
                                p = dp(depo): p(1) = p(1) + x: p(2) = p(2) + dg: dp(depo) = p
                            End If
                        End If
                    End If
                Next i
            End If
        End If
    Next kx
    ' Ilin stok sayfalari ayri zamanlarda yenilendiyse (yalniz biri guncellendi) karsilastirma yapilmaz:
    ' depolar arasi transfer sahte "yeni giris" katmani olusturup eski atil stogu yeni gosterirdi.
    ' Tarih en eski kaynagin tarihi olur; butun sayfalar yenilenince karsilastirma yapilir.
    If tMin > 0 And tarih - tMin > TimeSerial(1, 0, 0) Then tarih = tMin
    Set ATIL_STOK_OKU = d
End Function

' il -> Array(stok sozlugu, stok tarihi)
Private Function ATIL_TUM_STOK() As Object
    Dim d As Object, il As Variant, t As Date, s As Object
    If AN_ONB_GECERLI() Then
        If Not mOnbTum Is Nothing Then Set ATIL_TUM_STOK = mOnbTum: Exit Function
    End If
    Set d = CreateObject("Scripting.Dictionary")
    Set mAtilYuk = CreateObject("Scripting.Dictionary"): mAtilYuk.CompareMode = vbTextCompare
    For Each il In ADM_ILLER()
        AKSORA_BEKLE_NABIZ
        Set s = ATIL_STOK_OKU(CStr(il), t)
        d.Add CStr(il), Array(s, t)
    Next il
    Set ATIL_TUM_STOK = d
    If AN_ONB_GECERLI() Then Set mOnbTum = d
End Function

'---------- ATIL STOK: kayit ----------
' UY|malzeme -> Array(0 sistem, 1 girisler, 2 eski giris, 3 son 6 ay sarf, 4 son sarf, 5 SAP stogu, 6 son stok, 7 son stok tarihi,
'   8 aylik sarf "yyyymm=miktar;..." (eski kayitta bos))
' okuma: Array(CDP okuma tarihi, EPP okuma tarihi) (0: okunmadi). Eski kural kaydi (ATIL1) okunmaz.
Private Function ATIL_KAYIT_OKU(ByRef okuma As Variant) As Object
    Dim d As Object, ws As Worksheet, son As Long, v As Variant, r As Long, k As String
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    okuma = Array(0#, 0#)
    mAtilEski = False
    Set ATIL_KAYIT_OKU = d
    Set ws = ADM_WS(ATIL_KAYIT)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Range("A1").Value2) = "ATIL1" Then mAtilEski = True
    If CStr(ws.Range("A1").Value2) <> "ATIL2" Then Exit Function
    okuma = Array(ATIL_HUCRE_SAYI(ws.Range("D1").Value2), ATIL_HUCRE_SAYI(ws.Range("E1").Value2))
    son = AKSORA_SON_SATIR(ws, 2)
    If son < 3 Then Exit Function
    v = ws.Range("A3:K" & son).Value2
    For r = 1 To UBound(v, 1)
        k = CStr(v(r, 2)) & "|" & CStr(v(r, 3))
        If Len(CStr(v(r, 2))) > 0 And Len(CStr(v(r, 3))) > 0 Then
            d(k) = Array(CStr(v(r, 1)), CStr(v(r, 4)), ATIL_HUCRE_SAYI(v(r, 5)), ATIL_HUCRE_SAYI(v(r, 6)), ATIL_HUCRE_SAYI(v(r, 7)), _
                ATIL_HUCRE_SAYI(v(r, 8)), ATIL_HUCRE_SAYI(v(r, 9)), ATIL_HUCRE_SAYI(v(r, 10)), CStr(v(r, 11)))
        End If
    Next r
End Function

' Yeni stok verisi gelen illerde: stok onceki miktardan fazlaysa fark yeni giris sayilir (stok tarihiyle katman eklenir);
' azalis FIFO ile en eski stoktan duser. SAP okumasindan sonra stoga giren yeni malzemeler de eklenir. Donus: azalan kalem.
Private Function ATIL_KARSILASTIR(ByVal tum As Object) As Long
    Dim ws As Worksheet, son As Long, v As Variant, r As Long, il As Variant, uy As String, k As Variant, s As Object
    Dim cq As Double, td As Double, degisti As Boolean, n As Long, var As Object, yeni As Collection, a() As Variant, c As Long
    Dim ok0 As Double, ok1 As Double, sis As String, x As Variant
    Set ws = ADM_WS(ATIL_KAYIT)
    If ws Is Nothing Then Exit Function
    If CStr(ws.Range("A1").Value2) <> "ATIL2" Then Exit Function
    ok0 = ATIL_HUCRE_SAYI(ws.Range("D1").Value2): ok1 = ATIL_HUCRE_SAYI(ws.Range("E1").Value2)
    son = AKSORA_SON_SATIR(ws, 2)
    Set var = CreateObject("Scripting.Dictionary"): var.CompareMode = vbTextCompare
    If son >= 3 Then
        v = ws.Range("A3:J" & son).Value2
        For r = 1 To UBound(v, 1)
            uy = CStr(v(r, 2))
            var(uy & "|" & CStr(v(r, 3))) = True
            il = ATIL_TESIS_IL(uy)
            If Len(il) > 0 And Len(CStr(v(r, 3))) > 0 Then
                If tum.Exists(CStr(il)) Then
                    td = CDbl(tum(CStr(il))(1))
                    If td > 0 And td > ATIL_HUCRE_SAYI(v(r, 10)) + 0.00001 Then
                        Set s = tum(CStr(il))(0)
                        cq = 0
                        If s.Exists(uy & "|" & CStr(v(r, 3))) Then cq = s(uy & "|" & CStr(v(r, 3)))(0)
                        If cq > ATIL_HUCRE_SAYI(v(r, 9)) + 0.0005 Then
                            v(r, 4) = ATIL_KATMAN_EKLE(CStr(v(r, 4)), td, cq - ATIL_HUCRE_SAYI(v(r, 9)))
                        ElseIf cq < ATIL_HUCRE_SAYI(v(r, 9)) - 0.0005 Then
                            n = n + 1
                        End If
                        v(r, 9) = cq: v(r, 10) = td
                        degisti = True
                    End If
                End If
            End If
        Next r
    End If
    ' SAP okumasindan sonra stoga giren malzemeler: tamami yeni giris.
    Set yeni = New Collection
    For Each il In ADM_ILLER()
        If tum.Exists(CStr(il)) Then
            td = CDbl(tum(CStr(il))(1))
            Set s = tum(CStr(il))(0)
            For Each k In s.keys
                uy = Left$(CStr(k), InStr(CStr(k), "|") - 1)
                sis = ATIL_SISTEM(uy)
                If Not var.Exists(CStr(k)) And td > 0 And ((sis = "CDP" And ok0 > 0) Or (sis = "EPP" And ok1 > 0)) Then
                    yeni.Add Array(sis, uy, Mid$(CStr(k), InStr(CStr(k), "|") + 1), ATIL_KATMAN_EKLE(vbNullString, td, s(k)(0)), 0#, 0#, 0#, 0#, s(k)(0), td)
                End If
            Next k
        End If
    Next il
    If degisti Then ws.Range("A3:J" & son).Value2 = v
    If yeni.Count > 0 Then
        ReDim a(1 To yeni.Count, 1 To 10)
        r = 0
        For Each x In yeni
            r = r + 1
            For c = 1 To 10: a(r, c) = x(c - 1): Next c
        Next x
        ws.Range("A" & application.Max(3, son + 1)).Resize(yeni.Count, 10).Value2 = a
        ws.Range("C1").Value2 = application.Max(3, son + 1) + yeni.Count - 3
    End If
    If degisti Or yeni.Count > 0 Then AKSORA_SAP_TESHIS_YAZ "Atýl stok karþýlaþtýrma: " & n & " kalemde stok düþtü, " & yeni.Count & " yeni kalem"
    ATIL_KARSILASTIR = n
End Function

' Malzeme stoklari guncellenince cagrilir (kayit yoksa hicbir sey yapmaz).
Public Sub AKSORA_ATIL_STOK_KONTROL()
    Dim ac As Boolean
    On Error Resume Next
    If ADM_WS(ATIL_KAYIT) Is Nothing Then Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Stok deðiþimleri atýl kaydýna iþleniyor...", 0.3)
    ATIL_KARSILASTIR ATIL_TUM_STOK()
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
End Sub

Private Sub ATIL_KAYDET(ByVal sonuc As Object, ByVal tum As Object, ByVal oturumlar As Object)
    Dim ws As Worksheet, son As Long, eski As Variant, r As Long, c As Long, n As Long, a() As Variant, il As Variant, sis As Variant
    Dim uy As String, s As Object, k As Variant, o As Variant, td As Double, yazildi As Object, kap As Long, ok As Variant
    Set ws = Y_KAYIT_SAYFA(ATIL_KAYIT)
    ok = Array(0#, 0#)
    If CStr(ws.Range("A1").Value2) = "ATIL2" Then ok = Array(ATIL_HUCRE_SAYI(ws.Range("D1").Value2), ATIL_HUCRE_SAYI(ws.Range("E1").Value2))
    son = AKSORA_SON_SATIR(ws, 2)
    kap = application.Max(son, 3) + sonuc.Count
    For Each il In ADM_ILLER()
        kap = kap + tum(CStr(il))(0).Count
    Next il
    ReDim a(1 To kap, 1 To 11)
    ' Bu okumada sorgulanmayan sistemin satirlari korunur (yalniz yeni kural kaydi).
    If son >= 3 And CStr(ws.Range("A1").Value2) = "ATIL2" Then
        eski = ws.Range("A3:K" & son).Value2
        For r = 1 To UBound(eski, 1)
            If Len(CStr(eski(r, 1))) > 0 And Not oturumlar.Exists(CStr(eski(r, 1))) Then
                n = n + 1
                For c = 1 To 11: a(n, c) = eski(r, c): Next c
            End If
        Next r
    End If
    Set yazildi = CreateObject("Scripting.Dictionary"): yazildi.CompareMode = vbTextCompare
    For Each sis In oturumlar.keys
        For Each il In ADM_ILLER()
            uy = AKSORA_IL_TESIS(CStr(il), CStr(sis))
            Set s = tum(CStr(il))(0): td = CDbl(tum(CStr(il))(1))
            For Each k In s.keys
                If Left$(CStr(k), Len(uy) + 1) = uy & "|" Then
                    n = n + 1
                    a(n, 1) = CStr(sis): a(n, 2) = uy: a(n, 3) = Mid$(CStr(k), Len(uy) + 2)
                    If sonuc.Exists(CStr(k)) Then
                        o = sonuc(CStr(k))
                        a(n, 4) = CStr(o(0)): a(n, 5) = CDbl(o(1)): a(n, 6) = CDbl(o(2)): a(n, 7) = CDbl(o(3))
                        If UBound(o) >= 4 Then a(n, 11) = CStr(o(4))
                    Else
                        a(n, 4) = vbNullString: a(n, 5) = 0#: a(n, 6) = 0#: a(n, 7) = 0#
                    End If
                    a(n, 8) = s(k)(0): a(n, 9) = s(k)(0): a(n, 10) = td
                    yazildi(CStr(k)) = True
                End If
            Next k
            ' Stogu olmayan ama son 6 ayda sarf eden uretim yeri (transfer onerisi icin talep).
            For Each k In sonuc.keys
                If Left$(CStr(k), Len(uy) + 1) = uy & "|" And Not yazildi.Exists(CStr(k)) Then
                    o = sonuc(CStr(k))
                    n = n + 1
                    a(n, 1) = CStr(sis): a(n, 2) = uy: a(n, 3) = Mid$(CStr(k), Len(uy) + 2)
                    a(n, 4) = CStr(o(0)): a(n, 5) = CDbl(o(1)): a(n, 6) = CDbl(o(2)): a(n, 7) = CDbl(o(3))
                    If UBound(o) >= 4 Then a(n, 11) = CStr(o(4))
                    a(n, 8) = 0#: a(n, 9) = 0#: a(n, 10) = td
                    yazildi(CStr(k)) = True
                End If
            Next k
        Next il
        If CStr(sis) = "CDP" Then ok(0) = CDbl(Now) Else ok(1) = CDbl(Now)
    Next sis
    ws.Cells.Clear
    ws.Range("A:D").NumberFormat = "@": ws.Range("K:K").NumberFormat = "@"
    ws.Range("A1").Value2 = "ATIL2": ws.Range("B1").Value2 = CDbl(Now): ws.Range("C1").Value2 = n
    ws.Range("D1").NumberFormat = "General": ws.Range("D1").Value2 = ok(0): ws.Range("E1").Value2 = ok(1)
    ws.Range("A2:J2").Value2 = Array("Sistem", "ÜY", "Malzeme", "Giriþler (24 ay)", "Eski giriþ", "Son 6 ay sarf", "Son sarf", _
        "SAP stoðu", "Son stok", "Son stok tarihi")
    ws.Range("K2").Value2 = "Aylýk sarf (6 ay)"
    If n > 0 Then ws.Range("A3").Resize(n, 11).Value2 = a
End Sub

'---------- ATIL STOK: SAP (MB51) ----------
Private Sub ATIL_PANO_YAZ(ByVal metin As String)
    Dim h As LongPtr, p As LongPtr, boyut As LongPtr, i As Long, acik As Boolean
    application.CutCopyMode = False
    ' Pano baska program (pano yoneticisi, OneDrive, Teams, SAP GUI) tarafindan kisa sure tutulabilir: 20 sn'ye kadar beklenir.
    For i = 1 To 210
        If ATIL_OpenClipboard(0) <> 0 Then acik = True: Exit For
        If i Mod 10 = 0 Then DoEvents: ATIL_SURE_YAZ
        ATIL_Sleep IIf(i <= 20, 50, 100)
    Next i
    If Not acik Then Err.Raise vbObjectError + 2330, , "Pano 20 sn boyunca açýlamadý (baþka bir program kullanýyor). Pano kullanan programlarý (Teams ekran paylaþýmý, pano yöneticisi) kapatýp tekrar deneyin."
    ATIL_EmptyClipboard
    boyut = (Len(metin) + 1) * 2
    h = ATIL_GlobalAlloc(&H2, boyut)
    If h = 0 Then ATIL_CloseClipboard: Err.Raise vbObjectError + 2331, , "Pano için bellek ayrýlamadý."
    p = ATIL_GlobalLock(h)
    If p = 0 Then ATIL_GlobalFree h: ATIL_CloseClipboard: Err.Raise vbObjectError + 2332, , "Pano belleði kilitlenemedi."
    ATIL_CopyMemory p, StrPtr(metin), boyut
    ATIL_GlobalUnlock h
    If ATIL_SetClipboardData(13, h) = 0 Then
        ATIL_GlobalFree h: ATIL_CloseClipboard
        Err.Raise vbObjectError + 2333, , "Malzeme listesi panoya yazýlamadý."
    End If
    ATIL_CloseClipboard
End Sub

Private Sub ATIL_SURE_YAZ(Optional ByVal zorla As Boolean = False)
    Dim sn As Long
    If Not mAtilKart Then Exit Sub
    If Not zorla And Abs(Timer - mAtilSon) < 1 Then Exit Sub
    mAtilSon = Timer
    sn = CLng(Fix(ATIL_GECEN(mAtilBas)))
    On Error Resume Next
    If AKSORA_ADMIN_IPTAL Then mAtilMetin = "Durduruluyor: çalýþan SAP sorgusu bitince duracak..."
    frmSplash.lblDurum.Caption = mAtilMetin & vbLf & "Geçen süre: " & (sn \ 60) & " dk " & Format$(sn Mod 60, "00") & " sn"
    frmSplash.Repaint
    On Error GoTo 0
End Sub

Private Sub ATIL_ADIM(ByVal y As Double, ByVal metin As String)
    application.StatusBar = "AKSORA | Atýl stok | " & metin
    If Len(metin) > 72 Then metin = Left$(metin, 69) & "..."
    If AKSORA_ADMIN_IPTAL Then metin = "Durduruluyor: çalýþan SAP sorgusu bitince duracak..."
    mAtilMetin = metin
    AKSORA_BEKLE_ILERLE y, metin
    ATIL_SURE_YAZ True
End Sub

Private Sub ATIL_BEKLE(ByVal ses As Object)
    Dim b As Double
    b = Timer
    Do While ses.Busy
        If ATIL_GECEN(b) > 1800 Then Err.Raise vbObjectError + 2334, , "SAP 30 dakika içinde yanýt vermedi."
        ATIL_SURE_YAZ
        DoEvents: ATIL_Sleep 50
    Loop
End Sub

Private Function ATIL_SBAR(ByVal ses As Object) As String
    On Error Resume Next
    ATIL_SBAR = ses.FindById("wnd[0]/sbar").Text & " [" & ses.FindById("wnd[0]/sbar").MessageType & "]"
    On Error GoTo 0
End Function

Private Sub ATIL_METIN_TOPLA(ByVal kok As Object, ByRef metin As String, ByVal derinlik As Long)
    Dim i As Long, c As Object, t As String, adet As Long
    If derinlik > 5 Then Exit Sub
    On Error Resume Next
    adet = kok.Children.Count
    For i = 0 To adet - 1
        Set c = Nothing
        Set c = kok.Children(i)
        If Not c Is Nothing Then
            t = vbNullString: t = Trim$(c.Text)
            If Len(t) > 0 And c.Type <> "GuiButton" Then metin = metin & " " & t
            If c.ContainerType Then ATIL_METIN_TOPLA c, metin, derinlik + 1
        End If
    Next i
    On Error GoTo 0
End Sub

' "Kayit yok" bildirimi mi? (yetki ve diger hatalar bos sonuc sayilmaz)
Private Function ATIL_BOS_MESAJ(ByVal t As String) As Boolean
    t = " " & Y_NORM(t) & " "
    If InStr(t, "YETKI") > 0 Or InStr(t, "AUTHORI") > 0 Then Exit Function
    If InStr(t, "BELGE") = 0 And InStr(t, "DOCUMENT") = 0 And InStr(t, "KAYIT") = 0 And InStr(t, "VERI") = 0 And InStr(t, "DATA") = 0 Then Exit Function
    ATIL_BOS_MESAJ = (InStr(t, " YOK") > 0 Or InStr(t, "BULUNAMADI") > 0 Or InStr(t, "BULUNMADI") > 0 Or InStr(t, "NOT FOUND") > 0 Or _
        InStr(t, " NO ") > 0 Or InStr(t, "MEVCUT DEGIL") > 0)
End Function

Private Sub ATIL_COKLU_SIL(ByVal ses As Object)
    Dim sil As Object, tip As String
    Set sil = ses.FindById("wnd[1]/tbar[0]/btn[16]", False)
    If sil Is Nothing Then Err.Raise vbObjectError + 2335, , "SAP çoklu seçim penceresi açýlmadý."
    tip = Y_NORM(sil.Tooltip)
    If InStr(tip, "DELETE") = 0 And InStr(tip, "SIL") = 0 Then _
        Err.Raise vbObjectError + 2336, , "SAP çoklu seçim temizleme düðmesi doðrulanamadý: " & sil.Tooltip
    sil.Press
    ATIL_BEKLE ses
End Sub

Private Function ATIL_TARIH(ByVal t As String) As Date
    Dim g As Long, a As Long, y As Long
    On Error GoTo Yok
    t = Trim$(t)
    If t Like "##.##.####" Or t Like "##/##/####" Or t Like "##-##-####" Then
        g = CLng(Left$(t, 2)): a = CLng(Mid$(t, 4, 2)): y = CLng(Mid$(t, 7, 4))
    ElseIf t Like "####-##-##" Or t Like "####.##.##" Or t Like "####/##/##" Then
        y = CLng(Left$(t, 4)): a = CLng(Mid$(t, 6, 2)): g = CLng(Mid$(t, 9, 2))
    ElseIf t Like "########" Then
        y = CLng(Left$(t, 4)): a = CLng(Mid$(t, 5, 2)): g = CLng(Right$(t, 2))
    Else
        Exit Function
    End If
    If y < 1990 Or a < 1 Or a > 12 Or g < 1 Or g > 31 Then Exit Function
    ATIL_TARIH = DateSerial(y, a, g)
    If Month(ATIL_TARIH) <> a Or ATIL_TARIH > Date + 1 Then ATIL_TARIH = 0
Yok:
End Function

Private Function ATIL_SAYI_DENE(ByVal v As Variant, ByRef x As Double) As Boolean
    Dim t As String, eksi As Boolean
    x = 0
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If VarType(v) <> vbString And IsNumeric(v) Then x = CDbl(v): ATIL_SAYI_DENE = True: Exit Function
    t = Replace(Replace(Trim$(CStr(v)), ChrW(160), ""), " ", "")
    If Len(t) = 0 Then Exit Function
    If Right$(t, 1) = "-" Then eksi = True: t = Left$(t, Len(t) - 1)
    If Left$(t, 1) = "-" Then eksi = True: t = Mid$(t, 2)
    Static re As Object
    If re Is Nothing Then Set re = CreateObject("VBScript.RegExp"): re.Pattern = "^[0-9]{1,3}([.][0-9]{3})+$"
    If InStr(t, ",") = 0 Then
        If re.Test(t) Then t = Replace(t, ".", "")
    End If
    If InStr(t, ",") > 0 Then
        If InStrRev(t, ".") > InStrRev(t, ",") Then
            t = Replace(t, ",", "")
        Else
            t = Replace(Replace(t, ".", ""), ",", ".")
        End If
    End If
    If Len(t) = 0 Or t Like "*[!0-9.]*" Or (Len(t) - Len(Replace(t, ".", ""))) > 1 Then Exit Function
    x = Val(t)
    If eksi Then x = -x
    ATIL_SAYI_DENE = True
End Function

' Atil kapsami: D001-D999, K001-K999 ve uretim yerinin yuklenici depolari (yuklenici stok sayfasindan; M### / BO## kaliplari).
Private Function ATIL_KAPSAM_MI(ByVal uy As String, ByVal depo As String) As Boolean
    depo = UCase$(Trim$(depo))
    If Len(depo) = 0 Then Exit Function
    If ATIL_DEPO_MI(depo) Then ATIL_KAPSAM_MI = True: Exit Function
    If depo Like "M###" Or depo Like "BO##" Then ATIL_KAPSAM_MI = True: Exit Function
    If Not mAtilYuk Is Nothing Then
        If mAtilYuk.Exists(uy) Then ATIL_KAPSAM_MI = mAtilYuk(uy).Exists(depo)
    End If
End Function

' MB51 hareket turleri: girisler (iptalleriyle) ve sarflar (iptalleriyle).
Private Function ATIL_GIRIS_TUR() As String()
    ATIL_GIRIS_TUR = Split("101,102,105,106,501,502,561,562,301,302,303,304,305,306", ",")
End Function

Private Function ATIL_SARF_TUR() As String()
    ATIL_SARF_TUR = Split("201,202,221,222,241,242,261,262,281,282,291,292", ",")
End Function

' Satirin uretim yerine giris katkisi: 101/105/501/561 ve iptalleri isaretiyle; uretim yerleri arasi
' transferde yalniz alan taraf (301/303/305 arti, iptalleri 302/304/306 eksi). 313/315 depo yerleri arasi
' transfer oldugundan giris sayilmaz (yasi sifirlamaz).
Private Function ATIL_GIRIS_KATKI(ByVal itu As String, ByVal q As Double) As Double
    Select Case Trim$(itu)
        Case "101", "102", "105", "106", "501", "502", "561", "562": ATIL_GIRIS_KATKI = q
        Case "301", "303", "305": If q > 0 Then ATIL_GIRIS_KATKI = q
        Case "302", "304", "306": If q < 0 Then ATIL_GIRIS_KATKI = q
    End Select
End Function

' Aylik net sarf "yyyymm=miktar;..." (yalniz artilar).
Private Function ATIL_AY_METIN(ByVal aylar As Object, ByVal mal As String) As String
    Dim d As Object, k As Variant, t As String
    If Not aylar.Exists(mal) Then Exit Function
    Set d = aylar(mal)
    For Each k In d.keys
        If CDbl(d(k)) > 0.0005 Then t = t & IIf(Len(t) > 0, ";", "") & CStr(k) & "=" & Replace(CStr(Round(CDbl(d(k)), 3)), ",", ".")
    Next k
    ATIL_AY_METIN = t
End Function

Private Function ATIL_DS(ByVal d As Object, ByVal k As String) As Double
    If d.Exists(k) Then ATIL_DS = ATIL_HUCRE_SAYI(d(k))
End Function

' Gun gun net girisler "yyyymmdd=miktar;..." (iptal gunu kendinden onceki girislerden duser).
Private Function ATIL_KATMAN_METIN(ByVal katman As Object, ByVal mal As String) As String
    Dim gun As Object, k As Variant, n As Long, a() As String, q() As Double, i As Long, j As Long, ts As String, tq As Double, d As Double, t As String
    If Not katman.Exists(mal) Then Exit Function
    Set gun = katman(mal)
    n = gun.Count
    If n = 0 Then Exit Function
    ReDim a(1 To n): ReDim q(1 To n)
    For Each k In gun.keys
        i = i + 1: a(i) = CStr(k): q(i) = CDbl(gun(k))
    Next k
    For i = 2 To n
        ts = a(i): tq = q(i): j = i - 1
        Do While j >= 1
            If a(j) <= ts Then Exit Do
            a(j + 1) = a(j): q(j + 1) = q(j): j = j - 1
        Loop
        a(j + 1) = ts: q(j + 1) = tq
    Next i
    For i = 1 To n
        If q(i) < 0 Then
            d = -q(i): q(i) = 0
            For j = i - 1 To 1 Step -1
                If d <= 0 Then Exit For
                If q(j) > 0 Then
                    If q(j) >= d Then
                        q(j) = q(j) - d: d = 0
                    Else
                        d = d - q(j): q(j) = 0
                    End If
                End If
            Next j
        End If
    Next i
    For i = 1 To n
        If q(i) > 0.0000001 Then t = t & IIf(Len(t) > 0, ";", "") & a(i) & "=" & Trim$(Str$(Round(q(i), 6)))
    Next i
    ATIL_KATMAN_METIN = t
End Function

' Eldeki stok FIFO ile girislere dagitilir (en yeni giris once kalir). 24 aydan eski giristen kalan (ya da
' girisi bulunmayan) kisim ATIL; 21-24 ay 3 AY KALDI; 18-21 ay 6 AY KALDI. q(1) 6 ay, q(2) 3 ay, q(3) atil miktari.
Private Function ATIL_YAS(ByVal stok As Double, ByVal girisler As String, ByVal eski As Double, ByRef q() As Double, _
    ByRef gunMetin As String, ByRef girisMetin As String) As Long
    Dim p As Variant, i As Long, j As Long, n As Long, t() As Double, m() As Double, tt As Double, mm As Double, k As Long
    Dim kalan As Double, al As Double, s24 As Double, s21 As Double, s18 As Double, enYeniAtil As Double, enEskiRisk As Double
    ReDim q(0 To 3)
    gunMetin = vbNullString: girisMetin = vbNullString
    s24 = CDbl(DateAdd("m", -24, Date)): s21 = CDbl(DateAdd("m", -21, Date)): s18 = CDbl(DateAdd("m", -18, Date))
    If Len(girisler) > 0 Then
        p = Split(girisler, ";")
        ReDim t(0 To UBound(p)): ReDim m(0 To UBound(p))
        For i = 0 To UBound(p)
            k = InStr(p(i), "=")
            If k = 9 Then
                t(n) = CDbl(DateSerial(Val(Left$(p(i), 4)), Val(Mid$(p(i), 5, 2)), Val(Mid$(p(i), 7, 2))))
                m(n) = Val(Mid$(p(i), k + 1))
                n = n + 1
            End If
        Next i
    End If
    ' en yeni giris once
    For i = 1 To n - 1
        tt = t(i): mm = m(i): j = i - 1
        Do While j >= 0
            If t(j) >= tt Then Exit Do
            t(j + 1) = t(j): m(j + 1) = m(j): j = j - 1
        Loop
        t(j + 1) = tt: m(j + 1) = mm
    Next i
    kalan = stok
    For i = 0 To n - 1
        If kalan <= 0.0000001 Then Exit For
        If m(i) > 0 Then
            al = m(i)
            If al > kalan Then al = kalan
            kalan = kalan - al
            If t(i) < s24 Then
                q(3) = q(3) + al
                If t(i) > enYeniAtil Then enYeniAtil = t(i)
            ElseIf t(i) < s21 Then
                q(2) = q(2) + al
                If enEskiRisk = 0 Or t(i) < enEskiRisk Then enEskiRisk = t(i)
            ElseIf t(i) < s18 Then
                q(1) = q(1) + al
                If enEskiRisk = 0 Or t(i) < enEskiRisk Then enEskiRisk = t(i)
            End If
        End If
    Next i
    If kalan > 0.0000001 Then
        q(3) = q(3) + kalan
        If eski > enYeniAtil Then enYeniAtil = eski
    End If
    If q(3) > 0.0000001 Then
        ATIL_YAS = 3
        If enYeniAtil > 0 Then
            gunMetin = Format$(application.Max(0, Date - DateAdd("m", 24, CDate(Int(enYeniAtil)))), "#,##0") & " gündür atýl"
            girisMetin = Format$(CDate(Int(enYeniAtil)), "dd.mm.yyyy")
        Else
            gunMetin = "5+ yýldýr giriþ yok": girisMetin = "5 yýldan eski"
        End If
    ElseIf q(2) > 0.0000001 Or q(1) > 0.0000001 Then
        If q(2) > 0.0000001 Then ATIL_YAS = 2 Else ATIL_YAS = 1
        gunMetin = Format$(DateAdd("m", 24, CDate(Int(enEskiRisk))) - Date, "#,##0") & " gün kaldý"
        girisMetin = Format$(CDate(Int(enEskiRisk)), "dd.mm.yyyy")
    End If
End Function

Private Function ATIL_KATMAN_EKLE(ByVal girisler As String, ByVal tarih As Double, ByVal miktar As Double) As String
    ATIL_KATMAN_EKLE = girisler & IIf(Len(girisler) > 0, ";", "") & Format$(CDate(Int(tarih)), "yyyymmdd") & "=" & Trim$(Str$(Round(miktar, 6)))
End Function

' Coklu secim penceresi: temizlenir, degerler panodan yuklenir, ilk deger kontrol edilir.
Private Sub ATIL_COKLU_PANO(ByVal ses As Object, ByVal alan As String, ByRef degerler() As String)
    Dim b As Object, t As String, ilk As String
    ATIL_PANO_YAZ Join(degerler, vbCrLf)
    ses.FindById("wnd[0]/usr/btn%_" & alan & "_%_APP_%-VALU_PUSH").Press: ATIL_BEKLE ses
    ATIL_COKLU_SIL ses
    On Error Resume Next
    ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    On Error GoTo 0
    Set b = ses.FindById("wnd[1]/tbar[0]/btn[24]", False)
    If b Is Nothing Then Err.Raise vbObjectError + 2343, , "SAP 'Panodan yükle' düðmesi bulunamadý."
    t = Y_NORM(b.Tooltip)
    If InStr(t, "PANO") = 0 And InStr(t, "CLIPBOARD") = 0 Then Err.Raise vbObjectError + 2344, , "SAP 'Panodan yükle' düðmesi doðrulanamadý: " & b.Tooltip
    b.Press: ATIL_BEKLE ses
    On Error Resume Next
    ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA").Select
    ilk = ses.FindById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1,0]").Text
    On Error GoTo 0
    If Y_KOD_NORM(ilk) <> Y_KOD_NORM(degerler(LBound(degerler))) Then _
        Err.Raise vbObjectError + 2345, , alan & " listesi SAP'ye aktarýlamadý (ilk deðer '" & ilk & "', beklenen '" & degerler(LBound(degerler)) & "')."
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press: ATIL_BEKLE ses
End Sub

' MB51 listesi satirlari: Array(malzeme, kayit tarihi, miktar (isaretli), hareket turu, depo, degerleme turu).
' Ayni degerli bos hucreler (ALV siralama gosterimi) bir onceki satirdan tamamlanir; ara toplam ve baska uretim yeri satirlari alinmaz.
Private Sub ATIL_MB51_COZ(ByRef v As Variant, ByVal uy As String, ByVal sat As Collection, ByVal etiket As String)
    Dim r As Long, c As Long, h As Long, cM As Long, cT As Long, cQ As Long, cU As Long, cI As Long, cD As Long, cV As Long
    Dim a As String, t As String, x As Double, son As Long, n As Long, dolu As Long, enIyi As Long
    Dim ham As String, mal As String, tar As Date, uyx As String, itu As String, depo As String, dgr As String, q As Double
    Dim sMal As String, sTar As Date, sUy As String, sItu As String, sDepo As String, sDgr As String, atla As Long, okunan As Long, hb As String
    For r = 1 To application.Min(UBound(v, 1), 40)
        For c = 1 To UBound(v, 2)
            If Y_NORM(CStr(v(r, c))) = "MALZEME" Or Y_NORM(CStr(v(r, c))) = "MATERIAL" Then h = r: cM = c: Exit For
        Next c
        If h > 0 Then Exit For
    Next r
    If h = 0 Then Err.Raise vbObjectError + 2340, , "MB51 listesinde 'Malzeme' sütunu bulunamadý (/ATIL düzenini kontrol edin)."
    hb = vbNullString
    For c = 1 To UBound(v, 2): hb = hb & " | " & CStr(v(h, c)): Next c
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | baþlýk (satýr " & h & ")" & hb
    For c = 1 To UBound(v, 2)
        a = Y_NORM(CStr(v(h, c)))
        If cT = 0 And (InStr(a, "KAYIT") > 0 Or InStr(a, "KYT") > 0 Or InStr(a, "POSTING") > 0) Then cT = c
        If cQ = 0 And (a = "MIKTAR" Or a = "QUANTITY") Then cQ = c
        If cU = 0 And (InStr(a, "URETIM") > 0 Or a = "PLANT" Or a = "UY") Then cU = c
        If cI = 0 And (a = "ITU" Or a = "HTU" Or a = "BWART" Or InStr(a, "HAREKET") > 0 Or InStr(a, "HRK") > 0 Or InStr(a, "MOVEMENT") > 0 Or a = "MVT") Then cI = c
        If cD = 0 And (a = "DPYR" Or InStr(a, "DEPO") > 0 Or a = "SLOC" Or InStr(a, "STOR") > 0) Then cD = c
        If cV = 0 And (InStr(a, "DGRL") > 0 Or InStr(a, "DEGERLEME") > 0 Or InStr(a, "VALUATION") > 0) Then cV = c
    Next c
    If cQ = 0 Then
        For c = 1 To UBound(v, 2)
            a = Y_NORM(CStr(v(h, c)))
            If c <> cM And c <> cT And (InStr(a, "MIK") > 0 Or InStr(a, "MKT") > 0 Or InStr(a, "QUANT") > 0) Then cQ = c: Exit For
        Next c
    End If
    son = application.Min(UBound(v, 1), h + 300)
    If cT = 0 Then
        enIyi = 0
        For c = 1 To UBound(v, 2)
            If c <> cM Then
                n = 0: dolu = 0
                For r = h + 1 To son
                    t = Trim$(CStr(v(r, c)))
                    If Len(t) > 0 Then
                        dolu = dolu + 1
                        If ATIL_TARIH(t) > 0 Then n = n + 1
                    End If
                Next r
                If n > enIyi And n * 10 >= dolu * 9 Then enIyi = n: cT = c
            End If
        Next c
    End If
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | sütunlar: malzeme " & cM & ", tarih " & cT & ", miktar " & cQ & ", ÜY " & cU & ", ÝTü " & cI & ", depo " & cD & ", dðrl " & cV
    If cT = 0 Or cQ = 0 Or cI = 0 Or cD = 0 Then Err.Raise vbObjectError + 2341, , _
        "MB51 /ATIL düzeninde kayýt tarihi, miktar, hareket türü (ÝTü) ve depo yeri sütunlarý gerekli. Baþlýk:" & Left$(hb, 200)
    For r = h + 1 To UBound(v, 1)
        If (r And 4095) = 0 Then AKSORA_BEKLE_NABIZ: ATIL_SURE_YAZ
        ham = Trim$(CStr(v(r, cM)))
        If Y_NORM(ham) = "MALZEME" Or Y_NORM(ham) = "MATERIAL" Then GoTo Sonraki
        t = Trim$(CStr(v(r, cT)))
        If Left$(ham, 1) = "*" Or Left$(t, 1) = "*" Then atla = atla + 1: GoTo Sonraki
        If Len(ham) = 0 And Len(t) = 0 And Len(Trim$(CStr(v(r, cQ)))) = 0 Then GoTo Sonraki
        mal = Y_KOD_NORM(ham)
        If Len(mal) = 0 Then mal = sMal
        ' Bos hucre yalniz ALV tekrar gizlemesinde (malzeme hucresi de bos) onceki satirdan tamamlanir;
        ' depo yeri bos gercek hareket (ornek: dogrudan tuketime giris) baska depoya yazilmaz.
        If Len(t) = 0 And Len(ham) = 0 Then tar = sTar Else tar = ATIL_TARIH(t)
        uyx = vbNullString
        If cU > 0 Then uyx = Trim$(CStr(v(r, cU)))
        If Len(uyx) = 0 And Len(ham) = 0 Then uyx = sUy
        itu = Trim$(CStr(v(r, cI)))
        If Len(itu) = 0 And Len(ham) = 0 Then itu = sItu
        depo = UCase$(Trim$(CStr(v(r, cD))))
        If Len(depo) = 0 And Len(ham) = 0 Then depo = sDepo
        dgr = vbNullString
        If cV > 0 Then dgr = Trim$(CStr(v(r, cV)))
        If Len(dgr) = 0 And Len(ham) = 0 Then dgr = sDgr
        If Len(mal) = 0 Or tar = 0 Or Len(itu) = 0 Then atla = atla + 1: GoTo Sonraki
        sMal = mal: sTar = tar: sUy = uyx: sItu = itu: sDepo = depo: sDgr = dgr
        If Len(uyx) > 0 And uyx <> uy Then atla = atla + 1: GoTo Sonraki
        q = 0
        If Not ATIL_SAYI_DENE(v(r, cQ), q) Then atla = atla + 1: GoTo Sonraki
        sat.Add Array(mal, CDbl(tar), q, itu, depo, dgr)
        okunan = okunan + 1
Sonraki:
    Next r
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | okunan " & okunan & " satýr, atlanan " & atla
End Sub

' MB51 (/ATIL duzeni): uretim yeri, hareket turleri ve malzemeler panodan, tarih araligi; depo filtresi yok
' (satirlar okunurken atil kapsamina gore suzulur). Donus: satirlar (ATIL_MB51_COZ).
' MB51 liste gosterimi "Duz liste" secilir: /ATIL duzeni duz listede kayitli; hiyerarsik gosterimde
' SAP "Herhangi bir duzenek bulunamadi" der. Once standart alan (RFLAT_L), yoksa secim ekranindaki
' radyo dugmeleri metninden aranir. Bulunan secenekler teshise yazilir.
Private Sub ATIL_DUZ_LISTE(ByVal ses As Object, ByVal etiket As String)
    Dim usr As Object, o As Object, t As String, liste As String, n As Long, deneme As Long, hata As String
    ' Oturumdaki ilk erisimde ekran hazir olmayabilir: kisa bekleyip bir kez daha denenir.
    For deneme = 1 To 2
        hata = vbNullString: liste = vbNullString
        On Error GoTo Hata
        Set o = Nothing
        Set o = ses.FindById("wnd[0]/usr/radRFLAT_L", False)
        If Not o Is Nothing Then
            o.Select
            AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | liste gösterimi: " & o.Text & " (RFLAT_L)"
            Exit Sub
        End If
        Set usr = ses.FindById("wnd[0]/usr")
        For n = 0 To CLng(usr.Children.Count) - 1
            Set o = usr.Children.ElementAt(n)
            If CStr(o.Type) = "GuiRadioButton" Then
                t = Y_NORM(CStr(o.Text))
                liste = liste & " | " & o.Name & "=" & o.Text
                If InStr(t, "DUZ") > 0 Or InStr(t, "FLAT") > 0 Then
                    o.Select
                    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | liste gösterimi: " & o.Text & " (" & o.Name & ")"
                    Exit Sub
                End If
            End If
        Next n
        AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | 'Düz liste' seçeneði bulunamadý. Radyo düðmeleri:" & liste
        Exit Sub
Sonraki:
        On Error GoTo 0
        ATIL_BEKLE ses
        application.Wait Now + TimeSerial(0, 0, 1)
    Next deneme
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | liste gösterimi seçilemedi: " & hata & " | bulunanlar:" & liste
    Exit Sub
Hata:
    hata = Err.Description
    Resume Sonraki
End Sub

Private Function ATIL_MB51(ByVal ses As Object, ByVal uy As String, ByRef mallar() As String, ByVal bas As Date, ByVal son As Date, _
    ByVal etiket As String, ByRef turler() As String) As Collection
    Dim sat As Collection, alv As Object, v As Variant, n As Long, pop As Object, t As String, b0 As Double, varsayilan As Boolean
    Set sat = New Collection
    Set ATIL_MB51 = sat
    n = UBound(mallar) - LBound(mallar) + 1
    If n <= 0 Then Exit Function
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | " & n & " malzeme, " & Format$(bas, "dd.mm.yyyy") & " - " & Format$(son, "dd.mm.yyyy") & " | ÝTü " & Join(turler, ",")
    ses.StartTransaction "MB51": ATIL_BEKLE ses
    If ses.FindById("wnd[0]/usr/ctxtWERKS-LOW", False) Is Nothing Then _
        Err.Raise vbObjectError + 2342, , "MB51 seçim ekraný açýlmadý: " & ATIL_SBAR(ses)
    ATIL_COKLU_PANO ses, "MATNR", mallar
    ses.FindById("wnd[0]/usr/ctxtWERKS-LOW").Text = uy
    On Error Resume Next
    ses.FindById("wnd[0]/usr/ctxtWERKS-HIGH").Text = vbNullString
    On Error GoTo 0
    ' Depo: tum depolar (onceki secim temizlenir)
    ses.FindById("wnd[0]/usr/btn%_LGORT_%_APP_%-VALU_PUSH").Press: ATIL_BEKLE ses
    ATIL_COKLU_SIL ses
    ses.FindById("wnd[1]/tbar[0]/btn[8]").Press: ATIL_BEKLE ses
    On Error Resume Next
    ses.FindById("wnd[0]/usr/ctxtLGORT-LOW").Text = vbNullString
    ses.FindById("wnd[0]/usr/ctxtLGORT-HIGH").Text = vbNullString
    On Error GoTo 0
    ATIL_COKLU_PANO ses, "BWART", turler
    ses.FindById("wnd[0]/usr/ctxtBUDAT-LOW").Text = Format$(bas, "dd.mm.yyyy")
    ses.FindById("wnd[0]/usr/ctxtBUDAT-HIGH").Text = Format$(son, "dd.mm.yyyy")
    ATIL_DUZ_LISTE ses, etiket
    ses.FindById("wnd[0]/usr/ctxtALV_DEF").Text = "/ATIL"
Calistir:
    b0 = Timer
    ses.FindById("wnd[0]/tbar[1]/btn[8]").Press
    ATIL_BEKLE ses
    AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | sorgu " & AKSORA_SAYI_METNI(ATIL_GECEN(b0), "0.0") & " sn | " & ATIL_SBAR(ses)
    Set pop = ses.FindById("wnd[1]", False)
    If Not pop Is Nothing Then
        t = vbNullString
        On Error Resume Next
        t = pop.Text
        On Error GoTo 0
        ATIL_METIN_TOPLA pop, t, 0
        AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | SAP penceresi: " & t
        If ATIL_BOS_MESAJ(t) Then
            pop.SendVKey 0: ATIL_BEKLE ses
            Exit Function
        End If
        Err.Raise vbObjectError + 2346, , "SAP penceresi: " & Trim$(t)
    End If
    Set alv = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell", False)
    If alv Is Nothing And Not ses.FindById("wnd[0]/usr/ctxtWERKS-LOW", False) Is Nothing Then
        t = ATIL_SBAR(ses)
        If ATIL_BOS_MESAJ(t) Then Exit Function
        ' /ATIL duzeni bu liste gosteriminde yoksa varsayilan duzen denenir (gerekli sutunlar basliktan bulunur).
        If InStr(Y_NORM(t), "DUZEN") > 0 And Not varsayilan Then
            AKSORA_SAP_TESHIS_YAZ "Atýl " & etiket & " | /ATIL düzeni kullanýlamadý (" & t & "); varsayýlan düzenle deneniyor"
            varsayilan = True
            ses.FindById("wnd[0]/usr/ctxtALV_DEF").Text = vbNullString
            GoTo Calistir
        End If
        If InStr(Y_NORM(t), "DUZEN") > 0 Then Err.Raise vbObjectError + 2347, , _
            "Bu sistemde MB51 için /ATIL düzeni yok (SAP: " & t & "). MB51'de listeyi açýp sütunlarý ayarlayýn ve /ATIL adýyla kaydedin."
        Err.Raise vbObjectError + 2347, , "MB51 sonuç vermedi: " & t
    End If
    v = AKSORA_ALV_METIN_AL(ses, alv, "ATIL_" & uy)
    If Not IsArray(v) Then Err.Raise vbObjectError + 2348, , "MB51 listesi okunamadý (yerel dosya). Ayrýntý: AKSORA_SAP_TESHIS.txt"
    ATIL_MB51_COZ v, uy, sat, etiket
End Function

Private Function ATIL_DIZI(ByVal d As Object) As String()
    Dim a() As String, k As Variant, i As Long
    If d.Count = 0 Then Err.Raise vbObjectError + 2349, , "Sorgulanacak malzeme listesi boþ."
    ReDim a(0 To d.Count - 1)
    For Each k In d.keys
        a(i) = CStr(k): i = i + 1
    Next k
    ATIL_DIZI = a
End Function

Public Sub AKSORA_ATIL_SAP_OKU()
    Dim tum As Object, il As Variant, sis As Variant, ses As Object, msg As String, oturumlar As Object
    Dim sonuc As Object, adim As String, eksik As String, atlanan As String, bugun As Date
    Dim adet As Long, toplam As Long, no As Long, metin As String, ev As Boolean, sure As Double, ozet As String
    Dim tamam As Object, hatalar As String, okuma As Variant, kayitTmp As Object, bugunOk As String, digeri As String, cvp As VbMsgBoxResult
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ADMIN_TOPLU Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error GoTo Hata
    adim = "Stok verisi okunuyor"
    mAtilKart = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Stok verisi ve SAP oturumlarý hazýrlanýyor...", 0.1)
    Set tum = ATIL_TUM_STOK()
    For Each il In ADM_ILLER()
        If tum(CStr(il))(1) = 0 Then eksik = eksik & ", " & Y_IL_ADI(CStr(il))
    Next il
    If Len(eksik) > 0 Then
        If mAtilKart Then AKSORA_BEKLE_KAPAT
        mAtilKart = False
        MsgBox "Þu illerin Malzeme stok verisi yok: " & Mid$(eksik, 3) & vbCrLf & _
            "Önce Malzeme biriminde bütün illerin stoklarýný güncelleyin.", vbExclamation, "AKSORA | ATIL STOK"
        Exit Sub
    End If
    adim = "SAP oturumlarý"
    Set oturumlar = CreateObject("Scripting.Dictionary")
    For Each sis In Array("CDP", "EPP")
        msg = vbNullString
        Set ses = Nothing
        Set ses = SAP_SESSION_AL(CStr(sis), msg)
        If ses Is Nothing Then atlanan = atlanan & ", " & sis Else oturumlar.Add CStr(sis), ses
    Next sis
    If mAtilKart Then AKSORA_BEKLE_KAPAT
    mAtilKart = False
    If oturumlar.Count = 0 Then
        MsgBox "SAP CDP ve EPP oturumu bulunamadý." & vbCrLf & "SAP'ye giriþ yapýn, sonra tekrar deneyin.", vbExclamation, "AKSORA | ATIL STOK"
        Exit Sub
    End If
    ' Bir sistem bugun okunduysa (onceki okuma kismen kaldiysa) yalniz digeri okunabilir.
    If oturumlar.Count = 2 Then
        Set kayitTmp = ATIL_KAYIT_OKU(okuma)
        If Int(CDbl(okuma(0))) = CDbl(Date) And Int(CDbl(okuma(1))) <> CDbl(Date) Then bugunOk = "CDP": digeri = "EPP"
        If Int(CDbl(okuma(1))) = CDbl(Date) And Int(CDbl(okuma(0))) <> CDbl(Date) Then bugunOk = "EPP": digeri = "CDP"
        Set kayitTmp = Nothing
        If Len(bugunOk) > 0 Then
            cvp = MsgBox(bugunOk & " bugün okundu." & vbCrLf & vbCrLf & "Yalnýz " & digeri & " okunsun mu?" & vbCrLf & vbCrLf & _
                "Evet: yalnýz " & digeri & "   •   Hayýr: ikisi de   •   Ýptal: vazgeç", vbYesNoCancel + vbQuestion, "AKSORA | ATIL STOK")
            If cvp = vbCancel Then Exit Sub
            If cvp = vbYes Then oturumlar.Remove bugunOk
        End If
    End If
    If MsgBox("SAP'den atýl stok verisi okunacak (MB51)." & vbCrLf & vbCrLf & _
        "Sistem: " & Join(oturumlar.keys, " + ") & _
        IIf(Len(atlanan) > 0, vbCrLf & "Açýk oturum yok: " & Mid$(atlanan, 3) & " (eski kaydý korunur)", "") & vbCrLf & _
        "Süre: üretim yeri baþýna yaklaþýk 1-4 dakika." & vbCrLf & vbCrLf & "Baþlasýn mý?", vbYesNo + vbQuestion, "AKSORA | ATIL STOK") <> vbYes Then Exit Sub
    AKSORA_ADMIN_IPTAL = False
    AKSORA_ATIL_CALISIYOR = True
    mAtilBas = Timer: mAtilSon = 0
    mAtilKart = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "SAP'den okunuyor...", 0.02, True)
    AKSORA_SAP_OLE_SESSIZ True
    ev = application.EnableEvents
    application.EnableEvents = False
    bugun = Date
    Set sonuc = CreateObject("Scripting.Dictionary"): sonuc.CompareMode = vbTextCompare
    toplam = oturumlar.Count * 15
    Set tamam = CreateObject("Scripting.Dictionary")
    For Each sis In oturumlar.keys
        If AKSORA_ADMIN_IPTAL Then
            hatalar = hatalar & vbCrLf & "• " & sis & ": baþlatýlmadý (durduruldu)"
        Else
            metin = ATIL_SISTEM_OKU(oturumlar(sis), CStr(sis), tum, sonuc, bugun, adet, toplam)
            ' Gecici hata (pano mesgul, SAP penceresi) bir kez daha denenir.
            If Len(metin) > 0 And Not AKSORA_ADMIN_IPTAL Then
                AKSORA_SAP_TESHIS_YAZ "Atýl " & sis & " yeniden deneniyor: " & metin
                On Error Resume Next
                If Not oturumlar(sis).FindById("wnd[1]", False) Is Nothing Then oturumlar(sis).FindById("wnd[1]/tbar[0]/btn[12]").Press
                On Error GoTo Hata
                ATIL_ADIM 0.03 + 0.92 * adet / toplam, sis & " yeniden deneniyor (geçici hata)..."
                DoEvents: ATIL_Sleep 2000
                toplam = toplam + 15
                metin = ATIL_SISTEM_OKU(oturumlar(sis), CStr(sis), tum, sonuc, bugun, adet, toplam)
            End If
            If Len(metin) = 0 Then tamam.Add CStr(sis), oturumlar(sis) Else hatalar = hatalar & vbCrLf & "• " & sis & ": " & metin
            metin = vbNullString
        End If
    Next sis
    If tamam.Count = 0 Then adim = "SAP okumasý": Err.Raise vbObjectError + 2351, , Mid$(hatalar, 3)
    adim = "Kayýt yazýlýyor"
    ATIL_ADIM 0.96, "Kayýt yazýlýyor..."
    ATIL_KAYDET sonuc, tum, tamam
    sure = ATIL_GECEN(mAtilBas)
    AKSORA_SAP_TESHIS_YAZ "Atýl stok SAP okumasý tamamlandý: " & AKSORA_SAYI_METNI(sure, "0") & " sn"
    adim = "Sayfa hazýrlanýyor"
    ATIL_ADIM 0.98, "Sayfa hazýrlanýyor..."
    GoTo Bitir
Hata:
    no = Err.Number: metin = Err.Description
    Resume Bitir
Bitir:
    On Error Resume Next
    If Not oturumlar Is Nothing And AKSORA_ATIL_CALISIYOR Then
        For Each sis In oturumlar.keys
            Set ses = oturumlar(sis)
            If Not ses.FindById("wnd[1]", False) Is Nothing Then ses.FindById("wnd[1]/tbar[0]/btn[12]").Press
            ' SAP hafizasindaki malzeme (MAT) bosaltilir; yoksa sonraki MB52 stok okumalari tek malzemeyle calisir.
            ses.StartTransaction "MB51"
            AKSORA_SAP_MALZEME_TEMIZLE ses
            ses.FindById("wnd[0]").SendVKey 0
            ses.StartTransaction "SESSION_MANAGER"
        Next sis
    End If
    AKSORA_SAP_OLE_SESSIZ False
    AKSORA_ARKA_PLAN_BITIR
    If AKSORA_ATIL_CALISIYOR Then application.EnableEvents = ev
    application.StatusBar = False
    If no = 0 Then
        ATIL_GOSTER "TUM"
        ozet = ATIL_OZET_METNI()
    End If
    If mAtilKart Then AKSORA_BEKLE_KAPAT
    mAtilKart = False
    AKSORA_ATIL_CALISIYOR = False
    AKSORA_ADMIN_IPTAL = False
    On Error GoTo 0
    If no <> 0 Then
        MsgBox "Atýl stok SAP okumasý tamamlanamadý." & vbCrLf & vbCrLf & "Adým: " & adim & vbCrLf & "Hata: " & metin & vbCrLf & vbCrLf & _
            "Önceki atýl stok kaydý deðiþtirilmedi.", vbExclamation, "AKSORA | ATIL STOK"
    ElseIf Len(hatalar) > 0 Then
        MsgBox "Atýl stok SAP okumasý kýsmen tamamlandý." & vbCrLf & "Okunan ve kaydedilen: " & Join(tamam.keys, " + ") & vbCrLf & _
            "Süre: " & AKSORA_SURE_METNI(sure) & vbCrLf & vbCrLf & "Okunamayan (eski kaydý korunur):" & hatalar & vbCrLf & vbCrLf & ozet, vbExclamation, "AKSORA | ATIL STOK"
    ElseIf adet > 0 Then
        MsgBox "Atýl stok SAP okumasý tamamlandý." & vbCrLf & "Süre: " & AKSORA_SURE_METNI(sure) & vbCrLf & vbCrLf & ozet, vbInformation, "AKSORA | ATIL STOK"
    End If
End Sub

'---------- ATIL STOK: hesap ----------

Private Function ATIL_DURUM_ADI(ByVal durum As Long) As String
    Select Case durum
        Case 3: ATIL_DURUM_ADI = "ATIL"
        Case 2: ATIL_DURUM_ADI = "3 AY KALDI"
        Case 1: ATIL_DURUM_ADI = "6 AY KALDI"
        Case 0: ATIL_DURUM_ADI = "NORMAL"
        Case Else: ATIL_DURUM_ADI = "SORGULANMADI"
    End Select
End Function

' Kalemler: Array(0 il, 1 sistem, 2 UY, 3 malzeme, 4 tanim, 5 kategori, 6 stok, 7 birim, 8 stok degeri, 9 riskli girisin tarihi,
'   10 durum (-1 sorgulanmadi, 0 normal, 1 6 ay, 2 3 ay, 3 atil), 11 gun metni, 12 stok degisimi, 13 oneri ili, 14 oneri nedeni,
'   15 depolar, 16 stogu dustu mu, 17 gonderilecek, 18 oncelik, 19 deger (1 6 ay, 2 3 ay, 3 atil), 20 miktar (ayni sira),
'   21 riskli miktar, 22 riskli deger)
' bilgi: Array(okuma CDP, okuma EPP, sorgulanmayan, stogu dusen, stok tarihleri, eski kural kaydi var mi)
Public Function ATIL_HESAPLA(ByRef bilgi As Variant) As Collection
    Dim tum As Object, kayit As Object, okuma As Variant, c As Collection, il As Variant, s As Object, k As Variant
    Dim uy As String, mal As String, sis As String, o As Variant, kr As Variant, durum As Long, q() As Double
    Dim gunMetin As String, girisMetin As String, degisim As String, qq As Variant, quy As String, enIyi As Double
    Dim oIl As String, oMetin As String, qs As Double, sorgusuz As Long, dusen As Long, tarihler As String, ok As Double, s2 As Object
    Dim sarf6 As Double, oSarf As Double, oStok As Double, oMik As Double, oOnc As String, birim As String, bd As Double
    Dim risk As Double, dustu As Boolean, tv As Variant, tq As Variant
    If AN_ONB_GECERLI() Then
        If Not mOnbHes Is Nothing Then Set ATIL_HESAPLA = mOnbHes: bilgi = mOnbBilgi: Exit Function
    End If
    Set c = New Collection
    Set tum = ATIL_TUM_STOK()
    ATIL_KARSILASTIR tum
    Set kayit = ATIL_KAYIT_OKU(okuma)
    For Each il In ADM_ILLER()
        If tum(CStr(il))(1) > 0 Then
            tarihler = tarihler & "  " & Y_IL_ADI(CStr(il)) & " " & Format$(tum(CStr(il))(1), "dd.mm.yyyy")
        Else
            tarihler = tarihler & "  " & Y_IL_ADI(CStr(il)) & " -"
        End If
        Set s = tum(CStr(il))(0)
        For Each k In s.keys
            AKSORA_BEKLE_NABIZ
            uy = Left$(CStr(k), InStr(CStr(k), "|") - 1): mal = Mid$(CStr(k), InStr(CStr(k), "|") + 1)
            sis = ATIL_SISTEM(uy)
            o = s(k)
            If sis = "CDP" Then ok = okuma(0) Else ok = okuma(1)
            oIl = vbNullString: oMetin = vbNullString: degisim = vbNullString: gunMetin = vbNullString: girisMetin = vbNullString
            oSarf = 0: oStok = 0: oMik = 0: oOnc = vbNullString: birim = CStr(o(2)): risk = 0: dustu = False
            tv = Array(0#, 0#, 0#, 0#): tq = Array(0#, 0#, 0#, 0#)
            If o(0) > 0 Then bd = o(1) / o(0) Else bd = 0
            If kayit.Exists(CStr(k)) And ok > 0 Then
                kr = kayit(CStr(k))
                durum = ATIL_YAS(o(0), CStr(kr(1)), kr(2), q, gunMetin, girisMetin)
                tq = Array(0#, q(1), q(2), q(3))
                tv = Array(0#, q(1) * bd, q(2) * bd, q(3) * bd)
                risk = q(1) + q(2) + q(3)
                If o(0) > kr(5) + 0.0005 Then
                    degisim = "ARTTI"
                ElseIf o(0) < kr(5) - 0.0005 Then
                    degisim = "DÜÞTÜ": dustu = True: dusen = dusen + 1
                Else
                    degisim = "DEÐÝÞMEDÝ"
                End If
                ' Transfer onerisi (ayni sistem): son 6 ayda bu malzemeyi kullanan iller icinden eksigi en buyuk olan
                ' (6 aylik kullanim - stogu). YUKSEK: stogu 6 aylik kullanimindan az; DUSUK: stogu yetiyor.
                If risk > 0.0000001 Then
                    For Each qq In ADM_ILLER()
                        If CStr(qq) <> CStr(il) Then
                            quy = AKSORA_IL_TESIS(CStr(qq), sis)
                            If kayit.Exists(quy & "|" & mal) Then
                                sarf6 = kayit(quy & "|" & mal)(3)
                                If sarf6 > 0 Then
                                    qs = 0
                                    Set s2 = tum(CStr(qq))(0)
                                    If s2.Exists(quy & "|" & mal) Then qs = s2(quy & "|" & mal)(0)
                                    If Len(oIl) = 0 Or sarf6 - qs > enIyi Or (sarf6 - qs = enIyi And sarf6 > oSarf) Then
                                        enIyi = sarf6 - qs: oIl = CStr(qq): oSarf = sarf6: oStok = qs
                                    End If
                                End If
                            End If
                        End If
                    Next qq
                    If Len(oIl) > 0 Then
                        If oSarf > oStok + 0.0005 Then
                            oOnc = "YÜKSEK": oMik = application.Min(risk, oSarf - oStok)
                        Else
                            oOnc = "DÜÞÜK": oMik = application.Min(risk, oSarf)
                        End If
                        oMetin = "Son 6 ayda " & ADM_ADET(oSarf) & " " & birim & " kullandý; deposunda " & _
                            IIf(oStok > 0, ADM_ADET(oStok) & " " & birim & " var.", "hiç yok.")
                    End If
                End If
            Else
                durum = -1: sorgusuz = sorgusuz + 1
            End If
            c.Add Array(CStr(il), sis, uy, mal, CStr(o(3)), CStr(o(4)), o(0), birim, o(1), girisMetin, durum, gunMetin, degisim, oIl, oMetin, o(5), _
                dustu, oMik, oOnc, tv, tq, risk, risk * bd)
        Next k
    Next il
    bilgi = Array(okuma(0), okuma(1), sorgusuz, dusen, Trim$(tarihler), mAtilEski)
    Set ATIL_HESAPLA = c
    If AN_ONB_GECERLI() Then Set mOnbHes = c: mOnbBilgi = bilgi
End Function

Private Function ATIL_OZET_METNI() As String
    Dim c As Collection, bilgi As Variant, x As Variant, t(0 To 3) As Double, a(0 To 3) As Long, n As Long, j As Long
    Set c = ATIL_HESAPLA(bilgi)
    For Each x In c
        If x(10) >= 1 Then
            For j = 1 To 3
                If x(20)(j) > 0.0000001 Then t(j) = t(j) + x(19)(j): a(j) = a(j) + 1
            Next j
            If Len(x(13)) > 0 Then n = n + 1
        End If
    Next x
    ATIL_OZET_METNI = "ATIL: " & a(3) & " kalem, " & ADM_TL(t(3)) & vbCrLf & _
        "3 AY KALDI: " & a(2) & " kalem, " & ADM_TL(t(2)) & vbCrLf & _
        "6 AY KALDI: " & a(1) & " kalem, " & ADM_TL(t(1)) & vbCrLf & _
        "Transfer önerisi: " & n & " kalem"
End Function

'---------- ATIL STOK: sayfa ----------
Public Sub AKSORA_ATIL_AC()
    Dim ac As Boolean, no As Long, metin As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Atýl stok sayfasý hazýrlanýyor...")
    On Error GoTo Son
    ATIL_GOSTER vbNullString
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Atýl stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_ATIL_IL_SEC()
    Dim ac As Boolean, no As Long, metin As String, il As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error Resume Next
    il = ThisWorkbook.Worksheets(ATIL_SAYFA).Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(il) = 0 Then il = "TUM"
    ac = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Atýl stok sayfasý hazýrlanýyor...")
    On Error GoTo Son
    ATIL_GOSTER il
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Atýl stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Private Sub ATIL_GOSTER(ByVal il As String, Optional ByVal gorunum As String = "")
    Dim ws As Worksheet, ev As Boolean, hesap As XlCalculation, no As Long, metin As String
    ev = application.EnableEvents: hesap = application.Calculation
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    AN_ONB_AC
    Set ws = ATIL_SAYFA_KUR()
    If Len(gorunum) > 0 Then ws.Unprotect "": ws.Range("BX21").Value2 = gorunum
    If Len(il) = 0 Then il = CStr(ws.Range("BX20").Value2)
    If Len(il) = 0 Then il = "TUM"
    ATIL_CIZ ws, il
    ws.Visible = xlSheetVisible
    mYonlendirme = True
    ws.Select Replace:=True
    mYonlendirme = False
    AKSORA_YATIRIM_GORUNUM ws
Cikis:
    ' Temizlik hatasi Hata'ya donup sonsuz donguye girmesin.
    On Error Resume Next
    AN_ONB_KAPAT
    mYonlendirme = False
    application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If no <> 0 Then Err.Raise no, "ATIL_GOSTER", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Resume Cikis
End Sub

' Secili satirdaki malzemenin atil hesabinin kaniti: uretim yerindeki son 24 ay girisleri (yeni -> eski),
' her girisin stokta kalan payi ve yasi, girisle karsilanmayan kisim, sarf. MB51 ile elle karsilastirma icin.
Public Sub AKSORA_ATIL_KANIT()
    Dim ws As Worksheet, r As Long, ilAd As String, sis As String, mal As String, il As String, x As Variant, uy As String
    Dim kayit As Object, okuma As Variant, kr As Variant, s As Object, td As Date, o As Variant, stok As Double, birim As String
    Dim q() As Double, gm As String, gi As String, durum As Long, p As Variant, i As Long, j As Long, n As Long
    Dim tt() As Double, mm() As Double, a As Double, b As Double, kalan As Double, al As Double, sinif As String
    Dim s24 As Double, s21 As Double, s18 As Double, m As String, satir As String, gizli As Long, kart As Boolean
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error GoTo Hata
    Set ws = ActiveSheet
    If ws.Name <> ATIL_SAYFA Then Exit Sub
    r = ActiveCell.Row
    ilAd = Trim$(CStr(ws.Cells(r, 1).Value2)): sis = UCase$(Trim$(CStr(ws.Cells(r, 2).Value2))): mal = Y_KOD_NORM(ws.Cells(r, 3).Value2)
    For Each x In ADM_ILLER()
        If Y_IL_ADI(CStr(x)) = ilAd Then il = CStr(x)
    Next x
    If Len(il) = 0 Or (sis <> "CDP" And sis <> "EPP") Or Len(mal) = 0 Then
        MsgBox "Önce TRANSFER ÖNERÝLERÝ ya da ATIL LÝSTESÝ tablosunda bir malzeme satýrýna týklayýn, sonra KANIT düðmesine basýn.", vbInformation, "AKSORA | ATIL KANIT"
        Exit Sub
    End If
    kart = AKSORA_BEKLE_AC("AKSORA  |  ATIL KANIT", mal & " hesabý hazýrlanýyor...", 0.4)
    uy = AKSORA_IL_TESIS(il, sis)
    Set kayit = ATIL_KAYIT_OKU(okuma)
    If Not kayit.Exists(uy & "|" & mal) Then
        If kart Then AKSORA_BEKLE_KAPAT
        MsgBox mal & " için " & sis & " SAP okumasý kaydý yok.", vbInformation, "AKSORA | ATIL KANIT"
        Exit Sub
    End If
    kr = kayit(uy & "|" & mal)
    Set s = ATIL_STOK_OKU(il, td)
    If s.Exists(uy & "|" & mal) Then
        o = s(uy & "|" & mal): stok = o(0): birim = CStr(o(2))
        m = mal & "  " & CStr(o(3)) & vbCrLf & ilAd & " " & sis & " (" & uy & ")  •  Stok: " & ADM_ADET(stok) & " " & birim & vbCrLf & _
            "Depolar: " & ATIL_DEPO_METIN(o(5)) & vbCrLf & vbCrLf
    Else
        m = mal & vbCrLf & ilAd & " " & sis & " (" & uy & ")  •  Stok: 0" & vbCrLf & vbCrLf
    End If
    durum = ATIL_YAS(stok, CStr(kr(1)), kr(2), q, gm, gi)
    s24 = CDbl(DateAdd("m", -24, Date)): s21 = CDbl(DateAdd("m", -21, Date)): s18 = CDbl(DateAdd("m", -18, Date))
    ' Girisler yeni -> eski; stok en yeni giristen baslayarak dagitilir (ATIL_YAS ile ayni).
    If Len(CStr(kr(1))) > 0 Then
        p = Split(CStr(kr(1)), ";")
        ReDim tt(0 To UBound(p)): ReDim mm(0 To UBound(p))
        For i = 0 To UBound(p)
            If InStr(p(i), "=") = 9 Then
                tt(n) = CDbl(DateSerial(Val(Left$(p(i), 4)), Val(Mid$(p(i), 5, 2)), Val(Mid$(p(i), 7, 2))))
                mm(n) = Val(Mid$(p(i), 10)): n = n + 1
            End If
        Next i
        For i = 1 To n - 1
            a = tt(i): b = mm(i): j = i - 1
            Do While j >= 0
                If tt(j) >= a Then Exit Do
                tt(j + 1) = tt(j): mm(j + 1) = mm(j): j = j - 1
            Loop
            tt(j + 1) = a: mm(j + 1) = b
        Next i
    End If
    m = m & "SON 24 AY GÝRÝÞLERÝ (yeni " & ChrW(8594) & " eski)  •  stokta kalan pay" & vbCrLf
    If n = 0 Then m = m & "  Giriþ yok." & vbCrLf
    kalan = stok
    For i = 0 To n - 1
        al = mm(i)
        If al > kalan Then al = kalan
        If al < 0 Then al = 0
        kalan = kalan - al
        If tt(i) < s24 Then
            sinif = "ATIL"
        ElseIf tt(i) < s21 Then
            sinif = "3 AY KALDI"
        ElseIf tt(i) < s18 Then
            sinif = "6 AY KALDI"
        Else
            sinif = "normal"
        End If
        satir = "  " & Format$(CDate(tt(i)), "dd.mm.yyyy") & "   +" & ADM_ADET(mm(i)) & "   " & ChrW(8594) & " stokta " & ADM_ADET(al) & IIf(al > 0, "  (" & sinif & ")", "")
        If i < 8 Then m = m & satir & vbCrLf Else gizli = gizli + 1
    Next i
    If gizli > 0 Then m = m & "  ... ve " & gizli & " giriþ daha" & vbCrLf
    If kalan > 0.0000001 Then
        m = m & "  Giriþlerle karþýlanmayan: " & ADM_ADET(kalan) & " " & birim & "  " & ChrW(8594) & " 24 aydan eski (ATIL)" & vbCrLf
        If kr(2) > 0 Then
            m = m & "  2-5 yýl önceki son giriþ: " & Format$(CDate(Int(kr(2))), "dd.mm.yyyy") & vbCrLf
        Else
            m = m & "  Son 5 yýlda giriþ yok." & vbCrLf
        End If
    End If
    m = m & vbCrLf & "SONUÇ: " & ATIL_DURUM_ADI(durum) & "   •   ATIL " & ADM_ADET(q(3)) & "  •  3 AY " & ADM_ADET(q(2)) & "  •  6 AY " & ADM_ADET(q(1)) & _
        IIf(Len(gm) > 0, "   •   " & gm, "") & vbCrLf
    m = m & "Son 6 ay sarf: " & ADM_ADET(CDbl(kr(3))) & IIf(kr(4) > 0, "  (son sarf " & Format$(CDate(Int(kr(4))), "dd.mm.yyyy") & ")", "") & vbCrLf & vbCrLf
    m = m & "MB51: ÜY " & uy & ", ÝTü 101-106/501/561/301-306, " & Format$(DateAdd("m", -24, Date), "dd.mm.yyyy") & " - " & _
        Format$(Date, "dd.mm.yyyy") & " (311 depo transferi giriþ sayýlmaz)."
    If kart Then AKSORA_BEKLE_KAPAT
    MsgBox m, vbInformation, "AKSORA | ATIL KANIT"
    Exit Sub
Hata:
    If kart Then AKSORA_BEKLE_KAPAT
    MsgBox "Kanýt hazýrlanamadý: " & Err.Description, vbExclamation, "AKSORA | ATIL KANIT"
End Sub

Public Sub AKSORA_ATIL_GORUNUM()
    Dim ac As Boolean, no As Long, metin As String, gor As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error Resume Next
    gor = ThisWorkbook.Worksheets(ATIL_SAYFA).Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(gor) = 0 Then gor = "OZET"
    ac = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "Atýl stok sayfasý hazýrlanýyor...")
    On Error GoTo Son
    ATIL_GOSTER vbNullString, gor
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Atýl stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

' Transfer onerileri PDF: gecici cikti sayfasinda ozet + tablo (kurum yazicisina uygun acik gri / siyah),
' filtre uygulanmissa yalniz gorunen satirlar. Yatay, tek sayfa genisligi, baslik satiri her sayfada.
Public Sub AKSORA_ATIL_TRANSFER_PDF()
    Dim ws As Worksheet, cs As Worksheet, onceki As Object, adres As String, yol As String, il As String, ilAd As String, metin As String
    Dim v As Variant, a() As Variant, kol As Variant, hedef As Object, k As Variant, r As Long, n As Long, c As Long, ilk As Long
    Dim yuksek As Long, toplam As Double, ozet As String, ev As Boolean, sc As Boolean, da As Boolean, no As Long
    Dim anah() As String, idx() As Long, adlar() As String, kart As Boolean
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating: da = application.DisplayAlerts
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(ATIL_SAYFA)
    adres = CStr(ws.Range("BX22").Value2)
    If Len(adres) = 0 Or CStr(ws.Range("BX21").Value2) <> "TRANSFER" Then
        MsgBox "Önce TRANSFER ÖNERÝLERÝ sekmesini açýn.", vbInformation, "AKSORA | ATIL STOK": Exit Sub
    End If
    il = CStr(ws.Range("BX20").Value2)
    If il = "TUM" Or Len(il) = 0 Then ilAd = "Tüm Ýller" Else ilAd = Y_IL_ADI(il)
    ' Gorunen satirlar: Nereden, Malzeme, Tanim, Depo, Riskli miktar, Riskli deger, Durum, Nereye, Gonderilecek, Oncelik, Neden (Sistem cikmaz).
    v = ws.Range(adres).Value2
    ilk = ws.Range(adres).Row
    kol = Array(1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    ReDim a(1 To UBound(v, 1), 1 To 11)
    Set hedef = CreateObject("Scripting.Dictionary"): hedef.CompareMode = vbTextCompare
    For r = 2 To UBound(v, 1)
        If Not ws.Rows(ilk + r - 1).Hidden Then
            n = n + 1
            For c = 0 To 10: a(n, c + 1) = v(r, kol(c)): Next c
            toplam = toplam + ATIL_HUCRE_SAYI(v(r, 7))
            If CStr(v(r, 11)) = "YÜKSEK" Then yuksek = yuksek + 1
            hedef(CStr(v(r, 9))) = hedef(CStr(v(r, 9))) + 1
        End If
    Next r
    If n = 0 Then MsgBox "Yazdýrýlacak satýr yok (filtreyi kontrol edin).", vbInformation, "AKSORA | ATIL STOK": Exit Sub
    kart = AKSORA_BEKLE_AC("AKSORA  |  ATIL STOK", "PDF hazýrlanýyor (" & n & " satýr)...", 0.3)
    ' Nereye ozeti (cok olandan aza)
    ReDim anah(1 To hedef.Count): ReDim idx(1 To hedef.Count): ReDim adlar(1 To hedef.Count)
    r = 0
    For Each k In hedef.keys
        r = r + 1: idx(r) = r: adlar(r) = CStr(k)
        anah(r) = Format$(99999 - hedef(k), "00000") & CStr(k)
    Next k
    Y_SIRALA anah, idx, 1, hedef.Count
    For r = 1 To hedef.Count
        ozet = ozet & IIf(r > 1, "   •   ", "") & adlar(idx(r)) & ": " & hedef(adlar(idx(r))) & " kalem"
    Next r
    application.EnableEvents = False: application.ScreenUpdating = False: application.DisplayAlerts = False
    Set onceki = ActiveSheet
    Set cs = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
    With cs.Cells.Font
        .Name = "Aptos": .Size = 9: .Color = vbBlack
    End With
    cs.Range("A1").Value2 = "ATIL STOK TRANSFER ÖNERÝLERÝ   |   " & ilAd
    cs.Range("A1").Font.Size = 14: cs.Range("A1").Font.bold = True
    cs.Range("A2").Value2 = "Çoruh EDAÞ  •  Malzeme Stok Birimi  •  " & Format$(Now, "dd.mm.yyyy hh:nn") & _
        "  •  Atýl: 24 aydan eski stok (D, K ve yüklenici depolarý)  •  Öneri: ayný sistemde son 6 ayda bu malzemeyi kullanan il"
    cs.Range("A2").Font.Color = RGB(64, 64, 64)
    cs.Range("A3").Value2 = "ÖZET:  " & n & " öneri   •   Toplam riskli deðer: " & ADM_TL(toplam) & "   •   YÜKSEK öncelik: " & yuksek & "   •   DÜÞÜK öncelik: " & (n - yuksek)
    cs.Range("A3").Font.bold = True: cs.Range("A3").Font.Size = 10
    cs.Range("A4").Value2 = "NEREYE:  " & ozet
    cs.Range("A5").Value2 = "YÜKSEK: alacak ilin deposu 6 aylýk kullanýmýna yetmiyor; gönderilecek = eksik kadar.   DÜÞÜK: alacak ilin stoðu yetiyor."
    cs.Range("A5").Font.Italic = True: cs.Range("A5").Font.Color = RGB(64, 64, 64)
    cs.Range("A7:K7").Value2 = Array("Nereden", "Malzeme", "Taným", "Depo", "Riskli miktar", "Riskli deðer", "Durum", "Nereye", "Gönderilecek", "Öncelik", "Neden")
    With cs.Range("A7:K7")
        .Font.bold = True: .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter: .WrapText = True
    End With
    cs.Rows(7).RowHeight = 24
    cs.Range(cs.Cells(8, 2), cs.Cells(7 + n, 2)).NumberFormat = "@"
    cs.Range("A8").Resize(n, 11).Value2 = a
    With cs.Range(cs.Cells(7, 1), cs.Cells(7 + n, 11))
        .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin: .Borders.Color = RGB(166, 166, 166)
        .VerticalAlignment = xlCenter
    End With
    cs.Range(cs.Rows(8), cs.Rows(7 + n)).RowHeight = 15
    cs.Range(cs.Cells(8, 4), cs.Cells(7 + n, 4)).WrapText = True
    ' Kesirsiz miktarda virgul gosterilmez (ekrandaki tablolarla ayni bicim).
    AKSORA_V13_MIKTAR_BICIMI cs.Range(cs.Cells(8, 5), cs.Cells(7 + n, 5))
    cs.Range(cs.Cells(8, 6), cs.Cells(7 + n, 6)).NumberFormat = "#,##0.00 ""TL"""
    cs.Range(cs.Cells(8, 1), cs.Cells(7 + n, 2)).HorizontalAlignment = xlCenter
    cs.Range(cs.Cells(8, 4), cs.Cells(7 + n, 4)).HorizontalAlignment = xlCenter
    cs.Range(cs.Cells(8, 7), cs.Cells(7 + n, 10)).HorizontalAlignment = xlCenter
    cs.Range(cs.Cells(8, 7), cs.Cells(7 + n, 10)).Font.bold = True
    cs.Range(cs.Cells(7, 6), cs.Cells(7 + n, 6)).HorizontalAlignment = xlRight
    kol = Array(9, 11, 32, 20, 9, 13, 10, 10, 11, 8, 41)
    For c = 0 To 10: cs.Columns(c + 1).ColumnWidth = kol(c): Next c
    For r = 8 To 7 + n
        If Len(CStr(cs.Cells(r, 4).Value2)) > 20 Then cs.Rows(r).AutoFit
    Next r
    yol = AKSORA_KAYIT_KLASORU() & application.PathSeparator & "ATIL_TRANSFER_" & IIf(il = "TUM" Or Len(il) = 0, "TUM_ILLER", il) & "_" & Format$(Now, "yyyymmdd_hhnn") & ".pdf"
    On Error Resume Next
    With cs.PageSetup
        .PrintArea = cs.Range(cs.Cells(1, 1), cs.Cells(7 + n, 11)).Address
        .PrintTitleRows = "$7:$7"
        .Orientation = xlLandscape
        .Zoom = False: .FitToPagesWide = 1: .FitToPagesTall = False
        .LeftMargin = application.InchesToPoints(0.3): .RightMargin = application.InchesToPoints(0.3)
        .TopMargin = application.InchesToPoints(0.4): .BottomMargin = application.InchesToPoints(0.5)
        .CenterFooter = "Sayfa &P / &N"
        .LeftFooter = "AKSORA  •  Atýl Stok Transfer Önerileri  •  " & ilAd
    End With
    On Error GoTo Hata
    cs.ExportAsFixedFormat Type:=xlTypePDF, Filename:=yol, Quality:=xlQualityStandard, IncludeDocProperties:=False, _
        IgnorePrintAreas:=False, OpenAfterPublish:=True
Cikis:
    On Error Resume Next
    If Not cs Is Nothing Then cs.Delete
    If Not onceki Is Nothing Then onceki.Activate
    If kart Then AKSORA_BEKLE_KAPAT
    application.DisplayAlerts = da: application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If no <> 0 Then MsgBox "PDF oluþturulamadý: " & metin, vbExclamation, "AKSORA | ATIL STOK"
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    Resume Cikis
End Sub

' Malzemenin bulundugu tum depolar ve miktarlari: "D001 (3), K002 (5)".
Private Function ATIL_DEPO_METIN(ByVal dp As Object) As String
    Dim k As Variant, n As Long
    For Each k In dp.keys
        n = n + 1
        ATIL_DEPO_METIN = ATIL_DEPO_METIN & IIf(n > 1, ", ", "") & CStr(k) & " (" & ADM_ADET(CDbl(dp(k)(1))) & ")"
    Next k
End Function

Public Function AKSORA_ATIL_SAYFA(ByVal sh As Object) As Boolean
    If Not TypeOf sh Is Worksheet Then Exit Function
    If sh.Name <> ATIL_SAYFA Then Exit Function
    AKSORA_ATIL_SAYFA = (CStr(sh.Range("BX1").Value2) = ATIL_ETIKET)
End Function

' Sayfa musteri rapor sayfasinin basligiyla kurulur (veri tasinmaz).
Private Function ATIL_SAYFA_KUR() As Worksheet
    Dim ws As Worksheet, kaynak As Worksheet, gor As XlSheetVisibility, onceki As Collection, sh As Object, varmi As Boolean
    Dim da As Boolean
    Set ws = ADM_WS(ATIL_SAYFA)
    If Not ws Is Nothing Then
        If CStr(ws.Range("BX1").Value2) <> ATIL_ETIKET Then Err.Raise vbObjectError + 2360, , "Ayný isimde baþka bir sayfa var: " & ATIL_SAYFA
        Set ATIL_SAYFA_KUR = ws
        Exit Function
    End If
    Set kaynak = ADM_WS("M_CDP_RAPOR")
    If kaynak Is Nothing Then Err.Raise vbObjectError + 2361, , "Sayfa þablonu (M_CDP_RAPOR) bulunamadý."
    da = application.DisplayAlerts: application.DisplayAlerts = False
    gor = kaynak.Visible
    kaynak.Visible = xlSheetVisible
    Set onceki = New Collection
    For Each sh In ThisWorkbook.Sheets
        onceki.Add sh.Name, sh.Name
    Next sh
    kaynak.Copy after:=kaynak
    For Each sh In ThisWorkbook.Sheets
        varmi = False
        On Error Resume Next
        varmi = (Len(onceki(sh.Name)) > 0)
        On Error GoTo 0
        If Not varmi Then Set ws = sh: Exit For
    Next sh
    kaynak.Visible = gor
    application.DisplayAlerts = da
    If ws Is Nothing Then Err.Raise vbObjectError + 2362, , "Kopya sayfa bulunamadý."
    ws.Name = ATIL_SAYFA
    ws.Unprotect ""
    ws.Range("BX1:BX40").ClearContents
    ws.Range("BX1").Value2 = ATIL_ETIKET
    On Error Resume Next
    ws.Move after:=ThisWorkbook.Worksheets("YÜKLENÝCÝ_RAPOR")
    On Error GoTo 0
    Set ATIL_SAYFA_KUR = ws
End Function

Private Sub ATIL_DURUM_RENK(ByVal r As Range)
    Y_DURUM_KF r, "ATIL", RGB(220, 38, 38), RGB(255, 255, 255)
    Y_DURUM_KF r, "3 AY KALDI", RGB(254, 215, 170), RGB(154, 52, 18)
    Y_DURUM_KF r, "6 AY KALDI", RGB(254, 243, 199), RGB(146, 64, 14)
    r.HorizontalAlignment = xlCenter
End Sub

Private Sub ATIL_CIZ(ByVal ws As Worksheet, ByVal ilSec As String)
    Dim c As Collection, bilgi As Variant, x As Variant, w As Double, s As Shape, y As Double, kw As Double, bw As Double
    Dim i As Long, j As Long, r As Long, a1 As Long, n As Long, k As Variant, o As Variant, t As Variant, il As Variant
    Dim kpiT(0 To 3) As Double, kpiA(0 To 3) As Long, oneriA As Long, oneriT As Double, kats As Variant, katD As Object
    Dim karne As Object, kapsam As Object, dp As Object, dk As Variant, anah() As String, idx() As Long, satir() As Variant
    Dim liste As Collection, sw As Double, enAz As Double, enCok As Double, ayrim As Boolean
    Dim gec As Boolean, okMetin As String, kod As String, tablo() As Variant
    Dim gor As String, yuksek As Long, gAd As Variant, gKod As Variant, bw2 As Double, t2 As Long, pay As Double
    ws.Unprotect ""
    Y_SAYFA_SIFIRLA ws
    ' Sablondan gelen AKSORA amblemi (filigran) bu sayfada kullanilmaz.
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(ws.Shapes(i).Name, "FILIGRAN") > 0 Then ws.Shapes(i).Delete
    Next i
    Set s = Y_SEKIL(ws, "btnAnaPanelDon")
    If Not s Is Nothing Then s.AlternativeText = vbNullString: s.OnAction = BIRIM_MAKRO("ANA_PANELE_DON")
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | ATIL STOK  |  " & IIf(ilSec = "TUM", "Tüm Ýller", Y_IL_ADI(ilSec))
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    sw = application.Max(170, w - 1285)
    Y_KOLONLAR ws, Array(90, 60, 95, 230, 140, 95, 115, 105, 125, 105, 125, sw)
    ws.Range("BX20").Value2 = ilSec
    gor = CStr(ws.Range("BX21").Value2)
    ' Varsayilan gorunum NE YAPILMALI; transfer ayrintisi TRANSFER LISTESI sayfasinda.
    If gor <> "LISTE" And gor <> "OZET" Then gor = "ISLEM"
    Set c = ATIL_HESAPLA(bilgi)
    ' kapsam degeri (atil orani icin) ve karne
    Set kapsam = CreateObject("Scripting.Dictionary"): kapsam.CompareMode = vbTextCompare
    Set karne = CreateObject("Scripting.Dictionary"): karne.CompareMode = vbTextCompare
    Set katD = CreateObject("Scripting.Dictionary"): katD.CompareMode = vbTextCompare
    Set liste = New Collection
    For Each x In c
        If ilSec = "TUM" Or x(0) = ilSec Then
            ' karne anahtari: tum illerde il, tek ilde depo
            If ilSec = "TUM" Then
                kod = x(0)
                kapsam(kod) = kapsam(kod) + x(8)
            Else
                Set dp = x(15)
                For Each dk In dp.keys
                    kapsam(x(1) & "|" & dk) = kapsam(x(1) & "|" & dk) + dp(dk)(2)
                    If Not karne.Exists(x(1) & "|" & dk) Then karne.Add x(1) & "|" & dk, Array(CStr(dp(dk)(0)), 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                Next dk
            End If
            If x(10) >= 1 Then
                ' Bir malzemenin stogu birden fazla basamakta olabilir (ornek: bir kismi ATIL, bir kismi 3 AY):
                ' her basamak kendi miktar ve degeriyle sayilir.
                For t2 = 1 To 3
                    If x(20)(t2) > 0.0000001 Then kpiT(t2) = kpiT(t2) + x(19)(t2): kpiA(t2) = kpiA(t2) + 1
                Next t2
                If Len(x(13)) > 0 Then oneriA = oneriA + 1: oneriT = oneriT + x(22)
                If x(18) = "YÜKSEK" Then yuksek = yuksek + 1
                liste.Add x
                ' kategori: Array(atil kalem, atil deger, 3ay kalem, 3ay deger, 6ay kalem, 6ay deger, stogu dusen, riskli kalem)
                If Not katD.Exists(x(5)) Then katD.Add x(5), Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                o = katD(x(5))
                For t2 = 1 To 3
                    If x(20)(t2) > 0.0000001 Then j = (3 - t2) * 2: o(j) = o(j) + 1: o(j + 1) = o(j + 1) + x(19)(t2)
                Next t2
                If x(16) Then o(6) = o(6) + 1
                o(7) = o(7) + 1
                katD(x(5)) = o
                ' karne: Array(ad, atil kalem, atil deger, 3ay kalem, 3ay deger, 6ay kalem, 6ay deger, stogu dusen, riskli kalem)
                If ilSec = "TUM" Then
                    If Not karne.Exists(kod) Then karne.Add kod, Array(Y_IL_ADI(kod), 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
                    o = karne(kod)
                    For t2 = 1 To 3
                        If x(20)(t2) > 0.0000001 Then j = (3 - t2) * 2: o(1 + j) = o(1 + j) + 1: o(2 + j) = o(2 + j) + x(19)(t2)
                    Next t2
                    If x(16) Then o(7) = o(7) + 1
                    o(8) = o(8) + 1
                    karne(kod) = o
                Else
                    ' Depo karnesi: riskli deger depolara stok payiyla dagitilir.
                    Set dp = x(15)
                    For Each dk In dp.keys
                        o = karne(x(1) & "|" & dk)
                        If x(6) > 0 Then pay = dp(dk)(1) / x(6) Else pay = 0
                        For t2 = 1 To 3
                            If x(20)(t2) > 0.0000001 Then j = (3 - t2) * 2: o(1 + j) = o(1 + j) + 1: o(2 + j) = o(2 + j) + x(19)(t2) * pay
                        Next t2
                        If x(16) Then o(7) = o(7) + 1
                        o(8) = o(8) + 1
                        karne(x(1) & "|" & dk) = o
                    Next dk
                End If
            End If
        End If
    Next x
    If ilSec = "TUM" Then
        For Each il In ADM_ILLER()
            If Not karne.Exists(CStr(il)) Then karne.Add CStr(il), Array(Y_IL_ADI(CStr(il)), 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        Next il
    End If
    gec = (bilgi(0) > 0 Or bilgi(1) > 0)

    '--- KPI
    ws.Range("BZ21").Value2 = ADM_TL_KISA(kpiT(3)) & "  •  " & kpiA(3) & " kalem"
    ws.Range("BZ22").Value2 = ADM_TL_KISA(kpiT(2)) & "  •  " & kpiA(2) & " kalem"
    ws.Range("BZ23").Value2 = ADM_TL_KISA(kpiT(1)) & "  •  " & kpiA(1) & " kalem"
    ' 4. kart: NE YAPILMALI hesabindan satis / hurda / iade adayi (hicbir ilde kullanilmayan atil).
    t = Array(0#, 0#)
    For Each x In ATIL_ISLEM_HESAPLA(c, ilSec)
        If x(0) = 5 Then t(0) = t(0) + 1: t(1) = t(1) + x(7)
    Next x
    ws.Range("BZ24").Value2 = AKSORA_SAYI_METNI(Round(t(1) / 1000000, 1), "#,##0.0") & " mn TL  •  " & t(0) & " kalem"
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10: ws.Rows(8).RowHeight = 30: ws.Rows(9).RowHeight = 20: ws.Rows(10).RowHeight = 30: ws.Rows(11).RowHeight = 8
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "ATIL (24 AYDAN ESKÝ STOK)", "$BZ$21", RGB(220, 38, 38)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "3 AY KALDI (21-24 AY)", "$BZ$22", RGB(234, 88, 12)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "6 AY KALDI (18-21 AY)", "$BZ$23", RGB(202, 138, 4)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "SAT / HURDA / ÝADE ADAYI", "$BZ$24", RGB(124, 58, 237)
    For i = 1 To 4
        Set s = Y_SEKIL(ws, "Y_KPI_DEGER_" & i)
        If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 15
    Next i

    '--- il secimi ve islemler
    y = ws.Rows(8).Top + 1
    bw = application.Min(118, (w - 2 * 150 - 7 * 6) / 6)
    i = 0
    For Each il In Array("TUM", "ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
        Set s = Y_DUGME(ws, "ATIL_IL_" & i, IIf(CStr(il) = "TUM", "TÜM ÝLLER", ATIL_IL_BUYUK(CStr(il))), i * (bw + 6), y, bw, 28, _
            IIf(CStr(il) = ilSec, RGB(0, 125, 110), RGB(148, 163, 184)), "AKSORA_ATIL_IL_SEC")
        s.AlternativeText = CStr(il)
        i = i + 1
    Next il
    Y_DUGME ws, "ATIL_YENILE", "YENÝLE", w - 300 - 6, y, 150, 28, RGB(20, 103, 173), "AKSORA_ATIL_AC"
    Y_DUGME ws, "ATIL_SAP", "SAP'DEN OKU (MB51)", w - 150, y, 150, 28, RGB(12, 45, 79), "AKSORA_ATIL_SAP_OKU"
    If gec Then
        okMetin = "Atýl: eldeki stoðun 24 aydan eski giriþe ait kýsmý  •  D, K ve yüklenici depolarý   |   SAP okumasý: CDP " & _
            IIf(bilgi(0) > 0, Format$(bilgi(0), "dd.mm.yyyy"), "-") & "  •  EPP " & IIf(bilgi(1) > 0, Format$(bilgi(1), "dd.mm.yyyy"), "-") & _
            "   |   Stok verisi: " & bilgi(4)
        If bilgi(3) > 0 Then okMetin = okMetin & "   |   Stoðu düþen: " & bilgi(3) & " kalem"
        If bilgi(2) > 0 Then okMetin = okMetin & "   |   SAP okumasýndan sonra gelen " & bilgi(2) & " kalem sonraki okumada deðerlendirilecek"
    ElseIf bilgi(5) Then
        okMetin = "Atýl kuralý deðiþti (giriþ yaþý + yüklenici depolarý). Saðdaki 'SAP'DEN OKU (MB51)' ile yeniden okuyun (SAP'de CDP ve EPP oturumu açýk olmalý)."
    Else
        okMetin = "Henüz SAP okumasý yok. Saðdaki 'SAP'DEN OKU (MB51)' düðmesine basýn (SAP'de CDP ve EPP oturumu açýk olmalý)."
    End If
    ws.Range("A9").Value2 = okMetin
    With ws.Range("A9").Font
        .Name = "Aptos": .Size = 9: .Italic = True: .Color = IIf(gec, RGB(100, 116, 139), RGB(185, 28, 28)): .bold = Not gec
    End With
    ws.Range("A9").VerticalAlignment = xlCenter
    If gec Then AN_YAS_YAZ ws, AN_EN_ESKI(CDbl(bilgi(0)), CDbl(bilgi(1))), "ATIL STOK (MB51)"
    If Not gec Then GoTo Son

    '--- gorunum sekmeleri: ozet / transfer / liste (tek bolum gosterilir, sayfa kisa kalir)
    gAd = Array("NE YAPILMALI", "ÖZET  •  kategori ve karne", "ATIL LÝSTESÝ  •  " & liste.Count & " kalem", "TRANSFER LÝSTESÝ  " & ChrW(8594))
    gKod = Array("ISLEM", "OZET", "LISTE", "GIT")
    bw2 = (w - 3 * 8) / 4
    y = ws.Rows(10).Top + 1
    For i = 0 To 3
        If CStr(gKod(i)) = "GIT" Then
            Set s = Y_DUGME(ws, "ATIL_GOR_" & i, CStr(gAd(i)), i * (bw2 + 8), y, bw2, 28, RGB(39, 160, 94), "AKSORA_TRANSFER_AC")
            s.TextFrame2.TextRange.Font.Size = 11
        Else
            Set s = Y_DUGME(ws, "ATIL_GOR_" & i, CStr(gAd(i)), i * (bw2 + 8), y, bw2, 28, _
                IIf(CStr(gKod(i)) = gor, RGB(12, 45, 79), RGB(226, 232, 240)), "AKSORA_ATIL_GORUNUM")
            s.TextFrame2.TextRange.Font.Size = 11
            If CStr(gKod(i)) <> gor Then s.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(12, 45, 79)
        End If
        s.AlternativeText = CStr(gKod(i))
    Next i
    If gor = "ISLEM" Then
        ATIL_ISLEM_CIZ ws, c, ilSec
        GoTo Son
    End If
    If gor <> "OZET" Then GoTo Ayrinti

    '--- kategori ozeti
    r = 12
    Y_BOLUM ws, r, "KATEGORÝ ÖZETÝ", "Atýl ve atýla yaklaþan malzemelerin kategorilere daðýlýmý.", 12
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array("Kategori", "Kalem", "ATIL kalem", "ATIL deðeri", "3 AY kalem", _
        "3 AY deðeri", "6 AY kalem", "6 AY deðeri", "Toplam risk deðeri", "Risk payý", "Stoðu düþen", "Yorum")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    kats = Array("KABLO", "DÝREK", "PANO", "HÜCRE", "TRAFO", "DÝÐER")
    enCok = kpiT(1) + kpiT(2) + kpiT(3)
    For i = 0 To 5
        r = r + 1
        ws.Cells(r, 1).Value2 = kats(i)
        If katD.Exists(kats(i)) Then o = katD(kats(i)) Else o = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#)
        ws.Cells(r, 2).Value2 = o(7)
        ws.Cells(r, 3).Value2 = o(0): ws.Cells(r, 4).Value2 = Round(o(1), 2)
        ws.Cells(r, 5).Value2 = o(2): ws.Cells(r, 6).Value2 = Round(o(3), 2)
        ws.Cells(r, 7).Value2 = o(4): ws.Cells(r, 8).Value2 = Round(o(5), 2)
        ws.Cells(r, 9).Value2 = Round(o(1) + o(3) + o(5), 2)
        If enCok > 0 Then ws.Cells(r, 10).Value2 = (o(1) + o(3) + o(5)) / enCok Else ws.Cells(r, 10).Value2 = 0
        ws.Cells(r, 11).Value2 = o(6)
        If o(7) = 0 Then
            ws.Cells(r, 12).Value2 = "Riskli malzeme yok"
        ElseIf o(1) >= o(3) And o(1) >= o(5) Then
            ws.Cells(r, 12).Value2 = "Aðýrlýk atýlda: transfer / ihtiyaç kontrolü"
        Else
            ws.Cells(r, 12).Value2 = "Atýla yaklaþýyor: kullaným planlayýn"
        End If
    Next i
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12))
    ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 22
    r = r + 1
    ws.Cells(r, 1).Value2 = "TOPLAM"
    For j = 2 To 9
        ws.Cells(r, j).Formula = "=SUM(" & ws.Range(ws.Cells(a1, j), ws.Cells(r - 1, j)).Address(False, False) & ")"
    Next j
    ws.Cells(r, 10).Value2 = IIf(enCok > 0, 1, 0)
    ws.Cells(r, 11).Formula = "=SUM(" & ws.Range(ws.Cells(a1, 11), ws.Cells(r - 1, 11)).Address(False, False) & ")"
    Y_TOPLAM_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12))
    ws.Rows(r).RowHeight = 22
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).Font.bold = True
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).IndentLevel = 1
    ws.Range(ws.Cells(a1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 8), ws.Cells(r, 9)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(a1, 3), ws.Cells(r, 4)).Font.Color = RGB(185, 28, 28)
    I_FIRE_ISI ws.Range(ws.Cells(a1, 9), ws.Cells(r - 1, 9))

    '--- il / depo karnesi
    r = r + 2
    If ilSec = "TUM" Then
        Y_BOLUM ws, r, "ÝL KARNESÝ", "Atýl oraný = atýl deðer / kapsamdaki (D, K ve yüklenici depolarý) stok deðeri. Sýra: atýl oraný düþükten yükseðe.", 12
    Else
        Y_BOLUM ws, r, "DEPO KARNESÝ  –  " & Y_IL_ADI(ilSec), "Riskli malzemelerin depolara daðýlýmý. Sýra: toplam risk deðeri yüksekten düþüðe.", 12
    End If
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array(IIf(ilSec = "TUM", "Ýl", "Depo"), IIf(ilSec = "TUM", "Kalem", "Sistem"), _
        "ATIL kalem", "ATIL deðeri", "3 AY kalem", "3 AY deðeri", "6 AY kalem", "6 AY deðeri", "Toplam risk deðeri", "Atýl oraný", _
        "Stoðu düþen", IIf(ilSec = "TUM", "Deðerlendirme", "Depo adý"))
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    n = karne.Count
    If n > 0 Then
        ReDim anah(1 To n): ReDim idx(1 To n): ReDim satir(1 To n)
        i = 0
        For Each k In karne.keys
            i = i + 1: idx(i) = i: satir(i) = k
            o = karne(k)
            If ilSec = "TUM" Then
                If kapsam(CStr(k)) > 0 Then t = o(2) / kapsam(CStr(k)) Else t = 0
                anah(i) = Format$(Round(t * 1000000, 0), "000000000000") & CStr(k)
            Else
                anah(i) = Format$(999999999999# - Round(o(2) + o(4) + o(6), 0), "000000000000") & CStr(k)
            End If
        Next k
        Y_SIRALA anah, idx, 1, n
        enAz = -1: enCok = -1
        For i = 1 To n
            k = satir(idx(i)): o = karne(k)
            r = r + 1
            If kapsam(CStr(k)) > 0 Then t = o(2) / kapsam(CStr(k)) Else t = 0
            If ilSec = "TUM" Then
                ws.Cells(r, 1).Value2 = o(0)
                ws.Cells(r, 2).Value2 = o(8)
                If enAz < 0 Then enAz = t
                enCok = t
            Else
                ws.Cells(r, 1).NumberFormat = "@": ws.Cells(r, 1).Value2 = Mid$(CStr(k), InStr(CStr(k), "|") + 1)
                ws.Cells(r, 2).Value2 = Left$(CStr(k), InStr(CStr(k), "|") - 1)
                ws.Cells(r, 12).Value2 = o(0)
            End If
            ws.Cells(r, 3).Value2 = o(1): ws.Cells(r, 4).Value2 = Round(o(2), 2)
            ws.Cells(r, 5).Value2 = o(3): ws.Cells(r, 6).Value2 = Round(o(4), 2)
            ws.Cells(r, 7).Value2 = o(5): ws.Cells(r, 8).Value2 = Round(o(6), 2)
            ws.Cells(r, 9).Value2 = Round(o(2) + o(4) + o(6), 2)
            ws.Cells(r, 10).Value2 = t
            ws.Cells(r, 11).Value2 = o(7)
        Next i
        If ilSec = "TUM" Then
            ayrim = (enCok > enAz + 0.0005)
            For i = 1 To n
                t = ws.Cells(a1 + i - 1, 10).Value2
                If i = 1 And ayrim Then
                    ws.Cells(a1 + i - 1, 12).Value2 = ChrW(9733) & " EN ÝYÝ"
                ElseIf i = n And ayrim Then
                    ws.Cells(a1 + i - 1, 12).Value2 = ChrW(9660) & " EN KÖTÜ"
                ElseIf t < 0.05 Then
                    ws.Cells(a1 + i - 1, 12).Value2 = "ÝYÝ"
                ElseIf t < 0.15 Then
                    ws.Cells(a1 + i - 1, 12).Value2 = "ORTA"
                Else
                    ws.Cells(a1 + i - 1, 12).Value2 = "ZAYIF"
                End If
            Next i
        End If
        Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12))
        ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 22
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).Font.bold = True
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ws.Cells(a1, 8), ws.Cells(r, 9)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).NumberFormat = "0.0%"
        ws.Range(ws.Cells(a1, 3), ws.Cells(r, 4)).Font.Color = RGB(185, 28, 28)
        I_FIRE_ISI ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9))
        If ilSec = "TUM" Then I_DURUM_RENK ws.Range(ws.Cells(a1, 12), ws.Cells(r, 12))
    End If

    GoTo Son

Ayrinti:
    n = liste.Count
    If gor = "TRANSFER" Then
        ' Siralama: oncelik (YUKSEK once), durum (ATIL once), deger.
        j = 0
        For Each x In liste
            If Len(x(13)) > 0 Then j = j + 1
        Next x
        r = 12
        Y_BOLUM ws, r, "TRANSFER ÖNERÝLERÝ (" & j & ")", "Ayný sistemde (CDP-CDP, EPP-EPP). YÜKSEK: alacak ilin deposu 6 aylýk kullanýmýna yetmiyor.", 10
        ws.Rows(r).RowHeight = 30
        r = r + 1
        Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array("Nereden", "Sistem", "Malzeme", "Taným", "Depo", _
            "Riskli miktar", "Riskli deðer", "Durum", "Nereye", "Gönderilecek", "Öncelik", "Neden")
        ws.Rows(r).RowHeight = 30
        a1 = r
        If j = 0 Then
            r = r + 1
            ws.Cells(r, 1).Value2 = "Öneri yok: riskli malzemeleri ayný sistemde son 6 ayda kullanan baþka il bulunmadý."
            ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(100, 116, 139)
        Else
            ReDim anah(1 To j): ReDim idx(1 To j): ReDim satir(1 To j)
            i = 0
            For Each x In liste
                If Len(x(13)) > 0 Then
                    i = i + 1: idx(i) = i: satir(i) = x
                    anah(i) = IIf(x(18) = "YÜKSEK", "0", "1") & CStr(9 - x(10)) & Format$(999999999999# - Round(x(22), 0), "000000000000") & x(0) & x(3)
                End If
            Next x
            Y_SIRALA anah, idx, 1, j
            ReDim tablo(1 To j, 1 To 12)
            For i = 1 To j
                x = satir(idx(i))
                tablo(i, 1) = Y_IL_ADI(CStr(x(0))): tablo(i, 2) = x(1): tablo(i, 3) = CStr(x(3)): tablo(i, 4) = x(4)
                tablo(i, 5) = ATIL_DEPO_METIN(x(15))
                tablo(i, 6) = x(21): tablo(i, 7) = Round(x(22), 2)
                tablo(i, 8) = ATIL_DURUM_ADI(CLng(x(10)))
                tablo(i, 9) = Y_IL_ADI(CStr(x(13)))
                tablo(i, 10) = ADM_ADET(CDbl(x(17))) & " " & x(7)
                tablo(i, 11) = x(18)
                tablo(i, 12) = x(14)
            Next i
            ATIL_TABLO_YAZ ws, a1 + 1, tablo
            r = a1 + j
            Y_TABLO_GOVDE ws.Range(ws.Cells(a1 + 1, 1), ws.Cells(r, 12))
            ws.Range(ws.Rows(a1 + 1), ws.Rows(r)).RowHeight = 20
            ws.Range(ws.Cells(a1 + 1, 1), ws.Cells(r, 1)).Font.bold = True
            ws.Range(ws.Cells(a1 + 1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(a1 + 1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
            ' Depo: tum depolar gorunur; uzun listede satir yuksekligi metne gore acilir.
            ws.Range(ws.Cells(a1 + 1, 5), ws.Cells(r, 5)).WrapText = True
            For i = a1 + 1 To r
                If Len(CStr(ws.Cells(i, 5).Value2)) > 24 Then
                    ws.Rows(i).AutoFit
                    If ws.Rows(i).RowHeight < 20 Then ws.Rows(i).RowHeight = 20
                End If
            Next i
            ws.Range(ws.Cells(a1 + 1, 9), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(a1 + 1, 4), ws.Cells(r, 4)).WrapText = False
            ws.Range(ws.Cells(a1 + 1, 12), ws.Cells(r, 12)).WrapText = False
            AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1 + 1, 6), ws.Cells(r, 6))
            ws.Range(ws.Cells(a1 + 1, 7), ws.Cells(r, 7)).NumberFormat = "#,##0.00 ""TL"""
            ATIL_DURUM_RENK ws.Range(ws.Cells(a1 + 1, 8), ws.Cells(r, 8))
            With ws.Range(ws.Cells(a1 + 1, 9), ws.Cells(r, 10))
                .Font.bold = True: .Font.Color = RGB(20, 103, 173)
            End With
            Y_DURUM_KF ws.Range(ws.Cells(a1 + 1, 11), ws.Cells(r, 11)), "YÜKSEK", RGB(220, 252, 231), RGB(21, 128, 61)
            Y_DURUM_KF ws.Range(ws.Cells(a1 + 1, 11), ws.Cells(r, 11)), "DÜÞÜK", RGB(241, 245, 249), RGB(100, 116, 139)
            ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12)).AutoFilter
            ' PDF cikti: basliktan son satira (filtrede gizlenen satirlar basilmaz).
            ws.Range("BX22").Value2 = ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12)).Address
            Set s = Y_DUGME(ws, "ATIL_PDF", "PDF ÇIKTI", w - 150, ws.Rows(12).Top + 1, 150, 28, RGB(15, 118, 110), "AKSORA_ATIL_TRANSFER_PDF")
            Set s = Y_DUGME(ws, "ATIL_KANIT", "SEÇÝLÝ MALZEME: KANIT", w - 150 - 8 - 190, ws.Rows(12).Top + 1, 190, 28, RGB(20, 103, 173), "AKSORA_ATIL_KANIT")
        End If
    Else
        ' tum riskli malzemeler
        r = 12
        Y_BOLUM ws, r, "ATIL VE ATILA YAKLAÞAN MALZEMELER (" & n & ")", "Riskli miktar: stoðun 18 aydan eski giriþlere ait kýsmý. Sýra: durum (ATIL önce), sonra riskli deðer. Baþlýklardan filtrelenebilir.", 12
        ws.Rows(r).RowHeight = 30
        r = r + 1
        Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array("Ýl", "Sistem", "Malzeme", "Taným", "Kategori", _
            "Stok", "Riskli miktar", "Riskli deðer", "Riskli giriþ", "Gün sayacý", "Durum", "Transfer önerisi")
        ws.Rows(r).RowHeight = 30
        a1 = r
        If n = 0 Then
            r = r + 1
            ws.Cells(r, 1).Value2 = "Atýl ya da atýla yaklaþan malzeme yok."
            ws.Cells(r, 1).Font.Italic = True: ws.Cells(r, 1).Font.Color = RGB(100, 116, 139)
        Else
            ReDim anah(1 To n): ReDim idx(1 To n): ReDim satir(1 To n)
            i = 0
            For Each x In liste
                i = i + 1: idx(i) = i: satir(i) = x
                anah(i) = CStr(9 - x(10)) & Format$(999999999999# - Round(x(22), 0), "000000000000") & x(0) & x(3)
            Next x
            Y_SIRALA anah, idx, 1, n
            ReDim tablo(1 To n, 1 To 12)
            For i = 1 To n
                x = satir(idx(i))
                ATIL_SATIR tablo, i, x
                If Len(x(13)) > 0 Then tablo(i, 12) = ChrW(8594) & " " & Y_IL_ADI(CStr(x(13))) & "  (" & x(18) & ")"
            Next i
            ATIL_TABLO_YAZ ws, a1 + 1, tablo
            r = a1 + n
            ATIL_TABLO_BICIM ws, a1 + 1, r
            ws.Range(ws.Cells(a1 + 1, 11), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
            ws.Range(ws.Cells(a1 + 1, 12), ws.Cells(r, 12)).Font.Color = RGB(20, 103, 173)
            ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12)).AutoFilter
            Set s = Y_DUGME(ws, "ATIL_KANIT", "SEÇÝLÝ MALZEME: KANIT", w - 190, ws.Rows(12).Top + 1, 190, 28, RGB(20, 103, 173), "AKSORA_ATIL_KANIT")
        End If
    End If
Son:
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Private Sub ATIL_SATIR(ByRef tablo() As Variant, ByVal r As Long, ByVal x As Variant)
    tablo(r, 1) = Y_IL_ADI(CStr(x(0)))
    tablo(r, 2) = x(1)
    tablo(r, 3) = CStr(x(3))
    tablo(r, 4) = x(4)
    tablo(r, 5) = x(5)
    tablo(r, 6) = x(6)
    tablo(r, 7) = x(21)
    tablo(r, 8) = Round(x(22), 2)
    tablo(r, 9) = CStr(x(9))
    tablo(r, 10) = x(11)
    tablo(r, 11) = ATIL_DURUM_ADI(CLng(x(10)))
End Sub

' Malzeme kodu ve tarih metin olarak yazilir (bastaki sifirlar / tarih donusumu olmaz).
Private Sub ATIL_TABLO_YAZ(ByVal ws As Worksheet, ByVal a1 As Long, ByRef tablo() As Variant)
    Dim m As Long
    m = UBound(tablo, 1)
    ws.Range(ws.Cells(a1, 3), ws.Cells(a1 + m - 1, 3)).NumberFormat = "@"
    ws.Range(ws.Cells(a1, 9), ws.Cells(a1 + m - 1, 9)).NumberFormat = "@"
    ws.Cells(a1, 1).Resize(m, 12).Value2 = tablo
End Sub

Private Sub ATIL_TABLO_BICIM(ByVal ws As Worksheet, ByVal a1 As Long, ByVal r As Long)
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12))
    ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 20
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).Font.bold = True
    ws.Range(ws.Cells(a1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 9), ws.Cells(r, 10)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).WrapText = False
    ws.Range(ws.Cells(a1, 12), ws.Cells(r, 12)).WrapText = False
    AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6))
    AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7))
    ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).Font.bold = True
    ws.Range(ws.Cells(a1, 8), ws.Cells(r, 8)).NumberFormat = "#,##0.00 ""TL"""
    ATIL_DURUM_RENK ws.Range(ws.Cells(a1, 11), ws.Cells(r, 11))
End Sub

' Bir sistemin bes uretim yeri. Her uretim yerinde: 1) son 24 ayin girisleri (kendi stogu), 2) stogu bu girislerden
' fazla olanlarda 24-60 ay onceki son giris, 3) son 6 ay sarfi (sistemin tum malzemeleri; transfer onerisi icin).
' Basarida bos, hatada "adim | hata" doner (diger sistem etkilenmez).
Private Function ATIL_SISTEM_OKU(ByVal ses As Object, ByVal sis As String, ByVal tum As Object, ByVal sonuc As Object, _
    ByVal bugun As Date, ByRef adet As Long, ByVal toplam As Long) As String
    Dim il As Variant, s As Object, k As Variant, uy As String, adim As String, birlesim As Object, kendi As Object, kalan As Object
    Dim mallar() As String, sat As Collection, x As Variant, katman As Object, gun As Object, eski As Object, sarf As Object, son6 As Object
    Dim g As Double, gk As String, toplamR As Double, y As Variant, tGiris() As String, tSarf() As String, stokT As Date
    Dim aylar As Object
    On Error GoTo Hata
    tGiris = ATIL_GIRIS_TUR(): tSarf = ATIL_SARF_TUR()
    Set birlesim = CreateObject("Scripting.Dictionary"): birlesim.CompareMode = vbTextCompare
    For Each il In ADM_ILLER()
        uy = AKSORA_IL_TESIS(CStr(il), sis)
        Set s = tum(CStr(il))(0)
        For Each k In s.keys
            If Left$(CStr(k), Len(uy) + 1) = uy & "|" Then birlesim(Mid$(CStr(k), Len(uy) + 2)) = True
        Next k
    Next il
    For Each il In ADM_ILLER()
        uy = AKSORA_IL_TESIS(CStr(il), sis)
        Set s = tum(CStr(il))(0)
        Set kendi = CreateObject("Scripting.Dictionary"): kendi.CompareMode = vbTextCompare
        For Each k In s.keys
            If Left$(CStr(k), Len(uy) + 1) = uy & "|" Then kendi(Mid$(CStr(k), Len(uy) + 2)) = True
        Next k
        Set katman = CreateObject("Scripting.Dictionary"): katman.CompareMode = vbTextCompare
        Set eski = CreateObject("Scripting.Dictionary"): eski.CompareMode = vbTextCompare
        Set sarf = CreateObject("Scripting.Dictionary"): sarf.CompareMode = vbTextCompare
        Set son6 = CreateObject("Scripting.Dictionary"): son6.CompareMode = vbTextCompare
        Set aylar = CreateObject("Scripting.Dictionary"): aylar.CompareMode = vbTextCompare
        ' 1) son 24 ay girisleri
        If AKSORA_ADMIN_IPTAL Then Err.Raise vbObjectError + 2350, , "Kullanýcý durdurdu."
        If kendi.Count > 0 Then
            adim = sis & " " & uy & " (" & Y_IL_ADI(CStr(il)) & ") 1/3: giriþler 24 ay, " & kendi.Count & " malzeme"
            ATIL_ADIM 0.03 + 0.92 * adet / toplam, adim
            mallar = ATIL_DIZI(kendi)
            ' Girisler ilin stok verisi tarihine kadar: sonraki girisler stokta yok (stok artisi KARSILASTIR'da katman olur).
            stokT = CDate(Int(tum(CStr(il))(1)))
            If stokT <= 0 Or stokT > bugun Then stokT = bugun
            Set sat = ATIL_MB51(ses, uy, mallar, DateAdd("m", -24, stokT), stokT, sis & " " & uy & " 1/3", tGiris)
            For Each x In sat
                If ATIL_KAPSAM_MI(uy, CStr(x(4))) And Not ATIL_PARTI_DISI(CStr(x(5))) Then
                    g = ATIL_GIRIS_KATKI(CStr(x(3)), CDbl(x(2)))
                    If g <> 0 Then
                        If Not katman.Exists(CStr(x(0))) Then katman.Add CStr(x(0)), CreateObject("Scripting.Dictionary")
                        Set gun = katman(CStr(x(0)))
                        gk = Format$(CDate(x(1)), "yyyymmdd")
                        gun(gk) = ATIL_DS(gun, gk) + g
                    End If
                End If
            Next x
        End If
        adet = adet + 1
        ' 2) stogu 24 ay girislerinden fazla olanlar: 24-60 ay onceki son giris
        Set kalan = CreateObject("Scripting.Dictionary"): kalan.CompareMode = vbTextCompare
        For Each k In kendi.keys
            toplamR = 0
            If katman.Exists(CStr(k)) Then
                For Each y In katman(CStr(k)).items
                    toplamR = toplamR + CDbl(y)
                Next y
            End If
            If s(uy & "|" & CStr(k))(0) > toplamR + 0.0005 Then kalan(CStr(k)) = True
        Next k
        If kalan.Count > 0 Then
            If AKSORA_ADMIN_IPTAL Then Err.Raise vbObjectError + 2350, , "Kullanýcý durdurdu."
            adim = sis & " " & uy & " (" & Y_IL_ADI(CStr(il)) & ") 2/3: eski giriþler 2-5 yýl, " & kalan.Count & " malzeme"
            ATIL_ADIM 0.03 + 0.92 * adet / toplam, adim
            mallar = ATIL_DIZI(kalan)
            Set sat = ATIL_MB51(ses, uy, mallar, DateAdd("m", -60, bugun), DateAdd("m", -24, bugun) - 1, sis & " " & uy & " 2/3", tGiris)
            For Each x In sat
                If ATIL_KAPSAM_MI(uy, CStr(x(4))) And Not ATIL_PARTI_DISI(CStr(x(5))) Then
                    If ATIL_GIRIS_KATKI(CStr(x(3)), CDbl(x(2))) > 0 Then
                        If CDbl(x(1)) > ATIL_DS(eski, CStr(x(0))) Then eski(CStr(x(0))) = CDbl(x(1))
                    End If
                End If
            Next x
        End If
        adet = adet + 1
        ' 3) son 6 ay sarfi (yuklenici dahil; diger illerin talebi)
        If birlesim.Count > 0 Then
            If AKSORA_ADMIN_IPTAL Then Err.Raise vbObjectError + 2350, , "Kullanýcý durdurdu."
            adim = sis & " " & uy & " (" & Y_IL_ADI(CStr(il)) & ") 3/3: sarf 6 ay, " & birlesim.Count & " malzeme"
            ATIL_ADIM 0.03 + 0.92 * adet / toplam, adim
            mallar = ATIL_DIZI(birlesim)
            Set sat = ATIL_MB51(ses, uy, mallar, DateAdd("m", -6, bugun), bugun, sis & " " & uy & " 3/3", tSarf)
            For Each x In sat
                If ATIL_KAPSAM_MI(uy, CStr(x(4))) Then
                    sarf(CStr(x(0))) = ATIL_DS(sarf, CStr(x(0))) - CDbl(x(2))
                    If CDbl(x(2)) < 0 And CDbl(x(1)) > ATIL_DS(son6, CStr(x(0))) Then son6(CStr(x(0))) = CDbl(x(1))
                    ' Aylik net sarf (kritik stok: duzenli / tek seferlik sarf ayrimi).
                    If Not aylar.Exists(CStr(x(0))) Then aylar.Add CStr(x(0)), CreateObject("Scripting.Dictionary")
                    Set gun = aylar(CStr(x(0)))
                    gk = Format$(CDate(x(1)), "yyyymm")
                    gun(gk) = ATIL_DS(gun, gk) - CDbl(x(2))
                End If
            Next x
        End If
        adet = adet + 1
        For Each k In kendi.keys
            sonuc(uy & "|" & CStr(k)) = Array(ATIL_KATMAN_METIN(katman, CStr(k)), ATIL_DS(eski, CStr(k)), ATIL_DS(sarf, CStr(k)), _
                ATIL_DS(son6, CStr(k)), ATIL_AY_METIN(aylar, CStr(k)))
        Next k
        For Each k In sarf.keys
            If Not kendi.Exists(CStr(k)) And ATIL_DS(sarf, CStr(k)) > 0 Then _
                sonuc(uy & "|" & CStr(k)) = Array(vbNullString, 0#, ATIL_DS(sarf, CStr(k)), ATIL_DS(son6, CStr(k)), ATIL_AY_METIN(aylar, CStr(k)))
        Next k
    Next il
    Exit Function
Hata:
    ATIL_SISTEM_OKU = adim & " | " & Err.Description
End Function

'==================== SATINALMA KONTROLU (ME5A acik talepler, ME2N acik siparisler) ====================
' SAP'den (CDP ve EPP oturumu) acik satinalma talepleri ve teslim alinmamis siparisler okunur, _SATINALMA kaydina yazilir.
' Eslestirme: ayni sistemde ayni ilin serbest stogu (D / K depolari, YENI ve kullanilabilir partiler), atil stok
' (ATIL_HESAPLA) ve diger illerin serbest stogu. Stoktan karsilanabilen miktar x stok birim degeri = tasarruf (TL).

Private Function SA_GRID_ARA(ByVal o As Object) As Object
    Dim i As Long, n As Long, c As Object, tip As String, alt As String
    On Error Resume Next
    tip = o.Type: alt = o.SubType
    If tip = "GuiShell" And alt = "GridView" Then Set SA_GRID_ARA = o: Exit Function
    n = -1: n = o.Children.Count
    On Error GoTo 0
    For i = 0 To n - 1
        Set c = Nothing
        On Error Resume Next
        Set c = SA_GRID_ARA(o.Children(i))
        On Error GoTo 0
        If Not c Is Nothing Then Set SA_GRID_ARA = c: Exit Function
    Next i
End Function

Private Function SA_GRID_BUL(ByVal ses As Object) As Object
    Dim g As Object
    On Error Resume Next
    Set g = ses.FindById("wnd[0]/usr/cntlGRID1/shellcont/shell")
    If g Is Nothing Then Set g = SA_GRID_ARA(ses.FindById("wnd[0]/usr"))
    On Error GoTo 0
    Set SA_GRID_BUL = g
End Function

Private Function SA_KOLON_VAR(ByVal g As Object, ByVal id As String) As Boolean
    Dim co As Object, i As Long
    On Error Resume Next
    Set co = g.ColumnOrder
    For i = 0 To co.Count - 1
        If UCase$(CStr(co(i))) = UCase$(id) Then SA_KOLON_VAR = True: Exit Function
    Next i
End Function

' ALV duzen secimi (standart "Duzen sec" penceresi). Bulunamazsa pencere kapatilir; liste varsayilan duzenle okunur.
Private Function SA_DUZEN_YUKLE(ByVal ses As Object, ByVal g As Object, ByVal duzen As String) As Boolean
    Dim p As Object, i As Long, n As Long
    On Error Resume Next
    g.PressToolbarContextButton "&MB_VARIANT"
    g.SelectContextMenuItem "&LOAD"
    Set p = ses.FindById("wnd[1]/usr/ssubD0500_SUBSCREEN:SAPLSLVC_DIALOG:0501/cntlG51_CONTAINER/shellcont/shell")
    If p Is Nothing Then
        ses.FindById("wnd[1]").Close
        AKSORA_SAP_TESHIS_YAZ "ME2N düzen seçimi açýlamadý (" & duzen & ")"
        Exit Function
    End If
    n = p.RowCount
    For i = 0 To n - 1
        If UCase$(Trim$(CStr(p.GetCellValue(i, "VARIANT")))) = UCase$(duzen) Then
            p.SetCurrentCell i, "TEXT"
            p.ClickCurrentCell
            If Not ses.FindById("wnd[1]", False) Is Nothing Then p.DoubleClickCurrentCell
            SA_DUZEN_YUKLE = True
            Exit For
        End If
    Next i
    If Not SA_DUZEN_YUKLE Then
        ses.FindById("wnd[1]").Close
        AKSORA_SAP_TESHIS_YAZ "ME2N düzeni bulunamadý: " & duzen
    End If
    Err.Clear
End Function

' Bir sistemin ME5A + ME2N okumasi. Donus: hata metni (bos = tamam). Basarisiz okuma listeye yarim satir birakmaz.
Private Function SA_SISTEM_OKU(ByVal ses As Object, ByVal sis As String, ByVal liste As Collection, ByRef nT As Long, ByRef nS As Long) As String
    Dim c1 As Collection, c2 As Collection, x As Variant, a() As Variant, j As Long, ilk As String, son As String
    On Error GoTo Hata
    ilk = IIf(sis = "CDP", "7151", "7051"): son = IIf(sis = "CDP", "7155", "7055")
    AKSORA_BEKLE_ILERLE IIf(sis = "CDP", 0.1, 0.55), sis & "  •  açýk satýnalma talepleri (ME5A)  (SAP yanýtý bekleniyor)"
    Set c1 = SA_ME5A(ses, ilk, son, sis)
    AKSORA_BEKLE_ILERLE IIf(sis = "CDP", 0.3, 0.75), sis & "  •  teslim alýnmamýþ sipariþler (ME2N)  (SAP yanýtý bekleniyor)"
    Set c2 = SA_ME2N(ses, ilk, son, sis)
    For Each x In c1
        ReDim a(0 To 17): a(0) = "T": a(1) = sis
        For j = 0 To UBound(x): a(j + 2) = x(j): Next j
        liste.Add a: nT = nT + 1
    Next x
    For Each x In c2
        ReDim a(0 To 17): a(0) = "S": a(1) = sis
        For j = 0 To UBound(x): a(j + 2) = x(j): Next j
        liste.Add a: nS = nS + 1
    Next x
    Exit Function
Hata:
    SA_SISTEM_OKU = Err.Description & " (" & Err.Number & ")"
End Function

' Islemi acar ve beklenen alan gorunene kadar dener. Acik kalan bilgi penceresi (wnd[1]) Iptal ile kapatilir;
' guvenlik / izin penceresine dokunulmaz, hata verilir.
Private Sub SA_TX_AC(ByVal ses As Object, ByVal tx As String, ByVal alanId As String, ByVal sis As String)
    Dim deneme As Long, p As Object, baslik As String, ekran As String, sb As String
    For deneme = 1 To 3
        On Error Resume Next
        Set p = ses.FindById("wnd[1]", False)
        If Not p Is Nothing Then
            baslik = p.Text
            If InStr(1, Y_NORM(baslik), "GUVENLIK") > 0 Or InStr(1, Y_NORM(baslik), "SECURITY") > 0 Or InStr(1, Y_NORM(baslik), "IZIN") > 0 Then
                On Error GoTo 0
                Err.Raise vbObjectError + 9810, , tx & " (" & sis & "): SAP güvenlik penceresi açýk (" & baslik & "). Pencereyi SAP'de yanýtlayýp tekrar deneyin."
            End If
            AKSORA_SAP_TESHIS_YAZ tx & " " & sis & " açýk pencere kapatýlýyor: " & baslik
            ses.FindById("wnd[1]/tbar[0]/btn[12]").Press
        End If
        Err.Clear
        If deneme = 1 Then ses.StartTransaction tx Else ses.SendCommand "/n" & tx
        On Error GoTo 0
        If Not ses.FindById(alanId, False) Is Nothing Then Exit Sub
        DoEvents
    Next deneme
    On Error Resume Next
    ekran = ses.FindById("wnd[0]").Text: sb = ses.FindById("wnd[0]/sbar").Text
    On Error GoTo 0
    AKSORA_SAP_TESHIS_YAZ tx & " " & sis & " seçim ekraný açýlmadý: ekran '" & ekran & "', durum '" & sb & "'"
    Err.Raise vbObjectError + 9811, , tx & " (" & sis & ") seçim ekraný açýlamadý. Ekran: " & ekran & IIf(Len(sb) > 0, "  •  SAP: " & sb, "")
End Sub

Private Sub SA_YAZ(ByVal ses As Object, ByVal id As String, ByVal deger As String, ByVal adim As String)
    Dim o As Object
    Set o = ses.FindById(id, False)
    If o Is Nothing Then
        AKSORA_SAP_TESHIS_YAZ adim & ": alan bulunamadý " & id
        Err.Raise vbObjectError + 9812, , adim & ": SAP ekranýnda alan bulunamadý (" & Mid$(id, InStrRev(id, "/") + 1) & ")."
    End If
    o.Text = deger
End Sub

Private Sub SA_BAS(ByVal ses As Object, ByVal id As String, ByVal adim As String)
    Dim o As Object
    Set o = ses.FindById(id, False)
    If o Is Nothing Then
        AKSORA_SAP_TESHIS_YAZ adim & ": düðme bulunamadý " & id
        Err.Raise vbObjectError + 9813, , adim & ": SAP ekranýnda Çalýþtýr düðmesi bulunamadý."
    End If
    o.Press
End Sub

Private Sub SA_ALAN_BOSALT(ByVal ses As Object, ByVal ad As String)
    On Error Resume Next
    ses.FindById("wnd[0]/usr/" & ad).Text = vbNullString
    On Error GoTo 0
End Sub

' tanim: Array(Array(anahtar, "ID1|ID2", "BASLIK1|BASLIK2"), ...). Donus: anahtar -> kolon id (bulunamazsa "").
Private Function SA_KOLONLAR(ByVal g As Object, ByVal tanim As Variant, ByVal etiket As String) As Object
    Dim d As Object, idler As Object, t As Variant, id As Variant, b As Variant, liste As String, nt As String, bul As String, nb As String
    Dim co As Object, i As Long, v As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set idler = CreateObject("Scripting.Dictionary"): idler.CompareMode = vbTextCompare
    Set co = g.ColumnOrder
    For i = 0 To co.Count - 1
        id = CStr(co(i))
        nt = vbNullString
        On Error Resume Next
        nt = Y_NORM(CStr(g.GetDisplayedColumnTitle(CStr(id))))
        On Error GoTo 0
        idler(CStr(id)) = nt
        liste = liste & " | " & CStr(id) & "=" & nt
    Next i
    AKSORA_SAP_TESHIS_YAZ etiket & " sütunlar:" & liste
    For Each t In tanim
        bul = vbNullString
        For Each id In Split(CStr(t(1)), "|")
            If Len(CStr(id)) > 0 Then If idler.Exists(CStr(id)) Then bul = CStr(id)
            If Len(bul) > 0 Then Exit For
        Next id
        ' Gruplanmis sutun (ME2N: satinalma belgesi) ColumnOrder'da olmayabilir: id hucre okumasiyla denenir.
        If Len(bul) = 0 Then
            For Each id In Split(CStr(t(1)), "|")
                If Len(CStr(id)) > 0 Then
                    v = Empty
                    On Error Resume Next
                    v = g.GetCellValue(0, CStr(id))
                    If Err.Number = 0 And Len(Trim$(CStr(v))) > 0 Then bul = CStr(id)
                    Err.Clear
                    On Error GoTo 0
                End If
                If Len(bul) > 0 Then Exit For
            Next id
        End If
        If Len(bul) = 0 Then
            For Each b In Split(CStr(t(2)), "|")
                nb = Y_NORM(CStr(b))
                For Each id In idler.keys
                    If idler(id) = nb Then bul = CStr(id): Exit For
                Next id
                If Len(bul) > 0 Then Exit For
            Next b
        End If
        d(CStr(t(0))) = bul
    Next t
    Set SA_KOLONLAR = d
End Function

' Izgara satirlari: Array(anahtar sirasinda degerler). Bos malzeme / belge satirlari (ara toplam) alinmaz.
Private Function SA_IZGARA_OKU(ByVal g As Object, ByVal kol As Object, ByVal anahtarlar As Variant) As Collection
    Dim c As Collection, n As Long, i As Long, j As Long, a() As Variant, v As Variant, idMal As String, idBel As String, gor As Object
    Set c = New Collection
    Set gor = CreateObject("Scripting.Dictionary")
    n = 0
    On Error Resume Next
    n = g.RowCount
    On Error GoTo 0
    idMal = kol("MAL"): idBel = kol("BELGE")
    For i = 0 To n - 1
        If (i Mod 25) = 0 Then
            On Error Resume Next
            g.FirstVisibleRow = i
            On Error GoTo 0
            AKSORA_BEKLE_NABIZ
        End If
        ReDim a(0 To UBound(anahtarlar))
        For j = 0 To UBound(anahtarlar)
            v = vbNullString
            If Len(kol(anahtarlar(j))) > 0 Then
                On Error Resume Next
                v = g.GetCellValue(i, kol(anahtarlar(j)))
                On Error GoTo 0
            End If
            a(j) = Trim$(CStr(v))
        Next j
        ' Ara toplam / grup satirlari ayni belge-kalemi tekrar dondurebilir: belge + kalem bir kez alinir.
        If Len(a(SA_IDX(anahtarlar, "MAL"))) > 0 Then
            If Len(a(0)) = 0 Then
                c.Add a
            ElseIf Not gor.Exists(a(0) & "|" & a(1)) Then
                gor(a(0) & "|" & a(1)) = True: c.Add a
            End If
        End If
    Next i
    Set SA_IZGARA_OKU = c
End Function

Private Function SA_IDX(ByVal anahtarlar As Variant, ByVal ad As String) As Long
    Dim j As Long
    For j = 0 To UBound(anahtarlar)
        If anahtarlar(j) = ad Then SA_IDX = j: Exit Function
    Next j
    SA_IDX = -1
End Function

Private Function SA_ME5A_ANAHTAR() As Variant
    SA_ME5A_ANAHTAR = Array("BELGE", "KALEM", "UY", "DEPO", "SAG", "MG", "MAL", "TANIM", "KISI", "MIKTAR", "BIRIM", "TARIH", _
        "SIPARIS", "DEGER", "PB", "BASLIK")
End Function

Private Function SA_ME2N_ANAHTAR() As Variant
    SA_ME2N_ANAHTAR = Array("BELGE", "KALEM", "UY", "DEPO", "SAG", "MG", "MAL", "TANIM", "KISI", "MIKTAR", "BIRIM", "TARIH", _
        "FIYAT", "ACIK", "ACIKDEGER", "PB")
End Function

' Acik satinalma talepleri (ME5A). Liste yoksa (talep yok) bos koleksiyon.
Private Function SA_ME5A(ByVal ses As Object, ByVal ilk As String, ByVal son As String, ByVal sis As String) As Collection
    Dim g As Object, kol As Object, sb As String, tip As String
    Set SA_ME5A = New Collection
    SA_TX_AC ses, "ME5A", "wnd[0]/usr/ctxtS_WERKS-LOW", sis
    SA_ALAN_BOSALT ses, "ctxtBA_BANFN-LOW": SA_ALAN_BOSALT ses, "ctxtBA_BANFN-HIGH"
    SA_ALAN_BOSALT ses, "ctxtS_MATNR-LOW": SA_ALAN_BOSALT ses, "ctxtS_MATNR-HIGH"
    SA_YAZ ses, "wnd[0]/usr/ctxtS_WERKS-LOW", ilk, "ME5A " & sis
    SA_YAZ ses, "wnd[0]/usr/ctxtS_WERKS-HIGH", son, "ME5A " & sis
    SA_BAS ses, "wnd[0]/tbar[1]/btn[8]", "ME5A " & sis
    Set g = SA_GRID_BUL(ses)
    If g Is Nothing Then
        On Error Resume Next
        sb = ses.FindById("wnd[0]/sbar").Text: tip = ses.FindById("wnd[0]/sbar").MessageType
        On Error GoTo 0
        AKSORA_SAP_TESHIS_YAZ "ME5A " & sis & " liste yok: " & sb
        If tip = "E" Or tip = "A" Then Err.Raise vbObjectError + 9801, , "ME5A (" & sis & "): " & sb
        Exit Function
    End If
    Set kol = SA_KOLONLAR(g, Array( _
        Array("BELGE", "BANFN", "SA TALEBI|SATINALMA TALEBI"), Array("KALEM", "BNFPO", "KALEM|TALEP KALEMI"), _
        Array("UY", "WERKS", "UY|URETIM YERI"), Array("DEPO", "LGORT", "DPYR|DEPO YERI"), _
        Array("SAG", "EKGRP", "SAG|SATINALMA GRUBU"), Array("MG", "MATKL", "MAL GRUBU"), _
        Array("MAL", "MATNR", "MALZEME"), Array("TANIM", "TXZ01", "KISA METIN"), _
        Array("KISI", "AFNAM", "TALEP EDEN"), Array("MIKTAR", "MENGE", "MIKTAR|TALEP MIKTARI"), _
        Array("BIRIM", "MEINS", "BIRIM|OB|OLCU BIRIMI"), Array("TARIH", "BADAT", "TALEP TARIHI"), _
        Array("SIPARIS", "EBELN", "SATINALMA SIPARISI"), Array("DEGER", "GSWRT|RLWRT", "TOPLAM DEGER"), _
        Array("PB", "WAERS", "PB|PARA BIRIMI"), Array("BASLIK", "", "BASLIK METNI")), "ME5A " & sis)
    If Len(kol("MAL")) = 0 Or Len(kol("UY")) = 0 Or Len(kol("MIKTAR")) = 0 Then _
        Err.Raise vbObjectError + 9802, , "ME5A (" & sis & ") listesinde malzeme / üretim yeri / miktar sütunu bulunamadý (teþhis dosyasýnda sütunlar yazýlý)."
    Set SA_ME5A = SA_IZGARA_OKU(g, kol, SA_ME5A_ANAHTAR())
End Function

' Teslim alinmamis siparisler (ME2N, secim parametresi WE101, liste kapsami ALV).
Private Function SA_ME2N(ByVal ses As Object, ByVal ilk As String, ByVal son As String, ByVal sis As String) As Collection
    Dim g As Object, kol As Object, sb As String, tip As String
    Set SA_ME2N = New Collection
    SA_TX_AC ses, "ME2N", "wnd[0]/usr/ctxtLISTU", sis
    SA_ALAN_BOSALT ses, "ctxtEN_EBELN-LOW": SA_ALAN_BOSALT ses, "ctxtEN_EBELN-HIGH"
    SA_ALAN_BOSALT ses, "ctxtEM_MATNR-LOW": SA_ALAN_BOSALT ses, "ctxtEM_MATNR-HIGH"
    SA_ALAN_BOSALT ses, "ctxtS_MATNR-LOW": SA_ALAN_BOSALT ses, "ctxtS_MATNR-HIGH"
    SA_YAZ ses, "wnd[0]/usr/ctxtLISTU", "ALV", "ME2N " & sis
    ses.FindById("wnd[0]").SendVKey 0
    SA_YAZ ses, "wnd[0]/usr/ctxtSELPA-LOW", "WE101", "ME2N " & sis
    SA_YAZ ses, "wnd[0]/usr/ctxtS_WERKS-LOW", ilk, "ME2N " & sis
    SA_YAZ ses, "wnd[0]/usr/ctxtS_WERKS-HIGH", son, "ME2N " & sis
    SA_BAS ses, "wnd[0]/tbar[1]/btn[8]", "ME2N " & sis
    Set g = SA_GRID_BUL(ses)
    If g Is Nothing Then
        On Error Resume Next
        sb = ses.FindById("wnd[0]/sbar").Text: tip = ses.FindById("wnd[0]/sbar").MessageType
        On Error GoTo 0
        AKSORA_SAP_TESHIS_YAZ "ME2N " & sis & " liste yok: " & sb
        If tip = "E" Or tip = "A" Then Err.Raise vbObjectError + 9803, , "ME2N (" & sis & "): " & sb
        Exit Function
    End If
    ' Satinalma belgesi sutunu gorunmuyorsa (/AKSORA duzeni bu kullanicida varsayilan degil) duzen AKSORA tarafindan secilir.
    If Not SA_KOLON_VAR(g, "EBELN") Then
        If SA_DUZEN_YUKLE(ses, g, "/AKSORA") Then Set g = SA_GRID_BUL(ses)
    End If
    Set kol = SA_KOLONLAR(g, Array( _
        Array("BELGE", "EBELN", "SATINALMA BELGESI|SAS"), Array("KALEM", "EBELP", "KALEM"), _
        Array("UY", "WERKS", "UY|URETIM YERI"), Array("DEPO", "LGORT", "DPYR|DEPO YERI"), _
        Array("SAG", "EKGRP", "SAS G.|SAS GRUBU|SATINALMA GRUBU"), Array("MG", "MATKL", "MAL GRUBU"), _
        Array("MAL", "EMATN|MATNR", "MALZEME"), Array("TANIM", "TXZ01", "KISA METIN"), _
        Array("KISI", "SUPERFIELD|NAME1|LIFNR", "SATICI/TESLIMAT YAPAN URT.YERI|SATICI"), _
        Array("MIKTAR", "MENGE", "SAS MIKTARI|SA SIPARISI MIKTARI"), Array("BIRIM", "MEINS", "SOB|SAS OLCU BIRIMI"), _
        Array("TARIH", "BEDAT", "BELGE TARIHI"), Array("FIYAT", "NETPR", "NET FIYAT"), _
        Array("ACIK", "MGLIEF", "TSL.YAPILACAK|TESLIMATI YAPILACAK|TESLIM EDILECEK"), _
        Array("ACIKDEGER", "WTLIEF", "TESLIMATI YAPILACAK (DEGER)|TESLIM EDILECEK DEGER"), _
        Array("PB", "WAERS", "PB|PARA BIRIMI")), "ME2N " & sis)
    If Len(kol("MAL")) = 0 Or Len(kol("UY")) = 0 Or Len(kol("ACIK")) = 0 Then _
        Err.Raise vbObjectError + 9804, , "ME2N (" & sis & ") listesinde malzeme / üretim yeri / teslim edilecek miktar sütunu bulunamadý (teþhis dosyasýnda sütunlar yazýlý)."
    Set SA_ME2N = SA_IZGARA_OKU(g, kol, SA_ME2N_ANAHTAR())
End Function

Private Sub SA_KAYDET(ByVal liste As Collection, ByVal bilgi As String)
    Dim ws As Worksheet, v() As Variant, i As Long, j As Long, x As Variant, onceki As Object
    Set onceki = ActiveSheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SA_KAYIT)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SA_KAYIT
        ws.Visible = xlSheetVeryHidden
        On Error Resume Next
        onceki.Activate
        On Error GoTo 0
    End If
    ws.Visible = xlSheetVeryHidden
    ws.Cells.Clear
    ws.Range("A1").Value2 = "SA1": ws.Range("B1").Value = Now: ws.Range("C1").Value2 = bilgi
    If liste.Count > 0 Then
        ReDim v(1 To liste.Count, 1 To 18)
        For i = 1 To liste.Count
            x = liste(i)
            For j = 0 To UBound(x): v(i, j + 1) = x(j): Next j
        Next i
        ws.Range("A3").Resize(liste.Count, 18).NumberFormat = "@"
        ws.Range("A3").Resize(liste.Count, 18).Value2 = v
    End If
End Sub

' Kayit: Array(0 tur T/S, 1 sistem, 2..17 anahtarlar). Talep: BELGE..BASLIK; siparis: BELGE..PB.
Private Function SA_KAYIT_OKU(ByRef tarih As Date, ByRef bilgi As String) As Collection
    Dim ws As Worksheet, son As Long, v As Variant, i As Long, j As Long, a() As Variant
    Set SA_KAYIT_OKU = New Collection
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SA_KAYIT)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    If CStr(ws.Range("A1").Value2) <> "SA1" Then Exit Function
    If IsDate(ws.Range("B1").Value) Then tarih = CDate(ws.Range("B1").Value)
    bilgi = CStr(ws.Range("C1").Value2)
    son = AKSORA_SON_SATIR(ws, 1)
    If son < 3 Then Exit Function
    v = ws.Range("A3:R" & son).Value2
    For i = 1 To UBound(v, 1)
        ReDim a(0 To 17)
        For j = 1 To 18: a(j - 1) = CStr(v(i, j)): Next j
        SA_KAYIT_OKU.Add a
    Next i
End Function

' SAP'den oku: Yonetim Merkezi "SATINALMA KONTROLU" dugmesi.
Public Sub AKSORA_SATINALMA_OKU()
    Dim ac As Boolean, ev As Boolean, sis As Variant, ses As Object, hata As String, liste As Collection, c As Collection
    Dim x As Variant, a() As Variant, j As Long, bilgi As String, no As Long, metin As String, nT As Long, nS As Long
    Dim oncekiFiltre As LongPtr, bosFiltre As LongPtr, bas As Worksheet, eskiT As Date, eskiB As String, eksikSis As String, okunan As Long
    If MsgBox("Açýk satýnalma talepleri (ME5A) ve teslim alýnmamýþ sipariþler (ME2N) CDP ve EPP sistemlerinden okunacak." & vbCrLf & _
        "SAP'de iki sisteme de giriþ yapýlmýþ olmalýdýr. Devam edilsin mi?", vbQuestion + vbYesNo, "AKSORA | SATINALMA KONTROLÜ") <> vbYes Then Exit Sub
    Set bas = ActiveSheet
    ac = AKSORA_BEKLE_AC("AKSORA  |  SATINALMA KONTROLÜ", "SAP baðlantýsý hazýrlanýyor...", 0.03)
    ev = application.EnableEvents
    ' Uzun SAP raporunda "Excel bir OLE eylemini ... bekliyor" penceresi cikmasin (mesaj filtresi gecici kapatilir).
    ADMW_CoRegisterMessageFilter 0, oncekiFiltre
    On Error GoTo Hata
    Set liste = New Collection
    For Each sis In Array("CDP", "EPP")
        hata = vbNullString
        Set ses = SAP_SESSION_AL(CStr(sis), hata)
        If ses Is Nothing Then
            bilgi = bilgi & CStr(sis) & ": SAP oturumu yok (" & hata & ")  "
        Else
            ' Sistem okunamazsa bir kez daha denenir; yine olmazsa o sistemin onceki kaydi korunur (diger sistem kaybolmaz).
            hata = SA_SISTEM_OKU(ses, CStr(sis), liste, nT, nS)
            If Len(hata) > 0 And Not AKSORA_ADMIN_IPTAL Then
                AKSORA_SAP_TESHIS_YAZ "Satýnalma " & sis & " yeniden deneniyor: " & hata
                AKSORA_BEKLE_ILERLE IIf(sis = "CDP", 0.12, 0.57), CStr(sis) & "  •  yeniden deneniyor (geçici hata)..."
                On Error Resume Next
                ses.SendCommand "/n"
                On Error GoTo Hata
                DoEvents
                hata = SA_SISTEM_OKU(ses, CStr(sis), liste, nT, nS)
            End If
            On Error Resume Next
            ses.SendCommand "/n"
            On Error GoTo Hata
            If Len(hata) = 0 Then
                okunan = okunan + 1
                bilgi = bilgi & CStr(sis) & ": tamam  "
            Else
                AKSORA_SAP_TESHIS_YAZ "Satýnalma " & sis & " okunamadý: " & hata
                For Each x In SA_KAYIT_OKU(eskiT, eskiB)
                    If x(1) = CStr(sis) Then
                        liste.Add x
                        If x(0) = "T" Then nT = nT + 1 Else nS = nS + 1
                    End If
                Next x
                bilgi = bilgi & CStr(sis) & ": OKUNAMADI, önceki kayýt korundu (" & Left$(hata, 120) & ")  "
                eksikSis = eksikSis & IIf(Len(eksikSis) > 0, ", ", "") & CStr(sis)
            End If
        End If
    Next sis
    ' Hicbir sistem okunamadiysa onceki kayit oldugu gibi kalir.
    If okunan = 0 Then Err.Raise vbObjectError + 9815, , Trim$(bilgi)
    ADMW_CoRegisterMessageFilter oncekiFiltre, bosFiltre
    application.EnableEvents = False
    SA_KAYDET liste, Trim$(bilgi)
    ' Gizli kayit sayfasi olusunca Excel baska sayfaya gecmesin: baslanan sayfaya donulur.
    On Error Resume Next
    ThisWorkbook.Activate
    bas.Activate
    On Error GoTo Hata
    application.EnableEvents = ev
    If ac Then AKSORA_BEKLE_KAPAT
    MsgBox IIf(Len(eksikSis) > 0, "Satýnalma kontrolü kýsmen okundu (" & eksikSis & " okunamadý, önceki kaydý korundu).", "Satýnalma kontrolü okundu.") & _
        vbCrLf & vbCrLf & "Açýk talep kalemi: " & nT & vbCrLf & "Açýk sipariþ kalemi: " & nS & _
        vbCrLf & vbCrLf & Trim$(bilgi), IIf(Len(eksikSis) > 0, vbExclamation, vbInformation), "AKSORA | SATINALMA KONTROLÜ"
    If bas.Name = ADM Then
        AKSORA_ADMIN_YENILE
    ElseIf bas.Name = SA_SAYFA Then
        AKSORA_SATINALMA_AC
    End If
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    ADMW_CoRegisterMessageFilter oncekiFiltre, bosFiltre
    If Not bas Is Nothing Then ThisWorkbook.Activate: bas.Activate
    application.EnableEvents = ev
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    MsgBox "Satýnalma kontrolü okunamadý." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA | SATINALMA KONTROLÜ"
End Sub

' Eslestirme. Tasarruf YALNIZ atil stoktan: once ayni ilin atil stogu, sonra ayni sistemde diger illerin atil stogu.
' Atil miktar kalemler arasinda paylastirilir (ayni stok iki kez sayilmaz); talepler once, sonra siparisler.
' Serbest (calisan) stok bilgi olarak gosterilir, tasarrufa katilmaz. Stok birimi farkliysa eslestirilmez.
' Donus: Collection of Array(0 tur, 1 sistem, 2 il, 3 UY, 4 depo, 5 belge, 6 kalem, 7 tarih, 8 kisi/satici,
'   9 malzeme, 10 tanim, 11 acik miktar, 12 birim, 13 ayni il serbest, 14 ayni il atil, 15 transfer ili, 16 transfer miktari,
'   17 karsilanabilir (atil), 18 stok birim degeri TL, 19 tasarruf TL, 20 belge tutari, 21 PB, 22 durum, 23 not,
'   24 atil kaynaklari "ILKODU=miktar;..." (kendi ili dahil; transfer listesi ve kritik stok havuzu icin))
Private Function SA_ANALIZ(ByVal kayit As Collection) As Collection
    Dim kay As String
    Dim tum As Object, havuz As Object, c As Collection, bilgi As Variant, x As Variant, r As Variant, il As String, uy As String
    Dim mal As String, sis As String, miktar As Double, serbest As Double, yuk As Double, atil As Double, bd As Double
    Dim oIl As String, oMik As Double, il2 As Variant, uy2 As String, s2 As Double, karsi As Double, durum As String, nt As String
    Dim i As Long, tutar As Double, pb As String, y2 As Double, b2 As Double, tur As Variant, ihtiyac As Double, a1 As Double
    Dim al As Double, enIl As String, enMik As Double, sIl As String, sMik As Double, stBirim As String, k As String, o As Variant
    Dim ilkHavuz As Object, atilTop As Double
    Set SA_ANALIZ = New Collection
    If kayit.Count = 0 Then Exit Function
    Set tum = ATIL_TUM_STOK()
    Set havuz = CreateObject("Scripting.Dictionary"): havuz.CompareMode = vbTextCompare
    Set ilkHavuz = CreateObject("Scripting.Dictionary"): ilkHavuz.CompareMode = vbTextCompare
    If Not ADM_WS(ATIL_KAYIT) Is Nothing Then
        On Error Resume Next
        Set c = ATIL_HESAPLA(bilgi)
        On Error GoTo 0
        If Not c Is Nothing Then
            For Each x In c
                If IsArray(x(20)) Then If x(20)(3) > 0.0005 Then havuz(CStr(x(2)) & "|" & CStr(x(3))) = CDbl(x(20)(3)): ilkHavuz(CStr(x(2)) & "|" & CStr(x(3))) = CDbl(x(20)(3))
            Next x
        End If
    End If
    For Each tur In Array("T", "S")
        For Each r In kayit
            If r(0) = tur Then
                i = i + 1: If (i And 63) = 0 Then AKSORA_BEKLE_NABIZ
                uy = Trim$(CStr(r(4))): mal = Y_KOD_NORM(r(8)): sis = ATIL_SISTEM(uy): il = ATIL_TESIS_IL(uy)
                If tur = "T" Then
                    miktar = SA_SAYI(r(11)): tutar = SA_SAYI(r(15)): pb = CStr(r(16))
                Else
                    miktar = SA_SAYI(r(15)): tutar = SA_SAYI(r(16)): pb = CStr(r(17))
                End If
                serbest = 0: yuk = 0: atil = 0: atilTop = 0: bd = 0: oIl = vbNullString: oMik = 0: nt = vbNullString: karsi = 0: kay = vbNullString
                sIl = vbNullString: sMik = 0: stBirim = vbNullString
                If Len(il) > 0 And Len(mal) > 0 And miktar > 0.0005 Then
                    k = uy & "|" & mal
                    SA_STOK tum, il, k, serbest, yuk, bd
                    stBirim = SA_STOK_BIRIM(tum, sis, mal)
                    If bd <= 0 Then bd = SA_BIRIM_DEGER(tum, sis, mal)
                    If havuz.Exists(k) Then atil = havuz(k): atilTop = ilkHavuz(k)
                    If Len(stBirim) > 0 And UCase$(stBirim) <> UCase$(Trim$(CStr(r(12)))) Then
                        nt = "Birim farklý (talep " & CStr(r(12)) & ", stok " & stBirim & "): eþleþtirilmedi"
                    Else
                        ' 1) ayni ilin atil stogu
                        ihtiyac = miktar
                        a1 = application.Min(ihtiyac, atil)
                        If a1 > 0.0005 Then havuz(k) = atil - a1: ihtiyac = ihtiyac - a1: karsi = a1: kay = il & "=" & AN_SAYI_YAZ(a1)
                        ' 2) diger illerin atil stogu (en buyukten), paylastirmali
                        Do While ihtiyac > 0.0005
                            enIl = vbNullString: enMik = 0
                            For Each il2 In ADM_ILLER()
                                If CStr(il2) <> il Then
                                    uy2 = AKSORA_IL_TESIS(CStr(il2), sis) & "|" & mal
                                    If havuz.Exists(uy2) Then If havuz(uy2) > enMik Then enMik = havuz(uy2): enIl = CStr(il2)
                                End If
                            Next il2
                            If enMik <= 0.0005 Then Exit Do
                            al = application.Min(ihtiyac, enMik)
                            havuz(AKSORA_IL_TESIS(enIl, sis) & "|" & mal) = enMik - al
                            ihtiyac = ihtiyac - al: oMik = oMik + al: karsi = karsi + al
                            kay = kay & IIf(Len(kay) > 0, ";", "") & enIl & "=" & AN_SAYI_YAZ(al)
                            If Len(oIl) = 0 Then oIl = Y_IL_ADI(enIl) Else If InStr(oIl, Y_IL_ADI(enIl)) = 0 Then oIl = oIl & ", " & Y_IL_ADI(enIl)
                        Loop
                    End If
                    ' Serbest stok (bilgi): diger illerde en buyuk
                    For Each il2 In ADM_ILLER()
                        If CStr(il2) <> il Then
                            s2 = 0: y2 = 0: b2 = -1
                            SA_STOK tum, CStr(il2), AKSORA_IL_TESIS(CStr(il2), sis) & "|" & mal, s2, y2, b2
                            If s2 > sMik Then sMik = s2: sIl = Y_IL_ADI(CStr(il2))
                        End If
                    Next il2
                End If
                If karsi > 0.0005 And oMik <= 0.0005 Then
                    durum = "ATIL STOKTAN KARÞILANIR"
                ElseIf karsi > 0.0005 Then
                    durum = IIf(karsi - oMik > 0.0005, "ATIL STOKTAN KARÞILANIR", "BAÞKA ÝLÝN ATIL STOÐUNDAN")
                ElseIf serbest > 0.0005 Then
                    durum = "AYNI ÝLDE SERBEST STOK VAR"
                ElseIf sMik > 0.0005 Then
                    durum = "BAÞKA ÝLDE SERBEST STOK VAR"
                Else
                    durum = "STOK YOK"
                End If
                ' Eksi teslim edilecek: siparisten fazla teslim alinmis (tasarrufa girmez, ayrica incelenmeli).
                If miktar < -0.0005 Then durum = "FAZLA TESLÝM ALINMIÞ"
                If durum = "BAÞKA ÝLDE SERBEST STOK VAR" Then nt = Trim$(nt & IIf(Len(nt) > 0, "  •  ", "") & sIl & ": " & ADM_ADET(sMik) & " " & CStr(r(12)) & " serbest")
                If yuk > 0.0005 Then nt = Trim$(nt & IIf(Len(nt) > 0, "  •  ", "") & "Yüklenici depolarýnda " & ADM_ADET(yuk) & " " & CStr(r(12)))
                If tur = "T" Then
                    If Len(CStr(r(14))) > 0 Then nt = Trim$(nt & IIf(Len(nt) > 0, "  •  ", "") & "Kýsmen sipariþte: " & CStr(r(14)))
                End If
                SA_ANALIZ.Add Array(r(0), r(1), Y_IL_ADI(il), uy, r(5), r(2), r(3), r(13), r(10), r(8), r(9), miktar, r(12), _
                    serbest, atilTop, oIl, oMik, karsi, Round(bd, 4), Round(karsi * bd, 2), tutar, pb, durum, nt, kay)
            End If
        Next r
    Next tur
End Function

' Stok birimi (ayni sistemde malzemenin stokta oldugu ilk ilden).
Private Function SA_STOK_BIRIM(ByVal tum As Object, ByVal sis As String, ByVal mal As String) As String
    Dim il As Variant, k As String, s As Object
    For Each il In ADM_ILLER()
        If tum.Exists(CStr(il)) Then
            Set s = tum(CStr(il))(0)
            k = AKSORA_IL_TESIS(CStr(il), sis) & "|" & mal
            If s.Exists(k) Then SA_STOK_BIRIM = Trim$(CStr(s(k)(2))): If Len(SA_STOK_BIRIM) > 0 Then Exit Function
        End If
    Next il
End Function

Private Function SA_SAYI(ByVal v As Variant) As Double
    Dim x As Double
    If ATIL_SAYI_DENE(v, x) Then SA_SAYI = x
End Function

' Ilin serbest (D / K depolari) ve yuklenici stogu; birim deger (TL).
Private Sub SA_STOK(ByVal tum As Object, ByVal il As String, ByVal k As String, ByRef serbest As Double, ByRef yuk As Double, _
    ByRef bd As Double)
    Dim s As Object, o As Variant, dp As Object, d As Variant
    If Not tum.Exists(il) Then Exit Sub
    Set s = tum(il)(0)
    If Not s.Exists(k) Then Exit Sub
    o = s(k)
    Set dp = o(5)
    For Each d In dp.keys
        If ATIL_DEPO_MI(CStr(d)) Then serbest = serbest + dp(d)(1) Else yuk = yuk + dp(d)(1)
    Next d
    If bd >= 0 And o(0) > 0 Then bd = o(1) / o(0)
End Sub

Private Function SA_BIRIM_DEGER(ByVal tum As Object, ByVal sis As String, ByVal mal As String) As Double
    Dim il As Variant, k As String, o As Variant, s As Object
    For Each il In ADM_ILLER()
        If tum.Exists(CStr(il)) Then
            Set s = tum(CStr(il))(0)
            k = AKSORA_IL_TESIS(CStr(il), sis) & "|" & mal
            If s.Exists(k) Then
                o = s(k)
                If o(0) > 0 Then SA_BIRIM_DEGER = o(1) / o(0): Exit Function
            End If
        End If
    Next il
End Function

' Excel: Ozet, Acik Talepler, Acik Siparisler (masaustu).
Public Sub AKSORA_SATINALMA_EXCEL()
    Dim kayit As Collection, an As Collection, tarih As Date, bilgi As String, wb As Workbook, ac As Boolean, ev As Boolean
    Dim veri() As Variant, x As Variant, n As Long, i As Long, j As Long, anah() As String, idx() As Long, tur As Variant
    Dim ust As String, sonuc As String, no As Long, metin As String, c As Worksheet, hz As Range, oz As Object, il As Variant
    Dim kk As Variant, sira As Long, t As Variant, toplam(1 To 10) As Double, j2 As Variant
    AKSORA_DONUS_KAYDET True
    Set kayit = SA_KAYIT_OKU(tarih, bilgi)
    If kayit.Count = 0 Then MsgBox "Önce SATINALMA KONTROLÜ ile SAP'den okuyun.", vbInformation, "AKSORA": Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  SATIN ALMA", "Stok ve atýl stokla eþleþtiriliyor...", 0.1)
    ev = application.EnableEvents
    AN_ONB_AC
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    Set an = SA_ANALIZ(kayit)
    AKSORA_BEKLE_ILERLE 0.7, "Excel hazýrlanýyor..."
    AKSORA_DONUS_KAYDET
    Set wb = Workbooks.Add(xlWBATWorksheet)
    wb.Worksheets.Add After:=wb.Worksheets(1)
    wb.Worksheets.Add After:=wb.Worksheets(2)
    ' Ozet (il bazinda)
    Set oz = CreateObject("Scripting.Dictionary")
    For Each il In ADM_ILLER(): oz(Y_IL_ADI(CStr(il))) = Array(0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#, 0#): Next il
    For Each x In an
        If oz.Exists(x(2)) Then
            t = oz(x(2))
            sira = IIf(x(0) = "T", 0, 3)
            t(sira) = t(sira) + 1
            If x(17) > 0.0005 Then t(sira + 1) = t(sira + 1) + 1: t(sira + 2) = t(sira + 2) + x(19): t(6) = t(6) + x(19)
            Select Case x(22)
                Case "FAZLA TESLÝM ALINMIÞ": t(7) = t(7) + 1
                Case "AYNI ÝLDE SERBEST STOK VAR", "BAÞKA ÝLDE SERBEST STOK VAR": t(8) = t(8) + 1
                Case "STOK YOK": t(9) = t(9) + 1
            End Select
            oz(x(2)) = t
        End If
    Next x
    ReDim veri(1 To oz.Count + 1, 1 To 11)
    i = 0
    For Each kk In oz.keys
        i = i + 1: t = oz(kk): veri(i, 1) = kk
        For j = 0 To 9: veri(i, j + 2) = SA_YUVARLA(t(j), 2): toplam(j + 1) = toplam(j + 1) + t(j): Next j
    Next kk
    i = i + 1: veri(i, 1) = "TOPLAM"
    For j = 1 To 10: veri(i, j + 1) = SA_YUVARLA(toplam(j), 2): Next j
    ust = "SATINALMA KONTROLÜ  |  Özet  |  SAP okumasý: " & Format$(tarih, "dd.mm.yyyy hh:nn")
    Set c = wb.Worksheets(1)
    AKSORA_EXCEL_SAYFA c, Array("Ýl", "Açýk talep kalemi", "Atýl stoktan karþýlanabilir (talep)", "Talep tasarrufu (TL)", _
        "Açýk sipariþ kalemi", "Atýl stoktan karþýlanabilir (sipariþ)", "Sipariþ tasarrufu (TL)", "Toplam tasarruf (TL)", _
        "Fazla teslim", "Serbest stok var (bilgi)", "Stok yok"), veri, i, Array(14, 14, 18, 20, 14, 18, 20, 20, 12, 16, 12), _
        Array("@", "#,##0", "#,##0", "#,##0.00 ""TL""", "#,##0", "#,##0", "#,##0.00 ""TL""", "#,##0.00 ""TL""", "#,##0", "#,##0", "#,##0"), _
        Array(xlLeft, xlCenter, xlCenter, xlRight, xlCenter, xlCenter, xlRight, xlRight, xlCenter, xlCenter, xlCenter), ust, _
        "Tasarruf: atýl stoktan karþýlanabilen miktar (paylaþtýrmalý) x AKSORA stok birim deðeri", "Özet", Empty
    c.Range(c.Cells(i + 4, 1), c.Cells(i + 4, 11)).Font.bold = True
    c.Range(c.Cells(i + 4, 1), c.Cells(i + 4, 11)).Interior.Color = RGB(226, 236, 247)
    c.Range(c.Cells(5, 8), c.Cells(i + 4, 8)).Font.bold = True
    c.Range(c.Cells(5, 8), c.Cells(i + 4, 8)).Font.Color = RGB(21, 128, 61)
    ' Talepler ve siparisler (durum onceligi, sonra tasarruf buyukten kucuge)
    For Each tur In Array("T", "S")
        n = 0
        For Each x In an
            If x(0) = tur Then n = n + 1
        Next x
        Set c = wb.Worksheets(IIf(tur = "T", 2, 3))
        If n = 0 Then
            ReDim veri(1 To 1, 1 To 20): veri(1, 10) = "Açýk kayýt yok.": n = 1
        Else
            ReDim anah(1 To n): ReDim idx(1 To n): ReDim veri(1 To n, 1 To 20)
            i = 0
            For Each x In an
                If x(0) = tur Then
                    i = i + 1: idx(i) = i
                    veri(i, 1) = x(22): veri(i, 2) = x(2): veri(i, 3) = x(3): veri(i, 4) = x(4): veri(i, 5) = x(5): veri(i, 6) = x(6)
                    veri(i, 7) = x(7): veri(i, 8) = x(8): veri(i, 9) = x(9): veri(i, 10) = x(10): veri(i, 11) = x(11): veri(i, 12) = x(12)
                    veri(i, 11) = SA_YUVARLA(x(11), 3)
                    veri(i, 13) = SA_YUVARLA(x(13), 3): veri(i, 14) = SA_YUVARLA(x(14), 3): veri(i, 15) = x(15)
                    veri(i, 16) = SA_YUVARLA(x(16), 3): veri(i, 17) = SA_YUVARLA(x(17), 3)
                    veri(i, 18) = SA_YUVARLA(x(19), 2): veri(i, 19) = SA_YUVARLA(x(20), 2): veri(i, 20) = x(21) & IIf(Len(x(23)) > 0, "  •  " & x(23), "")
                    anah(i) = SA_DURUM_SIRA(CStr(x(22))) & I_SIRA_ANAHTAR(CDbl(x(19)))
                End If
            Next x
            Y_SIRALA anah, idx, 1, n
            Dim s2() As Variant
            ReDim s2(1 To n, 1 To 20)
            For i = 1 To n
                For j = 1 To 20: s2(i, j) = veri(idx(i), j): Next j
            Next i
            veri = s2
        End If
        ust = "SATINALMA KONTROLÜ  |  " & IIf(tur = "T", "Açýk satýnalma talepleri (ME5A)", "Teslim alýnmamýþ sipariþler (ME2N)") & _
            "  |  SAP okumasý: " & Format$(tarih, "dd.mm.yyyy hh:nn")
        AKSORA_EXCEL_SAYFA c, Array("Durum", "Ýl", "ÜY", "Depo", IIf(tur = "T", "Talep no", "Sipariþ no"), "Kalem", _
            IIf(tur = "T", "Talep tarihi", "Belge tarihi"), IIf(tur = "T", "Talep eden", "Satýcý"), "Malzeme", "Kýsa metin", _
            IIf(tur = "T", "Talep miktarý", "Teslim edilecek"), "Birim", "Ayný ilde serbest stok", "Ayný ilde atýl stok", _
            "Atýl transfer ili", "Atýl transfer miktarý", "Atýl stoktan karþýlanabilir", "Tasarruf (TL)", IIf(tur = "T", "Talep tutarý", "Açýk tutar"), "PB / not"), _
            veri, n, Array(24, 11, 7, 7, 12, 7, 11, 20, 13, 40, 11, 7, 12, 12, 11, 11, 12, 16, 15, 30), _
            Array("@", "@", "@", "@", "@", "@", "@", "@", "0", "@", "General", "@", "General", "General", "@", "General", "General", _
                "#,##0.00 ""TL""", "General", "@"), _
            Array(xlCenter, xlLeft, xlCenter, xlCenter, xlCenter, xlCenter, xlCenter, xlLeft, xlCenter, xlLeft, xlRight, xlCenter, _
                xlRight, xlRight, xlCenter, xlRight, xlRight, xlRight, xlRight, xlLeft), ust, _
            "Tasarruf yalnýz atýl stoktan (paylaþtýrmalý)  |  Kýrmýzý: ayný ilin atýl stoðu  |  Turuncu: baþka ilin atýl stoðu (transfer)  |  Sarý: serbest stok var (bilgi)  |  Mor: sipariþten fazla teslim alýnmýþ", _
            IIf(tur = "T", "Açýk Talepler", "Açýk Sipariþler"), Empty
        ' Miktarlar: tam sayida ondalik ayirac yok, kusuratta virgul (altin kural).
        For Each j2 In Array(11, 13, 14, 16, 17)
            AKSORA_V13_MIKTAR_BICIMI c.Range(c.Cells(5, j2), c.Cells(n + 4, j2))
        Next j2
        ' Belge tutari SAP para biriminde (PB sutunu); TL disi olabilir, TL eki yalniz TL satirlarda.
        For i = 1 To n
            Set hz = c.Cells(i + 4, 1)
            Select Case CStr(hz.Value2)
                Case "ATIL STOKTAN KARÞILANIR": hz.Interior.Color = RGB(220, 38, 38): hz.Font.Color = vbWhite: hz.Font.bold = True
                Case "BAÞKA ÝLÝN ATIL STOÐUNDAN": hz.Interior.Color = RGB(253, 186, 116): hz.Font.Color = RGB(124, 45, 18): hz.Font.bold = True
                Case "AYNI ÝLDE SERBEST STOK VAR", "BAÞKA ÝLDE SERBEST STOK VAR": hz.Interior.Color = RGB(254, 249, 195): hz.Font.Color = RGB(113, 63, 18)
                Case "FAZLA TESLÝM ALINMIÞ": hz.Interior.Color = RGB(237, 233, 254): hz.Font.Color = RGB(91, 33, 182): hz.Font.bold = True
            End Select
            ' Belge tutari SAP para biriminde: TRY -> TL, digerleri kendi para birimiyle (USD, EUR...).
            t = Trim$(Split(CStr(c.Cells(i + 4, 20).Value2) & " ", " ")(0))
            If t = "TRY" Or t = "TL" Or Len(t) = 0 Then
                c.Cells(i + 4, 19).NumberFormat = "#,##0.00 ""TL"""
            Else
                c.Cells(i + 4, 19).NumberFormat = "#,##0.00 """ & t & """"
            End If
        Next i
    Next tur
    AKSORA_BEKLE_ILERLE 0.85, "Kritik stok sayfalarý..."
    KS_EXCEL_EKLE wb
    wb.Worksheets(1).Activate
    sonuc = AKSORA_EXCEL_BITIR(wb, AKSORA_MASAUSTU() & application.PathSeparator & _
        AKSORA_DOSYA_ADI("Satýn Alma - " & Format$(Date, "dd.mm.yyyy")) & ".xlsx", False)
    AN_ONB_KAPAT
    application.EnableEvents = ev: application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    If Len(sonuc) > 0 Then application.StatusBar = "Masaüstüne kaydedildi: " & sonuc
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    application.EnableEvents = ev: application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    AN_ONB_KAPAT
    MsgBox "Satýnalma kontrolü Excel'e aktarýlamadý." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

' Satici: basindaki SAP satici numarasi atilir (sayfada okunur boyutta kalsin; tam hali Excel'de).
Private Function SA_SATICI_AD(ByVal v As Variant) As String
    Dim s As String, p As Long
    s = Trim$(CStr(v)): p = InStr(s, " ")
    If p > 1 Then If IsNumeric(Left$(s, p - 1)) Then s = Trim$(Mid$(s, p + 1))
    SA_SATICI_AD = s
End Function

Private Function SA_YUVARLA(ByVal v As Variant, ByVal basamak As Long) As Variant
    If IsNumeric(v) And Not IsEmpty(v) And VarType(v) <> vbString Then SA_YUVARLA = Round(CDbl(v), basamak) Else SA_YUVARLA = v
End Function

Private Function SA_DURUM_SIRA(ByVal d As String) As String
    Select Case d
        Case "ATIL STOKTAN KARÞILANIR": SA_DURUM_SIRA = "1"
        Case "BAÞKA ÝLÝN ATIL STOÐUNDAN": SA_DURUM_SIRA = "2"
        Case "AYNI ÝLDE SERBEST STOK VAR": SA_DURUM_SIRA = "3"
        Case "BAÞKA ÝLDE SERBEST STOK VAR": SA_DURUM_SIRA = "4"
        Case "FAZLA TESLÝM ALINMIÞ": SA_DURUM_SIRA = "5"
        Case Else: SA_DURUM_SIRA = "6"
    End Select
End Function

' Yonetim Merkezi bolumu: il bazinda acik talep / siparis ve stoktan karsilanabilir tutar.
Public Function ADM_SA_CIZ(ByVal ws As Worksheet, ByVal r As Long, ByVal secIl As String) As Long
    Dim kayit As Collection, an As Collection, tarih As Date, bilgi As String, x As Variant, p As Long, iller As Variant, ilk As Long
    Dim v(0 To 4, 0 To 5) As Double, j As Long, durum As String
    ADM_SA_CIZ = r
    On Error GoTo Hata
    Set kayit = SA_KAYIT_OKU(tarih, bilgi)
    r = r + 2
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "SATINALMA KONTROLÜ   (ME5A açýk talepler • ME2N teslim alýnmamýþ sipariþler • atýl stokla  •  düðmeler üst çubukta)"
    r = r + 1
    If kayit.Count = 0 Then
        ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Henüz okunmadý: üst çubuktaki SATINALMA: SAP'DEN OKU ile açýk talepler ve sipariþler alýnýr; stoktan karþýlanabilenler burada görünür.", _
            RGB(244, 248, 252), RGB(71, 85, 105), 9, False, xlLeft
        ADM_SA_CIZ = r
        Exit Function
    End If
    ADM_ILERLE 0.97, "Satýnalma kontrolü eþleþtiriliyor..."
    Set an = SA_ANALIZ(kayit)
    iller = ADM_ILLER()
    For Each x In an
        For p = 0 To 4
            If x(2) = Y_IL_ADI(CStr(iller(p))) Then Exit For
        Next p
        If p <= 4 Then
            j = IIf(x(0) = "T", 0, 3)
            v(p, j) = v(p, j) + 1
            If x(17) > 0.0005 Then v(p, j + 1) = v(p, j + 1) + 1: v(p, j + 2) = v(p, j + 2) + x(19)
        End If
    Next x
    ilk = r
    ADM_BASLIKLAR ws, r, 2, Array("Ýl", "Açýk talep", "Atýl stoktan karþýlanabilir", "Talep tasarrufu", "", "Açýk sipariþ", _
        "Atýl stoktan karþýlanabilir", "Sipariþ tasarrufu", "", "Toplam tasarruf", "", "Durum")
    ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 9, r, 10)).Merge: ws.Range(ADM_ADRES(r, 11, r, 12)).Merge
    For p = 0 To 4
        r = r + 1
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = IIf(CStr(iller(p)) = secIl, RGB(226, 236, 247), vbWhite)
        ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 9, r, 10)).Merge: ws.Range(ADM_ADRES(r, 11, r, 12)).Merge
        ws.Cells(r, 2).Value2 = Y_IL_ADI(CStr(iller(p))): ws.Cells(r, 2).Font.bold = True
        ws.Cells(r, 3).Value2 = v(p, 0): ws.Cells(r, 4).Value2 = v(p, 1): ws.Cells(r, 5).Value2 = Round(v(p, 2), 2)
        ws.Cells(r, 7).Value2 = v(p, 3): ws.Cells(r, 8).Value2 = v(p, 4): ws.Cells(r, 9).Value2 = Round(v(p, 5), 2)
        ws.Cells(r, 11).Value2 = Round(v(p, 2) + v(p, 5), 2)
        If v(p, 1) + v(p, 4) > 0 Then durum = "KARÞILANABÝLÝR" Else durum = IIf(v(p, 0) + v(p, 3) > 0, "STOK YOK", "AÇIK KAYIT YOK")
        ws.Cells(r, 13).Value2 = durum
    Next p
    r = r + 1
    ws.Range(ADM_ADRES(r, 5, r, 6)).Merge: ws.Range(ADM_ADRES(r, 9, r, 10)).Merge: ws.Range(ADM_ADRES(r, 11, r, 12)).Merge
    ws.Cells(r, 2).Value2 = "TOPLAM"
    For Each x In Array(3, 4, 5, 7, 8, 9, 11)
        ws.Cells(r, CLng(x)).Formula = "=SUM(" & ADM_ADRES(ilk + 1, CLng(x), r - 1, CLng(x)) & ")"
    Next x
    ws.Range(ADM_ADRES(r, 2, r, 13)).Font.bold = True
    ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = RGB(226, 236, 247)
    ws.Range(ADM_ADRES(ilk + 1, 2, r, 13)).HorizontalAlignment = xlCenter
    For Each x In Array(5, 9, 11)
        ws.Range(ADM_ADRES(ilk + 1, CLng(x), r, CLng(x))).NumberFormat = "#,##0.00 ""TL"""
    Next x
    ws.Range(ADM_ADRES(ilk + 1, 11, r, 11)).Font.Color = RGB(21, 128, 61)
    ws.Range(ADM_ADRES(ilk + 1, 2, r, 13)).ShrinkToFit = True
    Y_DURUM_KF ws.Range(ADM_ADRES(ilk + 1, 13, r - 1, 13)), "KARÞILANABÝLÝR", RGB(254, 226, 226), RGB(185, 28, 28)
    Y_DURUM_KF ws.Range(ADM_ADRES(ilk + 1, 13, r - 1, 13)), "STOK YOK", RGB(241, 245, 249), RGB(100, 116, 139)
    ADM_KENAR ws.Range(ADM_ADRES(ilk, 2, r, 13)), True
    r = r + 1
    ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "SAP okumasý: " & Format$(tarih, "dd.mm.yyyy hh:nn") & "  •  " & bilgi & _
        "  •  Tasarruf yalnýz ATIL stoktan, paylaþtýrmalý (ayný stok iki kez sayýlmaz; ayný sistem CDP / EPP); serbest stok tasarrufa dahil deðil.", RGB(244, 248, 252), RGB(100, 116, 139), 8, False, xlLeft
    ws.Cells(r, 2).Font.Italic = True
    ADM_SA_CIZ = r
    Exit Function
Hata:
    On Error Resume Next
    r = r + 1
    ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Satýnalma kontrolü hesaplanamadý: " & Err.Description, RGB(254, 243, 199), RGB(146, 64, 14), 9, False, xlLeft
    ADM_SA_CIZ = r
End Function

'---------- SATINALMA KONTROLU: Malzeme Stok birimi sayfasi ----------
Public Function AKSORA_SATINALMA_SAYFA(ByVal sh As Object) As Boolean
    If Not TypeOf sh Is Worksheet Then Exit Function
    If sh.Name <> SA_SAYFA Then Exit Function
    AKSORA_SATINALMA_SAYFA = (CStr(sh.Range("BX1").Value2) = SA_ETIKET)
End Function

Public Sub AKSORA_SATINALMA_AC()
    Dim ac As Boolean, no As Long, metin As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  SATINALMA KONTROLÜ", "Satýnalma kontrolü hazýrlanýyor...")
    On Error GoTo Son
    SA_GOSTER vbNullString
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Satýnalma kontrolü sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_SATINALMA_IL_SEC()
    Dim ac As Boolean, no As Long, metin As String, il As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error Resume Next
    il = ThisWorkbook.Worksheets(SA_SAYFA).Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(il) = 0 Then il = "TUM"
    ac = AKSORA_BEKLE_AC("AKSORA  |  SATINALMA KONTROLÜ", "Satýnalma kontrolü hazýrlanýyor...")
    On Error GoTo Son
    SA_GOSTER il
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Satýnalma kontrolü sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Private Sub SA_GOSTER(ByVal il As String)
    Dim ws As Worksheet, ev As Boolean, hesap As XlCalculation, no As Long, metin As String
    ev = application.EnableEvents: hesap = application.Calculation
    AN_ONB_AC
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    Set ws = SA_SAYFA_KUR()
    AN_ESKI_SAYFA_SIL KS_SAYFA, KS_ETIKET
    If Len(il) = 0 Then il = CStr(ws.Range("BX20").Value2)
    If Len(il) = 0 Then il = "TUM"
    SA_CIZ ws, il
    ws.Visible = xlSheetVisible
    mYonlendirme = True
    ws.Select Replace:=True
    mYonlendirme = False
    AKSORA_YATIRIM_GORUNUM ws
Cikis:
    On Error Resume Next
    AN_ONB_KAPAT
    mYonlendirme = False
    application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If no <> 0 Then Err.Raise no, "SA_GOSTER", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Resume Cikis
End Sub

' Sayfa musteri rapor sayfasinin basligiyla kurulur (atil stok sayfasi gibi).
Private Function SA_SAYFA_KUR() As Worksheet
    Dim ws As Worksheet, kaynak As Worksheet, gor As XlSheetVisibility, onceki As Collection, sh As Object, varmi As Boolean
    Dim da As Boolean
    Set ws = ADM_WS(SA_SAYFA)
    If Not ws Is Nothing Then
        If CStr(ws.Range("BX1").Value2) <> SA_ETIKET Then Err.Raise vbObjectError + 2370, , "Ayný isimde baþka bir sayfa var: " & SA_SAYFA
        Set SA_SAYFA_KUR = ws
        Exit Function
    End If
    Set kaynak = ADM_WS("M_CDP_RAPOR")
    If kaynak Is Nothing Then Err.Raise vbObjectError + 2371, , "Sayfa þablonu (M_CDP_RAPOR) bulunamadý."
    da = application.DisplayAlerts: application.DisplayAlerts = False
    gor = kaynak.Visible
    kaynak.Visible = xlSheetVisible
    Set onceki = New Collection
    For Each sh In ThisWorkbook.Sheets
        onceki.Add sh.Name, sh.Name
    Next sh
    kaynak.Copy after:=kaynak
    For Each sh In ThisWorkbook.Sheets
        varmi = False
        On Error Resume Next
        varmi = (Len(onceki(sh.Name)) > 0)
        On Error GoTo 0
        If Not varmi Then Set ws = sh: Exit For
    Next sh
    kaynak.Visible = gor
    application.DisplayAlerts = da
    If ws Is Nothing Then Err.Raise vbObjectError + 2372, , "Kopya sayfa bulunamadý."
    ws.Name = SA_SAYFA
    ws.Unprotect ""
    ws.Range("BX1:BX40").ClearContents
    ws.Range("BX1").Value2 = SA_ETIKET
    On Error Resume Next
    ws.Move after:=ThisWorkbook.Worksheets(ATIL_SAYFA)
    On Error GoTo 0
    Set SA_SAYFA_KUR = ws
End Function

Private Sub SA_CIZ(ByVal ws As Worksheet, ByVal ilSec As String)
    ' Tek sayfa, uc soru: ne satin alinmali (kritik stok), hangi talep / siparis gereksiz (atil stok var), ne fazla teslim alindi.
    ' Transferler TRANSFER LISTESI'nde; tam listeler EXCEL'E AKTAR'da.
    Dim kayit As Collection, an As Collection, ks As Collection, tarih As Date, bilgi As String, okuma As Variant, saT As Date
    Dim x As Variant, w As Double, s As Shape, y As Double, kw As Double, bw As Double, i As Long, r As Long, a1 As Long, n As Long
    Dim il As Variant, sw As Double, sonuc() As Variant, sec As Boolean, anah() As String, idx() As Long, blok() As Variant
    Dim nAl As Long, tAl As Double, nIp As Long, tIp As Double, nTr As Long, tTr As Double, nF As Long, tF As Double, bilinmez As Boolean
    ws.Unprotect ""
    Y_SAYFA_SIFIRLA ws
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(ws.Shapes(i).Name, "FILIGRAN") > 0 Then ws.Shapes(i).Delete
    Next i
    Set s = Y_SEKIL(ws, "btnAnaPanelDon")
    If Not s Is Nothing Then s.AlternativeText = vbNullString: s.OnAction = BIRIM_MAKRO("ANA_PANELE_DON")
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | SATIN ALMA  |  " & IIf(ilSec = "TUM", "Tüm Ýller", Y_IL_ADI(ilSec))
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    sw = application.Max(180, w - 1155)
    Y_KOLONLAR ws, Array(105, 95, 120, 110, 300, 110, 120, 65, 130, sw)
    ws.Range("BX20").Value2 = ilSec
    Set kayit = SA_KAYIT_OKU(tarih, bilgi)
    If kayit.Count > 0 Then Set an = SA_ANALIZ(kayit) Else Set an = New Collection
    Set ks = KS_ANALIZ(okuma, saT)
    For Each x In ks
        If ilSec = "TUM" Or x(0) = ilSec Then
            If x(14) > 0.0005 Then nAl = nAl + 1: tAl = tAl + x(16)
            If x(18) > 0.0005 Then nTr = nTr + 1: tTr = tTr + x(18) * x(15)
        End If
        If x(22) < 0 Then bilinmez = True
    Next x
    For Each x In an
        If ilSec = "TUM" Or ATIL_TESIS_IL(CStr(x(3))) = ilSec Then
            If x(17) > 0.0005 Then nIp = nIp + 1: tIp = tIp + x(19)
            If x(22) = "FAZLA TESLÝM ALINMIÞ" Then nF = nF + 1: tF = tF + Abs(x(20))
        End If
    Next x
    ws.Range("BZ21").Value2 = nAl & " kalem  •  " & AKSORA_SAYI_METNI(Round(tAl, 2), "#,##0.00") & " TL"
    ws.Range("BZ22").Value2 = nIp & " kalem  •  " & AKSORA_SAYI_METNI(Round(tIp, 2), "#,##0.00") & " TL"
    ws.Range("BZ23").Value2 = nTr & " kalem  •  " & AKSORA_SAYI_METNI(Round(tTr, 2), "#,##0.00") & " TL"
    ws.Range("BZ24").Value2 = nF & " kalem"
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10: ws.Rows(8).RowHeight = 30: ws.Rows(9).RowHeight = 18
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "SATIN ALINMASI GEREKEN (KRÝTÝK STOK)", "$BZ$21", RGB(12, 45, 79)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "GEREKSÝZ TALEP / SÝPARÝÞ (ATIL VAR)", "$BZ$22", RGB(220, 38, 38)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "TRANSFERLE ÇÖZÜLEN KRÝTÝK STOK", "$BZ$23", RGB(39, 160, 94)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "FAZLA TESLÝM ALINAN SÝPARÝÞ", "$BZ$24", RGB(124, 58, 237)
    For i = 1 To 3
        Set s = Y_SEKIL(ws, "Y_KPI_DEGER_" & i)
        If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 14
    Next i
    y = ws.Rows(8).Top + 1
    bw = application.Min(118, (w - 3 * 150 - 8 * 6) / 6)
    i = 0
    For Each il In Array("TUM", "ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
        Set s = Y_DUGME(ws, "SA_IL_" & i, IIf(CStr(il) = "TUM", "TÜM ÝLLER", ATIL_IL_BUYUK(CStr(il))), i * (bw + 6), y, bw, 28, _
            IIf(CStr(il) = ilSec, RGB(0, 125, 110), RGB(148, 163, 184)), "AKSORA_SATINALMA_IL_SEC")
        s.AlternativeText = CStr(il)
        i = i + 1
    Next il
    Y_DUGME ws, "SA_TRANSFER", "TRANSFER LÝSTESÝ", w - 450 - 12, y, 150, 28, RGB(39, 160, 94), "AKSORA_TRANSFER_AC"
    Y_DUGME ws, "SA_SAP", "SAP'DEN OKU (ME5A / ME2N)", w - 300 - 6, y, 150, 28, RGB(20, 103, 173), "AKSORA_SATINALMA_OKU"
    Y_DUGME ws, "SA_EXCEL", "EXCEL'E AKTAR", w - 150, y, 150, 28, RGB(33, 115, 70), "AKSORA_SATINALMA_EXCEL"
    ws.Range("A9").Value2 = "Satýn alma: stok 3 aydan az yetiyor (son 6 ay ortalama sarf; tek seferlik proje sarfý hariç)  •  Gereksiz: açýk talep / sipariþ kendi ilinin ya da baþka ilin ATIL stoðundan karþýlanýr  •  " & _
        "SAP: " & IIf(kayit.Count > 0, Format$(tarih, "dd.mm.yyyy hh:nn"), "ME5A / ME2N okunmadý") & IIf(bilinmez, "  •  Tek seferlik sarf ayrýmý için ATIL STOK'ta SAP'den yeniden okuyun", "")
    With ws.Range("A9").Font
        .Name = "Aptos": .Size = 9: .Italic = True: .Color = RGB(100, 116, 139): .bold = False
    End With
    If kayit.Count > 0 Then AN_YAS_YAZ ws, CDbl(tarih), "SATINALMA (ME5A / ME2N)"
    If IsArray(okuma) Then AN_YAS_YAZ ws, AN_EN_ESKI(CDbl(okuma(0)), CDbl(okuma(1))), "SARF VERÝSÝ (MB51)"
    '--- 1) satin alinmasi gerekenler
    r = 11
    Y_BOLUM ws, r, "1  •  SATIN ALINMASI GEREKENLER  •  " & nAl & " kalem  •  " & AKSORA_SAYI_METNI(Round(tAl, 2), "#,##0.00") & " TL", _
        "En acil üstte. Atýl / fazla stokla karþýlanan kýsým düþüldü (TRANSFER LÝSTESÝ).", 10
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)), Array("Ýl", "Ne kadar yeter", "Stok", "Malzeme", "Kýsa metin", _
        "Alýnacak miktar", "Açýk talep / sipariþ", "Birim", "Tahmini tutar", "Durum")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    If nAl > 0 Then
        ReDim sonuc(1 To nAl, 1 To 10)
        i = 0
        For Each x In ks
            If (ilSec = "TUM" Or x(0) = ilSec) And x(14) > 0.0005 Then
                i = i + 1
                sonuc(i, 1) = x(1) & " " & x(2): sonuc(i, 2) = AN_SURE_METNI(x(10)): sonuc(i, 3) = Round(x(7) + x(8), 3): sonuc(i, 4) = x(4)
                sonuc(i, 5) = x(5): sonuc(i, 6) = x(14): sonuc(i, 7) = Round(x(11) + x(12), 3): sonuc(i, 8) = x(6): sonuc(i, 9) = x(16)
                sonuc(i, 10) = x(19)
            End If
        Next x
        ws.Range(ws.Cells(a1, 2), ws.Cells(a1 + nAl - 1, 2)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 4), ws.Cells(a1 + nAl - 1, 5)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 8), ws.Cells(a1 + nAl - 1, 8)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 1), ws.Cells(a1 + nAl - 1, 10)).Value2 = sonuc
        r = a1 + nAl - 1
        SA_TABLO_BICIM ws, a1, r
        AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7))
        ws.Range(ws.Cells(a1, 2), ws.Cells(r, 2)).HorizontalAlignment = xlCenter
        AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 3), ws.Cells(r, 3))
        Y_DURUM_KF ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)), "KRÝTÝK", RGB(220, 38, 38), RGB(255, 255, 255)
        Y_DURUM_KF ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)), "AZALIYOR", RGB(253, 186, 116), RGB(124, 45, 18)
        ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).HorizontalAlignment = xlCenter
    Else
        r = a1
        ws.Cells(r, 5).Value2 = IIf(ks.Count = 0, "Sarf verisi yok: ATIL STOK sayfasýnda SAP'DEN OKU (MB51).", "Satýn alýnmasý gereken kritik malzeme yok.")
        ws.Cells(r, 5).Font.Italic = True: ws.Cells(r, 5).Font.Color = RGB(100, 116, 139)
    End If
    '--- 2) iptal / azaltilacak talep ve siparisler (atil stoktan karsilanir)
    r = r + 2
    Y_BOLUM ws, r, "2  •  ÝPTAL / AZALTILACAK TALEP VE SÝPARÝÞLER  •  " & nIp & " kalem  •  " & AKSORA_SAYI_METNI(Round(tIp, 2), "#,##0.00") & " TL", _
        "Atýl stok var: talebi iptal et / sipariþi azalt; baþka ildeyse TRANSFER LÝSTESÝ'ndeki nakli yap.", 10
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)), Array("Ýl", "Tür", "Belge no", "Malzeme", "Kýsa metin", _
        "Açýk miktar", "Atýldan karþýlanýr", "Birim", "Kazanç", "Atýl stok nerede")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    If nIp > 0 Then
        ReDim anah(1 To nIp): ReDim idx(1 To nIp): ReDim blok(1 To nIp, 1 To 10)
        i = 0
        For Each x In an
            If (ilSec = "TUM" Or ATIL_TESIS_IL(CStr(x(3))) = ilSec) And x(17) > 0.0005 Then
                i = i + 1: idx(i) = i
                blok(i, 1) = x(2) & " " & x(1): blok(i, 2) = IIf(x(0) = "T", "TALEP", "SÝPARÝÞ")
                blok(i, 3) = x(5) & IIf(Len(x(6)) > 0, " / " & x(6), ""): blok(i, 4) = x(9): blok(i, 5) = x(10)
                blok(i, 6) = x(11): blok(i, 7) = x(17): blok(i, 8) = x(12): blok(i, 9) = x(19)
                If x(16) > 0.0005 And x(17) - x(16) > 0.0005 Then
                    blok(i, 10) = x(2) & " + " & x(15)
                ElseIf x(16) > 0.0005 Then
                    blok(i, 10) = x(15) & " (nakil)"
                Else
                    blok(i, 10) = x(2) & " (kendi deposu)"
                End If
                anah(i) = I_SIRA_ANAHTAR(CDbl(x(19)))
            End If
        Next x
        Y_SIRALA anah, idx, 1, nIp
        ReDim sonuc(1 To nIp, 1 To 10)
        For i = 1 To nIp
            For n = 1 To 10: sonuc(i, n) = blok(idx(i), n): Next n
        Next i
        ws.Range(ws.Cells(a1, 2), ws.Cells(a1 + nIp - 1, 5)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 8), ws.Cells(a1 + nIp - 1, 8)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 10), ws.Cells(a1 + nIp - 1, 10)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 1), ws.Cells(a1 + nIp - 1, 10)).Value2 = sonuc
        r = a1 + nIp - 1
        SA_TABLO_BICIM ws, a1, r
        ws.Range(ws.Cells(a1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
        AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7))
        ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)).Font.Color = RGB(21, 128, 61)
        ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).ShrinkToFit = True
    Else
        r = a1
        ws.Cells(r, 5).Value2 = IIf(kayit.Count = 0, "ME5A / ME2N okunmadý: SAP'DEN OKU.", "Atýl stoktan karþýlanabilen açýk talep / sipariþ yok.")
        ws.Cells(r, 5).Font.Italic = True: ws.Cells(r, 5).Font.Color = RGB(100, 116, 139)
    End If
    '--- 3) fazla teslim alinmis siparisler
    r = r + 2
    Y_BOLUM ws, r, "3  •  FAZLA TESLÝM ALINAN SÝPARÝÞLER  •  " & nF & " kalem", "Sipariþten fazla teslim alýnmýþ: iade / fatura kontrolü.", 10
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 10)), Array("Ýl", "ÜY", "Sipariþ no / kalem", "Malzeme", "Kýsa metin", _
        "Fazla miktar", "Belge tarihi", "Birim", "Fazla tutar", "Satýcý")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    If nF > 0 Then
        ReDim sonuc(1 To nF, 1 To 10)
        i = 0
        For Each x In an
            If x(22) = "FAZLA TESLÝM ALINMIÞ" And (ilSec = "TUM" Or ATIL_TESIS_IL(CStr(x(3))) = ilSec) Then
                i = i + 1
                sonuc(i, 1) = x(2): sonuc(i, 2) = x(3): sonuc(i, 3) = x(5) & IIf(Len(x(6)) > 0, " / " & x(6), ""): sonuc(i, 4) = x(9)
                sonuc(i, 5) = x(10): sonuc(i, 6) = -x(11): sonuc(i, 7) = x(7): sonuc(i, 8) = x(12): sonuc(i, 9) = Abs(x(20))
                sonuc(i, 10) = SA_SATICI_AD(x(8))
            End If
        Next x
        ws.Range(ws.Cells(a1, 2), ws.Cells(a1 + nF - 1, 5)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 7), ws.Cells(a1 + nF - 1, 8)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 10), ws.Cells(a1 + nF - 1, 10)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 1), ws.Cells(a1 + nF - 1, 10)).Value2 = sonuc
        r = a1 + nF - 1
        SA_TABLO_BICIM ws, a1, r
        ws.Range(ws.Cells(a1, 2), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).Font.Color = RGB(91, 33, 182)
        ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)).Font.Color = RGB(91, 33, 182)
        ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10)).ShrinkToFit = True
    Else
        r = a1
        ws.Cells(r, 5).Value2 = "Fazla teslim alýnmýþ sipariþ yok."
        ws.Cells(r, 5).Font.Italic = True: ws.Cells(r, 5).Font.Color = RGB(100, 116, 139)
    End If
    r = r + 2
    ws.Cells(r, 1).Value2 = "Tam listeler (tüm açýk talepler / sipariþler, yolda olanlar, yeterli ve fazla stok): EXCEL'E AKTAR.  Tutarlar: miktar x AKSORA stok birim deðeri."
    ws.Cells(r, 1).Font.Size = 9: ws.Cells(r, 1).Font.Color = RGB(71, 85, 105)
    Y_FILIGRAN ws
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

' Ortak tablo bicimi (10 sutun): kenar/zebra, il ve malzeme ortada, 6-7 miktar, 8 birim, 9 TL.
Private Sub SA_TABLO_BICIM(ByVal ws As Worksheet, ByVal a1 As Long, ByVal r As Long)
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 10))
    ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 21
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 1)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).IndentLevel = 1
    ws.Range(ws.Cells(a1, 8), ws.Cells(r, 8)).HorizontalAlignment = xlCenter
    AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6))
    ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).Font.bold = True
    ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)).Font.bold = True
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
End Sub

'---------- KRITIK STOK (Malzeme Stok birimi) ----------
' Kaynaklar (ek SAP okumasi yok): stok = CDP/EPP/ILCE/YUKLENICI stok sayfalari (D/K + yuklenici depolari, ATIL STOK ile ayni kapsam);
' sarf = ATIL STOK SAP okumasi (MB51, son 6 ay; 201/221/241/261/281/291 ve iptalleri); acik talep / siparis = SATINALMA KONTROL okumasi.
' Aylik sarf = son 6 ay sarf / 6. Kapsam = stok / aylik sarf. Hedef: 3 aylik ihtiyac.
' KRITIK: stok 1 aydan az yeter ve yoldakilerle 3 ayi bulmaz. AZALIYOR: 3 aydan az ve yoldakilerle 3 ayi bulmaz.
' YOLDA: stok 3 aydan az ama acik talep / siparisle 3 ay karsilanir. YETERLI: 3-12 ay. FAZLA STOK: 12 aydan fazla.
' DUZENSIZ SARF: son 6 ayin 3'unden azinda sarf var (proje / tek seferlik cikis); stok 3 aydan az olsa da kritik sayilmaz.
'   (Aylik dagilim ATIL STOK SAP okumasiyla gelir; eski kayitta yoksa butun sarf duzenli kabul edilir.)
' Onerilen talep = 3 aylik ihtiyac - (stok + acik talep + acik siparis) - diger illerin atil stogundan transfer (ayni sistem, paylastirmali).
' Donus: Array(0 il kodu, 1 il, 2 sistem, 3 UY, 4 malzeme, 5 tanim, 6 birim, 7 serbest stok, 8 yuklenici stok, 9 aylik sarf,
'   10 kapsam ay (-1: sarf yok), 11 acik talep, 12 acik siparis, 13 yoldakilerle kapsam, 14 onerilen talep, 15 birim deger,
'   16 tahmini tutar, 17 atil transfer metni, 18 atil transfer miktari, 19 durum, 20 son 6 ay sarf, 21 son sarf tarihi,
'   22 sarf olan ay sayisi (-1: bilinmiyor), 23 atil transfer kaynaklari "ILKODU=miktar;...")
Private Function KS_ANALIZ(ByRef okuma As Variant, ByRef saTarih As Date) As Collection
    Dim tum As Object, kayit As Object, havuz As Object, acik As Object, c As Collection, x As Variant, r As Variant, bilgi As Variant
    Dim k As Variant, uy As String, mal As String, sis As String, il As String, kr As Variant, ok As Double, sarf6 As Double
    Dim aylik As Double, serbest As Double, yuk As Double, bd As Double, tanim As String, birim As String, acT As Double, acS As Double
    Dim kap As Double, kapG As Double, ihtiyac As Double, durum As String, oneri As Double, tMik As Double, tMetin As String
    Dim il2 As Variant, uy2 As String, enIl As String, enMik As Double, al As Double, ham As Collection, idx() As Long
    Dim anah() As String, n As Long, i As Long, a As Variant, sBilgi As String, q As Double, ak As String, aySay As Long
    Dim saKayit As Collection, kk As Variant, tKay As String, hk As String, atilIlk As Object, fHavuz As Object, fz As Double
    Set KS_ANALIZ = New Collection
    Set kayit = ATIL_KAYIT_OKU(okuma)
    If kayit.Count = 0 Then Exit Function
    Set tum = ATIL_TUM_STOK()
    ' Diger illerin atil stogu (transfer havuzu): ATIL STOK hesabiyla ayni (en eski katman).
    Set havuz = CreateObject("Scripting.Dictionary"): havuz.CompareMode = vbTextCompare
    On Error Resume Next
    Set c = ATIL_HESAPLA(bilgi)
    On Error GoTo 0
    If Not c Is Nothing Then
        For Each x In c
            If IsArray(x(20)) Then If x(20)(3) > 0.0005 Then havuz(CStr(x(2)) & "|" & CStr(x(3))) = CDbl(x(20)(3))
        Next x
    End If
    Set atilIlk = CreateObject("Scripting.Dictionary"): atilIlk.CompareMode = vbTextCompare
    For Each kk In havuz.keys
        atilIlk(CStr(kk)) = havuz(kk)
    Next kk
    Set fHavuz = CreateObject("Scripting.Dictionary"): fHavuz.CompareMode = vbTextCompare
    ' Acik talep / siparis (UY|malzeme -> Array(talep, siparis, birim))
    Set acik = CreateObject("Scripting.Dictionary"): acik.CompareMode = vbTextCompare
    Set saKayit = SA_KAYIT_OKU(saTarih, sBilgi)
    ' Satinalma kontrolunde acik talep / siparis icin ayrilan atil stok kritik stok transferinde tekrar kullanilmaz.
    If saKayit.Count > 0 Then
        For Each x In SA_ANALIZ(saKayit)
            If UBound(x) >= 24 Then
                For Each kk In AN_KAYNAKLAR(CStr(x(24)))
                    hk = AKSORA_IL_TESIS(CStr(kk(0)), CStr(x(1))) & "|" & Y_KOD_NORM(x(9))
                    If havuz.Exists(hk) Then havuz(hk) = application.Max(0, havuz(hk) - kk(1))
                Next kk
            End If
        Next x
    End If
    For Each r In saKayit
        ak = Trim$(CStr(r(4))) & "|" & Y_KOD_NORM(r(8))
        If r(0) = "T" Then q = SA_SAYI(r(11)) Else q = SA_SAYI(r(15))
        If q > 0.0005 And Len(ak) > 2 Then
            If Not acik.Exists(ak) Then acik.Add ak, Array(0#, 0#, Trim$(CStr(r(12))), Trim$(CStr(r(9))))
            a = acik(ak)
            If r(0) = "T" Then a(0) = a(0) + q Else a(1) = a(1) + q
            acik(ak) = a
        End If
    Next r
    ' 1. gecis: hesap
    Set ham = New Collection
    For Each k In kayit.keys
        n = n + 1: If (n And 255) = 0 Then AKSORA_BEKLE_NABIZ
        kr = kayit(k)
        uy = Left$(CStr(k), InStr(CStr(k), "|") - 1): mal = Mid$(CStr(k), InStr(CStr(k), "|") + 1)
        sis = ATIL_SISTEM(uy): il = ATIL_TESIS_IL(uy)
        If sis = "CDP" Then ok = okuma(0) Else ok = okuma(1)
        sarf6 = kr(3)
        If ok > 0 And sarf6 > 0.0005 And Len(il) > 0 Then
            serbest = 0: yuk = 0: bd = 0
            SA_STOK tum, il, CStr(k), serbest, yuk, bd
            If bd <= 0 Then bd = SA_BIRIM_DEGER(tum, sis, mal)
            KS_MAL_BILGI tum, sis, mal, tanim, birim
            acT = 0: acS = 0
            If acik.Exists(CStr(k)) Then
                a = acik(CStr(k))
                If Len(birim) = 0 Or Len(CStr(a(2))) = 0 Or UCase$(CStr(a(2))) = UCase$(birim) Then acT = a(0): acS = a(1)
                If Len(tanim) = 0 Then tanim = CStr(a(3))
                If Len(birim) = 0 Then birim = CStr(a(2))
            End If
            aylik = sarf6 / 6
            kap = (serbest + yuk) / aylik
            kapG = (serbest + yuk + acT + acS) / aylik
            aySay = KS_AY_SAYISI(CStr(kr(8)))
            ' Aylik dagilim yoksa da kesin olan: adetli malzemede 6 ayda 3 adetten az sarf en fazla 2 ayda olabilir.
            If aySay < 0 And sarf6 < 2.9995 And KS_YUKARI(0.5, birim) = 1 Then aySay = Int(sarf6 + 0.0005)
            If aySay >= 0 And aySay < 3 And kap < 3 Then
                durum = "DÜZENSÝZ SARF"
            ElseIf kap < 1 And kapG < 3 Then
                durum = "KRÝTÝK"
            ElseIf kap < 3 And kapG < 3 Then
                durum = "AZALIYOR"
            ElseIf kap < 3 Then
                durum = "YOLDA"
            ElseIf kap <= 12 Then
                durum = "YETERLÝ"
            Else
                durum = "FAZLA STOK"
            End If
            ham.Add Array(il, Y_IL_ADI(il), sis, uy, mal, tanim, birim, serbest, yuk, aylik, kap, acT, acS, kapG, 0#, bd, 0#, _
                vbNullString, 0#, durum, sarf6, kr(4), aySay, vbNullString)
        End If
    Next k
    If ham.Count = 0 Then Exit Function
    ' 2. gecis: en kritikten baslayarak atil transfer paylastirmasi ve onerilen talep.
    ReDim anah(1 To ham.Count): ReDim idx(1 To ham.Count)
    i = 0
    For Each x In ham
        i = i + 1: idx(i) = i
        anah(i) = KS_DURUM_SIRA(CStr(x(19))) & Format$(application.Min(x(10), 9999) * 1000, "0000000000")
    Next x
    Y_SIRALA anah, idx, 1, ham.Count
    For i = 1 To ham.Count
        x = ham(idx(i))
        If x(19) = "KRÝTÝK" Or x(19) = "AZALIYOR" Then
            ihtiyac = 3 * x(9) - (x(7) + x(8) + x(11) + x(12))
            ' Adetli malzemede transfer de satin alma da tam sayi (1,5 adet transfer onerilmez).
            If ihtiyac > 0.0005 Then ihtiyac = KS_YUKARI(ihtiyac, CStr(x(6)))
            tMik = 0: tMetin = vbNullString: tKay = vbNullString
            Do While ihtiyac > 0.0005
                enIl = vbNullString: enMik = 0
                For Each il2 In ADM_ILLER()
                    If CStr(il2) <> CStr(x(0)) Then
                        uy2 = AKSORA_IL_TESIS(CStr(il2), CStr(x(2))) & "|" & CStr(x(4))
                        If havuz.Exists(uy2) Then If havuz(uy2) > enMik Then enMik = havuz(uy2): enIl = CStr(il2)
                    End If
                Next il2
                If enMik <= 0.0005 Then Exit Do
                al = application.Min(ihtiyac, enMik)
                havuz(AKSORA_IL_TESIS(enIl, CStr(x(2))) & "|" & CStr(x(4))) = enMik - al
                ihtiyac = ihtiyac - al: tMik = tMik + al
                tMetin = tMetin & IIf(Len(tMetin) > 0, ", ", "") & Y_IL_ADI(enIl) & ": " & ADM_ADET(al)
                tKay = tKay & IIf(Len(tKay) > 0, ";", "") & enIl & "=" & AN_SAYI_YAZ(al)
            Loop
            ' Atil yetmezse: diger illerin FAZLA stogu (serbest stogun atil disi ve 12 aylik ihtiyacini asan kismi).
            Do While ihtiyac > 0.0005
                enIl = vbNullString: enMik = 0
                For Each il2 In ADM_ILLER()
                    If CStr(il2) <> CStr(x(0)) Then
                        uy2 = AKSORA_IL_TESIS(CStr(il2), CStr(x(2))) & "|" & CStr(x(4))
                        If Not fHavuz.Exists(uy2) Then fHavuz(uy2) = KS_FAZLA(tum, kayit, atilIlk, CStr(il2), uy2, CStr(x(6)))
                        If fHavuz(uy2) > enMik Then enMik = fHavuz(uy2): enIl = CStr(il2)
                    End If
                Next il2
                If enMik <= 0.0005 Then Exit Do
                al = application.Min(ihtiyac, enMik)
                fHavuz(AKSORA_IL_TESIS(enIl, CStr(x(2))) & "|" & CStr(x(4))) = enMik - al
                ihtiyac = ihtiyac - al: tMik = tMik + al
                tMetin = tMetin & IIf(Len(tMetin) > 0, ", ", "") & Y_IL_ADI(enIl) & " (fazla): " & ADM_ADET(al)
                tKay = tKay & IIf(Len(tKay) > 0, ";", "") & "F:" & enIl & "=" & AN_SAYI_YAZ(al)
            Loop
            If ihtiyac > 0.0005 Then oneri = KS_YUKARI(ihtiyac, CStr(x(6))) Else oneri = 0
            x(14) = oneri: x(16) = Round(oneri * x(15), 2): x(17) = tMetin: x(18) = tMik: x(23) = tKay
        End If
        KS_ANALIZ.Add x
    Next i
End Function

Private Function KS_DURUM_SIRA(ByVal d As String) As String
    Select Case d
        Case "KRÝTÝK": KS_DURUM_SIRA = "1"
        Case "AZALIYOR": KS_DURUM_SIRA = "2"
        Case "YOLDA": KS_DURUM_SIRA = "3"
        Case "DÜZENSÝZ SARF": KS_DURUM_SIRA = "4"
        Case "YETERLÝ": KS_DURUM_SIRA = "5"
        Case Else: KS_DURUM_SIRA = "6"
    End Select
End Function

' Son 6 ayda sarf olan ay sayisi; aylik dagilim yoksa -1.
Private Function KS_AY_SAYISI(ByVal t As String) As Long
    Dim p As Variant, n As Long, x As Double
    If Len(Trim$(t)) = 0 Then KS_AY_SAYISI = -1: Exit Function
    For Each p In Split(t, ";")
        If InStr(CStr(p), "=") > 0 Then
            x = Val(Mid$(CStr(p), InStr(CStr(p), "=") + 1))
            If x > 0.0005 Then n = n + 1
        End If
    Next p
    KS_AY_SAYISI = n
End Function

' Bir ilin fazla stogu: serbest (D/K) stok - atil kismi - 12 aylik ihtiyac; adetli birimde asagi tam sayi.
Private Function KS_FAZLA(ByVal tum As Object, ByVal kayit As Object, ByVal atilIlk As Object, ByVal il As String, _
    ByVal k As String, ByVal birim As String) As Double
    Dim serbest As Double, yuk As Double, bd As Double, aylik As Double, atil As Double, f As Double
    bd = -1
    SA_STOK tum, il, k, serbest, yuk, bd
    If serbest <= 0.0005 Then Exit Function
    If kayit.Exists(k) Then aylik = application.Max(0, CDbl(kayit(k)(3))) / 6
    If atilIlk.Exists(k) Then atil = atilIlk(k)
    f = serbest - atil - 12 * aylik
    If f <= 0.0005 Then Exit Function
    If KS_YUKARI(0.5, birim) = 1 Then f = Int(f + 0.000001)
    KS_FAZLA = f
End Function

' Adetli birimlerde yukari tam sayiya, digerlerinde 2 basamaga yukari yuvarlanir.
Private Function KS_YUKARI(ByVal x As Double, ByVal birim As String) As Double
    Select Case UCase$(Trim$(birim))
        Case "ADT", "AD", "SET", "TK", "ST", "PAK", "KUT", "TOP", "RUL", "CFT", "ÇFT"
            KS_YUKARI = -Int(-Round(x, 6))
        Case Else
            KS_YUKARI = -Int(-Round(x * 100, 6)) / 100
    End Select
End Function

' Malzeme tanimi ve birimi: once ayni sistemde stogu olan ilden.
Private Sub KS_MAL_BILGI(ByVal tum As Object, ByVal sis As String, ByVal mal As String, ByRef tanim As String, ByRef birim As String)
    Dim il As Variant, k As String, s As Object
    tanim = vbNullString: birim = vbNullString
    For Each il In ADM_ILLER()
        If tum.Exists(CStr(il)) Then
            Set s = tum(CStr(il))(0)
            k = AKSORA_IL_TESIS(CStr(il), sis) & "|" & mal
            If s.Exists(k) Then
                If Len(tanim) = 0 Then tanim = Trim$(CStr(s(k)(3)))
                If Len(birim) = 0 Then birim = Trim$(CStr(s(k)(2)))
                If Len(tanim) > 0 And Len(birim) > 0 Then Exit Sub
            End If
        End If
    Next il
End Sub

Private Function KS_KAPSAM_METNI(ByVal x As Double) As Variant
    If x < 0 Then KS_KAPSAM_METNI = "-" Else KS_KAPSAM_METNI = Round(x, 1)
End Function

Public Function AKSORA_KRITIK_SAYFA(ByVal sh As Object) As Boolean
    If Not TypeOf sh Is Worksheet Then Exit Function
    If sh.Name <> KS_SAYFA Then Exit Function
    AKSORA_KRITIK_SAYFA = (CStr(sh.Range("BX1").Value2) = KS_ETIKET)
End Function

Public Sub AKSORA_KRITIK_AC()
    Dim ac As Boolean, no As Long, metin As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    ' Kritik stok SATIN ALMA sayfasinda (eski dugmeler bu sayfayi acar).
    ac = AKSORA_BEKLE_AC("AKSORA  |  SATIN ALMA", "Satýn alma sayfasý hazýrlanýyor...")
    On Error GoTo Son
    SA_GOSTER vbNullString
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Kritik stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_KRITIK_IL_SEC()
    Dim ac As Boolean, no As Long, metin As String, il As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error Resume Next
    il = ThisWorkbook.Worksheets(KS_SAYFA).Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(il) = 0 Then il = "TUM"
    ac = AKSORA_BEKLE_AC("AKSORA  |  KRÝTÝK STOK", "Kritik stok hesaplanýyor...")
    On Error GoTo Son
    SA_GOSTER il
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Kritik stok sayfasý açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

' Satin alma Excel'ine kritik stok sayfalari: "Satýn Alýnacaklar" (kritik + azalan) ve "Stok Kapsamý" (son 6 ayda sarfi olan tum malzemeler).
Private Sub KS_EXCEL_EKLE(ByVal wb As Workbook)
    Dim an As Collection, okuma As Variant, saTarih As Date, c As Worksheet, veri() As Variant, x As Variant, n As Long, i As Long
    Dim sayfa As Long, ust As String, hz As Range, j2 As Variant
    Set an = KS_ANALIZ(okuma, saTarih)
    If an.Count = 0 Then Exit Sub
    For sayfa = 1 To 2
        n = 0
        For Each x In an
            If sayfa = 2 Or x(14) > 0.0005 Then n = n + 1
        Next x
        Set c = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        If n = 0 Then
            ReDim veri(1 To 1, 1 To 21): veri(1, 6) = "Kayýt yok.": n = 1
        Else
            ReDim veri(1 To n, 1 To 21)
            i = 0
            For Each x In an
                If sayfa = 2 Or x(14) > 0.0005 Then
                    i = i + 1
                    veri(i, 1) = x(19): veri(i, 2) = x(1): veri(i, 3) = x(2): veri(i, 4) = x(3): veri(i, 5) = x(4): veri(i, 6) = x(5)
                    veri(i, 7) = x(6): veri(i, 8) = Round(x(7) + x(8), 3): veri(i, 9) = Round(x(20), 3): veri(i, 10) = Round(x(9), 2)
                    veri(i, 11) = AN_SURE_METNI(x(10)): veri(i, 12) = Round(x(11), 3): veri(i, 13) = Round(x(12), 3)
                    veri(i, 14) = x(17): veri(i, 15) = Round(x(18), 3): veri(i, 16) = x(14): veri(i, 17) = Round(x(15), 2)
                    veri(i, 18) = x(16)
                    If x(21) > 0 Then veri(i, 19) = Format$(CDate(x(21)), "dd.mm.yyyy")
                    If x(22) >= 0 Then veri(i, 20) = CStr(x(22)) Else veri(i, 20) = "-"
                    veri(i, 21) = AN_SURE_METNI(x(13))
                End If
            Next x
        End If
        ust = "SATIN ALMA  |  " & IIf(sayfa = 1, "Satýn alýnacaklar (kritik ve azalan)", "Stok kapsamý (son 6 ayda sarfý olan tüm malzemeler)")
        AKSORA_EXCEL_SAYFA c, Array("Durum", "Ýl", "Sistem", "ÜY", "Malzeme", "Kýsa metin", "Birim", "Stok (D/K + yük.)", _
            "Son 6 ay sarf", "Aylýk ortalama sarf", "Ne kadar yeter", "Açýk talep", "Açýk sipariþ", "Transfer (il: miktar)", "Transfer miktarý", _
            "Alýnacak miktar", "Stok birim deðeri (TL)", "Tahmini tutar (TL)", "Son sarf tarihi", "Sarf olan ay sayýsý", "Yoldakilerle ne kadar yeter"), _
            veri, n, Array(14, 11, 7, 7, 13, 40, 7, 12, 11, 11, 10, 10, 10, 28, 11, 12, 15, 17, 12, 10, 11), _
            Array("@", "@", "@", "@", "0", "@", "@", "General", "General", "General", "@", "General", "General", "@", "General", _
                "General", "#,##0.00 ""TL""", "#,##0.00 ""TL""", "@", "@", "@"), _
            Array(xlCenter, xlLeft, xlCenter, xlCenter, xlCenter, xlLeft, xlCenter, xlRight, xlRight, xlRight, xlCenter, xlRight, xlRight, _
                xlLeft, xlRight, xlRight, xlRight, xlRight, xlCenter, xlCenter, xlCenter), ust, _
            "Ne kadar yeter = stok / aylýk ortalama sarf (1 ay = 30 gün)  |  Alýnacak = 3 aylýk ihtiyaç - stok - yoldakiler - transfer (atýl veya fazla stok)", _
            IIf(sayfa = 1, "Satýn Alýnacaklar", "Stok Kapsamý"), Empty
        For Each j2 In Array(8, 9, 10, 12, 13, 15, 16)
            AKSORA_V13_MIKTAR_BICIMI c.Range(c.Cells(5, j2), c.Cells(n + 4, j2))
        Next j2
        For i = 1 To n
            Set hz = c.Cells(i + 4, 1)
            Select Case CStr(hz.Value2)
                Case "KRÝTÝK": hz.Interior.Color = RGB(220, 38, 38): hz.Font.Color = vbWhite: hz.Font.bold = True
                Case "AZALIYOR": hz.Interior.Color = RGB(253, 186, 116): hz.Font.Color = RGB(124, 45, 18): hz.Font.bold = True
                Case "YOLDA": hz.Interior.Color = RGB(219, 234, 254): hz.Font.Color = RGB(30, 64, 175)
                Case "DÜZENSÝZ SARF": hz.Interior.Color = RGB(226, 232, 240): hz.Font.Color = RGB(51, 65, 85)
                Case "FAZLA STOK": hz.Interior.Color = RGB(237, 233, 254): hz.Font.Color = RGB(91, 33, 182)
            End Select
        Next i
    Next sayfa
End Sub

Public Sub AKSORA_KRITIK_EXCEL()
    AKSORA_SATINALMA_EXCEL
End Sub

'---------- Analiz onbellegi ve veri yasi ----------
Public Sub AN_ONB_AC()
    If mOnbDerinlik <= 0 Or Not AN_ONB_GECERLI() Then
        mOnbDerinlik = 0
        Set mOnbTum = Nothing: Set mOnbHes = Nothing: mOnbBilgi = Empty
        mOnbZaman = Timer
    End If
    mOnbDerinlik = mOnbDerinlik + 1
End Sub

Public Sub AN_ONB_KAPAT()
    mOnbDerinlik = mOnbDerinlik - 1
    If mOnbDerinlik <= 0 Then
        mOnbDerinlik = 0
        Set mOnbTum = Nothing: Set mOnbHes = Nothing: mOnbBilgi = Empty
    End If
End Sub

Private Function AN_ONB_GECERLI() As Boolean
    Dim fark As Double
    If mOnbDerinlik <= 0 Then Exit Function
    fark = Timer - mOnbZaman
    If fark < 0 Then fark = fark + 86400
    AN_ONB_GECERLI = (fark < 120)
End Function

' Veri yasi: tarih AN_ESKI_GUN gunden eskiyse sayfanin bilgi satiri (A9) kirmizi uyariyla baslar.
Private Sub AN_YAS_YAZ(ByVal ws As Worksheet, ByVal tarih As Double, ByVal ne As String)
    Dim g As Long
    If tarih <= 0 Then Exit Sub
    g = Int(CDbl(Now)) - Int(tarih)
    If g < AN_ESKI_GUN Then Exit Sub
    ws.Range("A9").Value2 = ChrW(9888) & "  " & ne & " " & g & " GÜN ÖNCE OKUNDU – SAP'DEN YENÝDEN OKUYUN   |   " & CStr(ws.Range("A9").Value2)
    With ws.Range("A9").Font
        .Italic = False: .bold = True: .Color = RGB(185, 28, 28)
    End With
End Sub

Private Function AN_EN_ESKI(ByVal a As Double, ByVal b As Double) As Double
    If a <= 0 Then AN_EN_ESKI = b: Exit Function
    If b <= 0 Then AN_EN_ESKI = a: Exit Function
    If a < b Then AN_EN_ESKI = a Else AN_EN_ESKI = b
End Function

' Sayi -> nokta ondalikli metin (yerel ayardan bagimsiz; AN_SAYI_OKU ile geri okunur).
Private Function AN_SAYI_YAZ(ByVal x As Double) As String
    AN_SAYI_YAZ = Replace(CStr(Round(x, 4)), ",", ".")
End Function

Private Function AN_SAYI_OKU(ByVal t As String) As Double
    AN_SAYI_OKU = Val(Trim$(t))
End Function

' "IL=miktar;F:IL=miktar" -> Collection of Array(il kodu, miktar, kaynak: "A" atil / "F" fazla stok)
Private Function AN_KAYNAKLAR(ByVal t As String) As Collection
    Dim p As Variant, q() As String
    Set AN_KAYNAKLAR = New Collection
    If Len(Trim$(t)) = 0 Then Exit Function
    For Each p In Split(t, ";")
        q = Split(CStr(p), "=")
        If UBound(q) = 1 Then
            If AN_SAYI_OKU(q(1)) > 0.0005 Then
                If Left$(q(0), 2) = "F:" Then
                    AN_KAYNAKLAR.Add Array(Mid$(q(0), 3), AN_SAYI_OKU(q(1)), "F")
                Else
                    AN_KAYNAKLAR.Add Array(q(0), AN_SAYI_OKU(q(1)), "A")
                End If
            End If
        End If
    Next p
End Function

'==================== TRANSFER LISTESI (Malzeme Stok birimi) ====================
' Tek soru: hangi malzeme, nereden, nereye, ne kadar, ne kazandiririz. Kaynak: SATINALMA KONTROLU (acik talep / siparis
' atil stoktan karsilanir) ve KRITIK STOK (azalan malzeme diger ilin atil stogundan). Atil stok paylastirmali; ayni stok iki kez
' onerilmez (kritik stok satinalmadan kalan atil stogu kullanir). Kazanc = miktar x AKSORA stok birim degeri (satin alinmayan tutar).
' Donus: kazanc buyukten kucuge Collection of Array(0 nereden kod, 1 nereye kod, 2 sistem, 3 malzeme, 4 tanim, 5 miktar, 6 birim,
'   7 kazanc TL, 8 ne yapilacak, 9 gerekce, 10 tur: "TRANSFER" / "KENDÝ ATIL STOÐU")
Private Function AN_TRANSFERLER(ByRef saTarih As Date, ByRef okuma As Variant) As Collection
    Dim d As Object, saKayit As Collection, x As Variant, kk As Variant, hedef As String, k As String, o As Variant, sBilgi As String
    Dim an As Collection, anah() As String, idx() As Long, kod() As String, i As Long, n As Long, gerekce As String, tur As String
    Set AN_TRANSFERLER = New Collection
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    Set saKayit = SA_KAYIT_OKU(saTarih, sBilgi)
    If saKayit.Count > 0 Then
        For Each x In SA_ANALIZ(saKayit)
            If UBound(x) >= 24 Then
                hedef = ATIL_TESIS_IL(CStr(x(3)))
                For Each kk In AN_KAYNAKLAR(CStr(x(24)))
                    gerekce = IIf(x(0) = "T", "Talep ", "Sipariþ ") & CStr(x(5)) & IIf(Len(CStr(x(6))) > 0, "/" & CStr(x(6)), "")
                    AN_TR_EKLE d, CStr(kk(0)), hedef, CStr(x(1)), Y_KOD_NORM(x(9)), CStr(x(10)), kk(1), CStr(x(12)), kk(1) * x(18), gerekce, (x(0) = "T"), False
                Next kk
            End If
        Next x
    End If
    Set an = KS_ANALIZ(okuma, saTarih)
    For Each x In an
        If (x(19) = "KRÝTÝK" Or x(19) = "AZALIYOR") And Len(CStr(x(23))) > 0 Then
            For Each kk In AN_KAYNAKLAR(CStr(x(23)))
                ' Alan ilin stogu bugunku tuketimle kac ay yeter (stok / son 6 ay ortalama aylik sarf).
                If x(7) + x(8) <= 0.0005 Then
                    gerekce = Y_IL_ADI(CStr(x(0))) & "'da stok yok; ayda ort. " & ADM_ADET(Round(x(9), 1)) & " " & CStr(x(6)) & " kullanýlýyor"
                Else
                    gerekce = Y_IL_ADI(CStr(x(0))) & " stoðu " & AN_SURE_METNI(x(10)) & " yeter (stok " & _
                        ADM_ADET(Round(x(7) + x(8), 3)) & ", ayda ort. " & ADM_ADET(Round(x(9), 1)) & " " & CStr(x(6)) & " kullanýlýyor)"
                End If
                AN_TR_EKLE d, CStr(kk(0)), CStr(x(0)), CStr(x(2)), CStr(x(4)), CStr(x(5)), kk(1), CStr(x(6)), kk(1) * x(15), gerekce, False, (kk(2) = "F")
            Next kk
        End If
    Next x
    n = d.Count
    If n = 0 Then Exit Function
    ReDim anah(1 To n): ReDim idx(1 To n): ReDim kod(1 To n)
    For Each kk In d.keys
        i = i + 1: idx(i) = i: kod(i) = CStr(kk)
        anah(i) = I_SIRA_ANAHTAR(d(kk)(7))
    Next kk
    Y_SIRALA anah, idx, 1, n
    For i = 1 To n
        o = d(kod(idx(i)))
        If o(0) = o(1) Then
            o(10) = "KENDÝ ATIL STOÐU"
            o(8) = "Depodaki atýl stoðu kullan" & IIf(o(11), "; talebi iptal et / azalt", "; sipariþi azalt")
        Else
            o(10) = "TRANSFER"
            o(8) = Y_IL_ADI(CStr(o(0))) & " " & ChrW(8594) & " " & Y_IL_ADI(CStr(o(1))) & " stok nakli" & IIf(o(13), "  (fazla stoktan)", "  (atýl stoktan)")
        End If
        AN_TRANSFERLER.Add Array(o(0), o(1), o(2), o(3), o(4), Round(o(5), 3), o(6), Round(o(7), 2), o(8), o(9), o(10))
    Next i
End Function

' Ayni rota + malzeme tek satir; gerekceler birlesir (en fazla 3 belge + adet).
Private Sub AN_TR_EKLE(ByVal d As Object, ByVal kaynak As String, ByVal hedef As String, ByVal sis As String, ByVal mal As String, _
    ByVal tanim As String, ByVal miktar As Double, ByVal birim As String, ByVal kazanc As Double, ByVal gerekce As String, ByVal talep As Boolean, _
    ByVal fazla As Boolean)
    Dim k As String, o As Variant
    If miktar <= 0.0005 Or Len(kaynak) = 0 Or Len(hedef) = 0 Then Exit Sub
    k = kaynak & "|" & hedef & "|" & sis & "|" & mal
    If d.Exists(k) Then
        o = d(k)
        o(5) = o(5) + miktar: o(7) = o(7) + kazanc: o(12) = o(12) + 1
        If o(12) <= 3 Then o(9) = o(9) & ", " & gerekce Else If o(12) = 4 Then o(9) = o(9) & " ..."
        If talep Then o(11) = True
        If fazla Then o(13) = True
    Else
        o = Array(kaynak, hedef, sis, mal, tanim, miktar, birim, kazanc, vbNullString, gerekce, vbNullString, talep, 1, fazla)
    End If
    d(k) = o
End Sub

Public Function AKSORA_TRANSFER_SAYFA(ByVal sh As Object) As Boolean
    If Not TypeOf sh Is Worksheet Then Exit Function
    If sh.Name <> TR_SAYFA Then Exit Function
    AKSORA_TRANSFER_SAYFA = (CStr(sh.Range("BX1").Value2) = TR_ETIKET)
End Function

Public Sub AKSORA_TRANSFER_AC()
    Dim ac As Boolean, no As Long, metin As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    ac = AKSORA_BEKLE_AC("AKSORA  |  TRANSFER LÝSTESÝ", "Transfer listesi hazýrlanýyor...")
    On Error GoTo Son
    TR_GOSTER vbNullString
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Transfer listesi açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Public Sub AKSORA_TRANSFER_IL_SEC()
    Dim ac As Boolean, no As Long, metin As String, il As String
    If AKSORA_V11_MESGUL Or AKSORA_V13_GECIS Or AKSORA_V13_ACILIS Or AKSORA_ATIL_CALISIYOR Then Exit Sub
    On Error Resume Next
    il = ThisWorkbook.Worksheets(TR_SAYFA).Shapes(CStr(application.Caller)).AlternativeText
    On Error GoTo 0
    If Len(il) = 0 Then il = "TUM"
    ac = AKSORA_BEKLE_AC("AKSORA  |  TRANSFER LÝSTESÝ", "Transfer listesi hazýrlanýyor...")
    On Error GoTo Son
    TR_GOSTER il
Son:
    no = Err.Number: metin = Err.Description
    If ac Then AKSORA_BEKLE_KAPAT
    If no <> 0 Then MsgBox "Transfer listesi açýlamadý: " & metin, vbExclamation, "AKSORA"
End Sub

Private Sub TR_GOSTER(ByVal il As String)
    Dim ws As Worksheet, ev As Boolean, hesap As XlCalculation, no As Long, metin As String
    ev = application.EnableEvents: hesap = application.Calculation
    AN_ONB_AC
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    application.Calculation = xlCalculationManual
    Set ws = AN_SAYFA_KUR(TR_SAYFA, TR_ETIKET, SA_SAYFA)
    If Len(il) = 0 Then il = CStr(ws.Range("BX20").Value2)
    If Len(il) = 0 Then il = "TUM"
    TR_CIZ ws, il
    ws.Visible = xlSheetVisible
    mYonlendirme = True
    ws.Select Replace:=True
    mYonlendirme = False
    AKSORA_YATIRIM_GORUNUM ws
Cikis:
    On Error Resume Next
    AN_ONB_KAPAT
    mYonlendirme = False
    application.Calculation = hesap
    application.EnableEvents = ev: application.ScreenUpdating = True
    On Error GoTo 0
    If no <> 0 Then Err.Raise no, "TR_GOSTER", metin
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not ws Is Nothing Then ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    Resume Cikis
End Sub

' Analiz sayfasi kurulumu (rapor sablonu M_CDP_RAPOR kopyasi); sonrakiAd sayfasindan sonra konur.
Private Function AN_SAYFA_KUR(ByVal ad As String, ByVal etiket As String, ByVal sonrakiAd As String) As Worksheet
    Dim ws As Worksheet, kaynak As Worksheet, gor As XlSheetVisibility, onceki As Object, sh As Object, da As Boolean
    Set ws = ADM_WS(ad)
    If Not ws Is Nothing Then
        If CStr(ws.Range("BX1").Value2) <> etiket Then Err.Raise vbObjectError + 2390, , "Ayný isimde baþka bir sayfa var: " & ad
        Set AN_SAYFA_KUR = ws
        Exit Function
    End If
    Set kaynak = ADM_WS("M_CDP_RAPOR")
    If kaynak Is Nothing Then Err.Raise vbObjectError + 2391, , "Sayfa þablonu (M_CDP_RAPOR) bulunamadý."
    da = application.DisplayAlerts: application.DisplayAlerts = False
    gor = kaynak.Visible
    kaynak.Visible = xlSheetVisible
    Set onceki = CreateObject("Scripting.Dictionary"): onceki.CompareMode = vbTextCompare
    For Each sh In ThisWorkbook.Sheets
        onceki(sh.Name) = True
    Next sh
    kaynak.Copy after:=kaynak
    For Each sh In ThisWorkbook.Sheets
        If Not onceki.Exists(sh.Name) Then Set ws = sh: Exit For
    Next sh
    kaynak.Visible = gor
    application.DisplayAlerts = da
    If ws Is Nothing Then Err.Raise vbObjectError + 2392, , "Kopya sayfa bulunamadý."
    ws.Name = ad
    ws.Unprotect ""
    ws.Range("BX1:BX40").ClearContents
    ws.Range("BX1").Value2 = etiket
    On Error Resume Next
    If Not ADM_WS(sonrakiAd) Is Nothing Then ws.Move after:=ThisWorkbook.Worksheets(sonrakiAd) Else ws.Move after:=ThisWorkbook.Worksheets(ATIL_SAYFA)
    On Error GoTo 0
    Set AN_SAYFA_KUR = ws
End Function

Private Sub TR_CIZ(ByVal ws As Worksheet, ByVal ilSec As String)
    Dim liste As Collection, saTarih As Date, okuma As Variant, x As Variant, w As Double, s As Shape, y As Double, kw As Double, bw As Double
    Dim i As Long, r As Long, a1 As Long, n As Long, il As Variant, sw As Double, sonuc() As Variant, top As Double, nT As Long, nK As Long
    Dim enIyi As String, sec As Boolean
    ws.Unprotect ""
    Y_SAYFA_SIFIRLA ws
    For i = ws.Shapes.Count To 1 Step -1
        If InStr(ws.Shapes(i).Name, "FILIGRAN") > 0 Then ws.Shapes(i).Delete
    Next i
    Set s = Y_SEKIL(ws, "btnAnaPanelDon")
    If Not s Is Nothing Then s.AlternativeText = vbNullString: s.OnAction = BIRIM_MAKRO("ANA_PANELE_DON")
    Y_YAZ Y_SEKIL(ws, "AKT1_RAPOR_BASLIK"), "AKSORA | TRANSFER LÝSTESÝ  |  " & IIf(ilSec = "TUM", "Tüm Ýller", Y_IL_ADI(ilSec))
    Set s = Y_SEKIL(ws, "AKT1_RAPOR_UST_ZEMIN")
    If s Is Nothing Then w = 1544 Else w = s.Width
    sw = application.Max(160, w - 1190)
    Y_KOLONLAR ws, Array(105, 105, 110, 330, 100, 65, 145, 230, sw)
    ws.Range("BX20").Value2 = ilSec
    Set liste = AN_TRANSFERLER(saTarih, okuma)
    For Each x In liste
        If ilSec = "TUM" Or x(0) = ilSec Or x(1) = ilSec Then
            n = n + 1: top = top + x(7)
            If x(10) = "TRANSFER" Then nT = nT + 1 Else nK = nK + 1
            If Len(enIyi) = 0 Then enIyi = Left$(CStr(x(4)), 28) & "  •  " & AKSORA_SAYI_METNI(x(7), "#,##0") & " TL"
        End If
    Next x
    If Len(enIyi) = 0 Then enIyi = "-"
    ws.Range("BZ21").Value2 = Round(top, 2): ws.Range("BZ21").NumberFormat = "#,##0.00 ""TL"""
    ws.Range("BZ22").Value2 = nT: ws.Range("BZ23").Value2 = nK: ws.Range("BZ22:BZ23").NumberFormat = "#,##0"
    ws.Range("BZ24").Value2 = enIyi
    ws.Range(ws.Rows(3), ws.Rows(6)).RowHeight = 16.95
    ws.Rows(7).RowHeight = 10: ws.Rows(8).RowHeight = 30: ws.Rows(9).RowHeight = 18
    y = ws.Rows(3).Top + 1
    kw = (w - 3 * 12) / 4
    Y_KPI ws, 1, 0, y, kw, 63.7, "TOPLAM KAZANÇ (SATIN ALINMAYAN)", "$BZ$21", RGB(39, 160, 94)
    Y_KPI ws, 2, kw + 12, y, kw, 63.7, "ÝLLER ARASI TRANSFER", "$BZ$22", RGB(35, 119, 204)
    Y_KPI ws, 3, 2 * (kw + 12), y, kw, 63.7, "KENDÝ ATIL STOÐUNDAN", "$BZ$23", RGB(220, 38, 38)
    Y_KPI ws, 4, 3 * (kw + 12), y, kw, 63.7, "EN ÇOK KAZANDIRAN", "$BZ$24", RGB(12, 45, 79)
    Set s = Y_SEKIL(ws, "Y_KPI_DEGER_4")
    If Not s Is Nothing Then s.TextFrame2.TextRange.Font.Size = 13
    y = ws.Rows(8).Top + 1
    bw = application.Min(118, (w - 150 - 7 * 6) / 6)
    i = 0
    For Each il In Array("TUM", "ARTVIN", "TRABZON", "GUMUSHANE", "GIRESUN", "RIZE")
        Set s = Y_DUGME(ws, "TR_IL_" & i, IIf(CStr(il) = "TUM", "TÜM ÝLLER", ATIL_IL_BUYUK(CStr(il))), i * (bw + 6), y, bw, 28, _
            IIf(CStr(il) = ilSec, RGB(0, 125, 110), RGB(148, 163, 184)), "AKSORA_TRANSFER_IL_SEC")
        s.AlternativeText = CStr(il)
        i = i + 1
    Next il
    Y_DUGME ws, "TR_EXCEL", "EXCEL'E AKTAR", w - 150, y, 150, 28, RGB(33, 115, 70), "AKSORA_TRANSFER_EXCEL"
    ws.Range("A9").Value2 = "Ýl seçilince o ilin GÖNDERECEÐÝ ve ALACAÐI malzemeler  •  Kazanç = miktar x AKSORA stok birim deðeri (satýn alýnmayan tutar)  •  " & _
        "Kaynak: açýk talep / sipariþ (SAP " & IIf(saTarih > 0, Format$(saTarih, "dd.mm.yyyy"), "okunmadý") & ") ve stoðu 3 aydan az yeten malzemeler"
    With ws.Range("A9").Font
        .Name = "Aptos": .Size = 9: .Italic = True: .Color = RGB(100, 116, 139): .bold = False
    End With
    If IsArray(okuma) Then AN_YAS_YAZ ws, AN_EN_ESKI(CDbl(okuma(0)), CDbl(okuma(1))), "ATIL STOK (MB51)"
    If saTarih > 0 Then AN_YAS_YAZ ws, CDbl(saTarih), "SATINALMA (ME5A / ME2N)"
    r = 11
    Y_BOLUM ws, r, "YAPILACAKLAR  •  " & n & " kalem  •  " & AKSORA_SAYI_METNI(Round(top, 2), "#,##0.00") & " TL kazanç", _
        "En çok kazandýran üstte. Baþlýktaki filtrelerle gönderen / alan il seçilebilir.", 9
    ws.Rows(r).RowHeight = 30
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 9)), Array("Nereden", "Nereye", "Malzeme", "Kýsa metin", "Miktar", "Birim", _
        "Kazanç", "Ne yapýlacak", "Gerekçe")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    If n > 0 Then
        ReDim sonuc(1 To n, 1 To 9)
        i = 0
        For Each x In liste
            If ilSec = "TUM" Or x(0) = ilSec Or x(1) = ilSec Then
                i = i + 1
                sonuc(i, 1) = Y_IL_ADI(CStr(x(0))) & " " & x(2): sonuc(i, 2) = Y_IL_ADI(CStr(x(1))) & " " & x(2): sonuc(i, 3) = x(3): sonuc(i, 4) = x(4)
                sonuc(i, 5) = x(5): sonuc(i, 6) = x(6): sonuc(i, 7) = x(7): sonuc(i, 8) = x(8): sonuc(i, 9) = x(9)
            End If
        Next x
        ws.Range(ws.Cells(a1, 3), ws.Cells(a1 + n - 1, 4)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 6), ws.Cells(a1 + n - 1, 6)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 8), ws.Cells(a1 + n - 1, 9)).NumberFormat = "@"
        ws.Range(ws.Cells(a1, 1), ws.Cells(a1 + n - 1, 9)).Value2 = sonuc
        r = a1 + n - 1
        Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 9))
        ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 22
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(a1, 1), ws.Cells(r, 2)).Font.bold = True
        ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).IndentLevel = 1
        ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).HorizontalAlignment = xlCenter
        AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5))
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).Font.bold = True
        ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).Font.Size = 11
        ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).NumberFormat = "#,##0.00 ""TL"""
        ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).Font.bold = True
        ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).Font.Color = RGB(21, 128, 61)
        ws.Range(ws.Cells(a1, 8), ws.Cells(r, 9)).ShrinkToFit = True
        ws.Range(ws.Cells(a1, 8), ws.Cells(r, 8)).Font.bold = True
        For i = 1 To n
            If sonuc(i, 1) = sonuc(i, 2) Then
                ws.Cells(a1 + i - 1, 8).Font.Color = RGB(185, 28, 28)
            Else
                ws.Cells(a1 + i - 1, 8).Font.Color = RGB(30, 64, 175)
            End If
        Next i
        ws.Range(ws.Cells(a1 - 1, 1), ws.Cells(r, 9)).AutoFilter
    Else
        r = a1
        ws.Cells(r, 4).Value2 = "Önerilecek transfer yok (SATINALMA KONTROL ve ATIL STOK okumalarýný güncel tutun)."
        ws.Cells(r, 4).Font.Italic = True: ws.Cells(r, 4).Font.Color = RGB(100, 116, 139)
    End If
    r = r + 2
    ws.Cells(r, 1).Value2 = "Mavi: iller arasý stok nakli.  Kýrmýzý: ayný ilde atýl stok var, talep / sipariþ gereksiz.  Atýl stok: 24 aydýr giriþi olmayan stok.  " & _
        "Fazla stok: gönderen ilin 12 aylýk ihtiyacýný aþan kýsmý (gönderen il sýkýntýya düþmez).  Ayný stok iki kez önerilmez."
    ws.Cells(r, 1).Font.Size = 9: ws.Cells(r, 1).Font.Color = RGB(71, 85, 105)
    Y_FILIGRAN ws
    ws.Protect Password:="", DrawingObjects:=True, Contents:=True, UserInterfaceOnly:=True, AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

' Excel: gonderen il bazinda sayfalar (her il kendi sayfasini uygular) + tum liste.
Public Sub AKSORA_TRANSFER_EXCEL()
    Dim liste As Collection, saTarih As Date, okuma As Variant, wb As Workbook, c As Worksheet, ac As Boolean, ev As Boolean
    Dim veri() As Variant, x As Variant, n As Long, i As Long, j As Long, il As Variant, sonuc As String, no As Long, metin As String
    Dim sayfalar As Collection, top As Double, ust As String
    AKSORA_DONUS_KAYDET True
    ac = AKSORA_BEKLE_AC("AKSORA  |  TRANSFER LÝSTESÝ", "Transfer listesi hazýrlanýyor...", 0.1)
    ev = application.EnableEvents
    AN_ONB_AC
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    Set liste = AN_TRANSFERLER(saTarih, okuma)
    If liste.Count = 0 Then
        AN_ONB_KAPAT
        application.EnableEvents = ev: application.ScreenUpdating = True
        If ac Then AKSORA_BEKLE_KAPAT
        MsgBox "Önerilecek transfer yok.", vbInformation, "AKSORA"
        Exit Sub
    End If
    Set sayfalar = New Collection
    sayfalar.Add "TUM"
    For Each il In ADM_ILLER()
        For Each x In liste
            If x(0) = CStr(il) Then sayfalar.Add CStr(il): Exit For
        Next x
    Next il
    AKSORA_DONUS_KAYDET
    Set wb = Workbooks.Add(xlWBATWorksheet)
    For i = 2 To sayfalar.Count: wb.Worksheets.Add After:=wb.Worksheets(wb.Worksheets.Count): Next i
    For i = 1 To sayfalar.Count
        il = sayfalar(i)
        n = 0: top = 0
        For Each x In liste
            If il = "TUM" Or x(0) = il Then n = n + 1: top = top + x(7)
        Next x
        ReDim veri(1 To n, 1 To 10)
        j = 0
        For Each x In liste
            If il = "TUM" Or x(0) = il Then
                j = j + 1
                veri(j, 1) = Y_IL_ADI(CStr(x(0))): veri(j, 2) = Y_IL_ADI(CStr(x(1))): veri(j, 3) = x(2): veri(j, 4) = x(3)
                veri(j, 5) = x(4): veri(j, 6) = x(5): veri(j, 7) = x(6): veri(j, 8) = x(7): veri(j, 9) = x(8): veri(j, 10) = x(9)
            End If
        Next x
        Set c = wb.Worksheets(i)
        If il = "TUM" Then
            ust = "TRANSFER LÝSTESÝ  |  Tüm iller  |  " & AKSORA_SAYI_METNI(Round(top, 2), "#,##0.00") & " TL kazanç"
        Else
            ust = "TRANSFER LÝSTESÝ  |  " & Y_IL_ADI(CStr(il)) & " gönderecek  |  " & AKSORA_SAYI_METNI(Round(top, 2), "#,##0.00") & " TL kazanç"
        End If
        AKSORA_EXCEL_SAYFA c, Array("Nereden", "Nereye", "Sistem", "Malzeme", "Kýsa metin", "Miktar", "Birim", "Kazanç (TL)", _
            "Ne yapýlacak", "Gerekçe"), veri, n, Array(12, 12, 8, 13, 42, 11, 7, 17, 34, 40), _
            Array("@", "@", "@", "0", "@", "General", "@", "#,##0.00 ""TL""", "@", "@"), _
            Array(xlCenter, xlCenter, xlCenter, xlCenter, xlLeft, xlRight, xlCenter, xlRight, xlLeft, xlLeft), ust, _
            "Kazanç = miktar x AKSORA stok birim deðeri  |  Ayný atýl stok iki kez önerilmez", IIf(il = "TUM", "Tüm Liste", Y_IL_ADI(CStr(il)) & " Gönderecek"), Empty
        AKSORA_V13_MIKTAR_BICIMI c.Range(c.Cells(5, 6), c.Cells(n + 4, 6))
        c.Range(c.Cells(5, 8), c.Cells(n + 4, 8)).Font.bold = True
        c.Range(c.Cells(5, 8), c.Cells(n + 4, 8)).Font.Color = RGB(21, 128, 61)
    Next i
    wb.Worksheets(1).Activate
    sonuc = AKSORA_EXCEL_BITIR(wb, AKSORA_MASAUSTU() & application.PathSeparator & _
        AKSORA_DOSYA_ADI("Transfer Listesi - " & Format$(Date, "dd.mm.yyyy")) & ".xlsx", False)
    AN_ONB_KAPAT
    application.EnableEvents = ev: application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    If Len(sonuc) > 0 Then application.StatusBar = "Masaüstüne kaydedildi: " & sonuc
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    AN_ONB_KAPAT
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    ThisWorkbook.Activate
    application.EnableEvents = ev: application.ScreenUpdating = True
    If ac Then AKSORA_BEKLE_KAPAT
    On Error GoTo 0
    MsgBox "Transfer listesi Excel'e aktarýlamadý." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

'---------- Yonetim Merkezi: analiz uyarilari ve kritik stok ozeti ----------
' Skor hesabindan once cagrilir (ADM_KUR). Veri yasi, fazla teslim, kritik stok ve bekleyen transfer kazanci.
Public Sub ADM_ANALIZ_UYARI(ByVal uy As Collection, ByVal secIl As String)
    Dim nIp As Long, tIp As Double, nStokYok As Long, enAcil As String
    Dim ws As Worksheet, t0 As Double, t1 As Double, g As Long, saTarih As Date, sBilgi As String, kayit As Collection, x As Variant
    Dim nF As Long, tF As Double, an As Collection, okuma As Variant, nK As Long, tK As Double, liste As Collection, tT As Double, nT As Long
    Dim ilAd As String
    Set mAdmIsler = New Collection
    On Error GoTo Son
    ilAd = IIf(Len(secIl) = 0, "Tüm iller", Y_IL_ADI(secIl))
    Set ws = ADM_WS(ATIL_KAYIT)
    If Not ws Is Nothing Then
        t0 = AN_EN_ESKI(ATIL_HUCRE_SAYI(ws.Range("D1").Value2), ATIL_HUCRE_SAYI(ws.Range("E1").Value2))
        If t0 > 0 Then
            g = Int(CDbl(Now)) - Int(t0)
            If g >= AN_ESKI_GUN Then
                ADM_UYAR uy, 2, "Malzeme • ATIL STOK SAP okumasý (MB51) " & g & " gün önce; atýl stok, transfer ve kritik stok eski veriyle hesaplanýyor.", "Malzeme | Analiz", "Analiz verisi"
                mAdmIsler.Add Array(60, "ATIL STOK'ta SAP'den oku (MB51): veri " & g & " gün eski, öneriler güncel deðil", 0#, "Malzeme > ATIL STOK")
            End If
        End If
    End If
    Set kayit = SA_KAYIT_OKU(saTarih, sBilgi)
    If kayit.Count > 0 Then
        g = Int(CDbl(Now)) - Int(CDbl(saTarih))
        If g >= AN_ESKI_GUN Then ADM_UYAR uy, 3, "Malzeme • SATINALMA KONTROL (ME5A / ME2N) " & g & " gün önce okundu.", "Malzeme | Analiz", "Analiz verisi"
        For Each x In SA_ANALIZ(kayit)
            If Len(secIl) = 0 Or ATIL_TESIS_IL(CStr(x(3))) = secIl Then
                If x(22) = "FAZLA TESLÝM ALINMIÞ" Then nF = nF + 1: tF = tF + Abs(x(20))
                If x(17) > 0.0005 Then nIp = nIp + 1: tIp = tIp + x(19)
            End If
        Next x
        If nIp > 0 Then mAdmIsler.Add Array(30, nIp & " açýk talep / sipariþi iptal et ya da azalt: ayný malzeme atýl stokta var", tIp, "Malzeme > SATIN ALMA (tablo 2)")
        If nF > 0 Then mAdmIsler.Add Array(40, nF & " fazla teslim alýnmýþ sipariþ kalemi için iade / fatura kontrolü yap", tF, "Malzeme > SATIN ALMA (tablo 3)")
        If nF > 0 Then ADM_UYAR uy, 2, "Satýnalma • " & ilAd & ": " & nF & " sipariþ kalemi sipariþten FAZLA teslim alýnmýþ; iade / fatura kontrolü yapýlmalý.", "Malzeme | Satýnalma", "Fazla teslim"
    End If
    Set an = KS_ANALIZ(okuma, saTarih)
    For Each x In an
        If (Len(secIl) = 0 Or x(0) = secIl) Then
            If x(19) = "KRÝTÝK" Then nK = nK + 1
            If x(14) > 0.0005 Then tK = tK + x(16)
            If x(19) = "KRÝTÝK" And x(7) + x(8) <= 0.0005 And x(14) > 0.0005 Then
                nStokYok = nStokYok + 1
                If Len(enAcil) = 0 Then enAcil = CStr(x(5)) & " (" & x(1) & ")"
            End If
        End If
    Next x
    If nK > 0 Then mAdmIsler.Add Array(10, "Satýn almayý baþlat: " & nK & " malzemenin stoðu 1 aydan az yetiyor" & IIf(nStokYok > 0, ", " & nStokYok & " tanesinin stoðu yok (ör. " & enAcil & ")", ""), tK, "Malzeme > SATIN ALMA (tablo 1)")
    If nK > 0 Then ADM_UYAR uy, IIf(nK >= 50, 1, 2), "Kritik stok • " & ilAd & ": " & nK & " malzemenin stoðu 1 aydan az yetiyor; satýn alýnmasý gereken " & ADM_TL(tK) & ".", "Malzeme | Kritik stok", "Kritik stok"
    Set liste = AN_TRANSFERLER(saTarih, okuma)
    For Each x In liste
        If Len(secIl) = 0 Or x(0) = secIl Or x(1) = secIl Then nT = nT + 1: tT = tT + x(7)
    Next x
    If nT > 0 Then mAdmIsler.Add Array(20, nT & " stok naklini / atýl kullanýmýný uygula (satýn alma yerine)", tT, "Malzeme > TRANSFER LÝSTESÝ")
    If nT > 0 Then ADM_UYAR uy, 3, "Transfer • " & ilAd & ": " & nT & " transfer / atýl kullaným önerisi bekliyor; uygulanýrsa " & ADM_TL(tT) & " satýn alma yapýlmaz (TRANSFER LÝSTESÝ).", "Malzeme | Transfer", "Transfer"
Son:
End Sub

' Birlestirilen eski analiz sayfasi (etiketliyse) silinir; ayni adli kullanici sayfasina dokunulmaz.
Private Sub AN_ESKI_SAYFA_SIL(ByVal ad As String, ByVal etiket As String)
    Dim ws As Worksheet, da As Boolean
    Set ws = ADM_WS(ad)
    If ws Is Nothing Then Exit Sub
    If CStr(ws.Range("BX1").Value2) <> etiket Then Exit Sub
    da = application.DisplayAlerts: application.DisplayAlerts = False
    On Error Resume Next
    ws.Visible = xlSheetVisible
    ws.Delete
    application.DisplayAlerts = da
    On Error GoTo 0
End Sub

' Stok suresi okunur metin: 1,9 ay -> "1 ay 27 gün"; 0,4 ay -> "12 gün"; 0 -> "Stok yok"; 24 ay ustu -> "2 yýl 3 ay".
Private Function AN_SURE_METNI(ByVal ay As Double) As String
    Dim g As Long, a As Long, d As Long, y As Long
    If ay <= 0 Then AN_SURE_METNI = "Stok yok": Exit Function
    g = CLng(Round(ay * 30, 0))
    If g < 1 Then AN_SURE_METNI = "1 günden az": Exit Function
    If g < 30 Then AN_SURE_METNI = g & " gün": Exit Function
    a = g \ 30: d = g Mod 30
    If a >= 24 Then
        y = a \ 12: a = a Mod 12
        AN_SURE_METNI = y & " yýl" & IIf(a > 0, " " & a & " ay", "")
    Else
        AN_SURE_METNI = a & " ay" & IIf(d > 0, " " & d & " gün", "")
    End If
End Function

' ATIL STOK "NE YAPILMALI": her atil / atila yaklasan kalem icin tek eylem satiri.
' Oncelik: 1 kendi ihtiyacinda kullan (acik talep/siparis), 2 transfer et (TRANSFER LISTESI ile ayni hesap), 3 once bunu kullan
' (ilde tuketiliyor: eski stok once), 4 beklet (baska il kullaniyor), 5 sat / hurda / iade degerlendir (hic kullanilmiyor),
' 6 once kullan (3 / 6 ay icinde atil olacak).
Private Function ATIL_ISLEM_HESAPLA(ByVal c As Collection, ByVal ilSec As String) As Collection
    Dim tr As Object, kd As Object, trM As Object, kayit As Object, okuma As Variant, saT As Date, ok2 As Variant, x As Variant
    Dim k As String, q As Double, bd As Double, kal As Double, isler As Collection, eylem As String, ayr As String, il2 As Variant
    Dim uy2 As String, s6 As Double, enIl As String, enS As Double, i As Long, n As Long, r As Long, a1 As Long, anah() As String
    Dim idx() As Long, sonuc() As Variant, o As Variant, top(1 To 6) As Double, adet(1 To 6) As Long, ad As Variant, kk As Variant
    Set tr = CreateObject("Scripting.Dictionary"): tr.CompareMode = vbTextCompare
    Set trM = CreateObject("Scripting.Dictionary"): trM.CompareMode = vbTextCompare
    Set kd = CreateObject("Scripting.Dictionary"): kd.CompareMode = vbTextCompare
    For Each x In AN_TRANSFERLER(saT, ok2)
        If InStr(CStr(x(8)), "fazla stoktan") = 0 Then
            k = AKSORA_IL_TESIS(CStr(x(0)), CStr(x(2))) & "|" & CStr(x(3))
            If x(0) = x(1) Then
                kd(k) = ATIL_DS(kd, k) + x(5)
            Else
                tr(k) = ATIL_DS(tr, k) + x(5)
                trM(k) = IIf(trM.Exists(k), trM(k) & ", ", "") & Y_IL_ADI(CStr(x(1))) & " " & ADM_ADET(x(5))
            End If
        End If
    Next x
    Set kayit = ATIL_KAYIT_OKU(okuma)
    Set isler = New Collection
    For Each x In c
        If x(10) >= 1 And (ilSec = "TUM" Or x(0) = ilSec) Then
            k = CStr(x(2)) & "|" & CStr(x(3))
            If x(6) > 0 Then bd = x(8) / x(6) Else bd = 0
            q = 0
            If IsArray(x(20)) Then q = x(20)(3)
            kal = q
            If kal > 0.0005 And kd.Exists(k) Then
                AN_ISLEM_EKLE isler, x, 1, "KENDÝ ÝHTÝYACINDA KULLAN", application.Min(kal, kd(k)), bd, "Açýk talep / sipariþ var: iptal et, depodaki stoðu kullan"
                kal = kal - application.Min(kal, kd(k))
            End If
            If kal > 0.0005 And tr.Exists(k) Then
                AN_ISLEM_EKLE isler, x, 2, "TRANSFER ET", application.Min(kal, tr(k)), bd, ChrW(8594) & " " & trM(k) & "  (TRANSFER LÝSTESÝ)"
                kal = kal - application.Min(kal, tr(k))
            End If
            If kal > 0.0005 Then
                s6 = 0
                If kayit.Exists(k) Then s6 = kayit(k)(3)
                If s6 > 0.0005 Then
                    AN_ISLEM_EKLE isler, x, 3, "ÖNCE BUNU KULLAN", kal, bd, "Ýlde ayda ort. " & ADM_ADET(Round(s6 / 6, 1)) & " " & x(7) & " kullanýlýyor: yeni stoktan önce bu eski stok çýkýlsýn"
                Else
                    enIl = vbNullString: enS = 0
                    For Each il2 In ADM_ILLER()
                        If CStr(il2) <> CStr(x(0)) Then
                            uy2 = AKSORA_IL_TESIS(CStr(il2), CStr(x(1))) & "|" & CStr(x(3))
                            If kayit.Exists(uy2) Then If kayit(uy2)(3) > enS Then enS = kayit(uy2)(3): enIl = CStr(il2)
                        End If
                    Next il2
                    If enS > 0.0005 Then
                        AN_ISLEM_EKLE isler, x, 4, "BEKLET", kal, bd, Y_IL_ADI(enIl) & " ayda ort. " & ADM_ADET(Round(enS / 6, 1)) & " " & x(7) & " kullanýyor; ihtiyaç doðunca transfer"
                    Else
                        AN_ISLEM_EKLE isler, x, 5, "SAT / HURDA / ÝADE", kal, bd, "Son 6 ayda hiçbir ilde kullanýlmadý"
                    End If
                End If
            End If
            If IsArray(x(20)) Then
                If x(20)(1) + x(20)(2) > 0.0005 Then AN_ISLEM_EKLE isler, x, 6, "ÖNCE KULLAN", x(20)(1) + x(20)(2), bd, "Atýla " & CStr(x(11))
            End If
        End If
    Next x
    Set ATIL_ISLEM_HESAPLA = isler
End Function

' Il bazinda ozet (Yonetim Merkezi atil tablosu): il kodu -> Array(kullanilabilir kalem (kendi ihtiyaci + transfer), sat/hurda kalem, sat/hurda TL)
Public Function ATIL_ISLEM_IL_OZET(ByVal c As Collection) As Object
    Dim d As Object, o As Variant, t As Variant
    Set d = CreateObject("Scripting.Dictionary"): d.CompareMode = vbTextCompare
    For Each o In ATIL_ISLEM_HESAPLA(c, "TUM")
        If d.Exists(CStr(o(13))) Then t = d(CStr(o(13))) Else t = Array(0#, 0#, 0#)
        If o(0) <= 2 Then t(0) = t(0) + 1
        If o(0) = 5 Then t(1) = t(1) + 1: t(2) = t(2) + o(7)
        d(CStr(o(13))) = t
    Next o
    Set ATIL_ISLEM_IL_OZET = d
End Function

Private Sub ATIL_ISLEM_CIZ(ByVal ws As Worksheet, ByVal c As Collection, ByVal ilSec As String)
    Dim isler As Collection, ayr As String, i As Long, n As Long, r As Long, a1 As Long, anah() As String
    Dim idx() As Long, sonuc() As Variant, o As Variant, top(1 To 6) As Double, adet(1 To 6) As Long, ad As Variant, kk As Variant
    Set isler = ATIL_ISLEM_HESAPLA(c, ilSec)
    n = isler.Count
    For Each o In isler
        top(o(0)) = top(o(0)) + o(7): adet(o(0)) = adet(o(0)) + 1
    Next o
    r = 12
    Y_BOLUM ws, r, "NE YAPILMALI  •  " & n & " satýr", "Her satýr tek iþ. Transferler TRANSFER LÝSTESÝ ile ayný hesap (ayný stok iki kez önerilmez).", 12
    ws.Rows(r).RowHeight = 30
    ' eylem ozeti (tek satir)
    r = r + 1
    ayr = vbNullString
    i = 0
    For Each ad In Array("Kendi ihtiyacýnda kullan", "Transfer et", "Önce bunu kullan", "Beklet", "Sat / hurda / iade", "Önce kullan (atýla yaklaþan)")
        i = i + 1
        If adet(i) > 0 Then ayr = ayr & IIf(Len(ayr) > 0, "     •     ", "") & ad & ": " & adet(i) & " kalem, " & AKSORA_SAYI_METNI(Round(top(i), 2), "#,##0.00") & " TL"
    Next ad
    ws.Cells(r, 1).Value2 = ayr
    ws.Cells(r, 1).Font.Size = 10: ws.Cells(r, 1).Font.bold = True: ws.Cells(r, 1).Font.Color = RGB(12, 45, 79)
    ws.Rows(r).RowHeight = 22
    r = r + 1
    Y_BASLIK_SATIRI ws.Range(ws.Cells(r, 1), ws.Cells(r, 12)), Array("Ýl", "Sistem", "Malzeme", "Kýsa metin", "Depo", "Miktar", _
        "Deðer", "Birim", "Ne yapýlmalý", "Durum", "En eski giriþ", "Ayrýntý")
    ws.Rows(r).RowHeight = 30
    a1 = r + 1
    If n = 0 Then
        ws.Cells(a1, 4).Value2 = "Atýl ya da atýla yaklaþan stok yok."
        ws.Cells(a1, 4).Font.Italic = True: ws.Cells(a1, 4).Font.Color = RGB(21, 128, 61)
        Exit Sub
    End If
    ReDim anah(1 To n): ReDim idx(1 To n)
    i = 0
    For Each o In isler
        i = i + 1: idx(i) = i
        anah(i) = CStr(o(0)) & I_SIRA_ANAHTAR(CDbl(o(7)))
    Next o
    Y_SIRALA anah, idx, 1, n
    ReDim sonuc(1 To n, 1 To 12)
    For i = 1 To n
        o = isler(idx(i))
        For kk = 1 To 12: sonuc(i, kk) = o(kk): Next kk
    Next i
    ws.Range(ws.Cells(a1, 2), ws.Cells(a1 + n - 1, 5)).NumberFormat = "@"
    ws.Range(ws.Cells(a1, 8), ws.Cells(a1 + n - 1, 12)).NumberFormat = "@"
    ws.Range(ws.Cells(a1, 1), ws.Cells(a1 + n - 1, 12)).Value2 = sonuc
    r = a1 + n - 1
    Y_TABLO_GOVDE ws.Range(ws.Cells(a1, 1), ws.Cells(r, 12))
    ws.Range(ws.Rows(a1), ws.Rows(r)).RowHeight = 20
    ws.Range(ws.Cells(a1, 1), ws.Cells(r, 3)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 4), ws.Cells(r, 4)).IndentLevel = 1
    ws.Range(ws.Cells(a1, 5), ws.Cells(r, 5)).ShrinkToFit = True
    AKSORA_V13_MIKTAR_BICIMI ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6))
    ws.Range(ws.Cells(a1, 6), ws.Cells(r, 6)).Font.bold = True
    ws.Range(ws.Cells(a1, 7), ws.Cells(r, 7)).NumberFormat = "#,##0.00 ""TL"""
    ws.Range(ws.Cells(a1, 8), ws.Cells(r, 11)).HorizontalAlignment = xlCenter
    ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)).ShrinkToFit = True
    ws.Range(ws.Cells(a1, 12), ws.Cells(r, 12)).ShrinkToFit = True
    Y_DURUM_KF ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)), "KENDÝ ÝHTÝYACINDA KULLAN", RGB(220, 38, 38), RGB(255, 255, 255)
    Y_DURUM_KF ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)), "TRANSFER ET", RGB(35, 119, 204), RGB(255, 255, 255)
    Y_DURUM_KF ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)), "ÖNCE BUNU KULLAN", RGB(220, 252, 231), RGB(21, 128, 61)
    Y_DURUM_KF ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)), "BEKLET", RGB(241, 245, 249), RGB(71, 85, 105)
    Y_DURUM_KF ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)), "SAT / HURDA / ÝADE", RGB(237, 233, 254), RGB(91, 33, 182)
    Y_DURUM_KF ws.Range(ws.Cells(a1, 9), ws.Cells(r, 9)), "ÖNCE KULLAN", RGB(254, 243, 199), RGB(146, 64, 14)
    ATIL_DURUM_RENK ws.Range(ws.Cells(a1, 10), ws.Cells(r, 10))
    ws.Range(ws.Cells(a1 - 1, 1), ws.Cells(r, 12)).AutoFilter
    r = r + 2
    ws.Cells(r, 1).Value2 = "Atýl: stoðun 24 aydan eski giriþe ait kýsmý. Önce kullan: 18-24 ay arasý (3 / 6 ay içinde atýl olacak). Deðer: miktar x AKSORA stok birim deðeri."
    ws.Cells(r, 1).Font.Size = 9: ws.Cells(r, 1).Font.Color = RGB(71, 85, 105)
End Sub

' Eylem satiri: Array(0 oncelik, 1 il, 2 sistem, 3 malzeme, 4 tanim, 5 depo, 6 miktar, 7 deger, 8 birim, 9 eylem, 10 durum, 11 giris, 12 ayrinti, 13 il kodu)
Private Sub AN_ISLEM_EKLE(ByVal sat As Collection, ByVal x As Variant, ByVal sira As Long, ByVal eylem As String, ByVal q As Double, _
    ByVal bd As Double, ByVal ayrinti As String)
    If q <= 0.0005 Then Exit Sub
    sat.Add Array(sira, Y_IL_ADI(CStr(x(0))), CStr(x(1)), CStr(x(3)), CStr(x(4)), ATIL_DEPO_METIN(x(15)), Round(q, 3), Round(q * bd, 2), _
        CStr(x(7)), eylem, ATIL_DURUM_ADI(CLng(x(10))), CStr(x(9)), ayrinti, CStr(x(0)))
End Sub

' Yonetim Merkezi en ustte: BUGUN YAPILACAK 5 IS. Malzeme analiz isleri (tutarli) + diger birimlerin kritik / uyari maddeleri;
' oncelik: kritik stok, kritik uyarilar, transfer, gereksiz talep, fazla teslim, veri yasi, diger uyarilar. Donus: son satir.
Public Function ADM_YAPILACAK_CIZ(ByVal ws As Worksheet, ByVal r As Long, ByVal uy As Collection, ByVal ilMetin As String) As Long
    Dim liste As Collection, x As Variant, i As Long, j As Long, n As Long, a() As Variant, t As Variant, ilk As Long, kat As Object
    On Error GoTo Son
    Set liste = New Collection
    Set kat = CreateObject("Scripting.Dictionary"): kat.CompareMode = vbTextCompare
    If Not mAdmIsler Is Nothing Then
        For Each x In mAdmIsler
            liste.Add x
        Next x
    End If
    Dim sev As Long
    For sev = 1 To 2
    For Each x In uy
        Select Case CStr(x(3))
            Case "Analiz verisi", "Kritik stok", "Fazla teslim", "Transfer"
                ' malzeme analiz isleri yukarida tutariyla
            Case Else
                ' Her konudan (demontaj, ISG, teminat...) en onemli tek madde: ilk 5 cesitli kalsin.
                If x(0) = sev And Not kat.Exists(CStr(x(3))) Then
                    kat(CStr(x(3))) = True
                    liste.Add Array(IIf(sev = 1, 15, 70), CStr(x(1)), 0#, CStr(x(2)))
                End If
        End Select
    Next x
    Next sev
    n = liste.Count
    ADM_BOLUM ws, r, ADM_ADRES(r, 2, r, 13), "BUGÜN YAPILACAK " & application.Min(5, n) & " ÝÞ   •   " & ilMetin & "   (en önemliden baþlayarak)"
    r = r + 1
    If n = 0 Then
        ADM_BLOK ws, ADM_ADRES(r, 2, r, 13), "Bekleyen önemli iþ yok.", RGB(220, 252, 231), RGB(21, 128, 61), 10, True, xlLeft
        ADM_YAPILACAK_CIZ = r
        Exit Function
    End If
    ReDim a(1 To n)
    i = 0
    For Each x In liste
        i = i + 1: a(i) = x
    Next x
    For i = 1 To n - 1
        For j = i + 1 To n
            If a(j)(0) < a(i)(0) Or (a(j)(0) = a(i)(0) And a(j)(2) > a(i)(2)) Then t = a(i): a(i) = a(j): a(j) = t
        Next j
    Next i
    ilk = r
    ADM_BASLIKLAR ws, r, 2, Array("No", "Yapýlacak iþ", "", "", "", "", "", "", "Tutar", "", "Nerede", "")
    ws.Range(ADM_ADRES(r, 3, r, 9)).Merge: ws.Range(ADM_ADRES(r, 10, r, 11)).Merge: ws.Range(ADM_ADRES(r, 12, r, 13)).Merge
    For i = 1 To application.Min(5, n)
        r = r + 1
        ws.Rows(r).RowHeight = IIf(Len(CStr(a(i)(1))) > 110, 34, 24)
        ws.Range(ADM_ADRES(r, 2, r, 13)).Interior.Color = IIf(i = 1, RGB(254, 242, 242), vbWhite)
        ws.Range(ADM_ADRES(r, 3, r, 9)).Merge: ws.Range(ADM_ADRES(r, 10, r, 11)).Merge: ws.Range(ADM_ADRES(r, 12, r, 13)).Merge
        ws.Cells(r, 2).Value2 = i
        ws.Cells(r, 3).NumberFormat = "@": ws.Cells(r, 3).Value2 = CStr(a(i)(1))
        If a(i)(2) > 0.005 Then ws.Cells(r, 10).Value2 = Round(a(i)(2), 2): ws.Cells(r, 10).NumberFormat = "#,##0.00 ""TL"""
        ws.Cells(r, 12).NumberFormat = "@": ws.Cells(r, 12).Value2 = CStr(a(i)(3))
        ws.Cells(r, 2).HorizontalAlignment = xlCenter: ws.Cells(r, 2).Font.bold = True: ws.Cells(r, 2).Font.Size = 12
        ws.Cells(r, 3).HorizontalAlignment = xlLeft: ws.Cells(r, 3).IndentLevel = 1: ws.Cells(r, 3).WrapText = True
        ws.Cells(r, 3).VerticalAlignment = xlCenter: ws.Cells(r, 3).Font.Size = 10
        ws.Cells(r, 3).Font.bold = (i = 1)
        ws.Cells(r, 10).HorizontalAlignment = xlRight: ws.Cells(r, 10).Font.bold = True
        ws.Cells(r, 12).HorizontalAlignment = xlCenter: ws.Cells(r, 12).ShrinkToFit = True: ws.Cells(r, 12).Font.Color = RGB(20, 103, 173)
    Next i
    ADM_KENAR ws.Range(ADM_ADRES(ilk, 2, r, 13)), True
    ADM_YAPILACAK_CIZ = r
    Exit Function
Son:
    ADM_YAPILACAK_CIZ = r
End Function
