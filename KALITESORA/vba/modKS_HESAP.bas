Attribute VB_Name = "modKS_HESAP"
Option Explicit

' Türetilmiþ alanlarýn tek hesap merkezi. Form, tablo düzenleme, toplu aktarým ve panel ayný kurallarý kullanýr.
'   Sevkiyat: TL tutar (kur tablosuyla), gecikme, teslim durumu, garanti bitiþi, garanti durumu
'   Arýza   : baðlý sevkiyat (otomatik seçim), garanti kapsamý, dönüþ / çözüm süresi, durum, önceki arýza
'   Malzeme : toplam sevk / arýzalý miktar, arýza oraný, son arýza tarihi

Private mUyariGun As Long
Private mKur As Object

Public Sub KS_ParamYukle()
    Dim lo As ListObject, a As Variant, i As Long, x As Double
    mUyariGun = CLng(KS_ParamSayi("pGarUyari", 60))
    Set mKur = CreateObject("Scripting.Dictionary")
    mKur("TRY") = 1#
    mKur("TL") = 1#
    mKur("") = 1#
    Set lo = KS_LO(T_KUR)
    If Not lo Is Nothing Then
        a = KS_Veri(lo)
        If IsArray(a) Then
            For i = 1 To UBound(a, 1)
                If KS_Sayi(a(i, 2), x) Then
                    If x > 0 Then mKur(KS_Anahtar(a(i, 1))) = x
                End If
            Next i
        End If
    End If
End Sub

' Para biriminin TL kuru; tanýmsýzsa 0.
Public Function KS_Kur(ByVal pb As Variant) As Double
    Dim k As String
    If mKur Is Nothing Then KS_ParamYukle
    k = KS_Anahtar(pb)
    If mKur.Exists(k) Then KS_Kur = mKur(k)
End Function

Public Function KS_ArizaAcik(ByVal durum As Variant) As Boolean
    Select Case KS_Metin(durum)
        Case "DEVAM EDÝYOR", "ÜRETÝCÝDE", "KISMEN ÇÖZÜLDÜ"
            KS_ArizaAcik = True
    End Select
End Function

Public Sub KS_SevkHesapla(ByRef a As Variant, ByVal i As Long)
    Dim plan As Date, gelis As Date, bit As Date, hp As Boolean, hg As Boolean
    Dim adet As Double, fiyat As Double, gar As Double, kalan As Long, kur As Double
    If mKur Is Nothing Then KS_ParamYukle
    hp = KS_Tarih(a(i, S_PLAN), plan)
    hg = KS_Tarih(a(i, S_GELIS), gelis)
    If KS_Bos(a(i, S_PB)) Then a(i, S_PB) = "TRY"
    kur = KS_Kur(a(i, S_PB))
    If KS_Sayi(a(i, S_ADET), adet) And KS_Sayi(a(i, S_FIYAT), fiyat) And kur > 0 Then
        a(i, S_TUTAR) = Round(adet * fiyat * kur, 2)
    Else
        a(i, S_TUTAR) = Empty
    End If
    If hg And hp Then
        a(i, S_GECIKME) = CLng(gelis - plan)
        If gelis <= plan Then a(i, S_TDURUM) = "ZAMANINDA" Else a(i, S_TDURUM) = "GECÝKMELÝ"
    ElseIf hp Then
        If Date > plan Then
            a(i, S_GECIKME) = CLng(Date - plan)
            a(i, S_TDURUM) = "GECÝKÝYOR"
        Else
            a(i, S_GECIKME) = Empty
            a(i, S_TDURUM) = "BEKLENÝYOR"
        End If
    ElseIf hg Then
        a(i, S_GECIKME) = Empty
        a(i, S_TDURUM) = "TESLÝM ALINDI"
    Else
        a(i, S_GECIKME) = Empty
        a(i, S_TDURUM) = "TARÝH YOK"
    End If
    If hg And KS_Sayi(a(i, S_GARAY), gar) Then
        bit = DateAdd("m", CLng(gar), gelis)
        a(i, S_GARBIT) = bit
        kalan = CLng(bit - Date)
        If kalan < 0 Then
            a(i, S_GDURUM) = "BÝTTÝ"
        ElseIf kalan <= mUyariGun Then
            a(i, S_GDURUM) = "BÝTÝYOR (" & kalan & " GÜN)"
        Else
            a(i, S_GDURUM) = "DEVAM EDÝYOR"
        End If
    Else
        a(i, S_GARBIT) = Empty
        If Not hg Then
            a(i, S_GDURUM) = "TESLÝM BEKLÝYOR"
        Else
            a(i, S_GDURUM) = "GARANTÝ GÝRÝLMEDÝ"
        End If
    End If
End Sub

Public Sub KS_ArizaHesapla(ByRef a As Variant, ByVal i As Long)
    Dim t As Date, bit As Date, bil As Date, don As Date, coz As Date
    Dim ht As Boolean, hb As Boolean, hd As Boolean, hc As Boolean, cz As String, dv As String
    ht = KS_Tarih(a(i, A_TARIH), t)
    hb = KS_Tarih(a(i, A_BILDIRIM), bil)
    hd = KS_Tarih(a(i, A_DONUS), don)
    hc = KS_Tarih(a(i, A_COZUM), coz)
    If KS_Metin(a(i, A_KAPKAYNAK)) <> "ELLE" Then
        a(i, A_KAPKAYNAK) = "OTOMATÝK"
        If ht And KS_Tarih(a(i, A_GARBIT), bit) Then
            If t <= bit Then a(i, A_KAPSAM) = "EVET" Else a(i, A_KAPSAM) = "HAYIR"
        Else
            a(i, A_KAPSAM) = "BELÝRSÝZ"
        End If
    End If
    If hb And hd Then a(i, A_DONUSGUN) = CLng(don - bil) Else a(i, A_DONUSGUN) = Empty
    If ht And hc Then a(i, A_COZUMGUN) = CLng(coz - t) Else a(i, A_COZUMGUN) = Empty
    cz = KS_Anahtar(a(i, A_COZULDU))
    dv = KS_Anahtar(a(i, A_DEVAM))
    If cz = "EVET" Or (hc And cz <> "HAYIR" And cz <> "KISMEN") Then
        a(i, A_DURUM) = "ÇÖZÜLDÜ"
    ElseIf cz = "KISMEN" Then
        a(i, A_DURUM) = "KISMEN ÇÖZÜLDÜ"
    ElseIf cz = "HAYIR" And dv = "HAYIR" Then
        a(i, A_DURUM) = "ÇÖZÜLEMEDÝ"
    ElseIf hb And Not hd Then
        a(i, A_DURUM) = "ÜRETÝCÝDE"
    Else
        a(i, A_DURUM) = "DEVAM EDÝYOR"
    End If
End Sub

' Arýzanýn baðlanacaðý sevkiyatý seçer: seri no eþleþmesi, yoksa arýza tarihinden önceki en son teslim.
' adaylar: ayný malzeme koduna sahip sevkiyat satýrlarýnýn "," ile ayrýlmýþ listesi.
Public Function KS_SevkSec(ByRef arS As Variant, ByVal adaylar As String, ByVal seri As String, _
        ByVal tarih As Variant) As Long
    Dim p As Variant, i As Long, n As Long, t As Date, ht As Boolean, d As Date, enIyi As Date, sec As Long
    Dim planSec As Long, planEnIyi As Date
    If Len(adaylar) = 0 Then Exit Function
    p = Split(adaylar, ",")
    seri = KS_Anahtar(seri)
    If Len(seri) > 0 Then
        For i = 0 To UBound(p)
            n = CLng(p(i))
            If Len(KS_Anahtar(arS(n, S_SERI))) > 0 Then
                If InStr(1, KS_Anahtar(arS(n, S_SERI)), seri, vbBinaryCompare) > 0 Then
                    KS_SevkSec = n
                    Exit Function
                End If
            End If
        Next i
    End If
    ht = KS_Tarih(tarih, t)
    For i = 0 To UBound(p)
        n = CLng(p(i))
        If KS_Tarih(arS(n, S_GELIS), d) Then
            If (Not ht Or d <= t) And (sec = 0 Or d > enIyi) Then
                sec = n
                enIyi = d
            End If
        ElseIf KS_Tarih(arS(n, S_PLAN), d) Then
            If (Not ht Or d <= t) And (planSec = 0 Or d > planEnIyi) Then
                planSec = n
                planEnIyi = d
            End If
        End If
    Next i
    If sec = 0 Then sec = planSec
    KS_SevkSec = sec
End Function

' Malzeme kodu -> sevkiyat satýrlarý ("1,5,9") sözlüðü.
Public Function KS_KodSevkleri(ByRef arS As Variant) As Object
    Dim d As Object, i As Long, k As String
    Set d = CreateObject("Scripting.Dictionary")
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            k = KS_Anahtar(arS(i, S_KOD))
            If Len(k) > 0 Then
                If d.Exists(k) Then d(k) = d(k) & "," & i Else d(k) = CStr(i)
            End If
        Next i
    End If
    Set KS_KodSevkleri = d
End Function

' Arýza satýrýna baðlý sevkiyat ve malzeme bilgilerini doldurur.
Public Sub KS_ArizaBagla(ByRef aA As Variant, ByVal i As Long, ByRef arS As Variant, ByVal dS As Object, _
        ByVal dKS As Object, ByRef aM As Variant, ByVal dM As Object)
    Dim si As Long, mi As Long, k As String
    k = KS_Anahtar(aA(i, A_KOD))
    If dM.Exists(k) Then
        mi = dM(k)
        aA(i, A_AD) = aM(mi, M_AD)
        aA(i, A_URETICI) = aM(mi, M_URETICI)
    End If
    If Len(KS_Anahtar(aA(i, A_SEVK))) > 0 Then
        If dS.Exists(KS_Anahtar(aA(i, A_SEVK))) Then si = dS(KS_Anahtar(aA(i, A_SEVK)))
    End If
    If si = 0 And dKS.Exists(k) Then
        si = KS_SevkSec(arS, dKS(k), KS_Metin(aA(i, A_SERI)), aA(i, A_TARIH))
        If si > 0 Then aA(i, A_SEVK) = arS(si, S_NO)
    End If
    If si > 0 Then
        aA(i, A_SEVKCI) = arS(si, S_SEVKCI)
        aA(i, A_YER) = arS(si, S_YER)
        aA(i, A_GARBIT) = arS(si, S_GARBIT)
        If KS_Bos(aA(i, A_AD)) Then aA(i, A_AD) = arS(si, S_AD)
    Else
        aA(i, A_GARBIT) = Empty
    End If
End Sub

' Tüm tablolarý baþtan hesaplar ve geri yazar. Kullanýcý alanlarýna dokunmaz.
Public Sub KS_HesaplaTumu()
    Dim loM As ListObject, loS As ListObject, loA As ListObject
    Dim aM As Variant, arS As Variant, aA As Variant
    Dim dM As Object, dS As Object, dKS As Object, grp As Object, k As Variant
    Dim i As Long, j As Long, n As Long, si As Long, mi As Long
    Dim mik As Double, d As Date, d2 As Date, enYakin As Date, bulundu As Boolean, sira As Long
    Dim sevkMik() As Double, arzMik() As Double, sonArz() As Date, uyeler As Variant, anahtar As String

    KS_Basla
    KS_ParamYukle
    Set loM = KS_LOZ(T_MALZ)
    Set loS = KS_LOZ(T_SEVK)
    Set loA = KS_LOZ(T_ARIZA)
    aM = KS_Veri(loM)
    arS = KS_Veri(loS)
    aA = KS_Veri(loA)

    Set dM = CreateObject("Scripting.Dictionary")
    If IsArray(aM) Then
        ReDim sevkMik(1 To UBound(aM, 1))
        ReDim arzMik(1 To UBound(aM, 1))
        ReDim sonArz(1 To UBound(aM, 1))
        For i = 1 To UBound(aM, 1)
            dM(KS_Anahtar(aM(i, M_KOD))) = i
            If KS_Bos(aM(i, M_PB)) Then aM(i, M_PB) = "TRY"
        Next i
    End If

    ' Sevkiyatlar
    Set dS = CreateObject("Scripting.Dictionary")
    If IsArray(arS) Then
        For i = 1 To UBound(arS, 1)
            dS(KS_Anahtar(arS(i, S_NO))) = i
            mi = 0
            If dM.Exists(KS_Anahtar(arS(i, S_KOD))) Then mi = dM(KS_Anahtar(arS(i, S_KOD)))
            If mi > 0 Then
                arS(i, S_AD) = aM(mi, M_AD)
                If KS_Bos(arS(i, S_GARAY)) And Not KS_Bos(aM(mi, M_GARANTI)) Then arS(i, S_GARAY) = aM(mi, M_GARANTI)
                If KS_Sayi(arS(i, S_ADET), mik) Then sevkMik(mi) = sevkMik(mi) + mik
            End If
            KS_SevkHesapla arS, i
            arS(i, S_ARIZA) = 0
        Next i
    End If
    Set dKS = KS_KodSevkleri(arS)

    ' Arýzalar
    Set grp = CreateObject("Scripting.Dictionary")
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            KS_ArizaBagla aA, i, arS, dS, dKS, aM, dM
            si = 0
            If dS.Exists(KS_Anahtar(aA(i, A_SEVK))) Then si = dS(KS_Anahtar(aA(i, A_SEVK)))
            If si > 0 Then
                If KS_Sayi(aA(i, A_ADET), mik) Then arS(si, S_ARIZA) = arS(si, S_ARIZA) + mik
            End If
            mi = 0
            If dM.Exists(KS_Anahtar(aA(i, A_KOD))) Then mi = dM(KS_Anahtar(aA(i, A_KOD)))
            If mi > 0 Then
                If KS_Sayi(aA(i, A_ADET), mik) Then arzMik(mi) = arzMik(mi) + mik
                If KS_Tarih(aA(i, A_TARIH), d) Then
                    If d > sonArz(mi) Then sonArz(mi) = d
                End If
            End If
            KS_ArizaHesapla aA, i
            ' Tekrar grubu: seri no varsa malzeme + seri, yoksa malzeme + sevkiyat.
            If Len(KS_Anahtar(aA(i, A_SERI))) > 0 Then
                anahtar = "S|" & KS_Anahtar(aA(i, A_KOD)) & "|" & KS_Anahtar(aA(i, A_SERI))
            Else
                anahtar = "V|" & KS_Anahtar(aA(i, A_KOD)) & "|" & KS_Anahtar(aA(i, A_SEVK))
            End If
            If grp.Exists(anahtar) Then
                grp(anahtar) = grp(anahtar) & "," & i
            Else
                grp(anahtar) = CStr(i)
            End If
        Next i
        For Each k In grp.Keys
            uyeler = Split(grp(k), ",")
            For i = 0 To UBound(uyeler)
                n = CLng(uyeler(i))
                bulundu = False
                sira = 1
                If KS_Tarih(aA(n, A_TARIH), d) Then
                    For j = 0 To UBound(uyeler)
                        If j <> i Then
                            If KS_Tarih(aA(CLng(uyeler(j)), A_TARIH), d2) Then
                                If d2 < d Or (d2 = d And j < i) Then
                                    sira = sira + 1
                                    If d2 < d Then
                                        If Not bulundu Or d2 > enYakin Then enYakin = d2
                                        bulundu = True
                                    End If
                                End If
                            End If
                        End If
                    Next j
                End If
                If bulundu Then aA(n, A_ONCEKI) = enYakin Else aA(n, A_ONCEKI) = Empty
                aA(n, A_TEKRAR) = sira
            Next i
        Next k
    End If

    ' Malzeme özetleri
    If IsArray(aM) Then
        For i = 1 To UBound(aM, 1)
            aM(i, M_SEVKMIK) = sevkMik(i)
            aM(i, M_ARIZAMIK) = arzMik(i)
            If sevkMik(i) > 0 Then aM(i, M_ORAN) = arzMik(i) / sevkMik(i) Else aM(i, M_ORAN) = Empty
            If sonArz(i) > 0 Then aM(i, M_SONARIZA) = sonArz(i) Else aM(i, M_SONARIZA) = Empty
        Next i
    End If

    KS_DiziGeriYaz loM, aM
    KS_DiziGeriYaz loS, arS
    KS_DiziGeriYaz loA, aA
    KS_MailBaglantilari loA
    KS_KilitAyarla loS, KS_DuzenSevk()
    KS_KilitAyarla loM, KS_DuzenMalz()
    KS_Bitir
End Sub

Private Sub KS_DiziGeriYaz(ByVal lo As ListObject, ByRef a As Variant)
    Dim korumali As Boolean
    If Not IsArray(a) Then Exit Sub
    korumali = lo.Parent.ProtectContents
    KS_Ac lo.Parent
    lo.DataBodyRange.Value = a
    If korumali Then KS_Kapat lo.Parent
End Sub

' Mail dosyasý sütununda yol varsa týklanabilir köprü yapar.
Public Sub KS_MailBaglantilari(ByVal lo As ListObject)
    Dim r As Range, c As Range, korumali As Boolean, yol As String
    If lo.DataBodyRange Is Nothing Then Exit Sub
    korumali = lo.Parent.ProtectContents
    KS_Ac lo.Parent
    Set r = lo.ListColumns(A_MAILDOSYA).DataBodyRange
    r.Hyperlinks.Delete
    For Each c In r.Cells
        yol = KS_Metin(c.Value)
        If Len(yol) > 0 Then
            lo.Parent.Hyperlinks.Add Anchor:=c, Address:=yol, TextToDisplay:=yol
        End If
    Next c
    r.Font.Size = 9
    If korumali Then KS_Kapat lo.Parent
End Sub

Public Sub KS_YENIDEN_HESAPLA()
    On Error GoTo Hata
    KS_Durum "Tüm kayýtlar yeniden hesaplanýyor..."
    KS_HesaplaTumu
    Application.StatusBar = False
    KS_Log "YENÝDEN HESAPLA", "", "Tüm türetilmiþ alanlar güncellendi"
    MsgBox "Tüm türetilmiþ alanlar yeniden hesaplandý.", vbInformation, KS_AD
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Hesaplama durdu: " & Err.Description, vbExclamation, KS_AD
End Sub

' Elle düzenlenebilir sütunlar yalnýz tablo satýrlarýnda açýktýr; tablonun altýna yazýlarak yarým satýr oluþmaz.
Public Sub KS_KilitAyarla(ByVal lo As ListObject, ByVal kolonlar As Variant)
    Dim ws As Worksheet, k As Variant, c As Long, korumali As Boolean
    Set ws = lo.Parent
    korumali = ws.ProtectContents
    KS_Ac ws
    For Each k In kolonlar
        c = lo.HeaderRowRange.Cells(1, CLng(k)).Column
        ws.Range(ws.Cells(lo.HeaderRowRange.Row + 1, c), ws.Cells(ws.Rows.Count, c)).Locked = True
        If Not lo.DataBodyRange Is Nothing Then lo.ListColumns(CLng(k)).DataBodyRange.Locked = False
    Next k
    If korumali Then KS_Kapat ws
End Sub

Public Function KS_DuzenSevk() As Variant
    KS_DuzenSevk = Array(S_GELIS, S_GARAY, S_IRS, S_SERI, S_NOT)
End Function

Public Function KS_DuzenMalz() As Variant
    KS_DuzenMalz = Array(M_KAT, M_BIRIM, M_GARANTI, M_NOT)
End Function

' SEVKÝYATLAR'da seçili (görünür) satýrlara gerçek geliþ tarihi yazar; dolu geliþ tarihleri deðiþmez.
' "P" yazýlýrsa her satýra kendi planlanan teslim tarihi (bugünden ileri olmayanlar) yazýlýr.
Public Sub KS_TOPLU_GELIS()
    Dim lo As ListObject, govde As Range, sec As Range, ar As Range, r As Range, satirlar As Object, k As Variant
    Dim a As Variant, s As String, d As Date, d2 As Date, plan As Boolean, i As Long, n As Long, atla As Long
    On Error GoTo Hata
    Set lo = KS_LOZ(T_SEVK)
    Set govde = lo.DataBodyRange
    If govde Is Nothing Then Exit Sub
    If ActiveSheet.Name <> lo.Parent.Name Then lo.Parent.Activate
    If TypeName(Selection) = "Range" Then Set sec = Intersect(Selection.EntireRow, govde)
    If sec Is Nothing Then
        MsgBox "Önce SEVKÝYATLAR tablosunda geliþ tarihi yazýlacak satýrlarý seçin." & vbCrLf & vbCrLf & _
            "Ýpucu: tabloyu süzün (ARA, filtre düðmeleri veya sütun oklarý), ilk satýra týklayýp " & _
            "Ctrl+Shift+Aþaðý ile görünen satýrlarý seçin.", vbInformation, KS_AD
        Exit Sub
    End If
    Set satirlar = CreateObject("Scripting.Dictionary")
    For Each ar In sec.Areas
        For Each r In ar.Rows
            If Not r.EntireRow.Hidden Then satirlar(r.Row - govde.Row + 1) = True
        Next r
    Next ar
    If satirlar.Count = 0 Then Exit Sub
    s = Trim$(InputBox(satirlar.Count & " sevkiyata gerçek geliþ tarihi yazýlacak." & vbCrLf & vbCrLf & _
        "Tarih yazýn (gg.aa.yyyy)" & vbCrLf & _
        "ya da P yazýn: her satýra kendi planlanan teslim tarihi yazýlýr." & vbCrLf & vbCrLf & _
        "Geliþ tarihi dolu olan satýrlar deðiþtirilmez.", KS_AD, Format$(Date, "dd.mm.yyyy")))
    If Len(s) = 0 Then Exit Sub
    plan = (UCase$(s) = "P")
    If Not plan Then
        If Not KS_Tarih(s, d) Then
            MsgBox "Tarih anlaþýlamadý: " & s, vbExclamation, KS_AD
            Exit Sub
        End If
        If d > Date Then
            MsgBox "Gerçek geliþ tarihi bugünden ileri olamaz.", vbExclamation, KS_AD
            Exit Sub
        End If
    End If
    a = KS_Veri(lo)
    KS_Basla
    KS_Ac lo.Parent
    For Each k In satirlar.Keys
        i = CLng(k)
        If KS_Bos(a(i, S_GELIS)) Then
            If plan Then
                If KS_Tarih(a(i, S_PLAN), d2) And d2 <= Date Then
                    govde.Cells(i, S_GELIS).Value = d2
                    n = n + 1
                Else
                    atla = atla + 1
                End If
            Else
                govde.Cells(i, S_GELIS).Value = d
                n = n + 1
            End If
        Else
            atla = atla + 1
        End If
    Next k
    KS_Kapat lo.Parent
    KS_Bitir
    KS_Log "TOPLU GELÝÞ", "", n & " sevkiyat" & IIf(plan, " (planlanan teslim tarihi)", " (" & Format$(d, "dd.mm.yyyy") & ")")
    KS_PanelYenile False
    MsgBox n & " sevkiyata gerçek geliþ tarihi yazýldý; teslim ve garanti durumlarý yeniden hesaplandý." & _
        IIf(atla > 0, vbCrLf & atla & " satýr atlandý (geliþ zaten dolu" & IIf(plan, " ya da planlanan tarih ileride", "") & ").", ""), _
        vbInformation, KS_AD
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Toplu geliþ yazýlamadý: " & Err.Description, vbExclamation, KS_AD
End Sub

' MALZEMELER'de seçili (görünür) satýrlara garanti süresi (ay) yazar; dolu garanti deðerleri deðiþmez.
' Malzemenin garantisi, garanti girilmemiþ sevkiyatlarýna kendiliðinden uygulanýr.
Public Sub KS_TOPLU_GARANTI()
    Dim lo As ListObject, govde As Range, sec As Range, ar As Range, r As Range, satirlar As Object, k As Variant
    Dim a As Variant, s As String, ay As Double, i As Long, n As Long, atla As Long
    On Error GoTo Hata
    Set lo = KS_LOZ(T_MALZ)
    Set govde = lo.DataBodyRange
    If govde Is Nothing Then Exit Sub
    If ActiveSheet.Name <> lo.Parent.Name Then lo.Parent.Activate
    If TypeName(Selection) = "Range" Then Set sec = Intersect(Selection.EntireRow, govde)
    If sec Is Nothing Then
        MsgBox "Önce MALZEMELER tablosunda garanti süresi yazýlacak satýrlarý seçin." & vbCrLf & vbCrLf & _
            "Ýpucu: GARANTÝSÝ GÝRÝLMEMÝÞ ya da sütun oklarýyla süzün (ör. ayný üretici), " & _
            "ilk satýra týklayýp Ctrl+Shift+Aþaðý ile görünen satýrlarý seçin.", vbInformation, KS_AD
        Exit Sub
    End If
    Set satirlar = CreateObject("Scripting.Dictionary")
    For Each ar In sec.Areas
        For Each r In ar.Rows
            If Not r.EntireRow.Hidden Then satirlar(r.Row - govde.Row + 1) = True
        Next r
    Next ar
    If satirlar.Count = 0 Then Exit Sub
    s = Trim$(InputBox(satirlar.Count & " malzemeye garanti süresi yazýlacak." & vbCrLf & vbCrLf & _
        "Garanti süresini ay olarak yazýn (ör. 24). Garantisi olmayan için 0." & vbCrLf & vbCrLf & _
        "Garantisi dolu olan malzemeler deðiþtirilmez.", KS_AD, "24"))
    If Len(s) = 0 Then Exit Sub
    If Not KS_Sayi(s, ay) Then
        MsgBox "Sayý anlaþýlamadý: " & s, vbExclamation, KS_AD
        Exit Sub
    End If
    If ay < 0 Or ay > 240 Or ay <> Int(ay) Then
        MsgBox "Garanti süresi 0 - 240 arasý tam ay olmalý.", vbExclamation, KS_AD
        Exit Sub
    End If
    a = KS_Veri(lo)
    KS_Basla
    KS_Ac lo.Parent
    For Each k In satirlar.Keys
        i = CLng(k)
        If KS_Bos(a(i, M_GARANTI)) Then
            govde.Cells(i, M_GARANTI).Value = ay
            n = n + 1
        Else
            atla = atla + 1
        End If
    Next k
    KS_Kapat lo.Parent
    KS_Bitir
    KS_Log "TOPLU GARANTÝ", "", n & " malzeme (" & ay & " ay)"
    KS_PanelYenile False
    MsgBox n & " malzemeye " & ay & " ay garanti yazýldý; garanti girilmemiþ sevkiyatlarýna uygulandý." & _
        IIf(atla > 0, vbCrLf & atla & " satýr atlandý (garanti zaten dolu).", ""), vbInformation, KS_AD
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Toplu garanti yazýlamadý: " & Err.Description, vbExclamation, KS_AD
End Sub
