Attribute VB_Name = "modKS_HESAP"
Option Explicit

' Türetilmiþ alanlarýn tek hesap merkezi. Form, toplu aktarým ve panel ayný kurallarý kullanýr.
'   Sevkiyat: tutar, gecikme, teslim durumu, garanti bitiþi (gerçek geliþ + garanti ayý), garanti durumu
'   Arýza   : garanti kapsamý, üretici dönüþ süresi, çözüm süresi, durum, önceki arýza, tekrar sayýsý
'   Malzeme : toplam sevk / arýzalý miktar, arýza oraný, son arýza tarihi

Private mUyariGun As Long

Public Sub KS_ParamYukle()
    mUyariGun = CLng(KS_ParamSayi("pGarUyari", 60))
End Sub

Public Sub KS_SevkHesapla(ByRef a As Variant, ByVal i As Long)
    Dim plan As Date, gelis As Date, bit As Date, hp As Boolean, hg As Boolean
    Dim adet As Double, fiyat As Double, gar As Double, kalan As Long
    If mUyariGun = 0 Then KS_ParamYukle
    hp = KS_Tarih(a(i, S_PLAN), plan)
    hg = KS_Tarih(a(i, S_GELIS), gelis)
    If KS_Sayi(a(i, S_ADET), adet) And KS_Sayi(a(i, S_FIYAT), fiyat) Then
        a(i, S_TUTAR) = Round(adet * fiyat, 2)
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
            a(i, S_GDURUM) = "DEVAM"
        End If
    Else
        a(i, S_GARBIT) = Empty
        If hg Then a(i, S_GDURUM) = "SÜRE TANIMSIZ" Else a(i, S_GDURUM) = "TESLÝM BEKLÝYOR"
    End If
End Sub

Public Sub KS_ArizaHesapla(ByRef a As Variant, ByVal i As Long)
    Dim t As Date, bit As Date, bil As Date, don As Date, coz As Date
    Dim ht As Boolean, hb As Boolean, hd As Boolean, hc As Boolean
    ht = KS_Tarih(a(i, A_TARIH), t)
    hb = KS_Tarih(a(i, A_BILDIRIM), bil)
    hd = KS_Tarih(a(i, A_DONUS), don)
    hc = KS_Tarih(a(i, A_COZUM), coz)
    If ht And KS_Tarih(a(i, A_GARBIT), bit) Then
        If t <= bit Then a(i, A_KAPSAM) = "EVET" Else a(i, A_KAPSAM) = "HAYIR"
    Else
        a(i, A_KAPSAM) = "BELÝRSÝZ"
    End If
    If hb And hd Then a(i, A_DONUSGUN) = CLng(don - bil) Else a(i, A_DONUSGUN) = Empty
    If ht And hc Then a(i, A_COZUMGUN) = CLng(coz - t) Else a(i, A_COZUMGUN) = Empty
    If hc Then
        a(i, A_DURUM) = "KAPANDI"
    ElseIf hd Then
        a(i, A_DURUM) = "DÖNÜÞ ALINDI"
    ElseIf hb Then
        a(i, A_DURUM) = "ÜRETÝCÝDE"
    Else
        a(i, A_DURUM) = "AÇIK"
    End If
End Sub

' Tüm tablolarý baþtan hesaplar ve geri yazar. Kullanýcý alanlarýna dokunmaz.
Public Sub KS_HesaplaTumu()
    Dim loM As ListObject, loS As ListObject, loA As ListObject
    Dim aM As Variant, arS As Variant, aA As Variant
    Dim dM As Object, dS As Object, grp As Object, k As Variant
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
                If KS_Bos(arS(i, S_GARAY)) Then arS(i, S_GARAY) = aM(mi, M_GARANTI)
                If KS_Sayi(arS(i, S_ADET), mik) Then sevkMik(mi) = sevkMik(mi) + mik
            End If
            KS_SevkHesapla arS, i
            arS(i, S_ARIZA) = 0
        Next i
    End If

    ' Arýzalar
    Set grp = CreateObject("Scripting.Dictionary")
    If IsArray(aA) Then
        For i = 1 To UBound(aA, 1)
            si = 0
            If dS.Exists(KS_Anahtar(aA(i, A_SEVK))) Then si = dS(KS_Anahtar(aA(i, A_SEVK)))
            If si > 0 Then
                aA(i, A_KOD) = arS(si, S_KOD)
                aA(i, A_AD) = arS(si, S_AD)
                aA(i, A_SEVKCI) = arS(si, S_SEVKCI)
                aA(i, A_YER) = arS(si, S_YER)
                aA(i, A_GARBIT) = arS(si, S_GARBIT)
                If KS_Sayi(aA(i, A_ADET), mik) Then arS(si, S_ARIZA) = arS(si, S_ARIZA) + mik
            End If
            mi = 0
            If dM.Exists(KS_Anahtar(aA(i, A_KOD))) Then mi = dM(KS_Anahtar(aA(i, A_KOD)))
            If mi > 0 Then
                aA(i, A_URETICI) = aM(mi, M_URETICI)
                If KS_Sayi(aA(i, A_ADET), mik) Then arzMik(mi) = arzMik(mi) + mik
                If KS_Tarih(aA(i, A_TARIH), d) Then
                    If d > sonArz(mi) Then sonArz(mi) = d
                End If
            End If
            KS_ArizaHesapla aA, i
            ' Tekrar grubu: seri no varsa malzeme + seri, yoksa sevkiyat.
            If Len(KS_Anahtar(aA(i, A_SERI))) > 0 Then
                anahtar = "S|" & KS_Anahtar(aA(i, A_KOD)) & "|" & KS_Anahtar(aA(i, A_SERI))
            Else
                anahtar = "V|" & KS_Anahtar(aA(i, A_SEVK))
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

Public Sub KS_YENIDEN_HESAPLA()
    On Error GoTo Hata
    KS_Durum "Tüm kayýtlar yeniden hesaplanýyor..."
    KS_HesaplaTumu
    KS_Durum ""
    Application.StatusBar = False
    KS_Log "YENÝDEN HESAPLA", "", "Tüm türetilmiþ alanlar güncellendi"
    MsgBox "Tüm türetilmiþ alanlar yeniden hesaplandý.", vbInformation, KS_AD
    Exit Sub
Hata:
    KS_Bitir True
    MsgBox "Hesaplama durdu: " & Err.Description, vbExclamation, KS_AD
End Sub
