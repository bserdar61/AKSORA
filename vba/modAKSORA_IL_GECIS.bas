Attribute VB_Name = "modAKSORA_IL_GECIS"
Option Explicit

' Yalnizca stok veri alanlarini degistirir. Grafik, sekil ve gecmise dokunmaz.
' Diger iller icin arsiv yoksa gecis yapmaz; SAP guncelleme baglantisi degildir.
Private Const FIRST_DATA As Long = 7
Private Const LAST_COL As String = "L"

Public Sub AKSORA_IL_GECIS_V1(Optional ByVal secilenIl As String = "")
    Static aktarimSuruyor As Boolean
    Dim eski As String, hedef As String, girdi As String
    Dim kaynak As Variant, etiket As Variant, raporlar As Variant
    Dim ws As Worksheet, ar As Worksheet, kap As Worksheet
    Dim i As Long
    Dim olay As Boolean, ekran As Boolean
    Dim ayarDegisti As Boolean, geriAl As Boolean
    Dim hataMetni As String, geriHata As String

    If AKSORA_V11_MESGUL Or aktarimSuruyor Then Exit Sub
    olay = application.EnableEvents
    ekran = application.ScreenUpdating
    On Error GoTo Hata

    Set kap = ThisWorkbook.Worksheets("KAPAK")
    eski = AKSORA_CANLI_STOK_IL()
    If Len(Trim$(secilenIl)) > 0 Then
        girdi = secilenIl
    Else
        girdi = InputBox( _
            "Geçilecek il: ARTVIN / RIZE / TRABZON / GUMUSHANE / GIRESUN", _
            "AKSORA ÝL GEÇÝÞ", AKSORA_AKTIF_IL())
    End If

    If Len(Trim$(girdi)) = 0 Then Exit Sub

    hedef = AKG_IL(girdi)

    If Len(hedef) = 0 Then
        MsgBox "Geçersiz il. Ýþlem yapýlmadý.", vbExclamation, "AKSORA"
        Exit Sub
    End If

    If hedef = eski Then
        MsgBox "Bu ilin stoklarý zaten yüklü: " & eski, _
               vbInformation, "AKSORA"
        Exit Sub
    End If

    kaynak = Array("CDP_STOK", "EPP_STOK", "STOKLAR", _
                   "ILCE_STOK", "YUKLENICI_STOK")

    etiket = Array("CDP", "EPP", "TUM", "ILC", "YUK")

    raporlar = Array("CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", _
                    "ILCE_RAPORLAR", "YUKLENICI_RAPOR")

    AKG_HEDEF_ARSIV_KONTROL hedef

    For i = 0 To 4
        Set ws = AKG_SAYFA(CStr(kaynak(i)))
        Set ar = AKG_SAYFA("_AKLIVE_" & eski & "_" & etiket(i))

        If ws Is Nothing Or ar Is Nothing Then
            Err.Raise vbObjectError + 830, , _
                      "Mevcut stok veya kaynak il arþivi eksik."
        End If

        If ws.ProtectContents Or ar.ProtectContents Then
            Err.Raise vbObjectError + 831, , _
                      "Stok veya kaynak arþiv veri alaný korumalý."
        End If

        Set ws = AKG_SAYFA(CStr(raporlar(i)))

        If ws Is Nothing Then
            Err.Raise vbObjectError + 832, , _
                      "Rapor sayfasý eksik: " & raporlar(i)
        End If

        If ws.ProtectContents Then
            Err.Raise vbObjectError + 833, , _
                      "Rapor veri alaný korumalý: " & ws.Name
        End If
    Next i

    If Len(Trim$(secilenIl)) = 0 Then
        If MsgBox( _
        hedef & " ana paneline geçilsin mi?", _
        vbYesNo + vbQuestion, "AKSORA | Ýl Seçimi") <> vbYes Then Exit Sub
    End If

    aktarimSuruyor = True
    ayarDegisti = True
    application.EnableEvents = False
    application.ScreenUpdating = False

    For i = 0 To 4
        AKG_KOPYALA AKG_SAYFA(CStr(kaynak(i))), _
            AKG_SAYFA("_AKLIVE_" & eski & "_" & etiket(i))
    Next i

    geriAl = True

    For i = 0 To 4
        AKG_KOPYALA _
            AKG_SAYFA("_AKLIVE_" & hedef & "_" & etiket(i)), _
            AKG_SAYFA(CStr(kaynak(i)))
    Next i

    kap.Range("AZ1").Value2 = hedef

    If AKSORA_CANLI_STOK_IL() <> hedef Then
        Err.Raise vbObjectError + 834, , _
                  "Yüklenen stoklarýn il doðrulamasý baþarýsýz."
    End If

    AKG_RAPORLARI_YENILE
        AKSORA_V8_KAPAK_GORUNUM

    geriAl = False
    application.EnableEvents = olay
    application.ScreenUpdating = ekran
    ayarDegisti = False
    aktarimSuruyor = False

    If Len(Trim$(secilenIl)) = 0 Then
        MsgBox hedef & " stoklarý ve raporlarý hazýr.", _
               vbInformation, "AKSORA"
    End If
    
    Exit Sub

Hata:
    aktarimSuruyor = False
    hataMetni = Err.Description
    On Error Resume Next

    If geriAl Then
        For i = 0 To 4
            Err.Clear

            AKG_KOPYALA _
                AKG_SAYFA("_AKLIVE_" & eski & "_" & etiket(i)), _
                AKG_SAYFA(CStr(kaynak(i)))

            If Err.Number <> 0 Then
                geriHata = geriHata & vbCrLf & _
                           "Stok geri alma hatasý: " & kaynak(i)
            End If
        Next i

        Err.Clear
        kap.Range("AZ1").Value2 = eski
        If Err.Number <> 0 Then
            geriHata = geriHata & vbCrLf & "Ýl seçimi geri alýnamadý."
        End If

        If Len(geriHata) = 0 Then
            Err.Clear
            AKG_RAPORLARI_YENILE
            If Err.Number <> 0 Then
                geriHata = geriHata & vbCrLf & _
                           "Önceki ilin raporlarý yenilenemedi."
            End If
        End If
    End If

    If ayarDegisti Then
        application.EnableEvents = olay
        application.ScreenUpdating = ekran
    End If

    On Error GoTo 0

    MsgBox "Ýl geçiþi tamamlanamadý." & vbCrLf & _
           hataMetni & geriHata, vbExclamation, "AKSORA"
End Sub

Private Function AKG_IL(ByVal deger As Variant) As String
    Dim il As String
    il = AKFIX_NORMAL(CStr(deger))
    Select Case il
        Case "ARTVIN", "RIZE", "TRABZON", "GUMUSHANE", "GIRESUN": AKG_IL = il
    End Select
End Function

Private Function AKG_SAYFA(ByVal ad As String) As Worksheet
    Dim w As Worksheet, x As String, y As String
    x = AKG_NORMAL(ad)
    For Each w In ThisWorkbook.Worksheets
        y = AKG_NORMAL(w.Name)
        If x = y Then Set AKG_SAYFA = w: Exit Function
    Next w
End Function

Private Function AKG_NORMAL(ByVal s As String) As String
    s = UCase$(Trim$(s))
    s = Replace(s, ChrW(304), "I")
    s = Replace(s, ChrW(305), "I")
    s = Replace(s, ChrW(199), "C")
    s = Replace(s, ChrW(231), "C")
    s = Replace(s, ChrW(220), "U")
    s = Replace(s, ChrW(252), "U")
    s = Replace(s, ChrW(350), "S")
    s = Replace(s, ChrW(351), "S")
    s = Replace(s, ChrW(286), "G")
    s = Replace(s, ChrW(287), "G")
    s = Replace(s, ChrW(214), "O")
    s = Replace(s, ChrW(246), "O")
    AKG_NORMAL = s
End Function

Private Sub AKG_KOPYALA(ByVal src As Worksheet, ByVal dst As Worksheet)
    Dim son As Long, hedefSon As Long
    Dim c As Long, r As Long, adet As Long
    Dim kaynakArsiv As Boolean, hedefArsiv As Boolean
    Dim il As String, tur As String, tesis As String
    Dim cdp As String, epp As String
    Dim bilgi As Variant, kontrol As Variant, tamToplam As Variant

    If src Is Nothing Or dst Is Nothing Then
        Err.Raise vbObjectError + 741, "AKG_KOPYALA", _
                  "Arþiv veya stok sayfasý bulunamadý."
    End If

    If dst.ProtectContents Then
        Err.Raise vbObjectError + 742, "AKG_KOPYALA", _
                  "Hedef sayfa korumalý: " & dst.Name
    End If

    kaynakArsiv = (Left$(src.Name, 8) = "_AKLIVE_")
    hedefArsiv = (Left$(dst.Name, 8) = "_AKLIVE_")

    If kaynakArsiv = hedefArsiv Then
        Err.Raise vbObjectError + 743, "AKG_KOPYALA", _
                  "Stok ve arþiv eþleþmesi geçersiz."
    End If

    If kaynakArsiv Then
        il = CStr(src.Range("B1").Value2)
        tur = CStr(src.Range("C1").Value2)

        If CStr(src.Range("A1").Value2) <> "AKSORA_CANLI_ARSIV_V1" Then
            Err.Raise vbObjectError + 744, "AKG_KOPYALA", _
                      "Arþiv kimliði geçersiz: " & src.Name
        End If

        If src.Name <> "_AKLIVE_" & il & "_" & tur Then
            Err.Raise vbObjectError + 745, "AKG_KOPYALA", _
                      "Arþiv adý ve kimliði uyuþmuyor."
        End If

        bilgi = src.Range("A3:A5").Value2
    Else
        il = CStr(dst.Range("B1").Value2)
        tur = CStr(dst.Range("C1").Value2)

        If CStr(dst.Range("A1").Value2) <> "AKSORA_CANLI_ARSIV_V1" Then
            Err.Raise vbObjectError + 746, "AKG_KOPYALA", _
                      "Hedef arþiv kimliði geçersiz."
        End If

        If dst.Name <> "_AKLIVE_" & il & "_" & tur Then
            Err.Raise vbObjectError + 747, "AKG_KOPYALA", _
                      "Hedef arþiv adý ve kimliði uyuþmuyor."
        End If

        bilgi = src.Range("C2:C4").Value2
    End If

    cdp = AKSORA_IL_TESIS(il, "CDP")
    epp = AKSORA_IL_TESIS(il, "EPP")

    Select Case tur
        Case "CDP", "EPP", "TUM", "ILC", "YUK"
        Case Else
            Err.Raise vbObjectError + 748, "AKG_KOPYALA", _
                      "Arþiv türü geçersiz."
    End Select

    son = FIRST_DATA - 1
    hedefSon = FIRST_DATA - 1

    For c = 1 To 12
        r = src.Cells(src.Rows.Count, c).End(xlUp).Row
        If r > son Then son = r

        r = dst.Cells(dst.Rows.Count, c).End(xlUp).Row
        If r > hedefSon Then hedefSon = r
    Next c

    If son > 250000 Or hedefSon > 250000 Then
        Err.Raise vbObjectError + 749, "AKG_KOPYALA", _
                  "Beklenmeyen veri aralýðý."
    End If

    tamToplam = CDec(0)
    If son >= FIRST_DATA Then
        kontrol = src.Range("A" & FIRST_DATA & ":L" & son).Value2
        For r = 1 To UBound(kontrol, 1)
        If IsError(kontrol(r, 7)) Or IsError(kontrol(r, 1)) Then GoTo YANLIS_IL
        If Len(Trim$(CStr(kontrol(r, 7)))) > 0 Then
            adet = adet + 1
            If IsError(kontrol(r, 11)) Then GoTo YANLIS_IL
            If Not IsNumeric(kontrol(r, 11)) Then GoTo YANLIS_IL
            tamToplam = tamToplam + CDec(kontrol(r, 11))
            tesis = Trim$(CStr(kontrol(r, 1)))

            Select Case tur
                Case "CDP"
                    If tesis <> cdp Then GoTo YANLIS_IL
                Case "EPP"
                    If tesis <> epp Then GoTo YANLIS_IL
                Case Else
                    If tesis <> cdp And tesis <> epp Then GoTo YANLIS_IL
            End Select
        End If
    Next r
    End If

    If adet = 0 Then
        Err.Raise vbObjectError + 750, "AKG_KOPYALA", _
                  "Kaynak boþ; stoklar deðiþtirilmedi: " & src.Name
    End If

    If hedefSon >= FIRST_DATA Then
        dst.Range("A" & FIRST_DATA & ":L" & hedefSon).ClearContents
    End If

    dst.Range("A" & FIRST_DATA & ":L" & son).Value2 = kontrol

    If hedefArsiv Then
        ' C3/C4 filtrelenmis KPI olabilir. Arsiv tum satirlari tasir.
        bilgi(2, 1) = adet
        bilgi(3, 1) = CDbl(tamToplam)
        dst.Range("A3:A5").Value2 = bilgi
        dst.Range("D1").Value = Now
    Else
        dst.Range("C2:C4").Value2 = bilgi
    End If

    Exit Sub

YANLIS_IL:
    Err.Raise vbObjectError + 751, "AKG_KOPYALA", _
              "Baþka ile ait stok bulundu: " & src.Name & _
              ", satýr " & r
End Sub
Private Sub AKG_HEDEF_ARSIV_KONTROL(ByVal il As String)

    Dim tipler As Variant
    Dim kaynaklar As Variant

    Dim ar As Worksheet
    Dim ws As Worksheet

    Dim i As Long
    Dim r As Long
    Dim son As Long
    Dim adet As Long

    Dim toplam As Double
    Dim tesis As String

    Dim cdp As String
    Dim epp As String


    tipler = Array("CDP", "EPP", "TUM", "ILC", "YUK")

    kaynaklar = Array( _
        "CDP_STOK", _
        "EPP_STOK", _
        "STOKLAR", _
        "ILCE_STOK", _
        "YUKLENICI_STOK" _
    )


    cdp = AKSORA_IL_TESIS(il, "CDP")
    epp = AKSORA_IL_TESIS(il, "EPP")


    For i = 0 To 4

        Set ar = AKG_SAYFA( _
            "_AKLIVE_" & il & "_" & tipler(i) _
        )

        Set ws = AKG_SAYFA(CStr(kaynaklar(i)))


        If ar Is Nothing Or ws Is Nothing Then

            Err.Raise _
                vbObjectError + 760, _
                "AKG_HEDEF_ARSIV_KONTROL", _
                "Stok veya hedef arþiv sayfasý eksik."

        End If


        If ar.ProtectContents Or ws.ProtectContents Then

            Err.Raise _
                vbObjectError + 761, _
                "AKG_HEDEF_ARSIV_KONTROL", _
                "Stok veya arþiv veri alaný korumalý."

        End If


        If CStr(ar.Range("A1").Value2) <> _
                "AKSORA_CANLI_ARSIV_V1" Or _
           CStr(ar.Range("B1").Value2) <> il Or _
           CStr(ar.Range("C1").Value2) <> _
                CStr(tipler(i)) Then

            Err.Raise _
                vbObjectError + 762, _
                "AKG_HEDEF_ARSIV_KONTROL", _
                "Hedef arþiv kimliði uyuþmuyor: " & _
                ar.Name

        End If


        If Not IsNumeric(ar.Range("A3").Value2) Then _
            GoTo BILGI_HATALI

        If CDbl(ar.Range("A3").Value2) <= 0 Then _
            GoTo BILGI_HATALI

        If Not IsNumeric(ar.Range("A4").Value2) Then _
            GoTo BILGI_HATALI

        If Not IsNumeric(ar.Range("A5").Value2) Then _
            GoTo BILGI_HATALI


        son = _
            ar.Cells( _
                ar.Rows.Count, _
                "G" _
            ).End(xlUp).Row


        If son > 250000 Then _
            GoTo BILGI_HATALI


        adet = 0
        toplam = 0


        For r = 7 To son

            If Len( _
                Trim$( _
                    CStr( _
                        ar.Cells(r, "G").Value2 _
                    ) _
                ) _
            ) > 0 Then


                adet = adet + 1


                tesis = _
                    Trim$( _
                        CStr( _
                            ar.Cells(r, "A").Value2 _
                        ) _
                    )


                Select Case CStr(tipler(i))

                    Case "CDP"

                        If tesis <> cdp Then _
                            GoTo BILGI_HATALI


                    Case "EPP"

                        If tesis <> epp Then _
                            GoTo BILGI_HATALI


                    Case Else

                        If tesis <> cdp _
                           And tesis <> epp Then

                            GoTo BILGI_HATALI

                        End If

                End Select


                If Not IsNumeric( _
                    ar.Cells(r, "K").Value2 _
                ) Then

                    GoTo BILGI_HATALI

                End If


                toplam = _
                    toplam + _
                    CDbl( _
                        ar.Cells(r, "K").Value2 _
                    )

            End If

        Next r


        If adet = 0 Then _
            GoTo BILGI_HATALI


        If CDbl(ar.Range("A4").Value2) <> adet Then _
            GoTo BILGI_HATALI


        If Abs( _
            CDbl(ar.Range("A5").Value2) - toplam _
        ) > 0.02 Then

            GoTo BILGI_HATALI

        End If

    Next i


    Exit Sub


BILGI_HATALI:

    Err.Raise _
        vbObjectError + 763, _
        "AKG_HEDEF_ARSIV_KONTROL", _
        "Hedef arþiv boþ veya tutarsýz: " & _
        ar.Name & _
        vbCrLf & _
        "Ýl geçiþi yapýlmadý."

End Sub

Private Sub AKG_RAPORLARI_YENILE()
    Dim stoklar As Variant, raporlar As Variant
    Dim ws As Worksheet, rp As Worksheet
    Dim veri As Variant
    Dim i As Long, n As Long

    stoklar = Array("CDP_STOK", "EPP_STOK", "STOKLAR", _
                    "ILCE_STOK", "YUKLENICI_STOK")

    raporlar = Array("CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", _
                     "ILCE_RAPORLAR", "YUKLENICI_RAPOR")

    For i = 0 To 4
        Set ws = AKG_SAYFA(CStr(stoklar(i)))
        Set rp = AKG_SAYFA(CStr(raporlar(i)))

        If ws Is Nothing Or rp Is Nothing Then
            Err.Raise vbObjectError + 820, _
                      "AKG_RAPORLARI_YENILE", _
                      "Stok veya rapor sayfasý eksik."
        End If

        If rp.ProtectContents Then
            Err.Raise vbObjectError + 821, _
                      "AKG_RAPORLARI_YENILE", _
                      "Rapor veri alaný korumalý: " & rp.Name
        End If

        n = ws.Cells(ws.Rows.Count, "G").End(xlUp).Row - 6

        If n <= 0 Then
            Err.Raise vbObjectError + 822, _
                      "AKG_RAPORLARI_YENILE", _
                      "Rapor kaynaðý boþ: " & ws.Name
        End If

        Select Case i
            Case 0
                CDP_V2_RAPOR_OLUSTUR ws, rp, n
            Case 1
                EPP_V2_RAPOR_OLUSTUR ws, rp, n
            Case 2
                veri = ws.Range("A7").Resize(n, 12).Value2
                BIR_V2_RAPOR_OLUSTUR rp, veri
            Case 3
                ILCE_V2_RAPOR_OLUSTUR ws, rp, n
            Case 4
                YUK_V2_RAPOR_OLUSTUR ws, rp, n
        End Select

        rp.Range("B3").Value2 = ws.Range("C2").Value2
        RAPOR_GORUNUM_TEK rp.Name
    Next i

    For i = 0 To 4
        Set rp = AKG_SAYFA(CStr(raporlar(i)))
        AKSORA_RAPOR_AMBLEM_HIZALA rp.Name
    Next i
    AKFIX_STOKLARI_YENILE
    AKSORA_STOK_IL_BASLIKLARI_YENILE
       AKSORA_KARS_IL_GECIS_YENILE
    PANEL_YENILE
    AKSORA_V10_TARIH_GOSTER
End Sub


Public Sub AKSORA_V11_CANLI_YENILE(ByVal il As String)
    Dim kaynak As Variant, etiket As Variant, rapor As Variant, i As Long, n As Long
    Dim ws(0 To 4) As Worksheet, eski(0 To 4) As Variant, son(0 To 4) As Long
    Dim yaziliyor As Boolean, Hata As String, geri As String
    Dim meta(0 To 4) As Variant, j As Long
    On Error GoTo Hata
    If AKSORA_CANLI_STOK_IL() <> il Then Err.Raise vbObjectError + 1101, , "Canlý il uyuþmuyor."
    AKG_HEDEF_ARSIV_KONTROL il
    kaynak = Array("CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK")
    etiket = Array("CDP", "EPP", "TUM", "ILC", "YUK")
    rapor = Array("CDP_RAPOR", "EPP_RAPOR", "RAPORLAR", "ILCE_RAPORLAR", "YUKLENICI_RAPOR")
    For i = 0 To 4
        Set ws(i) = AKG_SAYFA(CStr(kaynak(i)))
        If ws(i).ProtectContents Or AKG_SAYFA(CStr(rapor(i))).ProtectContents Then
            Err.Raise vbObjectError + 1106, , "Canlý stok veya rapor korumalý."
        End If
        son(i) = ws(i).Cells(ws(i).Rows.Count, "G").End(xlUp).Row
        If son(i) < 7 Then son(i) = 7
        eski(i) = ws(i).Range("A7:L" & son(i)).Value2
        meta(i) = ws(i).Range("C2:C4").Value2
    Next i
    yaziliyor = True
    For i = 0 To 4
        AKG_KOPYALA AKG_SAYFA("_AKLIVE_" & il & "_" & etiket(i)), ws(i)
    Next i
    AKG_RAPORLARI_YENILE
    AKSORA_V11_BICIMLER
    Exit Sub
Hata:
    Hata = Err.Description
    If yaziliyor Then
        On Error Resume Next
        For i = 0 To 4
            Err.Clear
            n = ws(i).Cells(ws(i).Rows.Count, "G").End(xlUp).Row
            ws(i).Range("A7:L" & application.Max(7, n, son(i))).ClearContents
            ws(i).Range("A7:L" & son(i)).Value2 = eski(i)
            For j = 1 To 3: ws(i).Cells(j + 1, "C").Value2 = meta(i)(j, 1): Next j
            If Err.Number <> 0 Then geri = geri & vbCrLf & "Canlý geri alma baþarýsýz: " & ws(i).Name
        Next i
        If Len(geri) = 0 Then
            Err.Clear
            AKG_RAPORLARI_YENILE
            If Err.Number <> 0 Then geri = "Önceki rapor görünümü yenilenemedi."
        End If
        On Error GoTo 0
    End If
    Err.Raise vbObjectError + 1107, "AKSORA_V11_CANLI_YENILE", Hata & vbCrLf & geri
End Sub

Public Sub AKSORA_CANLI_ARSIV_BILGISINI_ONAR()
    Dim il As String, kaynak As Variant, tip As Variant, i As Long
    Dim ev As Boolean, sc As Boolean, no As Long, aciklama As String
    il = AKSORA_CANLI_STOK_IL()
    kaynak = Array("CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK")
    tip = Array("CDP", "EPP", "TUM", "ILC", "YUK")
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    application.EnableEvents = False: application.ScreenUpdating = False
    For i = 0 To 4
        AKG_KOPYALA AKG_SAYFA(CStr(kaynak(i))), AKG_SAYFA("_AKLIVE_" & il & "_" & CStr(tip(i)))
    Next i
    AKG_HEDEF_ARSIV_KONTROL il
Cikis:
    application.EnableEvents = ev: application.ScreenUpdating = sc
    If no <> 0 Then Err.Raise no, "AKSORA_CANLI_ARSIV_BILGISINI_ONAR", aciklama
    Exit Sub
Hata:
    no = Err.Number: aciklama = Err.Description: Resume Cikis
End Sub
