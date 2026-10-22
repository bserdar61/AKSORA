Attribute VB_Name = "modAKSORA_ARSIV"
Option Explicit

' Aktif ilin bes stok sayfasini ayri VeryHidden canli arsivlere kaydeder.
' Ilk _AK_* arsivleri, 180 gunluk gecmis, grafikler ve stok sayfalari korunur.
' Il degistirmez; mevcut giris/kilit akisini bozmaz.
Public Sub AKSORA_IL_CANLI_ARSIV_KAYDET()
    Dim il As String, cdp As String, epp As String
    Dim kaynak As Variant, tip As Variant, i As Long, r As Long, son As Long
    Dim ws As Worksheet, hedef As Worksheet, ad As String, tesis As String
    Dim HATALI As Long, sayac As Long, eskiEkran As Boolean, eskiOlay As Boolean
    Dim uygulamaDegisti As Boolean, hataMesaji As String
    On Error GoTo Hata
    il = AKSORA_CANLI_STOK_IL()
    cdp = AKSORA_IL_TESIS(il, "CDP")
    epp = AKSORA_IL_TESIS(il, "EPP")
    kaynak = Array("CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK")
    tip = Array("CDP", "EPP", "TUM", "ILC", "YUK")

    ' Hicbir seyi degistirmeden once bes sayfayi ve tesisleri kontrol et.
    For i = 0 To 4
        Set ws = AK_CANLI_SAYFA(CStr(kaynak(i)))
        If ws Is Nothing Then Err.Raise vbObjectError + 950, , "Kaynak sayfa yok: " & CStr(kaynak(i))
        son = AKSORA_SON_SATIR(ws, "A")
        If son < 7 Then Err.Raise vbObjectError + 951, , "Stok sayfasi bos: " & ws.Name
        For r = 7 To son
            If Len(Trim$(CStr(ws.Cells(r, "G").Value2))) > 0 Then
                tesis = Trim$(CStr(ws.Cells(r, "A").Value2))
                If CStr(tip(i)) = "CDP" Then
                    If tesis <> cdp Then HATALI = HATALI + 1
                ElseIf CStr(tip(i)) = "EPP" Then
                    If tesis <> epp Then HATALI = HATALI + 1
                Else
                    If tesis <> cdp And tesis <> epp Then HATALI = HATALI + 1
                End If
            End If
        Next r
    Next i
    If HATALI > 0 Then Err.Raise vbObjectError + 952, , "Baska ile ait " & HATALI & " stok satiri var; arsiv alinmadi."

    eskiEkran = application.ScreenUpdating
    eskiOlay = application.EnableEvents
    uygulamaDegisti = True
    application.ScreenUpdating = False
    application.EnableEvents = False
    For i = 0 To 4
        Set ws = AK_CANLI_SAYFA(CStr(kaynak(i)))
        ad = "_AKLIVE_" & il & "_" & CStr(tip(i))
        Set hedef = Nothing
        On Error Resume Next
        Set hedef = ThisWorkbook.Worksheets(ad)
        On Error GoTo Hata
        If hedef Is Nothing Then
            Set hedef = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            hedef.Name = ad
        Else
            If CStr(hedef.Range("A1").Value2) <> "AKSORA_CANLI_ARSIV_V1" Or _
               CStr(hedef.Range("B1").Value2) <> il Or _
               CStr(hedef.Range("C1").Value2) <> CStr(tip(i)) Then
                Err.Raise vbObjectError + 953, , "Arsiv kimligi uyusmuyor: " & ad
            End If
            hedef.Range("A1:L" & AKSORA_SON_SATIR(hedef, "A")).ClearContents
        End If
        son = AKSORA_SON_SATIR(ws, "A")
        hedef.Range("A1").Value2 = "AKSORA_CANLI_ARSIV_V1"
        hedef.Range("B1").Value2 = il
        hedef.Range("C1").Value2 = CStr(tip(i))
        hedef.Range("D1").Value = Now
        hedef.Range("A3:A5").Value2 = ws.Range("C2:C4").Value2
        hedef.Range("A7:L" & son).Value2 = ws.Range("A7:L" & son).Value2
        hedef.Visible = xlSheetVeryHidden
        sayac = sayac + 1
    Next i
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    MsgBox il & ": " & sayac & " canli stok arsivi kaydedildi." & vbCrLf & _
           "Ilk arsiv, stoklar, grafikler ve 180 gunluk gecmis korundu.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    hataMesaji = Err.Description
    On Error Resume Next
    If uygulamaDegisti Then
        application.EnableEvents = eskiOlay
        application.ScreenUpdating = eskiEkran
    End If
    MsgBox "Canli arsiv durduruldu: " & hataMesaji, vbExclamation, "AKSORA"
End Sub

Private Function AK_CANLI_SAYFA(ByVal asciiAd As String) As Worksheet
    Dim ws As Worksheet, ad As String
    For Each ws In ThisWorkbook.Worksheets
        ad = ws.Name
        ad = Replace(ad, ChrW$(304), "I")
        ad = Replace(ad, ChrW$(305), "I")
        ad = Replace(ad, ChrW$(199), "C")
        ad = Replace(ad, ChrW$(231), "C")
        ad = Replace(ad, ChrW$(220), "U")
        ad = Replace(ad, ChrW$(252), "U")
        If StrComp(UCase$(ad), asciiAd, vbBinaryCompare) = 0 Then
            Set AK_CANLI_SAYFA = ws
            Exit Function
        End If
    Next ws
End Function

' Il bazli ilk stok arsivi. Yalnizca yeni VeryHidden sayfalar olusturur.
' Mevcut stok, rapor, sekil ve _STOK_HAFIZA_* sayfalarina dokunmaz.
' Ilk arsiv korunur: ayni il/sayfa icin ikinci kez ustune yazmaz.
Public Sub AKSORA_IL_STOK_ILK_ARSIV()
    Dim il As String, cdp As String, epp As String
    Dim kaynaklar As Variant, kodlar As Variant, i As Long
    Dim ws As Worksheet, arsiv As Worksheet, ad As String
    Dim son As Long, r As Long, tesis As String, HATALI As Long
    Dim kaydedilen As Long, mevcut As Long
    Dim eskiEkran As Boolean, eskiOlay As Boolean
    Dim hataNo As Long, hataAciklama As String
    On Error GoTo Hata
        il = AKSORA_CANLI_STOK_IL()
    cdp = AKSORA_IL_TESIS(il, "CDP")
    epp = AKSORA_IL_TESIS(il, "EPP")
    kaynaklar = Array("CDP_STOK", "EPP_STOK", "STOKLAR", "ILCE_STOK", "YUKLENICI_STOK")
    kodlar = Array("CDP", "EPP", "TUM", "ILC", "YUK")
    ' Tum kaynaklari ve tesis kodlarini degisiklikten ONCE dogrula.
    For i = 0 To UBound(kaynaklar)
        Set ws = AKSORA_ARSIV_SAYFA(CStr(kaynaklar(i)))
        If ws Is Nothing Then Err.Raise vbObjectError + 920, , "Kaynak sayfa bulunamadi: " & CStr(kaynaklar(i))
        son = AKSORA_SON_SATIR(ws, "A")
        For r = 7 To son
            If Len(Trim$(CStr(ws.Cells(r, "G").Value2))) > 0 Then
                tesis = Trim$(CStr(ws.Cells(r, "A").Value2))
                If CStr(kodlar(i)) = "CDP" Then
                    If tesis <> cdp Then HATALI = HATALI + 1
                ElseIf CStr(kodlar(i)) = "EPP" Then
                    If tesis <> epp Then HATALI = HATALI + 1
                Else
                    If tesis <> cdp And tesis <> epp Then HATALI = HATALI + 1
                End If
            End If
        Next r
    Next i
    If HATALI > 0 Then Err.Raise vbObjectError + 921, , "Baska ile ait " & HATALI & " stok satiri var. Arsiv olusturulmadi."
    eskiEkran = application.ScreenUpdating
    eskiOlay = application.EnableEvents
    application.ScreenUpdating = False
    application.EnableEvents = False
    For i = 0 To UBound(kaynaklar)
        Set ws = AKSORA_ARSIV_SAYFA(CStr(kaynaklar(i)))
        ad = "_AK_" & il & "_" & CStr(kodlar(i))
        If AKSORA_ARSIV_VAR(ad) Then
            mevcut = mevcut + 1
        Else
            Set arsiv = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
            arsiv.Name = ad
            son = AKSORA_SON_SATIR(ws, "A")
            arsiv.Range("A1").Value2 = "AKSORA_IL_ARSIV_V2"
            arsiv.Range("B1").Value2 = il
            arsiv.Range("C1").Value2 = CStr(kodlar(i))
            arsiv.Range("D1").Value = Now
            arsiv.Range("A3:A5").Value2 = ws.Range("C2:C4").Value2
            If son >= 7 Then arsiv.Range("A7:L" & son).Value2 = ws.Range("A7:L" & son).Value2
            arsiv.Visible = xlSheetVeryHidden
            kaydedilen = kaydedilen + 1
        End If
    Next i
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    MsgBox il & ": " & kaydedilen & " ilk arsiv olusturuldu; " & mevcut & " mevcut arsiv korundu." & vbCrLf & _
           "Mevcut stoklar, grafikler ve 180 gunluk gecmis degismedi.", vbInformation, "AKSORA IL ARSIV"
    Exit Sub
Hata:
    hataNo = Err.Number: hataAciklama = Err.Description
    On Error Resume Next
    application.EnableEvents = eskiOlay
    application.ScreenUpdating = eskiEkran
    MsgBox "Arsiv durduruldu: " & hataAciklama, vbExclamation, "AKSORA IL ARSIV"
End Sub
Private Function AKSORA_ARSIV_VAR(ByVal ad As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(ad)
    On Error GoTo 0
    AKSORA_ARSIV_VAR = Not ws Is Nothing
End Function
Private Function AKSORA_ARSIV_SAYFA(ByVal asciiAd As String) As Worksheet
    Dim ws As Worksheet, ad As String
    For Each ws In ThisWorkbook.Worksheets
        ad = ws.Name
        ' Unicode kodlariyla eslestir: VBA modulunun dosya kodlamasindan bagimsiz.
        ad = Replace(ad, ChrW$(304), "I") ' I with dot
        ad = Replace(ad, ChrW$(305), "I") ' dotless i
        ad = Replace(ad, ChrW$(199), "C") ' C cedilla
        ad = Replace(ad, ChrW$(231), "C")
        ad = Replace(ad, ChrW$(220), "U") ' U diaeresis
        ad = Replace(ad, ChrW$(252), "U")
        ad = UCase$(ad)
        If StrComp(ad, asciiAd, vbBinaryCompare) = 0 Then
            Set AKSORA_ARSIV_SAYFA = ws
            Exit Function
        End If
    Next ws
End Function

' Bos il arsivlerini olusturur; SAP verisi cekmez ve il degistirmez.
' Artvin arsivleri, mevcut stoklar, raporlar, sekiller ve 180 gunluk gecmis korunur.
Public Sub AKSORA_DIGER_IL_ARSIV_HAZIRLA()
    Dim iller As Variant, turler As Variant, il As Variant, tur As Variant
    Dim ad As String, ws As Worksheet, mevcut As Worksheet
    Dim yeni As Long, eski As Long, olay As Boolean, ekran As Boolean
    Dim ayarDegisti As Boolean, Hata As String
    On Error GoTo Hata

    If AKSORA_AKTIF_IL() <> "ARTVIN" Then
        MsgBox "Hazirlik yalnizca ARTVIN aktifken yapilir. Hicbir sey degismedi.", vbExclamation, "AKSORA": Exit Sub
    End If
    iller = Array("RIZE", "TRABZON", "GUMUSHANE", "GIRESUN")
    turler = Array("CDP", "EPP", "TUM", "ILC", "YUK")
    If ThisWorkbook.ProtectStructure Then
        MsgBox "Calisma kitabi yapisi korumali. Islem yapilmadi.", vbExclamation, "AKSORA": Exit Sub
    End If

    ' Tum hedefleri onceden kontrol et: mevcut arsivleri asla silme/uzerine yazma.
    For Each il In iller
        For Each tur In turler
            ad = "_AKLIVE_" & CStr(il) & "_" & CStr(tur)
            Set mevcut = Nothing
            On Error Resume Next
            Set mevcut = ThisWorkbook.Worksheets(ad)
            On Error GoTo Hata
            If Not mevcut Is Nothing Then
                If CStr(mevcut.Range("A1").Value2) <> "AKSORA_CANLI_ARSIV_V1" Or _
                   CStr(mevcut.Range("B1").Value2) <> CStr(il) Or _
                   CStr(mevcut.Range("C1").Value2) <> CStr(tur) Then
                    Err.Raise vbObjectError + 991, , "Mevcut arsiv kimligi uyusmuyor: " & ad
                End If
                eski = eski + 1
            End If
        Next tur
    Next il

    olay = application.EnableEvents: ekran = application.ScreenUpdating
    ayarDegisti = True
    application.EnableEvents = False: application.ScreenUpdating = False
    For Each il In iller
        For Each tur In turler
            ad = "_AKLIVE_" & CStr(il) & "_" & CStr(tur)
            Set mevcut = Nothing
            On Error Resume Next
            Set mevcut = ThisWorkbook.Worksheets(ad)
            On Error GoTo Hata
            If mevcut Is Nothing Then
                Set ws = ThisWorkbook.Worksheets.Add(after:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
                ws.Name = ad
                ws.Range("A1").Value2 = "AKSORA_CANLI_ARSIV_V1"
                ws.Range("B1").Value2 = CStr(il)
                ws.Range("C1").Value2 = CStr(tur)
                ws.Range("D1").Value2 = "BOS - SAP GUNCELLEMESI BEKLENIYOR"
                ws.Visible = xlSheetVeryHidden
                yeni = yeni + 1
            End If
        Next tur
    Next il
    application.EnableEvents = olay: application.ScreenUpdating = ekran
    MsgBox "Diger iller: " & yeni & " bos canli arsiv olusturuldu; " & eski & " mevcut arsiv korundu." & vbCrLf & _
           "SAP verisi henuz yok. IL GECIS makrosunu simdilik calistirmayin.", vbInformation, "AKSORA"
    Exit Sub
Hata:
    Hata = Err.Description
    On Error Resume Next
    If ayarDegisti Then
        application.EnableEvents = olay: application.ScreenUpdating = ekran
    End If
    MsgBox "Arsiv hazirligi durduruldu: " & Hata, vbExclamation, "AKSORA"
End Sub
