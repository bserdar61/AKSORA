Attribute VB_Name = "modKS_YUKLEYICI"
Option Explicit

' Tek modül içe aktarýlýr, gerisini bu yükleyici yapar:
'   1) Seçilen klasördeki tüm modKS_*.bas modüllerini içe aktarýr (varsa eskisini deðiþtirir).
'   2) ThisWorkbook'a açýlýþ kodunu (Workbook_Open -> KS_ACILIS) ekler.
'   3) KALITESORA_KUR kurulumunu baþlatýr.
' Gerekli ayar: Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarlarý > Makro Ayarlarý >
'               "VBA proje nesne modeline eriþime güven" iþaretli olmalý.

Public Sub KALITESORA_YUKLE()
    Dim proje As Object, bilesen As Object, klasor As String, f As String, ad As String, adet As Long
    Dim cm As Object, sep As String, adim As String, dlg As Object, metin As String
    On Error GoTo Hata

    adim = "VBA projesine eriþim"
    Set proje = ThisWorkbook.VBProject
    adet = proje.VBComponents.Count
    adet = 0

    adim = "Klasör seçimi"
    sep = Application.PathSeparator
    Set dlg = Application.FileDialog(4)   ' msoFileDialogFolderPicker
    dlg.Title = "KALÝTESORA .bas dosyalarýnýn bulunduðu klasörü seçin"
    If Len(ThisWorkbook.Path) > 0 Then dlg.InitialFileName = ThisWorkbook.Path & sep
    If dlg.Show <> -1 Then Exit Sub
    klasor = dlg.SelectedItems(1)
    If Right$(klasor, 1) = sep Then klasor = Left$(klasor, Len(klasor) - 1)
    If Len(Dir$(klasor & sep & "modKS_CEKIRDEK.bas")) = 0 Then
        MsgBox "Seçilen klasörde modKS_CEKIRDEK.bas bulunamadý." & vbCrLf & klasor, vbExclamation, "KALÝTESORA"
        Exit Sub
    End If

    f = Dir$(klasor & sep & "modKS_*.bas")
    Do While Len(f) > 0
        ad = Left$(f, Len(f) - 4)
        If StrComp(ad, "modKS_YUKLEYICI", vbTextCompare) <> 0 Then
            adim = "Ýçe aktarma: " & f
            Set bilesen = Nothing
            On Error Resume Next
            Set bilesen = proje.VBComponents(ad)
            On Error GoTo Hata
            If Not bilesen Is Nothing Then proje.VBComponents.Remove bilesen
            proje.VBComponents.Import klasor & sep & f
            adet = adet + 1
        End If
        f = Dir$()
    Loop

    adim = "ThisWorkbook açýlýþ kodu"
    Set cm = proje.VBComponents(ThisWorkbook.CodeName).CodeModule
    If cm.CountOfLines > 0 Then metin = cm.Lines(1, cm.CountOfLines)
    If InStr(1, metin, "KS_ACILIS", vbTextCompare) = 0 Then
        If InStr(1, metin, "Workbook_Open", vbTextCompare) > 0 Then
            MsgBox "ThisWorkbook içinde baþka bir Workbook_Open var; açýlýþ kodu eklenmedi." & vbCrLf & _
                "Ýçine elle 'KS_ACILIS' satýrýný ekleyin.", vbExclamation, "KALÝTESORA"
        Else
            cm.AddFromString "Private Sub Workbook_Open()" & vbCrLf & _
                "    KS_ACILIS" & vbCrLf & "End Sub"
        End If
    End If
    If InStr(1, metin, "KS_DEGISIKLIK", vbTextCompare) = 0 Then
        cm.AddFromString "Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)" & vbCrLf & _
            "    KS_DEGISIKLIK Sh, Target" & vbCrLf & "End Sub"
    End If

    MsgBox adet & " modül içe aktarýldý. Þimdi kurulum baþlayacak.", vbInformation, "KALÝTESORA"
    Application.Run "'" & ThisWorkbook.Name & "'!KALITESORA_KUR"
    Exit Sub
Hata:
    MsgBox "Yükleme durdu." & vbCrLf & "Adým: " & adim & vbCrLf & "Hata: " & Err.Description & vbCrLf & vbCrLf & _
        "VBA eriþim hatasýysa: Dosya > Seçenekler > Güven Merkezi > Güven Merkezi Ayarlarý > Makro Ayarlarý > " & _
        "'VBA proje nesne modeline eriþime güven' kutusunu iþaretleyin.", vbExclamation, "KALÝTESORA"
End Sub
