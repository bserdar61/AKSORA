Attribute VB_Name = "modAKSORA_VBA_AKTAR"
Option Explicit

Public Sub AKSORA_VBA_DISARI_AKTAR()

    Dim proje As Object, bilesen As Object
    Dim fso As Object, kabuk As Object
    Dim masaustu As String, temel As String, klasor As String
    Dim uzanti As String, adim As String
    Dim adet As Long, sira As Long

    On Error GoTo Hata

    adim = "VBA projesine eriþim"
    Set proje = ThisWorkbook.VBProject

    If proje.Protection <> 0 Then
        MsgBox "VBA projesi kilitli. Önce proje kilidini açýn.", _
               vbExclamation, "AKSORA"
        Exit Sub
    End If

    ' Baþlamadan proje eriþimini doðrula.
    adet = proje.VBComponents.Count
    adet = 0

    adim = "Masaüstünde klasör oluþturma"
    Set kabuk = CreateObject("WScript.Shell")
    Set fso = CreateObject("Scripting.FileSystemObject")

    masaustu = kabuk.SpecialFolders("Desktop")
    temel = fso.BuildPath(masaustu, _
            "AKSORA_VBA_" & Format$(Now, "yyyymmdd_hhnnss"))

    klasor = temel
    sira = 1

    Do While fso.FolderExists(klasor) Or fso.FileExists(klasor)
        klasor = temel & "_" & CStr(sira)
        sira = sira + 1
    Loop

    fso.CreateFolder klasor

    For Each bilesen In proje.VBComponents

        adim = "Dýþarý aktarma: " & bilesen.Name

        Select Case bilesen.Type
            Case 1
                uzanti = ".bas"  ' Standart modül
            Case 2, 100
                uzanti = ".cls"  ' Sýnýf, sayfa ve çalýþma kitabý
            Case 3
                uzanti = ".frm"  ' Kullanýcý formu
            Case Else
                Err.Raise vbObjectError + 710, , _
                    "Desteklenmeyen bileþen türü: " & bilesen.Type
        End Select

        bilesen.Export fso.BuildPath(klasor, _
                                    bilesen.Name & uzanti)
        adet = adet + 1

    Next bilesen

    MsgBox adet & " VBA bileþeni dýþarý aktarýldý." & _
           vbCrLf & vbCrLf & klasor & vbCrLf & vbCrLf & _
           "Bu klasörün tamamýný ZIP yapýp gönderin.", _
           vbInformation, "AKSORA"

    Exit Sub

Hata:
    MsgBox "Aktarým tamamlanamadý." & vbCrLf & _
           "Adým: " & adim & vbCrLf & _
           "Hata: " & Err.Description & vbCrLf & _
           "Aktarýlan bileþen: " & adet & vbCrLf & _
           "Klasör: " & klasor, _
           vbExclamation, "AKSORA"

End Sub

