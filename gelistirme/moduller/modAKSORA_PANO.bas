Attribute VB_Name = "modAKSORA_PANO"
Option Explicit

'==========================================================
' AKSORA v3.1  |  modAKSORA_PANO
'
' YONETICI PANOSU (WEB): Yonetim Merkezi'ndeki karne, yapilacak isler, uyarilar ve
' il karsilastirmasi tek bir HTML dosyasina yazilir (masaustu) ve tarayicida acilir.
' Dosya tek basina calisir (internet ya da ek dosya gerekmez): e-posta ile gonderilebilir,
' telefonda / tablette acilabilir, acik ve koyu temayi destekler.
'
' Sayfa sablonu modAKSORA_PANO_SABLON icindedir (gelistirme/pano/pano_sablon.html
' dosyasindan tools/aksora_vba.py "sablon" komutuyla uretilir; elle duzenlenmez).
'==========================================================

Private Const PANO_YER As String = "/*VERI*/null/*VERI*/"

' JSON metni: tirnak icinde, ozel karakterler kacisli.
Public Function AKSORA_PANO_METIN(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCrLf, "\n")
    s = Replace(s, vbCr, "\n")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, " ")
    s = Replace(s, "</", "<\/")
    AKSORA_PANO_METIN = """" & s & """"
End Function

' JSON sayisi: ondalik ayirac daima nokta; sayi degilse null.
Public Function AKSORA_PANO_SAYI(ByVal v As Variant) As String
    Dim t As String
    AKSORA_PANO_SAYI = "null"
    If IsEmpty(v) Then Exit Function
    If IsError(v) Then Exit Function
    If VarType(v) = vbString Then Exit Function
    If Not IsNumeric(v) Then Exit Function
    t = Trim$(Str$(CDbl(v)))
    If Left$(t, 1) = "." Then t = "0" & t
    If Left$(t, 2) = "-." Then t = "-0" & Mid$(t, 2)
    AKSORA_PANO_SAYI = t
End Function

Public Sub AKSORA_PANO_AC()
    Dim veri As String, sayfa As String, yol As String, akis As Object, no As Long, metin As String
    veri = AKSORA_KPI_PANO_JSON()
    If Len(veri) = 0 Then
        MsgBox "Pano için önce RAPORU YENİLE düğmesine basın.", vbInformation, "AKSORA | YÖNETİCİ PANOSU"
        Exit Sub
    End If
    On Error GoTo Hata
    sayfa = AKSORA_PANO_SABLON()
    If InStr(1, sayfa, PANO_YER, vbBinaryCompare) = 0 Then Err.Raise vbObjectError + 3101, , "Pano şablonu bozuk (veri yeri bulunamadı)."
    sayfa = Replace(sayfa, PANO_YER, veri, 1, 1, vbBinaryCompare)
    yol = AKSORA_MASAUSTU() & Application.PathSeparator & _
        AKSORA_DOSYA_ADI("AKSORA Yönetici Panosu - " & Format$(Now, "dd.mm.yyyy hh.nn")) & ".html"
    Set akis = CreateObject("ADODB.Stream")
    akis.Type = 2
    akis.Charset = "utf-8"
    akis.Open
    akis.WriteText sayfa
    akis.SaveToFile yol, 2
    akis.Close
    Set akis = Nothing
    AKSORA_GUNLUK_YAZ "YÖNETİCİ", "Yönetici panosu (web) oluşturuldu"
    Application.StatusBar = "Masaüstüne kaydedildi: " & yol
    On Error Resume Next
    CreateObject("WScript.Shell").Run """" & yol & """", 1, False
    If Err.Number <> 0 Then
        Err.Clear
        MsgBox "Pano masaüstüne kaydedildi ancak tarayıcı açılamadı." & vbCrLf & vbCrLf & yol, vbInformation, "AKSORA | YÖNETİCİ PANOSU"
    End If
    Exit Sub
Hata:
    no = Err.Number: metin = Err.Description
    On Error Resume Next
    If Not akis Is Nothing Then akis.Close
    On Error GoTo 0
    MsgBox "Yönetici panosu oluşturulamadı." & vbCrLf & vbCrLf & metin & " (" & no & ")", vbExclamation, "AKSORA"
End Sub

' Sinama (kurulum betigi cagirir). Donus: "OK" ya da basarisiz kontrolun adi.
Public Function AKSORA_PANO_SINAMA() As String
    Dim bos As Variant, s As String
    On Error GoTo Hata
    If AKSORA_PANO_SAYI(0.5) <> "0.5" Then AKSORA_PANO_SINAMA = "sayi-yarim": Exit Function
    If AKSORA_PANO_SAYI(-0.25) <> "-0.25" Then AKSORA_PANO_SINAMA = "sayi-eksi": Exit Function
    If AKSORA_PANO_SAYI(1389056399.4) <> "1389056399.4" Then AKSORA_PANO_SINAMA = "sayi-buyuk": Exit Function
    If AKSORA_PANO_SAYI(bos) <> "null" Then AKSORA_PANO_SINAMA = "sayi-bos": Exit Function
    If AKSORA_PANO_SAYI("12") <> "null" Then AKSORA_PANO_SINAMA = "sayi-metin": Exit Function
    If AKSORA_PANO_METIN("a""b\c") <> """a\""b\\c""" Then AKSORA_PANO_SINAMA = "metin-kacis": Exit Function
    If AKSORA_PANO_METIN("x" & vbLf & "y") <> """x\ny""" Then AKSORA_PANO_SINAMA = "metin-satir": Exit Function
    s = AKSORA_PANO_SABLON()
    If Len(s) < 5000 Then AKSORA_PANO_SINAMA = "sablon-kisa": Exit Function
    If InStr(1, s, PANO_YER, vbBinaryCompare) = 0 Then AKSORA_PANO_SINAMA = "sablon-veri-yeri": Exit Function
    If Right$(RTrim$(Replace(s, vbLf, "")), 7) <> "</html>" Then AKSORA_PANO_SINAMA = "sablon-son": Exit Function
    AKSORA_PANO_SINAMA = "OK"
    Exit Function
Hata:
    AKSORA_PANO_SINAMA = "beklenmeyen hata: " & Err.Description
End Function
