VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sayfa12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit



Private mKarsFiltreImza As String
Private mKarsTakipAktif As Boolean
Private mKarsZaman As Date
Private mKarsMakro As String
Private mKarsPlanli As Boolean

Public Sub STOK_KARS_V13_TAKIP_AC()
    STOK_KARS_FILTRE_TAKIP_DURDUR
    mKarsFiltreImza = STOK_KARS_FILTRE_IMZA()
    mKarsTakipAktif = True
    STOK_KARS_FILTRE_TAKIP_BASLAT
End Sub

Private Sub Worksheet_Activate()
    If AKSORA_V13_ACILIS Or AKSORA_V13_GECIS Or AKSORA_V11_MESGUL Then Exit Sub

    STOK_KARS_FILTRE_TAKIP_DURDUR

    mKarsFiltreImza = STOK_KARS_FILTRE_IMZA()
    mKarsTakipAktif = True

    STOK_KARS_FILTRE_TAKIP_BASLAT

End Sub

Private Sub Worksheet_Deactivate()

    STOK_KARS_FILTRE_TAKIP_DURDUR

End Sub

Public Sub STOK_KARS_FILTRE_TAKIP_DURDUR()

    mKarsTakipAktif = False

    If mKarsPlanli Then
        On Error Resume Next
        application.OnTime EarliestTime:=mKarsZaman, _
            Procedure:=mKarsMakro, Schedule:=False
        On Error GoTo 0
    End If

    mKarsPlanli = False

End Sub

Public Sub STOK_KARS_FILTRE_TAKIP_BASLAT()

    Dim yeniImza As String
    Dim eskiOlay As Boolean

    eskiOlay = application.EnableEvents
    On Error GoTo Hata

    If mKarsPlanli Then
        On Error Resume Next
        application.OnTime EarliestTime:=mKarsZaman, _
            Procedure:=mKarsMakro, Schedule:=False
        On Error GoTo Hata
        mKarsPlanli = False
    End If

    If Not mKarsTakipAktif Then Exit Sub

    If Me.Visible <> xlSheetVisible Then
        STOK_KARS_FILTRE_TAKIP_DURDUR
        Exit Sub
    End If

    If Not (Me Is ActiveSheet) Then
        STOK_KARS_FILTRE_TAKIP_DURDUR
        Exit Sub
    End If

    yeniImza = STOK_KARS_FILTRE_IMZA()

    If yeniImza <> mKarsFiltreImza Then
        application.EnableEvents = False
        STOK_KARS_KPI_GUNCELLE
        AKFIX_EN_USTE Me
        application.EnableEvents = eskiOlay
        mKarsFiltreImza = yeniImza
    End If

    mKarsZaman = Now + TimeSerial(0, 0, 1)
    mKarsMakro = "'" & Replace(ThisWorkbook.Name, "'", "''") & _
                "'!STOK_KARS_FILTRE_TAKIP_TICK"

    application.OnTime EarliestTime:=mKarsZaman, _
        Procedure:=mKarsMakro, Schedule:=True

    mKarsPlanli = True
    Exit Sub

Hata:
    application.EnableEvents = eskiOlay
    STOK_KARS_FILTRE_TAKIP_DURDUR

End Sub

Private Function STOK_KARS_FILTRE_IMZA() As String

    Dim sonSatir As Long
    Dim r As Long
    Dim imza As String

    sonSatir = Me.Cells(Me.Rows.Count, "D").End(xlUp).Row

    If sonSatir < 16 Then
        STOK_KARS_FILTRE_IMZA = "0"
        Exit Function
    End If

    For r = 16 To sonSatir
        If Not Me.Rows(r).Hidden Then
            imza = imza & "|" & CStr(r)
        End If
    Next r

    STOK_KARS_FILTRE_IMZA = imza

End Function




