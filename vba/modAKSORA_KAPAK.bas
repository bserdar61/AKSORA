Attribute VB_Name = "modAKSORA_KAPAK"
Option Explicit

'==========================================================
' AKSORA - KAPAK / GIRIS TEK MODUL
' V5 + V8 BIRLESTIRILMIS YAPI
'==========================================================

Public AKSORA_GIRIS_ACIK As Boolean

Private Const KAPAK_SAYFA As String = "KAPAK"
Private Const PRE As String = "AKSORA_V4_"

Private mResizing As Boolean


'==========================================================
' KURULUM / DUZENLEME
'==========================================================
Public Sub AKSORA_V8_KUR()

    On Error GoTo Hata

    AKSORA_V8_DUZENLE

    MsgBox "AKSORA kapak hazir.", vbInformation, "AKSORA"

    Exit Sub

Hata:

    MsgBox "Kapak kurulum hatasi: " & Err.Description, _
           vbCritical, "AKSORA"

End Sub


'==========================================================
' KAPAK DUZENLE
'==========================================================
Public Sub AKSORA_V8_DUZENLE()

    Dim ws As Worksheet
    Dim sh As Shape
    Dim i As Long

    Dim ev As Boolean
    Dim sc As Boolean
    Dim chosen As String

    On Error GoTo Hata

    If mResizing Then Exit Sub

    mResizing = True

    ev = application.EnableEvents
    sc = application.ScreenUpdating

    application.EnableEvents = False
    application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets(KAPAK_SAYFA)

    chosen = UCase$(Trim$(CStr(ws.Range("AZ1").Value)))

    If Len(chosen) = 0 Then chosen = "ARTVIN"


    '------------------------------------------------------
    ' ESKI SLOGAN NESNELERI VARSA TEMIZLE
    '------------------------------------------------------
    On Error Resume Next

    ws.Shapes(PRE & "SLOGAN_BG").Delete
    ws.Shapes(PRE & "PHOTO_SLOGAN").Delete

    On Error GoTo Hata


    '------------------------------------------------------
    ' IL BUTONLARI
    '------------------------------------------------------
    For i = 1 To 5

        Set sh = ws.Shapes(PRE & "IL_" & i)

        sh.OnAction = _
            "'" & Replace(ThisWorkbook.Name, "'", "''") & _
            "'!AKSORA_V8_IL_SEC"

    Next i


    '------------------------------------------------------
    ' GIRIS BUTONU
    '------------------------------------------------------
    ws.Shapes(PRE & "ENTER").OnAction = _
        "'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!AKSORA_V8_GIRIS"


    '------------------------------------------------------
    ' ALT BILGI
    '------------------------------------------------------
    With ws.Shapes(PRE & "CREDIT")

        .TextFrame2.TextRange.Text = _
            "Version: v3.0  |  Design by Burhan SERDAR"

        .TextFrame2.TextRange.Font.Size = 10

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(181, 201, 216)

    End With


    AKSORA_V8_TARIH_YAZ ws

    AKSORA_V8_YERLESIM ws

    ws.Range("AZ1").Value = chosen

    AKSORA_V8_SECIM_GORUNUM ws

    AKSORA_V10_TARIH_GOSTER


    ws.Move Before:=ThisWorkbook.Worksheets(1)

    ws.Activate

    AKSORA_V8_KAPAK_GORUNUM


Cikis:

    application.EnableEvents = ev
    application.ScreenUpdating = sc

    mResizing = False

    Exit Sub


Hata:

    MsgBox "AKSORA kapak: " & Err.Description, _
           vbExclamation, "AKSORA"

    Resume Cikis

End Sub


'==========================================================
' SON GUNCELLEME TARIHI
'==========================================================
Private Sub AKSORA_V8_TARIH_YAZ(ByVal ws As Worksheet)

    Dim t As String
    Dim v As Variant

    Dim s As Shape
    Dim old As Shape
    Dim p As Shape

    Dim y As Double
    Dim ht As Double


    t = "Son stok guncellemesi: Kayit bulunamadi"


    On Error Resume Next


    v = ThisWorkbook.Worksheets("STOKLAR").Range("C2").Value


    If IsDate(v) Then

        t = "Son stok guncellemesi: " & _
            Format$(CDate(v), "dd.mm.yyyy hh:nn")

    Else

        v = ThisWorkbook.Worksheets("CDP_STOK").Range("C2").Value

        If IsDate(v) Then

            t = "Son stok guncellemesi: " & _
                Format$(CDate(v), "dd.mm.yyyy hh:nn")

        End If

    End If


    Set old = Nothing

    Set old = ws.Shapes("AKSORA_V7_TARIH")

    If Not old Is Nothing Then old.Delete


    Set old = Nothing

    Set old = ws.Shapes("AKSORA_V8_TARIH")

    If Not old Is Nothing Then old.Delete


    Set p = ws.Shapes(PRE & "STATUS")


    y = p.Top + p.Height + 6

    ht = 19


    Set s = ws.Shapes.AddTextbox( _
                msoTextOrientationHorizontal, _
                p.Left, _
                y, _
                p.Width, _
                ht)


    With s

        .Name = "AKSORA_V8_TARIH"

        .line.Visible = msoFalse

        .Fill.Visible = msoFalse

        .TextFrame2.MarginLeft = 0
        .TextFrame2.MarginRight = 0
        .TextFrame2.MarginTop = 0
        .TextFrame2.MarginBottom = 0

        .TextFrame2.TextRange.Text = t

        .TextFrame2.TextRange.Font.Name = "Aptos"

        .TextFrame2.TextRange.Font.Size = 10

        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = _
            RGB(183, 206, 221)

        .TextFrame2.TextRange.ParagraphFormat.Alignment = _
            msoAlignCenter

        .TextFrame2.VerticalAnchor = msoAnchorMiddle

        .Placement = xlFreeFloating

    End With


    On Error GoTo 0

End Sub


'==========================================================
' KAPAK YERLESIM
'==========================================================
Private Sub AKSORA_V8_YERLESIM(ByVal ws As Worksheet)

    Dim st As Shape
    Dim dt As Shape
    Dim bt As Shape
    Dim cr As Shape
    Dim pnl As Shape

    Dim y As Double
    Dim bottomLimit As Double
    Dim gap As Double


    On Error GoTo Cikis


    Set st = ws.Shapes(PRE & "STATUS")

    Set dt = ws.Shapes("AKSORA_V8_TARIH")

    Set bt = ws.Shapes(PRE & "ENTER")

    Set cr = ws.Shapes(PRE & "CREDIT")

    Set pnl = ws.Shapes(PRE & "PANEL")


    st.Height = 22

    dt.Height = 19

    gap = 10


    dt.Top = st.Top + st.Height + 5

    dt.Left = st.Left

    dt.Width = st.Width


    bt.Top = dt.Top + dt.Height + gap

    cr.Top = bt.Top + bt.Height + 15


    bottomLimit = pnl.Top + pnl.Height - 7


    If cr.Top + cr.Height > bottomLimit Then

        y = cr.Top + cr.Height - bottomLimit

        st.Top = st.Top - y
        dt.Top = dt.Top - y
        bt.Top = bt.Top - y
        cr.Top = cr.Top - y

    End If


    If st.Top < pnl.Top Then

        st.Top = pnl.Top + 4

        dt.Top = st.Top + st.Height + 3

        bt.Top = dt.Top + dt.Height + 5

        cr.Top = bt.Top + bt.Height + 6

    End If


Cikis:

End Sub


'==========================================================
' SECILI IL GORUNUMU
'==========================================================
Private Sub AKSORA_V8_SECIM_GORUNUM(ByVal ws As Worksheet)

    Dim keys(1 To 5) As String
    Dim names(1 To 5) As String

    Dim i As Long
    Dim key As String


    keys(1) = "GIRESUN"
    keys(2) = "TRABZON"
    keys(3) = "RIZE"
    keys(4) = "GUMUSHANE"
    keys(5) = "ARTVIN"


    names(1) = _
        "G" & ChrW(304) & "RESUN"

    names(2) = _
        "TRABZON"

    names(3) = _
        "R" & ChrW(304) & "ZE"

    names(4) = _
        "G" & ChrW(220) & _
        "M" & ChrW(220) & _
        ChrW(350) & "HANE"

    names(5) = _
        "ARTV" & ChrW(304) & "N"


    key = UCase$(Trim$(CStr(ws.Range("AZ1").Value)))


    For i = 1 To 5

        With ws.Shapes(PRE & "IL_" & i)

            If keys(i) = key Then

                .Fill.ForeColor.RGB = RGB(32, 78, 76)

                .line.ForeColor.RGB = RGB(115, 223, 157)

                .line.Weight = 2.5


                ws.Shapes(PRE & "STATUS") _
                    .TextFrame2.TextRange.Text = _
                    "SE" & ChrW(199) & _
                    ChrW(304) & _
                    "L" & ChrW(304) & _
                    " " & ChrW(304) & _
                    "L: " & names(i)

            Else

                .Fill.ForeColor.RGB = RGB(40, 79, 111)

                .line.ForeColor.RGB = RGB(83, 133, 168)

                .line.Weight = 1

            End If

        End With

    Next i

End Sub


'==========================================================
' IL SEC
'==========================================================
Public Sub AKSORA_V8_IL_SEC()
    Dim ws As Worksheet, ix As Long, keys(1 To 5) As String
    Dim ev As Boolean, sc As Boolean, hataMetni As String
    If AKSORA_V11_MESGUL Then Exit Sub
    ev = application.EnableEvents: sc = application.ScreenUpdating
    On Error GoTo Hata
    Set ws = ThisWorkbook.Worksheets(KAPAK_SAYFA)
    ix = CLng(ws.Shapes(CStr(application.Caller)).AlternativeText)
    keys(1) = "GIRESUN": keys(2) = "TRABZON": keys(3) = "RIZE"
    keys(4) = "GUMUSHANE": keys(5) = "ARTVIN"
    If ix < 1 Or ix > 5 Then Exit Sub
    ' Secim tek seferde cizilir; hesaplama olaylari baska sayfaya gecis yaptirmaz.
    application.ScreenUpdating = False
    application.EnableEvents = False
    ws.Range("AZ1").Value = keys(ix)
    AKSORA_V8_SECIM_GORUNUM ws
    AKSORA_V10_TARIH_GOSTER
Cikis:
    On Error Resume Next
    If Not (ActiveSheet Is ws) And Not ws Is Nothing Then ws.Activate
    application.EnableEvents = ev
    application.ScreenUpdating = True
    On Error GoTo 0
    If Len(hataMetni) > 0 Then MsgBox "Il secimi: " & hataMetni, vbExclamation, "AKSORA"
    Exit Sub
Hata:
    hataMetni = Err.Description
    Resume Cikis
End Sub


'==========================================================
' ANA GIRIS
'==========================================================
Public Sub AKSORA_V8_GIRIS()

    Static mesgul As Boolean

    Dim hedef As String
    Dim olay As Boolean

    Dim hataMetni As String

    Dim bas As Double


    If mesgul _
       Or AKSORA_V11_MESGUL _
       Or AKSORA_V13_ACILIS Then

        Exit Sub

    End If


    olay = application.EnableEvents

    bas = Timer

    mesgul = True


    On Error GoTo Hata


    application.EnableEvents = False

    application.ScreenUpdating = False


    AKSORA_V13_GECIS = True


    hedef = AKSORA_AKTIF_IL()


    If AKSORA_CANLI_STOK_IL() <> hedef Then

        AKSORA_IL_GECIS_V1 hedef

        If AKSORA_CANLI_STOK_IL() <> hedef Then
            GoTo Cikis
        End If

    End If


    AKSORA_GIRIS_AC


    If Not AKSORA_GIRIS_ACIK Then
        GoTo Cikis
    End If


    PANEL_YENILE


    ThisWorkbook.Worksheets("ANA PANEL").Select Replace:=True


    AKSORA_V13_GORUNUM _
        ThisWorkbook.Worksheets("ANA PANEL")


Cikis:

    ' Il secimi sonrasi her zaman ana panelde (veya basarisizsa kapakta) kalinir;
    ' gecis sirasinda acilan rapor/stok sayfasinda kalinmaz.
    On Error Resume Next
    If ActiveSheet.Name <> "ANA PANEL" Then
        If AKSORA_GIRIS_ACIK And AKSORA_CANLI_STOK_IL() = AKSORA_AKTIF_IL() Then
            ThisWorkbook.Worksheets("ANA PANEL").Select Replace:=True
            AKSORA_V13_GORUNUM ThisWorkbook.Worksheets("ANA PANEL")
        Else
            ThisWorkbook.Worksheets("KAPAK").Activate
        End If
    End If
    On Error GoTo 0

    AKSORA_V13_GECIS = False


    application.EnableEvents = olay

    application.ScreenUpdating = True


    mesgul = False


    AKSORA_V13_SURE_KAYDET _
        "Il girisi", _
        bas


    If Len(hataMetni) > 0 Then

        MsgBox _
            "Ana panel acilamadi." & _
            vbCrLf & hataMetni, _
            vbExclamation, _
            "AKSORA"

    End If


    Exit Sub


Hata:

    hataMetni = Err.Description

    Resume Cikis

End Sub



Private Sub AKSORA_GIRIS_AC()

    Dim kap As Worksheet
    Dim ws As Worksheet

    Dim nm As String
    Dim r As Long

    Dim ev As Boolean
    Dim scr As Boolean


    On Error GoTo Hata


    Set kap = ThisWorkbook.Worksheets(KAPAK_SAYFA)


    If Not AKSORA_SAP_GUNCELLEME_IZNI() Then
        Exit Sub
    End If


    ev = application.EnableEvents

    scr = application.ScreenUpdating


    application.EnableEvents = False

    application.ScreenUpdating = False


    '------------------------------------------------------
    ' KAYITLI SAYFALARI AC
    '------------------------------------------------------
    For r = 2 To 100

        nm = Trim$(CStr(kap.Cells(r, "BA").Value))


        If Len(nm) = 0 Then Exit For


        Set ws = Nothing


        On Error Resume Next

        Set ws = ThisWorkbook.Worksheets(nm)

        On Error GoTo Hata


        If Not ws Is Nothing Then
            ws.Visible = xlSheetVisible
        End If

    Next r


    ThisWorkbook.Worksheets("ANA PANEL").Visible = _
        xlSheetVisible


    kap.Move Before:=ThisWorkbook.Worksheets(1)


    AKSORA_GIRIS_ACIK = True


    ThisWorkbook.Worksheets("ANA PANEL").Activate


    With ThisWorkbook.Windows(1)

        .DisplayWorkbookTabs = True

        .DisplayHorizontalScrollBar = True

        .DisplayVerticalScrollBar = True

        .DisplayHeadings = False

        .DisplayGridlines = False

    End With


    application.DisplayStatusBar = True

    application.StatusBar = False


    application.EnableEvents = ev

    application.ScreenUpdating = scr


    Exit Sub


Hata:

    application.EnableEvents = True

    application.ScreenUpdating = True


    MsgBox _
        "Giris hatasi (" & _
        Err.Number & "): " & _
        Err.Description, _
        vbCritical, _
        "AKSORA"

End Sub


Public Sub AKSORA_KAPAK_KILITLE()

    Dim ws As Worksheet
    Dim kap As Worksheet

    Dim n As Long

    Dim ev As Boolean
    Dim scr As Boolean


    On Error GoTo Hata


    ev = application.EnableEvents

    scr = application.ScreenUpdating


    application.EnableEvents = False

    application.ScreenUpdating = False


    Set kap = ThisWorkbook.Worksheets(KAPAK_SAYFA)


    AKSORA_GIRIS_ACIK = False


    '------------------------------------------------------
    ' ILK ACILISTA GORUNUR SAYFALARI KAYDET
    '------------------------------------------------------
    If Len(CStr(kap.Range("BA1").Value)) = 0 Then

        kap.Range("BA1").Value = "AKSORA_SAYFALAR"


        n = 2


        For Each ws In ThisWorkbook.Worksheets

            If ws.Name <> KAPAK_SAYFA _
               And ws.Visible = xlSheetVisible Then

                kap.Cells(n, "BA").Value = ws.Name

                n = n + 1

            End If

        Next ws

    End If


    kap.Visible = xlSheetVisible


    kap.Move Before:=ThisWorkbook.Worksheets(1)


    kap.Activate


    '------------------------------------------------------
    ' KAPAK BUTON BAGLANTILARI
    '------------------------------------------------------
    On Error Resume Next


    kap.Shapes(PRE & "ENTER").OnAction = _
        "'" & Replace(ThisWorkbook.Name, "'", "''") & _
        "'!AKSORA_V8_GIRIS"


    On Error GoTo Hata


    '------------------------------------------------------
    ' DIGER SAYFALARI GIZLE
    '------------------------------------------------------
    For Each ws In ThisWorkbook.Worksheets

        If ws.Name <> KAPAK_SAYFA _
           And ws.Visible = xlSheetVisible Then

            ws.Visible = xlSheetVeryHidden

        End If

    Next ws


    AKSORA_KAPAK_EKRAN_SABITLE


    application.EnableEvents = ev

    application.ScreenUpdating = scr


    Exit Sub


Hata:

    application.EnableEvents = True

    application.ScreenUpdating = True


    MsgBox _
        "Giris kilidi hatasi: " & _
        Err.Description, _
        vbCritical, _
        "AKSORA"

End Sub


'==========================================================
' KAPAK EKRAN SABITLE
'==========================================================
Private Sub AKSORA_KAPAK_EKRAN_SABITLE()

    Dim ws As Worksheet


    On Error Resume Next


    Set ws = ThisWorkbook.Worksheets(KAPAK_SAYFA)


    ws.ScrollArea = "A1:AY49"


    ws.Activate


    ActiveWindow.ScrollRow = 1

    ActiveWindow.ScrollColumn = 1


    ActiveWindow.DisplayHorizontalScrollBar = False

    ActiveWindow.DisplayVerticalScrollBar = False

    ActiveWindow.DisplayHeadings = False

    ActiveWindow.DisplayGridlines = False


    ws.Rows("50:1048576").Hidden = True


    ws.Range("A1").Select


    On Error GoTo 0

End Sub


'==========================================================
' KAPAK GORUNUM
'==========================================================
Public Sub AKSORA_V8_KAPAK_GORUNUM()

    On Error Resume Next


    If ActiveSheet.Name <> KAPAK_SAYFA Then
        Exit Sub
    End If


    AKSORA_KAPAK_EKRAN_SABITLE


    ActiveWindow.DisplayWorkbookTabs = False

    ActiveWindow.DisplayHorizontalScrollBar = False

    ActiveWindow.DisplayVerticalScrollBar = False


    ActiveWindow.ScrollRow = 1

    ActiveWindow.ScrollColumn = 1


    AKSORA_EKRAN_SIGDIR ThisWorkbook.Worksheets(KAPAK_SAYFA)

    On Error GoTo 0

End Sub


'==========================================================
' YENIDEN BOYUTLANDIR
'==========================================================
Public Sub AKSORA_V8_YENIDEN_BOYUTLANDIR()

    On Error GoTo Cikis


    If mResizing Then Exit Sub


    AKSORA_V8_DUZENLE


Cikis:

End Sub
'==========================================================
' IL GUNCELLEME TARIHLERI
'==========================================================

Private Function AKSORA_IL_TARIH_SIRA(ByVal kod As String) As Long

    Select Case AKFIX_NORMAL(kod)

        Case "GIRESUN"
            AKSORA_IL_TARIH_SIRA = 2

        Case "TRABZON"
            AKSORA_IL_TARIH_SIRA = 3

        Case "RIZE"
            AKSORA_IL_TARIH_SIRA = 4

        Case "GUMUSHANE"
            AKSORA_IL_TARIH_SIRA = 5

        Case "ARTVIN"
            AKSORA_IL_TARIH_SIRA = 6

    End Select

End Function


Public Sub AKSORA_V10_KUR()

    Dim il As Variant
    Dim tarih As Date

    For Each il In Array( _
        "GIRESUN", _
        "TRABZON", _
        "RIZE", _
        "GUMUSHANE", _
        "ARTVIN")

        tarih = AKSORA_V11_IL_TARIH(CStr(il))

        If tarih > 0 Then
            AKSORA_V10_TARIH_KAYDET CStr(il), tarih
        End If

    Next il

    AKSORA_V10_TARIH_GOSTER

End Sub


Public Sub AKSORA_V10_TARIH_GOSTER()

    Dim ws As Worksheet
    Dim tarih As Date
    Dim metin As String

    Set ws = ThisWorkbook.Worksheets("KAPAK")

    tarih = _
        AKSORA_V11_IL_TARIH( _
            CStr(ws.Range("AZ1").Value2) _
        )

    If tarih > 0 Then

        metin = _
            "Son stok güncellemesi: " & _
            Format$(tarih, "dd.mm.yyyy hh:nn")

    Else

        metin = _
            "Son stok güncellemesi: Henüz güncellenmedi"

    End If

    ws.Shapes("AKSORA_V8_TARIH") _
        .TextFrame2.TextRange.Text = metin

End Sub


Public Sub AKSORA_V10_TARIH_KAYDET( _
    ByVal ilKodu As String, _
    ByVal veriTarihi As Date)

    Dim ws As Worksheet
    Dim r As Long

    Dim ev As Boolean
    Dim korumali As Boolean

    Dim no As Long
    Dim aciklama As String


    ev = application.EnableEvents

    On Error GoTo Hata


    r = AKSORA_IL_TARIH_SIRA(ilKodu)


    If r = 0 Or veriTarihi <= 0 Then
        Exit Sub
    End If


    Set ws = ThisWorkbook.Worksheets("KAPAK")


    korumali = ws.ProtectContents


    application.EnableEvents = False


    ws.Unprotect


    ws.Cells(r, "BB").Value2 = _
        AKFIX_NORMAL(ilKodu)


    ws.Cells(r, "BC").Value = _
        veriTarihi


    ws.Cells(r, "BC").NumberFormat = _
        "dd.mm.yyyy hh:mm"


    ws.Columns("BB:BC").Hidden = True


Cikis:

    On Error Resume Next


    If korumali Then
        AKSORA_KAPAK_KORU
    End If


    application.EnableEvents = ev


    On Error GoTo 0


    If no <> 0 Then

        Err.Raise _
            no, _
            "AKSORA_V10_TARIH_KAYDET", _
            aciklama

    End If

    Exit Sub


Hata:

    no = Err.Number

    aciklama = Err.Description

    Resume Cikis

End Sub


