VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmfaturalistesi 
   Caption         =   "UserForm1"
   ClientHeight    =   13485
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20655
   OleObjectBlob   =   "frmfaturalistesi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmfaturalistesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnkapat_Click()
Unload Me
End Sub
Private Sub UserForm_Initialize()
    txtfaturakodu.Value = "AF0000" & Sheets("Tanimlamalar").Range("E2").Value + 1
    txttarih.Value = Format(Date, "dd.mm.yyyy")
End Sub
Private Sub txttarih_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' 0-9 ve . dýþýnda her þeyi engelle
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 46) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txttarih_Change()
    Static kilit As Boolean
    If kilit Then Exit Sub
    kilit = True

    Dim s As String
    s = Replace(txttarih.Value, ".", "")

    If IsNumeric(s) Then
        Select Case Len(s)
            Case 3
                txttarih.Value = Left(s, 2) & "." & Mid(s, 3)
            Case 5
                txttarih.Value = Left(s, 2) & "." & Mid(s, 3, 2) & "." & Mid(s, 5)
            Case 8
                txttarih.Value = Left(s, 2) & "." & Mid(s, 3, 2) & "." & Mid(s, 5, 4)
        End Select
    End If

    txttarih.SelStart = Len(txttarih.Value)
    kilit = False
End Sub

Private Sub txttarih_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Dim d As Date
    Dim bugun As Date

    bugun = Date

    On Error GoTo Hata
    d = DateValue(Replace(txttarih.Value, ".", "/"))

    ' GEÇMÝÞ YIL uyarýsý
    If Year(d) < Year(bugun) Then
        MsgBox "Uyarý: Geçmiþ bir yýla ait tarih girdiniz.", vbInformation

    ' GELECEK (ayný yýl ama farklý ay)
    ElseIf Year(d) = Year(bugun) And Month(d) <> Month(bugun) Then
        MsgBox "Uyarý: Bulunduðunuz ay dýþýndaki bir tarih girdiniz.", vbInformation

    ' GELECEK YIL (hangi ay olursa olsun)
    ElseIf Year(d) > Year(bugun) Then
        MsgBox "Uyarý: Gelecek bir yýla ait tarih girdiniz.", vbInformation
    End If

    txttarih.Value = Format(d, "dd.mm.yyyy")
    Exit Sub

Hata:
    MsgBox "Geçerli bir tarih giriniz (GG.AA.YYYY).", vbCritical
End Sub



