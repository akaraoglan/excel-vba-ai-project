VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmstokmenu 
   Caption         =   "Stok Ýþlemleri"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2145
   OleObjectBlob   =   "frmstokmenu.frx":0000
End
Attribute VB_Name = "frmstokmenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'Bu userformu sola formun soluna sabitler ve 82 birim aþaðýya getirir.
Private Sub UserForm_Initialize()
Me.Left = frmAnaForm.Left + 127
Me.Top = frmAnaForm.Top + 82

End Sub

'Stok tanýmlama isimli butona týklandýðýnda stok listesini açar.
Private Sub btnstoktanimlama_Click()
Unload Me
frmstoklistesi.Show
End Sub


'Stok menülerini geçerken her bir butonun üzerine gelindiðinde yeþil renge döndürür.

' --- STOK TANIMLAMA BUTONU ---
Private Sub btnstoktanimlama_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Sadece bu butonu yeþil yap, diðerlerini varsayýlan renge çek
    Call ResetButtonColors
    btnstoktanimlama.BackColor = &HFF00&
End Sub

' --- STOK GÝRÝÞÝ BUTONU ---
Private Sub btnstokgirisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnstokgirisi.BackColor = &HFF00&
End Sub

' --- STOK ÇIKIÞI BUTONU ---
Private Sub btnstokcikisi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnstokcikisi.BackColor = &HFF00&
End Sub

' --- STOK DURUMU BUTONU ---
Private Sub btnstokdurumu_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnstokdurumu.BackColor = &HFF00&
End Sub

' --- FORMUN BOÞLUÐUNA GELÝNDÝÐÝNDE ---
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Fare butonlarýn dýþýna çýktýðýnda hepsini eski haline döndür
    Call ResetButtonColors
End Sub
' --- YARDIMCI PROSEDÜR ---
Sub ResetButtonColors()
    ' Renk kodlarýnýn sonuna & ekledik
    btnstoktanimlama.BackColor = &H8000000F
    btnstokgirisi.BackColor = &H8000000F
    btnstokcikisi.BackColor = &H8000000F
    btnstokdurumu.BackColor = &H8000000F
End Sub
