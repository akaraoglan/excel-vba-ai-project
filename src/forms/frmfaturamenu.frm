VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmfaturamenu 
   Caption         =   "Fatura Ýþlemleri"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2145
   OleObjectBlob   =   "frmfaturamenu.frx":0000
End
Attribute VB_Name = "frmfaturamenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'Bu userformu sola formun soluna sabitler ve 82 birim aþaðýya getirir.
Private Sub UserForm_Initialize()
Me.Left = frmAnaForm.Left + 482
Me.Top = frmAnaForm.Top + 82

End Sub

'fatura tanýmlama isimli butona týklandýðýnda fatura listesini açar.
Private Sub btnalisfaturasi_Click()
Unload Me
frmfaturalistesi.Show
End Sub


'fatura menülerini geçerken her bir butonun üzerine gelindiðinde yeþil renge döndürür.

' --- fatura TANIMLAMA BUTONU ---
Private Sub btnalisfaturasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Sadece bu butonu yeþil yap, diðerlerini varsayýlan renge çek
    Call ResetButtonColors
    btnalisfaturasi.BackColor = &HFF00&
End Sub

' --- fatura GÝRÝÞÝ BUTONU ---
Private Sub btnsatisfaturasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnsatisfaturasi.BackColor = &HFF00&
End Sub

' --- fatura ÇIKIÞI BUTONU ---
Private Sub btnalisiadefaturasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnalisiadefaturasi.BackColor = &HFF00&
End Sub

' --- fatura DURUMU BUTONU ---
Private Sub btnsatisiadefaturasi_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnsatisiadefaturasi.BackColor = &HFF00&
End Sub

' --- FORMUN BOÞLUÐUNA GELÝNDÝÐÝNDE ---
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Fare butonlarýn dýþýna çýktýðýnda hepsini eski haline döndür
    Call ResetButtonColors
End Sub
' --- YARDIMCI PROSEDÜR ---
Sub ResetButtonColors()
    ' Renk kodlarýnýn sonuna & ekledik
    btnalisfaturasi.BackColor = &H8000000F
    btnsatisfaturasi.BackColor = &H8000000F
    btnalisiadefaturasi.BackColor = &H8000000F
    btnsatisiadefaturasi.BackColor = &H8000000F
End Sub
