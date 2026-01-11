VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmcarimenu 
   Caption         =   "Cari Ýþlemler"
   ClientHeight    =   4470
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   2145
   OleObjectBlob   =   "frmcarimenu.frx":0000
End
Attribute VB_Name = "frmcarimenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'Bu userformu sola formun soluna sabitler ve 82 birim aþaðýya getirir.
Private Sub UserForm_Initialize()
Me.Left = frmAnaForm.Left + 245
Me.Top = frmAnaForm.Top + 82

End Sub

'cari tanýmlama isimli butona týklandýðýnda cari listesini açar.
Private Sub btncaritanimlama_Click()
Unload Me
frmcarilistesi.Show
End Sub


'cari menülerini geçerken her bir butonun üzerine gelindiðinde yeþil renge döndürür.

' --- cari TANIMLAMA BUTONU ---
Private Sub btncaritanimlama_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Sadece bu butonu yeþil yap, diðerlerini varsayýlan renge çek
    Call ResetButtonColors
    btncaritanimlama.BackColor = &HFF00&
End Sub

' --- cari GÝRÝÞÝ BUTONU ---
Private Sub btnborclucariler_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnborclucariler.BackColor = &HFF00&
End Sub

' --- cari ÇIKIÞI BUTONU ---
Private Sub btnalacaklicariler_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call ResetButtonColors
    btnalacaklicariler.BackColor = &HFF00&
End Sub



' --- FORMUN BOÞLUÐUNA GELÝNDÝÐÝNDE ---
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    ' Fare butonlarýn dýþýna çýktýðýnda hepsini eski haline döndür
    Call ResetButtonColors
End Sub
' --- YARDIMCI PROSEDÜR ---
Sub ResetButtonColors()
    ' Renk kodlarýnýn sonuna & ekledik
    btncaritanimlama.BackColor = &H8000000F
    btnborclucariler.BackColor = &H8000000F
    btnalacaklicariler.BackColor = &H8000000F
    
End Sub
