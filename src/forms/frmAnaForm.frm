VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAnaForm 
   Caption         =   "EXCEL MUHASEBE 1.0"
   ClientHeight    =   8130
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   20475
   OleObjectBlob   =   "frmAnaForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAnaForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btncariislemler_Click()

Unload Me
frmcarilistesi.Show
End Sub


'Çarpý Ýþaretini Ýþlevsiz Hale Getirme
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then Cancel = True
End Sub

'Kapat isimli butona týklandýðýnda Userformu kaydeder ve kapatýr.
Private Sub btnkapat_Click()
'Application.ActiveWorkbook.Save
Unload Me
End Sub

'Stok iþlemleri isimli butona týklandýðýnda btnstokislemleri isimli menüyü (Userformu) açar
Private Sub btnstokislemleri_Click()
frmstokmenu.Show
End Sub

