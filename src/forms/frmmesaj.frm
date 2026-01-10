VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmmesaj 
   Caption         =   "Mesaj"
   ClientHeight    =   1665
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7860
   OleObjectBlob   =   "frmmesaj.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmmesaj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tamam isimli butona týklandýðýnda mesaj kutusunu kapatýr.
Private Sub btntamam_Click()
Unload Me
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
