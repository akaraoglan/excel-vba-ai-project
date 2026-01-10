VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmstoklistesi 
   Caption         =   "Stok Listesi"
   ClientHeight    =   9900.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   OleObjectBlob   =   "frmstoklistesi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmstoklistesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Çarpý Ýþaretini Ýþlevsiz Hale Getirme
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
'Kapat týklandýðýnda Userformu kapatýr.
Private Sub btnkapat_Click()
Unload Me
End Sub
'Form yüklendiðinde listeyi getirir.
Private Sub UserForm_Initialize()

StoklariListele

End Sub

' Stok sayfasýndaki deðerleri listboxa getiriyoruz. Burada yaptýðýmýz iþlem stok sayfasýndaki deðerleri TMP isimli gizli sayfaya yazýyoruz ve oradan çekiyoruz. Bu sayede listboxa
' parçalý bir þekilde dilediðimiz sütunu getirebiliriz
Sub StoklariListele()

    Dim ws As Worksheet, tmp As Worksheet
    Dim sonSatir As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Stok")
    Set tmp = ThisWorkbook.Sheets("TMP")

    tmp.Cells.Clear

    ' Baþlýklar
    tmp.Range("A1:F1").Value = Array("Stok Kodu", "Açýklama", "Birimi", "Alýþ Fiyatý", "Satýþ Fiyatý", "KDV")

    sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If sonSatir < 2 Then Exit Sub

    ' Verileri parça parça al
    For i = 2 To sonSatir
        tmp.Cells(i, 1).Value = ws.Cells(i, "A").Value
        tmp.Cells(i, 2).Value = ws.Cells(i, "B").Value
        tmp.Cells(i, 3).Value = ws.Cells(i, "C").Value
        tmp.Cells(i, 4).Value = ws.Cells(i, "D").Value
        tmp.Cells(i, 5).Value = ws.Cells(i, "E").Value
        tmp.Cells(i, 6).Value = ws.Cells(i, "I").Value
    Next i

    ' ListBox
    With lststoklar
        .Clear
        .ColumnCount = 6
        .ColumnWidths = "100;300;80;100;100;100"
        .ColumnHeads = True
        .RowSource = "TMP!A2:F" & sonSatir
    End With

    ' Ýstersen TMP'yi gizle
    tmp.Visible = xlSheetVeryHidden

End Sub

'Listboxa çift týklandýðýnda deðiþiklik yapmak için önce textboxlarý ve comboboxlarý getiriyoruz. Comboboxlarýmýz eðer içine deðer yazýlamayacak þekilde ise gelmesi için önce bu
'fonksiyonu aktif hale getiriyoruz. Ayrýca eðer comboboxlarýmýz UCASE fonksiyonu (Büyük harfe çevirme) içeriyorsa doðru gelmesi için ayarlanmýþtýr.
Private Sub ComboSec(cb As MSForms.ComboBox, deger As String)
    Dim i As Integer
    For i = 0 To cb.ListCount - 1
        ' UCase kullanarak her iki tarafý da büyük harfe çevirip karþýlaþtýrýn
        If UCase(Trim(cb.List(i))) = UCase(Trim(deger)) Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

'Listboxumuzdaki deðere çift týklandýðýnda deðerleri getiriyoruz.
Private Sub lstStoklar_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If lststoklar.ListIndex = -1 Then Exit Sub

    With frmstoktanimlama
        .txtstokkodu.Value = lststoklar.List(lststoklar.ListIndex, 0)
        .txtaciklama.Value = lststoklar.List(lststoklar.ListIndex, 1)
        .txtalis.Value = lststoklar.List(lststoklar.ListIndex, 3)
        .txtsatis.Value = lststoklar.List(lststoklar.ListIndex, 4)

        ' ?? ÖNEMLÝ KISIM
        ComboSec .cbbirim, lststoklar.List(lststoklar.ListIndex, 2)
        ComboSec .cbkdv, lststoklar.List(lststoklar.ListIndex, 5)
        frmstoktanimlama.lblislem.Caption = "Düzeltme"
        frmstoktanimlama.btnkaydet.Caption = "Güncelle" 'Listboxtaki deðere çift týklandýðýnda Kaydet butonunun ismini güncelle yapar.

        .Show
    End With

End Sub

Private Sub btnekle_Click()
frmstoktanimlama.lblislem.Caption = "Yeni" 'Eðer frmstoktanimlama sayfasýndaki label deðeri yeni ise fmrstoktanimlama isimli formu açar.
frmstoktanimlama.Show
End Sub



























































'Sub StoklariListeleE()


    'Dim ws As Worksheet
    'Dim sonSatir As Long
    'Dim i As Long
    'Dim arr()

    'Set ws = ThisWorkbook.Sheets("Stok")

    'sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    'If sonSatir < 2 Then Exit Sub

    ' 6 sütunluk dizi (A-E + I)
    'ReDim arr(1 To sonSatir - 1, 1 To 6)

    'For i = 2 To sonSatir
        'arr(i - 1, 1) = ws.Cells(i, "A").Value
        'arr(i - 1, 2) = ws.Cells(i, "B").Value
        'arr(i - 1, 3) = ws.Cells(i, "C").Value
        'arr(i - 1, 4) = ws.Cells(i, "D").Value
        'arr(i - 1, 5) = ws.Cells(i, "E").Value
        'arr(i - 1, 6) = ws.Cells(i, "I").Value ' 9. sütun
    'Next i

    'With lststoklar
        '.Clear
        '.ColumnCount = 6
        '.ColumnWidths = "100;300;80;100;100;100"
        '.List = arr
    'End With

'End Sub








