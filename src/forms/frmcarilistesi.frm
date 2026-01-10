VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmcarilistesi 
   Caption         =   "Cari Listesi"
   ClientHeight    =   9900.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15855
   OleObjectBlob   =   "frmcarilistesi.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmcarilistesi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Function UCaseTR(ByVal txt As String) As String
    Dim result As String
    result = txt
    
    ' Önce küçük Türkçe karakterleri tek tek büyük halleriyle deðiþtiriyoruz
    result = Replace(result, "i", "Ý")
    result = Replace(result, "ý", "I")
    result = Replace(result, "ç", "Ç")
    result = Replace(result, "þ", "Þ")
    result = Replace(result, "ö", "Ö")
    result = Replace(result, "ü", "Ü")
    result = Replace(result, "ð", "Ð")
    
    ' En son geri kalan standart Ýngilizce karakterleri büyütüyoruz
    UCaseTR = UCase(result)
End Function
Function LCaseTR(ByVal txt As String) As String
    Dim result As String
    result = txt
    
    ' Önce büyük Türkçe karakterleri tek tek küçük halleriyle deðiþtiriyoruz
    result = Replace(result, "Ý", "i")
    result = Replace(result, "I", "ý")
    result = Replace(result, "Ç", "ç")
    result = Replace(result, "Þ", "þ")
    result = Replace(result, "Ö", "ö")
    result = Replace(result, "Ü", "ü")
    result = Replace(result, "Ð", "ð")
    
    ' En son geri kalan standart Ýngilizce karakterleri küçültüyoruz
    LCaseTR = LCase(result)
End Function





Private Sub Frame1_Click()

End Sub

'Çarpý Ýþaretini Ýþlevsiz Hale Getirme
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
If CloseMode = vbFormControlMenu Then Cancel = True
End Sub
'Kapat týklandýðýnda Userformu kapatýr.
Private Sub btnkapat_Click()
Unload Me
frmAnaForm.Show
End Sub
'Form yüklendiðinde listeyi getirir.
Private Sub UserForm_Initialize()

CariListele

End Sub

' Cari sayfasýndaki deðerleri listboxa getiriyoruz. Burada yaptýðýmýz iþlem cari sayfasýndaki deðerleri TMP isimli gizli sayfaya yazýyoruz ve oradan çekiyoruz. Bu sayede listboxa
' parçalý bir þekilde dilediðimiz sütunu getirebiliriz
Sub CariListele()

    Dim ws As Worksheet, tmp As Worksheet
    Dim sonSatir As Long
    Dim i As Long

    Set ws = ThisWorkbook.Sheets("Cari")
    Set tmp = ThisWorkbook.Sheets("TMP")

    tmp.Cells.Clear

    ' Baþlýklar
    tmp.Range("A1:G1").Value = Array("CARÝ KODU", "ADI ÜNVANI", "VERGÝ DAÝRESÝ", "VERGÝ NUMARASI", "TELEFONU", "EMAÝL", "ADRESÝ")

    sonSatir = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If sonSatir < 2 Then Exit Sub

    ' Verileri parça parça al
    For i = 2 To sonSatir
        tmp.Cells(i, 1).Value = ws.Cells(i, "A").Value
        tmp.Cells(i, 2).Value = ws.Cells(i, "B").Value
        tmp.Cells(i, 3).Value = ws.Cells(i, "C").Value
        tmp.Cells(i, 4).Value = ws.Cells(i, "D").Value
        tmp.Cells(i, 5).Value = ws.Cells(i, "E").Value
        tmp.Cells(i, 6).Value = ws.Cells(i, "F").Value
        tmp.Cells(i, 7).Value = ws.Cells(i, "G").Value
    Next i

    ' ListBox
    With lstcariler
        .Clear
        .ColumnCount = 7
        .ColumnWidths = "100;300;80;100;100;100;100"
        .ColumnHeads = True
        .RowSource = "TMP!A2:G" & sonSatir
    End With

    ' Ýstersen TMP'yi gizle
    tmp.Visible = xlSheetVeryHidden

End Sub

'Listboxa çift týklandýðýnda deðiþiklik yapmak için önce textboxlarý ve comboboxlarý getiriyoruz. Comboboxlarýmýz eðer içine deðer yazýlamayacak þekilde ise gelmesi için önce bu
'fonksiyonu aktif hale getiriyoruz. Ayrýca eðer comboboxlarýmýz ucasetr fonksiyonu (Büyük harfe çevirme) içeriyorsa doðru gelmesi için ayarlanmýþtýr.
Private Sub ComboSec(cb As MSForms.ComboBox, deger As String)
    Dim i As Integer
    For i = 0 To cb.ListCount - 1
        ' ucasetr kullanarak her iki tarafý da büyük harfe çevirip karþýlaþtýrýn
        If UCaseTR(Trim(cb.List(i))) = UCaseTR(Trim(deger)) Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub
'Listboxa çift týklandýðýnda deðiþiklik yapmak için önce textboxlarý ve comboboxlarý getiriyoruz. Comboboxlarýmýz eðer içine deðer yazýlamayacak þekilde ise gelmesi için önce bu
'fonksiyonu aktif hale getiriyoruz. Ayrýca eðer comboboxlarýmýz ucasetr fonksiyonu (küçük harfe çevirme) içeriyorsa doðru gelmesi için ayarlanmýþtýr.
Private Sub ComboSec1(cb As MSForms.ComboBox, deger As String)
    Dim i As Integer
    For i = 0 To cb.ListCount - 1
        ' Lcasetr kullanarak her iki tarafý da küçük harfe çevirip karþýlaþtýrýn
        If LCaseTR(Trim(cb.List(i))) = LCaseTR(Trim(deger)) Then
            cb.ListIndex = i
            Exit Sub
        End If
    Next i
End Sub

'Listboxumuzdaki deðere çift týklandýðýnda deðerleri getiriyoruz.
Private Sub lstcariler_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    If lstcariler.ListIndex = -1 Then Exit Sub

    With frmcaritanimlama
        .txtcarikodu.Value = lstcariler.List(lstcariler.ListIndex, 0)
        .txtadunvan.Value = lstcariler.List(lstcariler.ListIndex, 1)
        .txtvergidairesi.Value = lstcariler.List(lstcariler.ListIndex, 2)
        .txtvergino.Value = lstcariler.List(lstcariler.ListIndex, 3)
        .txttelefon.Value = lstcariler.List(lstcariler.ListIndex, 4)
        .txtemail.Value = lstcariler.List(lstcariler.ListIndex, 5)
        .txtadres.Value = lstcariler.List(lstcariler.ListIndex, 6)
          

        frmcaritanimlama.lblislem.Caption = "Düzeltme"
        frmcaritanimlama.btnkaydet.Caption = "Güncelle" 'Listboxtaki deðere çift týklandýðýnda Kaydet butonunun ismini güncelle yapar.

        .Show
    End With

End Sub

Private Sub btnekle_Click()
frmcaritanimlama.lblislem.Caption = "Yeni" 'Eðer frmcaritanimlama sayfasýndaki label deðeri yeni ise fmrcaritanimlama isimli formu açar.
frmcaritanimlama.Show
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








