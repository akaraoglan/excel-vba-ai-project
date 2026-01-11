VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmstoktanimlama 
   Caption         =   "Stok Tanýmlama Formu"
   ClientHeight    =   7320
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435.001
   OleObjectBlob   =   "frmstoktanimlama.frx":0000
End
Attribute VB_Name = "frmstoktanimlama"
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

Private Sub Image1_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Data As MSForms.DataObject, ByVal X As Single, ByVal Y As Single, ByVal DragState As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)

End Sub

'Bu userformu ana formun tam ortasýna sabitler ve gerekli ayarlamalar yapýlýr.
Private Sub UserForm_Initialize()
    ' Yatayda (Soldan saða) ortalamak için:
    Me.Left = frmAnaForm.Left + (frmAnaForm.Width - Me.Width) / 2
    
    ' Dikeyde (Yukarýdan aþaðýya) ortalamak için:
    Me.Top = frmAnaForm.Top + (frmAnaForm.Height - Me.Height) / 2
    
    
Frame1.Enabled = True
btnkaydet.Enabled = True
btniptal.Enabled = True
txtaciklama.SetFocus 'Her yeni kayýt yapýldýktan sonra aciklama isimli textboxa odaklanýr.
txtstokkodu.Value = "STK00000" & Sheets("Tanimlamalar").Range("C2").Value + 1 'Kayýt sýrasýnda stok kodu otomatik olarak STK00000 + Tanýmlamar sayfasýndaki deðerin bir fazlasý gelir.
End Sub

Private Sub btnKaydet_Click()
Dim X As Long
Dim sor As Byte

'--- 1. KRÝTÝK ALAN KONTROLÜ (Zorunlu alanlar) ---
If txtstokkodu.Value = "" Or txtaciklama.Value = "" Or cbbirim.Value = "" Or cbkdv.Value = "" Then
    MsgBox "Lütfen zorunlu alanlarý (Stok Kodu, Açýklama, Birim) doldurunuz!", vbExclamation
    Exit Sub
End If

'--- 2. SAYISAL KONTROL (Eðer doluysa sayý mý?) ---
'Alýþ fiyatý doluysa ve sayý deðilse hata ver
If txtalis.Value <> "" And Not IsNumeric(txtalis.Value) Then
    MsgBox "Alýþ fiyatý sayýsal bir deðer olmalýdýr!", vbCritical
    txtalis.SetFocus
    Exit Sub
End If

'Satýþ fiyatý doluysa ve sayý deðilse hata ver
If txtsatis.Value <> "" And Not IsNumeric(txtsatis.Value) Then
    MsgBox "Satýþ fiyatý sayýsal bir deðer olmalýdýr!", vbCritical
    txtsatis.SetFocus
    Exit Sub
End If




'--- 3. ONAY MEKANÝZMASI ---
If lblislem.Caption = "Yeni" Then
sor = MsgBox("Stok Kaydedilsin mi?", vbQuestion + vbYesNo, "KAYDET")
If sor = vbNo Then Exit Sub

'--- 4. BOÞ SATIR BULMA VE KAYIT ---
'Döngü yerine daha hýzlý olan 'End(xlUp)' yöntemini kullanalým
X = Sheets("Stok").Cells(Rows.Count, "A").End(xlUp).Row + 1

With Sheets("Stok")
    .Range("A" & X).Value = UCaseTR(txtstokkodu.Value)
    .Range("B" & X).Value = UCaseTR(txtaciklama.Value)
    .Range("C" & X).Value = UCaseTR(cbbirim.Value)
    .Range("D" & X).Value = txtalis.Value 'Fiyatlarda ucasetr kullanmaya gerek yok
    .Range("E" & X).Value = txtsatis.Value
    .Range("I" & X).Value = cbkdv.Value
    
End With

'--- 5. BÝTÝÞ ÝÞLEMLERÝ ---
btniptal_Click 'Formu temizlemek için
frmmesaj.lblmesaj.Caption = "Stok Kaydedildi ... "
Sheets("Tanimlamalar").Range("C2").Value = Sheets("Tanimlamalar").Range("C2").Value + 1
Unload frmstoktanimlama
frmstoklistesi.lststoklar.RowSource = ""
frmstoklistesi.StoklariListele
frmmesaj.Show
Exit Sub
End If

If lblislem.Caption = "Düzeltme" Then 'Hocanýn yazdýðý gibi büyük/küçük harfe dikkat et
    sor = MsgBox("Stok Güncellensin mi?", vbQuestion + vbYesNo, "GÜNCELLE")
    If sor = vbNo Then Exit Sub

    '--- MEVCUT KAYDI BULMA ---
    Dim bul As Range
    'Stok sayfasýnýn A sütununda txtStokKodu'nu arýyoruz
    Set bul = Sheets("Stok").Range("A:A").Find(What:=txtstokkodu.Value, LookIn:=xlValues, LookAt:=xlWhole)

    If Not bul Is Nothing Then
        'Eðer stok kodu bulunduysa o satýrý X deðiþkenine ata
        X = bul.Row
    Else
        'Eðer stok kodu bulunamazsa (hata olmamasý için) uyarý ver ve çýk
        MsgBox "Güncellenecek stok kodu bulunamadý!", vbCritical
        Exit Sub
    End If

    '--- GÜNCELLEME ÝÞLEMÝ ---
    With Sheets("Stok")
        .Range("A" & X).Value = UCaseTR(txtstokkodu.Value)
        .Range("B" & X).Value = UCaseTR(txtaciklama.Value)
        .Range("C" & X).Value = UCaseTR(cbbirim.Value)
        .Range("D" & X).Value = txtalis.Value
        .Range("E" & X).Value = txtsatis.Value
        .Range("I" & X).Value = cbkdv.Value
    End With

    '--- BÝTÝÞ ÝÞLEMLERÝ ---
    btniptal_Click
    frmmesaj.lblmesaj.Caption = "Stok Güncellendi ... "

    Unload frmstoktanimlama
    'Liste yenileme iþlemleri
    frmstoklistesi.lststoklar.RowSource = ""
    frmstoklistesi.StoklariListele
    frmmesaj.Show
End If
End Sub
Private Sub btniptal_Click()

Unload Me

End Sub
Private Sub btnkapat_Click()
Unload Me
End Sub






