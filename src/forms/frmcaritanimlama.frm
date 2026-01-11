VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmcaritanimlama 
   Caption         =   "Cari Tanýmlama Formu"
   ClientHeight    =   8760.001
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9435.001
   OleObjectBlob   =   "frmcaritanimlama.frx":0000
End
Attribute VB_Name = "frmcaritanimlama"
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

Private Sub txtadunvan_Change()

End Sub

Private Sub txtemail_Change()

End Sub

'Bu userformu ana formun tam ortasýna sabitler ve gerekli ayarlamalar yapýlýr.
Private Sub UserForm_Initialize()
    ' Yatayda (Soldan saða) ortalamak için:
    Me.Left = frmcarilistesi.Left + (frmAnaForm.Width - Me.Width) / 2
    
    ' Dikeyde (Yukarýdan aþaðýya) ortalamak için:
    Me.Top = frmcarilistesi.Top + (frmAnaForm.Height - Me.Height) / 2
    
    
Frame1.Enabled = True
btnkaydet.Enabled = True
btniptal.Enabled = True
txtadunvan.SetFocus 'Her yeni kayýt yapýldýktan sonra aciklama isimli textboxa odaklanýr.
txtcarikodu.Value = "CR00000" & Sheets("Tanimlamalar").Range("D2").Value + 1 'Kayýt sýrasýnda cari kodu otomatik olarak CR00000 + Tanýmlamar sayfasýndaki deðerin bir fazlasý gelir.
End Sub

Private Sub btnKaydet_Click()
Dim X As Long
Dim sor As Byte

'--- 1. KRÝTÝK ALAN KONTROLÜ (Zorunlu alanlar) ---
If txtcarikodu.Value = "" Or txtadunvan.Value = "" Or txttelefon.Value = "" Or txtadres.Value = "" Then
    MsgBox "Lütfen zorunlu alanlarý (Cari Kodu, AdÜnvan, Telefon, Adres) doldurunuz!", vbExclamation
    Exit Sub
End If

'--- 2. SAYISAL KONTROL (Eðer doluysa sayý mý?) ---
'Telefon doluysa ve sayý deðilse hata ver
If txttelefon.Value <> "" And Not IsNumeric(txttelefon.Value) Then
    MsgBox "Telefon no sayýsal bir deðer olmalýdýr!", vbCritical
    txttelefon.SetFocus
    Exit Sub
End If

'Vergi Numarasý doluysa ve sayý deðilse hata ver
If txtvergino.Value <> "" And Not IsNumeric(txtvergino.Value) Then
    MsgBox "Vergi Numarasý sayýsal bir deðer olmalýdýr!", vbCritical
    txtvergino.SetFocus
    Exit Sub
End If




'--- 3. ONAY MEKANÝZMASI ---
If lblislem.Caption = "Yeni" Then
sor = MsgBox("Cari Kaydedilsin mi?", vbQuestion + vbYesNo, "KAYDET")
If sor = vbNo Then Exit Sub

'--- 4. BOÞ SATIR BULMA VE KAYIT ---
'Döngü yerine daha hýzlý olan 'End(xlUp)' yöntemini kullanalým
X = Sheets("Cari").Cells(Rows.Count, "A").End(xlUp).Row + 1

With Sheets("Cari")
    .Range("A" & X).Value = UCaseTR(txtcarikodu.Value)
    .Range("B" & X).Value = UCaseTR(txtadunvan.Value)
    .Range("C" & X).Value = UCaseTR(txtvergidairesi.Value)
    .Range("D" & X).Value = txtvergino.Value 'Fiyatlarda ucasetr kullanmaya gerek yok
    .Range("E" & X).Value = txttelefon.Value
    .Range("F" & X).Value = LCaseTR(txtemail.Value)
    .Range("G" & X).Value = UCaseTR(txtadres.Value)
    
End With

'--- 5. BÝTÝÞ ÝÞLEMLERÝ ---
btniptal_Click 'Formu temizlemek için
frmmesaj.lblmesaj.Caption = "Cari Kaydedildi ... "
Sheets("Tanimlamalar").Range("D2").Value = Sheets("Tanimlamalar").Range("D2").Value + 1
Unload frmcaritanimlama
frmcarilistesi.lstcariler.RowSource = ""
frmcarilistesi.CariListele
frmmesaj.Show
Exit Sub
End If

If lblislem.Caption = "Düzeltme" Then 'Hocanýn yazdýðý gibi büyük/küçük harfe dikkat et
    sor = MsgBox("Cari Güncellensin mi?", vbQuestion + vbYesNo, "GÜNCELLE")
    If sor = vbNo Then Exit Sub

    '--- MEVCUT KAYDI BULMA ---
    Dim bul As Range
    'Cari sayfasýnýn A sütununda txtcariKodu'nu arýyoruz
    Set bul = Sheets("Cari").Range("A:A").Find(What:=txtcarikodu.Value, LookIn:=xlValues, LookAt:=xlWhole)

    If Not bul Is Nothing Then
        'Eðer cari kodu bulunduysa o satýrý X deðiþkenine ata
        X = bul.Row
    Else
        'Eðer cari kodu bulunamazsa (hata olmamasý için) uyarý ver ve çýk
        MsgBox "Güncellenecek cari kodu bulunamadý!", vbCritical
        Exit Sub
    End If

    '--- GÜNCELLEME ÝÞLEMÝ ---
    With Sheets("Cari")
    .Range("A" & X).Value = UCaseTR(txtcarikodu.Value)
    .Range("B" & X).Value = UCaseTR(txtadunvan.Value)
    .Range("C" & X).Value = UCaseTR(txtvergidairesi.Value)
    .Range("D" & X).Value = txtvergino.Value 'Fiyatlarda ucasetr kullanmaya gerek yok
    .Range("E" & X).Value = txttelefon.Value
    .Range("F" & X).Value = LCaseTR(txtemail.Value)
    .Range("G" & X).Value = UCaseTR(txtadres.Value)
    
End With


    '--- BÝTÝÞ ÝÞLEMLERÝ ---
    btniptal_Click
    frmmesaj.lblmesaj.Caption = "Cari Güncellendi ... "

    Unload frmcaritanimlama
    'Liste yenileme iþlemleri
    frmcarilistesi.lstcariler.RowSource = ""
    frmcarilistesi.CariListele
    frmmesaj.Show
End If
End Sub
Private Sub btniptal_Click()

Unload Me

End Sub
Private Sub btnkapat_Click()
Unload Me
End Sub






