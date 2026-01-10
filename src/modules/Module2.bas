Attribute VB_Name = "Module2"
Public Sub ExportVBA_To_GitHub()
    Dim comp As Object
    Dim basePath As String
    
    basePath = "C:\Users\arif\Documents\Excel Github\excel-vba-ai-project\src\"
    
    For Each comp In ThisWorkbook.VBProject.VBComponents
        Select Case comp.Type
            Case 1 ' Standard Module
                comp.Export basePath & "modules\" & comp.Name & ".bas"
            
            Case 2 ' Class Module
                comp.Export basePath & "classes\" & comp.Name & ".cls"
            
            Case 3 ' UserForm
                comp.Export basePath & "forms\" & comp.Name & ".frm"
            
            Case 100 ' Sheets & ThisWorkbook
                comp.Export basePath & "sheets\" & comp.Name & ".cls"
        End Select
    Next comp
    
    MsgBox "VBA kodlarý GitHub klasörüne export edildi.", vbInformation
End Sub

Public Sub Export_Commit_Push()
    ' 1. Önce tüm VBA kodlarýný export et
    Call ExportVBA_To_GitHub
    
    ' 2. Git commit + push yapan bat dosyasýný çalýþtýr
    Shell "cmd /c ""C:\Users\arif\Documents\Excel Github\excel-vba-ai-project\push_vba.bat""", vbHide
    
    MsgBox "Export + Git commit + push tamamlandý.", vbInformation
End Sub

