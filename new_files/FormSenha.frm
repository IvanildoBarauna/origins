Attribute VB_Name = "FormSenha"
Attribute VB_Base = "0{C46AD212-8FED-4A07-9343-C2B2EFF0E946}{07253684-5801-4058-8E16-4FFA68B1359C}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnsair_Click()
Unload FormSenha
End Sub

Private Sub btOk_Click()
If txtsenha = "vdpsafra2016" Then
    
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name <> "HOME" Then
            ws.Visible = xlSheetVisible 'Mostar todas
        End If
    Next

    
    
    Unload FormSenha
Else
    MsgBox "Senha digitada incorreta. Verifique!", vbInformation, "Informação"
    Exit Sub
 End If
 
End Sub

