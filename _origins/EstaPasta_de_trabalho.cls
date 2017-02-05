Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

DEACTIVATE_APP

End Sub

Private Sub Workbook_Open()

ACTIVATE_APP

Plan9.Visible = True
Plan9.Select
Range("D3").Select

Dim USER As String

USER = Environ("USERNAME")

Select Case USER

Case Is = "ijuni002", "jteno002", "jmart003"

Plan9.Visible = False
ActiveWindow.DisplayWorkbookTabs = True
Plan1.Activate
Range("e13").Select

DEACTIVATE_APP


        If Hour(Now()) > 0 And Hour(Now()) <= 12 Then
            MsgBox "Bom dia " & USER & "! Seja bem vindo(a)!", vbInformation, "Acesso Autorizado"
        End If
        If Hour(Now()) > 12 And Hour(Now()) < 18 Then
            MsgBox "Boa tarde " & USER & "! Seja bem vindo(a)!", vbInformation, "Acesso Autorizado"
        End If
        If Hour(Now()) >= 18 Then
            MsgBox "Boa noite " & USER & "! Seja bem vindo(a)!", vbInformation, "Acesso Autorizado"
        End If
        
    Case Else
    
    ActiveWorkbook.Save
    
    MsgBox "Acesso não autorizado, procure o propietário do arquivo ou o seu supervisor!", vbCritical, "Restrição de Acesso"
    
    DEACTIVATE_APP
    
    ActiveWorkbook.Close False
    
    End Select
        
End Sub
