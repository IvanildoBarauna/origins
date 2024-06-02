Attribute VB_Name = "mdMain"
Option Explicit
Option Private Module
Const MASTER_USER As String = "DIRETOR"
Const MASTER_PASS As String = "DIR123"

Private Function isValidUser(pUser As String) As Boolean
    If pUser = MASTER_USER Then isValidUser = True
End Function

Private Function isValidPass(pPass As String) As Boolean
    If pPass = MASTER_PASS Then isValidPass = True
End Function

Public Sub ShowProducts()
    Dim InformedUser    As String
    Dim informedPass    As String
    Dim UserValid       As Boolean, PassValid As Boolean
    
    InformedUser = Excel.Application.InputBox("Informe o usuário de acesso", "Acesso a base de produtos.", Type:=2)
    
    UserValid = isValidUser(InformedUser)
    
    If UserValid = False Then
        MsgBox "Usuário inválido, acesso negado!", vbCritical, "Valida Acesso"
        Exit Sub
    End If
    
    informedPass = Excel.Application.InputBox("Senha?", "Acesso a base de produtos.", Type:=2)
    
    PassValid = isValidPass(informedPass)
    
    If PassValid = False Then
        MsgBox "Senha inválida, acesso não permitido!", vbCritical, "Valida Acesso"
        Exit Sub
    End If
    
    wsMarkups.Visible = PassValid
    wsProdutos.Visible = PassValid
    
    Application.OnTime VBA.Now() + VBA.TimeValue("00:00:10"), "ValidaSheet"
End Sub

Sub ValidaSheet()
    If Not ActiveSheet.CodeName = "wsProdutos" And Not ActiveSheet.CodeName = "wsMarkups" Then
        wsMarkups.Visible = xlSheetVeryHidden
        wsProdutos.Visible = xlSheetVeryHidden
        Exit Sub
    End If

    If Not ActiveSheet.CodeName = "wsProdutos" Then
        wsProdutos.Visible = xlSheetVeryHidden
    Else
        wsMarkups.Visible = xlSheetVeryHidden
    End If
End Sub
