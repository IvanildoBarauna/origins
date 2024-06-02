Attribute VB_Name = "HIPERLINKS_BOTÕES"
Sub LinkPS()

Plan3.Select
Range("A4").Select

End Sub

Sub LinkPRINTERS()

Plan4.Select
Range("A4").Select

End Sub

Public Sub Menu()

Plan1.Select
Range("A1").Select

End Sub

Sub ViraramPesquisas()

Plan6.Select
Range("A3").Select

End Sub

Public Sub AbrirLog()
Attribute AbrirLog.VB_ProcData.VB_Invoke_Func = "I\n14"

'CTRL+SHIFT+I

shtLOG.Visible = True
shtLOG.Unprotect ["#P@ssw0rd1"]
shtLOG.Activate
shtLOG.Range("a2").Select

MsgBox "Olá, " & Environ("USERNAME") & "! Bem vindo ao LOG de Atividades da Planilha"

End Sub

Sub FecharLog()
Attribute FecharLog.VB_ProcData.VB_Invoke_Func = "O\n14"

'CTRL+SHIFT+O

shtLOG.Visible = False
shtLOG.Protect ["#P@ssw0rd1"]

Call Menu

MsgBox "Log de Atividades Fechado!"

End Sub
