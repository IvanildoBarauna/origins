Attribute VB_Name = "HIPERLINKS_BOTÕES"
Option Private Module
Sub LinkPS()

Plan3.Select
Range("A3").Select

End Sub

Public Sub Menu()

Home.Select
Range("A15").Select

End Sub

Sub ViraramPesquisas()

Dim USER As String

USER = Environ("USERNAME")

Select Case USER

'Autorizados na abertura do LOG

      'Junior&Cris = PC Pessoal
      'ijuni002 = Ivanildo
      'vsous001 = Super Vinicius
      'dsilv060 = Coordinator Daniela
      'rsouz023 = Super Reginaldo
      'mmelo002 = Super Melo | Retirado Excessão
      'jsilv061 = Super Jefte

Case Is = "Junior&Cris", "jsilv061", "ijuni002", "vsous001", "dsilv060", "rsouz023"


Plan6.Select
Range("A3").Select

Case Else

LOG "TENTATIVA DE ACESSO A BASE DE [VIRARAM PESQUISAS]"

MsgBox "ACESSO NÃO PERMITIDO", vbCritical, AppName

End Select

End Sub

Public Sub AbrirLog()
Attribute AbrirLog.VB_ProcData.VB_Invoke_Func = "I\n14"

Dim USER As String

USER = Environ("USERNAME")

Select Case USER

'Autorizados na abertura do LOG

      'Junior&Cris = PC Pessoal
      'ijuni002 = Ivanildo
      'vsous001 = Super Vinicius
      'dsilv060 = Coordinator Daniela
      'rsouz023 = Super Reginaldo
      'mmelo002 = Super Melo | Retirado Excessão
      'jsilv061 = Super Jefte

Case Is = "Junior&Cris", "jsilv061", "ijuni002", "vsous001", "dsilv060", "rsouz023"

'CTRL+SHIFT+I

shtLOG.Visible = True
shtLOG.Activate
shtLOG.Range("a2").Select

MsgBox "Olá, " & USER & "! Bem vindo ao LOG de Atividades da Planilha", vbInformation, AppName

Case Else

LOG "TENTATIVA DE ACESSO AO LOG"

MsgBox "ACESSO NÃO PERMITIDO", vbCritical, AppName

End Select

End Sub

Sub FecharLog()
Attribute FecharLog.VB_ProcData.VB_Invoke_Func = "O\n14"

'CTRL+SHIFT+O

If shtLOG.FilterMode Then

ActiveSheet.ShowAllData

End If

shtLOG.Visible = False
Call Menu

MsgBox "Log de Atividades Fechado!", vbInformation, AppName

End Sub

Public Sub ACTIVATE_()

With Application
    
    .AskToUpdateLinks = False
    .DisplayAlerts = False
    .ScreenUpdating = False

End With

End Sub

Public Sub DEACTIVATE_()

With Application

    .AskToUpdateLinks = True
    .DisplayAlerts = True
    .ScreenUpdating = True

End With

End Sub
