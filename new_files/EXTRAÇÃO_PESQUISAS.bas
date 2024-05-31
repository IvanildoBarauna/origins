Attribute VB_Name = "EXTRAÇÃO_PESQUISAS"
Sub ExtraçãoBasePesquisas()
Attribute ExtraçãoBasePesquisas.VB_ProcData.VB_Invoke_Func = "Q\n14"

ACTIVATE_

Dim USER As String

USER = Environ("USERNAME")

Select Case USER

'Autorizados na extração

      'Junior&Cris = PC Pessoal
      'ijuni002 = Ivanildo
      'vsous001 = Super Vinicius
      'dsilv060 = Coordinator Daniela
      'rsouz023 = Super Reginaldo
      'mmelo002 = Super Melo | Retirado Excessão
      'jsilv061 = Super Jefte

Case Is = "Junior&Cris", "jsilv061", "ijuni002", "vsous001", "dsilv060", "rsouz023"

'ATALHO DE TECLADO CRTL+Q'

Plan5.Range("A:EY").ClearContents

LOG "CONTEÚDO DE BASE DE PESQUISAS EXCLUÍDO"

ChDir ("\\10.166.2.17\shareportal\HP-CONSUMER\Supervisores\Jefte Soares")
Workbooks.Open Filename:="\\10.166.2.17\shareportal\HP-CONSUMER\Supervisores\Jefte Soares\Qualidade.xlsx"

Windows("Qualidade.xlsx").Activate
Sheets("Base").Select
Range("A:EY").Copy
Windows("FORÇA TAREFA - QUALIDADE HPC PRINTERS.xlsm").Activate
Plan5.Range("A1").PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

Windows("Qualidade.xlsx").Activate
Sheets("Base").Select
Application.CutCopyMode = False

Workbooks("Qualidade.xlsx").Close False

Windows("FORÇA TAREFA - QUALIDADE HPC PRINTERS.xlsm").Activate

Sheets("BASE_QUALIDADE").Select
Range("A1").Select

LOG "BASE DE PESQUISAS EXTRAÍDA"

Calculate

MsgBox "PESQUISAS EXTRAÍDAS COM SUCESSO!!", vbInformation, AppName

Case Else

LOG "TENTATIVA DE ACESSO A BASE DE PESQUISAS"

MsgBox "ACESSO NÃO PERMITIDO", vbCritical, AppName

DEACTIVATE_

End Select

End Sub

