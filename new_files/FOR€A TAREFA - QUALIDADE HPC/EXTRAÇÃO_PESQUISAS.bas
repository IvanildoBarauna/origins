Attribute VB_Name = "EXTRAÇÃO_PESQUISAS"

Sub ExtraçãoBasePesquisas()
Attribute ExtraçãoBasePesquisas.VB_ProcData.VB_Invoke_Func = "q\n14"

'ATALHO DE TECLADO CRTL+Q'

Plan5.Range("A:EY").ClearContents

LOG "CONTEÚDO DE BASE DE PESQUISAS EXCLUÍDO"

ChDir ("\\10.166.2.17\shareportal\HP-CONSUMER\Relatórios\Publicado\Qualidade\2016")
Workbooks.Open Filename:="\\10.166.2.17\shareportal\HP-CONSUMER\Relatórios\Publicado\Qualidade\2016\Consolidado Qualidade_11.xlsx"

Windows("Consolidado Qualidade_11.xlsx").Activate
Sheets("Base").Select
Range("A:EY").Copy
Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
Sheets("BASE_QUALIDADE").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

Windows("Consolidado Qualidade_11.xlsx").Activate
Sheets("Base").Select
Application.CutCopyMode = False

Workbooks("Consolidado Qualidade_11.xlsx").Close

Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate

Plan5.Range("A1").Select

LOG "BASE DE PESQUISAS EXTRAÍDA"

MsgBox "PESQUISAS EXTRAÍDAS COM SUCESSO!!"


End Sub

