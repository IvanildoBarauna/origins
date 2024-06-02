Attribute VB_Name = "EXTRAÇÃO_CASOS"
Public Const AppName As String = "| FORÇA TAREFA QUALIDADE |"

Option Private Module

Sub ExtraçãoDeDados()

ACTIVATE_

Dim USER As String

USER = Environ("USERNAME")

'LIMPEZA DOS DADOS'

Application.StatusBar = "Limpando dados ... "

Plan2.Select
Rows("6:6").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete

  LOG "CONTEÚDO DE BASE GERAL EXCLUÍDO"


'ABRIR OS ARQUIVOS'

Application.StatusBar = "Abrindo fontes de dados ... "

'ABRIR PLAN  IPG'

ChDir ("\\10.166.108.10\users\Public\Documents\Equipe Callback\IPG")
Workbooks.Open Filename:="\\10.166.108.10\users\Public\Documents\Equipe Callback\IPG\Casos IPG - Janeiro.xlsx"

'ABRIR PLAN  PSG'

ChDir ("\\10.166.108.10\users\Public\Documents\Equipe Callback\PSG")
Workbooks.Open Filename:="\\10.166.108.10\users\Public\Documents\Equipe Callback\PSG\Casos PSG - Janeiro.xlsx"

'Retira filtro'

Call Filtro

Application.StatusBar = "Consolidando casos..."

'RETIRAR DADOS IPG'

'EXIBIR COLUNAS OCULTAS'

Windows("Casos IPG - Janeiro").Activate
Sheets("Base").Select
Cells.Select
    Selection.EntireColumn.Hidden = False
    
    
   'FILTRO PARA TIRAR AS LINHAS SEM CASOS'
    
    Sheets("Base").Range("A1:S1").AutoFilter Field:=4, Criteria1:="<>"
    
    'Cópia dos dados
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    'Consolidar
    
    Windows("FORÇA TAREFA - QUALIDADE HPC PCS.xlsm").Activate
    Sheets("BASE_GERAL").Select
   Range("A6").PasteSpecial xlPasteValuesAndNumberFormats
   
        
      'RETIRAR DADOS PSG'

'EXIBIR COLUNAS OCULTAS'

Windows("Casos PSG - Janeiro").Activate
Sheets("Base").Select
    Cells.Select
    Selection.EntireColumn.Hidden = False
    
    
   'FILTRO PARA TIRAR AS LINHAS SEM CASOS'
    
        Sheets("Base").Range("A1:V1").AutoFilter Field:=4, Criteria1:="<>"
    
      'Cópia dos dados
    
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    

    Windows("FORÇA TAREFA - QUALIDADE HPC PCS.xlsm").Activate
    Sheets("BASE_GERAL").Select
     Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial xlPasteFormulasAndNumberFormats
        
        LOG "BASE DE CALL BACK EXTRAÍDA"

   Application.StatusBar = "Salvando o arquivo ... "
   
  Workbooks("Casos IPG - Janeiro").Close False
  Workbooks("Casos PSG - Janeiro").Close False
   
   Application.StatusBar = False
   
    Home.Activate

    Calculate
   
    DEACTIVATE_
   
    MsgBox "EXTRAÇÃO DE CASOS CONCLUÍDA COM SUCESSO!", vbInformation, AppName

    
End Sub


Sub Filtro()

'Verificar se a tabela está filtrada'

Windows("Casos IPG - Janeiro").Activate
If Sheets("Base").FilterMode Then

Sheets("Base").ShowAllData

End If

Windows("Casos PSG - Janeiro").Activate
If Sheets("Base").FilterMode Then

Sheets("Base").ShowAllData

End If

End Sub

