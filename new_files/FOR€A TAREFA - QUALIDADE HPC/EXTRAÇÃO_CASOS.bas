Attribute VB_Name = "EXTRAÇÃO_CASOS"
Option Private Module

Sub ExtraçãoDeDados()
Attribute ExtraçãoDeDados.VB_ProcData.VB_Invoke_Func = " \n14"

'LIMPEZA DOS DADOS'

Application.StatusBar = "Limpando dados ... "

  Sheets("BASE_GERAL").Select
  Range("A6:S5000").ClearContents
  
  LOG "CONTEÚDO DE BASE GERAL EXCLUÍDO"
  
  Range("a6").Select

'ABRIR OS ARQUIVOS'

Application.StatusBar = "Abrindo fontes de dados ... "

'ABRIR PLAN DO JEFTE'

ChDir ("\\10.166.108.10\users\Public\Documents\Equipe Callback\Jefte")
Workbooks.Open Filename:="\\10.166.108.10\users\Public\Documents\Equipe Callback\Jefte\Jefte Novembro1.xlsx"

'ABRIR PLAN DO LOULA'

ChDir ("\\10.166.108.10\users\Public\Documents\Equipe Callback\Vinicius")
Workbooks.Open Filename:="\\10.166.108.10\users\Public\Documents\Equipe Callback\Vinicius\Vinicius Novembro.xlsx"

'ABRIR PLAN DO MELO'

ChDir ("\\10.166.108.10\users\Public\Documents\Equipe Callback\Melo")
Workbooks.Open Filename:="\\10.166.108.10\users\Public\Documents\Equipe Callback\Melo\Melo Novembro.xlsx"

'ABRIR PLAN DO REGINALDO'

ChDir ("\\10.166.108.10\users\Public\Documents\Equipe Callback\Reginaldo")
Workbooks.Open Filename:="\\10.166.108.10\users\Public\Documents\Equipe Callback\Reginaldo\Reginaldo Novembro.xlsx"

'ABRIR PLAN DO COMMERCIAL'

ChDir ("\\10.166.108.10\users\Public\Documents\Equipe Callback\Comercial")
Workbooks.Open Filename:="\\10.166.108.10\users\Public\Documents\Equipe Callback\Comercial\Comercial Novembro.xlsx"


'Retira filtro'

Call Filtro

Application.StatusBar = "Consolidando casos da equipe do Jefte ..."

'RETIRAR DADOS JEFTE'

'EXIBIR COLUNAS OCULTAS'

Windows("Jefte Novembro1.xlsx").Activate
Sheets("Base").Select
Cells.Select
    Selection.EntireColumn.Hidden = False
    
    'Exclui coluna 2º Retorno'
    
    Columns("T:T").Select
    Selection.Delete
    
    'FILTRO PARA TIRAR AS LINHAS SEM CASOS'
    
    Range("a1").Select
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=4, Criteria1:="<>"
    
      'FILTRO IPG'
      
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="IPG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
   Range("A5").PasteSpecial xlPasteValuesAndNumberFormats
   
        
         'FILTRO PSG'
         
         Windows("Jefte Novembro1.xlsx").Activate
       Sheets("Base").Select
        
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="PSG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
    Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
        
        
        'RETIRAR DADOS DO LOULA'

'EXIBIR COLUNAS OCULTAS'

Windows("Vinicius Novembro.xlsx").Activate
Sheets("Base").Select
Cells.Select
    Selection.EntireColumn.Hidden = False
    
    'Exclui coluna 2º Retorno'
    
    Columns("T:T").Select
    Selection.Delete
    
    'FILTRO PARA TIRAR AS LINHAS SEM CASOS'
    
    Range("a1").Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=4, Criteria1:="<>"
    
      'FILTRO IPG'
      
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="IPG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
   Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
   
        
         'FILTRO PSG'
         
         Windows("Vinicius Novembro.xlsx").Activate
       Sheets("Base").Select
        
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="PSG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
    Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
        
            'RETIRAR DADOS DO REGINALDO'

'EXIBIR COLUNAS OCULTAS'

Windows("Reginaldo Novembro.xlsx").Activate
Sheets("Base").Select
Cells.Select
    Selection.EntireColumn.Hidden = False
    
    'Exclui coluna 2º Retorno'
    
    Columns("T:T").Select
    Selection.Delete
    
    'FILTRO PARA TIRAR AS LINHAS SEM CASOS'
    
    Range("a1").Select
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=4, Criteria1:="<>"
    
      'FILTRO IPG'
      
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="IPG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
   Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
   
        
         'FILTRO PSG'
         
         Windows("Reginaldo Novembro.xlsx").Activate
       Sheets("Base").Select
        
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="PSG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
    Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
        
        
        'RETIRAR DADOS DO MELO'

'EXIBIR COLUNAS OCULTAS'

Windows("Melo Novembro.xlsx").Activate
Sheets("Base").Select
Cells.Select
    Selection.EntireColumn.Hidden = False
    
    'Exclui coluna 2º Retorno'
    
    Columns("T:T").Select
    Selection.Delete
    
    'FILTRO PARA TIRAR AS LINHAS SEM CASOS'
    
    Range("a1").Select
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=4, Criteria1:="<>"
    
      'FILTRO IPG'
      
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="IPG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
   Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
   
        
         'FILTRO PSG'
         
         Windows("Melo Novembro.xlsx").Activate
       Sheets("Base").Select
        
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="PSG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
    Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
        
        
        
        'RETIRAR DADOS DE COMMERCIAL'

'EXIBIR COLUNAS OCULTAS'

Windows("Comercial Novembro.xlsx").Activate
Sheets("Base").Select
Cells.Select
    Selection.EntireColumn.Hidden = False
    
    'Exclui coluna 2º Retorno'
    
    Columns("T:T").Select
    Selection.Delete
    
    'FILTRO PARA TIRAR AS LINHAS SEM CASOS'
    
    Range("a1").Select
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=4, Criteria1:="<>"
    
      'FILTRO IPG'
      
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="IPG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
   Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
   
        
         'FILTRO PSG'
         
         Windows("Comercial Novembro.xlsx").Activate
       Sheets("Base").Select
        
    ActiveSheet.Range("$A$1:$S$2000").AutoFilter Field:=18, Criteria1:="PSG"
    Range("A1").Select
    Range(Selection, Selection.End(xlToRight)).Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    
    Windows("FORÇA TAREFA - QUALIDADE HPC.xlsm").Activate
    Sheets("BASE_GERAL").Select
    Range("A1048576").End(xlUp).Select
    ActiveCell.Offset(1, 0).Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
        
        Range("a5").Select
        
        
   Workbooks("Jefte Novembro1.xlsx").Close
   Workbooks("Vinicius Novembro.xlsx").Close
   Workbooks("Melo Novembro.xlsx").Close
   Workbooks("Reginaldo Novembro.xlsx").Close
   Workbooks("Comercial Novembro.xlsx").Close
   
 
   

   Application.StatusBar = "Salvando o arquivo ... "
   
   ActiveWorkbook.Save
   
   Application.StatusBar = False
   
   Sheets("home").Select
   Range("B14").Select
   
   Calculate
   
   LOG "BASE DE CALLBACK EXTRAÍDA"
   
   MsgBox "EXTRAÇÃO DE CASOS CONCLUÍDA COM SUCESSO"

    
End Sub



Sub Filtro()

'Verificar se a tabela está filtrada'


Windows("Jefte Novembro1.xlsx").Activate
Sheets("Base").Select
If ActiveSheet.FilterMode Then

ActiveSheet.ShowAllData

End If

Windows("Vinicius Novembro.xlsx").Activate
Sheets("Base").Select
If ActiveSheet.FilterMode Then

ActiveSheet.ShowAllData

End If

Windows("Vinicius Novembro.xlsx").Activate
Sheets("Base").Select
If ActiveSheet.FilterMode Then

ActiveSheet.ShowAllData

End If

Windows("Melo Novembro.xlsx").Activate
Sheets("Base").Select
If ActiveSheet.FilterMode Then

ActiveSheet.ShowAllData

End If

Windows("Reginaldo Novembro.xlsx").Activate
Sheets("Base").Select
If ActiveSheet.FilterMode Then

ActiveSheet.ShowAllData

End If

Windows("Comercial Novembro.xlsx").Activate
Sheets("Base").Select
If ActiveSheet.FilterMode Then

ActiveSheet.ShowAllData

End If


End Sub




 
    
    
    
