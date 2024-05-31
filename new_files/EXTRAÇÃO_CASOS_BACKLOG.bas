Attribute VB_Name = "EXTRAÇÃO_CASOS_BACKLOG"
Public Sub EXTRAÇÃO_CASOS()
Attribute EXTRAÇÃO_CASOS.VB_ProcData.VB_Invoke_Func = "E\n14"

ACTIVATE_

Application.StatusBar = "Limpando dados ... "

'Limpar Base de Fupo'

 Plan4.Select
 
 On Error Resume Next
 
 Range("CASOS_FUPO[CASE ID], CASOS_FUPO[STATUS ATUAL], CASOS_FUPO[OBSERVAÇÃO]").ClearContents
 Rows("4:4").Select
 Range(Selection, Selection.End(xlDown)).Select
 Selection.Delete
 Range("A3").Select
 
 'Limpar BACKLOG BASE
 
 Plan3.Select
 
 Rows("3:3").Select
 Range(Selection, Selection.End(xlDown)).Select
 Selection.Delete
 Range("B2:Q2").ClearContents
 Range("B2").Select
 
 'LIMPAR CALLBACK SLA'
 
  Plan2.Select
 
 Rows("3:3").Select
 Range(Selection, Selection.End(xlDown)).Select
 Selection.Delete
 Range("A2:M2").ClearContents
 Range("A2").Select
 
 'LIMPAR BASE MARCOS
 
   Plan6.Select
 
 Rows("3:3").Select
 Range(Selection, Selection.End(xlDown)).Select
 Selection.Delete
 Range("A2:G2").ClearContents
 

    'Abrir o arquivo CALLBACK SLA para extração'
    
    Application.StatusBar = "Extraindo casos agendados ... "

ChDir ("C:\Users\ijuni002\Desktop")
Workbooks.Open Filename:="C:\Users\ijuni002\Desktop\CALLBACK_SLA.xlsx"

'Ativar janela do arquivo CALLBACK SLA'

Windows("CALLBACK_SLA.xlsx").Activate

'Extração de dados do arquivo ativo'

Range("A2:M2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Windows("BASE_CASOS").Activate
Plan2.Range("A2").PasteSpecial xlPasteValuesAndNumberFormats


        Range("A1").Select
        
        'Abrir arquivo BACKLOG GERAL para extração'
        
        Application.StatusBar = "Extraindo casos sem SLA ... "

ChDir ("C:\Users\ijuni002\Desktop")
Workbooks.Open Filename:="C:\Users\ijuni002\Desktop\BACKLOG_GERAL.xlsx"

' Ativar janela do arquivo'

Windows("BACKLOG_GERAL.xlsx").Activate

'Cópia dos dados'

Range("A2:P2").Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Copy
Windows("BASE_CASOS").Activate
Plan3.Range("B2").PasteSpecial xlPasteValuesAndNumberFormats

Application.CutCopyMode = False

'Abrir arquivo Backlog WEB'

Application.StatusBar = "Extraindo casos de WEB ... "
        
 ChDir ("C:\Users\ijuni002\Desktop")
 Workbooks.Open Filename:="C:\Users\ijuni002\Desktop\BACKLOG_WEB.xlsx"
 
 'Ativar janela e extrair dados'
        
        Windows("BACKLOG_WEB.xlsx").Activate
        
        If Range("A2").Value <> "" Then
        
        Range("A2:P2").Select
        
        If Range("a3").Value <> "" Then
        Range(Selection, Selection.End(xlDown)).Select
        Selection.Copy
        
        Else
         
         Range("a2:p2").Copy
         
         End If
         
        Windows("BASE_CASOS").Activate
        
        'Neste ponto o código seleciona a célula A1048576 e sobe até a primeira célula com dados'
        Sheets("BACKLOG_BASE").Select
        Range("B1048576").End(xlUp).Select
        
        ' Após selecionada a última célula com dados o código desce uma linha para inserir os novos dados, fazendo então, a concatenação'
        ActiveCell.Offset(1, 0).Select
        ActiveCell.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
        
        Application.CutCopyMode = False
        
        Else
        
        On Error Resume Next
        
        End If
        
        'Copiar uma célula em branco e colar apenas os valores desta célula com a operação de adição.
'Isso resulta em "limpar" o formato em que o dado foi recebido e adiciona o formato "nenhum" ao intervalo em questão'
'Exemplo: Se um valor foi recebido como texto e precisa ser inserido como número, fazemos esta operação na tentativa de limpar o formato atual'
        
        'Neste ponto, o código classifica a planilha Backlog Base na coluna V que é a dos CALL BACKS e de A a Z para que os números fiquem primeiro'
        
        Application.StatusBar = "Organizando dados ... "
        
        Call teste
        
   Call ClassificarA_Z
    
    'Fechando arquivos em aberto, usados para extração'
    
    Application.StatusBar = "Fechando fonte de dados ... "
    
        
        Workbooks("CALLBACK_SLA.xlsx").Close False
        Workbooks("BACKLOG_GERAL.xlsx").Close False
        Workbooks("BACKLOG_WEB.xlsx").Close False
        
        
        
     'Estas instruções inserem os dados da base do Backlog já extraída para a Base de Fupo'
     
     Application.StatusBar = "Consolidando casos para tratativa, aguarde ... "
     
     Windows("BASE_CASOS").Activate

    Plan3.Select
    
        Selection.AutoFilter
        Application.CommandBars("Selection and Visibility").Visible = False
        ActiveSheet.ListObjects("BACKLOG_BASE").Range.AutoFilter Field:=8, Criteria1 _
        :="Pendente de retorno de chamada"
    
    Range("C2").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Copy
    Plan4.Select
    Range("A3").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    
    Range("A3").Select
    
    Plan3.Select
    
    Application.CutCopyMode = False
    
    Call teste
    
   
    
    Range("A1").Select
    
    Workbooks("Callback SLA.xlsx").Save
    Workbooks("Callback SLA.xlsx").Close
    
    Application.StatusBar = "Processo Concluído!"

    Sheets("CAPA").Select
    Range("A1").Select
    
    
    'SALVA O ARQUIVO'
    
    Application.StatusBar = "Calculando fórmulas e salvando BASE_CASOS ... "
    
    Calculate
    
    Application.StatusBar = False

    ActiveWorkbook.Save
    
      DEACTIVATE_
      
      CAPA
        
    MsgBox "EXTRAÇÃO DE DADOS CONCLUÍDA COM SUCESSO", vbInformation, AppName
    
End Sub

Sub teste()

'Verificar se a tabela está filtrada'

Windows("BASE_CASOS.xlsm").Activate
Plan3.Select
If ActiveSheet.FilterMode Then

ActiveSheet.ShowAllData

End If

End Sub


Sub ClassificarA_Z()


Plan3.Select

Call teste

    Range("BACKLOG_BASE[[#Headers],[Callback Tempo Programado (Agente)]]").Select
    ActiveWorkbook.Worksheets("BACKLOG_BASE").ListObjects("BACKLOG_BASE").Sort. _
        SortFields.Clear
    ActiveWorkbook.Worksheets("BACKLOG_BASE").ListObjects("BACKLOG_BASE").Sort. _
        SortFields.Add Key:=Range( _
        "BACKLOG_BASE[[#All],[Callback Tempo Programado (Agente)]]"), SortOn:= _
        xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("BACKLOG_BASE").ListObjects("BACKLOG_BASE").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("BACKLOG_BASE[[#Headers],[Callback Tempo Programado (Agente)]]").Select
    
    
End Sub

