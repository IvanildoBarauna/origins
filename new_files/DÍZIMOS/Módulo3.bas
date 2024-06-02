Attribute VB_Name = "Módulo3"
Sub copiar()
Attribute copiar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' copiar Macro
'

'
    Range("Tabela2[[#All],[NOME DIZIMISTA/OFERTANTE]:[DESCRIÇÃO]]").Select
    Selection.Copy
    Sheets("RELATÓRIO").Select
    Application.Run "CADASTRO_MEMBROS.xlsm!SalvamentoProgramado"
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "ID"
    Range("A3").Select
    Selection.ClearContents
    Range("A3").Select
    ActiveCell.FormulaR1C1 = "NOME"
    Range("A3").Select
    Selection.ClearContents
    Range("B3").Select
    ActiveCell.FormulaR1C1 = "NOME"
    Range("C3").Select
End Sub
