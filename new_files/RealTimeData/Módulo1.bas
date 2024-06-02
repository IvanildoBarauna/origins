Attribute VB_Name = "Módulo1"
Sub Pos()
Attribute Pos.VB_ProcData.VB_Invoke_Func = " \n14"
    ThisWorkbook.Sheets("Posição de Custódia").Range("C7").Select
End Sub
Sub acom()
Attribute acom.VB_ProcData.VB_Invoke_Func = " \n14"
    ThisWorkbook.Sheets("Acompanhamento de mercado").Range("F6").Select
End Sub
Sub voltar()
Attribute voltar.VB_ProcData.VB_Invoke_Func = " \n14"
    ThisWorkbook.Sheets("RTD").Range("F14").Select
End Sub
Sub Macro1()
    ThisWorkbook.Sheets("Relatórios").Range("K6").Select
End Sub

Sub trader()
    ThisWorkbook.Sheets("Planilha do Trader").Range("J4").Select
End Sub

Sub IRDayTrade()
    ThisWorkbook.Sheets("IR Day Trade").Range("D2:I3").Select
End Sub
