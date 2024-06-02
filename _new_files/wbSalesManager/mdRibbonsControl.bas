Attribute VB_Name = "mdRibbonsControl"
Option Explicit

Sub RibbonCallBack(xControl As IRibbonControl)
    Dim CallName As String
    
    CallName = "'" & ThisWorkbook.Name & "'!" & "mdRibbonsControl." & xControl.ID
    
    Application.Run CallName
End Sub

Sub btnLancamentos()
    shCaixa.Activate
    frmlançamentos.Show False
End Sub

Sub btnPedidos(): shPedidos.Activate: End Sub

Sub btContagem()
    wsContagem.Activate
    Call mdInserções.Inserir
End Sub

Sub btClear(): mdInserções.ClearData: End Sub
Sub btFechamento(): Call mdFechamentos.InserirFechamento: End Sub
