Attribute VB_Name = "mdRibbonsControl"
Option Explicit
'Callback for btnLancamentos onAction
Sub btn_formLancamentos(control As IRibbonControl)
    shCaixa.Activate
    frmlançamentos.Show False
End Sub

'Callback for btnPedidos onAction
Sub btn_ViewPedidos(control As IRibbonControl): shPedidos.Activate: End Sub

'Callback for btnContagem onAction
Sub btn_Contagem(control As IRibbonControl): shContagem.Activate: End Sub

'Callback for btnClearData onAction
Sub btn_ClearData(control As IRibbonControl)
    shContagem.Activate
    Call ClearData
End Sub

'Callback for btListas onAction
Sub btn_Listas(control As IRibbonControl): sApoio.Activate: End Sub

'Callback for btTables onAction
Sub btnTables(control As IRibbonControl): wsTablesReport.Activate: End Sub

'Callback for btReport onAction
Sub btnReport(control As IRibbonControl): wsReports.Activate: End Sub
