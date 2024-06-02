Attribute VB_Name = "mdRibbonsControls"
Option Explicit

Sub RibbonsCallBacks(xControl As IRibbonControl)
    Dim CallName As String
    
    CallName = "'" & ThisWorkbook.Name & "'!" & "mdRibbonsControls." & xControl.ID
    
    Application.Run CallName
End Sub

Private Sub btLancamentos(): frmConsultas.Show: End Sub

Private Sub btreportfichas()
    On Error GoTo err:
    wsReportConsultas.Activate
    wsReportConsultas.PivotTables(1).RefreshTable
    VBA.MsgBox "O relatório está atualizado.", vbInformation
    Exit Sub
err:
    MsgBox "Não foi possível atualizar o relatório." & vbNewLine & _
            err.Number & "-" & err.Description, vbCritical
End Sub

Private Sub btRelatorio()
    On Error GoTo err:
    wsReportProcedimentos.Activate
    wsReportProcedimentos.PivotTables(1).RefreshTable
    VBA.MsgBox "O relatório está atualizado.", vbInformation
    Exit Sub
err:
    MsgBox "Não foi possível atualizar o relatório." & vbNewLine & _
            err.Number & "-" & err.Description, vbCritical
End Sub

Private Sub btprint(): frmExportReport.Show False: End Sub
Private Sub btProcedimentos(): frmProcedimentos.Show: End Sub
Private Sub btProfissional(): frmCadastroProfissional.Show: End Sub
Private Sub btProcedimento(): frmCadastroProcedimento.Show: End Sub
Private Sub btConsulta(): frmCadastroConsulta.Show: End Sub
Private Sub btCadastroView(): wsCadastros.Activate: End Sub

