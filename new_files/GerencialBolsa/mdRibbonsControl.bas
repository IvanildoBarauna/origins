Attribute VB_Name = "mdRibbonsControl"
Option Explicit

Sub RibbonCallBack(xControl As IRibbonControl)
    Const ModuleName    As String = "mdRibbonsControl."
    Dim CallName        As String
    
    CallName = "'" & ThisWorkbook.Name & "'!" & ModuleName & xControl.ID
    Application.Run CallName
End Sub

Sub btnImportacao(): Call mdImportExport.FileImport: End Sub
Sub btnExportacao()
    If MsgBox("A operação a seguir pode demorar um ou mais segundos, deseja continuar?", vbQuestion + vbYesNo) = vbYes Then
        Application.StatusBar = "Aguarde ... Exportando arquivos."
        Call ExportDataAllAgents
        Application.StatusBar = "Concluído!"
    End If
End Sub

Sub btnReport()
    On Error GoTo err:
    With wsDyn
        .Activate
        .PivotTables(1).RefreshTable
        .PivotTables(2).RefreshTable
    End With
    VBA.MsgBox "O relatório está atualizado.", vbInformation
    Exit Sub
err:
    MsgBox "Não foi possível atualizar o relatório." & vbNewLine & _
            err.Number & "-" & err.Description, vbCritical
End Sub

Sub btCadastroAgent(): frmCadastroAgents.Show: End Sub
Sub btCadastroRuas(): frmRelacaoRuas.Show: End Sub
Sub btTabelaBolsa(): shBD.Activate: End Sub
Sub btRelacaoAgents(): wsListaAgents.Activate: End Sub
Sub btRelacaoRuas(): wsRuasAgents.Activate: End Sub



