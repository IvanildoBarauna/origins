Attribute VB_Name = "barraprogres"
Sub BarraDeProgresso()
Dim i               As Long
Dim iUltimaLinha    As Long
Dim iPercentualConcluido As Double
    
    Application.ScreenUpdating = False
    
    iUltimaLinha = ActiveSheet.Range("A1").End(xlDown).Row
    
    frmBarraDeProgresso.Show False
    
    For i = 2 To iUltimaLinha
        iPercentualConcluido = i / iUltimaLinha
        With frmBarraDeProgresso
            .framePb.Caption = Format(iPercentualConcluido, "0%") & " Concluído"
            .progressBar.Width = iPercentualConcluido * (.framePb.Width - 10)
        End With
        
        DoEvents    'Permite que sejam visualizadas as mudanças nos controles do formulário
        
        ' O código da sua macro vai aqui
        Call MinhaMacro(ActiveSheet.Cells(i, 1))
    Next
    
    Unload frmBarraDeProgresso
    
    'Atualizar a dinamica da Macro
    
    Application.Calculation = xlCalculationAutomatic
    
    
    'MsgBox "Cálculos Atualizados.", vbInformation, "Orçamento 2016"

End Sub

Private Sub MinhaMacro(ByVal rCell As Range)

'Calculate
ActiveWorkbook.RefreshAll


End Sub

