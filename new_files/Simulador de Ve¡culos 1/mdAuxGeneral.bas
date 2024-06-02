Attribute VB_Name = "mdAuxGeneral"
Option Explicit
Public Sub CallForm()
    frmFinanciamento.Show
End Sub

Public Sub btnSimulateRotine()
    Dim ws                      As Excel.Worksheet
    Dim Entrada                 As Integer
    Dim Parcelas                As Integer
    Dim InstType                As String
    Dim Inst                    As String
    Dim JurosMes                As Double
    Dim JurosAno                As Double
    Dim ValorFinanciadoSJuros   As Double
    Dim ValorFinanciadoCJuros   As Double
    Dim ParcelaSJuros           As Double
    Dim ParcelaCJuros           As Double
    
    Set ws = sTabelas
    
    With frmFinanciamento
        Entrada = .txtEntrada.Value
        Parcelas = .txtParcelas.Value * 1
        Inst = .cboInstuicao.List(.cboInstuicao.ListIndex, 0)
        JurosMes = .cboInstuicao.List(.cboInstuicao.ListIndex, 1)
        JurosAno = .cboInstuicao.List(.cboInstuicao.ListIndex, 2)
        ValorFinanciadoSJuros = .txtPreco.Value - Entrada
        ValorFinanciadoCJuros = ValorFinanciadoSJuros * (1 + (JurosMes * Parcelas))
        ParcelaSJuros = ValorFinanciadoSJuros / Parcelas
        ParcelaCJuros = ValorFinanciadoCJuros / Parcelas
        Unload frmFinanciamento
    End With
    
    With frmSimulador
        .cboInst.Value = Inst
        .txtJurosMes.Value = FormatPercent(JurosMes, 2)
        .txtJurosAno.Value = FormatPercent(JurosAno, 2)
        .txtValorSJuros.Value = VBA.FormatCurrency(ValorFinanciadoSJuros, 2)
        .txtParcelaSJuros.Value = VBA.FormatCurrency(ParcelaSJuros, 2)
        .txtValorCJuros.Value = VBA.FormatCurrency(ValorFinanciadoCJuros, 2)
        .txtParcelaCJuros.Value = VBA.FormatCurrency(ParcelaCJuros, 2)
        .Show
    End With
End Sub
