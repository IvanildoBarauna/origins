Attribute VB_Name = "frmSimulador"
Attribute VB_Base = "0{61956AE3-D63F-4C55-98F5-76CE7F4988B9}{566AA1A7-1100-4430-B354-45E016F72911}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub btnSair_Click()
    Unload Me
End Sub

Private Sub btnUpdate_Click()
    Dim ws As Excel.Worksheet
    Dim nParcelas As Integer
    
    Set ws = sTabelas
    nParcelas = Me.txtValorSJuros.Value / Me.txtParcelaSJuros.Value
    
    With Me
        .txtJurosMes.Value = ws.Range("tb" & .cboInstType.Value & "s[Instituição]").Find(.cboInst.Value).Offset(0, 1).Value2
        .txtJurosAno.Value = ws.Range("tb" & .cboInstType.Value & "s[Instituição]").Find(.cboInst.Value).Offset(0, 2).Value2
        .txtValorCJuros.Value = FormatCurrency(.txtValorSJuros.Value * _
                (1 + (.txtJurosMes * nParcelas)), 2)
        .txtParcelaCJuros.Value = FormatCurrency(.txtValorCJuros.Value / nParcelas, 2)
        .txtJurosMes.Value = FormatPercent(.txtJurosMes.Value, 2)
        .txtJurosAno.Value = FormatPercent(.txtJurosAno.Value, 2)
    End With
End Sub

Private Sub btnVoltar_Click()
    Unload Me
    frmFinanciamento.Show
End Sub

Private Sub cboInstType_Change()
    Me.cboInst.Value = vbNullString
End Sub

Private Sub cboInstType_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    Me.cboInst.RowSource = ThisWorkbook.Names(Me.cboInstType.Value).Value
End Sub

Private Sub UserForm_Initialize()
    Dim lo As Excel.ListObject
    
    Set lo = sTabelas.ListObjects(1)
    
    With Me.cboInst
        .ColumnCount = lo.ListColumns.Count
        .ColumnHeads = True
        .RowSource = lo.DataBodyRange.Address(external:=1)
    End With
End Sub
