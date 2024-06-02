Attribute VB_Name = "frmFinanciamento"
Attribute VB_Base = "0{005875BB-42A2-43BA-889D-076ECB23CF1F}{F97A3088-8E4A-4A36-913B-5BEBA4134426}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btClear_Click()
    Call ClearFields(Me)
End Sub

Private Sub btn_Exit_Click()
    Unload Me
End Sub

Private Sub btnSpin_SpinUp()
    If Me.txtParcelas.Value = "" Then Me.txtParcelas.Value = 1
    If Me.txtParcelas.Value > 59 Then Exit Sub
    Me.txtParcelas.Value = Me.txtParcelas.Value + 1
End Sub
Private Sub btnSpin_SpinDown()
    If Me.txtParcelas.Value = "" Then Me.txtParcelas.Value = 1
    If Me.txtParcelas.Value < 2 Then Exit Sub
    Me.txtParcelas.Value = Me.txtParcelas.Value - 1
End Sub

Private Sub txtEntrada_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtEntrada.Value = Format(Me.txtEntrada.Value, "R$ #,##0.00")
End Sub

Private Sub txtPreco_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtPreco.Value = Format(Me.txtPreco.Value, "R$ #,##0.00")
End Sub

Private Sub UserForm_Initialize()
    Dim lo As Excel.ListObject
    
    Set lo = sTabelas.ListObjects(1)
    
    Call GetData(Me)
    
    With Me.cboInstuicao
        .ColumnCount = lo.ListColumns.Count
        .ColumnHeads = True
        .RowSource = lo.DataBodyRange.Address(external:=1)
    End With
End Sub

Private Sub btnSimular_Click()
    If ValidateEmptyControls(Me) Then
        MsgBox "Preencha todos os campos para iniciar a simulação", _
            vbExclamation, Me.Caption
    Else
        Call SaveData(Me)
        Call btnSimulateRotine
    End If
End Sub
