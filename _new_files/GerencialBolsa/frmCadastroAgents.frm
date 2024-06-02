Attribute VB_Name = "frmCadastroAgents"
Attribute VB_Base = "0{6D6215B5-5BCC-4968-ADD1-8839755420D9}{9AF6977D-F42E-4924-916E-0E8464704DDC}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnCadastra_Click()
    Dim oAgentClass As cAgent
    Dim item        As Integer
        
    On Error Resume Next
    item = Me.lstDados.List(Me.lstDados.ListIndex, 0)
    On Error GoTo 0
    
    If ValidateEmptyFields(Me) Then
        MsgBox "Todos os campos são obrigatórios", vbExclamation, Me.Caption
        Exit Sub
    End If
    
    Set oAgentClass = New cAgent
    
    With oAgentClass
        .NomeAgente = Me.txtNameAgent.Value
        .Funcional = Me.txtFuncional.Value
        Me.lstDados.RowSource = ""
        .SaveORChangeReg (item)
        Me.lstDados.RowSource = wsListaAgents.ListObjects(1).DataBodyRange.Address(external:=1)
    End With
    Call btnClear_Click
    Call ClearFields(Me, "TextBox")
    MsgBox "Agente cadastrado/alterado com sucesso.", vbInformation, Me.Caption
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    Call ClearFields(Me, "TextBox", "Label")
End Sub

Private Sub btnDelete_Click()
    Dim lo      As ListObject
    Dim item    As Integer
    
    Set lo = wsListaAgents.ListObjects(1)
    item = Me.lstDados.List(Me.lstDados.ListIndex, 0)
    
    If MsgBox("Tem certeza que deseja [EXCLUIR] o agente: " & Me.lstDados.List(item - 1, 2) & "?", vbQuestion + vbYesNo) = vbYes Then
        lo.ListRows(item).Delete
        Call btnClear_Click
        MsgBox "Registro excluído com sucesso.", vbInformation
    End If
End Sub

Private Sub lstDados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim Selected As Integer
    
    With Me
        .txtNameAgent.Value = .lstDados.List(Selected, 2)
        .txtFuncional.Value = .lstDados.List(Selected, 1)
        .lbctrl.Caption = "[ATENÇÃO] Modo de Alteração"
        .btnCadastra.Caption = "[ALTERAR]"
    End With
End Sub

Private Sub txtFuncional_Enter()
    If Me.txtNameAgent.Value = "" Then
        Me.lbctrl.Caption = "[ATENÇÃO] Modo de Cadastro"
        Me.btnCadastra.Caption = "[CADASTRAR]"
        If Not Me.lstDados.ListIndex < 0 Then Me.lstDados.Selected(Me.lstDados.ListIndex) = False
    End If
End Sub

Private Sub UserForm_Initialize()
    Dim lo As ListObject
    
    Set lo = wsListaAgents.ListObjects(1)
    
    With Me.lstDados
        .ColumnCount = lo.ListColumns.Count
        .ColumnHeads = True
        If Not lo.DataBodyRange Is Nothing Then .RowSource = lo.DataBodyRange.Address(external:=1)
    End With
    
End Sub
