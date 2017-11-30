Attribute VB_Name = "frmCadastroProcedimento"
Attribute VB_Base = "0{DB8991AD-40C2-45FD-92B5-3E689E5B6539}{7CF32D2D-65CD-4C08-986C-FD184A165C4A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btnAlterar_Click()
    Dim oProced As cNewProcedimento
    Dim item    As Integer
    Dim IDItem As String
    
    item = Me.lstProcedimentos.ListIndex: If item < 0 Then Exit Sub
    
    If MsgBox("Você tem certeza que deseja [ALTERAR] o registro de ID: " & IDItem, vbQuestion + vbYesNo) = vbYes Then
        Set oProced = New cNewProcedimento
        IDItem = Me.lstProcedimentos.List(item, 0)
        Me.lstProcedimentos.RowSource = ""
        
        With oProced
            .NomeProcedimento = Me.txtProcedimento.Value
            .CodigoProcedimento = Me.txtCodProcedimento.Value
            .SaveOrChangeData (IDItem)
            Call ClearFields(Me)
            Me.txtProcedimento.SetFocus
            Me.btnExcluir.Enabled = True
            Me.btnLancamento.Enabled = True
            Call PopulaListBox
            MsgBox "Registro alterado com sucesso.", vbInformation, Me.Caption
        End With
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    ClearFields Me
    Me.txtProcedimento.SetFocus
End Sub

Private Sub btnExcluir_Click()
    Dim item As Integer
    Dim ID   As String
    
    item = Me.lstProcedimentos.ListIndex: If item < 0 Then Exit Sub
    
    ID = Me.lstProcedimentos.List(item, 0)
    
    If MsgBox("Você tem certeza que deseja [EXCLUIR] o item de ID: " & ID, vbQuestion + vbYesNo) = vbYes Then
        Me.lstProcedimentos.RowSource = ""
        wsCadastros.ListObjects(Me.Tag).ListRows(ID).Delete
        Call PopulaListBox
        Me.btnLancamento.Enabled = True
        Me.btnExcluir.Enabled = True
        MsgBox "Registro excluído com sucesso.", vbInformation, Me.Caption
    End If
End Sub

Private Sub btnLancamento_Click()
    Dim oProcedimento As cNewProcedimento
    
    Set oProcedimento = New cNewProcedimento
        
    If Not ValidateEmptyControls(Me) Then
        Me.lstProcedimentos.RowSource = ""
        With oProcedimento
            .NomeProcedimento = Me.txtProcedimento.Value
            .CodigoProcedimento = Me.txtCodProcedimento.Value
            .SaveOrChangeData
        End With
        
        Call PopulaListBox
        Call ClearFields(Me)
        
        MsgBox "Registro efetuado com sucesso.", vbInformation, Me.Caption
    End If
End Sub

Private Sub lstProcedimentos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim item As Integer: item = Me.lstProcedimentos.ListIndex
    
    With Me
        .btnLancamento.Enabled = False
        .btnExcluir.Enabled = False
        .txtProcedimento.Value = .lstProcedimentos.List(item, 1)
        .txtCodProcedimento.Value = .lstProcedimentos.List(item, 2)
    End With
End Sub

Private Sub UserForm_Initialize()
    wsView.Activate
    Call PopulaListBox
End Sub

Private Sub PopulaListBox()
    Dim lo As ListObject
    Dim rng As Range
    
    Set rng = wsCadastros.ListObjects(Me.Tag).DataBodyRange
    
    With Me
        .lstProcedimentos.ColumnHeads = True
        .lstProcedimentos.ColumnCount = rng.Columns.Count
        .lstProcedimentos.RowSource = rng.Address(external:=1)
        .txtProcedimento.SetFocus
    End With
End Sub
