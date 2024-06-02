Attribute VB_Name = "frmPrincipal"
Attribute VB_Base = "0{607C0BBC-B745-4582-A492-EB92116F2AA0}{0B4D4057-C1A5-4F2D-92B4-6193FD03EC3B}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

'Inicializando o form

Private Sub UserForm_Initialize()
    On Error Resume Next
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set ws = sMain
    Set lo = ws.ListObjects(1)
    
    With Me
        .cboStatus.List = Array("PENDENTE", "CONCLUÍDA", "CANCELADA")
        With .lstDados
            .ColumnCount = lo.ListColumns.Count
            .ColumnHeads = True
            .RowSource = lo.DataBodyRange.Address(, , , True)
        End With
    End With
End Sub

'Rotinas de atribuição a botões

Private Sub btnNewTask_Click()
    Dim oTask As New TaskManager
    
    With oTask
        .aDataInicial = Me.txtDataInicial.Value
        .bDataLimite = Me.txtDataFinal.Value
        .cObjetivo = Me.txtDescription
        .dValor = Me.txtValor.Value * 1
        .eQuantidade = (DateValue(Me.txtDataFinal.Value) * 1) - _
                       (DateValue(Me.txtDataInicial.Value) * 1)
        .fstatus = VBA.IIf(Me.cboStatus.Value = "", "PENDENTE", Me.cboStatus.Value)
        .gDataConclusao = Me.txtDataConclusao.Value
         Me.lstDados.RowSource = ""
        .Save Modo:=NovaTarefa
    End With
    
    Me.lstDados.RowSource = sMain.ListObjects(1).DataBodyRange.Address(, , , True)
    ClearControls (Me)
    
    MsgBox "Nova tarefa adicionada com sucessso", vbInformation, Me.Caption
End Sub

Private Sub btnTaskLoad_Click()
    Dim nLin As Integer
    
    nLin = Me.lstDados.ListIndex
    
    If nLin < 0 Then
        MsgBox "Você precisa primeiro selecionar um item abaixo para atualizá-lo.", _
            vbExclamation, Me.Caption
        Exit Sub
    End If
    
    With Me
        .txtDataInicial.Value = Format(.lstDados.List(nLin, 0), "dd/mm/yyyy")
        .txtDataFinal.Value = Format(.lstDados.List(nLin, 1), "dd/mm/yyyy")
        .txtDescription.Value = .lstDados.List(nLin, 2)
        .txtValor.Value = .lstDados.List(nLin, 3)
        .txtQTD.Value = .lstDados.List(nLin, 4)
        .cboStatus.Value = .lstDados.List(nLin, 5)
        .txtDataConclusao.Value = .lstDados.List(nLin, 6)
    End With
End Sub

Private Sub btnTaskUpdate_Click()
    Dim oTask As New TaskManager
        
     If ValidateEmptyFields(Me) Then
        MsgBox "Selecione um item na lista abaixo e clique em [ATUALIZAR], altere os dados e clique em [ALTERAR]" _
                , vbExclamation, Me.Caption
     Exit Sub
     End If
        
    With oTask
        .aDataInicial = Me.txtDataInicial.Value
        .bDataLimite = Me.txtDataFinal.Value
        .cObjetivo = Me.txtDescription
        .dValor = Me.txtValor.Value * 1
        .eQuantidade = (DateValue(Me.txtDataFinal.Value) * 1) - _
                       (DateValue(Me.txtDataInicial.Value) * 1)
        .fstatus = VBA.IIf(Me.cboStatus.Value = "", "PENDENTE", Me.cboStatus.Value)
        .gDataConclusao = Me.txtDataConclusao.Value
        .Save Modo:=AlterarTarefa
    End With
    
    Call ClearControls(Me)
    
    MsgBox "Dados alterados com sucesso", vbInformation, Me.Caption
End Sub

Private Sub btnDelete_Click()
    Dim nLin As Integer
    
    nLin = Me.lstDados.ListIndex
    
    If nLin < 0 Then
        MsgBox "Você precisa primeiro selecionar um item abaixo para Excluí-lo.", _
            vbExclamation, Me.Caption
        Exit Sub
    Else
        If MsgBox("Você tem certeza que deseja [EXCLUIR] o item selecionado?", vbQuestion + vbYesNo) _
            = vbYes Then sMain.ListObjects(1).ListRows(nLin + 1).Delete
    End If

    MsgBox "Excluído com sucesso", vbInformation, Me.Caption
End Sub

Private Sub btnExit_Click()
    Unload Me
End Sub

'Rotinas para comportamento de controles

Private Sub txtDataConclusao_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim nCaracters As Integer
    
    With Me
        nCaracters = VBA.Len(.txtDataConclusao.Value)
        Select Case KeyAscii
            Case Asc("0") To Asc("9")
                If nCaracters = 2 Or nCaracters = 5 Then .txtDataConclusao.SelText = "/"
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

Private Sub txtDataFinal_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim nCaracters As Integer
    
    With Me
        nCaracters = VBA.Len(.txtDataFinal.Value)
        Select Case KeyAscii
            Case Asc("0") To Asc("9")
                If nCaracters = 2 Or nCaracters = 5 Then .txtDataFinal.SelText = "/"
            Case Else
                KeyAscii = 0
        End Select
    End With
End Sub

Private Sub txtDataInicial_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Dim nCaracters As Integer
    Const MaxCaracters As Integer = 10
    
    nCaracters = Len(Me.txtDataInicial.Value)
    
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            If nCaracters = 2 Or _
                nCaracters = 5 Then
                Me.txtDataInicial.SelText = "/"
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub
