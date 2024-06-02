Attribute VB_Name = "frmProcedimentos"
Attribute VB_Base = "0{1AD0079B-229C-4E30-8189-19503CCD5CE4}{B8D58259-5E65-4471-8D51-726D759C76EA}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnAlterar_Click()
    Dim SelectedItem As Long:
    Dim IDItem       As Long
    
    With Me
        SelectedItem = .lstProcedimentos.ListIndex
        If SelectedItem < 0 Or ValidateEmptyControls(Me) Then Exit Sub
        
        If MsgBox("Você tem certeza que deseja [ALTERAR] este registro?", vbQuestion + vbYesNo) = vbYes Then
            IDItem = .lstProcedimentos.List(SelectedItem, 0)
            .lstProcedimentos.RowSource = ""
            Call ChangeProcedimento(IDItem)
            Call ClearFields(Me)
            Call PopulaListBox
            .lstProcedimentos.Selected(SelectedItem) = True
            .cboProfissional.SetFocus
            .lbValida.Caption = "REGISTRO ALTERADO"
        End If
    End With
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    Call ClearFields(Me)
End Sub

Private Sub btnExcluir_Click()
    Dim SelectedItem As Long: SelectedItem = Me.lstProcedimentos.ListIndex
    Dim IDItem       As Long
    
    If SelectedItem < 0 Then Exit Sub
    IDItem = Me.lstProcedimentos.List(SelectedItem, 0)
    
    If MsgBox("Você tem certeza que deseja [EXCLUIR] este lançamento?", vbQuestion + vbYesNo) = vbYes Then
        Me.lstProcedimentos.RowSource = ""
        wsProcedimentos.ListObjects("tbProcedimentos").ListRows(IDItem).Delete
        Call PopulaListBox
    End If
End Sub

Private Sub btnLancamento_Click()
    If ValidateEmptyControls(Me) Then Exit Sub
    Me.lstProcedimentos.RowSource = ""
    Call InsertOrSumProcedimento
    Call ClearFields(Me)
    Call PopulaListBox
    Me.cboProfissional.SetFocus
    Me.lbValida.Caption = "FICHA LANÇADA"
End Sub

Private Sub btnSelect_Click()
    On Error Resume Next
    Me.lstProcedimentos.Selected(Me.lstProcedimentos.ListCount - 1) = True
    On Error GoTo 0
End Sub


Private Sub chk_otherdate_Click()
    If Me.chk_otherdate.Value Then
        Me.txt_databpa.Enabled = True
        Me.txt_databpa.SetFocus
    Else
        Me.txt_databpa.Enabled = False
    End If
End Sub

Private Sub lstProcedimentos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim item As Long
    
    item = Me.lstProcedimentos.ListIndex
    
    If item < 0 Then Exit Sub
    
    With Me
        .cboProfissional.Value = .lstProcedimentos.List(item, 1)
        .cboProcedimento.Value = .lstProcedimentos.List(item, 2)
        .txtQuantidade.Value = .lstProcedimentos.List(item, 3)
        .txt_databpa.Value = .lstProcedimentos.List(item, 4)
    End With

End Sub

Private Sub UserForm_Initialize()
    Dim ws      As Excel.Worksheet: Set ws = wsCadastros
    Dim loProc  As Excel.ListObject
    Dim loProf  As Excel.ListObject
    Dim rngProc As Excel.Range
    Dim rngProf As Excel.Range
    
    wsView.Activate
    
    With ws
        Set loProc = .ListObjects("tbCadastroProcedimento")
        Set loProf = .ListObjects("tbCadastroProfissional")
        Set rngProc = loProc.Range
    End With
    
    With Me
        .txt_databpa.Value = StartDateBPA()
        .txt_databpa.Enabled = False
        
        With .cboProcedimento
            .ColumnCount = 1
            With loProc
                Me.cboProcedimento.RowSource = .Application.Range(.DataBodyRange(1, _
                                                .ListColumns("PROCEDIMENTO").Index), _
                                                    .DataBodyRange(.ListRows.Count, _
                                                        .ListColumns("PROCEDIMENTO").Index)).Address(external:=1)
            End With
        End With
        
        With .cboProfissional
            .ColumnCount = 1
            With loProf
                Me.cboProfissional.RowSource = .Application.Range(.DataBodyRange(1, _
                                                .ListColumns("PROFISSIONAL").Index), _
                                                    .DataBodyRange(.ListRows.Count, _
                                                        .ListColumns("PROFISSIONAL").Index)).Address(external:=1)
            End With
            .SetFocus
        End With
    End With
    
    Call PopulaListBox
End Sub

Private Sub PopulaListBox()
    Dim loLancs As ListObject: Set loLancs = wsProcedimentos.ListObjects("tbProcedimentos")
    
    With Me.lstProcedimentos
        .ColumnHeads = True
        .ColumnCount = loLancs.ListColumns.Count
        On Error Resume Next
        .RowSource = loLancs.DataBodyRange.Address(external:=1)
    End With
End Sub

Private Function isExistsData(UniqueID As String) As Long
'------------------------------------------------------
'RotineType: Function / Exit Long
'Criacao: Ivanildo Junior
'Criada em: 25/03/2018 18:45
'Objetivo: Verifica se há um item em uma determinada matriz e retorna o número do index da linha
'Aplicacaoo: isExistsData("IDUnico")
'------------------------------------------------------
    Dim mtz         As Variant
    Dim lo          As Excel.ListObject
    Dim ws          As Excel.Worksheet
    Dim iCounter    As Long
    
    Set ws = wsProcedimentos
    Set lo = ws.ListObjects(1)
    
    If lo.DataBodyRange Is Nothing Then Exit Function
    
    mtz = lo.DataBodyRange.Value
    
    For iCounter = LBound(mtz, 1) To UBound(mtz, 1)
        item = mtz(iCounter, lo.ListColumns("PROFISSIONAL").Index) & _
                mtz(iCounter, lo.ListColumns("PROCEDIMENTO").Index) & _
                mtz(iCounter, lo.ListColumns("DATA INICIAL").Index)
        If item = UniqueID Then
            isExistsData = lo.DataBodyRange(iCounter, lo.ListColumns("ID").Index).Value2
            Exit For
        End If
    Next iCounter
    
End Function

Private Sub InsertOrSumProcedimento()
    Dim lo            As ListObject
    Dim lr            As ListRow
    Dim oProcedimento As cFichaProcedimento
    Dim QTD           As Long
    Dim itemunico     As String
    Dim item          As Long
    
    Set lo = wsProcedimentos.ListObjects("tbProcedimentos")
    Set oProcedimento = New cFichaProcedimento
    itemunico = Me.cboProfissional.Value & Me.cboProcedimento.Value & Me.txt_databpa.Value
    item = isExistsData(itemunico)
    
    If item > 0 Then
        Set lr = lo.ListRows(item)
    Else
        Set lr = lo.ListRows.Add
    End If
    
    QTD = VBA.IIf(item > 0, _
                    (Me.txtQuantidade.Value * 1) + lr.Range(lo.ListColumns("QUANTIDADE").Index).Value2, _
                        Me.txtQuantidade.Value * 1)
    With oProcedimento
        .ProfissionalNome = Me.cboProfissional.Value
        .ProcedimentoNome = Me.cboProcedimento.Value
        .Quantidade = QTD
        .DataInicial = Me.txt_databpa.Value
        .SaveOrChangeData lr.Index
    End With
End Sub

Private Sub ChangeProcedimento(RowIndex As Long)
    Dim oProcedimento As cFichaProcedimento
    
    Set oProcedimento = New cFichaProcedimento
    
    With oProcedimento
        .ProfissionalNome = Me.cboProfissional.Value
        .ProcedimentoNome = Me.cboProcedimento.Value
        .Quantidade = Me.txtQuantidade.Value
        .SaveOrChangeData (RowIndex)
    End With
End Sub

