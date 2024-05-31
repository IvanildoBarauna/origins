Attribute VB_Name = "frmRelacaoRuas"
Attribute VB_Base = "0{2001715F-0316-4FEE-ABA8-A8CA59CE4FA1}{33A4699C-5340-44D8-9CE2-47339F0D571E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub UserForm_Initialize()
    Dim lo As Excel.ListObject
    
    Set lo = wsRuasAgents.ListObjects(1)
    
    EnableControlsAntiqueCodes False
    
    With Me
        If Not wsListaAgents.ListObjects(1).DataBodyRange Is Nothing Then
            .ComboAgent.RowSource = wsListaAgents.ListObjects(1).ListColumns(3).DataBodyRange.Address(external:=1)
            .ComboAgentNewAddress.RowSource = wsListaAgents.ListObjects(1).ListColumns(3).DataBodyRange.Address(external:=1)
        End If
        .ComboLog.List = Array("RUA", "AVENIDA", "ALAMEDA", "VIELA", "ESTRADA")
        .lstDados.List = lo.Range.Value
        .lstDados.ColumnCount = lo.ListColumns.Count
        .lbUserInfo.Caption = Me.lstDados.ListCount - 1 & " registros localizados." '-1 do Cabeçalho
    End With
End Sub

Private Sub btnClearFields_Click()
    Dim iControl As MSForms.Control
       
    For Each iControl In Me.Controls
        Select Case VBA.TypeName(iControl)
            Case Is = "TextBox", "ComboBox": If iControl.Value <> "" Then iControl.Value = ""
            Case Is = "CheckBox": iControl.Value = False
        End Select
    Next iControl
    
End Sub

Private Sub btnClearSelection_Click()
    Dim iCounter As Variant
    
    For iCounter = 0 To Me.lstDados.ListCount
        If Me.lstDados.Selected(iCounter) Then Me.lstDados.Selected(iCounter) = False
    Next iCounter
End Sub

Private Sub btnFilterRemove_Click()
    With Me
        .lstDados.List = wsRuasAgents.ListObjects(1).Range.Value
        .ComboAgent.Value = ""
        .lbUserInfo.Caption = Me.lstDados.ListCount - 1 & " registros localizados."
    End With
End Sub

Private Sub btnNewLanc_Click()
    On Error GoTo err
    Dim oAgentAddress           As cAgentAddress
    Dim FinalArea               As String
    
    If Not ValidateEmptyFields(Me.fmInclusion) Then
        FinalArea = VBA.IIf(Not Me.chkCodesValidate.Value, _
                            Me.txtAreaCode.Value & "-" & Me.txtmAreaCode.Value, _
                            Me.txtAreaCode.Value & "-" & Me.txtmAreaCode.Value & " / " & Me.txtAreaCodeAnt.Value & "-" & Me.txtmAreaCodeAnt)
                            
        Set oAgentAddress = New cAgentAddress
        
        With oAgentAddress
            .aFuncional = Me.txtFunctional.Value
            .bNomeAgente = Me.ComboAgentNewAddress.Value
            .cAreaMicroArea = FinalArea
            .dEndereco = AbreviaLogradouro(Me.ComboLog.Value & " " & Me.txtRuaNome.Value)
            .eBairro = AbreviaBairros(Me.txtBairro.Value)
            .fCep = VBA.Format(Me.txtCEP.Value, "00000-000")
            .gDetalheAdicional = VBA.IIf(Me.txtAditionalDetail.Value = "", "N/A", VBA.UCase(Me.txtAditionalDetail.Value))
            .SaveOrChangeData
            Me.lstDados.List = wsRuasAgents.ListObjects(1).Range.Value
        End With
    End If
    
    MsgBox "Registro lançado com sucesso.", vbInformation, Me.Caption
    Exit Sub
err:
    MsgBox "Não foi possível lançar os dados na tabela, verifique o erro!" & vbNewLine & err.Number & "-" & err.Description, vbCritical, Me.Caption
End Sub

Private Sub btnSelectionDelete_Click()
    On Error GoTo GetErr
    Dim lo              As Excel.ListObject
    Dim iCounter        As Long
    Dim TotalSelected   As Long
    Dim sValidate       As String
    
    If Me.ComboAgent.Value <> "" Then Me.ComboAgent.Value = "": Exit Sub
    
    If isSelected(Me.lstDados, xlYes) Then
        TotalSelected = SelectedCounter(Me.lstDados, xlYes)
        sValidate = VBA.IIf(TotalSelected > 1, " registros que foram selecionados", " registro que foi selecionado")
        If MsgBox("[ATENÇÃO] Você está prestes a excluir " & TotalSelected & _
            sValidate & ", deseja continuar?", vbQuestion + vbYesNo) = vbYes Then
            Set lo = wsRuasAgents.ListObjects(1)
            For iCounter = 1 To Me.lstDados.ListCount - 1
                If Me.lstDados.Selected(iCounter) Then
                    lo.ListRows(iCounter).Range.ClearContents
                End If
            Next iCounter
            lo.DataBodyRange.SpecialCells(xlCellTypeBlanks).Rows.Delete
            Me.lstDados.List = lo.Range.Value
            MsgBox TotalSelected & " registros foram excluídos com sucesso.", vbInformation, Me.Caption
        Else
            MsgBox "Operação cancelada!", vbInformation, Me.Caption
        End If
    End If
    Exit Sub
GetErr:
    MsgBox "Não foi possível concluír a operação. " & vbNewLine & err.Description & "-" & err.Number, _
            vbCritical, Me.Caption
End Sub

Private Sub btnSelectionInvert_Click()
    Dim iCounter As Long
    
    If isSelected(Me.lstDados, xlYes) Then
        For iCounter = 1 To Me.lstDados.ListCount - 1
            Me.lstDados.Selected(iCounter) = Not Me.lstDados.Selected(iCounter)
        Next iCounter
    End If
End Sub

Private Sub chkCodesValidate_Click()
    EnableControlsAntiqueCodes Me.chkCodesValidate.Value
End Sub

Private Sub chkSelection_Click()
    Dim iCounter As Long
    
    For iCounter = 1 To Me.lstDados.ListCount - 1
        Me.lstDados.Selected(iCounter) = Me.chkSelection.Value
    Next iCounter
End Sub

Private Sub ComboAgent_Change()
    Dim lo As Excel.ListObject: Set lo = wsRuasAgents.ListObjects(1)
    
    With Me
        If Not .ComboAgent.Value = "" And Not lo.DataBodyRange Is Nothing Then
            .lstDados.List = FilterArray(lo.DataBodyRange.Value, lo.ListColumns("Nome do Agente").index, .ComboAgent.Value)
            .lbUserInfo.Caption = .lstDados.ListCount - 1 & " registros localizados." ' -1 do Cabeçalho
            .ComboAgentNewAddress.Value = .ComboAgent.Value
        Else
            .lstDados.List = lo.Range.Value
        End If
    End With
    
End Sub

Private Sub ComboAgentNewAddress_Change()
    Me.txtFunctional.Value = GetFunctional(Me.ComboAgentNewAddress.Value)
End Sub

Private Sub EnableControlsAntiqueCodes(booAux As Boolean)
    Dim iControl As MSForms.Control
    
    For Each iControl In Me.fmAntiquAreasCodes.Controls
        iControl.Enabled = booAux
    Next iControl
    
    Me.fmAntiquAreasCodes.Enabled = booAux
End Sub

Sub Main()
    With Me
        .ComboAgentNewAddress.ListIndex = 5
        .txtAreaCode.Value = "300"
        .txtmAreaCode.Value = "01"
        .ComboLog.ListIndex = 1
        .txtRuaNome.Value = "VERONA"
        .txtBairro.Value = "VILA DOS TOCA"
        .txtCEP.Value = "06813270"
        .txtAditionalDetail.Value = "OBS DOS TOCA"
    End With
End Sub

