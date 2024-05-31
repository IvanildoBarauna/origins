Attribute VB_Name = "frmConsultas"
Attribute VB_Base = "0{B28B9C6D-875E-42DD-8CA2-BF118A1C2E39}{21E54F65-E931-4892-93A5-7FCAD2CC9092}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub btnAlterar_Click()
    Dim lo      As ListObject: Set lo = wsConsultas.ListObjects("tbConsultas")
    Dim lr      As ListRow
    Dim SelectedItem    As Long
    Dim IDItem          As Long
    
    SelectedItem = Me.lstConsultas.ListIndex
    If SelectedItem < 0 Or ValidateEmptyControls(Me) Then Exit Sub
    
    If MsgBox("Você tem certeza que deseja [ALTERAR] este registro?", vbQuestion + vbYesNo) = vbYes Then
        With Me
            IDItem = .lstConsultas.List(SelectedItem, 0)
            Call InsertOrChangeConsulta(IDItem)
            Call ClearFields(Me)
            .lstConsultas.Selected(SelectedItem) = True
            .lbValida.Caption = "REGISTRO ALTERADO!"
        End With
    End If
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub

Private Sub btnClear_Click()
    Call ClearFields(Me)
End Sub

Private Sub btnExcluir_Click()
    Dim SelectedItem As Long: SelectedItem = Me.lstConsultas.ListIndex
    Dim IDItem       As Long
    
    If SelectedItem < 0 Then Exit Sub
    IDItem = Me.lstConsultas.List(SelectedItem, 0)
    
    If MsgBox("Você tem certeza que deseja [EXCLUIR] este lançamento?", vbQuestion + vbYesNo) = vbYes Then
        Me.lstConsultas.RowSource = ""
        wsConsultas.ListObjects("tbConsultas").ListRows(IDItem).Delete
        Call PopulaListBox
        Call ClearFields(Me)
    End If
End Sub

Private Sub btnSelect_Click()
    On Error Resume Next
    Me.lstConsultas.Selected(lstConsultas.ListCount - 1) = True
    On Error GoTo 0
End Sub

Private Sub lstConsultas_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim item As Long
    
    item = Me.lstConsultas.ListIndex
    
    If item < 0 Then Exit Sub
    
    With Me
        .cbo_prof.Value = .lstConsultas.List(item, 1)
        .txt_nascto.Value = VBA.Format(.lstConsultas.List(item, 2), "00\/00\/0000")
        .txt_databpa.Value = VBA.Format(.lstConsultas.List(item, 3), "00\/00\/0000")
    End With
End Sub

Private Sub UserForm_Initialize()
    Dim ws      As Excel.Worksheet
    Dim lo      As Excel.ListObject
    
    Set ws = wsCadastros
    Set lo = ws.ListObjects("tbCadastroConsultas")
    wsView.Activate
    With Me
        .cbo_prof.RowSource = lo.Application.Range(lo.DataBodyRange(1, 2), _
                                lo.Range(lo.ListRows.Count, 2)).Address(external:=1)
        Call PopulaListBox
        .btnLan.Enabled = False
        .txt_databpa = StartDateBPA()
        .txt_databpa.Enabled = False
        .cbo_prof.SetFocus
    End With
End Sub

Private Sub chk_otherdate_Click()
    If Me.chk_otherdate.Value Then
        Me.txt_databpa.Enabled = True
        Me.txt_databpa.SetFocus
    Else
        Me.txt_databpa.Enabled = False
    End If
End Sub

Private Sub cbo_prof_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With Me
        If .cbo_prof.Value = "" Then
            .lbValida.Caption = "#ERRO PROF = VAZIO"
            .lbValida.ForeColor = &HFF&
            .btnLan.Enabled = False
        Else
            .lbValida = ""
        End If
    End With
End Sub

Private Sub txt_nascto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With Me
        If .txt_nascto = "" Then
            .lbValida = "#ERRO DATA NASCTO = VAZIO"
            Cancel = True
            .btnLan.Enabled = False
        ElseIf Len(Replace(.txt_nascto, "/", "")) <> 8 Then
err:
            .lbValida = "#ERRO DATA NASCTO = INVALIDA."
            Cancel = True
        Else
            .lbValida = ""
            .btnLan.Enabled = True
            .txt_nascto.Value = VBA.Format(VBA.Replace(Me.txt_nascto.Value, "/", ""), "00\/00\/0000")
            If Not VBA.IsDate(.txt_nascto.Value) Then GoTo err
        End If
    End With
End Sub

Private Sub txt_databpa_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    With Me
        If .txt_databpa = "" Then
            .lbValida = "#ERRO DATA BPA = VAZIO"
            Cancel = True
            .btnLan.Enabled = False
        ElseIf Len(Replace(.txt_databpa, "/", "")) <> 8 Then
            .lbValida.Caption = "#ERRO DATA BPA = INVALIDA"
            .txt_databpa.Value = ""
            Cancel = True
            .btnLan.Enabled = False
        Else
            .lbValida = ""
            .btnLan.Enabled = True
            .txt_databpa.Value = VBA.Format(VBA.Replace(.txt_databpa, "/", ""), "00\/00\/0000")
        End If
    End With
End Sub

Private Sub btnLan_Click()
    With Me
        .lstConsultas.RowSource = ""
        Call InsertOrChangeConsulta
        .lbValida.Caption = "REGISTRO LANÇADO!"
        Call PopulaListBox
        .lstConsultas.Selected(.lstConsultas.ListCount - 1) = True
        Call ClearFields(Me)
        .cbo_prof.SetFocus
        .txt_databpa.Enabled = False
        .chk_otherdate.Value = False
    End With
End Sub

Private Sub InsertOrChangeConsulta(Optional RowIndex As Long = 0)
    Dim oConsulta As cFichaConsulta
    
    Set oConsulta = New cFichaConsulta
    
    With oConsulta
        .Profissional = Me.cbo_prof.Value
        .DataNascimento = Me.txt_nascto.Value
        .DataInicial = Me.txt_databpa.Value
        .SaveOrChangeData RowIndex
    End With
End Sub

Private Sub PopulaListBox()
    Dim lo As ListObject: Set lo = wsConsultas.ListObjects("tbConsultas")
    
    With Me.lstConsultas
        .ColumnHeads = True
        .ColumnCount = lo.ListColumns.Count
        On Error Resume Next
        .RowSource = lo.DataBodyRange.Address(external:=1)
    End With
    
End Sub

Public Function ValidateEmptyControls(ByRef FRM As UserForm) As Boolean
    Dim xControl As MSForms.control
    Dim sList    As String
    
    For Each xControl In FRM.Controls
        Select Case TypeName(xControl)
            Case "TextBox", "ComboBox"
                If xControl.Value = vbNullString Then
                    If Not ValidateEmptyControls Then _
                        ValidateEmptyControls = True
                    sList = sList & vbNewLine & xControl.Tag
                End If
        End Select
    Next xControl
    
    If ValidateEmptyControls Then MsgBox "Preencha os campos abaixo:" _
        & vbNewLine & sList
End Function

