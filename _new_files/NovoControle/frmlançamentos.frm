Attribute VB_Name = "frmlançamentos"
Attribute VB_Base = "0{B407210E-B175-4245-95AD-5C340D868E27}{952DFE9E-07AC-4D5C-A25C-004AC7EFF9A0}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub UserForm_Initialize()
    Dim rngLançamentos As Range: Set rngLançamentos = RangeToComboBox(3)
    Dim rngPagamentos  As Range: Set rngPagamentos = RangeToComboBox(4)
    
    With Me
        .txtData.Value = VBA.Date
        .cbolanc.RowSource = rngLançamentos.Address(External:=1)
        .cbopgto.RowSource = rngPagamentos.Address(External:=1)
        Call PopularListBox
    End With
    
    Call ZerarTextBoxes
End Sub

Private Sub btnAlterar_Click()
    Dim item As Integer: item = Me.lstLançamentos.ListIndex
    Dim ID   As Integer: ID = Me.lstLançamentos.List(item, 0)
    Dim lo   As ListObject: Set lo = shCaixa.ListObjects("fCaixa")
    
    If item >= 1 And MsgBox("Tem certeza que deseja [ALTERAR] o lançamento de ID: " & _
                    ID & " ?", vbQuestion + vbYesNo, Me.Caption) = vbYes And Not ValidateEmptyFields(Me) Then
        Call ChangeDataOnListObject(Me, ID)
        Call PopularListBox
        Call ClearFields(Me)
        Me.cbolanc.SetFocus
        Me.btnlanc.Enabled = True
        MsgBox "Registro alterado com sucesso.", vbInformation, Me.Caption
    End If
End Sub

Private Sub btnExcluir_Click()
    Dim item As Integer: item = Me.lstLançamentos.ListIndex
    Dim ID   As Integer: ID = Me.lstLançamentos.List(item, 0)
    Dim lo   As ListObject: Set lo = shCaixa.ListObjects("fCaixa")
    
    If item > 0 Then
        If MsgBox("Tem certeza que deseja [EXCLUIR] o lançamento de ID: " & _
                    ID & " ?", vbQuestion + vbYesNo, Me.Caption) = vbYes Then
            lo.ListRows(ID).Delete
            Call PopularListBox
            Call ClearFields(Me)
            MsgBox "Registro excluído com sucesso", vbInformation, Me.Caption
        End If
    End If
End Sub

Private Sub chkCong_Click()
    With Me
        If .chkCong.Value Then
            .chkCong.Caption = "Congelado!"
        Else
            .chkCong.Caption = "Congelado?"
        End If
        ValidateCheckBox .chkCong.Value
    End With
End Sub

Private Sub ValidateCheckBox(boo As Boolean)
    Dim Gramas  As Integer
    Dim QTD     As Integer
    
    With Me
        If boo Then
            If .txtvenda.Value = 0 Then Exit Sub: .txtvenda.SetFocus
             Gramas = VBA.Replace(VBA.Right(.cbodesc.Text, 3), "G", "") * 1
             QTD = 1000 / Gramas
            .txtpreco.Value = Format(.txtvenda.Value / QTD, "Currency")
            .txtqtdperdida.Value = 0
        Else
            .txtvenda.Value = Format(0, "Currency")
            .txtpreco.Value = Format(0, "Currency")
            .txtvenda.SetFocus
        End If
    End With
End Sub

Private Sub lstLançamentos_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim item As Integer: item = Me.lstLançamentos.ListIndex
    Dim xControl As MSForms.control
    
    With Me
        If item < 1 Then
            Exit Sub
        Else
            .txtData.Enabled = True
            .txtData.Value = .lstLançamentos.List(item, 1)
            .cbolanc.Text = .lstLançamentos.List(item, 2)
            .cbopgto.Text = .lstLançamentos.List(item, 3)
            .cbodesc.Text = .lstLançamentos.List(item, 4)
            For Each xControl In Me.Controls
                On Error Resume Next
                If VBA.TypeName(xControl) = "OptionButton" Then
                     If xControl.Caption = .lstLançamentos.List(item, 5) Then
                        xControl.Value = True
                     End If
                End If
                On Error GoTo 0
            Next xControl
            .txtvenda.Value = .lstLançamentos.List(item, 6)
            .txtpreco.Value = .lstLançamentos.List(item, 7)
            .txtqtdperdida.Value = .lstLançamentos.List(item, 9)
            .btnlanc.Enabled = False
        End If
    End With
End Sub

Private Sub btn_cancel_Click()
    Unload Me
End Sub

Private Sub btnCalc_Click()
    Shell "CALC.EXE"
End Sub

Public Sub btnlanc_Click()
    If ValidateEmptyFields(Me) Then
        MsgBox "Todos os campos são obrigatórios", vbExclamation
        Exit Sub
    Else
        Call SaveOnListObject(Me)
        Call PopularListBox
        Call ClearFields(Me)
        Me.cbolanc.SetFocus
        MsgBox "Dados lançados com sucesso!" & vbNewLine & vbNewLine & _
                CostValues, vbInformation, Me.Caption
    End If
End Sub

Private Sub cbolanc_Change()
    Me.cbodesc.Value = vbNullString
    AllTextBoxStatus True, xlNo
End Sub

Private Sub cbolanc_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    On Error Resume Next
    With Me
        .cbodesc.RowSource = RangeToComboBox(.cbolanc.Value).Address(External:=1)
        If .cbolanc.Value <> 1 Then
            Call AllTextBoxStatus(False, xlYes)
        End If
    End With
End Sub

Private Sub chk_data_Click()
    With Me
        If .chk_data.Value Then
            .txtData.Value = VBA.Constants.vbNullString
            .txtData.Enabled = True
            .txtData.SetFocus
        Else
            .txtData.Enabled = False
            .txtData.Value = Date
        End If
    End With
End Sub

Private Sub txtcusto_Enter()
    If Me.txtcusto.Value = VBA.Constants.vbNullString Then Me.txtcusto.Value = Format(Me.txtcusto.Value, "Currency")
End Sub

Private Sub txtcusto_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtcusto.Value = VBA.Format(Me.txtcusto.Value, "Currency")
End Sub

Private Sub txtData_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.txtData.MaxLength = 10
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            Select Case VBA.Len(Me.txtData)
                Case 2, 5
                    Me.txtData.SelText = "/"
            End Select
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtpreco_Enter()
    If Me.txtpreco.Value = vbNullString Then Me.txtpreco.Value = Format(Me.txtpreco.Value, "Currency")
End Sub

Private Sub txtpreco_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtpreco.Value = VBA.Format(Me.txtpreco.Value, "Currency")
End Sub

Private Sub txtvenda_Enter()
    If Me.txtvenda.Value = VBA.Constants.vbNullString Then Me.txtvenda.Value = Format(Me.txtvenda.Value, "Currency")
End Sub

Private Sub txtvenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    Me.txtvenda.Value = Format(txtvenda.Value, "Currency")
End Sub

Public Sub ZerarTextBoxes()
    Dim xCtrl As MSForms.control
    
    For Each xCtrl In Me.Controls
        If TypeName(xCtrl) = "TextBox" And xCtrl.Name <> "txtData" Then
            xCtrl.Value = VBA.Format(0, "Currency")
        End If
    Next xCtrl
    
End Sub

Public Sub AllTextBoxStatus(ByVal booVal As Boolean, ByVal WithExceptions As XlYesNoGuess)
    Dim xCtrl As MSForms.control
    
    If WithExceptions = xlYes Then
        For Each xCtrl In Me.Controls
            If TypeName(xCtrl) = "TextBox" Then
                If xCtrl.Name = "txtData" Or xCtrl.Name = "txtvenda" Then GoTo NextCtrl
                    xCtrl.Enabled = booVal
            End If
NextCtrl:
        Next xCtrl
    Else
        For Each xCtrl In Me.Controls
            If TypeName(xCtrl) = "TextBox" And xCtrl.Name <> "txtData" Then
                xCtrl.Enabled = booVal
            End If
        Next xCtrl
    End If
End Sub

Private Function WidthsToListBox(arr) As String
    Const Mult As Double = 8.5
    Dim iCounter As Integer
    Dim nLEN     As Integer
    Dim tmpArr() As String

    ReDim tmpArr(0 To UBound(arr, 2))
    
    For iCounter = 0 To UBound(arr, 2)
        If iCounter = 0 Then nLEN = 0 Else nLEN = VBA.Len(arr(0, iCounter)) * Mult
        tmpArr(iCounter) = VBA.CStr(nLEN)
    Next iCounter
        
    WidthsToListBox = VBA.Join(tmpArr, ";")
End Function

Private Sub PopularListBox()
    Dim ws             As Worksheet:    Set ws = shCaixa
    Dim lo             As ListObject:   Set lo = ws.ListObjects("fCaixa")
    Dim rngDados       As Range
    Dim arr
    
    With lo: Set rngDados = .Application.Range(lo.Range(1, 1), _
                            .DataBodyRange(.ListRows.Count, lo.ListColumns("QTD Perdida").index)): End With
    With Me.lstLançamentos
        .ColumnCount = rngDados.Columns.Count
        .List = FormatColumnsInArray(FormatColumnsInArray(FilterArrayWithDate(rngDados.Value2, 2), "Currency", 7, 8, 9), "DD/MM/YYYY", 2)
        .ColumnWidths = WidthsToListBox(.List)
        With Me
            .lbControle.Caption = VBA.IIf(.lbControle.Caption = "Total de Registros:", _
                                            .lbControle.Caption & " " & .lstLançamentos.ListCount - 1, _
                                                "Total de Registros: " & .lstLançamentos.ListCount - 1)
        End With
    End With
End Sub


