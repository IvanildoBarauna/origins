Attribute VB_Name = "frmLancamentos"
Attribute VB_Base = "0{D8BFC169-09C3-4B0B-8874-E1AD520C5287}{D2A01734-5CE7-4F46-B71E-70C8B65554C8}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub UserForm_Initialize()
    Dim ws As Worksheet
    Dim lo As ListObject
    
    Set ws = wsDados
    Set lo = ws.ListObjects(1)
    
    With Me
        .cboProduto.List = Array("RECARGA", "REVISTA", "JORNAL", "DIVERSOS")
        .lstDados.ColumnCount = lo.ListColumns.Count - 4
        If lo.ListRows.Count = 0 Then lo.ListRows.Add
        With .lstDados
            .List = Filtermtz(lo.DataBodyRange.Value, lo.ListColumns("DATA").index)
        End With
    End With
End Sub

Private Sub btnGoToEnd_Click()
    Me.lstDados.Selected(Me.lstDados.ListCount - 1) = True
End Sub

Private Sub btnLancar_Click()
    On Error GoTo Error
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    Dim tmpArr(1 To 1, 1 To 4) As Variant
    
    Set ws = wsDados
    Set lo = ws.ListObjects(1)
    
    If EmptyFields(Me) Then Exit Sub
    
    If Not VBA.IsNumeric(Me.txtValorVenda.Value) Then
        MsgBox "Valor de venda inválido", vbCritical, Me.Caption
        Exit Sub
    End If
        
    tmpArr(1, 1) = Application.WorksheetFunction.Max(lo.ListColumns("ID").DataBodyRange.Value2) + 1
    tmpArr(1, 2) = VBA.Now
    tmpArr(1, 3) = Me.cboProduto.Value
    tmpArr(1, 4) = Me.txtValorVenda.Value * 1
    
    
    If lo.ListRows.Count = 1 And lo.DataBodyRange(1, 1).Value2 = "" Then
        Set lr = lo.ListRows(1)
    Else
        Set lr = lo.ListRows.Add
    End If
    
    With lr
        .Application.Range(lo.DataBodyRange(lr.index, 1), _
            lo.DataBodyRange(lr.index, 4)).Value = tmpArr
    End With
    
    With Me
        .lstDados.List = Filtermtz(lo.DataBodyRange.Value, _
            lo.ListColumns("DATA").index)
        .lstDados.Selected(.lstDados.ListCount - 1) = True
        
        MsgBox "Venda de: " & .cboProduto.Value & vbNewLine & vbNewLine & _
            "Valor: R$ " & .txtValorVenda.Value & vbNewLine & vbNewLine & _
                "Lançada com sucesso.", vbInformation, .Caption
        
        .lstDados.List = Filtermtz(lo.DataBodyRange.Value, _
            lo.ListColumns("DATA").index)
        ClearFields Me
        .cboProduto.SetFocus
        Erase tmpArr
        Exit Sub
Error:
        ErrRaise ("Não foi possível realizar o lançanto.")
    End With
End Sub

Private Sub btnChange_Click()
    On Error GoTo Err
    Dim ws          As Worksheet
    Dim lo          As ListObject
    Dim SearchID    As Long
    Dim FoundRow    As Long
    Dim item        As Long
    
    Set ws = wsDados
    Set lo = ws.ListObjects(1)
    item = Me.lstDados.ListIndex
    
    If item < 0 Then
        GoTo ExitPoint
    ElseIf Me.cboProduto.Value = "" Or _
           Me.txtValorVenda.Value = "" Or _
           Me.btnLancar.Enabled Then
ExitPoint:
        MsgBox "Dê um duplo clique no item e insira os dados que deseja [ALTERAR].", vbExclamation, Me.Caption
        Exit Sub
    Else
        SearchID = Me.lstDados.List(item, 0)
        FoundRow = lo.ListColumns("ID").Range.Find(What:=SearchID, After:=lo.Range(1, 1), LookIn:=xlValues, LookAt _
            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, SearchFormat:=False).Row
        
        With ws
            .Cells(FoundRow, lo.ListColumns("PRODUTO").index + 1).Value = Me.cboProduto.Value
            .Cells(FoundRow, lo.ListColumns("VALOR").index + 1).Value = Me.txtValorVenda.Value * 1
        End With
        
        ClearFields Me
        Me.btnLancar.Enabled = True
        Me.lstDados.List = Filtermtz(lo.DataBodyRange.Value, _
            lo.ListColumns("DATA").index)
        Me.cboProduto.SetFocus
        MsgBox "Dados alterados com sucesso.", vbInformation, Me.Caption
        Exit Sub
    End If
Err:
    ErrRaise ("Não foi possível realizar a alteração.")
End Sub

Private Sub btnDelete_Click()
    On Error GoTo Err
    Dim ws          As Worksheet
    Dim lo          As ListObject
    Dim SearchID    As Long
    Dim FoundRow    As Long
    Dim item        As Long
    
    Set ws = wsDados
    Set lo = ws.ListObjects(1)
    item = Me.lstDados.ListIndex
    
    If item < 0 Then
        MsgBox "Dê um duplo clique no item e insira os dados que deseja [EXCLUIR].", vbExclamation, Me.Caption
        Exit Sub
    ElseIf item = 0 Then Exit Sub
    Else
        SearchID = Me.lstDados.List(item, 0)
        FoundRow = lo.ListColumns("ID").Range.Find(What:=SearchID, After:=lo.Range(1, 1), LookIn:=xlValues, LookAt _
            :=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, SearchFormat:=False).Row
        
        ws.Rows(FoundRow).EntireRow.Delete
        ClearFields Me
        MsgBox "Registro excluído com sucesso.", vbInformation, Me.Caption
        If Me.lstDados.ListCount < 3 Then
            Unload Me
            frmLancamentos.Show
        Else
            Me.lstDados.List = Filtermtz(lo.DataBodyRange.Value, _
                lo.ListColumns("DATA").index)
        End If
        Exit Sub
    End If
        
Err:
    ErrRaise ("Não foi possível excluir o registro.")
End Sub

Private Sub btnSair_Click()
    Unload Me
End Sub

Private Sub lstDados_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim item  As Long
    
    item = Me.lstDados.ListIndex
    
    If item < 1 Then Exit Sub
    
    With Me
        .cboProduto.Value = Me.lstDados.List(item, 2)
        .txtValorVenda.Value = Me.lstDados.List(item, 3)
        .cboProduto.SetFocus
        .btnLancar.Enabled = False
    End With
End Sub

Private Sub txtValorVenda_Enter()
    If Me.txtValorVenda.Value = "R$ 0.00" Then Me.txtValorVenda.Value = ""
End Sub

Private Sub txtValorVenda_Exit(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.txtValorVenda.Value = "" Then Me.txtValorVenda.Value = 0
    Me.txtValorVenda.Value = VBA.Format(Me.txtValorVenda.Value, "Currency")
End Sub

