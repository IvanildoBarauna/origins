Attribute VB_Name = "frmProcedimentos"
Attribute VB_Base = "0{FE8717D1-D69F-46AE-B2F6-0B3AC76C9547}{4253D45F-688E-43F4-A7A6-9508E9045756}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Private Sub btnCancel_Click(): Unload Me: End Sub
    
Private Sub btnExcluir_Click()
#If booAux Then
    Dim conn        As ADODB.Connection
    Dim rs          As ADODB.Recordset
#Else
    Dim conn        As Object
    Dim rs          As Object
#End If
    Dim itemunico As String
    Dim selected  As Long
    Dim DeleteID  As Long
    
    Set conn = DataBaseConnection()
    Set rs = myRecordSet()
    selected = Me.lstProcedimentos.ListIndex
    
    If selected < 0 Then
        Exit Sub
    ElseIf MsgBox("Tem certeza que deseja excluir permanentemente o item selecionado do banco de dados?", vbQuestion + vbYesNo) = vbYes Then
        itemunico = Me.lstProcedimentos.List(selected, 0) & Me.lstProcedimentos.List(selected, 1) & Me.lstProcedimentos.List(selected, 3)
        DeleteID = VBA.Split(isExistsData(itemunico), ";")(1)
        rs.Open "DELETE FROM tbProcedimentos WHERE ID = " & DeleteID, conn
        conn.Close
        Call UpdateListBox
        MsgBox "Item deletado do banco de dados!", vbInformation
    End If
End Sub

Private Sub btnLancamento_Click()
    Dim oFicha      As cFichaProcedimento

    If ValidateEmptyControls(Me) Or Not VBA.IsDate(Me.txt_databpa.Value) Then Exit Sub
    
    Set oFicha = New cFichaProcedimento
    
    With oFicha
        .ProfissionalNome = Me.cboProfissional.Value
        .ProcedimentoNome = Me.cboProcedimento.Value
        .Quantidade = Me.txtQuantidade.Value
        .DataInicial = Me.txt_databpa.Value
        .InsertOrSumReg (Me.cboProfissional.Value & Me.cboProcedimento.Value & Me.txt_databpa.Value)
        Call UpdateListBox
    End With
    
    Call ClearFields(Me)
    Me.cboProfissional.SetFocus
    Me.lbValida.Caption = "FICHA LANÇADA"
End Sub

Private Sub btnSelect_Click()
    Me.lstProcedimentos.selected(Me.lstProcedimentos.ListCount - 1) = True
End Sub

Private Sub chk_otherdate_Click()
    Me.txt_databpa.Enabled = Me.chk_otherdate.Value
End Sub

Private Sub txt_databpa_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Me.txt_databpa.MaxLength = 10
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            Select Case VBA.Len(Me.txt_databpa.Value)
                Case 2, 5
                    Me.txt_databpa.SelText = "/"
            End Select
        Case Else
            KeyAscii = 0
    End Select
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
    Call UpdateListBox
End Sub

Private Sub UpdateListBox()
#If booAux Then
    Dim conn        As ADODB.Connection
    Dim rs          As ADODB.Recordset
#Else
    Dim conn        As Object
    Dim rs          As Object
#End If
    Dim arrAux  As Variant
    Dim iCounter    As Long
    
    Set conn = DataBaseConnection()
    Set rs = myRecordSet()
    
    rs.Open "SELECT PROFESSIONAL, PROCEDIMENTO, QUANTIDADE, INITIAL_DATE FROM tbProcedimentos", conn
    
    If Not rs.EOF Then
        arrAux = rs.GetRows
        conn.Close
        arrAux = Array2DTranspose(arrAux)
        
        Me.lstProcedimentos.ColumnCount = UBound(arrAux, 2) + 1
        Me.lstProcedimentos.List = arrAux
    End If
End Sub
