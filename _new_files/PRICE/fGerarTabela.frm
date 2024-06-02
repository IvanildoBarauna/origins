Attribute VB_Name = "fGerarTabela"
Attribute VB_Base = "0{D6398B2A-21C2-4AFD-B2D2-87D3023D4F07}{A3F95232-99D4-497D-BDE3-18CC4964DB7A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CalcularPRICE()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim ValorTotal As Double
    Dim Entrada As Double
    Dim ValorFinanciado As Double
    Dim Prestacoes As Integer
    Dim Taxa As Double
        
    Application.ScreenUpdating = False
        
    ValorTotal = txtValor.Value
    Entrada = txtEntrada.Value
    ValorFinanciado = ValorTotal - Entrada
    Prestacoes = txtPrestacoes.Value
    Taxa = CDbl(txtJuros.Value) / 100

    
    With shtPRICE
        .Unprotect
        .Range("ValorTotal").Value2 = ValorTotal
        .Range("Entrada").Value2 = Entrada
        .Range("ValorFinanciado").Value2 = ValorFinanciado
        .Range("Taxa").Value2 = Taxa
        .Range("Prestacoes").Value2 = Prestacoes
        .Range("ValorPrestacao").FormulaR1C1 = "=-PMT(Taxa,Prestacoes,ValorFinanciado)"
    End With
    
    'Gerar a Tabela
    Set lo = shtPRICE.ListObjects("tbPRICE")
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    AddRows Prestacoes, lo
    With lo
        .DataBodyRange(1, .ListColumns("Saldo Inicial").Index).Value = ValorFinanciado
    End With
    
    shtPRICE.Protect
    
    Application.ScreenUpdating = True
    
End Sub

Public Sub AddRows(ByVal pNumRows As Long, pListObject As ListObject)
    With pListObject
        If pNumRows = 1 Then
            .ListRows.Add
        Else
            .Resize .Range.Resize(1 + .ListRows.Count + pNumRows, .ListColumns.Count)
        End If
    End With
End Sub

Private Function Validado() As Boolean
    If Not CDbl(txtValor.Text) > 0 Then
        MsgBox "Infome um valor para calcular o financiamento", vbCritical, "PRICE"
        txtValor.SetFocus
        Exit Function
    End If
    
    If CDbl(txtEntrada.Text) >= CDbl(txtValor.Text) Then
        MsgBox "A entrada deve ser menor que o valor do bem", vbCritical, "PRICE"
        txtEntrada.SetFocus
        Exit Function
    End If
    
    If Not Int(txtPrestacoes.Text) > 0 Then
        MsgBox "Infome o número de prestações do financimento", vbCritical, "PRICE"
        txtPrestacoes.SetFocus
        Exit Function
    End If
    
    If Not CDbl(txtJuros.Text) > 0 Then
        MsgBox "Infome a taxa de juros do financimento", vbCritical, "PRICE"
        txtJuros.SetFocus
        Exit Function
    End If
    
    Validado = True
    
End Function

Private Sub btnCalcular_Click()
    If Not Validado Then Exit Sub
    CalcularPRICE
    Unload Me
End Sub

Private Sub btnCancelar_Click()
    Unload Me
End Sub

Private Sub btnLimpar_Click()
    txtValor.Text = Format(0, "0.00")
    txtEntrada.Text = Format(0, "0.00")
    txtPrestacoes.Text = Format(0, "0")
    txtJuros.Text = Format(0, "0.00")
    txtValor.SetFocus
End Sub

Private Sub spPrestacoes_Change(): txtPrestacoes.Text = spPrestacoes.Value: End Sub

Private Sub txtPrestacoes_Change()
    If txtPrestacoes.Text = "" Then txtPrestacoes.Text = 0
    txtPrestacoes.Text = Format(txtPrestacoes.Text, "0")
    If CInt(txtPrestacoes.Text) < spPrestacoes.Min Then txtPrestacoes.Text = spPrestacoes.Min
    If CInt(txtPrestacoes.Text) > spPrestacoes.Max Then txtPrestacoes.Text = spPrestacoes.Max
    spPrestacoes.Value = txtPrestacoes.Text
End Sub

Private Sub txtPrestacoes_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    On Error Resume Next
    Select Case KeyCode
        Case 38: spPrestacoes.Value = spPrestacoes.Value + spPrestacoes.SmallChange: KeyCode = 0
        Case 40: spPrestacoes.Value = spPrestacoes.Value - spPrestacoes.SmallChange: KeyCode = 0
    End Select
End Sub

Private Sub txtPrestacoes_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            'Permite
        Case Else
            KeyAscii = 1
    End Select
End Sub

Private Sub txtValor_Change()
    Dim sValor As String
    If txtValor.Text = "" Then txtValor.Text = 0
    sValor = Replace(txtValor, ",", "")
    txtValor.Text = Format(CDbl(sValor) / 100, "0.00")
End Sub

Private Sub txtValor_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            'Permite
        Case Else
            KeyAscii = 1
    End Select
End Sub

Private Sub txtEntrada_Change()
    Dim sValor As String
    If txtEntrada.Text = "" Then txtEntrada.Text = 0
    sValor = Replace(txtEntrada, ",", "")
    txtEntrada.Text = Format(CDbl(sValor) / 100, "0.00")
End Sub

Private Sub txtEntrada_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            'Permite
        Case Else
            KeyAscii = 1
    End Select
End Sub

Private Sub txtJuros_Change()
    Dim sValor As String
    If txtJuros.Text = "" Then txtJuros.Text = 0
    sValor = Replace(txtJuros, ",", "")
    txtJuros.Text = Format(CDbl(sValor) / 100, "0.00")
End Sub

Private Sub txtJuros_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    Select Case KeyAscii
        Case Asc("0") To Asc("9")
            'Permite
        Case Else
            KeyAscii = 1
    End Select
End Sub

Private Sub UserForm_Initialize()
    With shtPRICE
        txtValor.Text = Format(.Range("ValorTotal").Value2, "0.00")
        txtEntrada.Text = Format(.Range("Entrada").Value2, "0.00")
        txtPrestacoes.Text = Format(.Range("Prestacoes").Value2, "0")
        txtJuros.Text = Format(.Range("Taxa").Value2 * 100, "0.00")
        txtValor.SetFocus
    End With
End Sub
