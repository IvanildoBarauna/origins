Attribute VB_Name = "fGerarTabela"
Attribute VB_Base = "0{787BF8E3-652A-4117-996C-B440A94CC8B9}{9C9E52CE-7A6A-4B86-9833-A0F47F59C633}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Sub CalcularSAC()
    Dim lo As ListObject
    Dim lr As ListRow
    Dim ValorTotal As Double
    Dim Entrada As Double
    Dim ValorFinanciado As Double
    Dim Prestacoes As Integer
    Dim Taxa As Double
        
    Application.ScreenUpdating = False
        
    ValorTotal = txtValor.Text
    Entrada = txtEntrada.Text
    ValorFinanciado = ValorTotal - Entrada
    Prestacoes = txtPrestacoes.Text
    Taxa = CDbl(txtJuros.Text) / 100
    
    With shtSAC
        .Unprotect
        .Range("ValorTotal").Value2 = ValorTotal
        .Range("Entrada").Value2 = Entrada
        .Range("ValorFinanciado").Value2 = ValorFinanciado
        .Range("Taxa").Value2 = Taxa
        .Range("Prestacoes").Value2 = Prestacoes
        .Range("ValorAmortizacao").FormulaR1C1 = "=ValorFinanciado/Prestacoes"
    End With
    
    'Gerar a Tabela
    Set lo = shtSAC.ListObjects("tbSAC")
    If Not lo.DataBodyRange Is Nothing Then lo.DataBodyRange.Delete
    AddRows Prestacoes, lo
    With lo
        .DataBodyRange(1, .ListColumns("Saldo Inicial").Index).Value = ValorFinanciado
    End With
    
    shtSAC.Protect
    
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
        MsgBox "Infome um valor para calcular o financiamento", vbCritical, "SAC"
        txtValor.SetFocus
        Exit Function
    End If
    
    If CDbl(txtEntrada.Text) >= CDbl(txtValor.Text) Then
        MsgBox "A entrada deve ser menor que o valor do bem", vbCritical, "SAC"
        txtEntrada.SetFocus
        Exit Function
    End If
    
    If Not Int(txtPrestacoes.Text) > 0 Then
        MsgBox "Infome o número de prestações do financimento", vbCritical, "SAC"
        txtPrestacoes.SetFocus
        Exit Function
    End If
    
    If Not CDbl(txtJuros.Text) > 0 Then
        MsgBox "Infome a taxa de juros do financimento", vbCritical, "SAC"
        txtJuros.SetFocus
        Exit Function
    End If
    
    Validado = True
    
End Function

Private Sub btnCalcular_Click()
    If Not Validado Then Exit Sub
    CalcularSAC
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
    With shtSAC
        txtValor.Text = Format(.Range("ValorTotal").Value2, "0.00")
        txtEntrada.Text = Format(.Range("Entrada").Value2, "0.00")
        txtPrestacoes.Text = Format(.Range("Prestacoes").Value2, "0")
        txtJuros.Text = Format(.Range("Taxa").Value2 * 100, "0.00")
        txtValor.SetFocus
    End With
End Sub
