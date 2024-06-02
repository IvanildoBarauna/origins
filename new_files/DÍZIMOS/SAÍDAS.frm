Attribute VB_Name = "SAÍDAS"
Attribute VB_Base = "0{5D87D81D-B008-4F09-941F-C24183F258E7}{41911E0F-452B-461E-8FDD-5F7131638046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub DATA_VENCIMENTO_Change()
    'Formata : dd/mm/aa
    If Len(DATA_VENCIMENTO) = 2 Or Len(DATA_VENCIMENTO) = 5 Then
        DATA_VENCIMENTO.Text = DATA_VENCIMENTO.Text & "/"
         DATA_VENCIMENTO.SelStart = Len(DATA_VENCIMENTO)
    End If
End Sub

Private Sub DATA_VENCIMENTO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
   DATA_VENCIMENTO.MaxLength = 10
 
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
 
End Sub

Private Sub DATA_PAGAMENTO_Change()
    'Formata : dd/mm/aa
    If Len(DATA_PAGAMENTO) = 2 Or Len(DATA_PAGAMENTO) = 5 Then
        DATA_PAGAMENTO.Text = DATA_PAGAMENTO.Text & "/"
         DATA_PAGAMENTO.SelStart = Len(DATA_PAGAMENTO)
    End If
End Sub

Private Sub DATA_PAGAMENTO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
   DATA_PAGAMENTO.MaxLength = 10
 
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
 
End Sub

Private Sub VALOR_DOCUMENTO_Enter()
    If VALOR_DOCUMENTO.Text = "" Then
        VALOR_DOCUMENTO.Text = "0,00"
    End If
End Sub
Private Sub VALOR_DOCUMENTO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        VALOR_DOCUMENTO.Text = Replace(VALOR_DOCUMENTO.Text, ",", "")
        VALOR_DOCUMENTO.Text = Replace(VALOR_DOCUMENTO.Text, ".", "")
        
        If Mid(VALOR_DOCUMENTO.Text, 1, 1) = "0" Then
            VALOR_DOCUMENTO.Text = Mid(VALOR_DOCUMENTO.Text, 2, Len(VALOR_DOCUMENTO.Text))
        End If
        
        VALOR_DOCUMENTO.Text = Mid(VALOR_DOCUMENTO.Text, 1, Len(VALOR_DOCUMENTO.Text) - 1) & "," & Mid(VALOR_DOCUMENTO.Text, Len(VALOR_DOCUMENTO.Text), 2)
        a = InStr(1, VALOR_DOCUMENTO.Text, ",", vbTextCompare)
        b = Mid(VALOR_DOCUMENTO.Text, 1, a - 1)
        
        For x = 1 To Len(b)
            cont = cont + 1
            c = Mid(b, Len(b) - x + 1, 1) & c
            If cont = 3 Then
                If Len(b) > 3 Then
                    c = "." & c
                    cont = 0
                End If
            End If
        Next x
        
        If Mid(c, 1, 1) = "." Then
            c = Mid(c, 2, Len(c))
        End If
        
        VALOR_DOCUMENTO.Text = c & Mid(VALOR_DOCUMENTO.Text, Len(VALOR_DOCUMENTO.Text) - 1, 2)
        
    Else
    
        KeyAscii = 0
        
    End If
End Sub
'Parte importante que eu não havia percebido.

Private Sub VALOR_DOCUMENTO_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then
        VALOR_DOCUMENTO.Text = Replace(VALOR_DOCUMENTO.Text, ",", "")
        VALOR_DOCUMENTO.Text = Replace(VALOR_DOCUMENTO.Text, ".", "")
        
        If Len(VALOR_DOCUMENTO.Text) > 3 Then
            VALOR_DOCUMENTO.Text = Mid(VALOR_DOCUMENTO.Text, 1, Len(VALOR_DOCUMENTO.Text) - 2) & "," & Mid(VALOR_DOCUMENTO.Text, Len(VALOR_DOCUMENTO.Text) - 1, 2)
        Else
        
            If Len(VALOR_DOCUMENTO.Text) = 3 Then
                VALOR_DOCUMENTO.Text = Mid(VALOR_DOCUMENTO, 1, 1) & "," & Mid(VALOR_DOCUMENTO.Text, 2, 2)
            Else
            
                If Len(VALOR_DOCUMENTO.Text) = 2 Then
                    VALOR_DOCUMENTO.Text = "0," & Mid(VALOR_DOCUMENTO.Text, 1, 2)
                Else
                    If Len(VALOR_DOCUMENTO.Text) = 0 Then
                        VALOR_DOCUMENTO.Text = "0,00" & Mid(VALOR_DOCUMENTO.Text, 1, 2)
                        Exit Sub
                    End If
                End If
            End If
            
        End If
        
        a = InStr(1, VALOR_DOCUMENTO.Text, ",", vbTextCompare)
        b = Mid(VALOR_DOCUMENTO.Text, 1, a - 1)
        
        For x = 1 To Len(b)
            cont = cont + 1
            c = Mid(b, Len(b) - x + 1, 1) & c
            If cont = 3 Then
                If Len(b) > 3 Then
                    c = "." & c
                    cont = 0
                End If
            End If
        Next x
        
        If Mid(c, 1, 1) = "." Then
            c = Mid(c, 2, Len(c))
        End If
        
        VALOR_DOCUMENTO.Text = c & Mid(VALOR_DOCUMENTO.Text, Len(VALOR_DOCUMENTO.Text) - 2, 3)
End If
End Sub

Private Sub VALOR_PAGO_Enter()
    If VALOR_PAGO.Text = "" Then
        VALOR_PAGO.Text = "0,00"
    End If
    
    If VALOR_DOCUMENTO > 0 Then
VALOR_PAGO.Text = VALOR_DOCUMENTO
End If
    
End Sub
Private Sub VALOR_PAGO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        VALOR_PAGO.Text = Replace(VALOR_PAGO.Text, ",", "")
        VALOR_PAGO.Text = Replace(VALOR_PAGO.Text, ".", "")
        
        If Mid(VALOR_PAGO.Text, 1, 1) = "0" Then
            VALOR_PAGO.Text = Mid(VALOR_PAGO.Text, 2, Len(VALOR_PAGO.Text))
        End If
        
        VALOR_PAGO.Text = Mid(VALOR_PAGO.Text, 1, Len(VALOR_PAGO.Text) - 1) & "," & Mid(VALOR_PAGO.Text, Len(VALOR_PAGO.Text), 2)
        a = InStr(1, VALOR_PAGO.Text, ",", vbTextCompare)
        b = Mid(VALOR_PAGO.Text, 1, a - 1)
        
        For x = 1 To Len(b)
            cont = cont + 1
            c = Mid(b, Len(b) - x + 1, 1) & c
            If cont = 3 Then
                If Len(b) > 3 Then
                    c = "." & c
                    cont = 0
                End If
            End If
        Next x
        
        If Mid(c, 1, 1) = "." Then
            c = Mid(c, 2, Len(c))
        End If
        
        VALOR_PAGO.Text = c & Mid(VALOR_PAGO.Text, Len(VALOR_PAGO.Text) - 1, 2)
        
    Else
    
        KeyAscii = 0
        
    End If
End Sub
'Parte importante que eu não havia percebido.

Private Sub VALOR_PAGO_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then
        VALOR_PAGO.Text = Replace(VALOR_PAGO.Text, ",", "")
        VALOR_PAGO.Text = Replace(VALOR_PAGO.Text, ".", "")
        
        If Len(VALOR_PAGO.Text) > 3 Then
            VALOR_PAGO.Text = Mid(VALOR_PAGO.Text, 1, Len(VALOR_PAGO.Text) - 2) & "," & Mid(VALOR_PAGO.Text, Len(VALOR_PAGO.Text) - 1, 2)
        Else
        
            If Len(VALOR_PAGO.Text) = 3 Then
                VALOR_PAGO.Text = Mid(VALOR_PAGO, 1, 1) & "," & Mid(VALOR_PAGO.Text, 2, 2)
            Else
            
                If Len(VALOR_PAGO.Text) = 2 Then
                    VALOR_PAGO.Text = "0," & Mid(VALOR_PAGO.Text, 1, 2)
                Else
                    If Len(VALOR_PAGO.Text) = 0 Then
                        VALOR_PAGO.Text = "0,00" & Mid(VALOR_PAGO.Text, 1, 2)
                        Exit Sub
                    End If
                End If
            End If
            
        End If
        
        a = InStr(1, VALOR_PAGO.Text, ",", vbTextCompare)
        b = Mid(VALOR_PAGO.Text, 1, a - 1)
        
        For x = 1 To Len(b)
            cont = cont + 1
            c = Mid(b, Len(b) - x + 1, 1) & c
            If cont = 3 Then
                If Len(b) > 3 Then
                    c = "." & c
                    cont = 0
                End If
            End If
        Next x
        
        If Mid(c, 1, 1) = "." Then
            c = Mid(c, 2, Len(c))
        End If
        
        VALOR_PAGO.Text = c & Mid(VALOR_PAGO.Text, Len(VALOR_PAGO.Text) - 2, 3)
End If
End Sub

Private Sub DATA_PAGAMENTO_Enter()



If DATA_PAGAMENTO = "" Then
DATA_PAGAMENTO.Text = Date
End If

End Sub

Private Sub DATA_VENCIMENTO_Enter()



If DATA_VENCIMENTO = "" Then
DATA_VENCIMENTO.Text = Date
End If

End Sub









Private Sub DATA_Enter()


DATA = Date

End Sub

Private Sub UserForm_Activate()
    Dim cod
    ' Ativa a auto numeração no Form
    Sheets("SAÍDAS").Select
    cod = Range("D1001").End(xlUp).Offset(0, 0).Value
    Me.CODIGO = cod + 1
End Sub



Private Sub SALVAR_Click()

Dim cod
    
 

'Ativar a primeira planilha

    ' Adicionar dados na planilha ( Nesta parte, o codigo selecionará a Plan1 e gravará seus dados).
    Sheets("SAÍDAS").Select
'Aqui o codigo seleciona a linha em branco e inicia a gravação
    ' lembre se que poderá ser necessário a digitação de um zero (0) na primeira linha para iniciar a gravação já que os numeros começam com zero.
    Range("D1001").End(xlUp).Offset(1, 0).Select

'Procurar a primeira célula vazia
Do Until IsEmpty(ActiveCell) = True
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop

If NOMES.Text = "" Then
     Me.NOMES.SetFocus
     MsgBox ("Campo Obrigatório 'Tipo da Saída é Obrigatório'"), vbOKOnly, ("")
     Exit Sub
End If


'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.Offset(0, 0).Value = CODIGO.Value
ActiveCell.Offset(0, 1).Value = CENTRO.Value
ActiveCell.Offset(0, 2).Value = NOMES.Value
ActiveCell.Offset(0, 3).Value = RECIBO.Value
ActiveCell.Offset(0, 4).Value = DESCRICAO.Value
ActiveCell.Offset(0, 5).Value = DATA_VENCIMENTO.Value
ActiveCell.Offset(0, 6).Value = DATA_PAGAMENTO.Value
ActiveCell.Offset(0, 7).Value = VALOR_DOCUMENTO.Value
ActiveCell.Offset(0, 8).Value = VALOR_PAGO.Value
ActiveCell.Offset(0, 9).Value = DATA.Value



 ' Atualizar a auto-numeração lembrado que caso queira adicionar um label, troque a opção VALOR_DOCUMENTO pelo seu label
    cod = Range("D1001").End(xlUp).Offset(0, 0).Value
    
     'Aqui eis nosso contador automatico responsável pela contagem e numeração automatica.
    Me.CODIGO = cod + 1
    
      ' Apos adiconar os dados na planilha limpa o campo
Me.NOMES = Empty
Me.CENTRO = Empty
Me.RECIBO = Empty
Me.DESCRICAO = Empty
Me.DATA_VENCIMENTO = Empty
Me.DATA_PAGAMENTO = Empty
Me.VALOR_DOCUMENTO = Empty
Me.VALOR_PAGO = Empty
Me.DATA = Empty

For Each cell In [K2:L1000]
       If cell > "" Then
       numero = Str(cell.Value)
       cell.Activate
      ActiveCell.FormulaR1C1 = numero
    
  Else
  End If
  Next




NOMES.SetFocus

End Sub




 
