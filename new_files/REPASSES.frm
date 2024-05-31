Attribute VB_Name = "REPASSES"
Attribute VB_Base = "0{AD5010CE-9CBF-4DD7-A477-62A970855C7B}{45DB8CA9-B04C-4F29-8868-FC6AC3C916C9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

Private Sub DATA_CADASTRO_Change()
    'Formata : dd/mm/aa
    If Len(DATA_CADASTRO) = 2 Or Len(DATA_CADASTRO) = 5 Then
        DATA_CADASTRO.Text = DATA_CADASTRO.Text & "/"
         DATA_CADASTRO.SelStart = Len(DATA_CADASTRO)
    End If
End Sub

Private Sub DATA_CADASTRO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    'Limita a Qde de caracteres
   DATA_CADASTRO.MaxLength = 10
 
    'para permitir que apenas números sejam digitados
    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        KeyAscii = 0
    End If
 
End Sub
Private Sub DATA_CADASTRO_Enter()

DATA_CADASTRO.Text = Date

End Sub

Private Sub NOMES_Change()

End Sub

Private Sub REPASSE_Enter()
    If REPASSE.Text = "" Then
        REPASSE.Text = "0,00"
    End If
End Sub
Private Sub REPASSE_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        REPASSE.Text = Replace(REPASSE.Text, ",", "")
        REPASSE.Text = Replace(REPASSE.Text, ".", "")
        
        If Mid(REPASSE.Text, 1, 1) = "0" Then
            REPASSE.Text = Mid(REPASSE.Text, 2, Len(REPASSE.Text))
        End If
        
        REPASSE.Text = Mid(REPASSE.Text, 1, Len(REPASSE.Text) - 1) & "," & Mid(REPASSE.Text, Len(REPASSE.Text), 2)
        a = InStr(1, REPASSE.Text, ",", vbTextCompare)
        b = Mid(REPASSE.Text, 1, a - 1)
        
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
        
        REPASSE.Text = c & Mid(REPASSE.Text, Len(REPASSE.Text) - 1, 2)
        
    Else
    
        KeyAscii = 0
        
    End If
End Sub
'Parte importante que eu não havia percebido.

Private Sub REPASSE_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then
        REPASSE.Text = Replace(REPASSE.Text, ",", "")
        REPASSE.Text = Replace(REPASSE.Text, ".", "")
        
        If Len(REPASSE.Text) > 3 Then
            REPASSE.Text = Mid(REPASSE.Text, 1, Len(REPASSE.Text) - 2) & "," & Mid(REPASSE.Text, Len(REPASSE.Text) - 1, 2)
        Else
        
            If Len(REPASSE.Text) = 3 Then
                REPASSE.Text = Mid(REPASSE, 1, 1) & "," & Mid(REPASSE.Text, 2, 2)
            Else
            
                If Len(REPASSE.Text) = 2 Then
                    REPASSE.Text = "0," & Mid(REPASSE.Text, 1, 2)
                Else
                    If Len(REPASSE.Text) = 0 Then
                        REPASSE.Text = "0,00" & Mid(REPASSE.Text, 1, 2)
                        Exit Sub
                    End If
                End If
            End If
            
        End If
        
        a = InStr(1, REPASSE.Text, ",", vbTextCompare)
        b = Mid(REPASSE.Text, 1, a - 1)
        
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
        
        REPASSE.Text = c & Mid(REPASSE.Text, Len(REPASSE.Text) - 2, 3)
End If
End Sub



Private Sub DESCRICAO_Enter()

If REPASSE > "0" Then
DESCRICAO.Text = "REPASSE DE CONGREGAÇÃO"
End If

End Sub

Private Sub DATA_Enter()


DATA = Date

End Sub



Private Sub UserForm_Activate()
    Dim cod
    ' Ativa a auto numeração no Form
    Sheets("REPASSES").Select
    cod = Range("D1001").End(xlUp).Offset(0, 0).Value
    Me.CODIGO = cod + 1
End Sub



Private Sub SALVAR_Click()

Dim cod
    
 

'Ativar a primeira planilha

    ' Adicionar dados na planilha ( Nesta parte, o codigo selecionará a Plan1 e gravará seus dados).
    Sheets("REPASSES").Select
'Aqui o codigo seleciona a linha em branco e inicia a gravação
    ' lembre se que poderá ser necessário a digitação de um zero (0) na primeira linha para iniciar a gravação já que os numeros começam com zero.
    Range("D1001").End(xlUp).Offset(1, 0).Select

'Procurar a primeira célula vazia
Do
  If Not (IsEmpty(ActiveCell)) Then
      ActiveCell.Offset(1, 0).Select
  End If
Loop Until IsEmpty(ActiveCell) = True

If NOMES.Text = "" Then

 Me.NOMES.SetFocus
 
MsgBox ("Campo Obrigatório 'Nome da Congregação'"), vbOKOnly, ("")

     Exit Sub

End If


'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.Offset(0, 0).Value = CODIGO.Value
ActiveCell.Offset(0, 1).Value = NOMES.Value
ActiveCell.Offset(0, 2).Value = REPASSE.Value
ActiveCell.Offset(0, 3).Value = DESCRICAO.Value
ActiveCell.Offset(0, 4).Value = RECIBO.Value
ActiveCell.Offset(0, 5).Value = DATA_CADASTRO.Value
ActiveCell.Offset(0, 6).Value = DATA.Value


 ' Atualizar a auto-numeração lembrado que caso queira adicionar um label, troque a opção REPASSE pelo seu label
    cod = Range("D1001").End(xlUp).Offset(0, 0).Value
    
     'Aqui eis nosso contador automatico responsável pela contagem e numeração automatica.
    Me.CODIGO = cod + 1
    
      ' Apos adiconar os dados na planilha limpa o campo
Me.NOMES = Empty
Me.REPASSE = Empty
Me.DESCRICAO = Empty
Me.DATA = Empty
Me.RECIBO = Empty
Me.DATA_CADASTRO = Empty

For Each cell In [F2:F1000]
       If cell > "" Then
       numero = Str(cell.Value)
       cell.Activate
      ActiveCell.FormulaR1C1 = numero
    
  Else
  End If
  Next

End Sub




