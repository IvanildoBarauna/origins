Attribute VB_Name = "LANÇAMENTOS"
Attribute VB_Base = "0{6B328743-7913-4372-8AC5-4C595594BF11}{657535EE-C3EC-4C3C-A52D-3565AC9E85B7}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False



Private Sub CheckBox1_Click()

End Sub

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

Private Sub DIZIMO_Enter()
    If DIZIMO.Text = "" Then
        DIZIMO.Text = "0,00"
    End If
End Sub
Private Sub DIZIMO_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        DIZIMO.Text = Replace(DIZIMO.Text, ",", "")
        DIZIMO.Text = Replace(DIZIMO.Text, ".", "")
        
        If Mid(DIZIMO.Text, 1, 1) = "0" Then
            DIZIMO.Text = Mid(DIZIMO.Text, 2, Len(DIZIMO.Text))
        End If
        
        DIZIMO.Text = Mid(DIZIMO.Text, 1, Len(DIZIMO.Text) - 1) & "," & Mid(DIZIMO.Text, Len(DIZIMO.Text), 2)
        a = InStr(1, DIZIMO.Text, ",", vbTextCompare)
        b = Mid(DIZIMO.Text, 1, a - 1)
        
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
        
        DIZIMO.Text = c & Mid(DIZIMO.Text, Len(DIZIMO.Text) - 1, 2)
        
    Else
    
        KeyAscii = 0
        
    End If
End Sub
'Parte importante que eu não havia percebido.

Private Sub DIZIMO_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then
        DIZIMO.Text = Replace(DIZIMO.Text, ",", "")
        DIZIMO.Text = Replace(DIZIMO.Text, ".", "")
        
        If Len(DIZIMO.Text) > 3 Then
            DIZIMO.Text = Mid(DIZIMO.Text, 1, Len(DIZIMO.Text) - 2) & "," & Mid(DIZIMO.Text, Len(DIZIMO.Text) - 1, 2)
        Else
        
            If Len(DIZIMO.Text) = 3 Then
                DIZIMO.Text = Mid(DIZIMO, 1, 1) & "," & Mid(DIZIMO.Text, 2, 2)
            Else
            
                If Len(DIZIMO.Text) = 2 Then
                    DIZIMO.Text = "0," & Mid(DIZIMO.Text, 1, 2)
                Else
                    If Len(DIZIMO.Text) = 0 Then
                        DIZIMO.Text = "0,00" & Mid(DIZIMO.Text, 1, 2)
                        Exit Sub
                    End If
                End If
            End If
            
        End If
        
        a = InStr(1, DIZIMO.Text, ",", vbTextCompare)
        b = Mid(DIZIMO.Text, 1, a - 1)
        
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
        
        DIZIMO.Text = c & Mid(DIZIMO.Text, Len(DIZIMO.Text) - 2, 3)
End If
End Sub

Private Sub OFERTA_Enter()
    If OFERTA.Text = "" Then
        OFERTA.Text = "0,00"
    End If
End Sub
Private Sub OFERTA_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        OFERTA.Text = Replace(OFERTA.Text, ",", "")
        OFERTA.Text = Replace(OFERTA.Text, ".", "")
        
        If Mid(OFERTA.Text, 1, 1) = "0" Then
            OFERTA.Text = Mid(OFERTA.Text, 2, Len(OFERTA.Text))
        End If
        
        OFERTA.Text = Mid(OFERTA.Text, 1, Len(OFERTA.Text) - 1) & "," & Mid(OFERTA.Text, Len(OFERTA.Text), 2)
        a = InStr(1, OFERTA.Text, ",", vbTextCompare)
        b = Mid(OFERTA.Text, 1, a - 1)
        
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
        
        OFERTA.Text = c & Mid(OFERTA.Text, Len(OFERTA.Text) - 1, 2)
        
    Else
    
        KeyAscii = 0
        
    End If
End Sub
'Parte importante que eu não havia percebido.

Private Sub OFERTA_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then
        OFERTA.Text = Replace(OFERTA.Text, ",", "")
        OFERTA.Text = Replace(OFERTA.Text, ".", "")
        
        If Len(OFERTA.Text) > 3 Then
            OFERTA.Text = Mid(OFERTA.Text, 1, Len(OFERTA.Text) - 2) & "," & Mid(OFERTA.Text, Len(OFERTA.Text) - 1, 2)
        Else
        
            If Len(OFERTA.Text) = 3 Then
                OFERTA.Text = Mid(OFERTA, 1, 1) & "," & Mid(OFERTA.Text, 2, 2)
            Else
            
                If Len(OFERTA.Text) = 2 Then
                    OFERTA.Text = "0," & Mid(OFERTA.Text, 1, 2)
                Else
                    If Len(OFERTA.Text) = 0 Then
                        OFERTA.Text = "0,00" & Mid(OFERTA.Text, 1, 2)
                        Exit Sub
                    End If
                End If
            End If
            
        End If
        
        a = InStr(1, OFERTA.Text, ",", vbTextCompare)
        b = Mid(OFERTA.Text, 1, a - 1)
        
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
        
        OFERTA.Text = c & Mid(OFERTA.Text, Len(OFERTA.Text) - 2, 3)
End If
End Sub

Private Sub OFERTAESP_Enter()
    If OFERTAESP.Text = "" Then
        OFERTAESP.Text = "0,00"
    End If
End Sub
Private Sub OFERTAESP_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If KeyAscii > 47 And KeyAscii < 58 Then
        OFERTAESP.Text = Replace(OFERTAESP.Text, ",", "")
        OFERTAESP.Text = Replace(OFERTAESP.Text, ".", "")
        
        If Mid(OFERTAESP.Text, 1, 1) = "0" Then
            OFERTAESP.Text = Mid(OFERTAESP.Text, 2, Len(OFERTAESP.Text))
        End If
        
        OFERTAESP.Text = Mid(OFERTAESP.Text, 1, Len(OFERTAESP.Text) - 1) & "," & Mid(OFERTAESP.Text, Len(OFERTAESP.Text), 2)
        a = InStr(1, OFERTAESP.Text, ",", vbTextCompare)
        b = Mid(OFERTAESP.Text, 1, a - 1)
        
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
        
        OFERTAESP.Text = c & Mid(OFERTAESP.Text, Len(OFERTAESP.Text) - 1, 2)
        
    Else
    
        KeyAscii = 0
        
    End If
End Sub
'Parte importante que eu não havia percebido.

Private Sub OFERTAESP_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)

If KeyCode = 8 Or KeyCode = 46 Then
        OFERTAESP.Text = Replace(OFERTAESP.Text, ",", "")
        OFERTAESP.Text = Replace(OFERTAESP.Text, ".", "")
        
        If Len(OFERTAESP.Text) > 3 Then
            OFERTAESP.Text = Mid(OFERTAESP.Text, 1, Len(OFERTAESP.Text) - 2) & "," & Mid(OFERTAESP.Text, Len(OFERTAESP.Text) - 1, 2)
        Else
        
            If Len(OFERTAESP.Text) = 3 Then
                OFERTAESP.Text = Mid(OFERTAESP, 1, 1) & "," & Mid(OFERTAESP.Text, 2, 2)
            Else
            
                If Len(OFERTAESP.Text) = 2 Then
                    OFERTAESP.Text = "0," & Mid(OFERTAESP.Text, 1, 2)
                Else
                    If Len(OFERTAESP.Text) = 0 Then
                        OFERTAESP.Text = "0,00" & Mid(OFERTAESP.Text, 1, 2)
                        Exit Sub
                    End If
                End If
            End If
            
        End If
        
        a = InStr(1, OFERTAESP.Text, ",", vbTextCompare)
        b = Mid(OFERTAESP.Text, 1, a - 1)
        
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
        
        OFERTAESP.Text = c & Mid(OFERTAESP.Text, Len(OFERTAESP.Text) - 2, 3)
End If
End Sub



Private Sub NOMES_Change()

With Worksheets("CADASTROS").Range("C:C")

Set buscar = .Find(NOMES.Value, LookIn:=xlValues, LookAt:=xlPart)

If Not buscar Is Nothing Then

On Error Resume Next

CONGREGACAO.Value = buscar.Offset(0, 1).Value
OBREIRO.Value = buscar.Offset(0, 2).Value

End If

End With


End Sub



'On Error Resume Next
'CONGREGACAO = Application.WorksheetFunction.VLookup(CStr(NOMES), Plan2.Range("A1:D1001"), 2, 0)
'CONGREGACAO.Text = CStr(Application.VLookup(NOMES.Text, Plan2.Range("A1:D1001"), 4, 0))





Private Sub DESCRICAO_Enter()



If DIZIMO > "0" Then
DESCRICAO.Text = "DÍZIMO"
End If

End Sub

Private Sub DATA_Enter()


DATA = Date

End Sub



Private Sub UserForm_Activate()
    Dim cod
    ' Ativa a auto numeração no Form
    Sheets("ENTRADAS").Select
    cod = Range("D1001").End(xlUp).Offset(0, 0).Value
    Me.CODIGO = cod + 1
End Sub



Private Sub SALVAR_Click()

Dim cod
    
 

'Ativar a primeira planilha

    ' Adicionar dados na planilha ( Nesta parte, o codigo selecionará a Plan1 e gravará seus dados).
    Sheets("ENTRADAS").Select
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
 
MsgBox ("Campo Obrigatório 'Nome do Cadastrado'"), vbOKOnly, ("")

     Exit Sub

End If


'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.Offset(0, 0).Value = CODIGO.Value
ActiveCell.Offset(0, 1).Value = NOMES.Value
ActiveCell.Offset(0, 2).Value = CONGREGACAO.Value
ActiveCell.Offset(0, 3).Value = DIZIMO.Value
ActiveCell.Offset(0, 4).Value = OFERTA.Value
ActiveCell.Offset(0, 5).Value = OFERTAESP.Value
ActiveCell.Offset(0, 6).Value = DESCRICAO.Value
ActiveCell.Offset(0, 7).Value = RECIBO.Value
ActiveCell.Offset(0, 8).Value = DATA_CADASTRO.Value
ActiveCell.Offset(0, 9).Value = DATA.Value
ActiveCell.Offset(0, 11).Value = OBREIRO.Value


 ' Atualizar a auto-numeração lembrado que caso queira adicionar um label, troque a opção DIZIMO pelo seu label
    cod = Range("D1001").End(xlUp).Offset(0, 0).Value
    
     'Aqui eis nosso contador automatico responsável pela contagem e numeração automatica.
    Me.CODIGO = cod + 1
    
      ' Apos adiconar os dados na planilha limpa o campo
Me.NOMES = Empty
Me.CONGREGACAO = Empty
Me.DIZIMO = Empty
Me.OFERTA = Empty
Me.OFERTAESP = Empty
Me.DESCRICAO = Empty
Me.DATA = Empty
Me.RECIBO = Empty
Me.DATA_CADASTRO = Empty
Me.OBREIRO = Empty

For Each cell In [g2:I1000]
       If cell > "" Then
       numero = Str(cell.Value)
       cell.Activate
      ActiveCell.FormulaR1C1 = numero
    
  Else
  End If
  Next
  
  
End Sub



