Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{E4123693-963E-454F-A6A6-265E78E21D7C}{F4DB8480-5C8E-4E26-8A65-BC5896BF3445}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Sub CARREGAR1()
ComboBox2.AddItem "SULDESTE"
ComboBox2.AddItem "SUL"
ComboBox2.AddItem "TODOS"
ComboBox2 = "TODOS"

End Sub
Sub CARREGAR2()
Me.ComboBox1.Clear 'APAGA TODAS AS INFORMAÇOES
Me.ComboBox1.AddItem "TODOS"
    LIN = 2
    Do Until Plan1.Cells(LIN, 1) = "" ' PEGA A ULTIMA LINHA VAZIA
         Me.ComboBox1.AddItem Plan1.Cells(LIN, 3) 'ADD TUDO QUE TIVER NA COLUNA 3 DA PLANILHA
         LIN = LIN + 1
     Loop
Me.ComboBox1 = "TODOS"
End Sub
Sub CARREGAR3()
    Dim REGIAO As String, linha As Integer, colunaESTADO As Integer, colunaREGIAO As Integer
    linha = 2
    colunaESTADO = 3 ' COLUNA ESTADOS
    colunaREGIAO = 2 ' COLUNA REGIÃO
    ComboBox3.Clear

    If ComboBox2 = "TODOS" Then ' CONDIÇÃO DA COMBOBOX QUE VAI PESQUISAR
    CARREGAR4 '  FUNÇÃO 4 CASO A COMBOBOX REGIÃO SEJA TODOS
    'Exit Sub
    Else
    REGIAO = ComboBox2 ' FILTRO DA REGIÃO
    End If
    With Sheets("PEIA")
        Do While Not IsEmpty(.Cells(linha, colunaESTADO)) ' PESQUISA LINHA EM BRANCO
            If .Cells(linha, colunaREGIAO).Value = REGIAO Then ' FILTRO DA REGIÃO
            ComboBox3.AddItem .Cells(linha, colunaESTADO).Value ' ADICIONA O VALOR A COMBOBOX
            End If
            linha = linha + 1
        Loop
    End With
End Sub
Sub CARREGAR4()
Me.ComboBox3.Clear 'APAGA TODAS AS INFORMAÇOES
Me.ComboBox3.AddItem "TODOS"
    LIN = 2
    Do Until Plan1.Cells(LIN, 1) = "" ' PEGA A ULTIMA LINHA VAZIA
         Me.ComboBox3.AddItem Plan1.Cells(LIN, 3) 'ADD TUDO QUE TIVER NA COLUNA 3 DA PLANILHA
         LIN = LIN + 1
     Loop
Me.ComboBox3 = "TODOS"
End Sub
Private Sub ComboBox2_Change()
CARREGAR3
End Sub

Private Sub UserForm_Initialize()
CARREGAR1
CARREGAR2
CARREGAR3
End Sub
