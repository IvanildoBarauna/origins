Attribute VB_Name = "frmSQL"
Attribute VB_Base = "0{50B56314-500A-46C7-BAED-A75B2F8B39B5}{C729F4A9-5F58-4BEA-8290-35C326D8D9D3}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Option Compare Text
' Icons courtesy of Axialis Software - http://www.axialis.com

' ================================================================================================================
'                                                      FORMULÁRIO
' ================================================================================================================

' Inicialização do formulário

Private Sub UserForm_Initialize()
    ChDrive Left(ThisWorkbook.Path, 1)
    ChDir ThisWorkbook.Path
    lblAjuda = " Selecione o arquivo de dados, digite o código SQL e clique em 'Executar código SQL'."
    txtArquivoDados = ""
    txtArquivoImagem = ""
    
    If Range("ArquivoDados") <> "" Then
        txtArquivoDados = Range("ArquivoDados")
    End If
    
    If Range("ArquivoImagem") <> "" Then
        txtArquivoImagem = Range("ArquivoImagem")
    End If
    
    If Range("TamanhoFonte") = "" Then
        Range("TamanhoFonte") = 20
        txtTamanhoFonte = 20
        txtSQL.Font.Size = 20
    Else
        txtTamanhoFonte = Range("TamanhoFonte")
        txtSQL.Font.Size = CInt(txtTamanhoFonte)
    End If
    
    With cboFonte
        .AddItem "Consolas"
        .AddItem "Courier New"
        .AddItem "Lucida Sans Typewriter"
        .AddItem "Lucida Console"
        .ListIndex = 0
    End With
    
    If Range("NomeFonte") = Empty Then
        cboFonte.Text = "Consolas"
    Else
        cboFonte.Text = Range("NomeFonte")
    End If
    
    ' Especifica os textos para os botões de comando
    cmdProcurar.Caption = "Procurar" & vbCr & "banco"
    cmdProcurarImagem.Caption = "Procurar" & vbCr & "imagem"
    cmdExecutarSQL.Caption = "Executar" & vbCr & "código SQL"
    cmdLimparSQL.Caption = "Limpar" & vbCr & "código SQL"
    cmdRecuperarSQL.Caption = "Recuperar" & vbCr & "código SQL"
    cmdVisualizarImagem.Caption = "Visualizar" & vbCr & "imagem"
    cmdRestaurarFormulário.Caption = "Restaurar" & vbCr & "formulário"
    lblTítulo2 = Range("Versão")
    
    ' Carrega valores salvos na planilha "Config"
    lblResultado = ""
    chkGravarArquivoTexto = Range("SalvarArquivoTexto")
    chkIncluirTempo = Range("IncluirTempoRegistros")
    chkAjuda = True
    cmdLimparArquivoTexto.Enabled = chkGravarArquivoTexto
    cmdAbrirArquivoTexto.Enabled = chkGravarArquivoTexto
    txtSQL = Range("SQL")
    tglNegrito = Range("Negrito")
    chkAjuda = Range("Ajuda")
    optNãoMudar = Range("NãoMudar")
    optMinúsculas = Range("AlternarMinúsculas")
    optMaiúsculas = Range("AlternarMaiúsculas")
    optIniciaisMaiúsculas = Range("AlternarIniciais")
    chkLimparSQL = Range("InserirCláusulasPadrão")
    chkPontoVírgula = Range("PontoVírgula")
    chkRemoverLinhas = Range("RemoverLinhas")
    
    ' Pré-carrega imagem
    Load frmImagem
    If mpgSQL.Value = 1 Then txtSQL.SetFocus
End Sub

' Mudança de guia no controle multi-página - Se guia "SQL", leva o foco para a caixa de texto
Private Sub mpgSQL_Change()
    If mpgSQL.Value = 1 Then txtSQL.SetFocus
End Sub

' Mouse sobre o controle multi-página - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub mpgSQL_MouseMove(ByVal Index As Long, ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Restaurar_Botões
End Sub

' Botão de comando para os créditos (i) - Exibe os créditos do programa
Private Sub cmdCréditos_Click()
    Dim c As Range, m As String
    For Each c In Range("Créditos")
        m = m & c & vbCrLf
    Next c
    MsgBox m, vbInformation
End Sub

' Botão de comando para os créditos (i) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdCréditos_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para ver detalhes do autor e versão."
    cmdCréditos.BackColor = CinzaMédio
End Sub

' Botão de comando para Ajuda (?) - Exibe uma mensagem com orientações para as ações básicas
Private Sub cmdAjuda_Click()
    Dim c As Range
    Open ThisWorkbook.Path & "\Ajuda.txt" For Output As #1
    For Each c In Range("TextoAjuda")
        Print #1, c.Value & vbCr
    Next c
    Print #1, vbCrLf
    For Each c In Range("Créditos")
        Print #1, c
    Next c
    Close #1
    Shell "notepad.exe " & ThisWorkbook.Path & "\Ajuda.txt", 1
End Sub

' Botão de comando para Ajuda (?) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdAjuda_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para obter uma rápida ajuda sobre as principais ações deste formulário."
    cmdAjuda.BackColor = CinzaMédio
End Sub

' Oculta o formulário
Private Sub cmdFechar_Click()
    lblResultado = ""
    Me.Hide
End Sub

' Botão "Fechar" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdFechar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para fechar esta caixa de diálogo."
    cmdFechar.Font.Bold = True
    cmdFechar.BackColor = CinzaMédio
End Sub

' Formulário (fora de qualquer controle) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Restaurar_Botões
End Sub

' Moldura "Arquivo de dados" (fora de qualquer controle) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub fraArquivoDados_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Restaurar_Botões
End Sub

' Moldura "Arquivo de imagem" (fora de qualquer controle) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub fraArquivoImagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Restaurar_Botões
End Sub

' Moldura "Arquivo de texto" (fora de qualquer controle) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub fraArquivoTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    Call Restaurar_Botões
End Sub

' ================================================================================================================
'                                                Guia "Arquivos"
' ================================================================================================================

' ----------------------------------------------------
' Arquivo de dados
' ----------------------------------------------------

' Rótulo "Nome do arquivo:" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub lblArquivoDados_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Digite o nome do arquivo de dados na caixa ao lado ou clique no botão [Procurar banco]."
End Sub

' Caixa de texto "Arquivo de dados" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub txtArquivoDados_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Digite o nome do arquivo de dados ou clique no botão [Procurar banco]."
End Sub

' Botão "Procurar..." (arquivo de dados) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdProcurar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique aqui para selecionar um arquivo de dados do Access."
    cmdProcurar.Font.Bold = True
    cmdProcurar.BackColor = CinzaMédio
End Sub

' Botão "Procurar" (arquivo de dados) - Permite a seleção de um arquivo de dados do Access
Private Sub cmdProcurar_Click()
    Dim Arq As Variant
    Dim Filtro As String
    Filtro = "Bancos de dados do Access,*.accdb,Bancos de dados do Access 2003,*.mdb"
    ChDrive Left(ThisWorkbook.Path, 1)
    ChDir ThisWorkbook.Path
    Arq = Application.GetOpenFilename(FileFilter:=Filtro, Title:="Abrir banco de dados")
    If Arq = False Then
        txtArquivoDados = Range("ArquivoDados")
    Else
        txtArquivoDados = Arq
        Range("ArquivoDados") = Arq
        lblResultado = ""
    End If
End Sub

' ----------------------------------------------------
' Arquivo de imagem
' ----------------------------------------------------

' Rótulo "Nome do arquivo:" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub lblArquivoImagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Digite o nome do arquivo de imagem na caixa ao lado ou clique no botão [Procurar imagem]."
End Sub

' Caixa de Texto para nome de arquivo de imagem - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub txtArquivoImagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Digite o nome do arquivo de imagem (GIF, JPG ou BMP) ou clique no botão [Procurar imagem]."
End Sub

' Permite a seleção de um arquivo de imagem
Private Sub cmdProcurarImagem_Click()
    Dim Arq As Variant
    Dim Filtro As String
    Filtro = "Arquivos GIF,*.gif,Arquivos JPEG,*.jp*,Arquivos BMP,*.bmp"
    ChDrive Left(Directory(txtArquivoDados), 1)
    ChDir Directory(txtArquivoDados)
    Arq = Application.GetOpenFilename(FileFilter:=Filtro, Title:="Abrir imagem")
    If Arq = False Then
        txtArquivoImagem = Range("ArquivoImagem")
    Else
        txtArquivoImagem = Arq
        Range("ArquivoImagem") = Arq
        lblResultado = ""
    End If
End Sub

' Botão "Procurar..." (imagem) - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdProcurarImagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique aqui para selecionar um arquivo de imagem que representa as relações entre tabelas."
    cmdProcurarImagem.Font.Bold = True
    cmdProcurarImagem.BackColor = CinzaMédio
End Sub

' ----------------------------------------------------
' Arquivo de texto
' ----------------------------------------------------

' Determina se haverá gravação em arquivo de texto
Private Sub chkGravarArquivoTexto_Click()
    If chkGravarArquivoTexto Then
        chkIncluirTempo.Visible = True
        lblNomeTXT.Visible = True
        txtGravarArquivoTexto.Visible = True
        chkIncluirErros.Visible = True
        If Range("ArquivoTexto") = "" Then Range("ArquivoTexto") = "SQL.txt"
        txtGravarArquivoTexto = Range("ArquivoTexto")
        Range("SalvarArquivoTexto") = True
        cmdLimparArquivoTexto.Visible = True
        cmdAbrirArquivoTexto.Visible = True
    Else
        chkIncluirTempo.Visible = False
        txtGravarArquivoTexto.Visible = False
        chkIncluirErros.Visible = False
        lblNomeTXT.Visible = False
        Range("SalvarArquivoTexto") = False
        cmdLimparArquivoTexto.Visible = False
        cmdAbrirArquivoTexto.Visible = False
    End If
End Sub

' Caixa de seleção "[x] Gravar em arquivo de texto" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkGravarArquivoTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Marque esta opção para que cada comando SQL executado seja salvo em arquivo TXT."
End Sub

' Caixa de seleção "Incluir tempo e registros"
Private Sub chkIncluirTempo_Change()
    Range("IncluirTempoRegistros") = chkIncluirTempo
End Sub

' Caixa "Incluir tempo e registros" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkIncluirTempo_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Quando ativado, grava no arquivo de texto informação do tempo de processamento e quantidade de registros."
End Sub

Private Sub lblNomeTXT_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Digite ao lado o nome para o arquivo de texto que registrará cada comando SQL executado."
End Sub

' Caixa de texto com o nome do arquivo TXT - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub txtGravarArquivoTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Digite aqui o nome para o arquivo de texto que registrará cada comando SQL executado."
End Sub

' Caixa de seleção "Armazenar códigos com erros"
Private Sub chkIncluirErros_Change()
    Range("ArmazenarCódigosErrados") = chkIncluirErros
End Sub

' Caixa "Armazenar códigos com erros" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkIncluirErros_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Marque esta opção caso deseje gravar no arquivo TXT códigos SQL que geraram erros."
End Sub

' Limpa o conteúdo do arquivo de texto indicado
Private Sub cmdLimparArquivoTexto_Click()
    Dim Mens As String

    If Dir(ThisWorkbook.Path & "\" & txtGravarArquivoTexto) <> "" Then
        Mens = "Tem certeza que deseja excluir o arquivo " & txtGravarArquivoTexto & "?" & vbCr
        Mens = Mens & "Todos os registros de comandos SQL serão perdidos!"
        If MsgBox(Mens, vbYesNo + vbQuestion + vbDefaultButton2, "Apagar arquivo texto") = vbYes Then
            Kill ThisWorkbook.Path & "\" & txtGravarArquivoTexto
        End If
    Else
        MsgBox "Arquivo de texto não encontrado!" & vbCr & ThisWorkbook.Path & "\" & txtGravarArquivoTexto, vbExclamation
    End If
    Call Restaurar_Botões
End Sub

' Botão "Excluir TXT" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdLimparArquivoTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para excluir o arquivo de texto gerado."
    cmdLimparArquivoTexto.Font.Bold = True
    cmdLimparArquivoTexto.BackColor = CinzaMédio
End Sub

' Abre o arquivo de texto pelo bloco de notas
Private Sub cmdAbrirArquivoTexto_Click()
    If Dir(ThisWorkbook.Path & "\" & txtGravarArquivoTexto) <> "" Then
        Shell "notepad.exe " & ThisWorkbook.Path & "\" & txtGravarArquivoTexto, 1
    Else
        MsgBox "Arquivo de texto não encontrado!" & vbCr & _
            ThisWorkbook.Path & "\" & txtGravarArquivoTexto, vbExclamation
    End If
End Sub

' Botão "Abrir TXT" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdAbrirArquivoTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para abrir o arquivo de texto pelo Bloco de Notas."
    cmdAbrirArquivoTexto.Font.Bold = True
    cmdAbrirArquivoTexto.BackColor = CinzaMédio
End Sub

' ================================================================================================================
'                                                  Guia "SQL"
' ================================================================================================================

' Rótulo "Insira o código SQL..." - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub lblSQL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Na caixa de texto abaixo, digite o código SQL e depois clique no botão [Executar código SQL]."
End Sub

' Caixa de texto principal para digitação de comandos SQL - Ações quando o valor é alterado
Private Sub txtSQL_Change()
    Dim Cláusulas()                                      ' Tabela de cláusulas e funções SQL
    Dim p As Long
    p = txtSQL.SelStart

    Cláusulas() = Range("TabCláusulas")
    
    If optMaiúsculas Then
        p = txtSQL.SelStart
        For i = 1 To UBound(Cláusulas)
            txtSQL = Replace(txtSQL, Cláusulas(i, 1), Cláusulas(i, 2))
        Next i
        txtSQL.SelStart = p
    ElseIf optMinúsculas Then
        p = txtSQL.SelStart
        For i = 1 To UBound(Cláusulas)
            txtSQL = Replace(txtSQL, Cláusulas(i, 1), Cláusulas(i, 1))
        Next i
        txtSQL.SelStart = p
    ElseIf optIniciaisMaiúsculas Then
        p = txtSQL.SelStart
        For i = 1 To UBound(Cláusulas)
            txtSQL = Replace(txtSQL, Cláusulas(i, 1), Cláusulas(i, 3))
        Next i
        txtSQL.SelStart = p
    End If
End Sub

' Caixa de texto principal para digitação de comandos SQL - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub txtSQL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Digite aqui o comando SQL e depois clique em 'Executar código SQL'. É permitido usar múltiplas linhas."
End Sub

' Executa o código SQL inserido na caixa
Private Sub cmdExecutarSQL_Click()
    Dim T As Single                                                     ' Tempo coletado por Timer
    Dim TempoExec As Single                                             ' Tempo total do processamento
    Dim obj As New DataObject                                           ' Objeto de dados para cópia na área de transferência
    Dim Cn As ADODB.Connection                                          ' Objeto para conexão ADO
    Dim Rs As ADODB.Recordset                                           ' Objeto para Recordset ADO
    Dim verErro As Byte
    Erro_Núm = 0                                                        ' Limpa o número do erro
    Erro_Msg = ""                                                       ' Limpa a mensagem do erro
    lblResultado = ""                                                   ' Limpa o rótulo para tempo e quantidade
    
    ' Verifica se existe um arquivo de dados
    If txtArquivoDados = "" Then
        MsgBox "Selecione um arquivo de banco de dados do Access antes de executar a consulta!", vbExclamation
        mpgSQL.Value = 0
        txtArquivoDados.SetFocus
    ' Verifica de foi digitado algum comando SQL
    ElseIf Dir(txtArquivoDados) = "" Then
        MsgBox "Selecione um arquivo de banco de dados do Access antes de executar a consulta!", vbExclamation
        mpgSQL.Value = 0
        txtArquivoDados = ""
        txtArquivoDados.SetFocus
    ElseIf txtSQL = "" Then
        MsgBox "Digite uma instrução SQL válida!", vbExclamation
        txtSQL.SetFocus
    ' Executa o comando SQL digitado
    Else
        Application.ScreenUpdating = False
        If chkRemoverLinhas Then txtSQL = RemoveBlankLines(txtSQL)      ' Remove linhas em branco, se solicitado
        If chkPontoVírgula Then txtSQL = AddSemiColon(txtSQL)           ' Insere ponto-e-vírgula, se solicitado
        Sheets("Dados").Select                                          ' Seleciona a planilha
        Cells.Clear                                                     ' Limpa os dados anteriores
        Cells.ColumnWidth = 8                                           ' Todas as colunas com largura de 8 caracteres
        T = Timer                                                       ' Coleta o tempo antes da execução
        
        ' ===============================================================================================================
        Set Cn = New ADODB.Connection                                   ' Define a conexão com a base de dados
        Cn.Provider = "Microsoft.ace.OLEDB.12.0"                        ' Provedor
        Cn.ConnectionString = "Data Source=" & frmSQL.txtArquivoDados   ' Cadeia de conexão
        Cn.Open                                                         ' Abre a conexão
        SQL = frmSQL.txtSQL                                             ' Define a Instrução SQL
        Set Rs = New ADODB.Recordset                                    ' Define um novo objeto recorset
        On Error GoTo TratarErro                                        ' Ativa o tratamento de erros
        Rs.CursorLocation = adUseClient                                 ' Client-side cursor
        T = Timer                                                       ' Coleta o tempo antes da execução do comando SQL
        Rs.Open SQL, Cn, adOpenStatic                                   ' Abre o recordset como Static para correta contagem
        TempoExec = Timer - T ' ***********
        NúmReg = Rs.RecordCount                                         ' Número total de registros
        Range("A2").CopyFromRecordset Rs                                ' Copia as informações para a planilha
        On Error GoTo 0                                                 ' Desliga o tratamento de erros
        For i = 0 To Rs.Fields.Count - 1                                ' Preenche os cabeçalhos
            Cells(1, i + 1) = Rs(i).Name
        Next i
        Rs.Close                                                        ' Fecha o recordset
        Cn.Close                                                        ' Fecha a conexão
        ' ===============================================================================================================
        
        If Erro_Msg <> "" Then TempoExec = 0 'Else TempoExec = Timer - T ' Calcula o tempo total (se erro, tempo = 0)
        
        If chkGravarArquivoTexto Then                                   ' Se solicitada gravação em arquivo texto
            If Erro_Msg <> "" And chkIncluirErros Then
                Open ThisWorkbook.Path & "\" & txtGravarArquivoTexto For Append As #1
                Print #1, String(40, "-")
                Print #1, txtSQL
                Print #1, ">>> O código SQL acima gerou o erro " & Erro_Núm
                Print #1, ">>> " & Erro_Msg
                Erro_Núm = 0
                Erro_Msg = ""
                Print #1, Chr(13)
                Close #1
            ElseIf chkIncluirTempo And Erro_Msg = "" And Not chkIncluirErros Then
                Open ThisWorkbook.Path & "\" & txtGravarArquivoTexto For Append As #1
                Print #1, String(40, "-")
                Print #1, txtSQL
                Print #1, ">>> Tempo de execução: " & Format(TempoExec, "0.00") & _
                    " s  |  Registros extraídos: " & Format(Range("A1").CurrentRegion.Rows.Count - 1, "#,##0")
                Print #1, Chr(13)
                Close #1
            ElseIf Not chkIncluirErros And Erro_Msg = "" Then
                Open ThisWorkbook.Path & "\" & txtGravarArquivoTexto For Append As #1
                Print #1, String(40, "-")
                Print #1, txtSQL
                Print #1, Chr(13)
                Close #1
            End If
        End If
        
        With Range("A1").CurrentRegion.Rows(1)
            .Font.Bold = True
            .Interior.Color = RGB(240, 240, 240)
        End With
        If Erro_Msg = "" Then
            lblResultado = "Tempo de execução: " & Format(TempoExec, "0.00") & _
                " s   •   Registros extraídos: " & Format(NúmReg, "#,##0")
        Else
            lblResultado = ""
        End If
        
        Range("A1").CurrentRegion.Columns.AutoFit                   ' Aplica auto-ajuste às colunas
        Range("A1").CurrentRegion.Rows.AutoFit                      ' Aplica auto-ajuste às linhas
        
        Range("SQL") = txtSQL
        If chkCopiarTexto Then
            obj.SetText txtSQL.Text
            obj.PutInClipboard
        End If
    End If
    Range("A1").Select
    Set obj = Nothing
    Application.ScreenUpdating = True
    Exit Sub
    
TratarErro:
    Erro_Núm = Err.Number
    Erro_Msg = Err.Description
    verErro = MsgBox(Erro_Msg & vbNewLine & Erro_Núm, vbCritical, "Erro de Execução de Código")
    If verErro = vbOK Then frmSQL.txtSQL = SQL Else Exit Sub
End Sub

' Botão "Executar SQL" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdExecutarSQL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para executar o código SQL digitado na caixa acima. Veja o resultado da consulta diretamente na planilha."
    cmdExecutarSQL.Font.Bold = True
    cmdExecutarSQL.BackColor = CinzaMédio
End Sub

' Limpa o comando SQL digitado
Private Sub cmdLimparSQL_Click()
    If chkLimparSQL Then
        txtSQL = "Select * From " & vbCr & "Where " & vbCr & "Group By " & vbCr & "Having " & vbCr & "Order By "
        Application.SendKeys "^{HOME}{END}"
    Else
        txtSQL = ""
    End If
    lblResultado = ""
    txtSQL.SetFocus
End Sub

' Botão "Limpar código SQL" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdLimparSQL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para limpar a caixa acima. Nada será executado."
    cmdLimparSQL.Font.Bold = True
    cmdLimparSQL.BackColor = CinzaMédio
End Sub

' Recupera o último código SQL executado
Private Sub cmdRecuperarSQL_Click()
    If Range("SQL") <> "" Then txtSQL = Range("SQL")
    lblResultado = ""
    txtSQL.SetFocus
    txtSQL.SelStart = 0
End Sub

' Botão "Recuperar código SQL" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdRecuperarSQL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para recuperar o último comando SQL digitado na caixa acima, caso tenha sido apagado."
    cmdRecuperarSQL.Font.Bold = True
    cmdRecuperarSQL.BackColor = CinzaMédio
End Sub

' Visualiza o arquivo de imagem que representa o relacionamento entre as tabelas
Private Sub cmdVisualizarImagem_Click()
    If txtArquivoImagem = Empty Then
        MsgBox "Nenhum arquivo de imagem associado!", vbExclamation
    ElseIf Dir(txtArquivoImagem) = "" Then
        MsgBox "O arquivo especificado não existe!", vbExclamation
    Else
        frmImagem.imgSQL.Picture = LoadPicture(txtArquivoImagem)
        frmImagem.Width = frmImagem.imgSQL.Width + 12
        frmImagem.Height = frmImagem.imgSQL.Height + 30
        frmImagem.Show
    End If
    Call Restaurar_Botões
    txtSQL.SetFocus
End Sub

' Botão "Visualizar Imagem" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdVisualizarImagem_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para visualizar o arquivo de imagem que representa as relações entre as tabelas."
    cmdVisualizarImagem.Font.Bold = True
    cmdVisualizarImagem.BackColor = CinzaMédio
End Sub

' Rótulo "Fonte" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub lblFonte_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Escolha ao lado uma das fontes de espaçamento constante para usar na caixa de texto do código SQL."
End Sub

' Caixa de combinação "Fonte" - Ações quando há mudança de valor
Private Sub cboFonte_Change()
    txtSQL.Font.Name = cboFonte
    Range("NomeFonte") = cboFonte
End Sub

' Caixa de combinação "Fonte" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cboFonte_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Escolha a fonte de espaçamento constante para ser usada no código SQL."
End Sub

' Rótulo "Tamanho da fonte:" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub lblTamanhoFonte_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique nas setas ao lado para aumentar ou diminuir o tamanho da fonte para o código SQL."
End Sub ' Caixa de texto "tamanho da fonte"

' Caixa de texto "Tamanho da fonte" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub txtTamanhoFonte_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique nas setas ao lado para aumentar ou diminuir o tamanho da fonte para o código SQL."
End Sub

' Aumenta o tamanho da fonte no código SQL
Private Sub spnTamanhoFonte_SpinDown()
    If txtTamanhoFonte > 8 Then
        txtTamanhoFonte = CInt(txtTamanhoFonte) - 1
        txtSQL.Font.Size = CInt(txtTamanhoFonte)
        Range("TamanhoFonte") = CInt(txtTamanhoFonte)
        txtSQL.IntegralHeight = True
        txtSQL.SetFocus
    End If
End Sub

' Diminui o tamanho da fonte no código SQL
Private Sub spnTamanhoFonte_SpinUp()
    If txtTamanhoFonte < 36 Then
        txtTamanhoFonte = CInt(txtTamanhoFonte) + 1
        txtSQL.Font.Size = CInt(txtTamanhoFonte)
        Range("TamanhoFonte") = CInt(txtTamanhoFonte)
        txtSQL.IntegralHeight = True
        txtSQL.SetFocus
    End If
End Sub

' Qualquer alteração, manda o foco de volta à caixa de código SQL
Private Sub spnTamanhoFonte_Change()
    txtSQL.SetFocus
End Sub

' Negrito na caixa de comando SQL
Private Sub tglNegrito_Click()
    txtSQL.Font.Bold = tglNegrito
    Range("Negrito") = tglNegrito
    If mpgSQL.Value = 1 Then txtSQL.SetFocus
End Sub

' Botão "Negrito" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub tglNegrito_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    tglNegrito.Font.Bold = True
    tglNegrito.BackColor = CinzaMédio
    If tglNegrito Then
        lblAjuda = " Clique para desativar o negrito da caixa de código SQL."
    Else
        lblAjuda = " Clique para ativar o negrito da caixa de código SQL."
    End If
End Sub

' ================================================================================================================
'                                                Guia "Opções"
' ================================================================================================================

' Deixa visível a ajuda dos controles
Private Sub chkAjuda_Click()
    Range("Ajuda") = chkAjuda
    If chkAjuda Then frmSQL.Height = 406 Else frmSQL.Height = 376
End Sub

' Caixa "Ativar ajuda dos controles" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkAjuda_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Quando ativado, mostra um texto de ajuda quando o mouse passa sobre o controle."
End Sub

' Caixa "Ao limpar a caixa de texto SQL..." - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkLimparSQL_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Quando ativado, ao clicar no botão [Limpar código SQL], insere a sintaxe padrão como ponto de partida."
End Sub

' Caixa "Ao executar um código SQL, acrescentar ponto-e-vírgula se necessário" - Ações quando o botão é clicado
Private Sub chkPontoVírgula_Click()
    Range("PontoVírgula") = chkPontoVírgula
End Sub

' Caixa "Ao executar um código SQL, acrescentar ponto-e-vírgula se necessário" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkPontoVírgula_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Quando ativado, ao clicar no botão [Executar código SQL], acrescenta o ponto-e-vírgula no final da sintaxe."
End Sub

' Caixa "Ao executar um código SQL, remover linhas em branco" - Ações quando o botão é clicado
Private Sub chkRemoverLinhas_Click()
    Range("RemoverLinhas") = chkRemoverLinhas
End Sub

' Caixa "Ao executar um código SQL, remover linhas em branco" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkRemoverLinhas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Quando ativado, ao clicar no botão [Executar código SQL], remove linhas em branco da sintaxe."
End Sub

' Caixa "Ao executar um código SQL, copiar para a Área de Transferência" - Ações quando o botão é clicado
Private Sub chkCopiarTexto_Click()
    Range("CopiarTexto") = chkCopiarTexto
End Sub

' Caixa "Ao executar um código SQL, copiar para a Área de Transferência" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub chkCopiarTexto_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Quando ativado, ao clicar no botão [Executar código SQL], copia o conteúdo para a Área de Transferência do Windows."
End Sub

' Botão de opção "Não mudar nada"
Private Sub optNãoMudar_Click()
    Range("NãoMudar") = optNãoMudar
    Range("AlternarIniciais") = optIniciaisMaiúsculas
    Range("AlternarMaiúsculas") = optMaiúsculas
    Range("AlternarMinúsculas") = optMinúsculas
    txtSQL_Change
End Sub

' Botão de opção "Alternar Para Iniciais Maiúsculas"
Private Sub optIniciaisMaiúsculas_Click()
    Range("NãoMudar") = optNãoMudar
    Range("AlternarIniciais") = optIniciaisMaiúsculas
    Range("AlternarMaiúsculas") = optMaiúsculas
    Range("AlternarMinúsculas") = optMinúsculas
    txtSQL_Change
End Sub

' Botão de opção "ALTERNAR PARA LETRAS MAIÚSCULAS"
Private Sub optMaiúsculas_Click()
    Range("NãoMudar") = optNãoMudar
    Range("AlternarIniciais") = optIniciaisMaiúsculas
    Range("AlternarMaiúsculas") = optMaiúsculas
    Range("AlternarMinúsculas") = optMinúsculas
    txtSQL_Change
End Sub

' Botão de opção "alternar para letras minúculas"
Private Sub optMinúsculas_Click()
    Range("NãoMudar") = optNãoMudar
    Range("AlternarIniciais") = optIniciaisMaiúsculas
    Range("AlternarMaiúsculas") = optMaiúsculas
    Range("AlternarMinúsculas") = optMinúsculas
    txtSQL_Change
End Sub

' Botão de opção "Alternar Para Iniciais Maiúsculas" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub optIniciaisMaiúsculas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Marque esta opção para que as cláusulas SQL fiquem com iniciais maiúsculas. Exemplos: Select, From, Where..."
End Sub

' Botão de opção "ALTERNAR PARA LETRAS MAIÚSCULAS" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub optMaiúsculas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Marque esta opção para que as cláusulas SQL fiquem somente com letras maiúsculas. Exemplos: SELECT, FROM, WHERE..."
End Sub

' Botão de opção "alternar para letras minúsculas" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub optMinúsculas_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Marque esta opção para que as cláusulas SQL fiquem somente com letras minúsculas. Exemplos: select, from, where..."
End Sub

' Botão de opção "Não mudar nada" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub optNãoMudar_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Marque esta opção para que as cláusulas SQL não sejam alteradas para maiúsculas ou minúsculas durante a digitação."
End Sub

' Botão "Restaurar formulário" - Ações quando o botão é clicado
Private Sub cmdRestaurarFormulário_Click()
    Dim m As String
    m = "Tem certeza que deseja limpar todas as definições do formulário?" & vbCr
    m = m & "O formulário será inicializado com valores padrão."
    If MsgBox(m, vbYesNo + vbQuestion + vbDefaultButton2) = vbYes Then Call LimparDefinições
End Sub

' Botão "Restaurar formulário" - Ações quando o ponteiro do mouse passa sobre o controle
Private Sub cmdRestaurarFormulário_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    lblAjuda = " Clique para limpar todas as definições salvas no formulário e iniciar com valores padrão."
    cmdRestaurarFormulário.Font.Bold = True
    cmdRestaurarFormulário.BackColor = CinzaMédio
End Sub


' ================================================================================================================
'                                                OUTRAS ROTINAS
' ================================================================================================================

' Restaura as cores originais
Private Sub Restaurar_Botões()
    lblAjuda = ""
    tglNegrito.Font.Bold = False
    tglNegrito.BackColor = CinzaClaro
    cmdProcurar.Font.Bold = False
    cmdProcurar.BackColor = CinzaClaro
    cmdProcurarImagem.Font.Bold = False
    cmdProcurarImagem.BackColor = CinzaClaro
    cmdLimparArquivoTexto.Font.Bold = False
    cmdLimparArquivoTexto.BackColor = CinzaClaro
    cmdAbrirArquivoTexto.Font.Bold = False
    cmdAbrirArquivoTexto.BackColor = CinzaClaro
    cmdExecutarSQL.Font.Bold = False
    cmdExecutarSQL.BackColor = CinzaClaro
    cmdLimparSQL.Font.Bold = False
    cmdLimparSQL.BackColor = CinzaClaro
    cmdRecuperarSQL.Font.Bold = False
    cmdRecuperarSQL.BackColor = CinzaClaro
    cmdVisualizarImagem.Font.Bold = False
    cmdVisualizarImagem.BackColor = CinzaClaro
    cmdFechar.Font.Bold = False
    cmdFechar.BackColor = CinzaClaro
    cmdCréditos.BackColor = CinzaClaro
    cmdAjuda.BackColor = CinzaClaro
    cmdRestaurarFormulário.Font.Bold = False
    cmdRestaurarFormulário.BackColor = CinzaClaro
End Sub

' Limpa as definições para o formulário
Private Sub LimparDefinições()
    Range("ArquivoDados").ClearContents
    Range("ArquivoImagem").ClearContents
    Range("NomeFonte") = "Consolas"
    Range("TamanhoFonte") = 20
    Range("Negrito") = False
    Range("Ajuda") = True
    Range("SalvarArquivoTexto") = True
    Range("IncluirTempoRegistros") = True
    Range("ArmazenarCódigosErrados") = False
    Range("InserirCláusulasPadrão") = False
    Range("ArquivoTexto") = "SQL.txt"
    Range("SQL") = ""
    Range("ArmazenarCódigosErrados") = False
    Range("InserirCláusulasPadrão") = False
    Range("RemoverLinhas") = False
    Range("PontoVírgula") = False
    Range("CopiarTexto") = False
    Range("NãoMudar") = True
    Range("AlternarMinúsculas") = False
    Range("AlternarMaiúsculas") = False
    Range("AlternarIniciais") = False
    mpgSQL.Value = 0
    UserForm_Initialize
End Sub

