Attribute VB_Name = "Cadastro_Dados"
Public Mes As Variant
Public Ano As Variant
Public Operacao As Variant
Public rs As ADODB.Recordset
Public Connect As ADODB.connection

Function connection()

    Set Connect = New ADODB.connection

    Connect.ConnectionString = "provider=SQLOLEDB.1;Persist Security Info=True;Data Source=10.166.2.21;User ID=wfmRelatorios.user;Password=Cb7cXGAVO;Initial Catalog= wfmRelatorios"
    Connect.Open

End Function

Function desconect()

    Connect.Close
    Set Connect = Nothing
    
End Function

Sub InserirDadosMapa()

    Dim ultima_linha
    Dim SQL As String
    Dim rng As Range

        
        ultima_linha = Plan17.Range("C65000").End(xlUp).Row
        Mes = Month(Plan9.Cells(2, 5).Value)
        Ano = Year(Plan9.Cells(2, 5).Value)
        Operacao = Plan9.Cells(3, 5).Value
        Site = Plan9.Cells(1, 4).Value
        Set rng = Plan17.Range("C3:C" & ultima_linha)
           
        Call connection
        Set rs = New ADODB.Recordset
        
        For Each c In rng
        
        Application.StatusBar = "Aguarde, transferindo dados " & c.Row & " de " & ultima_linha - 1
        
        RE = Plan17.Cells(c.Row, 3).Value
        Lawson = Plan17.Cells(c.Row, 4).Value
        Login = Plan17.Cells(c.Row, 5).Value
        
        If Lawson = "-" Then
        Lawson = 0
        End If
        
        If Login = "-" Then
        Login = 0
        End If
        
        Nome = Plan17.Cells(c.Row, 6).Value
        Leader = Plan17.Cells(c.Row, 7).Value
        Coordinator = Plan17.Cells(c.Row, 8).Value
        Manager = Plan17.Cells(c.Row, 9).Value
        Status = Plan17.Cells(c.Row, 10).Value
        Celula = Plan17.Cells(c.Row, 11).Value
        Projeto = Plan17.Cells(c.Row, 12).Value
        Entrada = Format(Plan17.Cells(c.Row, 13).Value, "hh:mm:ss")
        saida = Format(Plan17.Cells(c.Row, 14).Value, "hh:mm:ss")
        Jornada = Format(Plan17.Cells(c.Row, 15).Value, "hh:mm:ss")
        HireDate = Format(Plan17.Cells(c.Row, 16).Value, "yyyy-mm-dd")
        
        If Status = "Desligado" Then
        DataDesligamento = Format(Plan17.Cells(c.Row, 17).Value, "yyyy-mm-dd")
        Else
        DataDesligamento = ""
        End If
        
        TipoAgente = Plan17.Cells(c.Row, 18).Value
        Email = Plan17.Cells(c.Row, 19).Value
        IEX = Plan17.Cells(c.Row, 20).Value
        Network = Plan17.Cells(c.Row, 21).Value
        
        ID = RE & "_" & Mes & "_" & Ano & "_" & Operacao
        
        If RE <> "-" Then
            
            SQL = "EXEC SP_InserirMapa '" & ID & "','" & Mes & "'," & Ano & ",'" & Operacao & "'," & RE & "," _
                & Lawson & "," & Login & ",'" & Nome & "','" & Leader & "','" & Coordinator & "','" & Manager & "','" _
                & Status & "','" & Celula & "','" & Projeto & "','" & Entrada & "','" & saida & "','" & Jornada & "','" _
                & HireDate & "','" & DataDesligamento & "','" & Email & "','" & TipoAgente & "','" & Site & "','" _
                & IEX & "','" & Network & "'"
                
            rs.Open SQL, Connect
        End If
        Next c
        Call desconect

Application.StatusBar = False
End Sub







