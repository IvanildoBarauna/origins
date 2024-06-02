Attribute VB_Name = "mdLOG"
Sub CRIARLOG(Ação As String)

Dim ws As Worksheet
    Set ws = shtLOG  'Aponta a variável para a Objeto Worksheet que deseja
    
Dim ÚltimaLinha As Long
    ÚltimaLinha = ws.Range("A1048576").End(xlUp).Row + 1 'Verifica qual a última linha preenchida e adiciona mais uma linha
    
Dim MomentoInteração As Date
     MomentoInteração = Now  'Configura a variável para receber a data e hora atual
     
Dim ComputerName As String
  ComputerName = Environ("COMPUTERNAME") 'Usado comando environ com expressão COMPUTERNAME recebendo o nome do computador
  
Dim UserName As String
    UserName = Environ("USERNAME") ' Usado comando environ com expressão USERNAME recebendo o nome de usuário logado

'Usado coleção cells apontando a linha para a variável UltimaLinha e número da coluna desejada recebendo cada uma das variáveis

With ws
        .Cells(ÚltimaLinha, 1) = MomentoInteração
        .Cells(ÚltimaLinha, 2) = UserName
        .Cells(ÚltimaLinha, 3) = ComputerName
        .Cells(ÚltimaLinha, 4) = Ação
End With

ws.Columns.EntireColumn.AutoFit 'Autoajusta as colunas

End Sub
