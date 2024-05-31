Attribute VB_Name = "mdTestes"
Option Explicit

Sub SalvarDados_Tabela(ws As Worksheet, FRM As MSForms.UserForm, strTAG As String)
Dim xControl As MSForms.Control
Dim lo As ListObject
Dim lr As ListRow

Set lo = ws.ListObjects("Tabela1")
Set lr = lo.ListRows.Add

     For Each xControl In FRM.Controls
          If xControl.Tag = strTAG Then
               lo.Range(lr.Index + 1, lo.ListColumns(xControl.Name).Index).Value = xControl.Value
          End If
     Next xControl

End Sub

Sub SalvarDados_SemTabela(ws As Worksheet, FRM As MSForms.UserForm, strTAG As String)
Dim xControl As MSForms.Control
Dim sColumAD As String
Dim LastRow As Integer

LastRow = ws.Range("A1048576").End(xlUp).Row + 1

     For Each xControl In FRM.Controls
          If xControl.Tag = strTAG Then
               sColumAD = VBA.Replace(VBA.Left(ws.Range(xControl.Name).Address, 2), "$", "")
               ws.Range(sColumAD & LastRow).Value2 = xControl.Value
          End If
     Next xControl
     
End Sub

Sub SalvarDados_MétodoComum_SemTabela()
Dim ws As Worksheet
Dim LastRow As Integer

Set ws = shClientes
LastRow = ws.Range("A1048576").End(xlUp).Row + 1
     
     With ws
          .Cells(LastRow, 1).Value2 = frmCadClientes.ClienteID.Value
          .Cells(LastRow, 2).Value2 = frmCadClientes.RazaoSocial.Value
          .Cells(LastRow, 3).Value2 = frmCadClientes.Telefone.Value
          .Cells(LastRow, 4).Value2 = frmCadClientes.Endereco.Value
          .Cells(LastRow, 5).Value2 = frmCadClientes.Cidade.Value
          .Cells(LastRow, 6).Value2 = frmCadClientes.Pais.Value
     End With
End Sub

Sub SalvarDados_MétodoComum_ComTabela()
Dim ws As Worksheet
Dim lo As ListObject
Dim lr As ListRow

Set ws = shClientes
Set lo = ws.ListObjects("Tabela1")
Set lr = lo.ListRows.Add

     With lo
          .Range(lr.Index + 1, 1).Value = frmCadClientes.ClienteID.Value
          .Range(lr.Index + 1, 2).Value = frmCadClientes.RazaoSocial.Value
          .Range(lr.Index + 1, 3).Value = frmCadClientes.Telefone.Value
          .Range(lr.Index + 1, 4).Value = frmCadClientes.Endereco.Value
          .Range(lr.Index + 1, 5).Value = frmCadClientes.Cidade.Value
          .Range(lr.Index + 1, 6).Value = frmCadClientes.Pais.Value
     End With
End Sub

Sub Medição()
Dim i As Integer
Dim sTime As Single
Dim eTime As Single
Dim rTime As String
Dim intLimit As Integer
intLimit = InputBox("Digite a quantidade de Loops:", "Medição de Rotinas")
sTime = VBA.Timer()
     For i = 1 To intLimit
          With frmCadClientes
               .Show
               .ClienteID = i
               .RazaoSocial = "NOME DA EMPRESA " & i
               .Telefone = "TELEFONE DO CLIENTE " & i
               .Endereco = "ENDEREÇO DO CLIENTE " & i
               .Cidade = "CIDADE DO CLIENTE " & i
               .Pais = "PAIS DO CLIENTE " & i
          End With
          Call SalvarDados_Tabela(shClientes, frmCadClientes, "cad_clientes")
     Next i
Unload frmCadClientes
shClientes.Cells.Columns.AutoFit
eTime = VBA.Timer()
rTime = Format(eTime - sTime, "0 segundos")
Debug.Print "Método Loop ForEach, Tempo com Tabela: " & rTime

sTime = VBA.Timer()
          For i = 1 To intLimit
          With frmCadClientes
               .Show
               .ClienteID = i
               .RazaoSocial = "NOME DA EMPRESA " & i
               .Telefone = "TELEFONE DO CLIENTE " & i
               .Endereco = "ENDEREÇO DO CLIENTE " & i
               .Cidade = "CIDADE DO CLIENTE " & i
               .Pais = "PAIS DO CLIENTE " & i
          End With
          Call SalvarDados_SemTabela(shClientes, frmCadClientes, "cad_clientes")
     Next i
     Unload frmCadClientes
     shClientes.Cells.Columns.AutoFit
eTime = VBA.Timer()
rTime = Format(eTime - sTime, "0 segundos")
Debug.Print "Método Loop ForEach, Tempo Sem Tabela: " & rTime

sTime = VBA.Timer()
          For i = 1 To intLimit
          With frmCadClientes
               .Show
               .ClienteID = i
               .RazaoSocial = "NOME DA EMPRESA " & i
               .Telefone = "TELEFONE DO CLIENTE " & i
               .Endereco = "ENDEREÇO DO CLIENTE " & i
               .Cidade = "CIDADE DO CLIENTE " & i
               .Pais = "PAIS DO CLIENTE " & i
          End With
          Call SalvarDados_MétodoComum_SemTabela
     Next i
     Unload frmCadClientes
     shClientes.Cells.Columns.AutoFit
eTime = VBA.Timer()
rTime = Format(eTime - sTime, "0 segundos")
Debug.Print "Método Comum, Tempo Sem Tabela: " & rTime

sTime = VBA.Timer()
          For i = 1 To intLimit
          With frmCadClientes
               .Show
               .ClienteID = i
               .RazaoSocial = "NOME DA EMPRESA " & i
               .Telefone = "TELEFONE DO CLIENTE " & i
               .Endereco = "ENDEREÇO DO CLIENTE " & i
               .Cidade = "CIDADE DO CLIENTE " & i
               .Pais = "PAIS DO CLIENTE " & i
          End With
          Call SalvarDados_MétodoComum_ComTabela
     Next i
     Unload frmCadClientes
     shClientes.Cells.Columns.AutoFit
eTime = VBA.Timer()
rTime = Format(eTime - sTime, "0 segundos")
Debug.Print "Método Comum, Tempo Com Tabela: " & rTime
End Sub


Public Sub ValidaCampos(FRM As MSForms.UserForm)
    Dim xControl As MSForms.Control

    For Each xControl In FRM.Controls
          If xControl.Tag = "cad_clientes" Then
                If xControl.Value = vbNullString Then MsgBox "Todos os campos são obrigatórios", vbExclamation
                Exit For
          End If
    Next xControl
End Sub

