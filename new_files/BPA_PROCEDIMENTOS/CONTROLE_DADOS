Attribute VB_Name = "CONTROLE_DADOS"
Option Explicit
Public Const NameApp As String = "Gerencial BPA"
Public Sub CriarlstUnica()
Application.ScreenUpdating = False
Dim wsFonte                   As Worksheet
Dim wsDestino               As Worksheet
Dim uRowFonte              As Integer
Dim uRowDestinoA           As Integer
Dim uRowDestinoB              As Integer

On Error Resume Next
     
     'Seta as Planilhas fonte e Destino
     Set wsFonte = shtDIGITAÇÃO
     Set wsDestino = shApoio
     
     If wsFonte.AutoFilterMode Then wsFonte.ListObjects("tbDIGITAÇÃO").AutoFilter.ShowAllData
     'Declara que a ultima linha da planilha fonte
     uRowFonte = wsFonte.Range("A1048576").End(xlUp).Row
     ' Declara a última linha da coluna A da planilha destino
     uRowDestinoA = wsDestino.Range("A1048576").End(xlUp).Row
     'Declara a última linha da coluna B da plnailha destino
     uRowDestinoB = wsDestino.Range("B1048576").End(xlUp).Row
     'Exclui os dados da coluna A com a variavel DESTINOA
     wsDestino.Range("A2:A" & uRowDestinoA).ClearContents
     'Exclui os dados da coluna A com a variavel DESTINOB
     wsDestino.Range("B2:B" & uRowDestinoB).ClearContents
     'Valida se há dados na planilha fonte
     If uRowFonte < 10 Then Exit Sub
     'Copia a coluna A da planilha FONTE e cola na coluna A da plan destino
     wsFonte.Range("A8:A" & uRowFonte).Copy Destination:=wsDestino.Range("A1")
     'Copia a coluna B da planilha FONTE e cola na coluna B da plan destino
     wsFonte.Range("B8:B" & uRowFonte).Copy Destination:=wsDestino.Range("B1")
     
     'Re-declara as variaveis de ultima linha da plan destino pois novos dados foram inseridos
     uRowDestinoA = wsDestino.Range("A1048576").End(xlUp).Row
     uRowDestinoB = wsDestino.Range("B1048576").End(xlUp).Row
     
     'Remove duplicadas coluna A
     wsDestino.Range("A1:A" & uRowDestinoA).RemoveDuplicates Columns:=1, Header:=xlYes
     'Remove duplicadas coluna B
     wsDestino.Range("B1:B" & uRowDestinoB).RemoveDuplicates Columns:=1, Header:=xlYes
Application.ScreenUpdating = True
End Sub

Public Sub InserirProfissional()
Application.ScreenUpdating = False
     Dim ws As Worksheet
     Dim uRow  As Integer
     Dim strProf As String
     Dim strCBO As String
     
     Set ws = shtPROF
     uRow = ws.Range("A1048576").End(xlUp).Row
     strProf = InputBox("Digite o nome do profissional:", NameApp)
     strCBO = InputBox("Informe o CBO do profissional: " & strProf, NameApp)
     
     If strProf = vbNullString Then
          MsgBox "Profissional invállido", vbCritical, NameApp
     ElseIf strCBO = vbNullString Then
          MsgBox "CBO inválido", vbCritical, NameApp
     Else
          ws.Range("A" & uRow + 1).Value = UCase(strProf)
          ws.Range("A" & uRow + 1).Offset(0, 1).Value = strCBO
          MsgBox "Profissional: " & strProf & ", CBO: " & strCBO & " adicionado com sucesso!" _
               , vbInformation, NameApp
     End If
Application.ScreenUpdating = True
End Sub

Public Sub InserirProcedimento()
Application.ScreenUpdating = False
     Dim ws As Worksheet
     Dim uRow  As Integer
     Dim strProced As String
     Dim strNproced As String
     
     Set ws = shtPROCED
     uRow = ws.Range("A1048576").End(xlUp).Row
     strProced = InputBox("Digite o nome do procedimento:", NameApp)
     strNproced = InputBox("Informe o número do procedimento: " & strProced, NameApp)
     
     If strProced = vbNullString Then
          MsgBox "Profissional invállido", vbCritical, NameApp
     ElseIf strNproced = vbNullString Then
          MsgBox "CBO inválido", vbCritical, NameApp
     Else
          ws.Range("A" & uRow + 1).Value = UCase(strProced)
          ws.Range("A" & uRow + 1).Offset(0, 1).Value = strNproced
          MsgBox "Procedimento: " & strProced & ", Número de Procedimento: " & strNproced & " adicionado com sucesso!" _
               , vbInformation, NameApp
     End If
Application.ScreenUpdating = True
End Sub

Public Sub PrintSheet()
Dim iVal As Integer
Dim ws As Worksheet

Application.ScreenUpdating = False

iVal = 0
Set ws = shtCONS

iVal = MsgBox("O relatório será impresso com os dados abaixo: " & vbNewLine _
     & vbNewLine & "ANO: " & ws.Range("ANO").Value2 & vbNewLine & "MÊS: " _
          & UCase(ws.Range("MÊS").Value2) & vbNewLine & vbNewLine & "Deseja prosseguir?", vbQuestion + vbYesNo)
     
     If iVal = vbYes Then
          ws.PrintOut
          MsgBox "Relatório enviado para a fila de impressão", vbInformation, NameApp
     Else: MsgBox "Selecione o ano e mês desejado", vbExclamation
     ws.Range("ANO").Select
     End If
     Application.ScreenUpdating = True
End Sub
