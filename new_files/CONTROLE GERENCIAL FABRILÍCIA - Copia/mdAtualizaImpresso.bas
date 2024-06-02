Attribute VB_Name = "mdAtualizaImpresso"
Option Explicit

Sub AtualizarImpresso()
      Dim linha         As Integer
      
      Application.ScreenUpdating = 0
      linha = Empty
      linha = shtPRINT.Range("A1048576").End(xlUp).Row
      shtPRINT.Select
      If shtPRINT.Range("A2").Value <> "" Then
            shtPRINT.Range("A2:G2", Range("A2:A" & linha)).Delete
            shtESTOQUE.ListObjects("tbESTOQUE").Range.AutoFilter Field:=4, Criteria1:="PRODUTO ACABADO"
            shtESTOQUE.Range("tbESTOQUE[CÓDIGO]" & ":" & "tbESTOQUE[DESCRIÇÃO]").Copy Destination:=shtPRINT.Range("A2")
            shtESTOQUE.Range("tbESTOQUE[CATEGORIA]").Copy Destination:=shtPRINT.Range("D2")
            shtESTOQUE.Range("tbESTOQUE[QUANTIDADE]").Copy Destination:=shtPRINT.Range("E2")
            shtPRINT.PrintOut
            shtHOME.Activate
            MsgBox "Relatório de Estoque atualizado e enviado para a fila de impressão!", vbInformation, "Fabrilícia - Controle de Estoque"
      End If
      Application.ScreenUpdating = 1
End Sub
