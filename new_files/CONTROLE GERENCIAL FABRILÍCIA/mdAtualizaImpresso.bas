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
Continue:
            shtESTOQUE.ListObjects("tbESTOQUE").Range.AutoFilter Field:=4, Criteria1:= _
              "PRODUTO ACABADO"
              shtESTOQUE.Range("tbESTOQUE[CÓDIGO]" & ":" & "tbESTOQUE[DESCRIÇÃO]").Copy Destination:= _
                  shtPRINT.Range("A2")
             shtESTOQUE.Range("tbESTOQUE[CATEGORIA]").Copy Destination:= _
                  shtPRINT.Range("D2")
             shtESTOQUE.Range("tbESTOQUE[QUANTIDADE]").Copy Destination:= _
                  shtPRINT.Range("E2")
             shtPRINT.PrintOut ActivePrinter:="HP Deskjet 2540 series (Rede)"
             shtHOME.Activate
            MsgBox "Relatório de Estoque atualizado e enviado para a fila de impressão!", vbInformation, "Fabrilícia - Controle de Estoque"
      Else: GoTo Continue
      End If
      Application.ScreenUpdating = 1
End Sub
