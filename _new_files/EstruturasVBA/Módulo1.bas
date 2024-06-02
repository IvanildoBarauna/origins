Attribute VB_Name = "Módulo1"
Option Explicit
Sub DoWhileInverso()
' Faça ... Enquanto inverso
      Dim resultado           As Long
      Dim ws                      As Worksheet
      Dim ultCel                 As Range
      
      resultado = Empty
      
      Set ws = sht
      Set ultCel = ws.Range("A1048576").End(xlUp)
      
      ultCel.Select
      
      Do While ActiveCell.Row >= 2
            If ActiveCell.Row = 2 Then
                 ActiveCell.Offset(0, 1).Value = "Resultado final: " & resultado + ActiveCell.Value
            Else
                  resultado = ActiveCell.Value + resultado
                  ActiveCell.Offset(0, 1).Value = "Acumulado em: " & resultado
                  ActiveCell.Offset(-1, 0).Select
            End If
      Loop
      ultCel.Offset(0, 1).Value = "o resultado é: " & resultado
      resultado = Empty
End Sub

Sub DoWhileComum()
' Faça ... Enquanto
      Dim resultado            As Long
      Dim ws                       As Worksheet
      
      Set ws = sht
      ws.Range("B1").Select
      
      Do While ActiveCell.Value <> ""
            resultado = ActiveCell.Value + resultado
            ActiveCell.Offset(1, 0).Select
      Loop
      ActiveCell.Offset(-1, 1).Value = "o resultado é: " & resultado
      resultado = Empty
End Sub

Sub DoUntil()
' Faça .... Até ... Limitador
      Dim resultado As Long
      Dim ws            As Worksheet
      
      Set ws = sht
      ws.Range("b1").Select
      
      Do Until resultado >= 4000
            resultado = ActiveCell.Value + resultado
            ActiveCell.Offset(0, 1).Value = "Acumulado em: " & resultado
            ActiveCell.Offset(1, 0).Select
      Loop
      ActiveCell.Offset(-1, 1).Value = "O resultado é: " & resultado
      resultado = Empty
End Sub

Sub DoWhileComIF()
      Dim i As Long
      Dim w As Worksheet
      Dim iNeg As Long
      
      i = Empty
      Set w = sht
      
      w.Range("A2").Select
      
      Do While ActiveCell.Value <> ""
            If ActiveCell.Value > 0 Then
                  i = ActiveCell.Value + i
            Else
                  iNeg = ActiveCell.Value + iNeg
            End If
            ActiveCell.Offset(1, 0).Select
      Loop
      MsgBox "Soma dos números positivos é: " & i _
            & Chr(13) & "Soma dos números negativos é: " & iNeg, vbInformation
End Sub

Sub ForNext()
      Dim i       As Long
      Dim w     As Worksheet
      Dim iTime As Date
      Dim fTime As Date
      Dim uLinha As Long
      
      Application.ScreenUpdating = False
      
      Set w = sht
      w.Range("A2").Select
      uLinha = w.Range("A1048576").End(xlUp).Row - 1
       iTime = Now
       
       w.Range("A2:A" & uLinha).ClearContents
      
      For i = 1 To w.Range("cont").Value
            ActiveCell.Value = i
            ActiveCell.Offset(1, 0).Select
      Next
      
      fTime = Now
      Range("cont").Offset(0, 2) = "Tempo total: " & Format(fTime - iTime, "hh:mm:ss")
      Application.ScreenUpdating = True
      MsgBox Range("cont").Value & " linhas foram preenchidas", vbInformation
End Sub

Sub ForNextStep()
      Dim i As Long
      Dim w As Worksheet
      Dim mx As Long
      Dim uLinha As Long
      
      Application.ScreenUpdating = 0
      
      Set w = sht
      mx = w.Range("cont").Value
      uLinha = w.Range("A1048576").End(xlUp).Row + 1
      w.Range("A2:A" & uLinha).ClearContents
      w.Range("A2").Select
      
      For i = mx To 1 Step -1
            ActiveCell.Value = i
            ActiveCell.Offset(1, 0).Select
      Next
      Application.ScreenUpdating = 1
End Sub
