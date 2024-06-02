Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub PasteFX()
Dim uRow       As Integer
Dim ws            As Worksheet

Set ws = Planilha1
uRow = ws.Range("O1048576").End(xlUp).Row

     With ws
          .Range("A2:O" & uRow).FormulaLocal = .Range("A2:O" & uRow).Value2
     End With
End Sub
