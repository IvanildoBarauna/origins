Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub ScaleToRange()
    Dim ws      As Worksheet
    Dim lRow    As Long
    Dim lCol    As Integer
    
    Set ws = Planilha1
    With ws
        lRow = .Cells(ws.Rows.Count, "E").End(xlUp).Row
        lCol = .Cells(8, ws.Columns.Count).End(xlToLeft).Column
        .Range(.Cells(lRow + 1, "E"), .Cells(lRow + 1, lCol)).Value2 = .Range("B8").Value2
    End With
End Sub
