Attribute VB_Name = "mdMain"
Option Explicit

Public Sub GeraTabuada()
    Dim ws          As Worksheet: Set ws = wsTabuadas
    Dim Number      As Long
    Dim iRow        As Integer
    
    On Error GoTo err
    Number = VBA.InputBox("Digite um número desejado para geração da tabuada:", "Gerador de Tabuada")

    If ws.Cells(1, 1).Value2 <> "" Then ws.Cells.Delete
    
    For iRow = 1 To 10
        With ws
            .Cells(iRow, 1).Value2 = Number
            .Cells(iRow, 2).Value2 = "X"
            .Cells(iRow, 3).Value2 = iRow
            .Cells(iRow, 4).Value2 = Number * iRow
        End With
    Next iRow
    
    MsgBox "Processo Concluído", vbInformation
    Exit Sub
err:
    MsgBox "Digite um valor válido.", vbExclamation
End Sub
