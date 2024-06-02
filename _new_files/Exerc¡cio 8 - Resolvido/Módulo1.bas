Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub CalculoDeValorFuturo()
    Dim intUltimaLinha As Integer
    Dim intResposta    As Integer
    Dim dblVP          As Double
    Dim intN           As Integer
    Dim dbli           As Double
    Dim dblVF          As Double
    Dim ws             As Worksheet
    Dim iRow           As Long
    
    intResposta = MsgBox("Deseja executar a macro?", vbQuestion + vbYesNo)
    
    If intResposta = 6 Then
        Set ws = wsCalculos
        intUltimaLinha = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        For iRow = 2 To intUltimaLinha
            dblVP = ws.Cells(iRow, 1).Value2
            dbli = ws.Cells(iRow, 3).Value2
            intN = ws.Cells(iRow, 2).Value2
            dblVF = dblVP * (1 + dbli) ^ intN
            ws.Cells(iRow, 4).Value2 = dblVF
        Next iRow
        MsgBox "Macro executada com sucesso", vbInformation
    End If
End Sub
