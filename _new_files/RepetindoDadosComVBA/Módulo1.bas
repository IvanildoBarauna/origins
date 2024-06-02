Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub Main()
    Dim ws          As Excel.Worksheet
    Dim lRow        As Long
    Dim iRow        As Long
    Dim SumCounter  As Integer
    Dim InitialTime As Single
    
    InitialTime = VBA.Timer
    Set ws = Planilha1
    lRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    For iRow = 1 To lRow
        SumCounter = SumCounter + ws.Cells(iRow, 1).Value2
    Next iRow
    
    Debug.Print "O valor total é: " & SumCounter
    Debug.Print "O tempo total foi de: " & VBA.Format(VBA.Timer - InitialTime, "0.000 segundos")
End Sub

Public Sub RepeatData(ByRef rng As Excel.Range, ByVal NumRepetitions)
    
    If NumRepetitions < 2 Then
        MsgBox "O número de repetições deve ser maior que 1!", vbExclamation
        Exit Sub
    Else
        rng.Resize(rng.Rows.Count, _
            rng.Columns.Count).Value2 = rng.Value
    End If
End Sub
