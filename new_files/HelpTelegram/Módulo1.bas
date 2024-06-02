Attribute VB_Name = "Módulo1"
Option Explicit

Sub Main()
    Dim wsFonte     As Excel.Worksheet
    Dim wsDestino   As Excel.Worksheet
    Dim LastRow     As Long
    Dim LastCol     As Integer
    Dim rng         As Excel.Range
    Dim InitialTime As Single
    
    InitialTime = VBA.Timer

    Set wsFonte = Planilha1
    Set wsDestino = Planilha2
    
    With wsFonte
        LastRow = .Range("A" & wsDestino.Rows.Count).End(xlUp).Row
        LastCol = .Range("XFD1").End(xlToLeft).Column
        Set rng = .Range("A1").Resize(LastRow, LastCol)
    End With
    
    With wsDestino.Range("A1").Resize(rng.Rows.Count, rng.Columns.Count)
        .ClearContents
        .Value2 = rng.Value2
    End With
    
    Debug.Print "Tempo total: " & VBA.Format(VBA.Timer - InitialTime, "0.00 s")
End Sub

