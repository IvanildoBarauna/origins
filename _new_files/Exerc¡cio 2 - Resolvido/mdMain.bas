Attribute VB_Name = "mdMain"
Option Explicit

Public Sub DeleteZeros()
    Dim ws          As Worksheet
    Dim LastRow     As Long
    Dim LastCol     As Integer
    Dim rng         As Range
    Dim iCell       As Range
    
    Set ws = wsDados
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    LastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    Set rng = ws.Range("A1").Resize(LastRow, LastCol)
    
    For Each iCell In rng
        If iCell.Value = 0 Then
            iCell.ClearContents
        End If
    Next iCell
    
    MsgBox "Processo concluído", vbInformation
End Sub
