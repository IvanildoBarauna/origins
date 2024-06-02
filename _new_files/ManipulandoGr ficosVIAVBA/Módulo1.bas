Attribute VB_Name = "Módulo1"
Option Explicit

Sub Main()
    Dim ws          As Excel.Worksheet
    Dim FinalRow    As Long
    Dim FinalCol    As Integer
    Dim rngDyn      As Excel.Range
    
    Set ws = Planilha1
    
    With ws
        FinalRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        FinalCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set rngDyn = .Range("A1").Resize(FinalRow, FinalCol)
        On Error Resume Next
        .ChartObjects(1).Delete
        On Error GoTo 0
        .Shapes.AddChart2(Style:=201, _
            XlChartType:=xlColumnClustered).Chart.SetSourceData rngDyn
    End With
    
End Sub
