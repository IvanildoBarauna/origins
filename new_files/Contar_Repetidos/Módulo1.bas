Attribute VB_Name = "Módulo1"
Option Explicit
Public Sub Main()
    Dim ws          As Worksheet
    Dim lRow        As Long
    Dim lCol        As Integer
    Dim iRow        As Long
    Dim iCol        As Integer
    Dim oDic        As Object
    Dim NewItem     As String
    Dim arr         As Variant
    Dim Key         As String
    Dim wFunction   As WorksheetFunction
    
    Set wFunction = Application.WorksheetFunction
    Set oDic = CreateObject("Scripting.Dictionary")
    Set ws = Planilha1
    
    With ws
        lRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        lCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        oDic.CompareMode = TextCompare
        
        For iRow = 1 To lRow
            For iCol = 1 To lCol
                NewItem = .Cells(iRow, iCol).Value2
                If Not oDic.Exists(NewItem) Then oDic.Add NewItem, NewItem
            Next iCol
        Next iRow
           
        If .Range("H3").Value = "" Or .Range("I3").Value2 = "" Then GoTo Continue

        .Range("H3", .Range("H3").End(xlDown)).ClearContents
        .Range("I3", .Range("I3").End(xlDown)).ClearContents
        
Continue:
        arr = oDic.Items
        
        For iRow = LBound(arr) To UBound(arr)
            .Cells(iRow + 3, "H").Value2 = arr(iRow)
            .Cells(iRow + 3, "I").Value2 = _
                    wFunction.CountIf(.Range("A1", .Cells(lRow, lCol)), arr(iRow))
        Next iRow
    End With
End Sub


