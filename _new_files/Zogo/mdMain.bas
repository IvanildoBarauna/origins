Attribute VB_Name = "mdMain"
Public Enum FunctionReturn
    game = 1
    Location = 2
    Category = 3
End Enum

Public Sub InsertNewGame()
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim lr As ListRow
    
    Set ws = shJogos
    Set lo = ws.ListObjects(1)
    Set lr = lo.ListRows.Add
    
    lr.Application.Range(lo.DataBodyRange.Cells(lr.Index, 1), _
        lo.DataBodyRange.Cells(lr.Index, 4)).Value = ArrToListRow
    lo.DataBodyRange.Cells(lr.Index, 1).Select
End Sub

Private Function ArrToListRow() As Variant
    Dim tmpArr(1 To 1, 1 To 4)  As Variant
    
    tmpArr(1, 1) = VBA.DateTime.Date
    tmpArr(1, 2) = RandomValue(game)
    tmpArr(1, 3) = RandomValue(Location)
    tmpArr(1, 4) = RandomValue(Category)
    
    ArrToListRow = tmpArr
    Erase tmpArr
End Function

Public Function RandomValue(ByVal ReturnType As FunctionReturn) As String
    Dim ws         As Worksheet
    Dim RandomLine As Integer
    Dim lRow       As Integer
    
    Set ws = shListas
    lRow = ws.Cells(ws.Rows.Count, ReturnType).End(xlUp).Row
    
    RandomLine = Application.WorksheetFunction.RandBetween(2, lRow)
    RandomValue = ws.Cells(RandomLine, ReturnType).Value2
End Function
