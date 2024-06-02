Attribute VB_Name = "mdAux"
Option Explicit

Public Sub CallForm()
    frmMain.Show False
End Sub
Public Function ArrtoListBox() As Variant
    Dim ws      As Worksheet
    Dim iRow    As Long
    Dim iCol    As Integer
    Dim lRow    As Long
    Dim lCol    As Integer
    Dim arr     As Variant
    
    Set ws = Planilha1
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ReDim arr(1 To lRow, 1 To lCol)
    
    For iRow = LBound(arr, 1) To UBound(arr, 1)
        For iCol = LBound(arr, 2) To UBound(arr, 2)
            arr(iRow, iCol) = ws.Cells(iRow, iCol).Value
        Next iCol
    Next iRow
    
    ArrtoListBox = arr
    Erase arr
End Function

Public Function RangeToListBox() As Excel.Range
    Dim ws      As Worksheet
    Dim iRow    As Long
    Dim iCol    As Integer
    Dim lRow    As Long
    Dim lCol    As Integer
    Dim arr     As Variant
    
    Set ws = Planilha1
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ReDim arr(1 To lRow, 1 To lCol)
    
    For iRow = LBound(arr, 1) To UBound(arr, 1)
        For iCol = LBound(arr, 2) To UBound(arr, 2)
            arr(iRow, iCol) = ws.Cells(iRow, iCol).Value
        Next iCol
    Next iRow
    
    With ws
        Set RangeToListBox = .Range(.Cells(LBound(arr, 1), LBound(arr, 2)), _
                             .Cells(UBound(arr, 1), UBound(arr, 2)))
    End With
        
    Erase arr
End Function

