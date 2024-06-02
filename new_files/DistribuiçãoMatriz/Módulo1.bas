Attribute VB_Name = "Módulo1"
Function ArrayCustomizedDist(QtdeItems As Integer, Cols As Integer, Rows As Long) As Variant
    Dim arrAux, arrInv, item: item = 1
    Dim ws          As Worksheet: Set ws = Planilha1
    Dim fDim        As Integer
    Dim iRow        As Integer
    Dim iCol        As Integer
    Dim iRowInv     As Integer
    Dim iColInv     As Integer
    
    ReDim arrAux(1 To Rows, 1 To Cols)
    ReDim arrInv(1 To Rows, 1 To Cols)
    
    For fDim = 1 To Rows
        For sDim = 1 To Rows
            If fDim = 1 Then
                arrAux(fDim, sDim) = item
                item = item + 1
            ElseIf sDim = 1 Then
                arrAux(fDim, sDim) = arrAux(fDim - 1, Cols) + 1
            Else
                arrAux(fDim, sDim) = arrAux(fDim, sDim - 1) + 1
            End If
        Next sDim
    Next fDim
    
    With ws
        .Range("A1").Resize(Rows, Cols).Value2 = arrAux
        .Range("A8").Resize(Rows, Cols).Value2 = InvertedArray(arrAux)
    End With
    
End Function

Public Function InvertedArray(SourceArray As Variant) As Variant
    Dim arrInv
    Dim iRow        As Integer
    Dim iCol        As Integer
    Dim iRowInv     As Integer
    Dim iColInv     As Integer
    Dim rngDestino  As Range
    
    If rngDestino Is Nothing Then Exit Sub
    
    ReDim arrInv(1 To UBound(SourceArray, 1), 1 To UBound(SourceArray, 2))
    
    iRowInv = UBound(SourceArray, 1)
    
    For iRow = 1 To UBound(SourceArray, 1)
        iColInv = UBound(SourceArray, 2)
        For iCol = 1 To UBound(SourceArray, 1)
            arrInv(iRow, iCol) = SourceArray(iRowInv, iColInv)
            iColInv = iColInv - 1
        Next iCol
        iRowInv = iRowInv - 1
    Next iRow
    
    InvertedArray = arrInv
End Function
