Attribute VB_Name = "mdFilterArrays"
Public Sub FiltrarArrayOneDimension()
    Dim MyArray As Variant
    Dim vRow    As Double
    Dim vCol    As Double
    Dim ws      As Worksheet
    Dim uRow    As Long
    Dim uCol    As Long
    Dim arr     As Variant
    
    Set ws = shtArrays
    uRow = ws.ListObjects("tbMatriz").ListRows.Count
    
    ReDim MyArray(1 To uRow)
    
    For vRow = 1 To UBound(MyArray)
        MyArray(vRow) = ws.ListObjects("tbMatriz").Range(vRow, 1).Value2
    Next vRow
    
    arr = VBA.Filter(MyArray, "SP", True, vbTextCompare)
    
    With wsEX
        .Range("A1", .Cells(UBound(MyArray), 1)).Value2 = arr
    End With
End Sub

Public Sub FiltrarArrayBIDimension()
    Dim MyArray As Variant
    Dim vRow    As Double
    Dim vCol    As Double
    Dim ws      As Worksheet
    Dim uRow    As Long
    Dim uCol    As Long
    Dim Arr2    As Variant
    
    Set ws = shtArrays
    uRow = ws.ListObjects("tbMatriz").ListRows.Count
    uCol = ws.ListObjects("tbMatriz").ListColumns.Count
    
    ReDim MyArray(1 To uRow, 1 To uCol)
    
    For vRow = 1 To UBound(MyArray)
        For vCol = 1 To UBound(MyArray, 2)
            MyArray(vRow, vCol) = ws.ListObjects("tbMatriz").Range(vRow, vCol).Value2
        Next vCol
    Next vRow
    
    Arr2 = Filter2DArray(mtz:=MyArray, nCol:=2, strFind:="SP", Header:=True)
    
    With wsEX
        .Range("A1", .Cells(UBound(Arr2, 1), UBound(Arr2, 2))).Value2 = Arr2
    End With
    
    Call LIMPAR_MEMORIA(MyArray, ws, Arr2)
End Sub

Function Filter2DArray(ByVal mtz As Variant, _
                       ByVal nCol As Long, _
                       ByVal strFind As String, _
                       ByVal Header As Boolean)
                       
On Error Resume Next
Dim tmpArr  As Variant
Dim i       As Long
Dim j       As Long
Dim arr     As Variant
Dim DIC     As Variant
Dim TmpStr  As String
Dim Tmp     As Variant
Dim Chk     As Boolean
Dim TmpVal  As Double
  
Set DIC = CreateObject("Scripting.Dictionary")
tmpArr = mtz
nCol = nCol + LBound(tmpArr, 2) - 1
Chk = (InStr("><=", Left(strFind, 1)) > 0)
  
  For i = LBound(tmpArr, 1) - Header To UBound(tmpArr, 1)
  
    If Chk Then
      TmpVal = CDbl(tmpArr(i, nCol))
      If Evaluate(TmpVal & strFind) Then DIC.Add i, ""
    Else
      If UCase(tmpArr(i, nCol)) Like UCase(strFind) Then DIC.Add i, ""
    End If
    
  Next i
  
  If DIC.Count > 0 Then
  
    Tmp = DIC.Keys
    ReDim arr(LBound(tmpArr, 1) To UBound(Tmp) + LBound(tmpArr, 1) _
                        - Header, LBound(tmpArr, 2) To UBound(tmpArr, 2))
                        
    For i = LBound(tmpArr, 1) - Header To UBound(Tmp) + LBound(tmpArr, 1) - Header
      For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
        arr(i, j) = tmpArr(Tmp(i - LBound(tmpArr, 1) + Header), j)
      Next j
    Next i
    
    If Header Then
    
      For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
        arr(LBound(tmpArr, 1), j) = tmpArr(LBound(tmpArr, 1), j)
      Next j
      
    End If
    
  End If
  
  Filter2DArray = arr
  
End Function
