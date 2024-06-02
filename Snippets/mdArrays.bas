## AiosaPlan.xslm
Attribute VB_Name = "mdArrays"
Option Explicit
Public Sub FiltrarArrayBIDimension()
    Dim MyArray     As Variant
    Dim vRow        As Double
    Dim vCol        As Double
    Dim ws          As Worksheet
    Dim lo          As ListObject
    Dim nRows       As Long
    Dim nCols       As Long
    Dim LastRow     As Long
    Dim LastColumn  As Long
    Dim Arr2        As Variant
    Dim sFilter     As String
    
    sFilter = VBA.InputBox("Digite a UF do estado que deseja filtrar:")
    If sFilter = vbNullString Then
        MsgBox "UF não informada.", vbExclamation
        Exit Sub
    End If

    Set ws = shFilter
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    LastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    ws.Range("A1", ws.Cells(LastRow, LastColumn)).Delete Shift:=xlToLeft
    
    Set lo = shVendas.ListObjects("tbVendas")
    nRows = lo.ListRows.Count
    nCols = lo.ListColumns.Count
    
    ReDim MyArray(1 To nRows, 1 To nCols)
    
    For vRow = 1 To UBound(MyArray)
        For vCol = 1 To UBound(MyArray, 2)
            MyArray(vRow, vCol) = lo.Range(vRow, vCol).Value2
        Next vCol
    Next vRow
    
    Arr2 = Filter2DArray(MyArray, 2, sFilter, True)

    With shFilter
        .Range("A1", .Cells(UBound(Arr2, 1), UBound(Arr2, 2))).Value2 = Arr2
    End With
    
    Erase MyArray
    Erase Arr2
    
    ws.Select
    MsgBox "Concluído", vbInformation
End Sub


Function Filter2DArray(ByVal sArray, ByVal ColIndex As Long, ByVal FindStr As String, ByVal HasTitle As Boolean)
  Dim tmpArr, i As Long, j As Long, arr, Dic, TmpStr, Tmp, Chk As Boolean, TmpVal As Double
  On Error Resume Next
  Set Dic = CreateObject("Scripting.Dictionary")
  tmpArr = sArray
  ColIndex = ColIndex + LBound(tmpArr, 2) - 1
  Chk = (InStr("><=", Left(FindStr, 1)) > 0)
  For i = LBound(tmpArr, 1) - HasTitle To UBound(tmpArr, 1)
    If Chk Then
      TmpVal = CDbl(tmpArr(i, ColIndex))
      If Evaluate(TmpVal & FindStr) Then Dic.Add i, ""
    Else
      If UCase(tmpArr(i, ColIndex)) Like UCase(FindStr) Then Dic.Add i, ""
    End If
  Next
  If Dic.Count > 0 Then
    Tmp = Dic.Keys
    ReDim arr(LBound(tmpArr, 1) To UBound(Tmp) + LBound(tmpArr, 1) - HasTitle, LBound(tmpArr, 2) To UBound(tmpArr, 2))
    For i = LBound(tmpArr, 1) - HasTitle To UBound(Tmp) + LBound(tmpArr, 1) - HasTitle
      For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
        arr(i, j) = tmpArr(Tmp(i - LBound(tmpArr, 1) + HasTitle), j)
      Next
    Next
    If HasTitle Then
      For j = LBound(tmpArr, 2) To UBound(tmpArr, 2)
        arr(LBound(tmpArr, 1), j) = tmpArr(LBound(tmpArr, 1), j)
      Next
    End If
  End If
  Filter2DArray = arr
End Function

