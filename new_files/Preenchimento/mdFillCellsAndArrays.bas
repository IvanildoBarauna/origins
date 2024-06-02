Attribute VB_Name = "mdFillCellsAndArrays"
Option Explicit
Dim vRows   As Double
Dim vCols   As Double
Dim INI     As Single
Dim FIM     As Single
Public Sub PreencherArrayValue2(ByVal sMensagem As String)
    Dim MyArray     As Variant
    Dim ws          As Worksheet
    Dim vRow        As Long
    Dim vCol        As Long
    Dim LowerLine   As Long
    Dim UpperLine   As Long
    Dim LowerCol    As Long
    Dim UpperCol    As Long
    
    If vRows = 0 Or vCols = 0 Then Exit Sub
    INI = VBA.Timer
    Set ws = shtArrays
    ReDim MyArray(1 To vRows, 1 To vCols)

    LowerLine = LBound(MyArray, 1)
    UpperLine = UBound(MyArray, 1)
    LowerCol = LBound(MyArray, 2)
    UpperCol = UBound(MyArray, 2)
    
    For vRow = 1 To vRows
        For vCol = 1 To vCols
            MyArray(vRow, vCol) = vRow * vCol
        Next vCol
    Next vRow

    With ws
        .Range(.Cells(LowerLine, LowerCol), .Cells(UpperLine, UpperCol)).Value2 = MyArray
    End With

    FIM = VBA.Timer

    Debug.Print sMensagem
    INI = 0
    FIM = 0
End Sub

Public Sub PreencherArrayValue(ByVal sMensagem As String)
    Dim MyArray     As Variant
    Dim ws          As Worksheet
    Dim vRow        As Long
    Dim vCol        As Long
    Dim LowerLine   As Long
    Dim UpperLine   As Long
    Dim LowerCol    As Long
    Dim UpperCol    As Long
    
    If vRows = 0 Or vCols = 0 Then Exit Sub
    INI = VBA.Timer
    Set ws = shtArrays
    ReDim MyArray(1 To vRows, 1 To vCols)

    LowerLine = LBound(MyArray, 1)
    UpperLine = UBound(MyArray, 1)
    LowerCol = LBound(MyArray, 2)
    UpperCol = UBound(MyArray, 2)
    
    For vRow = 1 To vRows
        For vCol = 1 To vCols
            MyArray(vRow, vCol) = vRow * vCol
        Next vCol
    Next vRow

    With ws
        .Range(.Cells(LowerLine, LowerCol), .Cells(UpperLine, UpperCol)).Value = MyArray
    End With

    FIM = VBA.Timer
    
    Debug.Print sMensagem
    INI = 0
    FIM = 0
End Sub


Public Sub PreencherCellsValue2(ByVal sMensagem As String)
    Dim ws      As Worksheet
    Dim vRow    As Long
    Dim vCol    As Long
    
    Set ws = shtCells
    
    INI = VBA.Timer
    
    If vRows = 0 Or vCols = 0 Then Exit Sub
    
    For vRow = 1 To vRows
        For vCol = 1 To vCols
            ws.Cells(vRow, vCol).Value2 = vRow * vCol
        Next vCol
    Next vRow
    
    FIM = VBA.Timer
    
    Debug.Print sMensagem
    INI = 0
    FIM = 0
End Sub

Public Sub PreencherCellsValue(ByVal sMensagem As String)
    Dim ws      As Worksheet
    Dim vRow    As Long
    Dim vCol    As Long
    
    Set ws = shtCells
    
    INI = VBA.Timer
    
    If vRows = 0 Or vCols = 0 Then Exit Sub
    
    For vRow = 1 To vRows
        For vCol = 1 To vCols
            ws.Cells(vRow, vCol).Value = vRow * vCol
        Next vCol
    Next vRow
    
    FIM = VBA.Timer
    
    Debug.Print sMensagem
    INI = 0
    FIM = 0
End Sub

Public Sub Medição()
    Dim ws      As Worksheet
    Dim LastRow As Long
    Dim LastCol As Long
    Dim sNum    As Integer
    
    DeleteUsedRangeAllSheets
    
    vRows = VBA.InputBox("Digite a quantidade de linhas da matriz:")
    vCols = VBA.InputBox("Digite a quantidade de colunas da matriz:")
    
    Call PreencherArrayValue(sMensagem:="Preenchimento Arrays com .Value: " & VBA.Format(FIM - INI, "0.00 SEGUNDOS"))
    Call PreencherArrayValue2(sMensagem:="Preenchimento Arrays com .Value2: " & VBA.Format(FIM - INI, "0.00 SEGUNDOS"))
    Call PreencherCellsValue(sMensagem:="Preenchimento Cells com .Value: " & VBA.Format(FIM - INI, "0.00 SEGUNDOS"))
    Call PreencherCellsValue2(sMensagem:="Preenchimento Cells com .Value2: " & VBA.Format(FIM - INI, "0.00 SEGUNDOS"))
End Sub
