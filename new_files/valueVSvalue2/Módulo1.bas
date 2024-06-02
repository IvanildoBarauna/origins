Attribute VB_Name = "Módulo1"
Option Explicit
Private vRows As Long
Private vCols As Integer

Public Sub FillingWithValue()
    Dim ws      As Worksheet
    Dim vRow    As Long
    Dim vCol    As Integer
    Dim INI     As Single
    Dim FIM     As Single
    
    Set ws = shValue
    
    INI = VBA.Timer
    
    For vRow = 1 To vRows
        For vCol = 1 To vCols
            ws.Cells(vRow, vCol).Value = Date
        Next vCol
    Next vRow
    
    FIM = VBA.Timer
    
    Debug.Print "Preenchimento de Células utilizando Cells().Value" & " = " & VBA.Format(FIM - INI, "0.00  segundos")
End Sub

Public Sub FillingWithValue2()
    Dim ws      As Worksheet
    Dim vRow    As Long
    Dim vCol    As Integer
    Dim INI     As Single
    Dim FIM     As Single
    
    Set ws = shValue2
    
    INI = VBA.Timer
    
    For vCol = 1 To vCols
        For vRow = 1 To vRows
            ws.Cells(vRow, vCol).Value2 = Date
        Next vRow
    Next vCol
    
    FIM = VBA.Timer
    
    Debug.Print "Preenchimento de Células utilizando Cells().Value2" & " = " & VBA.Format(FIM - INI, "0.00  segundos")
End Sub

Public Sub FillingWithCells()
    Dim ws      As Worksheet
    Dim vRow    As Long
    Dim vCol    As Integer
    Dim INI     As Single
    Dim FIM     As Single
    
    Set ws = shCells
    
    INI = VBA.Timer
    
    For vCol = 1 To vCols
        For vRow = 1 To vRows
            ws.Cells(vRow, vCol) = Date
        Next vRow
    Next vCol
    
    FIM = VBA.Timer
    
    Debug.Print "Preenchimento de Células utilizando Cells()" & " = " & VBA.Format(FIM - INI, "0.00  segundos")
End Sub


Public Sub Medicao()
    Call DeleteUsedRangeAllSheets
    vRows = VBA.InputBox("Quantidade de linhas:") * 1
    vCols = VBA.InputBox("Quantidade de colunas:") * 1
    
    Call FillingWithValue
    Call FillingWithValue2
    Call FillingWithCells
End Sub
