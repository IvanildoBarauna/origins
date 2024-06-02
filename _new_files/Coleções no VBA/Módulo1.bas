Attribute VB_Name = "Módulo1"
Option Explicit
Option Private Module
Sub Main()
#Const DESIGN_MODE = True
    Dim ws      As Excel.Worksheet
    Dim lRow    As Long
    Dim lCtrl   As Long
    Dim iTime   As Single
    Dim oCol    As Collection
    Dim MyArray As Variant
    
#If DESIGN_MODE Then
    Dim oDic    As Scripting.Dictionary
    Set oDic = New Scripting.Dictionary
#Else
    Dim oDic    As Object
    Set oDic = VBA.CreateObject("Scripting.Dictionary")
#End If

    Set ws = Planilha1
    lRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    Set oCol = New Collection
    ReDim MyArray(1 To lRow, 1 To 1) As Long
    
    iTime = VBA.Timer
    
    For lCtrl = 1 To lRow
        oCol.Add ws.Range("A" & lCtrl).Value2, VBA.Conversion.CStr(lCtrl)
    Next lCtrl
    
    Debug.Print "Desempenho Collection: " & VBA.Format(VBA.Timer - iTime, "0.00 segundos")
    
    iTime = VBA.Timer
    
    For lCtrl = 1 To lRow
        oDic.Add VBA.Conversion.CStr(lCtrl), ws.Range("A" & lCtrl).Value2
    Next lCtrl
    
    Debug.Print "Desempenho Dictionary: " & VBA.Format(VBA.Timer - iTime, "0.00 segundos")
    
    iTime = VBA.Timer
    
    For lCtrl = 1 To UBound(MyArray)
        MyArray(lCtrl, 1) = ws.Range("A" & lCtrl).Value2
    Next lCtrl
    
    Debug.Print "Desempenho Array: " & VBA.Format(VBA.Timer - iTime, "0.00 segundos")
End Sub

Sub MainTeste()
    Dim lCtrl As Long
    
    For lCtrl = 1 To 4 Step 1
        Debug.Print "------------ " & lCtrl & "º TESTE" & " ---------------"
        Call Main
    Next lCtrl
    
    Debug.Print "------------ CONCLUÍDO --------------"
End Sub

Public Sub ArrayList()
    Dim ws      As Worksheet
    Dim lRow    As Long
    Dim List    As Object
    Dim NewItem As String
    Dim Values  As Variant
    Dim iRow    As Long
    
    Set List = VBA.CreateObject("System.Collections.ArrayList")
    Set ws = Planilha1
    lRow = ws.Range("A" & ws.Rows.Count).End(xlUp).Row
    
    For iRow = 1 To lRow
        NewItem = ws.Cells(iRow, "A").Value2
        If Not List.contains(NewItem) Then
            List.Add NewItem
        End If
    Next iRow
    
    List.Sort
    Values = WorksheetFunction.Transpose(List.ToArray)
    
End Sub
