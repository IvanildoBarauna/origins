## AiosaPlan.xslm
Attribute VB_Name = "Módulo1"
Option Explicit
Public Sub UnicosDic()
    Dim dic     As Scripting.Dictionary
    Dim ws      As Worksheet
    Dim lRow    As Long
    Dim iRow    As Long
    Dim item    As String
    
    Set dic = New Scripting.Dictionary
    Set ws = Planilha1
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For iRow = 2 To lRow
        item = ws.Cells(iRow, 1).Value2
        If Not dic.Exists(item) Then dic.Add item, item
    Next iRow
    
    With ws
        .Range("B2").Resize(UBound(dic.Items) + 1, 1).Value2 = WorksheetFunction.Transpose(dic.Items)
    End With
    
End Sub

Public Sub UnicosCol()
    Dim col     As Collection
    Dim ws      As Worksheet
    Dim lRow    As Long
    Dim iRow    As Long
    Dim item    As String
    
    Set col = New Collection
    Set ws = Planilha1
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For iRow = 2 To lRow
        item = ws.Cells(iRow, 1).Value2
        On Error Resume Next
        col.Add item, item
        On Error GoTo 0
    Next iRow
    
    For iRow = 1 To col.Count
        ws.Cells(iRow + 1, "C").Value2 = col.item(iRow)
    Next iRow

End Sub

Public Sub UnicosArrayList()
    Dim MyList  As Object
    Dim ws      As Worksheet
    Dim lRow    As Long
    Dim iRow    As Long
    Dim item    As Integer
    
    Set MyList = VBA.CreateObject("System.Collections.ArrayList")
    Set ws = Planilha1
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For iRow = 2 To lRow
        item = ws.Cells(iRow, 1).Value2
        If Not MyList.Contains(item) Then MyList.Add item
    Next iRow
    
    With ws
        .Range("D2").Resize(UBound(MyList.toarray) + 1, 1).Value2 = WorksheetFunction.Transpose(MyList.toarray)
    End With
    
End Sub

Public Sub Teste()
    Dim iCounter As Integer
    Dim InitialTime As Single
    
    Planilha1.Range("B2:D11").ClearContents

        InitialTime = VBA.Timer
        Call UnicosDic
        Debug.Print VBA.Format(VBA.Timer - InitialTime, "0.00 segundos") & " Dictionary"
        InitialTime = VBA.Timer
        Call UnicosCol
        Debug.Print VBA.Format(VBA.Timer - InitialTime, "0.00 segundos") & " Collection"
        InitialTime = VBA.Timer
        Call UnicosArrayList
        Debug.Print VBA.Format(VBA.Timer - InitialTime, "0.00 segundos") & " ArrayList"
    
End Sub

