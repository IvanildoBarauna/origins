Attribute VB_Name = "Módulo1"
Option Explicit
Public Enum Mode
    Import = 1
    Export = 0
End Enum

Public Sub VBAReferences(ByVal xlMode As Mode)
    Dim ws          As Worksheet
    Dim lRow        As Long
    Dim iCounter    As Long
    Dim sGUID       As String
    Dim dblMin      As Double
    Dim dblMax      As Double
    Dim sValuesDesc As String
    
    Set ws = Planilha1
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    With ThisWorkbook.VBProject.References
        If xlMode = Export Then
            ws.Cells(2, 1).Resize(lRow, 4).ClearContents
            For iCounter = 1 To .Count
                sValuesDesc = .Item(iCounter).Name & " - " & .Item(iCounter).Description
                sGUID = .Item(iCounter).GUID
                dblMin = .Item(iCounter).Minor
                dblMax = .Item(iCounter).Major
                
                ws.Range("A" & iCounter + 1).Value2 = sValuesDesc
                ws.Range("B" & iCounter + 1).Value2 = sGUID
                ws.Range("C" & iCounter + 1).Value2 = dblMin
                ws.Range("D" & iCounter + 1).Value2 = dblMax
            Next iCounter
            ws.Range("A:D").EntireColumn.AutoFit
        Else
            For iCounter = 1 To .Count
                sValuesDesc = .Item(iCounter).Name & " - " & .Item(iCounter).Description
                sGUID = .Item(iCounter).GUID
                dblMin = .Item(iCounter).Minor
                dblMax = .Item(iCounter).Major
                
                .AddFromGuid sGUID, dblMax, dblMin
            Next iCounter
        End If
        
    End With
    
End Sub


