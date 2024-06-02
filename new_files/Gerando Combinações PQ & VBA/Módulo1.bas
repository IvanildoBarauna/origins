Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub GeraCombinações()
    Dim ws              As Worksheet
    Dim lo              As ListObject
    Dim iRow            As Long, iRow2 As Long
    Dim combine         As String
    
    Set ws = Planilha1
    Set lo = ws.ListObjects(1)
    
    ws.Range("E2", ws.Cells(2, "E").End(xlDown)).ClearContents
    
    For iRow = 1 To lo.ListRows.Count
        For iRow2 = 1 To lo.ListRows.Count
        
            If lo.DataBodyRange(iRow, 1).Value2 = lo.DataBodyRange(iRow2, 1).Value2 Then
                iRow2 = iRow2 + 1
            End If
            
            If combine = "" Then
                combine = lo.DataBodyRange(iRow, 1).Value2 & lo.DataBodyRange(iRow2, 1).Value2
            Else
                combine = combine & lo.DataBodyRange(iRow2, 1).Value2
            End If
            
        Next iRow2
        ws.Cells(iRow + 1, "E").Value2 = combine
        combine = ""
    Next iRow
    
    MsgBox "Concluído", vbInformation
End Sub
