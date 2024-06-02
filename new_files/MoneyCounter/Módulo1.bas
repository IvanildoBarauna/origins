Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub ClearData()
    Dim ws  As Worksheet
    Dim lo  As ListObject
    
    Set ws = shContagem
    Set lo = ws.ListObjects(1)
    If lo.ListRows.Count < 2 Then
        MsgBox "Não há dados para serem apagados.", vbExclamation
    Else
        With lo.ListRows(2)
        .Application.Range(.Range(1, 1), _
            .Range(lo.ListRows.Count - 1, lo.ListColumns.Count)).Rows.Delete
        lo.Application.Range(lo.DataBodyRange(1, 1), lo.DataBodyRange(1, 2)).ClearContents
        lo.DataBodyRange(1, 1).Select
        End With
        MsgBox "Valores reiniciados", vbInformation
    End If
End Sub


