Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub ImportFile()
    Dim fDialog     As FileDialog: Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    Dim fIndex      As Integer, sPath As String, ws As Worksheet, iSheet As Integer, iLine As String, iRow As Long
    
    With fDialog
        .AllowMultiSelect = True
        .Filters.Add "Texto", "*.txt", 1
        If .Show Then
        
            If ThisWorkbook.Sheets.Count < .SelectedItems.Count Then
                For iSheet = 1 To .SelectedItems.Count - ThisWorkbook.Sheets.Count
                    ThisWorkbook.Sheets.Add
                Next iSheet
            End If
            
            For Each ws In ThisWorkbook.Sheets
                ws.Cells.Delete
                fIndex = VBA.FreeFile()
                Open .SelectedItems(ws.Index) For Input As #fIndex
                
                Do While Not VBA.EOF(fIndex)
                    Line Input #fIndex, iLine
                    iRow = iRow + 1
                    ws.Cells(iRow, 1).Value2 = iLine
                Loop
                
                Close #fIndex
            Next ws
        End If
    End With
End Sub
