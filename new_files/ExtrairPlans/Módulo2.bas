Attribute VB_Name = "Módulo2"
Option Explicit
Public Const DIV = 1000
Public Function QuantItems(nItems) As Variant
    Dim arr As Variant
    Dim steps As Integer: steps = nItems / DIV
    Dim iCounter As Integer
    
    ReDim arr(1 To steps)
    For iCounter = 1 To steps
        arr(iCounter) = DIV * iCounter
    Next iCounter
    QuantItems = arr
End Function

Sub ExportToCSV()
    Dim ws      As Worksheet: Set ws = Planilha1
    Dim LastRow As Long: LastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).Row
    Dim item    As Variant
    Dim arr2    As Variant
    Dim wb      As Workbook
    Dim arr     As Variant: arr = QuantItems(LastRow)
    Dim iCounter As Integer
    Dim wsNew    As Worksheet
    Application.ScreenUpdating = False
    For iCounter = 1 To UBound(arr)
        If iCounter = 1 Then
            arr2 = ws.Range("A1").Resize(DIV - 1).Value2
        Else
            arr2 = ws.Range("A" & arr(iCounter - 1)).Resize(DIV).Value2
        End If
        Set wb = Application.Workbooks.Add
        Set wsNew = wb.Sheets(1)
        wsNew.Range("A1").Resize(UBound(arr2)).Value2 = arr2
        wb.SaveAs ThisWorkbook.Path & "\" & iCounter, FileFormat:=xlCSV
        wb.Close
    Next iCounter
    Application.ScreenUpdating = True
    MsgBox "Processo concluido", vbInformation
End Sub
