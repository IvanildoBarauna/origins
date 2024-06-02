Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub Main()
    Const sPath     As String = "D:\OneDrive\wbOrçamentoPessoal.xlsm"
    Dim wb          As Excel.Workbook
    Dim ws          As Excel.Worksheet
    Dim arrDados    As Variant
    Dim wsDestino   As Excel.Worksheet
    Dim iRow        As Long
    Dim iCol        As Integer
    Dim mtzResult   As Variant
    
    Set wb = Application.Workbooks.Open(sPath)
    Set ws = wb.Sheets(1)
    
    arrDados = ws.UsedRange.Value
    
    Set wsDestino = ThisWorkbook.Sheets.Add
    
    For iRow = 1 To UBound(arrDados, 1)
        For iCol = 1 To UBound(arrDados, 2)
            If Not VBA.IsError(arrDados(iRow, iCol)) Then mtzSize = mtzSize + 1
        Next iCol
    Next iRow
    
    With wsDestino
        .Range("A1").Resize(UBound(arrDados, 1), UBound(arrDados, 1)).Value2 = arrDados
    End With
    
End Sub
