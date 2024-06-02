Attribute VB_Name = "mdMain"
Option Explicit

Public Sub ConsolidaEletropaulo(wbPath As String, wsDestino As Worksheet)
    Dim wb          As Excel.Workbook
    Dim ws          As Excel.Worksheet
    Dim FinalCol    As Long
    Dim arr         As Variant
    Dim FinalRow    As Long
    Dim rng         As Excel.Range
    Dim lo          As Excel.ListObject
    
    If wsDestino.Range("A2").Value2 <> "" Then wsfDestino.ListObjects("tb" & VBA.Right(wbPath, 11)).Delete
    
    Set wb = Application.Workbooks.Open(Filename:=wbPath & ".csv")
    Set ws = wb.Sheets(1)
    ws.Rows("1:2").Delete
    ws.Range("A1:A3").TextToColumns ws.Range("A1"), xlDelimited, xlTextQualifierNone, True, Semicolon:=True
    ws.Columns(1).Delete
    FinalCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    FinalRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    arr = ws.Range("A1").Resize(FinalRow, FinalCol).Value2
    wb.Close SaveChanges:=False
    Set rng = wsDestino.Range("A1").Resize(FinalCol, FinalRow)
    rng.Value2 = Application.WorksheetFunction.Transpose(arr)
    Set lo = wsDestino.ListObjects.Add(xlSrcRange, rng, , , rng(1, 1))
    With lo
        .ListColumns(1).Range(1, 1).Value2 = "DataREF."
        .ListColumns(1).DataBodyRange.NumberFormatLocal = "dd/mm/aaaa"
        .ListColumns(1).Range.EntireColumn.AutoFit
        .ListColumns(2).Range(1, 1).Value2 = "VALOR R$"
        .ListColumns(2).DataBodyRange.NumberFormat = "_($* #,##0.00_);_($* (#,##0.00);_($* ""-""??_);_(@_)"
        .ListColumns(2).Range.EntireColumn.AutoFit
        .ListColumns(3).Range(1, 1).Value2 = "CONSUMO KW/H"
        .ListColumns(3).DataBodyRange.NumberFormat = "0.0"
        .ListColumns(3).Range.EntireColumn.AutoFit
        .Name = "tb" & VBA.Right(wbPath, 11)
    End With
End Sub

Public Sub ImportData()
    Application.ScreenUpdating = False
    Const sPath     As String = "D:\OneDrive\DADOS_AES\"
    Dim ws          As Excel.Worksheet:  Set ws = wsInstalacoes
    Dim lo          As Excel.ListObject: Set lo = ws.ListObjects(1)
    Dim iCounter    As Integer
    Dim CPF         As String
    
    For iCounter = 1 To lo.ListRows.Count
        CPF = lo.DataBodyRange(iCounter, lo.ListColumns("CPF").Index).Value2
        ConsolidaEletropaulo sPath & CPF, ThisWorkbook.Sheets(iCounter + 1)
    Next iCounter
    
    MsgBox "Todos os relatórios foram consolidados", vbInformation
    Application.ScreenUpdating = True
End Sub
