Attribute VB_Name = "mdMain"
Option Explicit
Public Sub Inserir()
    On Error GoTo ExitPoint
    Const ColumnQuant   As String = "QUANTIDADE"
    Const ColumnImport  As String = "IMPORTÂNCIA"
    Const appname       As String = "Contagem"
    Dim lo              As Excel.ListObject
    Dim import          As Variant
    Dim quant           As Variant
    Dim lr              As Excel.ListRow
    Dim ExistRow        As Long
    
    Do
        import = VBA.InputBox("Digite a importância:", appname)
        If Not VBA.IsNumeric(import) Or import = 0 Or import = "" Or Not isValidCurrency(import) Then GoTo ExitPoint
        
        quant = VBA.InputBox("Digite a quantidade:", appname)
        If Not VBA.IsNumeric(quant) Or quant = 0 Then GoTo ExitPoint
        
        Set lo = wsMain.ListObjects("tb" & appname)
        ExistRow = isValidItem(VBA.Conversion.CDbl(import), lo.ListColumns(ColumnImport).Index)
        
        If ExistRow > 0 Then Set lr = lo.ListRows(ExistRow) Else Set lr = lo.ListRows.Add
        
        With lr
            If ExistRow > 0 Then
                .Range(, lo.ListColumns(ColumnQuant).Index).Value2 = .Range(, lo.ListColumns(ColumnQuant).Index).Value2 + VBA.Conversion.CInt(quant)
            Else
                .Range(, lo.ListColumns(ColumnImport).Index).Value2 = VBA.Conversion.CDbl(import)
                .Range(, lo.ListColumns(ColumnQuant).Index).Value2 = VBA.Conversion.CInt(quant)
            End If
        End With
    Loop
ExitPoint:
    If wsMain.Range("C6").Value2 <> 0 Then
        MsgBox "Total em Dinheiro: " & VBA.FormatCurrency(wsMain.Range("C3").Value2, 2) & vbNewLine & _
                "Total em Moeda: " & VBA.FormatCurrency(wsMain.Range("C2").Value2, 2) & vbNewLine & _
                "Valor Total: " & VBA.FormatCurrency(wsMain.Range("C6").Value2, 2), vbInformation
    End If
End Sub
    
Private Function isValidItem(inputValue As Double, iCol As Integer) As Long
    Dim lo As Excel.ListObject
    Dim counter As Long
    
    Set lo = wsMain.ListObjects(1)
    
    For counter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(counter, iCol).Value2 = inputValue Then
            isValidItem = counter
            Exit For
        End If
    Next counter
    
End Function

Function isValidCurrency(ByVal import As Variant) As Boolean
    Dim arrAcept(1 To 11) As Double, item As Range
    
    On Error Resume Next
    import = VBA.Conversion.CDbl(import)
    On Error GoTo 0
    
    For Each item In wsAux.Range("Imports")
        If item.Value2 = import Then
            isValidCurrency = True
            Exit For
        End If
    Next item
    
End Function

Sub ClearData()
    wsMain.ListObjects("tbContagem").DataBodyRange.Delete
End Sub
