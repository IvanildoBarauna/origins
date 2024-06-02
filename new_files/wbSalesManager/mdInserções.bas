Attribute VB_Name = "mdInserções"
Option Explicit
Sub Inserir()
    On Error GoTo ExitPoint
    Const ColumnQuant   As String = "QUANTIDADE"
    Const ColumnImport  As String = "IMPORTÂNCIA"
    Const appname       As String = "Contagem"
    Dim lo              As Excel.ListObject
    Dim import          As Variant
    Dim quant           As Variant
    Dim lr              As Excel.ListRow
    Dim ExistRow        As Long
    Dim CurType         As String
    
    Do
        import = VBA.InputBox("Digite a importância:", appname, import)
        If Not VBA.IsNumeric(import) Or import = 0 Or import = "" Or Not isValidCurrency(import) Then GoTo ExitPoint
        
        CurType = CurrencyType(import)
        quant = VBA.InputBox("Qual a quantidade de " & VBA.Split(CurType, ";")(0) & " de " & FormatCurrency(import, 2) & " " & VBA.Split(CurType, ";")(1) & "?")
        If Not VBA.IsNumeric(quant) Or quant = 0 Then GoTo ExitPoint
        
        Set lo = wsContagem.ListObjects("tb" & appname)
        ExistRow = isValidItem(VBA.Conversion.CDbl(import), lo.ListColumns(ColumnImport).index)
        
        If ExistRow > 0 Then Set lr = lo.ListRows(ExistRow) Else Set lr = lo.ListRows.Add
        
        With lr
            If ExistRow > 0 Then
                .Range(, lo.ListColumns(ColumnQuant).index).Value2 = .Range(, lo.ListColumns(ColumnQuant).index).Value2 + VBA.Conversion.CInt(quant)
            Else
                .Range(, lo.ListColumns(ColumnImport).index).Value2 = VBA.Conversion.CDbl(import)
                .Range(, lo.ListColumns(ColumnQuant).index).Value2 = VBA.Conversion.CInt(quant)
            End If
        End With
        
        Application.ScreenUpdating = False
        With lo.Sort
            .SortFields.Clear
            .SortFields.Add2 Key:=lo.ListColumns("IMPORTÂNCIA").Range, SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .Apply
        End With
        Application.ScreenUpdating = True
    Loop
ExitPoint:
    If wsContagem.Range("C6").Value2 <> 0 Then
        MsgBox "Total em Dinheiro: " & VBA.FormatCurrency(wsContagem.Range("C3").Value2, 2) & vbNewLine & _
                "Total em Moeda: " & VBA.FormatCurrency(wsContagem.Range("C2").Value2, 2) & vbNewLine & _
                "Valor Total: " & VBA.FormatCurrency(wsContagem.Range("C6").Value2, 2), vbInformation
    End If
End Sub
    
Private Function isValidItem(inputValue As Double, iCol As Integer) As Long
    Dim lo As Excel.ListObject
    Dim counter As Long
    
    Set lo = wsContagem.ListObjects(1)
    
    For counter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(counter, iCol).Value2 = inputValue Then
            isValidItem = counter
            Exit For
        End If
    Next counter
    
End Function

Private Function isValidCurrency(ByVal import As Variant) As Boolean
    Dim arrAcept(1 To 11) As Double, item As Range
    
    On Error Resume Next
    import = VBA.Conversion.CDbl(import)
    On Error GoTo 0
    
    For Each item In wsContagemAux.Range("Imports")
        If item.Value2 = import Then
            isValidCurrency = True
            Exit For
        End If
    Next item
    
End Function

Function CurrencyType(Value As Variant) As String
    Dim sAux As String, sAux2 As String
    
    If isValidCurrency(Value) Then
        If Value < 2 Then sAux = "moedas" Else sAux = "notas"
        Select Case Value
            Case 0.05 To 0.5
                sAux2 = "centavos"
            Case 2 To 100
                sAux2 = "reais"
            Case Else
                sAux2 = "real"
        End Select
    End If
    
    CurrencyType = sAux & ";" & sAux2
End Function
Sub ClearData()
    With wsContagem
        If Not .ListObjects(1).DataBodyRange Is Nothing Then .ListObjects("tbContagem").DataBodyRange.Delete
        .Range("Troco").Value2 = 0
        .Range("Cartao").Value2 = 0
    End With
End Sub
