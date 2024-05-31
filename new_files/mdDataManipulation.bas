Attribute VB_Name = "mdDataManipulation"
Option Explicit
Private Sub InsertReg(ByRef item As String, _
              ByRef marca As String, _
              ByVal sessao As String, _
              ByRef price As Currency, _
              ByRef quantidade As Double)
              
    Dim lo      As Excel.ListObject
    Dim lr      As Excel.ListRow
    Dim loAux   As Excel.ListObject
    
    Excel.Application.EnableEvents = False
    
    Set loAux = wsApoios.ListObjects("tbSessao")
    Set lo = wsConsolidado.ListObjects(1)
    Set lr = lo.ListRows.Add
    
    With lr
        .Range(lo.ListColumns("ITEM").Index).Value2 = VBA.Strings.UCase(item)
        .Range(lo.ListColumns("MARCA").Index).Value2 = VBA.Strings.UCase(marca)
        .Range(lo.ListColumns("SESSÃO").Index).Value2 = sessao
        .Range(lo.ListColumns("DATA_REF").Index).Value2 = VBA.DateSerial(Year(Date), Month(Date), 15)
        .Range(lo.ListColumns("PREÇO").Index).Value2 = price
        .Range(lo.ListColumns("QTD").Index).Value2 = quantidade
        .Range(lo.ListColumns("VALIDA").Index).Value2 = "NÃO COMPRADO"
    End With
    Excel.Application.EnableEvents = True
End Sub

Private Sub SortListObject(lo As ListObject, pOrder As XlSortOrder, ParamArray ColumnsIndex())
    Dim iCol As Variant
    
    On Error GoTo err
    With lo.Sort
        .SortFields.Clear
        For Each iCol In ColumnsIndex
            .SortFields.Add2 lo.ListColumns(iCol).DataBodyRange, xlSortOnValues, pOrder, xlSortNormal
        Next iCol
        .Header = VBA.IIf(lo.ShowHeaders, xlYes, xlNo)
        .Apply
    End With
    Exit Sub
err:
    Debug.Print err.Description
End Sub

Sub InsertAndSortListObject()
    Dim ws      As Worksheet
    Dim item    As String
    Dim marca   As String
    Dim sessao  As String
    Dim price       As Currency
    Dim quantidade  As Double
    Dim rng As Range, iCell As Range
    
    On Error GoTo err
    
    Set ws = wsConsolidado
    Set rng = ws.Range("A5:E5")
    
    For Each iCell In rng
        If iCell.Value2 = "" Or iCell = 0 Then
            MsgBox "A célula " & iCell.AddressLocal & " está vazia ou com valor igual a 0.", vbExclamation
            Exit Sub
        End If
    Next iCell
    
    item = ws.Range("A5").Value2
    marca = ws.Range("B5").Value2
    sessao = ws.Range("C5").Value2
    price = ws.Range("D5").Value2
    quantidade = ws.Range("E5").Value2
    
    rng.Value2 = ""
    
    InsertReg item, marca, sessao, price, quantidade
    rng(1, 1).Select
    
    MsgBox "Produto cadastrado com sucesso.", vbInformation
    Exit Sub
err:
    MsgBox "Não foi possível realizar o cadastro. " & err.Description, vbCritical
End Sub

Sub CustomSort()
    Dim lo As ListObject
    
    Set lo = wsConsolidado.ListObjects(1)
    
    SortListObject lo, xlAscending, "DATA_REF", "Sessão", "item"
End Sub
