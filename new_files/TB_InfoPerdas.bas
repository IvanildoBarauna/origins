Attribute VB_Name = "TB_InfoPerdas"
Option Explicit


Sub Gravar_InfoPerdas()
    Dim vData           As Date
    Dim vProduto        As String
    Dim Perdas          As String
    Dim Qtd             As String
    Dim TotalItens      As Integer
    Dim iRow            As Integer
    Dim iCell           As Range
    Dim rng             As Range
    Dim lo              As ListObject
    Dim lr              As ListRow

    Set lo = wsPerdas.ListObjects(1)
    Set rng = wsFormulario.Range("nPerdas")

    For Each iCell In rng
        If Not iCell.Value = "" Then TotalItens = TotalItens + 1
    Next iCell
    
    vData = wsFormulario.Range("G2").Value
    vProduto = wsFormulario.Range("C4").Value2
    
    For iRow = rng.Cells(1, 1).Row To rng.Cells(1, 1).Row + (TotalItens - 1)
        Perdas = Range("B" & iRow).Value2
        Qtd = Range("C" & iRow).Value2
        Set lr = lo.ListRows.Add
        
        With lr
            .Range(lo.ListColumns("DATA").Index).Value = vData
            .Range(lo.ListColumns("PRODUTO").Index).Value2 = vProduto
            .Range(lo.ListColumns("ITEM").Index).Value2 = Perdas
            .Range(lo.ListColumns("QUANTIDADE").Index).Value2 = Qtd
        End With
    Next iRow
    
End Sub

