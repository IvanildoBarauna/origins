Attribute VB_Name = "TB_InfoParada"
Option Explicit
Sub Gravar_InfoParada()
    Dim vData           As Date
    Dim vProduto        As String
    Dim InicioParada    As Date
    Dim FimParada       As Date
    Dim CodParada       As String
    Dim Tempo_TT        As Date
    Dim TotalItens      As Integer
    Dim iRow            As Integer
    Dim lo              As ListObject
    Dim lr              As ListRow
    Dim iCell           As Range
    Dim rng             As Range
    
    Set lo = wsParadas.ListObjects(1)
    Set rng = wsFormulario.Range("nParadas")

    For Each iCell In rng
        If Not iCell.Value = "" Then TotalItens = TotalItens + 1
    Next iCell
    
    vData = wsFormulario.Range("G2").Value
    vProduto = wsFormulario.Range("C4").Value2
    
     For iRow = rng.Cells(1, 1).Row To rng.Cells(1, 1).Row + (TotalItens - 1)
        InicioParada = wsFormulario.Range("E" & iRow).Value2
        FimParada = wsFormulario.Range("F" & iRow).Value2
        CodParada = wsFormulario.Range("G" & iRow).Value2
        Tempo_TT = wsFormulario.Range("H" & iRow).Value2
        Set lr = lo.ListRows.Add
        With lr
            .Range(lo.ListColumns("DATA").Index).Value = vData
            .Range(lo.ListColumns("PRODUTO").Index).Value2 = vProduto
            .Range(lo.ListColumns("INICIO (H)").Index).Value2 = InicioParada
            .Range(lo.ListColumns("FINAL (H)").Index).Value2 = FimParada
            .Range(lo.ListColumns("CÓD. PARADA  MOTIVO").Index).Value2 = CodParada
            .Range(lo.ListColumns("TEMPO GASTO").Index).Value2 = Tempo_TT
        End With
    Next iRow
End Sub
