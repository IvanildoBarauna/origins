Attribute VB_Name = "TB_InfoProduto"
Option Explicit


Sub Gravar_InfoProduto()
    Dim lr          As ListRow
    Dim lo          As ListObject
    
    Set lo = wsProduto.ListObjects(1)
    Set lr = lo.ListRows.Add

    With lr
        .Range(lo.ListColumns("DATA").Index).Value = wsFormulario.Range("G2").Value
        .Range(lo.ListColumns("PRODUTO").Index).Value2 = wsFormulario.Range("C4").Value
        .Range(lo.ListColumns("LOTE").Index).Value2 = wsFormulario.Range("C5").Value
        .Range(lo.ListColumns("TOTAL PRODUZIDO").Index).Value2 = wsFormulario.Range("C6").Value
        .Range(lo.ListColumns("INÍCIO PRODUÇÃO").Index).Value2 = wsFormulario.Range("C7").Value
        .Range(lo.ListColumns("FINAL PRODUÇÃO").Index).Value2 = wsFormulario.Range("C8").Value
        .Range(lo.ListColumns("TEMPO TOTAL").Index).Value2 = wsFormulario.Range("C9").Value
        .Range(lo.ListColumns("OBSERVAÇÃO").Index).Value2 = wsFormulario.Range("E22").Value
    End With
End Sub
