Attribute VB_Name = "TB_InfoMP"
Option Explicit

Sub Gravar_InfoMP()
    Dim iCell           As Range
    Dim rng             As Range
    Dim vData           As Date
    Dim vProduto        As String
    Dim vMP             As String
    Dim vFornecedor     As String
    Dim TotalItens      As Integer
    Dim iRow            As Integer
    Dim lo              As ListObject
    Dim lr              As ListRow
    
    Set lo = wsMP.ListObjects(1)
    Set rng = wsFormulario.Range("Matéria_Prima_Utilizada")
    
    For Each iCell In rng
        If Not iCell.Value = "" Then TotalItens = TotalItens + 1
    Next iCell
    
    If Not VBA.IsDate(wsFormulario.Range("vData").Value) Then
        MsgBox "Data Inválida.", vbCritical
        Exit Sub
    End If

    vData = wsFormulario.Range("G2").Value
    vProduto = wsFormulario.Range("C4").Value2
    
    For iRow = rng.Cells(1, 1).Row To rng.Cells(1, 1).Row + (TotalItens - 1)
        vMP = wsFormulario.Range("E" & iRow).Value2
        vFornecedor = wsFormulario.Range("G" & iRow).Value2
        Set lr = lo.ListRows.Add
        
        With lr
            .Range(lo.ListColumns("DATA").Index).Value = vData
            .Range(lo.ListColumns("PRODUTO").Index).Value2 = vProduto
            .Range(lo.ListColumns("MATÉRIA PRIMA").Index).Value2 = vMP
            .Range(lo.ListColumns("FORNECEDOR").Index).Value2 = vFornecedor
        End With
    Next iRow
    
End Sub
