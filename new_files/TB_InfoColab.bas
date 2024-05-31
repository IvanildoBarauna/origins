Attribute VB_Name = "TB_InfoColab"
Option Explicit
Sub Gravar_InfoColab()
    Dim vData       As Date
    Dim vProduto    As String
    Dim vFunction   As String
    Dim vColab      As String
    Dim iRow        As Long
    Dim lo          As ListObject
    Dim lr          As ListRow
    Dim iCell       As Range
    Dim rng             As Range
    Dim TotalItens  As Integer

    Set lo = wsFuncao.ListObjects(1)
    Set rng = wsFormulario.Range("FUNÇÃO")
    
    For Each iCell In rng
        If Not iCell.Value = "" Then TotalItens = TotalItens + 1
    Next iCell
    
    
    If Not VBA.IsDate(wsFormulario.Range("vData").Value) Then
        MsgBox "Data Inválida.", vbCritical
        Exit Sub
    End If
    
    vData = wsFormulario.Range("vData").Value
    vProduto = wsFormulario.Range("vProduto").Value2
    
    For iRow = rng.Cells(1, 1).Row To rng.Cells(1, 1).Row + (TotalItens - 1)
           vFunction = wsFormulario.Range("B" & iRow).Value2
           vColab = Range("C" & iRow).Value2
           Set lr = lo.ListRows.Add
           With lr
                .Range(lo.ListColumns("DATA").Index).Value = vData
                .Range(lo.ListColumns("PRODUTO").Index).Value2 = vProduto
                .Range(lo.ListColumns("FUNÇÃO").Index).Value2 = vFunction
                .Range(lo.ListColumns("COLABORADOR").Index).Value2 = vColab
           End With
           Set lr = Nothing
    Next iRow


End Sub
