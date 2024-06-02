Attribute VB_Name = "TB_InfoTorque"
Option Explicit

Sub Gravar_InfoTorque()
    Dim vData       As Date
    Dim vProduto    As String
    Dim Horario     As Date
    Dim Torque      As String
    Dim TotalItens  As Integer
    Dim iRow        As Integer
    Dim lo          As ListObject
    Dim lr          As ListRow
    Dim rng         As Range
    Dim iCell       As Range

    Set lo = wsTorque.ListObjects(1)
    Set rng = wsFormulario.Range("hTorques")
    
    For Each iCell In rng
        If Not iCell.Value = "" Then TotalItens = TotalItens + 1
    Next iCell
    
    vData = wsFormulario.Range("G2").Value
    vProduto = wsFormulario.Range("C4").Value2
    
    For iRow = rng.Cells(1, 1).Row To rng.Cells(1, 1).Row + (TotalItens - 1)
        Horario = wsFormulario.Range("J" & iRow).Value2
        Torque = wsFormulario.Range("K" & iRow).Value2
        Set lr = lo.ListRows.Add
        With lr
            .Range(lo.ListColumns("DATA").Index).Value = vData
            .Range(lo.ListColumns("PRODUTO").Index).Value2 = vProduto
            .Range(lo.ListColumns("HORÁRIO").Index).Value2 = Horario
            .Range(lo.ListColumns("TORQUE").Index).Value2 = Torque
        End With
    Next iRow
End Sub
