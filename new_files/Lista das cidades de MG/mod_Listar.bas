Attribute VB_Name = "mod_Listar"
Sub CopiarCidades()
Dim LastRows         As Long   'achar a ultima linha
Dim i                As Long
Dim ws               As Worksheet
Dim Busca
Dim Linha            As Long
                                'variavelEsquerda = Left(Range("a2").Value, 2)
    Set ws = Sheet1
    Busca = Sheet1.Range("D5").Value
    LastRows = ws.Cells(Rows.Count, 1).End(3).Row
    Debug.Print LastRows
    Linha = 4
    For i = 4 To LastRows
                                   'If Left(Ws.Cells(i, 2).Value, 2) < Busca Then
    If ws.Cells(i, 1).Value <= Busca Then
                                 '    Ws.Cells(i, 1).Copy Sheet2.Cells(Linha, 1)
                                 '    Ws.Cells(i, 2).Copy Sheet2.Cells(Linha, 2)
         Sheet2.Cells(Linha, 1) = ws.Cells(i, 1).Value
         Sheet2.Cells(Linha, 2) = ws.Cells(i, 2).Value
    End If
    Linha = Linha + 1
    Next i
   
End Sub

Public Sub CitiesToRange()
    Dim oDic    As Object
    Dim ws      As Worksheet
    Dim lRow    As Long
    Dim iRow    As Long
    Dim sItem   As String
    Dim arr     As Variant
    
    Set oDic = VBA.CreateObject("Scripting.Dictionary")
    Set ws = Sheet1
    lRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    For iRow = 4 To lRow
        sItem = ws.Cells(iRow, 2).Value2
        If VBA.UCase(VBA.Left(sItem, 1)) = "A" Then
            oDic.Add sItem, sItem
        End If
    Next iRow
    
    arr = oDic.Items
    
    ws.Range("C4", ws.Cells(ws.Rows.Count, "C").End(xlUp)).ClearContents
    
    For iRow = LBound(arr) To UBound(arr)
        ws.Cells(iRow + 4, "C").Value2 = arr(iRow)
    Next iRow
    
    MsgBox "Total de " & UBound(arr) + 1 & " cidades encontradas!", vbInformation
End Sub

Public Sub CitiesToListObject()
    Dim oDic    As Object
    Dim ws      As Worksheet
    Dim lo      As ListObject
    Dim iRow    As Long
    Dim sItem   As String
    Dim arr     As Variant
    
    Set oDic = VBA.CreateObject("Scripting.Dictionary")
    Set ws = Sheet1
    Set lo = ws.ListObjects("tbCidades")
    
    For iRow = 1 To lo.ListRows.Count
        sItem = lo.DataBodyRange(iRow, 2).Value2
        If VBA.UCase(VBA.Left(sItem, 1)) = "A" Then
            oDic.Add sItem, sItem
        End If
    Next iRow
    
    arr = oDic.Items
    
    lo.ListColumns(3).DataBodyRange.ClearContents
    
    For iRow = LBound(arr) To UBound(arr)
        lo.DataBodyRange(iRow + 1, 3).Value2 = arr(iRow)
    Next iRow
    
    MsgBox "Total de " & UBound(arr) + 1 & " cidades encontradas!", vbInformation
End Sub
