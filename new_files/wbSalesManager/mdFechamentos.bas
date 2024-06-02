Attribute VB_Name = "mdFechamentos"
Option Explicit
Sub InserirFechamento()
    Dim lo      As Excel.ListObject
    Dim lr      As Excel.ListRow
    Dim xCtrl   As String, VendaReal, VendaEsperada As Double
    
    If isClosed() Then
        MsgBox "O dia já está fechado, os lançamentos já foram realizados", vbInformation
        Exit Sub
    End If
    
    xCtrl = wsContagem.Range("Type").Value2
    VendaReal = wsContagem.Range("Venda").Value2
    
    If xCtrl = "VENDA" And VendaReal > 0 Then
        If VBA.MsgBox("Tem certeza que deseja realizar o lançamento com os valores atuais?", vbQuestion + vbYesNo) = vbYes Then
            VendaEsperada = VendaTotal(VBA.DateTime.Date)
            If VendaEsperada <= 0 Then Exit Sub
            Set lo = wsFechamentos.ListObjects(1)
            Set lr = lo.ListRows.Add
            With lr
                .Range(lo.ListColumns("Data").index).Value2 = VBA.DateTime.Date
                .Range(lo.ListColumns("VendaReal").index).Value2 = VendaReal
                .Range(lo.ListColumns("VendaEsperada").index).Value2 = VendaEsperada
                .Range(lo.ListColumns("R$ Quebra").index).Value2 = VendaReal - VendaEsperada
                VBA.Interaction.MsgBox "Fechamento realizado com sucesso.", vbInformation
            End With
        End If
    End If
End Sub

Private Function VendaTotal(vDate As Date) As Double
    Dim lo       As Excel.ListObject
    Dim cData    As Excel.ListColumn
    Dim cLanc    As Excel.ListColumn
    Dim cVenda   As Excel.ListColumn
    Dim iCounter As Long
    Dim vSum     As Double

    Set lo = shCaixa.ListObjects(1)
    Set cData = lo.ListColumns("DataLançamento")
    Set cLanc = lo.ListColumns("Lançamento")
    Set cVenda = lo.ListColumns("Venda")
    
    For iCounter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(iCounter, cData.index).Value2 = vDate And _
            lo.DataBodyRange(iCounter, cLanc.index).Value2 = "RECEITA" Then
            vSum = vSum + lo.DataBodyRange(iCounter, cVenda.index).Value2
        End If
    Next iCounter
    
    VendaTotal = vSum
End Function

Function isClosed() As Boolean
    Dim lo          As Excel.ListObject
    Dim iCounter    As Long
    
    Set lo = wsFechamentos.ListObjects(1)
    
    For iCounter = 1 To lo.ListRows.Count
        If lo.DataBodyRange(iCounter, 1).Value = VBA.Date Then
            isClosed = True
            Exit For
        End If
    Next iCounter
    
End Function
