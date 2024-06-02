Attribute VB_Name = "Code"
Option Explicit
Public Sub casoProblematizador()
'Criei uma tabela com os dados conforme enunciado com código_cidade, número_veiculos_passeio, número_acidentes_vítimas
'para que possa ser lida pelo algorítimo em questão
'---------------------------------------------------------
'A varíavel MaxAcidentes mostra o número máximo de acidentes
'A varíavel MinAcidentes mostra o número mínimo de acidentes
'A varíavel CidadeMax mostra a cidade com maior número de acidentes
'A varíavel CidadeMin mostra a cidade com menor número de acidentes
'A variável AVGVeiculos mostra a média de veiculos de todas as cidades.
'A variável AVGAcidents mostra a média de acidentes nas cidades com menos de 2000 veículos de passeio.

    Dim lo As ListObject
    Dim MaxAcidentes As Integer
    Dim MinAcidentes As Integer
    Dim iRow        As Long
    Dim CidadeMax   As Integer
    Dim CidadeMin   As Integer
    Dim AVGVeiculos As Integer
    Dim iCounter    As Integer
    Dim SomaVeiculos As Integer
    Dim nAcidents As Integer
    Dim nVeiculos   As Integer
    Dim somaacidentes As Integer
    Dim div As Integer
    Dim AVGAcidents As Integer
    
    
    Set lo = Planilha1.ListObjects("fAcidentes")
    MaxAcidentes = Application.WorksheetFunction.Max(lo.ListColumns(3).DataBodyRange.Value2)
    MinAcidentes = Application.WorksheetFunction.Min(lo.ListColumns(3).DataBodyRange.Value2)
    
    For iRow = 1 To lo.ListRows.Count
        If lo.DataBodyRange(iRow, 3).Value2 = MaxAcidentes Then CidadeMax = lo.DataBodyRange(iRow, 1).Value2
    Next iRow
    
    For iRow = 1 To lo.ListRows.Count
        If lo.DataBodyRange(iRow, 2).Value2 = MinAcidentes Then CidadeMin = lo.DataBodyRange(iRow, 1).Value2
    Next iRow
    
    Do Until iRow <= lo.ListRows.Count
        SomaVeiculos = SomaVeiculos + lo.DataBodyRange(iRow, 2).Value2
        iRow = iRow + 1
    Loop
    
    AVGVeiculos = SomaVeiculos / lo.ListRows.Count
    
    For iRow = 1 To lo.ListRows.Count
        nAcidents = lo.DataBodyRange(iRow, 3).Value2
        nVeiculos = lo.DataBodyRange(iRow, 2).Value2
        If nVeiculos < 2000 Then
            somaacidentes = somaacidentes + nAcidents
            div = div + 1
        Else
            Exit For
        End If
    Next iRow
    
    AVGAcidents = somaacidentes / div
    
End Sub


