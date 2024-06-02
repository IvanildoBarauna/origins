Attribute VB_Name = "Lançamento"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Dim mData       As String
Dim mLancamento As String
Dim mPagamento  As String
Dim mDescricao  As String
Dim mLocalVenda As String
Dim mVenda      As String
Dim mPriceUN    As String
Dim mCustoKG    As String
Dim mPerdido    As String

Public Property Let Data(vData As String): mData = vData: End Property
Public Property Let Lancamento(sLancamento As String): mLancamento = sLancamento: End Property
Public Property Let Pagamento(sPagamento As String): mPagamento = sPagamento: End Property
Public Property Let Descricao(sDecricao As String): mDescricao = sDecricao: End Property
Public Property Let LocalVenda(sLocal As String): mLocalVenda = sLocal: End Property
Public Property Let Venda(lVenda As String): mVenda = lVenda: End Property
Public Property Let PriceUN(lPrice As String): mPriceUN = lPrice: End Property
Public Property Let CustoKG(lCusto As String): mCustoKG = lCusto: End Property
Public Property Let QTDPerdidos(lPerdidos As String): mPerdido = lPerdidos: End Property
Public Sub Save()
    Dim arr, lo As ListObject, lr As ListRow
    
    arr = ArrToListRow
    Set lo = shCaixa.ListObjects("fCaixa")
    Set lr = lo.ListRows.Add
    
    With lr
        .Application.Range(.Range(1, 2), _
            .Range(1, lo.ListColumns("QTD PERDIDA").index)).Value2 = _
                arr
    End With
End Sub

Public Sub Change(ByVal RowIndex As Integer)
    Dim arr, lo As ListObject, lr As ListRow
    
    arr = ArrToListRow
    Set lo = shCaixa.ListObjects("fCaixa")
    Set lr = lo.ListRows(RowIndex)
    
    With lo
        .Application.Range(.DataBodyRange(RowIndex, 2), _
                .DataBodyRange(RowIndex, lo.ListColumns("QTD PERDIDA").index)).Value2 = _
                    arr
    End With
End Sub

Private Function ArrToListRow() As Variant
    Dim tmpArray(1 To 1, 1 To 9), lo As ListObject: Set lo = shCaixa.ListObjects("fCaixa")
    
    tmpArray(1, lo.ListColumns("DataLançamento").index - 1) = CLng(VBA.DateValue(mData))
    tmpArray(1, lo.ListColumns("Lançamento").index - 1) = mLancamento
    tmpArray(1, lo.ListColumns("Tipo Pagamento").index - 1) = mPagamento
    tmpArray(1, lo.ListColumns("Descrição").index - 1) = mDescricao
    tmpArray(1, lo.ListColumns("LocalVenda").index - 1) = mLocalVenda
    tmpArray(1, lo.ListColumns("Venda").index - 1) = mVenda * 1
    tmpArray(1, lo.ListColumns("Preço/Un.").index - 1) = mPriceUN * 1
    tmpArray(1, lo.ListColumns("Custo/KG.").index - 1) = mCustoKG * 1
    tmpArray(1, lo.ListColumns("QTD PERDIDA").index - 1) = mPerdido * 1
    
    ArrToListRow = tmpArray
End Function
