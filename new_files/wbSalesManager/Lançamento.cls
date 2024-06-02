Attribute VB_Name = "Lançamento"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Type ClassType
     mData       As String
     mLancamento As String
     mPagamento  As String
     mDescricao  As String
     mLocalVenda As String
     mVenda      As String
     mPriceUN    As String
     mCustoKG    As String
     mPerdido    As String
End Type

Private This As ClassType

Public Property Let Data(Value As String): This.mData = Value: End Property
Public Property Let Lancamento(Value As String): This.mLancamento = Value: End Property
Public Property Let Pagamento(Value As String): This.mPagamento = Value: End Property
Public Property Let Descricao(Value As String): This.mDescricao = Value: End Property
Public Property Let LocalVenda(Value As String): This.mLocalVenda = Value: End Property
Public Property Let Venda(Value As String): This.mVenda = Value: End Property
Public Property Let PriceUN(Value As String): This.mPriceUN = Value: End Property
Public Property Let CustoKG(Value As String): This.mCustoKG = Value: End Property
Public Property Let QTDPerdidos(Value As String): This.mPerdido = Value: End Property

Public Sub Save()
    Dim arr, lo As Excel.ListObject, lr As Excel.ListRow
    
    arr = ArrToListRow
    Set lo = shCaixa.ListObjects("fCaixa")
    Set lr = lo.ListRows.Add
    
    With lr
        .Application.Range(.Range(1, lo.ListColumns("DataLançamento").index), _
                           .Range(1, lo.ListColumns("QTD PERDIDA").index)).Value2 = arr
    End With
End Sub

Public Sub Change(RowIndex As Integer)
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
    
    With This
        tmpArray(1, lo.ListColumns("DataLançamento").index - 1) = VBA.Conversion.CLng(VBA.DateValue(.mData))
        tmpArray(1, lo.ListColumns("Lançamento").index - 1) = .mLancamento
        tmpArray(1, lo.ListColumns("Tipo Pagamento").index - 1) = .mPagamento
        tmpArray(1, lo.ListColumns("Descrição").index - 1) = .mDescricao
        tmpArray(1, lo.ListColumns("LocalVenda").index - 1) = .mLocalVenda
        tmpArray(1, lo.ListColumns("Venda").index - 1) = .mVenda * 1
        tmpArray(1, lo.ListColumns("Preço/Un.").index - 1) = .mPriceUN * 1
        tmpArray(1, lo.ListColumns("Custo/KG.").index - 1) = .mCustoKG * 1
        
        If .mPerdido = "" Then
            tmpArray(1, lo.ListColumns("QTD PERDIDA").index - 1) = 0
        Else
            tmpArray(1, lo.ListColumns("QTD PERDIDA").index - 1) = .mPerdido * 1
        End If
    End With
    ArrToListRow = tmpArray
End Function
