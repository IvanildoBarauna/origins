Attribute VB_Name = "mdDataControl"
Option Explicit
Public Function RangeToComboBox(ByVal IDControle As Integer) As Range
    Const InitialRow        As Integer = 4
    Const iColIDLançamentos As String = "A"
    Const iColLançamentos   As String = "B"
    Const iColIDPagamentos  As String = "D"
    Const iColPagamentos    As String = "E"
    Const iColIDVenda       As String = "H"
    Const iColVenda         As String = "i"
    Const iColIDCusto       As String = "K"
    Const iColCusto         As String = "L"
    
    Dim ws              As Worksheet
    Dim rngTipoVenda    As Range, lRowTipoVenda   As Long 'ID1
    Dim rngTipoCusto    As Range, lRowTipoCusto   As Long 'ID2
    Dim rngLançamentos  As Range, lRowLançamentos As Long 'ID3
    Dim rngPagamentos   As Range, lRowPagamentos  As Long 'ID4
        
    Set ws = sApoio
    
    With ws
        lRowLançamentos = .Cells(.Rows.Count, iColIDLançamentos).End(xlUp).Row
        Set rngLançamentos = .Range(.Cells(InitialRow, iColIDLançamentos), .Cells(lRowLançamentos, iColLançamentos))
        lRowPagamentos = .Cells(.Rows.Count, iColIDPagamentos).End(xlUp).Row
        Set rngPagamentos = .Range(.Cells(InitialRow, iColIDPagamentos), .Cells(lRowPagamentos, iColPagamentos))
        lRowTipoVenda = .Cells(.Rows.Count, iColIDVenda).End(xlUp).Row
        Set rngTipoVenda = .Range(.Cells(InitialRow, iColIDVenda), .Cells(lRowTipoVenda, iColVenda))
        lRowTipoCusto = .Cells(.Rows.Count, iColIDCusto).End(xlUp).Row
        Set rngTipoCusto = .Range(.Cells(InitialRow, iColIDCusto), .Cells(lRowTipoCusto, iColCusto))
    End With
    
    Select Case IDControle
        Case 1: Set RangeToComboBox = rngTipoVenda
        Case 2: Set RangeToComboBox = rngTipoCusto
        Case 3: Set RangeToComboBox = rngLançamentos
        Case 4: Set RangeToComboBox = rngPagamentos
    End Select
    
End Function

Public Sub SaveOnListObject(ByVal FRMSource As MSForms.UserForm)
    Dim oLanc As Lançamento
    
    Set oLanc = New Lançamento

    With oLanc
        .Data = FRMSource.txtData.Value
        .Lancamento = FRMSource.cbolanc.Text
        .Pagamento = FRMSource.cbopgto.Text
        .Descricao = FRMSource.cbodesc.Text
        .LocalVenda = sOptButtons(FRMSource)
        .Venda = FRMSource.txtvenda.Value
        .PriceUN = FRMSource.txtpreco.Value
        .CustoKG = PreçoCusto()
        .QTDPerdidos = FRMSource.txtqtdperdida.Value
        .Save
    End With
    
    SortAndFormatListObject shCaixa, shCaixa.ListObjects("fCaixa"), 2, "DD/MM/YYYY"
End Sub

Public Sub ChangeDataOnListObject(ByVal FRMSource As MSForms.UserForm, ByVal iChange As Integer)
    Dim oLanc As Lançamento
    
    Set oLanc = New Lançamento
    
    With oLanc
        .Data = FRMSource.txtData.Value
        .Lancamento = FRMSource.cbolanc.Text
        .Pagamento = FRMSource.cbopgto.Text
        .Descricao = FRMSource.cbodesc.Text
        .LocalVenda = sOptButtons(FRMSource)
        .Venda = FRMSource.txtvenda.Value
        .PriceUN = FRMSource.txtpreco.Value
        .CustoKG = PreçoCusto()
        .QTDPerdidos = FRMSource.txtqtdperdida.Value
        .Change (iChange)
    End With
    
    SortAndFormatListObject shCaixa, shCaixa.ListObjects("fCaixa"), 2, "DD/MM/YYYY"
End Sub

Public Function sOptButtons(ByRef FRM As UserForm) As String
    Dim xControl As MSForms.control
    
    For Each xControl In FRM.Controls
        If TypeOf xControl Is MSForms.OptionButton Then
            If xControl.Value Then
                sOptButtons = xControl.Caption
                Exit Function
            End If
        End If
    Next xControl
End Function

Public Sub SortAndFormatListObject(ByRef ws As Worksheet, _
                                   ByRef lo As ListObject, _
                                   ByVal ColIndex As Integer, _
                                   ByVal sFormat As String)
    With lo
        .Sort.SortFields.Clear
        .Sort.SortFields.Add _
            Key:=lo.ListColumns(ColIndex).Range, SortOn:=xlSortOnValues, _
            Order:=xlAscending, DataOption:=xlSortNormal
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With
    
    lo.ListColumns(ColIndex).DataBodyRange.NumberFormat = sFormat
End Sub

Public Sub ClearFields(ByVal FRM As MSForms.UserForm)
    Dim xCtrl As MSForms.control
    
    For Each xCtrl In FRM.Controls
        Select Case VBA.TypeName(xCtrl)
            Case "ComboBox", "TextBox"
                If xCtrl.Name = "txtData" Then xCtrl.Value = VBA.Date Else xCtrl.Value = ""
            Case "OptionButton"
                xCtrl.Value = False
        End Select
    Next xCtrl
End Sub

Public Function ValidateEmptyFields(ByVal SourceFRM As UserForm) As Boolean
    Dim Field As MSForms.control

    For Each Field In SourceFRM.Controls
        Select Case VBA.TypeName(Field)
            Case "TextBox", "ComboBox"
                If Field.Value = vbNullString Then
                    Field.SetFocus
                    ValidateEmptyFields = True
                    Exit Function
                End If
        End Select
    Next Field
    
End Function

Public Function FilterArrayWithDate(ByVal mtz, _
                                    ByVal iCol As Integer)
                                    
'------------------------------------------------------
'RotineType: Function / Variant - Array
'Criacao: Ivanildo Junior
'Criada em: 10/03/2018 / 19:41
'Objetivo: Filtrar uma coluna de data de uma matriz com os dados da data atual
'Aplicacaoo: FilterArrayWithDate(YourArray, 4)
'------------------------------------------------------
                          
    Dim mtzResult   As Variant
    Dim index       As Long
    Dim RowCounter  As Long
    Dim ColCounter  As Integer
    Dim mtzSize     As Long
    Dim mtzValue    As Date
    Dim ValueDate   As Date
    Dim MYDate      As String
    
    For index = LBound(mtz, 1) To UBound(mtz, 1)
        On Error Resume Next
        mtzValue = mtz(index, iCol)
        On Error GoTo 0
        ValueDate = DateSerial(Year(mtzValue), Month(mtzValue), Day(mtzValue))
        MYDate = Year(ValueDate) & Month(ValueDate)
        If MYDate = Year(Date) & Month(Date) Then mtzSize = mtzSize + 1
    Next index
    
    mtzValue = 0
    ValueDate = 0
    MYDate = ""
    
    ReDim mtzResult(1 To mtzSize + 1, 1 To UBound(mtz, 2)) As String
    
    mtzResult(1, 1) = "IDLancto"
    mtzResult(1, 2) = "Data Lancto"
    mtzResult(1, 3) = "Lançamento"
    mtzResult(1, 4) = "TipoPagamento"
    mtzResult(1, 5) = "Descrição"
    mtzResult(1, 6) = "LocalVenda"
    mtzResult(1, 7) = "Valor/Venda"
    mtzResult(1, 8) = "PreçoUN"
    mtzResult(1, 9) = "CustoKG"
    mtzResult(1, 10) = "Perdidos"
    
    For index = LBound(mtz, 1) To UBound(mtz, 1)
        On Error Resume Next
        mtzValue = mtz(index, iCol)
        On Error GoTo 0
        ValueDate = DateSerial(Year(mtzValue), _
                    Month(mtzValue), Day(mtzValue))
        MYDate = Year(ValueDate) & Month(ValueDate)
        If MYDate = Year(Date) & Month(Date) Then
            RowCounter = RowCounter + 1
            mtzResult(RowCounter + 1, 1) = index - 1
            For ColCounter = 2 To UBound(mtzResult, 2)
                mtzResult(RowCounter + 1, ColCounter) = mtz(index, ColCounter)
            Next ColCounter
        End If
    Next index
    
    FilterArrayWithDate = mtzResult
    
End Function

Public Function FormatColumnsInArray(ByVal mtz, _
                                     ByVal sFormat As String, _
                                     ParamArray Cols() As Variant)
'------------------------------------------------------
'RotineType: Function / Variant - Array
'Criacao: Ivanildo Junior
'Criada em: 10/03/2018 / 19:41
'Objetivo: Formatar uma ou mais colunas de uma matriz de acordo com o formato passado como parâmetro
'Aplicacaoo: FormatColumnsInArray(YourArray,"Currency", 2,3,9,5)
'------------------------------------------------------
    Dim RowCounter  As Long
    Dim ColCounter  As Long
    Dim mtzResult   As Variant: mtzResult = mtz
    Dim ParamCounter As Integer

    For RowCounter = LBound(mtzResult, 1) To UBound(mtzResult, 1)
        For ColCounter = LBound(mtzResult, 2) To UBound(mtzResult, 2)
                For ParamCounter = 0 To UBound(Cols)
                    If Cols(ParamCounter) = ColCounter Then
                        mtzResult(RowCounter, ColCounter) = VBA.Format(mtzResult(RowCounter, ColCounter), sFormat)
                    End If
                Next ParamCounter
        Next ColCounter
    Next RowCounter
    
    FormatColumnsInArray = mtzResult
End Function

Public Function CostValues() As String
    Dim ws      As Worksheet: Set ws = shCaixa
    Dim lo      As ListObject: Set lo = ws.ListObjects("fCaixa")
    Dim iCount  As Integer
    Dim vSum    As Currency
    Dim vDate   As Date
    Dim iDate   As Integer: iDate = lo.ListColumns("DataLançamento").index
    Dim iCusto  As Integer: iCusto = lo.ListColumns("Custofinal").index
    Dim iLanc   As Integer: iLanc = lo.ListColumns("Lançamento").index
    Dim vCusto  As Currency
    Dim vLanc   As String
    Dim vSumKG  As Double
    Dim vKG     As Double
    Dim iKG     As Integer: iKG = lo.ListColumns("Qtd/Kg Total").index
    
    For iCount = 1 To lo.ListRows.Count
        With lo
            vDate = .DataBodyRange(iCount, iDate).Value2
            vCusto = .DataBodyRange(iCount, iCusto).Value2
            vLanc = .DataBodyRange(iCount, iLanc).Value2
            vKG = .DataBodyRange(iCount, iKG).Value2
        End With
        If vDate = DateValue(frmlançamentos.txtData.Value) And vLanc = "RECEITA" Then
            vSum = vSum + vCusto
            vSumKG = vSumKG + vKG
        End If
    Next iCount
    
    CostValues = "Quantidade de KG: " & VBA.FormatNumber(vSumKG, 3) & vbNewLine & " " & _
                 "Custo Total: " & VBA.Format(vSum, "Currency")
End Function

Function PreçoCusto() As String
    Dim ws      As Worksheet
    Dim lo      As ListObject
    Dim vDate   As String
    Dim counter As Integer
    Dim pID     As Double
    
    Set ws = shPedidos
    Set lo = ws.ListObjects(1)
    vDate = frmlançamentos.txtData.Value
    
    pID = Application.WorksheetFunction.Match(DateValue(vDate) * 1, lo.ListColumns(1).DataBodyRange.Value2, 1)
     
    PreçoCusto = VBA.FormatCurrency(lo.DataBodyRange(pID, lo.ListColumns("Preço/KG").index).Value2, 2)
End Function

