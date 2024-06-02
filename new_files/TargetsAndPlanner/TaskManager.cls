Attribute VB_Name = "TaskManager"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Dim mLO             As ListObject
Dim mDataInicio     As Date
Dim mDataLimite     As Date
Dim mObjetivo       As String
Dim mValor          As Integer
Dim mQuantidade     As Integer
Dim mStatus         As String
Dim mDataConclusao  As String

Public Enum TypeSave
    NovaTarefa = 0
    AlterarTarefa = 1
End Enum

Private Sub Class_Initialize()
    Set mLO = sMain.ListObjects("tbMain")
End Sub

Private Sub Class_Terminate()
    Set mLO = Nothing
End Sub

Public Property Let aDataInicial(sDataIni As String): mDataInicio = sDataIni: End Property
Public Property Let bDataLimite(sDataLim As String): mDataLimite = sDataLim: End Property
Public Property Let cObjetivo(sObjetivo As String): mObjetivo = sObjetivo: End Property
Public Property Let dValor(iValor As Integer): mValor = iValor: End Property
Public Property Let eQuantidade(sQTD As String): mQuantidade = sQTD: End Property
Public Property Let fstatus(strStatus As String): mStatus = strStatus: End Property
Public Property Let gDataConclusao(sDataConclusao As String): mDataConclusao = sDataConclusao: End Property

Public Sub Save(ByVal Modo As TypeSave)
    Dim lIndex As Integer
    
    With mLO
        If Modo = NovaTarefa Then
            lIndex = .ListRows.Add.Index
            .ListRows.Item(lIndex).Range.Value2 = ArrToListObject
        Else
            lIndex = frmPrincipal.lstDados.ListIndex + 1
            .ListRows.Item(lIndex).Range.Value2 = ArrToListObject
        End If
    End With
    
    mLO.Range.Columns.AutoFit
End Sub

Private Function ArrToListObject()
    Dim tmpArr(1 To 1, 1 To 7)
    
    With mLO
        tmpArr(1, .ListColumns("DataInício").Index) = mDataInicio * 1
        tmpArr(1, .ListColumns("DataVencimento").Index) = mDataLimite * 1
        tmpArr(1, .ListColumns("Descrição").Index) = mObjetivo
        tmpArr(1, .ListColumns("Valor").Index) = mValor
        tmpArr(1, .ListColumns("QuantidadeDias").Index) = mQuantidade
        tmpArr(1, .ListColumns("Status").Index) = mStatus
        tmpArr(1, .ListColumns("DataConclusão").Index) = mDataConclusao
    End With

    ArrToListObject = tmpArr
    
    Erase tmpArr
End Function
