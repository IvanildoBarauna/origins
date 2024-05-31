Attribute VB_Name = "cAgentAddress"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Type ClassType
    aFunctional      As Variant
    bAgentName       As String
    cAreaMicroArea   As String
    dAddress         As String
    eBairro          As String
    fCep             As String
    gAditionalDetail As String
End Type

Dim This As ClassType

Property Let aFuncional(Value As Variant): This.aFunctional = Value: End Property
Property Let bNomeAgente(Value As String): This.bAgentName = Value: End Property
Property Let cAreaMicroArea(Value As String): This.cAreaMicroArea = Value: End Property
Property Let dEndereco(Value As String): This.dAddress = Value: End Property
Property Let eBairro(Value As String): This.eBairro = Value: End Property
Property Let fCep(Value As String): This.fCep = Value: End Property
Property Let gDetalheAdicional(Value As String): This.gAditionalDetail = Value: End Property

Sub SaveOrChangeData(Optional RowIndex As Long = 0)
    On Error GoTo err
    Dim lo  As Excel.ListObject
    Dim lr  As Excel.ListRow
    Dim arr As Variant
    
    arr = ArrToListRow()
    Set lo = wsRuasAgents.ListObjects(1)
    
    If RowIndex > 0 Then Set lr = lo.ListRows(RowIndex) Else Set lr = lo.ListRows.Add
    
    With lr
        .Application.Range(.Range(1, 1), .Range(1, lo.ListColumns.Count)).Value = arr
    End With
    
    Exit Sub
err:
    Debug.Print "Não foi possível salvar os dados no ListObject. , " - " & err.Description"
End Sub

Private Function ArrToListRow()
    Dim Values(1 To 7)
    
    With This
        Values(1) = CStr(.aFunctional)
        Values(2) = .bAgentName
        Values(3) = .cAreaMicroArea
        Values(4) = .dAddress
        Values(5) = .eBairro
        Values(6) = .fCep
        Values(7) = .gAditionalDetail
    End With
    
    ArrToListRow = Values()
End Function
