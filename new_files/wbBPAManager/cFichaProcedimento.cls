Attribute VB_Name = "cFichaProcedimento"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Type ClassType
    sProcedimento As String
    sProfissional As String
    iQuantidade   As Integer
    sDataInicial  As String
End Type

Private This As ClassType

Public Property Let ProcedimentoNome(Value As String): This.sProcedimento = Value: End Property
Public Property Let ProfissionalNome(Value As String): This.sProfissional = Value: End Property
Public Property Let Quantidade(Value As String): This.iQuantidade = Value: End Property
Public Property Let DataInicial(Value As String): This.sDataInicial = Value: End Property

Public Sub SaveOrChangeData(Optional ByVal RowIndex As Long = 0)
    Dim lo As Excel.ListObject
    Dim lr As Excel.ListRow, arr
    
    arr = ArrToListRow
    Set lo = wsProcedimentos.ListObjects("tbProcedimentos")
    Set lr = lo.ListRows(RowIndex)
    
    With lr
        .Application.Range(.Range(1, 2), .Range(lo.ListColumns("DATA INICIAL").Index)).Value2 = arr
    End With
    ThisWorkbook.Save
End Sub

Private Function ArrToListRow() As Variant
    Dim tmpArr(1 To 1, 2 To 5)
    
    With This
        tmpArr(1, 2) = .sProfissional
        tmpArr(1, 3) = .sProcedimento
        tmpArr(1, 4) = .iQuantidade
        tmpArr(1, 5) = VBA.Replace(.sDataInicial, "/", "")
    End With
    
    ArrToListRow = tmpArr
End Function
