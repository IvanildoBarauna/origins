Attribute VB_Name = "cAgent"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private Type ClassType
    AgentName  As String
    sFunc      As String
End Type

Private This As ClassType

Public Property Let NomeAgente(sNome As String): This.AgentName = sNome: End Property
Public Property Let Funcional(sFuncional As String): This.sFunc = sFuncional: End Property

Sub SaveORChangeReg(Optional RowIndex As Integer = 0)
    Dim lo As Excel.ListObject
    Dim lr As Excel.ListRow
    
    Set lo = wsListaAgents.ListObjects(1)
    If RowIndex > 0 Then Set lr = lo.ListRows(RowIndex) Else Set lr = lo.ListRows.Add
    
    With lr
        .Range(lo.ListColumns("FUNCIONAL").index).Value2 = This.sFunc
        .Range(lo.ListColumns("NOME").index).Value2 = This.AgentName
    End With
    
    Call ClassificarTabela(lo, lo.ListColumns("NOME").index)
End Sub

Sub ClassificarTabela(lo As ListObject, iCol As Integer)
    With lo.Sort
        .SortFields.Clear
        .SortFields.Add lo.ListColumns(iCol).Range, xlSortOnValues, xlAscending, xlSortNormal
        .Apply
    End With
End Sub
