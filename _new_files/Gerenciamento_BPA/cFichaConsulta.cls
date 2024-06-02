Attribute VB_Name = "cFichaConsulta"
Attribute VB_Base = "0{FCFB3D2A-A0FA-1068-A738-08002B3371B5}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Type ClassType
    mProfissional As Variant
    mDataNasc     As Variant
    mDataINI      As Variant
End Type

Private This As ClassType

Public Property Let Profissional(Value As String): This.mProfissional = Value: End Property
Public Property Let DataNascimento(Value As String): This.mDataNasc = Value: End Property
Public Property Let DataInicial(Value As String): This.mDataINI = Value: End Property

Public Sub SaveOrChangeData(Optional ByVal RowIndex As Long = 0)
    Dim lo As Excel.ListObject, lr As Excel.ListRow, arr
    
    arr = ArrToListObject
    Set lo = wsConsultas.ListObjects("tbConsultas")
    
    If RowIndex > 0 Then
        Set lr = lo.ListRows(RowIndex)
    Else
        Set lr = lo.ListRows.Add
    End If

    With lr
        .Application.Range(.Range(1, 2), _
            .Range(1, lo.ListColumns("IDADE").Index)).Value2 = arr
    End With
    ThisWorkbook.Save
End Sub

Private Function ArrToListObject() As Variant
    Dim tmpArr(1 To 1, 2 To 6), lo As ListObject

    Set lo = wsConsultas.ListObjects("tbConsultas")
    
    With This
        tmpArr(1, lo.ListColumns("PROFISSIONAL").Index) = .mProfissional
        tmpArr(1, lo.ListColumns("DATA DE NASC.").Index) = VBA.Replace(.mDataNasc, "/", "")
        tmpArr(1, lo.ListColumns("DATA INICIAL").Index) = VBA.Replace(.mDataINI, "/", "")
        tmpArr(1, lo.ListColumns("IDADE").Index) = VBA.Int((VBA.Date - VBA.DateValue(.mDataNasc)) / 365.25)
    End With
    
    ArrToListObject = tmpArr
End Function
