Attribute VB_Name = "mdMain"
Function DADOS_PARA_STRING(ByRef arr, _
                           ByVal Delimiter As String) As String
    DADOS_PARA_STRING = VBA.Join(WorksheetFunction.Transpose(arr), Delimiter)
End Function

