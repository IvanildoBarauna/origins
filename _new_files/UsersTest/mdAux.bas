Attribute VB_Name = "mdAux"
Option Explicit

Public Function IsValidUserAndPass(ByVal sUser As String, _
                                   ByVal sPass As String) As Boolean
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim SearchRow As Long
    
    Set ws = shUsers
    
    On Error GoTo Err
    SearchRow = ws.Range("B:B").Find(sUser).Row

    If StrConv(ws.Cells(SearchRow, 3).Value2, vbUpperCase) = StrConv(sPass, vbUpperCase) Then
        IsValidUserAndPass = True
        Exit Sub
    End If
   
Err:
    MsgBox "Usuário informado é inválido", vbCritical
End Function


