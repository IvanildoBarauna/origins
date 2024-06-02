Attribute VB_Name = "mdMain"
Option Explicit
Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" _
                                                     (ByRef lpdwFlags As Long, _
                                                      ByVal lpszConnectionName As String, _
                                                      ByVal dwNameLen As Integer, _
                                                      ByVal dwReserved As Long) _
                                                      As Long
 
Dim sConnType As String * 255

Function isConected() As String
    Dim Ret As Long: Ret = InternetGetConnectedStateEx(Ret, sConnType, 254, 0)
    If Ret = 1 Then isConected = sConnType
End Function

Public Sub ErrRaise()
    On Error GoTo err
    Dim ws As Excel.Worksheet
    
    Set ws = ThisWorkbook.Charts(1)
    Exit Sub
err:
    If isConected <> "" Then
        ErrLib.SendMail ErrLib.ErrInfoToMail("mdMain", "ErrRaise", "Sem Comentários")
        VBA.MsgBox "Erro não tratado, um e-mail foi enviado ao administrador do projeto com todas as informações necessárias.", vbCritical
    Else
        VBA.MsgBox "Ocorreu um erro não esperado, capture esta tela (foto ou print screen) e informe o administrador." & vbNewLine & _
                    ErrLib.ErrInfoToMail("mdMain", "ErrRaise"), vbCritical
    End If
End Sub
