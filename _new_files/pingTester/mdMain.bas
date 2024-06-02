Attribute VB_Name = "mdMain"
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Type ping
    descricao   As String
    bufferSize  As String
    bufferTime  As String
    TTL         As String
End Type

Public Sub executaTeste()
    frmTestaConexao.Show 1
End Sub


