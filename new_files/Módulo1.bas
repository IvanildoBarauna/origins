Attribute VB_Name = "Módulo1"
Declare Function DisplaySize Lib "user32" Alias "GetSystemMetrics" (ByVal nIndex As Long) As Long


'Identifica a resolução do sistema operacional
Function VideoRes() As String
    Dim vidWidth
    Dim vidHeight

    'Chamada da API do Windows que identifica a resolução da tela
    vidWidth = DisplaySize(0)
    vidHeight = DisplaySize(1)

    Select Case (vidWidth * vidHeight)
        Case 307200
            VideoRes = "640 x 480"
        Case 480000
            VideoRes = "800 x 600"
        Case 786432
            VideoRes = "1024 x 768"
        Case 1024000
            VideoRes = "1280 x 800"
        Case Else
            VideoRes = "Outra resolução"
    End Select
End Function

'Sub que faz uso da function que identifica a resolução para determinar se ela é adequada
Public Sub CheckDisplayRes()
    Dim VideoInfo As String
    Dim Msg1 As String, Msg2 As String, Msg3 As String

    VideoInfo = VideoRes

    Msg1 = "A resolução atual está configurada em " & VideoInfo & Chr(10)
    Msg2 = "A melhor resolução para essa aplicação é 1024 x 768" & Chr(10)
    Msg3 = "Ajuste a resolução"

    Select Case VideoInfo
        Case "640 x 480"
            MsgBox Msg1 & Msg2 & Msg3
        Case "800 x 600"
            MsgBox Msg1 & Msg2
        Case "1024 x 768"
            MsgBox Msg1
        Case "1280 x 800"
            MsgBox Msg1 & Msg2
        Case Else
            MsgBox Msg2 & Msg3
    End Select
End Sub
