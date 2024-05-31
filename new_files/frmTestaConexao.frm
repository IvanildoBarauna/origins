Attribute VB_Name = "frmTestaConexao"
Attribute VB_Base = "0{31BA2B8C-1858-41DB-A860-7F078BD14E09}{235C47FD-2731-4E5E-94C2-31496DFEB454}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Function sPing(sHost) As ping
    Dim oPing As Object, oRetStatus As Object

    Set oPing = GetObject("winmgmts:{impersonationLevel=impersonate}").ExecQuery _
      ("select * from Win32_PingStatus where address = '" & sHost & "'")
 
    For Each oRetStatus In oPing
        If IsNull(oRetStatus.StatusCode) Or oRetStatus.StatusCode <> 0 Then
            sPing.descricao = ""
            sPing.bufferTime = 0
        Else
            sPing.descricao = "Sucesso"
            sPing.bufferSize = oRetStatus.bufferSize
            sPing.bufferTime = oRetStatus.ResponseTime
            sPing.TTL = oRetStatus.ResponseTimeToLive
        End If
    Next
End Function

Private Sub cmdReTestar_Click()
    testaRede
End Sub

Private Sub cmdSair_Click()
    Unload Me
End Sub

Private Sub testaRede()
    Dim iRow        As Long
    Dim lRow        As Long
    Dim dados       As ping
    Dim qtdErros    As Integer
    Dim lo          As Excel.ListObject
    Dim IP          As String
    Dim svrType     As String
    
    Set lo = wsIPS.ListObjects(1)
    lo.ListColumns(3).DataBodyRange.Value2 = ""
    
    With Me.lblTexto
        For iRow = 1 To lo.ListRows.Count
            IP = lo.DataBodyRange(iRow, 1).Value2
            svrType = lo.DataBodyRange(iRow, 2).Value2
            .Caption = "Testando conexão: " & IP & " - " & svrType
            VBA.DoEvents
            mdMain.Sleep 500
            dados = sPing(IP)
            If Not VBA.Trim(dados.descricao) = "" Then
                lo.DataBodyRange(iRow, 3).Value2 = dados.descricao & " - " & dados.bufferTime & "ms"
                lo.DataBodyRange(iRow, 3).Font.Color = vbGreen
                .ForeColor = vbBlack
            Else
                lo.DataBodyRange(iRow, 3).Value2 = "Erro"
                lo.DataBodyRange(iRow, 3).Font.Color = vbRed
                lblTexto.ForeColor = vbRed
                mdMain.Sleep 1000
            End If
        Next iRow
        
        qtdErros = Application.WorksheetFunction.CountIf(lo.ListColumns(3).DataBodyRange, "Erro")
        .Caption = "IPS TESTADOS COM SUCESSO, TOTAL DE: " & qtdErros & " erros."
    End With
    
End Sub

