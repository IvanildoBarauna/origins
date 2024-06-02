Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub SendMails()
    Const cName     As String = "[NOME DO CANDIDATO]"
    Const cAux      As String = "[DATA DO PROCESSO SELETIVO]"
    Const sMsg      As String = "[SAUDACAO]"
    Dim vDate       As Date
    Dim iRow        As Long, LastRow As Long
    Dim ws          As Worksheet
    Dim OutApp      As Object
    Dim MailItem    As Object
    Dim sCandidato  As String
    Dim sBody       As String
    Dim msg         As String
    
    If MsgBox("[AVISO] Por favor verifique se seu outlook está aberto para que seja continuado o processo." & vbNewLine & "Deseja continuar?", vbExclamation + vbYesNo) = vbYes Then
    
        On Error GoTo errDate
        vDate = VBA.InputBox("Digite a data do processo seletivo:")
        GoTo continue
errDate:
        MsgBox "Digite uma data válida.", vbExclamation
        Exit Sub
continue:
        On Error GoTo err
        
        Set OutApp = VBA.CreateObject("Outlook.Application")
        Set ws = wsCandidatos
        LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        
        Excel.Application.ScreenUpdating = False
        
        msg = strSaudacao()
        
        For iRow = 2 To LastRow
            Set MailItem = OutApp.CreateItem(0)
            sCandidato = ws.Cells(iRow, 1).Value2
            sBody = VBA.Replace(VBA.Replace(VBA.Replace(wsCorpo.Range("A2").Value2, cName, sCandidato), cAux, vDate), sMsg, msg)
            With MailItem
                .To = ws.Cells(iRow, 2).Value2
                .Subject = wsAssunto.Range("A2").Value
                .Display
                .HTMLBody = sBody & .HTMLBody
                .Send
            End With
            Set MailItem = Nothing
        Next iRow
            
        MsgBox "E-mails enviados com sucesso!", vbInformation
        Set OutApp = Nothing
        Excel.Application.ScreenUpdating = True
        Exit Sub
err:
        Excel.Application.ScreenUpdating = True
        MsgBox "Não foi possível concluir o processo." & vbNewLine & err.Description, vbCritical
    Set OutApp = Nothing
    Set MailItem = Nothing
    End If
End Sub

 Function strSaudacao() As String
    If VBA.Hour(Now()) > 0 And VBA.Hour(Now()) < 12 Then
        strSaudacao = "Bom dia"
    ElseIf VBA.Hour(Now()) > 12 And VBA.Hour(Now()) < 18 Then
        strSaudacao = "Boa tarde"
    Else
        strSaudacao = "Boa noite"
    End If
End Function
