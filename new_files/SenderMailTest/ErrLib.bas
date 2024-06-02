Attribute VB_Name = "ErrLib"
Option Explicit
#Const DESIGN_MODE = False

Sub SendMail(BodyMessage As String)
    Const SMTPSERVER        As String = "smtp.gmail.com"
    Const PORT              As Integer = 465
    Const USERNAME          As String = "juniorieq6@gmail.com"
    Const PASS              As String = "#P@ssw0rd9"
    Const DESTINATÁRIO      As String = "ivanildo.jnr@outlook.com"
    
    Const schema  As String = "http://schemas.microsoft.com/cdo/configuration/"
#If Not DESIGN_MODE Then
    Dim oMessage As Object, oConf As Object
    Set oMessage = VBA.Interaction.CreateObject("CDO.Message")
    Set oConf = VBA.Interaction.CreateObject("CDO.Configuration")
#Else
    Dim oMessage As CDO.Message, oConf As CDO.Configuration
    Set oMessage = New CDO.Message
    Set oConf = New CDO.Configuration
#End If
    
    With oConf.Fields
        .Item(schema & "sendusing") = 2
        .Item(schema & "smtpserver") = SMTPSERVER
        .Item(schema & "smtpserverport") = PORT
        .Item(schema & "smtpauthenticate") = 1
        .Item(schema & "sendusername") = USERNAME
        .Item(schema & "sendpassword") = PASS
        .Item(schema & "smtpusessl") = 1
        .Update
    End With
    
    With oMessage
        .To = DESTINATÁRIO
        .From = USERNAME
        .Subject = "[IMPORTANTE] Notificação de Erro em Projeto VBA" & ThisWorkbook.Name
        .TextBody = BodyMessage
        .AddAttachment
        Set .Configuration = oConf
        .Send
    End With
End Sub


Public Function ErrInfoToMail(ModuleName As String, RotineName As String, Optional Comment As String = "N/A") As String
        ErrInfoToMail = "DESCRIÇÃO DO ERRO: " & VBA.err.Description & vbNewLine & _
                        "NÚMERO DO ERRO: " & VBA.err.Number & vbNewLine & _
                        "NOME MÓDULO: " & ModuleName & vbNewLine & _
                        "NOME ROTINA: " & RotineName & vbNewLine & _
                        "PATH WBK: " & ThisWorkbook.FullName & vbNewLine & _
                        "NOME DO COMPUTADOR: " & VBA.Environ("computername") & vbNewLine & _
                        "NOME DO USUÁRIO: " & VBA.Environ("username") & vbNewLine & _
                        "NOME PLANILHA ATIVA: " & ThisWorkbook.ActiveSheet.Name & vbNewLine & _
                        "CODE NAME PLANILHA ATIVA: " & ThisWorkbook.ActiveSheet.CodeName & vbNewLine & _
                        "DATA E HORA: " & VBA.UCase(VBA.Format(VBA.Now, "DDDD DD/MM/YYYY HH:MM:SS")) & vbNewLine & _
                        "COMENTÁRIO: " & Comment
End Function
