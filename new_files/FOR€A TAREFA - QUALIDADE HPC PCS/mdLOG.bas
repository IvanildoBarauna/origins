Attribute VB_Name = "mdLOG"
Option Explicit
Option Private Module

Public Sub LOG(Acao As String)

Dim lo As ListObject
Dim lr As ListRow
Dim USER As String
Dim Computer As String

Set lo = shtLOG.ListObjects("tbLOG")
Set lr = lo.ListRows.Add

USER = Environ("USERNAME")
Computer = Environ("COMPUTERNAME")

Application.StatusBar = "Aguarde ... Registrando Atividade"

With lr

     .Range(lo.ListColumns("DATA/HORA").Index) = Now
     .Range(lo.ListColumns("LOGIN").Index) = USER
     .Range(lo.ListColumns("COMPUTADOR").Index) = UCase(Computer)
     .Range(lo.ListColumns("AÇÃO").Index) = UCase(Acao)

End With

Application.StatusBar = False


End Sub

