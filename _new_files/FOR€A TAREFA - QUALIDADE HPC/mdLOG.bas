Attribute VB_Name = "mdLOG"
Option Explicit
Public Sub LOG(Acao As String)

shtLOG.Unprotect ["#P@ssw0rd1"]

Dim lo As ListObject
Dim lr As ListRow
Dim USER As String
Dim Computer As String

Set lo = shtLOG.ListObjects("tbLOG")
Set lr = lo.ListRows.Add

USER = Environ("USERNAME")
Computer = Environ("COMPUTERNAME")


With lr

Application.StatusBar = "Aguarde ... Registrando Atividade"

.Range(lo.ListColumns("DATA/HORA").Index) = Now
.Range(lo.ListColumns("LOGIN").Index) = USER
.Range(lo.ListColumns("COMPUTADOR").Index) = UCase(Computer)
.Range(lo.ListColumns("AÇÃO").Index) = UCase(Acao)

End With

shtLOG.Protect ["#P@ssw0rd1"]

Application.StatusBar = False


End Sub

Sub test()

Dim Row As Range

shtLOG.Range("a1").Select
Row = shtLOG.Range("A" & Rows.Count).End(xlUp).Row + 1


End Sub



