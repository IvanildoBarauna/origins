Attribute VB_Name = "mdLOG_tbl"
Option Explicit
Public Sub LOG_tbl(Ação As String)

Dim lo As ListObject
Dim lr As ListRow
Dim USER As String
Dim Computer As String

Set lo = shtLOG_tbl.ListObjects("tb_LOG")
Set lr = lo.ListRows.Add

USER = Environ("USERNAME")
Computer = Environ("COMPUTERNAME")

With lr

     .Range(lo.ListColumns("DATA / HORA INT.").Index) = Now
     .Range(lo.ListColumns("LOGIN").Index) = USER
     .Range(lo.ListColumns("COMPUTER").Index) = UCase(Computer)
     .Range(lo.ListColumns("AÇÃO").Index) = UCase(Ação)
        
End With

shtLOG_tbl.Columns.EntireColumn.AutoFit

End Sub

