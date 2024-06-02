Attribute VB_Name = "mLOG_TAB"
Option Explicit

Sub CriarLOG_TAB(Ação As String)
      Dim lo As ListObject
      Dim lr As ListRow
      Dim usr As String
      Dim cp As String
      Set lo = shtLOG_TAB.ListObjects("tLOG")
      Set lr = lo.ListRows.Add
      
      usr = Environ("USERNAME")
      cp = Environ("COMPUTERNAME")
      
      With lr
            .Range(lo.ListColumns("HORA").Index) = Now
            .Range(lo.ListColumns("USUÁRIO").Index) = usr
            .Range(lo.ListColumns("COMPUTADOR").Index) = cp
            .Range(lo.ListColumns("AÇÃO").Index) = Ação
      End With
      
      shtLOG_TAB.Columns.EntireColumn.AutoFit
End Sub
