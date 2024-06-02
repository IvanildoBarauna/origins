Attribute VB_Name = "mLOG_TXT"
Option Explicit

Public Sub CriarLOG_TXT(Ação As String)
      Dim oLOG As cLOG
      Set oLOG = New cLOG
      
      With oLOG
            .fPath = ThisWorkbook.path & Application.PathSeparator & "LOG.txt"
            .Registrar Ação
            .Salvar
      End With
      
      ThisWorkbook.RefreshAll
      shtLOG_TXT.Columns.EntireColumn.AutoFit
End Sub

