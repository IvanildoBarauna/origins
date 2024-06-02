Attribute VB_Name = "mdLOG"
Option Explicit
Sub Registrar(str As String)
      Dim ws           As Worksheet
      Dim lo            As ListObject
      Dim lr             As ListRow
      Dim linha       As Long
      Dim usr          As String
      Dim cp           As String
      
      Set ws = shtLOG
      Set lo = ws.ListObjects("tbLOG")
      Set lr = lo.ListRows.Add
      usr = Environ("USERNAME")
      cp = Environ("computername")
      
      With lr
      
            If str = "SAÍDA" Then
                  .Range(lo.ListColumns("DATA / HORA MOVIMENTAÇÃO").Index) = Now
                  .Range(lo.ListColumns("USUÁRIO").Index) = usr
                  .Range(lo.ListColumns("COMPUTADOR").Index) = cp
                  .Range(lo.ListColumns("PRODUTO").Index) = shtOUT.txtdesc
                  .Range(lo.ListColumns("ESTOQUE ANTERIOR").Index) = shtOUT.txtestoque * 1 + shtOUT.txtqtd * 1
                  .Range(lo.ListColumns("ENTRADA/SAÍDA").Index) = shtOUT.txtqtd * 1
                  .Range(lo.ListColumns("ESTOQUE ATUAL").Index) = shtOUT.txtestoque * 1
                  .Range(lo.ListColumns("GRUPO").Index) = shtOUT.txtgrupo
                  .Range(lo.ListColumns("CATEGORIA").Index) = shtOUT.txtcategoria
                  .Range(lo.ListColumns("TIPO DE MOV.").Index) = str
            Else
                  .Range(lo.ListColumns("DATA / HORA MOVIMENTAÇÃO").Index) = Now
                  .Range(lo.ListColumns("USUÁRIO").Index) = usr
                  .Range(lo.ListColumns("COMPUTADOR").Index) = cp
                  .Range(lo.ListColumns("PRODUTO").Index) = shtIN.txtdesc
                  .Range(lo.ListColumns("ESTOQUE ANTERIOR").Index) = shtIN.txtestoque * 1 - shtIN.txtqtd * 1
                  .Range(lo.ListColumns("ENTRADA/SAÍDA").Index) = shtIN.txtqtd * 1
                  .Range(lo.ListColumns("ESTOQUE ATUAL").Index) = shtIN.txtestoque * 1
                  .Range(lo.ListColumns("GRUPO").Index) = shtIN.txtgrupo
                  .Range(lo.ListColumns("CATEGORIA").Index) = shtIN.txtcategoria
                  .Range(lo.ListColumns("TIPO DE MOV.").Index) = str
            End If
      End With
      ws.Columns.EntireColumn.AutoFit
End Sub
