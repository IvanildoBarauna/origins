Attribute VB_Name = "mdT"
Option Explicit
Sub AbastecerEstoque()
      Dim qtd, est, linha, total   As Long
      Dim lo                               As ListObject
      
      Set lo = shtESTOQUE.ListObjects("tbESTOQUE")
      
      With lo
            
                  If shtIN.txtcod = Empty Or shtIN.txtdesc = Empty Then
                        MsgBox "O campo código do produto está vazio, ou digite um valor válido" _
                              , vbExclamation, "Dados inválidos"
                        Exit Sub
                  Else
                        linha = shtIN.txtcod + 1
                        qtd = shtIN.txtqtd.Value * 1
                        est = .Range(linha, lo.ListColumns("QUANTIDADE").Index)
                        total = est + qtd
                        
                        .Range(linha, lo.ListColumns("QUANTIDADE").Index) = total
                        
                        Registrar "ENTRADA"
                        
                        MsgBox "Entrada registrada com sucesso!" _
                              , vbInformation, "ENTRADA DE ESTOQUE"
                              
'                              shtIN.txtcod = Empty
'                              shtIN.txtqtd = Empty
'
'                              qtd = Empty
'                              est = Empty
'                              linha = Empty
'                              total = Empty
            End If
      End With
End Sub

Sub SaídaEstoque()
      Dim qtd, est, linha, total As Long
      Dim lo As ListObject
      Set lo = shtESTOQUE.ListObjects("tbESTOQUE")
      
      With lo
            
                  If shtOUT.txtcod = Empty Or shtOUT.txtdesc = Empty Then
                        MsgBox "O campo código do produto está vazio, ou digite um valor válido" _
                              , vbExclamation, "Dados inválidos"
                        Exit Sub
                  Else
                        linha = shtOUT.txtcod + 1
                        qtd = shtOUT.txtqtd.Value * 1
                        est = .Range(linha, lo.ListColumns("QUANTIDADE").Index)
                        total = est - qtd
                        
                        .Range(linha, lo.ListColumns("QUANTIDADE").Index) = total
                        
                        Registrar "SAÍDA"
                        
                        MsgBox "Saída Registrada com Sucesso!", vbInformation, "SAÍDA DE ESTOQUE"
                              
'                              shtOUT.txtcod = Empty
'                              shtOUT.txtqtd = Empty
'
'                              qtd = Empty
'                              est = Empty
'                              linha = Empty
'                              total = Empty
            End If
      End With
End Sub

Sub AcrescentarProduto()
      Dim linha As Integer
      
      linha = shtESTOQUE.Range("A6").End(xlDown)
      
      shtADD.txtcod = linha
      
End Sub
