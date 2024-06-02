Attribute VB_Name = "Módulo1"
Option Explicit

Sub inserir()
'
'
'

'======Essa rotina eu escrevi apenas para popular as planilhas com os itens da planilha base===

'Ela não faz parte da macro que faz a busca


Dim pl(1 To 6) As Worksheet
Dim a As Long
Dim i As Long
Dim palavra As String
Dim vezes As Integer
Dim v As Integer
Dim coluna As Integer, linha As Integer


Set pl(1) = Planilha2
Set pl(2) = Planilha3
Set pl(3) = Planilha4
Set pl(4) = Planilha5
Set pl(5) = Planilha6
Set pl(6) = Planilha7


    For i = 5 To 9
    
    palavra = Planilha1.Cells(i, 2)
    
    
        For a = 1 To 6
        
        vezes = Application.WorksheetFunction.RandBetween(2, 5)
        
            For v = 1 To vezes
            
                linha = Application.WorksheetFunction.RandBetween(1, 15)
                coluna = Application.WorksheetFunction.RandBetween(1, 12)
            
                pl(a).Cells(linha, coluna) = palavra
            
            
            Next v
        
        
        Next a
    
    
    Next i

End Sub


Sub buscar()

Dim mbusca() As Variant
Dim ulinha As Long
Dim w As Worksheet, wb As Workbook
Dim palavraBuscada As String
Dim AreaDeBusca As Range, cel As Range
Dim enderecos As String
Dim i As Long

Set wb = ThisWorkbook

ulinha = wsDados.Cells(wsDados.Rows.Count, 2).End(3).Row

mbusca = wsDados.Range("b5:b" & ulinha).Value 'preenche a matriz com os itens à serem procurados

ReDim Preserve mbusca(1 To UBound(mbusca), 1 To 2) 'redimensiona a matriz para uma coluna à mais _
que é nessa coluna que serão colocados os endereços das palavras nas diversas planilhas


    For i = 1 To UBound(mbusca, 1) 'laço que percorrerá cada uma das palavras na matriz
    palavraBuscada = VBA.UCase(mbusca(i, 1)) 'atribui o item da matriz à variável
    enderecos = "" 'limpa a variável que armazenará os endereços encontrados
        For Each w In wb.Worksheets 'laço para percorrer todas as planilhas do livro
        
            If VBA.UCase(w.CodeName) <> VBA.UCase(wsDados.CodeName) Then 'para evitar a redundância de encontrar a palavra na planilha de busca
            
            Set AreaDeBusca = w.UsedRange 'setando a área de busca
            
                    For Each cel In AreaDeBusca.Cells 'percorrendo a área de busca à procura da string
                    
                        If VBA.UCase(cel.Value) = palavraBuscada Then 'se encontrar, armazena na variável de endereços
                        
                            enderecos = enderecos & w.Name & "!" & cel.Address & ";"
                        
                        End If
                    
                    
                    Next cel
            
            End If
        
        
        Next w
    
    
    enderecos = VBA.Left(enderecos, VBA.Len(enderecos) - 1) 'tirando o último separador...
    
    mbusca(i, 2) = enderecos 'preenche a coluna 2 da matriz com os endereços encontrados
    
    Next i


wsDados.Range("b5").Resize(UBound(mbusca), 2).Value = mbusca 'derruba a matriz na planilha de busca

'=======Maquiagem================
wsDados.Columns(3).ColumnWidth = 50
wsDados.Range("c5:c" & ulinha).WrapText = True
wsDados.Rows.AutoFit
wsDados.Range("b5:c" & ulinha).VerticalAlignment = xlCenter

'=======Maquiagem================

Erase mbusca 'game over na matriz

End Sub



