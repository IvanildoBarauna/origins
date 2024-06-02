Attribute VB_Name = "Módulo1"
Rem: | Rotine By Giovani Franco - Grupo Resenha VBA / Exemplo Select Case com Intervalos
Sub Mycase()
    Dim x As Variant
    Do
        x = Application.InputBox("Informe um número qualquer, ou clique em cancelar para sair.")
        If x = "" Or x = 0 Then GoTo saida
        If Not VBA.IsNumeric(x) Then GoTo saida
        
        Select Case x
            Case 1 To 9
                MsgBox "Número pertencente ás unidades"
            Case 10 To 99
                MsgBox "Número pertencente às dezenas"
            Case 100 To 999
                MsgBox "Número pertencente às centenas"
            Case 1000 To 9999
                MsgBox "Número pertencente às milhares"
            Case 10000 To 99999
                MsgBox "Número pertencente às dezenas de milhares"
            Case Else
                MsgBox "Número muito grande para catalogar"
        End Select
    Loop While x <> Empty
saida:
End Sub
