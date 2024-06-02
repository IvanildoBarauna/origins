Attribute VB_Name = "Principal"
Option Explicit

' Exibe o formulário de execução de comandos SQL
Sub Exibir_Formulário()
Attribute Exibir_Formulário.VB_ProcData.VB_Invoke_Func = "S\n14"
    frmSQL.Show vbModeless
End Sub

' Função que extrai o caminho a partir de um nome completo de arquivo
Function Directory(p As String) As String
    Directory = Left(p, InStrRev(p, "\"))
End Function

' Função que remove excesso de linhas em branco na sintaxe SQL
Function RemoveBlankLines(ByVal Texto As String) As String
    Do While InStr(Texto, vbCrLf & vbCrLf) > 0
        Texto = Replace(Texto, vbCrLf & vbCrLf, vbCrLf)
    Loop
    If Right(Texto, 2) = vbCrLf Then Texto = Left(Texto, Len(Texto) - 2)
    RemoveBlankLines = Texto
End Function

' Função que remove linhas em branco no final da sintaxe e adiciona ponto-e-vírgula
Function AddSemiColon(ByVal Texto As String) As String
    Do While Right(Texto, 2) = vbCrLf
        Texto = Left(Texto, Len(Texto) - 2)
    Loop
    If Right(Texto, 1) <> ";" Then Texto = RTrim(Texto) & ";"
    AddSemiColon = Texto
End Function
