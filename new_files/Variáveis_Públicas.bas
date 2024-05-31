Attribute VB_Name = "Variáveis_Públicas"
Option Explicit

Public i As Integer                                     ' Contador genérico
Public SQL As String                                    ' String SQL a ser processada
Public Mens As String                                   ' Mensagens genéricas

Public Erro_Núm As Long                                 ' Número do erro SQL
Public Erro_Msg As String                               ' Texto do erro SQL
Public NúmReg As Long                                   ' Total de registros do recordset

Public FSO As New FileSystemObject                      ' Objeto FileSystem
Public Pasta As Folder                                  ' Objeto tipo pasta
Public Arquivo As File                                  ' Objeto tipo arquivo

Public Const CinzaClaro2 As Long = 15790320             ' Cinza do formulário
Public Const CinzaClaro As Long = 16316664              ' Cinza para botão não selecionado
Public Const CinzaMédio As Long = 15263976              ' Cinza para botão selecionado

