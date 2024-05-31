Attribute VB_Name = "md"
Option Explicit

Sub teste()

Dim valor As String
Dim resultado As String

resultado = Range("b1").Value

valor = Range("a1").Value

If valor <> "" Then

MsgBox resultado, vbInformation, "Ahh puuufo"

Else

End If

End Sub
