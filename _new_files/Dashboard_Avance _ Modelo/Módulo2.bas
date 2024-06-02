Attribute VB_Name = "Módulo2"
Sub verificar()
If Folha30.Range("q1") = 1 Then
    MsgBox " Data atualizada até: " & Folha30.Range("q2"), vbInformation, "Alerta de data"
    
     'Application.Speech.Speak ("Base de Dados atualizado até ") & Folha30.Range("q2")
    Exit Sub
    End If
End Sub

Sub Moagem()

Plan31.[M61].Value = Folha5.[L13].Value
Plan31.[R61].Value = Folha30.[W4].Value

End Sub
