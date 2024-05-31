Attribute VB_Name = "Chamada"
Option Explicit

Sub chamar_macros()
    If wsFormulario.Range("vData").Value2 <> "" Then
        Call Gravar_InfoColab
        Call Gravar_InfoProduto
        Call Gravar_InfoMP
        Call Gravar_InfoParada
        Call Gravar_InfoPerdas
        Call Gravar_InfoTorque
        Call LimparDados
    Else
        MsgBox "Há campos vazios no formulário, favor preenchê-los.", vbExclamation
    End If
End Sub

Sub LimparDados(): wsFormulario.Range("rngDados").ClearContents: End Sub
