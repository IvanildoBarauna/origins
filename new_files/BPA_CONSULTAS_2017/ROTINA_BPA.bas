Attribute VB_Name = "ROTINA_BPA"

Sub Atualizar_Dados()
Attribute Atualizar_Dados.VB_ProcData.VB_Invoke_Func = " \n14"

    Sheets("DIGITAÇÃO").PivotTables("UPLOAD_BPA").PivotCache.Refresh
    
End Sub

Sub Zerar_Dados()

On Error Resume Next

Range("b6").Select
    Range(Selection, Selection.End(xlDown)).ClearContents
    
    Range("e6").Select
    Range(Selection, Selection.End(xlDown)).ClearContents

Range("b6").Select

Call Atualizar_Dados

MsgBox "DADOS REINICIADOS! INICIE UMA NOVA DIGITAÇÃO."

End Sub



