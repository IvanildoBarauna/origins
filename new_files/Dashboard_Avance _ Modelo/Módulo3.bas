Attribute VB_Name = "Módulo3"
Sub limparparadas()
Attribute limparparadas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' limparparadas Macro
'

'
    ActiveWindow.SmallScroll Down:=3
    Range("A27:A41").Select
    Selection.FormulaR1C1 = "FALSE"
    ActiveWindow.SmallScroll Down:=-3
    Range("A27").Select
End Sub
Sub marcartudopardas()
Attribute marcartudopardas.VB_ProcData.VB_Invoke_Func = " \n14"
'
' marcartudopardas Macro
'

'
    Range("A27").Select
    ActiveWindow.SmallScroll Down:=9
    Range("A27:A41").Select
    Selection.FormulaR1C1 = "TRUE"
    ActiveWindow.SmallScroll Down:=-6
    Range("A26").Select
End Sub
