Attribute VB_Name = "Módulo9"
Sub Macro2()
Attribute Macro2.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro2 Macro
'

'
    Range("E3:E9").Select
    ActiveWorkbook.Worksheets("ENTRADAS").ListObjects("Tabela2").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("ENTRADAS").ListObjects("Tabela2").Sort.SortFields. _
        Add Key:=Range("E3"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("ENTRADAS").ListObjects("Tabela2").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("E2").Select
End Sub
Sub Macro4()
Attribute Macro4.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro4 Macro
'

'
    Range("W9:AF15").Select
    ActiveWorkbook.Worksheets("RELATÓRIO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("RELATÓRIO").Sort.SortFields.Add Key:=Range("W9"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("RELATÓRIO").Sort
        .SetRange Range("W9:AF15")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub
Sub Macro5()
Attribute Macro5.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro5 Macro
'

'
    
End Sub
Sub Macro6()
Attribute Macro6.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro6 Macro
'

'
   
End Sub
