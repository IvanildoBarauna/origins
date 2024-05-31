Attribute VB_Name = "Módulo6"
Sub Copiar_Ordenar()
Attribute Copiar_Ordenar.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Copiar_Ordenar Macro
'


'
    Sheets("ENTRADAS").Select
    Range("D3:O601").Select
    Selection.Copy
    Range("D3").Select
    Sheets("RELATÓRIO_DÍZIMO").Select
    Range("A8").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
      Selection.Font.Size = 9
      
      'Sheets("RELATÓRIO").Range("W8:AF" & Range("AF1048576").End(xlUp).Row)
      
    'Range("W8:AF15").Select
    Range("B8:K" & Range("K1048576").End(xlUp).Row).Select
    ActiveWorkbook.Worksheets("RELATÓRIO_DÍZIMO").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("RELATÓRIO_DÍZIMO").Sort.SortFields.Add Key:=Range("B8"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("RELATÓRIO_DÍZIMO").Sort
        .SetRange Range("B8:K" & Range("K1048576").End(xlUp).Row) 'Range("A8:K")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
      
      
   
    Range("A8").Select
End Sub
Sub Copiar_Saídas()
'
' Copiar_Saídas Macro
'

'
    Sheets("RELATÓRIO_SAÍDAS").Select
    Range("A8:K1000").Select
    Selection.ClearContents
    Range("AG5").Select
    Sheets("SAÍDAS").Select
    Range("D3:N601").Select
      Selection.Copy
    Sheets("RELATÓRIO_SAÍDAS").Select
    Range("A8").Select
    Selection.PasteSpecial Paste:=xlPasteAllUsingSourceTheme, Operation:=xlNone _
        , SkipBlanks:=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    With Selection.Font
        .Name = "Calibri"
        .Size = 9
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    Range("A8").Select
End Sub
