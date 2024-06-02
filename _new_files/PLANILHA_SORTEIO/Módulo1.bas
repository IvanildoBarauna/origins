Attribute VB_Name = "Módulo1"
Public Sub AleatorioEntreFixo()
        Dim lUltimaLinhaAtiva As Long
              Application.Volatile
              lUltimaLinhaAtiva = Worksheets("Lista").Cells(Worksheets("Lista").Rows.Count, 1).End(xlUp).Row
        For i = 1 To 100
              Range("G7").FormulaR1C1 = "=VLOOKUP(RANDBETWEEN(1," & lUltimaLinhaAtiva & "),Lista!C[-6]:C[-5],2,0)"
        Next i
        Range("G7").Select
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
               :=False, Transpose:=False
        Application.CutCopyMode = False
End Sub
