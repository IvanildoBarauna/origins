Attribute VB_Name = "Módulo1"
Sub AtualizaMapa()
Attribute AtualizaMapa.VB_ProcData.VB_Invoke_Func = " \n14"

    Dim estado As Range
    
    For Each estado In Range("ESTADOS")
    
            cor = Cells(estado.Row, 6)
            ActiveSheet.Shapes(estado).Fill.ForeColor.RGB = Range(cor).Interior.Color
    Next estado
    
    
End Sub

Sub AtualizaMapa2()

    LimpaMapa
    
    estados = Split(ActiveSheet.Range("estados_empresa").Value, ",")
    cor = ActiveSheet.Range("cor_estados").Interior.Color
       
    For i = LBound(estados) To UBound(estados)
        
        ActiveSheet.Shapes(Trim(estados(i))).Fill.ForeColor.RGB = cor
    
    Next i
        
    
End Sub


Sub LimpaMapa()

    cor = ActiveSheet.Range("sem_cor").Interior.Color
    ActiveSheet.Shapes("AC").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("AL").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("AM").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("AP").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("BA").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("CE").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("DF").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("ES").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("GO").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("MA").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("MG").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("MS").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("MT").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("PA").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("PB").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("PE").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("PI").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("PR").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("RJ").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("RN").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("RO").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("RR").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("RS").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("SC").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("SE").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("SP").Fill.ForeColor.RGB = cor
    ActiveSheet.Shapes("TO").Fill.ForeColor.RGB = cor

End Sub

