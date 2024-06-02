Attribute VB_Name = "Módulo1"
Sub grafico()

With Worksheets("PAINEL.MES").ChartObjects(1).Chart

    Select Case Worksheets("PAINEL.MES").Range("D2").Value
        Case 1
            .ChartType = xlColumnClustered
        Case 2
            .ChartType = xlLineMarkers
        Case 3
            .ChartType = xlArea
    End Select
End With
End Sub
