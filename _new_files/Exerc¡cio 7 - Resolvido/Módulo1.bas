Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub ExportToPPT()
    Dim ws      As Worksheet
    Dim rng     As Range
    Dim oChart  As ChartObject
    Dim pwrApp  As Object
    Dim pres    As Object
    Dim sld1    As Object
    Dim sld2    As Object
    
    On Error GoTo err
    
    Set ws = wsPrincipal
    Set rng = ws.Range("B4:H11")
    Set oChart = ws.ChartObjects(1)
    
    Set pwrApp = CreateObject("PowerPoint.Application")
    Set pres = pwrApp.Presentations.Add
    Set sld1 = pres.Slides.Add(1, 12) 'ppLayoutBlank
    Set sld2 = pres.Slides.Add(2, 12) 'ppLayoutBlank
    
    rng.CopyPicture xlScreen
    
    With sld1
        .Shapes.Paste.Align msoAlignLefts, msoCTrue
        With .Shapes(1)
            .Height = 8.06
            .Width = 33.87
            .ScaleHeight 1.84, True
            .ScaleWidth 1.84, True
        End With
    End With
    
    oChart.CopyPicture xlScreen
    
    With sld2
        .Shapes.Paste
        With .Shapes(1)
            .Left = 3
            .Top = 2
            .ScaleHeight 2.6, True
        End With
    End With

    pres.SaveAs ThisWorkbook.Path & "\Apresentação de Resultados.pptx"
    pres.Close
    pwrApp.Quit
    MsgBox "Apresentação de Resultados realizada com sucesso.", vbInformation
    Exit Sub
err:
    MsgBox "Não foi possível realizar a exportação" & vbNewLine & err.Description, vbCritical
End Sub
    
