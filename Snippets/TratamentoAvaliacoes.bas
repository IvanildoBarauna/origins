Attribute VB_Name = "mdTratamentoAval"
Option Explicit
Public Sub TratamentoAvaliacoes()
    Dim LastRow     As Integer
    Dim rng             As Range
    Dim irng        As Range
    Dim lineup As Integer
    Dim linedown As Integer
    Dim OutLastRow As Integer
    Dim titulo As String
    Dim nota As String
    Dim vdata As String
    Dim autor As String
    Dim comentario As String
    Dim versao As String
    Dim feedbackdev As String
    
    LastRow = wsInput.Range("A1048576").End(xlUp).Row
    
    Set rng = wsInput.Range("A1:A" & LastRow)
    
    OutLastRow = wsOutput.Range("A1048576").End(xlUp).Row + 1
    
    wsOutput.Range("A2:h2", "A" & OutLastRow & ":" & "G" & OutLastRow).Delete
    
    For Each irng In rng
        If irng.Value2 = "Responder" Or irng.Value2 = "Relatar um problema" Then
            irng.Delete xlUp
        ElseIf irng.Value2 = "Brasil" Then
            lineup = irng.Row - 1
            linedown = irng.Row + 1
            wsInput.Range(lineup & ":" & lineup, linedown & ":" & linedown).Delete xlUp
        ElseIf VBA.Left(irng.Value2, 6) = "Vers‹o" Then
            'Tratamento normal das avalia�›es sem resposta do desenvolvedor
            If wsOutput.Cells(irng.Row - 2, 1).Value2 <> "Resposta a uma avalia�‹o anterior" Then
                titulo = wsInput.Cells(irng.Row - 4, 1).Value2
                nota = wsInput.Cells(irng.Row - 3, 1).Value2
                vdata = VBA.Right(wsInput.Cells(irng.Row - 2, 1).Value2, 17)
                autor = wsInput.Cells(irng.Row - 2, 1).Value2
                comentario = wsInput.Cells(irng.Row - 1, 1).Value2
                versao = Replace(Replace(irng.Value2, "Vers‹o ", ""), Right(irng.Value2, 4), "")
                OutLastRow = wsOutput.Range("A1048576").End(xlUp).Row + 1
                With wsOutput
                    .Cells(OutLastRow, 1).Value2 = Replace(Trim(vdata), "de", "")
                    .Cells(OutLastRow, 2).Value2 = Replace(Replace(Replace(autor, vdata, ""), "de ", ""), " Ð ", "")
                    .Cells(OutLastRow, 3).Value2 = nota
                    .Cells(OutLastRow, 4).Value2 = titulo
                    .Cells(OutLastRow, 5).Value2 = comentario
                    .Cells(OutLastRow, 6).Value2 = versao
                    .Cells(OutLastRow, 7).Value2 = Format(Now(), "mm/dd/yyyy hh:mm:ss")
                End With
            Else
            'Tratamento pra quando o loop achar uma resposta do desenvolvedor na avalia�‹o
                
                titulo = wsInput.Cells(irng.Row - 8, 1).Value2
                nota = wsInput.Cells(irng.Row - 7, 1).Value2
                vdata = VBA.Right(wsInput.Cells(irng.Row - 5, 1).Value2, 17)
                autor = wsInput.Cells(irng.Row - 5, 1).Value2
                comentario = wsInput.Cells(irng.Row - 4, 1).Value2
                versao = Replace(Replace(irng.Value2, "Vers‹o ", ""), Right(irng.Value2, 4), "")
                feedbackdev = comentario = wsInput.Cells(irng.Row - 1, 1).Value2
                OutLastRow = wsOutput.Range("A1048576").End(xlUp).Row + 1
                With wsOutput
                    .Cells(OutLastRow, 1).Value2 = Replace(Trim(vdata), "de", "")
                    .Cells(OutLastRow, 2).Value2 = Replace(Replace(Replace(autor, vdata, ""), "de ", ""), " Ð ", "")
                    .Cells(OutLastRow, 3).Value2 = nota
                    .Cells(OutLastRow, 4).Value2 = titulo
                    .Cells(OutLastRow, 5).Value2 = comentario
                    .Cells(OutLastRow, 6).Value2 = versao
                    .Cells(OutLastRow, 7).Value2 = Format(Now(), "mm/dd/yyyy hh:mm:ss")
                    .Cells(OutLastRow, 8).Value2 = feedbackdev
                End With
                
                
            End If
        End If
    Next irng
    
    Set irng = Nothing
    Set rng = Nothing
    
    Set rng = wsOutput.Range("A1:A" & OutLastRow)
    
    For Each irng In rng
        If irng.Value2 = "valia�‹o anterior" Then
            wsOutput.Cells(OutLastRow + 1, 5).Value2 = wsOutput.Cells(irng.Row, 3).Value2
            vdata = VBA.Right(wsOutput.Cells(irng.Row, 4).Value2, 17)
            wsOutput.Cells(OutLastRow + 1, 1).Value2 = Replace(Trim(vdata), "de", "")
            autor = Replace(wsOutput.Cells(irng.Row, 4).Value2, "Editada em ", "")
            wsOutput.Cells(OutLastRow + 1, 2).Value = Replace(Replace(Replace(autor, vdata, ""), "de ", ""), " Ð ", "")
            wsOutput.Cells(OutLastRow + 1, 8).Value = wsOutput.Cells(irng.Row, 5).Value2
            wsOutput.Cells(OutLastRow + 1, 7).Value = Format(Now(), "mm/dd/yyyy hh:mm:ss")
            wsOutput.Range(irng.Address).EntireRow.Delete
            OutLastRow = wsOutput.Range("A1048576").End(xlUp).Row
        End If
    Next irng
    
    MsgBox "Conclu’do", vbInformation
End Sub
