Attribute VB_Name = "mdMain"
Option Explicit

Public Sub ImportFile()
    On Error GoTo err
    Dim sPath   As String
    Dim iLinha  As String
    Dim ws      As Worksheet
    Dim iRow    As Long
    Dim rng     As Range
    Dim LastRow As Long
    Dim fDialog As FileDialog
    
    Set ws = wsDados
    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    
    With fDialog
        .Title = "Selecione o arquivo de texto do Exercício 6"
        .Filters.Add "Texto", "*.txt", 1
        .InitialFileName = ThisWorkbook.Path
        If .Show Then sPath = .SelectedItems(1) Else Exit Sub
    End With
    
    Application.ScreenUpdating = False
    
    If Not VBA.Right(sPath, 5) = "6.txt" Then
        MsgBox "O arquivo selecionado está incorreto, por favor selecione o arquivo 'Exercício 6.txt'", vbExclamation
        Exit Sub
    End If
    
    ws.Cells.ClearContents
    
    Open sPath For Input As #1
    
    Do Until VBA.EOF(1)
        Line Input #1, iLinha
        iRow = iRow + 1
        If iRow > 3 Then
            If Not iLinha = "" Then
                ws.Cells(iRow - 3, 1).Value2 = iLinha
            End If
        End If
    Loop

    Close #1
    
    With ws
        .Range("2:3,6:8").EntireRow.Delete
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        Set rng = .Range("A1").Resize(LastRow)
        rng.TextToColumns .Range("A1"), xlDelimited, ConsecutiveDelimiter:=True, Space:=True
        .Range("B:D,F:H,J:L").EntireColumn.Delete
        .Range("C3").Value2 = .Range("B3").Value2
        .Range("B3").Value2 = ""
        .Range("C10").Value2 = .Range("B10").Value2
        .Range("B10").Value2 = ""
        .Range("A1").Value2 = "Empresa"
        .Range("B1").Value2 = "Núm. Contrato"
        .Range("C1").Value2 = "Vl. Recebido"
        .Range("A1:C1").Font.Bold = True
        .Columns.EntireColumn.AutoFit
    End With

    MsgBox "Processo concluído.", vbInformation
    Application.ScreenUpdating = True
    Exit Sub
err:
    Application.ScreenUpdating = True
    MsgBox "Não foi possível concluir o processo." & vbNewLine & err.Description, vbCritical
End Sub


