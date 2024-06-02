Attribute VB_Name = "mdMain"
Option Explicit

Public Sub ExtrairFiliais()
    Dim ws          As Worksheet, wsNew As Worksheet
    Dim LastRow     As Long, LastCol As Integer
    Dim rng         As Range, rngNew As Range
    Dim sFilial     As String
    Dim iRow        As Long, iCol As Integer
    Dim Quant       As Long
    
    Set ws = wsFiliais
    If ws.FilterMode Then ws.ShowAllData
    
    With ws
        LastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        LastCol = .Cells(1, .Columns.Count).End(xlToLeft).Column
        Set rng = .Range(.Cells(1, 1), .Cells(LastRow, LastCol))
    End With
    
    sFilial = StrConv(VBA.InputBox("Digite o nome da Filial desejada:", "Extrair Filiais"), vbProperCase)
    
    If isExists(sFilial) Then
        Application.ScreenUpdating = False
        Quant = QTDValues(sFilial) + 1
        rng.AutoFilter Field:=2, Criteria1:=sFilial
        Set wsNew = CreatedSheet(sFilial)
        If Not wsNew Is Nothing Then
            wsNew.Cells.Delete shift:=xlUp
        Else
            Set wsNew = ThisWorkbook.Sheets.Add(after:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
            wsNew.Name = sFilial
        End If
        Set rngNew = wsNew.Range("A1").Resize(Quant, rng.Columns.Count)
        rng.CurrentRegion.Copy rngNew
        rngNew.Columns.EntireColumn.AutoFit
        ws.ShowAllData
        ws.Select
        Application.ScreenUpdating = True
        MsgBox "Os dados da filial de " & sFilial & " foram copiados com sucesso!", vbInformation
    Else
        MsgBox "Nenhuma filial válida foi informada, por favor verifique a filial digitada!", vbExclamation
    End If
End Sub

Private Function isExists(FilialValue As String) As Boolean
    Dim ws          As Worksheet: Set ws = wsFiliais
    Dim LastRow     As Long
    Dim rng         As Range
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Set rng = ws.Range("B1").Resize(LastRow)
    
    If Not FilialValue = "" And Not rng.Find(FilialValue, ws.Cells(1, 2), xlValues, True) Is Nothing Then
        isExists = True
    End If
    
End Function

 Function QTDValues(FilialValue) As Long
    Dim ws          As Worksheet: Set ws = wsFiliais
    Dim LastRow     As Long
    Dim rng         As Range
    Dim iRow        As Long
    Dim iCounter    As Integer
    
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    
    Set rng = ws.Range("B1").Resize(LastRow)
    
    For iRow = 1 To LastRow
        If ws.Cells(iRow, "B").Value2 = FilialValue Then iCounter = iCounter + 1
    Next iRow
    
    QTDValues = iCounter
End Function

Private Function CreatedSheet(wsName As String) As Worksheet
    Dim ws As Worksheet
    
    For Each ws In ThisWorkbook.Sheets
        If ws.Name = wsName Then Set CreatedSheet = ws: Exit For
    Next ws
    
End Function

