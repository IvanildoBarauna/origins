Attribute VB_Name = "mdFill"
Sub codificar()

Range("a3").Select

Do While ActiveCell.Offset(0, 2) = 0 ' 1 loopp

    ActiveCell.Offset(0, 1).Select
    Selection.Copy
    ActiveCell.Offset(0, -1).Select
    Selection.PasteSpecial 1
    ActiveCell.Offset(1, 0).Select
            
    Do While ActiveCell.Offset(0, 2).Value = "" '2 loop
        
        ActiveCell.Offset(-1, 0).Copy
        Selection.PasteSpecial 1
        ActiveCell.Offset(1, 0).PasteSpecial 1
          
    Loop '2 lopp
    

Loop ' 1 loop
End Sub

Sub CODIFICAR_FUNCIONANDO()
    Dim iTime   As Singleg
    Dim item    As Integer
    Dim ws      As Worksheet
    Dim LastRow As Long
    
    Set ws = Planilha1
    LastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    Application.ScreenUpdating = False
    If ws.Range("A3").Value2 <> "" Then Call DeleteData(1)
    iTime = VBA.Timer
    
    Range("A3").Select

    For item = 1 To 10932
        If ActiveCell.Offset(0, 2) = 0 Then
            ActiveCell.Offset(0, 1).Select
            Selection.Copy
            ActiveCell.Offset(0, -1).Select
            Selection.PasteSpecial 1
            ActiveCell.Offset(1, 0).Select
        Else
            ActiveCell.Offset(-1, 0).Copy
            Selection.PasteSpecial 1
            ActiveCell.Offset(1, 0).PasteSpecial 1
        End If
    Next item
    Application.ScreenUpdating = True
    MsgBox "TEMPO DECORRIDO: " & VBA.Format(Timer - iTime, "0.00 SEGUNDOS"), vbInformation
End Sub

Public Sub PreenchimentoComCells()
    Dim ws      As Worksheet
    Dim LastRow As Long
    Dim item    As Long
    Dim iTime   As Single
     
    iTime = Timer
        
    Set ws = Planilha1
    LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    Application.ScreenUpdating = False
    
    If ws.Range("A2").Value2 <> "" Then Call DeleteData(1)
    
    With ws
        For item = 2 To LastRow
            If .Cells(item, 3).Value2 = 0 Then
                .Cells(item, 1).Value2 = .Cells(item, 2).Value2
            Else
                .Cells(item, 1).Value2 = .Cells(item - 1, 1).Value2
            End If
        Next item
    End With
    Application.ScreenUpdating = True
    MsgBox "TEMPO DECORRIDO: " & VBA.Format(Timer - iTime, "0.00 SEGUNDOS"), vbInformation
End Sub

Public Sub PreenchimentoComMatriz()
    Dim ws          As Worksheet
    Dim LastRow     As Long
    Dim item        As Long
    Dim iTime       As Single
    Dim Values      As Variant
    
    iTime = Timer
        
    Set ws = Planilha1
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Range("A3").Value2 <> "" Then Call DeleteData(1)
    LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ReDim Values(2 To LastRow, 1 To 1) As Long
    
    With ws
        For item = 2 To LastRow
            If .Cells(item, 3).Value2 = 0 Then
                Values(item, 1) = .Cells(item, 2).Value2
            Else
                Values(item, 1) = Values(item - 1, 1)
            End If
        Next item
        .Range("A2", .Cells(LastRow, 1)).Value2 = Values
    End With
    MsgBox "TEMPO DECORRIDO: " & VBA.Format(Timer - iTime, "0.00 SEGUNDOS"), vbInformation
End Sub

Public Sub DeleteData(ByVal iColumn As Integer)
    Dim ws      As Worksheet
    Dim LastRow As Long
    
    Set ws = Planilha1
    
    With ws
         LastRow = .Cells(.Rows.Count, iColumn).End(xlUp).Row
        .Range("A2", .Cells(LastRow, iColumn)).Delete xlUp
    End With
End Sub

