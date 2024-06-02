Attribute VB_Name = "Módulo1"
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
Public Sub PreenchimentoComMatriz()
    Dim ws          As Worksheet
    Dim LastRow     As Long
    Dim item        As Long
    Dim iTime       As Single
    Dim Values      As Variant
    
    iTime = Timer
        
    Set ws = Planilha1
    LastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If ws.Range("A3").Value2 <> "" Then ws.Range("A3:A" & LastRow).Delete xlUp
    LastRow = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row
    
    ReDim Values(3 To LastRow, 1 To 1)
    
    With ws
        For item = 3 To LastRow
            If .Cells(item, 3).Value2 = 0 Then
                Values(item, 1) = .Cells(item, 2).Value2
            Else
                Values(item, 1) = Values(item - 1, 1)
            End If
        Next item
        .Range("A3", .Cells(LastRow, 1)).Value2 = Values
    End With
    MsgBox "TEMPO DECORRIDO: " & VBA.Format(Timer - iTime, "0.00 SEGUNDOS"), vbInformation
End Sub
