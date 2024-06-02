Attribute VB_Name = "Módulo1"
Sub AreaTransfer1()
    Dim linha As Integer
    Dim Round As Integer
    Dim resultado As Integer
    Dim i As Integer
    Dim vRow As Integer
    Dim LastRow As Long
    
    LastRow = zPlanilha.Cells(zPlanilha.Rows.Count, 2).End(xlUp).Row
    
    With zPlanilha
        For i = 3 To LastRow Step 5
            If .Cells(i, 2).Value2 = "" Then Exit For
            Round = Round + 1
            vRow = vRow + 1
        Next i
    End With
    
    For vRow = 1 To Round Step 5
        If zPlanilha.Cells(vRow, coluna + 2).Value2 = "" Then
            Exit Sub
        ElseIf zPlanilha.Cells(vRow, coluna + 5) = "" And zPlanilha.Cells(vRow, coluna + 12) = "" Then
            MsgBox "Está faltando registro de Bomb"
            Exit Sub
        End If
    Next vRow
    
    If MsgBox("Serão enviados " & Round & " Rounds" & " Continuar procedimento ?", _
        vbYesNo + vbQuestion, "Banco de Dados") = vbYes Then
        zPlanilha.Cells(69, coluna + 1) = Round
        AreaTransfer2
    Else
        resultado = 0
        Exit Sub
    End If
End Sub
Sub AreaTransfer2()
Dim vRow As Integer, vCol As Integer
Dim vLoop As Integer

vRow = 4
LastRow = zBD.Cells(zBD.Rows.Count, 1).End(xlUp).Row

For vLoop = 1 To zPlanilha.Cells(69, vCol + 1).Value2
    With zPlanilha
        .Cells(59, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(59, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(59, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(59, vCol + 4) = .Cells(vRow - 1, vCol + 2)
        .Cells(59, vCol + 5) = .Cells(vRow - 1, vCol + 3)
        .Cells(59, vCol + 6) = .Cells(vRow, vCol + 3)
        .Cells(59, vCol + 7) = .Cells(vRow + 1, vCol + 3)
        .Cells(59, vCol + 8) = .Cells(vRow + 2, vCol + 3)
        .Cells(60, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(60, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(60, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(60, vCol + 4) = .Cells(vRow - 1, vCol + 2)
        .Cells(60, vCol + 5) = .Cells(vRow - 1, vCol + 3)
        .Cells(60, vCol + 6) = .Cells(vRow, vCol + 4)
        .Cells(60, vCol + 7) = .Cells(vRow + 1, vCol + 4)
        .Cells(60, vCol + 8) = .Cells(vRow + 2, vCol + 4)
        .Cells(61, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(61, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(61, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(61, vCol + 4) = .Cells(vRow - 1, vCol + 2)
        .Cells(61, vCol + 5) = .Cells(vRow - 1, vCol + 3)
        .Cells(61, vCol + 6) = .Cells(vRow, vCol + 5)
        .Cells(61, vCol + 7) = .Cells(vRow + 1, vCol + 5)
        .Cells(61, vCol + 8) = .Cells(vRow + 2, vCol + 5)
        .Cells(62, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(62, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(62, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(62, vCol + 4) = .Cells(vRow - 1, vCol + 2)
        .Cells(62, vCol + 5) = .Cells(vRow - 1, vCol + 3)
        .Cells(62, vCol + 6) = .Cells(vRow, vCol + 6)
        .Cells(62, vCol + 7) = .Cells(vRow + 1, vCol + 6)
        .Cells(62, vCol + 8) = .Cells(vRow + 2, vCol + 6)
        .Cells(63, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(63, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(63, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(63, vCol + 4) = .Cells(vRow - 1, vCol + 2)
        .Cells(63, vCol + 5) = .Cells(vRow - 1, vCol + 3)
        .Cells(63, vCol + 6) = .Cells(vRow, vCol + 7)
        .Cells(63, vCol + 7) = .Cells(vRow + 1, vCol + 7)
        .Cells(63, vCol + 8) = .Cells(vRow + 2, vCol + 7)
        .Cells(64, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(64, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(64, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(64, vCol + 4) = .Cells(vRow - 1, vCol + 9)
        .Cells(64, vCol + 5) = .Cells(vRow - 1, vCol + 10)
        .Cells(64, vCol + 6) = .Cells(vRow, vCol + 10)
        .Cells(64, vCol + 7) = .Cells(vRow + 1, vCol + 10)
        .Cells(64, vCol + 8) = .Cells(vRow + 2, vCol + 10)
        .Cells(65, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(65, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(65, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(65, vCol + 4) = .Cells(vRow - 1, vCol + 9)
        .Cells(65, vCol + 5) = .Cells(vRow - 1, vCol + 10)
        .Cells(65, vCol + 6) = .Cells(vRow, vCol + 11)
        .Cells(65, vCol + 7) = .Cells(vRow + 1, vCol + 11)
        .Cells(65, vCol + 8) = .Cells(vRow + 2, vCol + 11)
        .Cells(66, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(66, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(66, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(66, vCol + 4) = .Cells(vRow - 1, vCol + 9)
        .Cells(66, vCol + 5) = .Cells(vRow - 1, vCol + 10)
        .Cells(66, vCol + 6) = .Cells(vRow, vCol + 12)
        .Cells(66, vCol + 7) = .Cells(vRow + 1, vCol + 12)
        .Cells(66, vCol + 8) = .Cells(vRow + 2, vCol + 12)
        .Cells(67, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(67, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(67, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(67, vCol + 4) = .Cells(vRow - 1, vCol + 9)
        .Cells(67, vCol + 5) = .Cells(vRow - 1, vCol + 10)
        .Cells(67, vCol + 6) = .Cells(vRow, vCol + 13)
        .Cells(67, vCol + 7) = .Cells(vRow + 1, vCol + 13)
        .Cells(67, vCol + 8) = .Cells(vRow + 2, vCol + 13)
        .Cells(68, vCol + 1) = .Cells(vRow - 1, vCol + 8)
        .Cells(68, vCol + 2) = .Cells(vRow, vCol + 21)
        .Cells(68, vCol + 3) = .Cells(vRow + 1, vCol + 15)
        .Cells(68, vCol + 4) = .Cells(vRow - 1, vCol + 9)
        .Cells(68, vCol + 5) = .Cells(vRow - 1, vCol + 10)
        .Cells(68, vCol + 6) = .Cells(vRow, vCol + 14)
        .Cells(68, vCol + 7) = .Cells(vRow + 1, vCol + 14)
        .Cells(68, vCol + 8) = .Cells(vRow + 2, vCol + 14)
        
        If .Cells(vRow - 1, vCol + 3) = "Defesa" Then
            .Range("Bomb") = .Cells(vRow - 1, vCol + 5)
        Else
            .Range("Bomb") = .Cells(vRow - 1, vCol + 12)
        End If
    End With
    
    zPlanilha.Range("ATDados").Copy Destination:=zBD.Cells(LastRow, 1)
    zPlanilha.Range("ATDados") = ""
    
    vRow = vRow + 5
    LastRow = LastRow + 10
Next vLoop
    zPlanilha.Range("A69") = ""
End Sub
