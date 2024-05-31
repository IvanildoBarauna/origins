Attribute VB_Name = "mdLOG_3"
Option Explicit

Sub LOG2(Ação As String)

Dim sht As Worksheet
    Set sht = shtLOG
    
 Dim line As Long
    line = sht.Cells(Cells.Rows.Count, 11).End(xlUp).Row + 1
    
Dim Hora As Date
    Hora = Now
    
Dim Usuário, Comp As String
    Usuário = Environ("USERNAME")
    Comp = Environ("COMPUTERNAME")
    
  With sht
        .Cells(line, 11) = Hora
        .Cells(line, 12) = Usuário
        .Cells(line, 13) = Comp
        .Cells(line, 14) = Ação
        .Columns.EntireColumn.AutoFit
  End With
  
End Sub
