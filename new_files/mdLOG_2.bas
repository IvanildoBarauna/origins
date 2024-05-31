Attribute VB_Name = "mdLOG_2"
Option Explicit

Sub LOG(Ação As String)

Dim wst As Worksheet
    Set wst = shtLOG

Dim linha As Long
    linha = wst.Range("F" & Rows.Count).End(xlUp).Row + 1

Dim Momento As Date
    Momento = Now
    
Dim USER As String
    USER = Environ("USERNAME")
    
Dim Computer As String
    Computer = Environ("COMPUTERNAME")
    
    
 With wst
 
        .Cells(linha, 6) = Momento
        .Cells(linha, 7) = Computer
        .Cells(linha, 8) = USER
        .Cells(linha, 9) = Ação
 
 End With

wst.Columns.EntireColumn.AutoFit

End Sub
