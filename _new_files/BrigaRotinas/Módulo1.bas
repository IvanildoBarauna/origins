Attribute VB_Name = "Módulo1"
Option Explicit

Sub Main()
Attribute Main.VB_ProcData.VB_Invoke_Func = " \n14"
    Const nLoops As Integer = 15
    Dim wbGiovani   As Workbook
    Dim wbIvanildo  As Workbook
    Dim iCounter    As Integer
    Dim ws          As Worksheet
    Dim iTime       As Single
    Dim booAux      As Boolean
        
    If Application.Workbooks.Count < 2 Then
        Set wbGiovani = Application.Workbooks.Open("Dessafio 01 - Giovani.xlsm")
        Set wbIvanildo = Application.Workbooks("Desafio 01 - Resolvido.xlsm")
    Else
        Set wbGiovani = Application.Workbooks("Dessafio 01 - Giovani.xlsm")
        Set wbIvanildo = Application.Workbooks("Desafio 01 - Resolvido.xlsm")
    End If
    
    Set ws = Planilha1
    
    For iCounter = 2 To nLoops
        iTime = VBA.Timer
        Application.Run "'" & wbGiovani.Name & "'" & "!Desafio"
        ws.Cells(iCounter, "A").Value2 = VBA.Format(VBA.Timer - iTime, "0.000") * 1
    Next iCounter
    
    For iCounter = 2 To nLoops
        iTime = VBA.Timer
        Application.Run "'" & wbIvanildo.Name & "'" & "!TranferData"
        ws.Cells(iCounter, "B").Value2 = VBA.Format(VBA.Timer - iTime, "0.000") * 1
    Next iCounter
End Sub
