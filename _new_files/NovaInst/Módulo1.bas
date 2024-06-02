Attribute VB_Name = "Módulo1"
Option Explicit

Public Sub NewInst()
    Dim ExcelApp    As Excel.Application
    Dim wbk         As Workbook
    Dim ws          As Worksheet
    
    Set ExcelApp = New Excel.Application
    
    ExcelApp.Visible = True
    
    Set wbk = ExcelApp.Workbooks.Add
    Set ws = wbk.Sheets(1)
    
    ws.Range("A1").Value = "Esta é uma nova Instância do Excel"
End Sub
