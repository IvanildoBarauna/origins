Attribute VB_Name = "Módulo1"
Option Explicit
Private Sub cmdImportExcel_Click()
    Dim importFileName As Variant
    Dim importWorkbook As Workbook
    Dim importSheet As Worksheet
    Dim importRange As Range
    
    importFileName = Application.GetOpenFilename(FileFilter:="Arquivo do Excel (*.xls; *.xlsx), *.xls;*.xlsx", Title:="Escolha um arquivo do Excel")
    
    If importFileName = False Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Set importWorkbook = Application.Workbooks.Open(importFileName)
    Set importSheet = importWorkbook.Worksheets(1)
    
    With importSheet
        Set importRange = .Range(.Range("A2"), .Range("I" & .Rows.Count).End(xlUp))
        importRange.Copy
    End With
    
    wsDados.Range("A1").PasteSpecial xlValues
    wsDados.Range("A1").PasteSpecial xlPasteFormats
    Application.CutCopyMode = False
    
    importWorkbook.Close
    
    Application.ScreenUpdating = True
End Sub

