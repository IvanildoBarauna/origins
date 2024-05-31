Attribute VB_Name = "mdAux"
Option Explicit
Public Sub SaveCopy()
    Dim wbk     As Workbook
    Dim ws      As Worksheet
    Dim shp     As Shape
    Dim Path    As String
        
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Path = VBA.Environ("USERPROFILE") & "\Desktop\CopiaSemShape.xlsm"
    ThisWorkbook.SaveCopyAs Path
    Set wbk = Application.Workbooks.Open(Filename:=Path)
    wbk.Sheets(1).Shapes.Range(Array("Button 1")).Delete
    wbk.Close savechanges:=True
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    
    MsgBox "Uma cópia deste arquivo foi realizada com sucesso!" & vbNewLine & _
            "Localize o arquivo no seguinte caminho: " & Path, vbInformation
End Sub

