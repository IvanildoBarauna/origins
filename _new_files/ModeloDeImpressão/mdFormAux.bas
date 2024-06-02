Attribute VB_Name = "mdFormAux"
Option Explicit
Public Sub AbrirForm()
    formImportador.Show True
End Sub

Public Function SheetsName(ByRef wbk As Workbook) As Variant
    Dim tmpArr      As Variant
    Dim lControl    As Long
    
    ReDim tmpArr(1 To wbk.Sheets.Count) As String
    
    For lControl = LBound(tmpArr) To UBound(tmpArr)
        tmpArr(lControl) = wbk.Sheets(lControl).Name
    Next lControl
    
    SheetsName = tmpArr
    Erase tmpArr
End Function

Public Sub PrintSheet(ByRef FRM As MSForms.UserForm)
    Dim wb As Excel.Workbook
    Dim ws As Excel.Worksheet
    
    Set wb = Application.Workbooks.Open(Application.GetOpenFilename())
    Set ws = wb.Sheets(formImportador.comboSheet.Value)
    
    If ValidateControls(FRM) Then Exit Sub
    
    ws.PrintOut
    wb.Close False
    
    Application.ScreenUpdating = True
End Sub

Public Function ValidateControls(ByRef FRM As MSForms.UserForm) As Boolean
    Dim xControl As MSForms.Control, sFields As String
            
    For Each xControl In FRM.Controls
        Select Case TypeName(xControl)
            Case "TextBox", "ComboBox"
                If xControl.Value = "" Then
                    sFields = sFields & vbNewLine & "Campo: " & xControl.Tag
                    ValidateControls = True
                End If
        End Select
    Next xControl
    
    If ValidateControls Then MsgBox "Os seguintes campos estão vazios: " & vbNewLine & sFields, _
        vbExclamation, "Validando Controles"
End Function

