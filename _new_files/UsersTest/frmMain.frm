Attribute VB_Name = "frmMain"
Attribute VB_Base = "0{18AB7CDC-50CE-422B-8B3D-764090DFD216}{D5881DA4-F9C7-4792-AB3F-BDE2C2BF8425}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False


Private Sub ControlsClear()
    Dim xControl As MSForms.Control
    
    For Each xControl In Me.Controls
        Select Case VBA.TypeName(xControl)
            Case "TextBox", "ComboBox"
                xControl.Value = vbNullString
        End Select
    Next xControl
End Sub

Private Function ValidarCampos() As Boolean
    Dim xControl As MSForms.Control, sFields As String
    
    For Each xControl In Me.Controls
        Select Case VBA.TypeName(xControl)
            Case "TextBox", "ComboBox"
            If xControl.Value = "" Then sFields = sFields & vbNewLine
            If Not ValidarCampos Then ValidarCampos = True
        End Select
    Next xControl
    
    If ValidarCampos Then MsgBox "Campos vazios: " & vbNewLine & sFields
End Function
