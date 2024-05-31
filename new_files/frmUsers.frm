Attribute VB_Name = "frmUsers"
Attribute VB_Base = "0{5814ADFB-30DA-4917-BECD-4E21F2A32AE9}{F685B484-D61B-401F-9D83-A7A674FAD528}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
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

