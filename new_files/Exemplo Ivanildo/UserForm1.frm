Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{53CF8227-13F4-4F37-94A1-58F1EC12CCDF}{78A01E24-190B-4AAF-B07B-893CDA45ACDD}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()
    Debug.Print ValidateControls(Me)
End Sub

Private Sub UserForm_Initialize()
    'PercorreControles
End Sub

Sub PercorreControles()
    
    Dim xCtrl As MSForms.Control
    
    For Each xCtrl In Me.Controls
        Debug.Print xCtrl.BackColor, xCtrl.Name
    Next xCtrl
    
End Sub


Public Function ValidateControls(ByRef FRM As MSForms.UserForm) As Boolean
    Dim xControl    As MSForms.Control
    Dim sFields     As String
            
    For Each xControl In Me.Controls
        Debug.Print xControl.Name, TypeName(xControl)
        If TypeOf xControl Is MSForms.TextBox Or TypeOf xControl Is MSForms.ComboBox Then
                If xControl.Value = "" Then
                    sFields = sFields & vbNewLine & "Campo: " & xControl.Tag
                    ValidateControls = True
                End If
        End If
    Next xControl
    
    If ValidateControls Then MsgBox "Os seguintes campos estão vazios: " & sFields, vbExclamation, Me.Caption

End Function
