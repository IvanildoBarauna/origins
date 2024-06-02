Attribute VB_Name = "mdFormsAux"
Option Explicit

Public Function ValidateEmptyControls(ByRef FRM As UserForm) As Boolean
    Dim xControl As MSForms.Control
    Dim sList    As String
    
    For Each xControl In FRM.Controls
        Select Case TypeName(xControl)
            Case "TextBox", "ComboBox"
                If xControl.Value = vbNullString Then
                    If Not ValidateEmptyControls Then _
                        ValidateEmptyControls = True
                    sList = sList & vbNewLine & xControl.Tag
                End If
        End Select
    Next xControl
    
    If ValidateEmptyControls Then MsgBox "Preencha os campos abaixo:" _
        & vbNewLine & sList
End Function

Public Sub ClearFields(ByRef FRM As MSForms.UserForm)
    Dim xControl As MSForms.Control
    
    For Each xControl In FRM.Controls
        Select Case VBA.TypeName(xControl)
            Case "TextBox", "ComboBox"
                xControl.Value = VBA.vbNullString
        End Select
    Next xControl
    
End Sub

Public Function OptButtonString(ByRef FRM As UserForm) As String
    Dim xControl As MSForms.Control
    
    For Each xControl In FRM.Controls
        If TypeOf xControl Is MSForms.OptionButton Then
            If xControl.Value Then
                OptButtonString = xControl.Caption
                Exit Function
            End If
        End If
    Next xControl
    
End Function

Public Sub SaveData(FRM As Object)
    Dim xCtrl As MSForms.Control
    
    For Each xCtrl In FRM.Controls
        Select Case VBA.TypeName(xCtrl)
            Case "TextBox", "ComboBox"
                VBA.SaveSetting FRM.Name, VBA.TypeName(xCtrl), xCtrl.Name, xCtrl.Value
        End Select
    Next xCtrl
End Sub

Public Sub GetData(FRM As Object)
    Dim xCtrl As MSForms.Control
    
    For Each xCtrl In FRM.Controls
        Select Case VBA.TypeName(xCtrl)
            Case "TextBox", "ComboBox"
                xCtrl.Value = VBA.GetSetting(FRM.Name, VBA.TypeName(xCtrl), xCtrl.Name, xCtrl.Value)
        End Select
    Next xCtrl
End Sub
