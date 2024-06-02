Attribute VB_Name = "mDataFormAux"
Option Explicit

Public Function ValidateEmptyFieldsNewTask(FRM As UserForm) As Boolean
    Const sTag As String = "NewTask"
    Dim xCtrl As Control
    
    For Each xCtrl In FRM.Controls
        Select Case TypeName(xCtrl)
            Case "TextBox", "ComboBox"
                If xCtrl.Tag = sTag Then
                    If xCtrl.Value = vbNullString Then
                        ValidateEmptyFieldsNewTask = True
                        Exit For
                    End If
                End If
        End Select
    Next xCtrl
End Function

Public Function ValidateEmptyFields(FRM As UserForm) As Boolean
    Dim xCtrl As Control
    
    For Each xCtrl In FRM.Controls
        Select Case TypeName(xCtrl)
            Case "TextBox", "ComboBox"
                If xCtrl.Value = vbNullString Then
                    ValidateEmptyFields = True
                    Exit For
                End If
        End Select
    Next xCtrl
End Function

Public Sub ClearControls(FRM As UserForm)
    Dim xCtrl As Control
    
    For Each xCtrl In FRM.Controls
        Select Case TypeName(xCtrl)
            Case "TextBox", "ComboBox"
                xCtrl.Value = vbNullString
        End Select
    Next xCtrl
End Sub

Public Sub FormCall()
    frmPrincipal.Show False
End Sub

