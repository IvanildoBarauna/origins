Attribute VB_Name = "formSampleMaskDate"
Attribute VB_Base = "0{7A90E479-D2E7-4503-A1A5-3D70CD910080}{24D47229-B7BB-4329-9197-1BF15E14C4A9}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub txtDate_Exit(ByVal Cancel As MSForms.ReturnBoolean)
Dim WSF As WorksheetFunction
    With Me
        VBA.Replace .txtDate.Value, "/", ""
        .txtDate.Value = VBA.Format(txtDate.Value, "00\/00\/0000")
    End With
End Sub

Private Sub txtDate_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    
'    txtDate.MaxLength = 10
'
'    Select Case KeyAscii
'        Case Asc("0") To Asc("9")
'            Select Case Len(txtDate)
'                Case 2, 5
'                txtDate.SelText = "/"
'            End Select
'        Case Else
'            KeyAscii = 0
'    End Select

End Sub
