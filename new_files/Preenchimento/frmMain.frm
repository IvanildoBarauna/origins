Attribute VB_Name = "frmMain"
Attribute VB_Base = "0{C9613EE0-81BE-49FF-88FC-339093763AB2}{02E9716C-BFA2-4E8D-95B5-8DA399E6F401}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub btn_1_Click()
Dim xControlTXT  As MSForms.Control

    For Each xControlTXT In Me.Controls
        If xControlTXT.Tag = "txts" Then xControlTXT.Value = UCase(xControlTXT.Value)
    Next xControlTXT
    
End Sub
