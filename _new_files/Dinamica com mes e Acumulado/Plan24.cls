Attribute VB_Name = "Plan24"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub OptionButton1_Click()
Plan2.[aa1].Value = 1
BarraDeProgresso
Plan2.Calculate
End Sub

Private Sub OptionButton2_Click()
Plan2.[aa1].Value = 2

BarraDeProgresso
Plan2.Calculate
End Sub
