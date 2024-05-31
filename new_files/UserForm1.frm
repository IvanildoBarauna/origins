Attribute VB_Name = "UserForm1"
Attribute VB_Base = "0{57EED488-E0FB-451E-A9F6-818F19DC8A23}{C1F4B324-FAC3-4EC4-BFE5-EBC234AE2B0D}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub CommandButton1_Click()
    Me.Hide
    UserForm2.Show
End Sub

Private Sub UserForm_Initialize()
    'Unload UserForms(UserForm2)
End Sub
