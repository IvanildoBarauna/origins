Attribute VB_Name = "UserForm2"
Attribute VB_Base = "0{2335601B-289D-4EFC-9292-82DE5878D480}{42BE38CF-DA68-443A-974B-953EF980C85A}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_Terminate()
    Load UserForm1
    UserForm1.Show
End Sub
