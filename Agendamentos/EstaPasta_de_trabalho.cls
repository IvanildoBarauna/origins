Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_Activate()
    FullScreen Ligar
End Sub

Private Sub Workbook_Deactivate()
    FullScreen Desligar
End Sub
