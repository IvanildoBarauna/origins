Attribute VB_Name = "wsReports"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Activate()
    Me.Range("A1").Select
    ThisWorkbook.RefreshAll
    ModoTelaCheia Ligado
    MsgBox "O relatório está atualizado. " & VBA.Now
End Sub

Private Sub Worksheet_Deactivate()
    ModoTelaCheia Desligado
End Sub
