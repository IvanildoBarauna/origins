Attribute VB_Name = "wsConsolidado"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    If Target.Address = "$C$5" And Not Target.Value2 = "" Then
        Excel.Application.EnableEvents = False
        Target.Value2 = wsApoios.ListObjects("tbSessao").DataBodyRange(Target.Value2, 2).Value2
        Excel.Application.EnableEvents = True
    End If
End Sub
