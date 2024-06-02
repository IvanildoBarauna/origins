Attribute VB_Name = "wbContagem"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_Open()
    With Excel.Application
        .OnKey "{INSERT}", "Inserir"
        .OnKey "^{DEL}", "ClearData"
        .OnKey "%{INSERT}", "InserirFechamento"
    End With
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    With Excel.Application
        .OnKey "{INSERT}"
        .OnKey "{DEL}"
        .OnKey "%{INSERT}"
    End With
End Sub



