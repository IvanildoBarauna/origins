Attribute VB_Name = "wbk"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)
CRIARLOG "ARQUIVO FECHADO"
LOG "ARQUIVO FECHADO"
LOG2 "ARQUIVO FECHADO"
LOG_tbl "ARQUIVO FECHADO"
wbk.Save
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
CRIARLOG "ARQUIVO SALVO"
LOG "ARQUIVO SALVO"
LOG2 "ARQUIVO SALVO"
LOG_tbl "ARQUIVO SALVO"
End Sub

Private Sub Workbook_Open()
CRIARLOG "ARQUIVO ABERTO"
LOG "ARQUIVO ABERTO"
LOG2 "ARQUIVO ABERTO"
LOG_tbl "ARQUIVO ABERTO"
End Sub


