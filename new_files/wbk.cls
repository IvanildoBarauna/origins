Attribute VB_Name = "wbk"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Private Sub Workbook_BeforeClose(Cancel As Boolean)

Application.StatusBar = ""

End Sub

Private Sub Workbook_Open()

shtDIGITAÇÃO.Activate
Application.StatusBar = ("Gerencial BPA 2017 - Todos os direitos reservados à Ivanildo Junior.  |  Nome do Usuário logado: " & Environ("USERNAME") & "  |  Nome do Computador: " & Environ("COMPUTERNAME"))


End Sub
