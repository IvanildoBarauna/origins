Attribute VB_Name = "EstaPasta_de_trabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)

DEACTIVATE_APP

DEACTIVATE_

Application.StatusBar = ""


End Sub

Private Sub Workbook_Open()

ActiveWorkbook.RefreshAll

ACTIVATE_APP

Application.DisplayFormulaBar = False

VALIDA_OPEN.Activate

Dim USER As String

USER = Environ("USERNAME")

Select Case USER

Case Is = "jsilv061", "ijuni002"


CAPA

Application.DisplayFormulaBar = True

Exit Sub

Case Else

MsgBox "Usuário não autorizado para abertura da planilha!", vbCritical, AppName

Application.DisplayFormulaBar = True

ThisWorkbook.Close False

End Select

End Sub
