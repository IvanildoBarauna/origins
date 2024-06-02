Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
    Call CriarLOG("Arquivo foi salvo")
End Sub

Private Sub Workbook_Open()
    Call CriarLOG("Arquivo aberto")
End Sub
