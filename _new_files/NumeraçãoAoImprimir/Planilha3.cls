Attribute VB_Name = "Planilha3"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Attribute VB_Control = "btnPrint, 3, 0, MSForms, CommandButton"
Option Explicit

Private Sub btnPrint_Click()
    Dim DocNum As Integer
    Dim ActualNum As String
    
    On Error GoTo Err:
    Me.PrintOut
    ActualNum = Me.Range("CONTROLE").Value2
    DocNum = VBA.Conversion.CInt(VBA.Right(ActualNum, 1))
    
    Me.Range("CONTROLE").Value2 = VBA.Replace(ActualNum, DocNum, DocNum + 1)
    
    MsgBox "Documento impresso com sucesso! Número de controle foi atualizado.", vbInformation
    Exit Sub
Err:
    MsgBox "Erro ao imprimir: " & VBA.Err.Description & " - " & VBA.Err.Number, vbCritical
End Sub
