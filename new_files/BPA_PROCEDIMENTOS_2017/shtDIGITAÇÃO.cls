Attribute VB_Name = "shtDIGITAÇÃO"
Attribute VB_Base = "0{00020820-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

Application.ScreenUpdating = False

If ActiveCell.Column = 2 And ActiveCell.Value <> "" Then

On Error GoTo Erro

qtd = InputBox("Digite a quantidade de procedimentos realizados em " & ActiveCell.Value & ":", App)

Erro:

    If qtd = Empty Or IsNumeric(qtd) = False Then

    MsgBox "Dados inválidos, digite a quantidade de procedimentos.", vbCritical, App
    
        Exit Sub
        
        Else

ActiveCell.Offset(0, 3).Select
ActiveCell.Value = qtd
ActiveCell.Offset(1, -4).Select

qtd = Empty

End If

End If

Application.ScreenUpdating = True

End Sub


