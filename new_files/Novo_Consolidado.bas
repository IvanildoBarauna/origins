Attribute VB_Name = "Novo_Consolidado"
Public a As Variant
Public d As Integer

Sub New_Consolidado()

ini = Time
d = 0
Application.ScreenUpdating = False
Plan17.Select
Range(Cells(3, 3), Cells(3000, 30)).Select
Selection.ClearContents
Calculate

a = "Production"
Call consolidaDados
d = Application.CountA(Range(Cells(3, 7), Cells(2000, 7)))

a = "Leaders"
Call consolidaDados
d = Application.CountA(Range(Cells(3, 7), Cells(2000, 7)))

a = "Staff"
Call consolidaDados

Worksheets("Production").Select
Cells(1, 1).Select
Calculate

Call InserirDadosMapa
fim = Time
Application.ScreenUpdating = True
MsgBox "Atualizado em: " & Format(fim - ini, "hh:mm:ss")

End Sub

Sub consolidaDados()

Plan17.Select
Cells(2, 3).Select

Do While ActiveCell <> ""

x = 1
y = ActiveCell.Value

Worksheets(a).Select
Z = Application.Match(y, Range(Cells(7, 1), Cells(7, 100)), 0)
Plan17.Select

If IsError(Z) <> True Then

Worksheets(a).Select
Range(Cells(8, Application.Match(y, Range(Cells(7, 1), Cells(7, 100)), 0)), Cells(Application.CountA(Range(Cells(7, 6), Cells(1500, 6))) + 6, Application.Match(y, Range(Cells(7, 1), Cells(7, 100)), 0))).Select
Selection.Copy
Plan17.Select
ActiveCell.Offset(1 + d, 0).Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
ActiveCell.Offset(-1 - d, 1).Select

Else

ActiveCell.Offset(0, 1).Select

End If

Loop

Plan17.Cells(1, 1).Select
Calculate
End Sub



