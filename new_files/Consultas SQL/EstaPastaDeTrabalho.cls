Attribute VB_Name = "EstaPastaDeTrabalho"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True
Option Explicit

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Application.OnKey "{INSERT}"
End Sub

Private Sub Workbook_Open()
    Application.OnKey "{INSERT}", "Exibir_Formulário"
    Planilha1.Select
    Cells.Clear
    Cells.ColumnWidth = 10
    Range("A1").Select
    With Range("B2")
        .Value = "Para executar o formulário pressione a tecla [Insert] ou use o atalho Ctrl + Shift + S."
        .Font.Bold = True
        .Font.Size = 14
        .Font.ColorIndex = 3
    End With
End Sub
